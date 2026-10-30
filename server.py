import asyncio
from contextlib import asynccontextmanager
from dataclasses import dataclass
import json
import logging
import os
import signal
import time
from typing import Optional
import uuid

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import httpx
from pydantic import BaseModel, Field
import torch
from transformers import AutoModel, AutoTokenizer

# Load environment variables from .env file
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("inference_server")

# Configuration
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()
GROQ_MODEL_NAME = os.environ.get("GROQ_MODEL_NAME", "qwen/qwen3.8-27b")
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"

EMBEDDING_MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"
LOG_FILE_PATH = "telemetry_metrics.jsonl"
DEFAULT_CACHE_TTL_SECONDS = 3600.0
AUTO_SHUTDOWN_SECONDS = float(os.environ.get("AUTO_SHUTDOWN_SECONDS", "0"))
INDEX_HTML_PATH = "index.html" if os.path.exists("index.html") else os.path.join("static", "index.html")

# Embedding model for Tier 2 semantic matching
logger.info("Loading embedding model '%s'...", EMBEDDING_MODEL_NAME)
emb_tokenizer = AutoTokenizer.from_pretrained(EMBEDDING_MODEL_NAME)
emb_model = AutoModel.from_pretrained(EMBEDDING_MODEL_NAME)
emb_model.eval()

def compute_embedding(prompt: str) -> Optional[torch.Tensor]:
    """Compute normalized mean-pooled sentence embedding for a prompt."""
    clean = prompt.strip()
    if not clean:
        return None
    try:
        encoded = emb_tokenizer([clean], padding=True, truncation=True, return_tensors="pt")
        with torch.no_grad():
            outputs = emb_model(**encoded)
            token_embeddings = outputs[0]
            mask = encoded["attention_mask"].unsqueeze(-1).expand(token_embeddings.size()).float()
            sum_embeddings = torch.sum(token_embeddings * mask, dim=1)
            sum_mask = torch.clamp(mask.sum(dim=1), min=1e-9)
            mean_pooled = sum_embeddings / sum_mask
            return torch.nn.functional.normalize(mean_pooled, p=2, dim=1)[0]
    except Exception as e:
        logger.error("Failed to compute embedding for prompt: %r. Error: %s", prompt, e)
        return None


def compute_cosine_similarity(vec_a: torch.Tensor, vec_b: torch.Tensor) -> float:
    return float(torch.dot(vec_a, vec_b).item())


def normalize_prompt(prompt: str) -> str:
    """Standardizes prompt text by lowercasing and trimming excessive whitespace."""
    return " ".join(prompt.strip().lower().split())


@dataclass(slots=True)
class CacheEntry:
    prompt: str
    max_new_tokens: int
    response_text: str
    embedding: Optional[torch.Tensor]
    timestamp: float
    prompt_tokens: int = 0
    generated_tokens: int = 0
    total_model_tokens: int = 0

    def is_expired(self, max_age_seconds: float) -> bool:
        return (time.time() - self.timestamp) > max_age_seconds


class TwoTierCache:
    """Two-tier cache combining exact string lookup (Tier 1) and semantic similarity (Tier 2)."""

    def __init__(self, default_ttl_seconds: float = DEFAULT_CACHE_TTL_SECONDS):
        self.default_ttl_seconds = default_ttl_seconds
        self.entries: dict[tuple[str, int], CacheEntry] = {}

    def _evict_expired(self, max_age_seconds: Optional[float] = None) -> None:
        ttl = max_age_seconds if max_age_seconds is not None else self.default_ttl_seconds
        now = time.time()
        expired = [k for k, v in self.entries.items() if (now - v.timestamp) > ttl]
        for k in expired:
            del self.entries[k]

    def lookup_exact(
        self,
        prompt: str,
        max_new_tokens: Optional[int] = None,
        max_age_seconds: Optional[float] = None
    ) -> Optional[CacheEntry]:
        self._evict_expired(max_age_seconds)
        clean = prompt.strip()

        # 1. Exact match with identical (clean, max_new_tokens)
        if max_new_tokens is not None and (clean, max_new_tokens) in self.entries:
            return self.entries[(clean, max_new_tokens)]

        # 2. Exact prompt string match (any token budget, e.g. pre-warmed / seeded entries)
        matching_clean = [entry for (p, _), entry in self.entries.items() if p == clean]
        if matching_clean:
            if max_new_tokens is not None:
                # Prefer entry closest to requested max_new_tokens, then most recent
                matching_clean.sort(key=lambda e: (abs(e.max_new_tokens - max_new_tokens), -e.timestamp))
            else:
                matching_clean.sort(key=lambda e: -e.timestamp)
            return matching_clean[0]

        # 3. Normalized prompt match (case-insensitive & whitespace-normalized)
        norm = normalize_prompt(clean)
        matching_norm = [entry for (p, _), entry in self.entries.items() if normalize_prompt(p) == norm]
        if matching_norm:
            if max_new_tokens is not None:
                matching_norm.sort(key=lambda e: (abs(e.max_new_tokens - max_new_tokens), -e.timestamp))
            else:
                matching_norm.sort(key=lambda e: -e.timestamp)
            return matching_norm[0]

        return None

    def lookup_semantic(
        self,
        prompt_embedding: torch.Tensor,
        max_new_tokens: int,
        similarity_threshold: float,
        max_age_seconds: Optional[float] = None
    ) -> tuple[Optional[CacheEntry], Optional[float]]:
        self._evict_expired(max_age_seconds)

        # Consider all active cached embeddings (including pre-warmed entries)
        candidates = [
            entry for entry in self.entries.values()
            if entry.embedding is not None
        ]
        if not candidates:
            return None, None

        best_entry: Optional[CacheEntry] = None
        best_sim = -1.0

        for entry in candidates:
            try:
                sim = compute_cosine_similarity(prompt_embedding, entry.embedding)
                # Pick best similarity; tie-break on token match
                if sim > best_sim or (
                    best_entry is not None
                    and abs(sim - best_sim) < 1e-4
                    and entry.max_new_tokens == max_new_tokens
                    and best_entry.max_new_tokens != max_new_tokens
                ):
                    best_sim = sim
                    best_entry = entry
            except Exception as e:
                logger.warning("Error comparing embedding against cached prompt '%s': %s", entry.prompt, e)

        if best_entry is not None and best_sim >= similarity_threshold:
            return best_entry, round(best_sim, 4)

        return None, round(best_sim, 4) if best_entry is not None else None

    def insert(
        self,
        prompt: str,
        max_new_tokens: int,
        response_text: str,
        embedding: Optional[torch.Tensor] = None,
        prompt_tokens: int = 0,
        generated_tokens: int = 0,
        total_model_tokens: int = 0
    ) -> CacheEntry:
        clean = prompt.strip()
        if embedding is None:
            embedding = compute_embedding(clean)

        entry = CacheEntry(
            prompt=clean,
            max_new_tokens=max_new_tokens,
            response_text=response_text,
            embedding=embedding,
            timestamp=time.time(),
            prompt_tokens=prompt_tokens,
            generated_tokens=generated_tokens,
            total_model_tokens=total_model_tokens
        )
        self.entries[(clean, max_new_tokens)] = entry
        return entry

    def clear(self, prompt: Optional[str] = None) -> int:
        if prompt:
            clean = prompt.strip()
            norm = normalize_prompt(clean)
            keys = [k for k in self.entries if k[0] == clean or normalize_prompt(k[0]) == norm]
            for k in keys:
                del self.entries[k]
            return len(keys)
        count = len(self.entries)
        self.entries.clear()
        return count


@dataclass(slots=True)
class LLMGenerationResult:
    generated_text: str
    processing_time_sec: float
    prompt_tokens: int
    generated_tokens: int
    total_model_tokens: int


# App State
two_tier_cache = TwoTierCache()
http_client: Optional[httpx.AsyncClient] = None


def append_telemetry_log(record: dict) -> None:
    with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(record) + "\n")


async def async_generate_llm(prompt: str, max_new_tokens: int) -> LLMGenerationResult:
    if not GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY is not configured. Please set GROQ_API_KEY in your .env file or environment variables.")
    start_time = time.time()
    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json"
    }

    client = http_client or httpx.AsyncClient(timeout=30.0)
    resp = await client.post(
        GROQ_API_URL,
        headers=headers,
        json={
            "model": GROQ_MODEL_NAME,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_new_tokens
        }
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Groq API error ({resp.status_code}): {resp.text}")

    r = resp.json()
    text = r["choices"][0]["message"]["content"] or ""
    usage = r.get("usage", {})
    p_tok = usage.get("prompt_tokens", len(prompt.split()))
    g_tok = usage.get("completion_tokens", len(text.split()))
    t_tok = usage.get("total_tokens", p_tok + g_tok)

    return LLMGenerationResult(
        generated_text=text,
        processing_time_sec=time.time() - start_time,
        prompt_tokens=p_tok,
        generated_tokens=g_tok,
        total_model_tokens=t_tok
    )


@asynccontextmanager
async def lifespan(app: FastAPI):
    global http_client
    http_client = httpx.AsyncClient(timeout=30.0)
    shutdown_task = None

    if AUTO_SHUTDOWN_SECONDS > 0:
        async def shutdown_after_timeout():
            await asyncio.sleep(AUTO_SHUTDOWN_SECONDS)
            logger.info("Automatic shutdown timer reached; stopping server.")
            os.kill(os.getpid(), signal.SIGINT)

        shutdown_task = asyncio.create_task(shutdown_after_timeout())

    try:
        yield
    finally:
        if shutdown_task is not None:
            shutdown_task.cancel()
            await asyncio.gather(shutdown_task, return_exceptions=True)
        if http_client is not None:
            await http_client.aclose()


app = FastAPI(
    title="LLM Inference Server - Two-Tier Cache Architecture",
    description="Inference server with Tier 1 exact match cache and Tier 2 semantic embedding cache.",
    lifespan=lifespan
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Request / Response Schemas
class GenerateRequest(BaseModel):
    prompt: str = Field(..., description="Input prompt text", min_length=1)
    max_new_tokens: int = Field(default=20, ge=1, le=100)


class GenerateResponse(BaseModel):
    request_id: str
    prompt: str
    generated_text: str
    queue_wait_time_ms: float = 0.0
    processing_time_ms: float
    total_latency_ms: float
    batch_size: int = 1
    cache_hit: bool
    match_type: str = Field(default="none", description="Match type: 'exact', 'semantic', or 'none'")
    similarity_score: Optional[float] = Field(default=None, description="Similarity score for semantic match")
    matched_prompt: Optional[str] = Field(default=None, description="Original prompt matched against")
    prompt_tokens: int
    generated_tokens: int
    total_model_tokens: int


class SeedCacheItem(BaseModel):
    prompt: str = Field(..., description="Prompt question or query")
    response_text: str = Field(..., description="Authoritative completion / answer")
    max_new_tokens: int = Field(default=50, ge=1, le=500)


class SeedCacheRequest(BaseModel):
    items: list[SeedCacheItem]


class VectorInspectRequest(BaseModel):
    prompt: str = Field(..., description="Query prompt to inspect against cache", min_length=1)
    top_k: int = Field(default=5, ge=1, le=20)
    similarity_threshold: Optional[float] = Field(default=0.90, ge=0.50, le=1.00)


def _record_and_build_response(
    request_id: str,
    prompt: str,
    generated_text: str,
    arrival_time: float,
    proc_time_ms: float,
    cache_hit: bool,
    match_type: str,
    similarity_score: Optional[float],
    matched_prompt: Optional[str],
    prompt_tokens: int,
    generated_tokens: int,
    total_model_tokens: int,
    enable_cache: bool,
    enable_semantic_cache: bool,
    similarity_threshold: float,
    batch_size: int = 1,
    queue_wait_ms: float = 0.0
) -> GenerateResponse:
    total_latency_ms = (time.time() - arrival_time) * 1000.0

    response = GenerateResponse(
        request_id=request_id,
        prompt=prompt,
        generated_text=generated_text,
        queue_wait_time_ms=round(max(queue_wait_ms, 0.0), 2),
        processing_time_ms=round(proc_time_ms, 2),
        total_latency_ms=round(total_latency_ms, 2),
        batch_size=batch_size,
        cache_hit=cache_hit,
        match_type=match_type,
        similarity_score=similarity_score,
        matched_prompt=matched_prompt,
        prompt_tokens=prompt_tokens,
        generated_tokens=generated_tokens,
        padded_tokens=0,
        total_model_tokens=total_model_tokens
    )

    append_telemetry_log({
        "request_id": request_id,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(arrival_time)),
        "prompt": prompt,
        "generated_text": generated_text,
        "prompt_char_len": len(prompt.strip()),
        "max_new_tokens": total_model_tokens,
        "queue_time_ms": round(max(queue_wait_ms, 0.0), 2),
        "processing_time_ms": round(proc_time_ms, 2),
        "total_latency_ms": round(total_latency_ms, 2),
        "batch_size": batch_size,
        "cache_hit": cache_hit,
        "match_type": match_type,
        "similarity_score": similarity_score,
        "matched_prompt": matched_prompt,
        "prompt_tokens": prompt_tokens,
        "generated_tokens": generated_tokens,
        "padded_tokens": 0,
        "total_model_tokens": total_model_tokens,
        "batching_enabled": False,
        "cache_enabled": enable_cache,
        "semantic_cache_enabled": enable_semantic_cache,
        "similarity_threshold": similarity_threshold
    })

    return response


# Endpoints
@app.get("/")
@app.get("/playground")
def playground():
    return FileResponse(INDEX_HTML_PATH)


@app.get("/api/status")
def status_api():
    return {
        "status": "ok",
        "stage": "Stage E: Two-Tier Cache Architecture (Tier 1 Exact + Tier 2 Semantic)",
        "generative_model": GROQ_MODEL_NAME,
        "log_file": LOG_FILE_PATH,
        "cache_entries": len(two_tier_cache.entries),
        "cache_ttl_seconds": two_tier_cache.default_ttl_seconds
    }


@app.get("/api/cache_entries")
def get_cache_entries():
    two_tier_cache._evict_expired()
    now = time.time()
    return [
        {
            "prompt": entry.prompt,
            "max_new_tokens": entry.max_new_tokens,
            "response_preview": entry.response_text[:80] + ("..." if len(entry.response_text) > 80 else ""),
            "timestamp": entry.timestamp,
            "age_seconds": round(now - entry.timestamp, 1),
            "has_embedding": entry.embedding is not None
        }
        for entry in two_tier_cache.entries.values()
    ]


@app.get("/api/logs")
def get_logs(limit: int = Query(default=50, ge=1, le=500)):
    if not os.path.exists(LOG_FILE_PATH):
        return []
    records = []
    with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line_str = line.strip()
            if line_str:
                try:
                    records.append(json.loads(line_str))
                except Exception:
                    pass
    return records[-limit:][::-1]


@app.get("/api/benchmark_results")
def get_benchmark_results():
    path = "benchmark_results.json"
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="No benchmark results found. Run benchmark.py first.")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to read benchmark results: {e}")


@app.get("/api/evaluation_results")
def get_evaluation_results():
    path = "semantic_evaluation_results.json"
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="No evaluation results found. Run evaluate_semantic_pooling.py first.")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to read evaluation results: {e}")


@app.get("/api/prompt_results")
def get_prompt_results():
    path = "semantic_prompt_results.json"
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="No prompt run results found. Run run_semantic_prompts.py first.")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to read prompt results: {e}")


@app.get("/api/request/{request_id}")
def get_request_by_id(request_id: str):
    if not os.path.exists(LOG_FILE_PATH):
        raise HTTPException(status_code=404, detail="No telemetry logs found.")
    with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line_str = line.strip()
            if line_str:
                try:
                    record = json.loads(line_str)
                    if record.get("request_id") == request_id:
                        return record
                except Exception:
                    pass
    raise HTTPException(status_code=404, detail=f"Request '{request_id}' not found.")


@app.post("/clear_cache")
def clear_cache(prompt: Optional[str] = Query(default=None, description="Specific prompt to evict from cache")):
    cleared_count = two_tier_cache.clear(prompt)
    return {
        "status": "cleared",
        "cleared_prompt": prompt,
        "cleared_count": cleared_count,
        "remaining_entries": len(two_tier_cache.entries)
    }


# ---------------------------------------------------------------------------
# Cache Seeding & Golden Pre-warming
# ---------------------------------------------------------------------------
DEFAULT_GOLDEN_FAQS = [
    {
        "prompt": "What is the boiling point of water at sea level?",
        "response_text": "At sea level (1 atmosphere of pressure / 101.325 kPa), the boiling point of pure water is exactly 100 degrees Celsius (212 degrees Fahrenheit).",
        "max_new_tokens": 30
    },
    {
        "prompt": "What is a large language model (LLM)?",
        "response_text": "A Large Language Model (LLM) is an artificial intelligence model trained on vast corpora of text using transformer architectures to comprehend, summarize, and generate human-like language.",
        "max_new_tokens": 40
    },
    {
        "prompt": "What is the difference between RAM and ROM?",
        "response_text": "RAM (Random Access Memory) is fast, volatile primary memory used for active application state and working memory. ROM (Read-Only Memory) is non-volatile memory storing permanent firmware and boot instructions.",
        "max_new_tokens": 45
    },
    {
        "prompt": "What is the capital of France?",
        "response_text": "Paris is the capital and most populous city of France, situated along the Seine River in northern central France.",
        "max_new_tokens": 25
    },
    {
        "prompt": "Explain Newton's third law of motion.",
        "response_text": "Newton's third law states that whenever an object exerts a force on a second object, the second object exerts an equal and opposite force on the first.",
        "max_new_tokens": 35
    },
    {
        "prompt": "How does photosynthesis work in plants?",
        "response_text": "Photosynthesis is the biochemical process by which green plants convert light energy, carbon dioxide, and water into chemical energy (glucose) and release oxygen as a byproduct.",
        "max_new_tokens": 40
    },
    {
        "prompt": "How do you reverse a list in Python?",
        "response_text": "You can reverse a list in Python in-place using `my_list.reverse()`, or non-destructively using slice notation `my_list[::-1]` or the `reversed(my_list)` iterator.",
        "max_new_tokens": 35
    },
    {
        "prompt": "How far away is the Moon from Earth?",
        "response_text": "The average distance between Earth and the Moon is approximately 384,400 kilometers (238,855 miles), varying slightly along its elliptical orbit.",
        "max_new_tokens": 30
    }
]


@app.post("/api/seed_cache")
async def seed_cache(payload: SeedCacheRequest):
    if not payload.items:
        raise HTTPException(status_code=400, detail="Items list cannot be empty.")

    loop = asyncio.get_running_loop()
    seeded = []

    for item in payload.items:
        clean_p = item.prompt.strip()
        clean_r = item.response_text.strip()
        if not clean_p or not clean_r:
            continue

        emb = await loop.run_in_executor(None, compute_embedding, clean_p)
        prompt_toks = max(len(clean_p.split()), 1)
        gen_toks = max(len(clean_r.split()), 1)

        entry = two_tier_cache.insert(
            prompt=clean_p,
            max_new_tokens=item.max_new_tokens,
            response_text=clean_r,
            embedding=emb,
            prompt_tokens=prompt_toks,
            generated_tokens=gen_toks,
            total_model_tokens=prompt_toks + gen_toks
        )
        seeded.append({
            "prompt": entry.prompt,
            "max_new_tokens": entry.max_new_tokens,
            "has_embedding": entry.embedding is not None
        })

    return {
        "status": "success",
        "seeded_count": len(seeded),
        "total_cache_entries": len(two_tier_cache.entries),
        "items": seeded
    }


@app.post("/api/seed_defaults")
async def seed_defaults():
    items = [SeedCacheItem(**item) for item in DEFAULT_GOLDEN_FAQS]
    return await seed_cache(SeedCacheRequest(items=items))


# ---------------------------------------------------------------------------
# Live Vector Similarity Inspector
# ---------------------------------------------------------------------------
@app.post("/api/vector_inspect")
async def vector_inspect(payload: VectorInspectRequest):
    clean_prompt = payload.prompt.strip()
    if not clean_prompt:
        raise HTTPException(status_code=400, detail="Prompt cannot be empty.")

    two_tier_cache._evict_expired()
    query_norm = normalize_prompt(clean_prompt)
    loop = asyncio.get_running_loop()
    query_emb = await loop.run_in_executor(None, compute_embedding, clean_prompt)
    if query_emb is None:
        raise HTTPException(status_code=500, detail="Failed to compute embedding for query.")

    candidates = [
        entry for entry in two_tier_cache.entries.values()
        if entry.embedding is not None
    ]

    now = time.time()
    neighbors = []
    threshold = payload.similarity_threshold or 0.90

    for entry in candidates:
        try:
            sim = compute_cosine_similarity(query_emb, entry.embedding)
            sim_score = round(float(sim), 4)
            sim_pct = round(max(0.0, min(100.0, sim_score * 100.0)), 1)
            is_exact = (entry.prompt.strip().lower() == clean_prompt.lower())
            is_norm_exact = (normalize_prompt(entry.prompt) == query_norm)

            neighbors.append({
                "prompt": entry.prompt,
                "normalized_prompt": normalize_prompt(entry.prompt),
                "max_new_tokens": entry.max_new_tokens,
                "response_preview": entry.response_text[:140] + ("..." if len(entry.response_text) > 140 else ""),
                "full_response": entry.response_text,
                "similarity_score": sim_score,
                "similarity_pct": sim_pct,
                "age_seconds": round(now - entry.timestamp, 1),
                "is_exact": is_exact,
                "is_norm_exact": is_norm_exact,
                "passes_threshold": (sim_score >= threshold)
            })
        except Exception as e:
            logger.warning("Error inspecting neighbor '%s': %s", entry.prompt, e)

    # Sort descending by similarity score
    neighbors.sort(key=lambda x: x["similarity_score"], reverse=True)
    top_neighbors = neighbors[:payload.top_k]

    return {
        "query": clean_prompt,
        "normalized_query": query_norm,
        "total_cache_entries": len(two_tier_cache.entries),
        "inspected_count": len(candidates),
        "similarity_threshold": threshold,
        "neighbors": top_neighbors
    }


@app.post("/generate", response_model=GenerateResponse)
async def generate(
    request: GenerateRequest,
    enable_cache: bool = Query(default=True),
    enable_semantic_cache: bool = Query(default=True),
    similarity_threshold: float = Query(default=0.90, ge=0.50, le=1.00),
    cache_ttl: float = Query(default=DEFAULT_CACHE_TTL_SECONDS, ge=1.0)
):
    from fastapi.params import Query as QueryParam
    if isinstance(enable_cache, QueryParam):
        enable_cache = True if enable_cache.default is ... else bool(enable_cache.default)
    if isinstance(enable_semantic_cache, QueryParam):
        enable_semantic_cache = True if enable_semantic_cache.default is ... else bool(enable_semantic_cache.default)
    if isinstance(similarity_threshold, QueryParam):
        similarity_threshold = 0.90 if similarity_threshold.default is ... else float(similarity_threshold.default)
    if isinstance(cache_ttl, QueryParam):
        cache_ttl = DEFAULT_CACHE_TTL_SECONDS if cache_ttl.default is ... else float(cache_ttl.default)

    clean_prompt = request.prompt.strip()
    if not clean_prompt:
        raise HTTPException(status_code=400, detail="Prompt cannot be empty.")

    request_id = f"req-{uuid.uuid4().hex[:8]}"
    arrival_time = time.time()

    # Tier 1: Exact Match Cache
    if enable_cache:
        exact_entry = two_tier_cache.lookup_exact(clean_prompt, request.max_new_tokens, max_age_seconds=cache_ttl)
        if exact_entry is not None:
            return _record_and_build_response(
                request_id=request_id,
                prompt=request.prompt,
                generated_text=exact_entry.response_text,
                arrival_time=arrival_time,
                proc_time_ms=0.0,
                batch_size=0,
                queue_wait_ms=0.0,
                cache_hit=True,
                match_type="exact",
                similarity_score=1.0,
                matched_prompt=exact_entry.prompt,
                prompt_tokens=exact_entry.prompt_tokens,
                generated_tokens=exact_entry.generated_tokens,
                total_model_tokens=0,
                enable_cache=enable_cache,
                enable_semantic_cache=enable_semantic_cache,
                similarity_threshold=similarity_threshold
            )

    # Tier 2: Semantic Similarity Cache
    prompt_emb: Optional[torch.Tensor] = None
    loop = asyncio.get_running_loop()

    if enable_cache and enable_semantic_cache and len(two_tier_cache.entries) > 0:
        try:
            prompt_emb = await loop.run_in_executor(None, compute_embedding, clean_prompt)
            if prompt_emb is not None:
                sem_entry, sim_score = two_tier_cache.lookup_semantic(
                    prompt_emb,
                    request.max_new_tokens,
                    similarity_threshold=similarity_threshold,
                    max_age_seconds=cache_ttl
                )
                if sem_entry is not None:
                    return _record_and_build_response(
                        request_id=request_id,
                        prompt=request.prompt,
                        generated_text=sem_entry.response_text,
                        arrival_time=arrival_time,
                        proc_time_ms=0.0,
                        batch_size=0,
                        queue_wait_ms=0.0,
                        cache_hit=True,
                        match_type="semantic",
                        similarity_score=sim_score,
                        matched_prompt=sem_entry.prompt,
                        prompt_tokens=sem_entry.prompt_tokens,
                        generated_tokens=sem_entry.generated_tokens,
                        total_model_tokens=0,
                        enable_cache=enable_cache,
                        enable_semantic_cache=enable_semantic_cache,
                        similarity_threshold=similarity_threshold
                    )
        except Exception as e:
            logger.warning("Semantic match lookup failed for '%s': %s", clean_prompt, e)

    # Full Cache Miss - Direct Model Inference (No Queue, No Batching)
    start_proc = time.time()
    try:
        gen_res = await async_generate_llm(clean_prompt, request.max_new_tokens)
    except Exception as err:
        raise HTTPException(status_code=500, detail=f"Model execution error: {err}")
    proc_time_ms = (time.time() - start_proc) * 1000.0

    if enable_cache:
        two_tier_cache.insert(
            prompt=clean_prompt,
            max_new_tokens=request.max_new_tokens,
            response_text=gen_res.generated_text,
            embedding=prompt_emb,
            prompt_tokens=gen_res.prompt_tokens,
            generated_tokens=gen_res.generated_tokens,
            total_model_tokens=gen_res.total_model_tokens
        )

    return _record_and_build_response(
        request_id=request_id,
        prompt=request.prompt,
        generated_text=gen_res.generated_text,
        arrival_time=arrival_time,
        proc_time_ms=proc_time_ms,
        batch_size=1,
        queue_wait_ms=0.0,
        cache_hit=False,
        match_type="none",
        similarity_score=None,
        matched_prompt=None,
        prompt_tokens=gen_res.prompt_tokens,
        generated_tokens=gen_res.generated_tokens,
        total_model_tokens=gen_res.total_model_tokens,
        enable_cache=enable_cache,
        enable_semantic_cache=enable_semantic_cache,
        similarity_threshold=similarity_threshold
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8004"))
    uvicorn.run(app, host="0.0.0.0", port=port)
