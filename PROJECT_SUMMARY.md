# Two-Tier LLM Inference & Semantic Caching Architecture

A detailed architectural summary and execution flow specification covering the primary server engine ([server.py](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py)), the environment configuration ([.env](file:///c:/Users/shrim/OneDrive/Desktop/FEP/.env)), and the web client interface ([index.html](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html)).

---

## 1. Project Overview

This project implements a low-latency **Two-Tier LLM Caching and Inference Engine** paired with a real-time web playground and observability dashboard. 

The primary objective is to dramatically reduce generation latency, operational compute costs, and remote API token expenditure by checking incoming queries against a two-stage in-memory cache before initiating remote inference calls to Groq's LLM endpoints:
* **Tier 1 (Exact Match Cache):** O(1) hash map lookup for identical prompts, returning cached responses with zero model execution time.
* **Tier 2 (Semantic Embedding Cache):** Dense vector similarity matching using local sentence embeddings (`all-MiniLM-L6-v2`) to serve paraphrased queries meeting or exceeding a configurable cosine similarity threshold.
* **Fallback (Remote Groq Inference):** Full cache misses dispatch directly to the upstream Groq LLM API (`qwen/qwen3.8-27b`). Newly generated outputs are cached along with their embedding representations.

---

## 2. Core Files Breakdown

### A. Environment Configuration: [`.env`](file:///c:/Users/shrim/OneDrive/Desktop/FEP/.env)
Contains runtime authentication and model selection configurations:
* `GROQ_API_KEY`: Secret authentication bearer token required for Groq Cloud API endpoints.
* `GROQ_MODEL_NAME`: Specifies the generative model identifier (defaults to `qwen/qwen3.8-27b`).

---

### B. Backend Engine: [`server.py`](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py)
A [FastAPI](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L259-L263) service listening by default on `http://127.0.0.1:8004`. It coordinates local vector embedding computation, two-tier cache evaluation, upstream Groq API calls, and telemetry logging.

#### Key Architectural Components:
1. **Local Embedding Model Loading** ([server.py:38-41](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L38-L41)):
   Loads `sentence-transformers/all-MiniLM-L6-v2` locally via Hugging Face `transformers` in evaluation mode (`emb_model.eval()`).
2. **Mean-Pooled Embeddings** ([compute_embedding](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L44-L62)):
   Tokenizes prompt strings, computes mean-pooled sentence representations weighted by attention masks, and applies $L_2$ normalization (`torch.nn.functional.normalize(..., p=2)`).
3. **Cosine Similarity** ([compute_cosine_similarity](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L64-L66)):
   Computes vector dot products between normalized tensors.
4. **Cache Data Structure** ([CacheEntry](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L68-L81)):
   A slotted dataclass containing the prompt text, `max_new_tokens`, completion text, embedding tensor, creation timestamp, token usage statistics, and TTL expiration logic.
5. **Two-Tier Cache Manager** ([TwoTierCache](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L83-L176)):
   * [_evict_expired()](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L90-L96): Evicts entries older than the configured TTL (default 3600 seconds).
   * [lookup_exact()](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L101-L135): Fast lookup checking exact `(prompt, max_new_tokens)` match first, falling back to exact and normalized prompt matches across any token budget (guaranteeing pre-warmed items hit cache).
   * [lookup_semantic()](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L137-L175): Scan over active cached embeddings (including pre-warmed entries), computing cosine similarity against the query embedding. Returns best entry if similarity $\ge$ `similarity_threshold`.
   * [insert()](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L177-L202): Adds new completions and embedding tensors to the cache.
   * [clear()](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L204-L215): Removes an individual prompt or purges all cached entries.
6. **Async LLM Client** ([async_generate_llm](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L197-L232)):
   Uses an asynchronous `httpx.AsyncClient` session to call `https://api.groq.com/openai/v1/chat/completions` with user prompts and token budgets.
7. **Telemetry Persistence** ([append_telemetry_log](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L192-L195)):
   Appends structured JSON records for every processed request to `telemetry_metrics.jsonl`.

#### Primary HTTP Endpoints:
* `GET /` & `GET /playground` ([server.py:358-362](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L358-L362)): Serves the single-page application [index.html](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html).
* `POST /generate` ([server.py:473-586](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L473-L586)): Core generation endpoint.
* `GET /api/status` ([server.py:364-374](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L364-L374)): Server status, cache size, and active model.
* `GET /api/cache_entries` ([server.py:376-391](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L376-L391)): Active cache entries and ages.
* `GET /api/logs` ([server.py:393-407](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L393-L407)): Telemetry records in reverse-chronological order.
* `GET /api/request/{request_id}` ([server.py:445-460](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L445-L460)): Detailed telemetry record for a specific request.
* `POST /clear_cache` ([server.py:462-471](file:///c:/Users/shrim/OneDrive/Desktop/FEP/server.py#L462-L471)): Cache eviction handler.

---

### C. Web Client Interface: [`index.html`](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html)
A modern client interface built with Tailwind CSS, Google Geist fonts, and Material Symbols.

#### User Interface Sections:
1. **Interactive Playground View (`#view-playground`)** ([index.html:193-389](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html#L193-L389)):
   * Prompt input area with quick paraphrase test chips.
   * Parameter controls: `max_tokens` slider (1–100), Tier 1 exact cache toggle, Tier 2 semantic cache toggle, and cosine threshold slider (0.70–0.99).
   * Completion response box with transparency alert banner showing match type and score.
   * Real-time metrics bar: Total latency (ms), model compute time (ms), queue time (ms), and token counts.
2. **Dashboard & Cache Inspector View (`#view-dashboard`)** ([index.html:391-556](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html#L391-L556)):
   * System health card with active cache capacity meter and active model name.
   * Active Cache Entries table with individual key eviction buttons.
   * Telemetry metrics bento: Average latency, tokens/sec (TPS), and cache hit rate percentage.
   * Recent Telemetry feed with direct inspection triggers.
3. **Request Inspector Modal (`#request-modal`)** ([index.html:1017-1160](file:///c:/Users/shrim/OneDrive/Desktop/FEP/index.html#L1017-L1160)):
   * Comprehensive drill-down for any `req-XXXXXXXX` record.
   * Shows prompt input, matched cache key, completion output, latency metrics, and pipeline configuration.
   * Quick utility actions: Open in Playground, Copy Prompt/Output, Copy JSON Telemetry, and Copy cURL command.

---

## 3. End-to-End Flow of Execution

```
                     ┌─────────────────────────────┐
                     │ Client Request (index.html) │
                     └──────────────┬──────────────┘
                                    │ POST /generate
                                    ▼
                     ┌─────────────────────────────┐
                     │  server.py: generate()      │
                     │  - Validate prompt          │
                     │  - Assign request_id        │
                     └──────────────┬──────────────┘
                                    │
                       Is Tier 1 Cache Enabled?
                                    │
                   YES ─────────────┴───────────── NO
                    │                               │
                    ▼                               │
      ┌───────────────────────────┐                 │
      │ lookup_exact() (Hash Map) │                 │
      └─────────────┬─────────────┘                 │
                    │                               │
            Tier 1 Hit?                             │
                    │                               │
           YES ─────┴───── NO                       │
            │               │                       │
            │               ▼                       ▼
            │     Is Tier 2 Semantic Enabled & Entries Exist?
            │               │
            │      YES ─────┴───── NO
            │       │               │
            │       ▼               │
            │ ┌───────────────────────────┐         │
            │ │ compute_embedding()       │         │
            │ │ (all-MiniLM-L6-v2)        │         │
            │ └─────────────┬─────────────┘         │
            │               ▼                       │
            │ ┌───────────────────────────┐         │
            │ │ lookup_semantic()         │         │
            │ │ (Cosine Sim >= Threshold) │         │
            │ └─────────────┬─────────────┘         │
            │               │                       │
            │       Tier 2 Hit?                     │
            │               │                       │
            │      YES ─────┴───── NO               │
            │       │               │               │
            ▼       ▼               ▼               ▼
      ┌───────────────────┐   ┌───────────────────────────┐
      │ Return Cache Hit  │   │ Full Cache Miss:          │
      │ - Exact: 100% Sim │   │ async_generate_llm()      │
      │ - Sem: Cosine Sim │   │ (Groq Cloud API)          │
      │ - Compute: 0.0 ms │   └─────────────┬─────────────┘
      └─────────┬─────────┘                 │
                │                           ▼
                │             ┌───────────────────────────┐
                │             │ Cache Insertion:          │
                │             │ two_tier_cache.insert()   │
                │             └─────────────┬─────────────┘
                │                           │
                ▼                           ▼
      ┌───────────────────────────────────────────────────┐
      │ _record_and_build_response()                      │
      │ 1. Log JSON record to telemetry_metrics.jsonl     │
      │ 2. Return GenerateResponse payload to Client      │
      └─────────────────────────┬─────────────────────────┘
                                │
                                ▼
      ┌───────────────────────────────────────────────────┐
      │ Client UI Update (index.html)                     │
      │ - Display completion text & match status banner   │
      │ - Update HUD metrics & refresh telemetry logs     │
      └───────────────────────────────────────────────────┘
```

### Detailed Sequence Breakdown:

1. **Bootstrapping**:
   * `server.py` invokes `load_dotenv()` to pull `GROQ_API_KEY` and `GROQ_MODEL_NAME` from `.env`.
   * Hugging Face tokenizer and `all-MiniLM-L6-v2` are loaded into memory and set to `eval()`.
   * An instance of `TwoTierCache` is initialized.
   * FastAPI lifespan manager instantiates a persistent `httpx.AsyncClient` session.
2. **Dispatching Prompt**:
   * The user enters prompt text in `index.html` and clicks **"Generate Output"**.
   * `runGeneration()` sends an HTTP POST request to `/generate?enable_cache=...&enable_semantic_cache=...&similarity_threshold=...`.
3. **Tier 1 Exact Match Evaluation**:
   * `server.py` trims the prompt and performs `two_tier_cache.lookup_exact(...)`.
   * If found: returns cached output immediately (`proc_time_ms = 0.0`, `match_type = "exact"`, `cache_hit = True`).
4. **Tier 2 Semantic Embedding Evaluation**:
   * If Tier 1 misses and semantic caching is active:
   * The server runs `compute_embedding()` asynchronously via `asyncio.get_running_loop().run_in_executor()`.
   * `lookup_semantic()` calculates cosine similarity across cached entries (including pre-warmed knowledge pairs).
   * If top score meets `similarity_threshold`: returns cached response (`proc_time_ms = 0.0`, `match_type = "semantic"`, `similarity_score = <score>`).
5. **Full Cache Miss & Remote Groq Inference**:
   * If both tiers miss: calls `async_generate_llm(...)`.
   * Sends the payload to Groq's completion API and tracks roundtrip latency.
   * If caching is enabled: stores the completion and prompt embedding in `TwoTierCache` via `insert()`.
6. **Telemetry & Client Rendering**:
   * Telemetry data is written to `telemetry_metrics.jsonl`.
   * Response JSON is returned to `index.html`, which updates the completion text, match banner, latency HUD, and telemetry tables.
