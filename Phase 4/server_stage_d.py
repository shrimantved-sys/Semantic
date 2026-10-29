import asyncio
import time
from contextlib import asynccontextmanager
from typing import NamedTuple, List, Dict, Tuple
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

MODEL_NAME = "gpt2"
MAX_BATCH_SIZE = 4
BATCH_TIMEOUT = 0.2

print(f"[Stage D] Loading model '{MODEL_NAME}' onto CPU...")
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
tokenizer.padding_side = "left"
if tokenizer.pad_token is None:
    tokenizer.pad_token = tokenizer.eos_token

model = AutoModelForCausalLM.from_pretrained(MODEL_NAME)
model.eval()
print("[Stage D] Model loaded successfully!")

# Exact-Match Cache: Key is (normalized_prompt, max_new_tokens) -> generated_text
exact_cache: Dict[Tuple[str, int], str] = {}

class RequestItem(NamedTuple):
    prompt: str
    max_new_tokens: int
    future: asyncio.Future
    arrival_time: float

class BatchResult(NamedTuple):
    generated_texts: List[str]
    processing_time_sec: float
    prompt_tokens_list: List[int]
    padded_tokens_list: List[int]
    generated_tokens_list: List[int]
    total_model_tokens_list: List[int]

request_queue: asyncio.Queue[RequestItem] = asyncio.Queue()

def sync_process_batch(prompts: List[str], max_new_tokens: int) -> BatchResult:
    start_time = time.time()
    inputs = tokenizer(prompts, padding=True, return_tensors="pt")
    batch_size = len(prompts)
    max_prompt_len_in_batch = inputs["input_ids"].shape[1]
    
    real_prompt_tokens = [
        int(mask.sum().item()) for mask in inputs["attention_mask"]
    ]
    padded_tokens = [
        max_prompt_len_in_batch - real_prompt_tokens[i]
        for i in range(batch_size)
    ]
    
    with torch.no_grad():
        outputs = model.generate(
            inputs["input_ids"],
            attention_mask=inputs["attention_mask"],
            max_new_tokens=max_new_tokens,
            pad_token_id=tokenizer.eos_token_id,
            do_sample=False
        )
    
    generated_texts = [
        tokenizer.decode(output, skip_special_tokens=True)
        for output in outputs
    ]
    
    generated_tokens = [
        len(outputs[i]) - max_prompt_len_in_batch
        for i in range(batch_size)
    ]
    total_model_tokens = [len(outputs[i]) for i in range(batch_size)]
    
    return BatchResult(
        generated_texts=generated_texts,
        processing_time_sec=time.time() - start_time,
        prompt_tokens_list=real_prompt_tokens,
        padded_tokens_list=padded_tokens,
        generated_tokens_list=generated_tokens,
        total_model_tokens_list=total_model_tokens
    )

async def batch_worker():
    print(f"[Stage D Worker] Worker started with exact-match cache support.")
    while True:
        first_item: RequestItem = await request_queue.get()
        batch: List[RequestItem] = [first_item]
        deadline = time.time() + BATCH_TIMEOUT
        
        while len(batch) < MAX_BATCH_SIZE:
            remaining = deadline - time.time()
            if remaining <= 0:
                break
            try:
                item = await asyncio.wait_for(request_queue.get(), timeout=remaining)
                batch.append(item)
            except asyncio.TimeoutError:
                break
        
        batch_size = len(batch)
        prompts = [item.prompt for item in batch]
        max_tokens = max(item.max_new_tokens for item in batch)
        
        try:
            loop = asyncio.get_running_loop()
            batch_res: BatchResult = await loop.run_in_executor(
                None, sync_process_batch, prompts, max_tokens
            )
            
            # Store outputs in exact-match cache and resolve futures
            for i, item in enumerate(batch):
                gen_text = batch_res.generated_texts[i]
                cache_key = (item.prompt.strip(), item.max_new_tokens)
                exact_cache[cache_key] = gen_text  # Save to cache
                
                if not item.future.done():
                    item.future.set_result({
                        "generated_text": gen_text,
                        "processing_time_sec": batch_res.processing_time_sec,
                        "batch_size": batch_size,
                        "prompt_tokens": batch_res.prompt_tokens_list[i],
                        "padded_tokens": batch_res.padded_tokens_list[i],
                        "generated_tokens": batch_res.generated_tokens_list[i],
                        "total_model_tokens": batch_res.total_model_tokens_list[i]
                    })
        except Exception as e:
            for item in batch:
                if not item.future.done():
                    item.future.set_exception(e)
        finally:
            for _ in range(batch_size):
                request_queue.task_done()

@asynccontextmanager
async def lifespan(app: FastAPI):
    worker_task = asyncio.create_task(batch_worker())
    yield
    worker_task.cancel()

app = FastAPI(
    title="LLM Inference Server - Stage D",
    description="Server with exact-match cache and static batching.",
    lifespan=lifespan
)

class GenerateRequest(BaseModel):
    prompt: str = Field(..., description="Input prompt text", min_length=1)
    max_new_tokens: int = Field(default=20, ge=1, le=100)

class GenerateResponse(BaseModel):
    prompt: str
    generated_text: str
    queue_wait_time_sec: float
    processing_time_sec: float
    batch_size: int
    cache_hit: bool
    prompt_tokens: int
    padded_tokens: int
    generated_tokens: int
    total_model_tokens: int

@app.get("/")
def health_check():
    return {
        "status": "ok",
        "stage": "Stage D: Exact-Match Cache",
        "cache_entries": len(exact_cache)
    }

@app.post("/generate", response_model=GenerateResponse)
async def generate(request: GenerateRequest):
    clean_prompt = request.prompt.strip()
    if not clean_prompt:
        raise HTTPException(status_code=400, detail="Prompt cannot be empty.")
    
    arrival_time = time.time()
    cache_key = (clean_prompt, request.max_new_tokens)
    
    # 1. Check Exact-Match Cache BEFORE queueing
    if cache_key in exact_cache:
        print(f"[Stage D Cache] HIT for prompt: '{clean_prompt[:25]}...'")
        cached_text = exact_cache[cache_key]
        prompt_toks = len(tokenizer.encode(clean_prompt))
        gen_toks = max(len(tokenizer.encode(cached_text)) - prompt_toks, 0)
        return GenerateResponse(
            prompt=request.prompt,
            generated_text=cached_text,
            queue_wait_time_sec=0.0,
            processing_time_sec=0.0,
            batch_size=0,
            cache_hit=True,
            prompt_tokens=prompt_toks,
            padded_tokens=0,
            generated_tokens=gen_toks,
            total_model_tokens=prompt_toks + gen_toks
        )
    
    # 2. Cache MISS: Queue request for batch model execution
    print(f"[Stage D Cache] MISS for prompt: '{clean_prompt[:25]}...'")
    loop = asyncio.get_running_loop()
    future = loop.create_future()
    
    item = RequestItem(
        prompt=clean_prompt,
        max_new_tokens=request.max_new_tokens,
        future=future,
        arrival_time=arrival_time
    )
    
    await request_queue.put(item)
    
    try:
        res_data = await future
    except Exception as err:
        raise HTTPException(status_code=500, detail=f"Model execution error: {str(err)}")
    
    queue_wait_time = time.time() - arrival_time
    return GenerateResponse(
        prompt=request.prompt,
        generated_text=res_data["generated_text"],
        queue_wait_time_sec=round(queue_wait_time, 4),
        processing_time_sec=round(res_data["processing_time_sec"], 4),
        batch_size=res_data["batch_size"],
        cache_hit=False,
        prompt_tokens=res_data["prompt_tokens"],
        padded_tokens=res_data["padded_tokens"],
        generated_tokens=res_data["generated_tokens"],
        total_model_tokens=res_data["total_model_tokens"]
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8003)
