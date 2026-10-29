import asyncio
import time
from contextlib import asynccontextmanager
from typing import NamedTuple, Optional
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

MODEL_NAME = "gpt2"
print(f"[Stage B] Loading model '{MODEL_NAME}' onto CPU...")
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
if tokenizer.pad_token is None:
    tokenizer.pad_token = tokenizer.eos_token

model = AutoModelForCausalLM.from_pretrained(MODEL_NAME)
model.eval()
print("[Stage B] Model loaded successfully!")

class RequestItem(NamedTuple):
    prompt: str
    max_new_tokens: int
    future: asyncio.Future
    arrival_time: float

# In-memory queue for decoupling HTTP requests from model execution
request_queue: asyncio.Queue[RequestItem] = asyncio.Queue()

class InferenceResult(NamedTuple):
    generated_text: str
    prompt_tokens: int
    generated_tokens: int
    total_tokens: int

def sync_process_request(prompt: str, max_new_tokens: int) -> InferenceResult:
    """Synchronous model execution function (runs in thread pool)."""
    inputs = tokenizer(prompt, return_tensors="pt")
    prompt_tokens = int(inputs["input_ids"].shape[1])
    with torch.no_grad():
        outputs = model.generate(
            **inputs,
            max_new_tokens=max_new_tokens,
            pad_token_id=tokenizer.eos_token_id,
            do_sample=False
        )
    total_tokens = len(outputs[0])
    generated_tokens = total_tokens - prompt_tokens
    generated_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return InferenceResult(
        generated_text=generated_text,
        prompt_tokens=prompt_tokens,
        generated_tokens=generated_tokens,
        total_tokens=total_tokens
    )

async def request_worker():
    """Background worker popping requests from queue and sending to model."""
    print("[Stage B Worker] Background worker task initialized.")
    while True:
        item: RequestItem = await request_queue.get()
        try:
            queue_time = time.time() - item.arrival_time
            print(f"[Stage B Worker] Popped prompt '{item.prompt[:20]}...' (queue wait: {queue_time:.3f}s)")
            
            # Offload synchronous PyTorch model inference to a thread pool
            loop = asyncio.get_running_loop()
            result: InferenceResult = await loop.run_in_executor(
                None, sync_process_request, item.prompt, item.max_new_tokens
            )
            
            if not item.future.done():
                item.future.set_result(result)
        except Exception as e:
            if not item.future.done():
                item.future.set_exception(e)
        finally:
            request_queue.task_done()

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Start background worker when server starts
    worker_task = asyncio.create_task(request_worker())
    yield
    # Cancel worker task when server shuts down
    worker_task.cancel()

app = FastAPI(
    title="LLM Inference Server - Stage B",
    description="Server with decoupled async request queue and background processing worker.",
    lifespan=lifespan
)

class GenerateRequest(BaseModel):
    prompt: str = Field(..., description="Input prompt text for text generation", min_length=1)
    max_new_tokens: int = Field(default=20, ge=1, le=100, description="Number of tokens to generate")

class GenerateResponse(BaseModel):
    prompt: str
    generated_text: str
    queue_wait_time_sec: float
    prompt_tokens: int
    generated_tokens: int
    total_tokens: int

@app.get("/")
def health_check():
    return {"status": "ok", "stage": "Stage B: Request Queue", "queue_size": request_queue.qsize()}

@app.post("/generate", response_model=GenerateResponse)
async def generate(request: GenerateRequest):
    if not request.prompt.strip():
        raise HTTPException(status_code=400, detail="Prompt cannot be empty.")
    
    arrival_time = time.time()
    loop = asyncio.get_running_loop()
    future = loop.create_future()
    
    item = RequestItem(
        prompt=request.prompt,
        max_new_tokens=request.max_new_tokens,
        future=future,
        arrival_time=arrival_time
    )
    
    # Enqueue request
    await request_queue.put(item)
    
    # Await response from background worker
    try:
        res: InferenceResult = await future
    except Exception as err:
        raise HTTPException(status_code=500, detail=f"Model execution error: {str(err)}")
    
    queue_wait_time = time.time() - arrival_time
    return GenerateResponse(
        prompt=request.prompt,
        generated_text=res.generated_text,
        queue_wait_time_sec=round(queue_wait_time, 4),
        prompt_tokens=res.prompt_tokens,
        generated_tokens=res.generated_tokens,
        total_tokens=res.total_tokens
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8001)
