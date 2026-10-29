from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

app = FastAPI(
    title="LLM Inference Server - Stage A",
    description="Basic single-endpoint server with synchronous CPU model execution."
)

MODEL_NAME = "gpt2"
print(f"[Stage A] Loading model '{MODEL_NAME}' onto CPU...")

tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
if tokenizer.pad_token is None:
    tokenizer.pad_token = tokenizer.eos_token

model = AutoModelForCausalLM.from_pretrained(MODEL_NAME)
model.eval()
print("[Stage A] Model loaded successfully and ready for inference!")

class GenerateRequest(BaseModel):
    prompt: str = Field(..., description="Input prompt text for text generation", min_length=1)
    max_new_tokens: int = Field(default=20, ge=1, le=100, description="Number of tokens to generate")

class GenerateResponse(BaseModel):
    prompt: str
    generated_text: str
    prompt_tokens: int
    generated_tokens: int
    total_tokens: int

@app.get("/")
def health_check():
    return {"status": "ok", "stage": "Stage A: Basic Server"}

@app.post("/generate", response_model=GenerateResponse)
def generate(request: GenerateRequest):
    if not request.prompt.strip():
        raise HTTPException(status_code=400, detail="Prompt cannot be empty.")
    
    # 1. Tokenize prompt input
    inputs = tokenizer(request.prompt, return_tensors="pt")
    prompt_token_count = int(inputs["input_ids"].shape[1])
    
    # 2. Run model generation synchronously
    with torch.no_grad():
        outputs = model.generate(
            **inputs,
            max_new_tokens=request.max_new_tokens,
            pad_token_id=tokenizer.eos_token_id,
            do_sample=False  # Greedy decoding for deterministic results
        )
    
    # 3. Decode output tokens back into string
    generated_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
    total_token_count = len(outputs[0])
    generated_token_count = total_token_count - prompt_token_count
    
    return GenerateResponse(
        prompt=request.prompt,
        generated_text=generated_text,
        prompt_tokens=prompt_token_count,
        generated_tokens=generated_token_count,
        total_tokens=total_token_count
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
