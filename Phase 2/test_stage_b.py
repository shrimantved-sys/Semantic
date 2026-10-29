import asyncio
import time
import httpx

SERVER_URL = "http://127.0.0.1:8001"

async def send_request(client: httpx.AsyncClient, req_id: int, prompt: str):
    print(f"[Client] Request #{req_id} sending: '{prompt}'")
    start = time.time()
    response = await client.post(
        f"{SERVER_URL}/generate",
        json={"prompt": prompt, "max_new_tokens": 15},
        timeout=30.0
    )
    total_time = time.time() - start
    data = response.json()
    print(f"[Client] Request #{req_id} completed in {total_time:.3f}s | Queue wait: {data['queue_wait_time_sec']}s")
    print(f"         Prompt Toks: {data.get('prompt_tokens')} | Gen Toks: {data.get('generated_tokens')} | Total Toks: {data.get('total_tokens')}")
    print(f"         Text: {data['generated_text']}\n")
    return data

async def main():
    prompts = [
        "The secret to learning AI is",
        "Python is popular because",
        "In a faraway galaxy",
        "Deep learning models require"
    ]
    
    print("=== Testing Stage B: Decoupled Queue with Concurrent Requests ===")
    async with httpx.AsyncClient() as client:
        # Fire all 4 requests concurrently
        tasks = [send_request(client, i+1, p) for i, p in enumerate(prompts)]
        results = await asyncio.gather(*tasks)
    
    print("=== Stage B Test Summary ===")
    print(f"Total concurrent requests processed: {len(results)}")
    for i, res in enumerate(results, 1):
        print(f" Request #{i}: Queue Wait = {res['queue_wait_time_sec']}s | Tokens: {res.get('prompt_tokens')} prompt + {res.get('generated_tokens')} gen = {res.get('total_tokens')} total")
    print("Stage B test passed successfully!")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except httpx.ConnectError:
        print("Error: Could not connect to server at http://127.0.0.1:8001.")
        print("Make sure server_stage_b.py is running!")
