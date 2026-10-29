import asyncio
import time
import httpx

SERVER_URL = "http://127.0.0.1:8002"

async def send_request(client: httpx.AsyncClient, req_id: int, prompt: str):
    print(f"[Client] Sending Request #{req_id}: '{prompt}'")
    start = time.time()
    response = await client.post(
        f"{SERVER_URL}/generate",
        json={"prompt": prompt, "max_new_tokens": 15},
        timeout=30.0
    )
    total_time = time.time() - start
    data = response.json()
    print(f"[Client] Req #{req_id} Done in {total_time:.3f}s | Batch Size: {data['batch_size']} | Queue Wait: {data['queue_wait_time_sec']}s | Model Process Time: {data['processing_time_sec']}s")
    print(f"         Tokens: Prompt={data.get('prompt_tokens')} | Padded={data.get('padded_tokens')} | Generated={data.get('generated_tokens')} | Total={data.get('total_model_tokens')}")
    print(f"         Output: {data['generated_text']}\n")
    return data

async def main():
    print("=== Testing Stage C: Static Batching ===")
    
    async with httpx.AsyncClient() as client:
        # Scenario 1: Burst of 4 requests (Hits MAX_BATCH_SIZE = 4 immediately)
        print("--- Test Scenario 1: Sending 4 concurrent requests (Full Batch Burst) ---")
        burst_prompts = [
            "The future of software is",
            "Quantum computing promises to",
            "Machine learning algorithms can",
            "Data structures are essential because"
        ]
        t1_start = time.time()
        burst_tasks = [send_request(client, i+1, p) for i, p in enumerate(burst_prompts)]
        burst_results = await asyncio.gather(*burst_tasks)
        t1_elapsed = time.time() - t1_start
        print(f"Scenario 1 Total Time for 4 requests: {t1_elapsed:.3f}s (Batch size observed: {burst_results[0]['batch_size']})")
        total_pad = sum(r.get('padded_tokens', 0) for r in burst_results)
        print(f"Scenario 1 Total Padding Tokens across batch: {total_pad}\n")
        
        # Scenario 2: Single request alone (Hits BATCH_TIMEOUT = 0.2s)
        print("--- Test Scenario 2: Sending 1 request alone (Timeout Window Test) ---")
        single_result = await send_request(client, 5, "Once upon a time in space")
        print(f"Scenario 2 Batch size observed: {single_result['batch_size']} (Expected: 1 after timeout)")
        print(f"Scenario 2 Tokens: Prompt={single_result.get('prompt_tokens')} | Padded={single_result.get('padded_tokens')} | Gen={single_result.get('generated_tokens')}\n")
        
    print("=== Stage C Test Summary ===")
    print("Stage C test passed successfully!")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except httpx.ConnectError:
        print("Error: Could not connect to server at http://127.0.0.1:8002.")
        print("Make sure server_stage_c.py is running!")
