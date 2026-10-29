import asyncio
import time
import httpx

SERVER_URL = "http://127.0.0.1:8003"

async def test_cache():
    print("=== Testing Stage D: Exact-Match Cache ===")
    prompt = "What is the capital of France?"
    
    async with httpx.AsyncClient() as client:
        # Request 1: First time sending prompt (Cache MISS expected)
        print("\n--- Request 1 (First attempt - Cache MISS expected) ---")
        t1_start = time.time()
        res1 = await client.post(f"{SERVER_URL}/generate", json={"prompt": prompt, "max_new_tokens": 15})
        t1_elapsed = time.time() - t1_start
        data1 = res1.json()
        print(f"Req #1 Total Time: {t1_elapsed:.4f}s | Cache Hit: {data1['cache_hit']} | Queue Wait: {data1['queue_wait_time_sec']}s")
        print(f"       Tokens: Prompt={data1.get('prompt_tokens')} | Padded={data1.get('padded_tokens')} | Generated={data1.get('generated_tokens')} | Total={data1.get('total_model_tokens')}")
        print(f"       Text: {data1['generated_text']}")
        assert data1['cache_hit'] is False, "Request 1 should be a Cache MISS"

        # Request 2: Exact same prompt (Cache HIT expected)
        print("\n--- Request 2 (Identical prompt - Cache HIT expected) ---")
        t2_start = time.time()
        res2 = await client.post(f"{SERVER_URL}/generate", json={"prompt": prompt, "max_new_tokens": 15})
        t2_elapsed = time.time() - t2_start
        data2 = res2.json()
        print(f"Req #2 Total Time: {t2_elapsed:.4f}s | Cache Hit: {data2['cache_hit']} | Queue Wait: {data2['queue_wait_time_sec']}s")
        print(f"       Tokens: Prompt={data2.get('prompt_tokens')} | Padded={data2.get('padded_tokens')} | Generated={data2.get('generated_tokens')} | Total={data2.get('total_model_tokens')}")
        print(f"       Text: {data2['generated_text']}")
        assert data2['cache_hit'] is True, "Request 2 should be a Cache HIT"
        assert data2['generated_text'] == data1['generated_text'], "Cached response text must match original response text"

        # Request 3: Slight whitespace variation (Should be normalized and HIT)
        print("\n--- Request 3 (Prompt with leading/trailing spaces - Normalized Cache HIT expected) ---")
        t3_start = time.time()
        res3 = await client.post(f"{SERVER_URL}/generate", json={"prompt": "  What is the capital of France?  ", "max_new_tokens": 15})
        t3_elapsed = time.time() - t3_start
        data3 = res3.json()
        print(f"Req #3 Total Time: {t3_elapsed:.4f}s | Cache Hit: {data3['cache_hit']} | Tokens: Prompt={data3.get('prompt_tokens')} | Gen={data3.get('generated_tokens')}")
        assert data3['cache_hit'] is True, "Request 3 should be a Cache HIT after string normalization"

    print("\n=== Stage D Test Summary ===")
    print(f"Initial request latency (MISS): {t1_elapsed:.4f}s")
    print(f"Cached request latency (HIT) : {t2_elapsed:.4f}s")
    speedup = t1_elapsed / max(t2_elapsed, 0.0001)
    print(f"Cache Speedup: {speedup:.1f}x faster!")
    print("Stage D test passed successfully!")

if __name__ == "__main__":
    try:
        asyncio.run(test_cache())
    except httpx.ConnectError:
        print("Error: Could not connect to server at http://127.0.0.1:8003.")
        print("Make sure server_stage_d.py is running!")
