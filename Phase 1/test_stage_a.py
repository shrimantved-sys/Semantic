import requests
import time

SERVER_URL = "http://127.0.0.1:8000"

def test_health():
    print("Testing GET / ...")
    response = requests.get(f"{SERVER_URL}/")
    print(f"Status: {response.status_code}")
    print(f"Response: {response.json()}\n")
    assert response.status_code == 200

def test_generation():
    print("Testing POST /generate ...")
    payload = {
        "prompt": "Artificial intelligence is",
        "max_new_tokens": 15
    }
    start_time = time.time()
    response = requests.post(f"{SERVER_URL}/generate", json=payload)
    elapsed = time.time() - start_time
    
    print(f"Status Code: {response.status_code}")
    print(f"Time Taken : {elapsed:.3f} seconds")
    print(f"Response   : {response.json()}\n")
    assert response.status_code == 200
    data = response.json()
    assert "generated_text" in data
    print("--- Stage A Token Usage ---")
    print(f"  Prompt Tokens    : {data.get('prompt_tokens')}")
    print(f"  Generated Tokens : {data.get('generated_tokens')}")
    print(f"  Total Tokens     : {data.get('total_tokens')}")
    print("---------------------------")
    print("Stage A test passed successfully!")

if __name__ == "__main__":
    try:
        test_health()
        test_generation()
    except requests.exceptions.ConnectionError:
        print("Error: Could not connect to the server at http://127.0.0.1:8000.")
        print("Please ensure server_stage_a.py is running!")
