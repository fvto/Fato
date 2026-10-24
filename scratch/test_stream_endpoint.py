import urllib.request
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = "http://localhost:5050/api/command/stream"
payload = json.dumps({"prompt": "Chọn nhà máy VH3 và lọc chuyền VL - D3 ngày hôm nay rồi tìm kiếm"}).encode('utf-8')
req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})

print("--- TESTING STREAMING SSE ENDPOINT ---")
try:
    with urllib.request.urlopen(req, timeout=60) as resp:
        print(f"Status: {resp.status}, Content-Type: {resp.headers.get('Content-Type')}")
        for raw_line in resp:
            line = raw_line.decode('utf-8').strip()
            if line.startswith("data: "):
                data = json.loads(line[6:])
                print(f"STEP {data.get('step')} | Status: {data.get('status')} | Msg: {data.get('msg')}")
                if data.get('intent'):
                    print("  Parsed intent:", data['intent']['summary'])
    print("--- STREAM COMPLETED SUCCESSFULLY! ---")
except Exception as e:
    print("Error:", e)
