import json
import time
from flask import Flask, Response, request, stream_with_context

app = Flask(__name__)

@app.route('/test_stream', methods=['POST'])
def test_stream():
    data = request.get_json() or {}
    prompt = data.get('prompt', '')

    def generate():
        yield f"data: {json.dumps({'step': 1, 'msg': f'Đã nhận câu lệnh: \"{prompt}\"'})}\n\n"
        time.sleep(0.5)
        yield f"data: {json.dumps({'step': 2, 'msg': 'Đang phân tích Nhà máy, Chuyền, Ngày...'})}\n\n"
        time.sleep(0.5)
        yield f"data: {json.dumps({'step': 3, 'msg': 'Đang tương tác với trình duyệt...'})}\n\n"
        time.sleep(0.5)
        yield f"data: {json.dumps({'step': 4, 'msg': 'Hoàn thành!', 'done': True})}\n\n"

    return Response(stream_with_context(generate()), mimetype='text/event-stream')

if __name__ == '__main__':
    # run quick check
    print("Test SSE generator defined successfully.")
