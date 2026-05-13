from flask import Flask, jsonify
import time

app = Flask(__name__)

@app.route('/')
def home():
    return jsonify({"status": "working", "time": time.time()})

@app.route('/test')
def test():
    return jsonify({"message": "Test endpoint working"})

if __name__ == '__main__':
    print("Starting test server on http://localhost:5004")
    app.run(debug=True, port=5004)
