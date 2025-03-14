from flask import Flask, request, jsonify
import subprocess

app = Flask(__name__)

@app.route('/run-python', methods=['POST'])
def run_python():
    try:
        # fetch_confluence.py を実行
        subprocess.run(["python3", "fetch_confluence.py"], check=True)
        return jsonify({"status": "success", "message": "Python script executed successfully"})
    except subprocess.CalledProcessError as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
