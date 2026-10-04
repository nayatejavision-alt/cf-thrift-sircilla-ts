from flask import Flask, request, jsonify
from flask_cors import CORS
import requests

app = Flask(__name__)
CORS(app)

# ఉచిత నిరంతర కనెక్షన్ కోసం
@app.route('/')
def home():
    return jsonify({"message": "Thrift Scheme API is running smoothly!"})

@app.route('/search', methods=['GET'])
def search_beneficiary():
    search_id = request.args.get('id', '').strip()
    if not search_id:
        return jsonify({"error": "ID ఇవ్వు"}), 400
        
    # తర్వాతి స్టెప్‌లో Google Cloud తో కనెక్ట్ చేస్తాం
    return jsonify({"status": "Success", "searched_id": search_id})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
