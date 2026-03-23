from flask import Flask, request, jsonify, send_from_directory 
import requests
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
public_dir = os.path.join(BASE_DIR, 'public')

app = Flask(__name__, static_folder=public_dir, static_url_path='')

@app.route('/')
def serve_index():
    return send_from_directory(public_dir, 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    if os.path.exists(os.path.join(public_dir, path)):
        return send_from_directory(public_dir, path)
    return "Not found", 404

@app.route('/api/search', methods=['POST'])
def proxy_search():
    data = request.json
    gstin = data.get('gstin')
    fy = data.get('fy', '2025')
    
    if not gstin:
        return jsonify({"error": "GSTIN is required"}), 400

    url = 'https://services.gst.gov.in/services/api/search/taxpayerReturnDetails'
    payload = {'gstin': gstin, 'fy': fy}
    
    headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }
    
    try:
        response = requests.post(url, json=payload, headers=headers)
        
        try:
            resp_data = response.json()
        except ValueError:
            resp_data = {"error": "Invalid API response format", "response_text": response.text[:300]}

        return jsonify(resp_data), response.status_code
    except requests.exceptions.RequestException as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("Starting GSTIN Viewer server...")
    print("--> Open http://localhost:5000 in your browser")
    app.run(port=5000, debug=True)
