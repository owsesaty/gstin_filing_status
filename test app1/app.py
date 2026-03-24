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

    resp_data, status_code = fetch_gstin_data(gstin, fy)
    return jsonify(resp_data), status_code

def fetch_gstin_data(gstin, fy):
    url = 'https://services.gst.gov.in/services/api/search/taxpayerReturnDetails'
    payload = {'gstin': gstin, 'fy': fy}
    headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=15)
        try:
            return response.json(), response.status_code
        except ValueError:
            return {"error": "Invalid API response format", "response_text": response.text[:300]}, response.status_code
    except requests.exceptions.RequestException as e:
        return {'error': str(e)}, 500

@app.route('/api/search_bulk', methods=['POST'])
def proxy_search_bulk():
    import concurrent.futures

    data = request.json
    gstins = data.get('gstins', [])
    fy = data.get('fy', '2025')
    
    if not gstins or not isinstance(gstins, list):
        return jsonify({"error": "A list of GSTINs is required"}), 400

    results = {}
    
    # Use ThreadPoolExecutor for concurrent requests safely handled by Python backend
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
        future_to_gstin = {executor.submit(fetch_gstin_data, gstin, fy): gstin for gstin in gstins}
        for future in concurrent.futures.as_completed(future_to_gstin):
            gstin = future_to_gstin[future]
            try:
                resp_data, status_code = future.result()
                if status_code == 200:
                    results[gstin] = resp_data
                else:
                    results[gstin] = {"error": f"Server responded with {status_code}", "details": resp_data}
            except Exception as exc:
                results[gstin] = {"error": str(exc)}

    return jsonify({"results": results}), 200

if __name__ == '__main__':
    print("Starting GSTIN Viewer server...")
    print("--> Open http://localhost:5000 in your browser")
    app.run(port=5000, debug=True)
