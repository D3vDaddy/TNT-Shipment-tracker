from flask import Flask, request, jsonify, render_template_string, send_file
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
import io
import uuid
import zipfile
from datetime import datetime

app = Flask(__name__)

# Temporary in-memory storage for generated Excel and ZIP files
generated_files = {}

# --- Helper: Robust Requests Session ---
def get_robust_session():
    """
    Creates a requests Session with automatic retry logic for transient network errors.
    This handles brief DNS failures, connection drops, and API rate limits.
    """
    session = requests.Session()
    # Retry 5 times, increasing the delay between retries automatically (1s, 2s, 4s...)
    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"]
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session

# --- HTML & CSS User Interface ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tracking Automator</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --success: #10b981;
            --success-hover: #059669;
            --accent: #8b5cf6;
            --accent-hover: #7c3aed;
            --bg-dark: #0f172a;
            --card-bg: rgba(255, 255, 255, 0.95);
            --text-main: #1e293b;
            --text-muted: #64748b;
        }

        body { 
            font-family: 'Inter', sans-serif; 
            margin: 0; 
            min-height: 100vh;
            display: flex; 
            justify-content: center; 
            align-items: center;
            background: linear-gradient(-45deg, #ee7752, #e73c7e, #23a6d5, #23d5ab);
            background-size: 400% 400%;
            animation: gradientBG 15s ease infinite;
            padding: 20px;
            box-sizing: border-box;
        }

        @keyframes gradientBG {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        .container { 
            background: var(--card-bg); 
            padding: 40px; 
            border-radius: 24px; 
            box-shadow: 0 25px 50px -12px rgba(0,0,0,0.25), 0 0 0 1px rgba(255,255,255,0.1) inset; 
            width: 100%; 
            max-width: 650px; 
            backdrop-filter: blur(10px);
            position: relative;
            overflow: hidden;
        }

        h2 { 
            color: var(--text-main); 
            margin-top: 0; 
            font-size: 28px;
            font-weight: 700;
            letter-spacing: -0.5px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        p { color: var(--text-muted); font-size: 15px; line-height: 1.6; margin-bottom: 25px; }
        
        textarea { 
            width: 100%; 
            height: 220px; 
            padding: 16px; 
            border: 2px solid #e2e8f0; 
            border-radius: 12px; 
            font-family: 'JetBrains Mono', monospace; 
            font-size: 14px;
            resize: vertical; 
            box-sizing: border-box; 
            transition: all 0.3s ease;
            background: #f8fafc;
            color: #334155;
        }
        textarea:focus { 
            outline: none; 
            border-color: var(--primary); 
            background: #ffffff;
            box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.15); 
        }
        
        button { 
            background-color: var(--primary); 
            color: white; 
            border: none; 
            padding: 16px 24px; 
            font-size: 16px; 
            border-radius: 12px; 
            cursor: pointer; 
            margin-top: 20px; 
            width: 100%; 
            font-weight: 600; 
            transition: all 0.2s ease;
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            box-shadow: 0 4px 6px -1px rgba(79, 70, 229, 0.2), 0 2px 4px -1px rgba(79, 70, 229, 0.1);
        }
        button:hover { 
            background-color: var(--primary-hover); 
            transform: translateY(-2px);
            box-shadow: 0 10px 15px -3px rgba(79, 70, 229, 0.3), 0 4px 6px -2px rgba(79, 70, 229, 0.15);
        }
        button:active { transform: translateY(0); }
        
        button.success-btn { background-color: var(--success); box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.2); }
        button.success-btn:hover { background-color: var(--success-hover); box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.3); }
        
        button.zip-btn { background-color: var(--accent); box-shadow: 0 4px 6px -1px rgba(139, 92, 246, 0.2); }
        button.zip-btn:hover { background-color: var(--accent-hover); box-shadow: 0 10px 15px -3px rgba(139, 92, 246, 0.3); }

        /* Screens */
        .view { display: none; opacity: 0; transition: opacity 0.4s ease; }
        .view.active { display: block; opacity: 1; }
        
        #view-processing { text-align: center; padding: 60px 0; }

        /* Modern Loader */
        .loader-wrapper { position: relative; width: 80px; height: 80px; margin: 0 auto 25px auto; }
        .loader {
            width: 100%; height: 100%;
            border: 4px solid #e0e7ff;
            border-top-color: var(--primary);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        .loader-inner {
            position: absolute; top: 10px; left: 10px; right: 10px; bottom: 10px;
            border: 4px solid transparent;
            border-top-color: var(--accent);
            border-radius: 50%;
            animation: spin 0.75s ease infinite reverse;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

        /* Summary Grid */
        .summary-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 30px; }
        @media(min-width: 500px) { .summary-grid { grid-template-columns: repeat(4, 1fr); } }
        
        .summary-box { 
            background: #ffffff; 
            padding: 20px 15px; 
            border-radius: 16px; 
            text-align: center; 
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
            transition: transform 0.3s ease;
        }
        .summary-box:hover { transform: translateY(-3px); }
        .summary-box.success { border-color: #a7f3d0; background: #ecfdf5; }
        .summary-box.error { border-color: #fecaca; background: #fef2f2; }
        .summary-box.purple { border-color: #ddd6fe; background: #f5f3ff; }
        
        .summary-value { font-size: 32px; font-weight: 800; margin-bottom: 8px; color: var(--text-main); }
        .summary-label { font-size: 12px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }

        /* Background Processing Indicator for ZIP */
        #bg-zip-processing {
            margin-top: 20px;
            padding: 20px;
            background: #f8fafc;
            border-radius: 12px;
            border: 1px dashed #cbd5e1;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
            color: #475569;
            font-weight: 600;
        }
        .small-spinner {
            width: 20px; height: 20px;
            border: 3px solid #e2e8f0; border-top-color: var(--accent);
            border-radius: 50%; animation: spin 1s linear infinite;
        }
        
        .reset-link { 
            display: block; text-align: center; margin-top: 25px; 
            color: var(--text-muted); text-decoration: none; font-size: 15px; font-weight: 600; cursor: pointer; 
            transition: color 0.2s;
        }
        .reset-link:hover { color: var(--primary); text-decoration: underline; }
        
        /* Icons via SVG */
        .icon { width: 20px; height: 20px; fill: currentColor; }
    </style>
</head>
<body>

<div class="container">
    
    <!-- STEP 1: INPUT SCREEN -->
    <div id="view-input" class="view active">
        <h2>
            <svg class="icon" style="width:28px;height:28px;color:var(--primary)" viewBox="0 0 24 24"><path d="M20 7h-4V4c0-1.103-.897-2-2-2h-4c-1.103 0-2 .897-2 2v3H4c-1.103 0-2 .897-2 2v11c0 1.103.897 2 2 2h16c1.103 0 2-.897 2-2V9c0-1.103-.897-2-2-2zM8 4h8v3H8V4zm12 16H4V9h16v11z"/><path d="M11 14h2v2h-2z"/></svg>
            Shipment Automator
        </h2>
        <p>Paste tracking numbers (one per line). We'll instantly map the exact matches to Excel, and extract PODs in the background.</p>
        <textarea id="tracking_numbers" placeholder="107211951&#10;107204448&#10;107005190..." required></textarea>
        <button id="start-btn" onclick="startProcessing()">
            Start Processing
            <svg class="icon" viewBox="0 0 24 24"><path d="M12 2L4.5 20.29l.71.71L12 18l6.79 3 .71-.71z"/></svg>
        </button>
    </div>

    <!-- STEP 2: PROCESSING SCREEN -->
    <div id="view-processing" class="view">
        <div class="loader-wrapper">
            <div class="loader"></div>
            <div class="loader-inner"></div>
        </div>
        <h3 style="margin:0; font-size: 22px; color: var(--text-main);">Fetching Tracker Data...</h3>
        <p style="margin-top:8px; color: var(--text-muted);">Parsing TNT records and generating Excel sheet.</p>
    </div>

    <!-- STEP 3: SUMMARY SCREEN -->
    <div id="view-results" class="view">
        <h2>🎉 Analysis Complete</h2>
        <p>Your Excel file is ready and downloading. PDF retrieval status is below.</p>
        
        <div class="summary-grid">
            <div class="summary-box">
                <div class="summary-value" id="sum-total">0</div>
                <div class="summary-label">Processed</div>
            </div>
            <div class="summary-box success">
                <div class="summary-value" id="sum-found" style="color:var(--success);">0</div>
                <div class="summary-label">Found</div>
            </div>
            <div class="summary-box error">
                <div class="summary-value" id="sum-blank" style="color:#ef4444;">0</div>
                <div class="summary-label">Blank</div>
            </div>
            <div class="summary-box purple">
                <div class="summary-value" id="sum-pods" style="color:var(--accent);">-</div>
                <div class="summary-label">PODs Zipped</div>
            </div>
        </div>

        <!-- Excel Button -->
        <button id="download-btn" class="success-btn" onclick="downloadFile(currentExcelId)">
            <svg class="icon" viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
            Download Excel Again
        </button>
        
        <!-- Background ZIP Processing Loader -->
        <div id="bg-zip-processing" style="display:none;">
            <div class="small-spinner"></div>
            <span>Bypassing security & zipping PODs in background...</span>
        </div>

        <!-- ZIP Button (Hidden initially) -->
        <button id="download-zip-btn" class="zip-btn" onclick="downloadFile(currentZipId)" style="display: none; margin-top: 15px;">
            <svg class="icon" viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
            Download Proof of Deliveries (ZIP)
        </button>
        
        <a class="reset-link" onclick="resetApp()">Process another batch</a>
    </div>

</div>

<script>
    let currentExcelId = null;
    let currentZipId = null;

    function switchView(viewId) {
        document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
        document.getElementById(viewId).classList.add('active');
    }

    async function startProcessing() {
        const text = document.getElementById('tracking_numbers').value.trim();
        if (!text) {
            alert("Please enter at least one tracking number.");
            return;
        }

        // Switch to Processing View
        switchView('view-processing');

        try {
            // -- STAGE 1: EXCEL GENERATION --
            const response = await fetch('/api/process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ tracking_numbers: text })
            });

            const data = await response.json();

            if (response.ok) {
                // Prepare Summary Stats
                document.getElementById('sum-total').innerText = data.summary.total;
                document.getElementById('sum-found').innerText = data.summary.found;
                document.getElementById('sum-blank').innerText = data.summary.not_found;
                currentExcelId = data.file_id;

                // Switch to Results View immediately
                switchView('view-results');
                
                // AUTO-DOWNLOAD EXCEL INSTANTLY
                setTimeout(() => { downloadFile(currentExcelId); }, 500);

                // -- STAGE 2: PDF ZIP GENERATION (Background) --
                if (data.valid_shipments && data.valid_shipments.length > 0) {
                    
                    document.getElementById('bg-zip-processing').style.display = 'flex';
                    document.getElementById('download-zip-btn').style.display = 'none';
                    document.getElementById('sum-pods').innerText = "...";

                    const podResp = await fetch('/api/generate_pod_zip', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ shipments: data.valid_shipments })
                    });
                    
                    const podData = await podResp.json();
                    
                    // Hide loader
                    document.getElementById('bg-zip-processing').style.display = 'none';

                    if (podData.zip_id) {
                        currentZipId = podData.zip_id;
                        document.getElementById('sum-pods').innerText = podData.count;
                        document.getElementById('download-zip-btn').style.display = 'flex'; // Use flex to keep icon alignment
                        
                        // Optional auto-download for ZIP too:
                        // setTimeout(() => { downloadFile(currentZipId); }, 500);
                    } else {
                        document.getElementById('sum-pods').innerText = "0";
                    }
                } else {
                    document.getElementById('sum-pods').innerText = "0";
                }
            } else {
                alert("Error processing data: " + data.error);
                resetApp();
            }
        } catch (error) {
            alert("Network error occurred or Server did not respond. Please try again.");
            resetApp();
        }
    }

    function downloadFile(fileId) {
        if (fileId) {
            window.location.href = '/download/' + fileId;
        }
    }

    function resetApp() {
        document.getElementById('tracking_numbers').value = '';
        currentExcelId = null;
        currentZipId = null;
        document.getElementById('bg-zip-processing').style.display = 'none';
        document.getElementById('download-zip-btn').style.display = 'none';
        document.getElementById('sum-pods').innerText = "-";
        switchView('view-input');
    }
</script>

</body>
</html>
"""

def parse_iso_date(date_string):
    """Converts ISO date to EXACTLY DD/MM/YYYY format."""
    if not date_string:
        return ""
    try:
        dt = datetime.fromisoformat(date_string)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return date_string

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/process', methods=['POST'])
def process_api():
    try:
        data = request.json
        raw_input = data.get('tracking_numbers', '')
        tracking_list = [t.strip() for t in raw_input.split('\n') if t.strip()]
        
        if not tracking_list:
            return jsonify({"error": "No numbers provided"}), 400

        # Setup strict dictionary preserving exact order
        results_dict = {
            trk: {
                "Tracking Number": trk,
                "Status": "",
                "Delivery Date": "",
                "Customer Ref": "",
                "_found": False
            } for trk in tracking_list
        }

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://www.tnt.com/"
        }

        # Use our robust session with retries built-in
        session = get_robust_session()

        for batch in chunks(tracking_list, 30):
            cons_param = ",".join(batch)
            url = f"https://www.tnt.com/api/v3/shipment?con={cons_param}&locale=en_GB&searchType=CON&channel=OPENTRACK"
            
            # Using session instead of generic requests.get + added timeout
            response = session.get(url, headers=headers, timeout=20)
            
            if response.status_code == 200:
                resp_data = response.json()
                
                # Check Consignments - ONLY EXACT MATCHES ALLOWED
                consignments = resp_data.get('tracker.output', {}).get('consignment', [])
                for item in consignments:
                    con_num = item.get('consignmentNumber')
                    
                    if con_num in results_dict:
                        results_dict[con_num]['_found'] = True
                        
                        # Store API data securely for Stage 2 (PDF fetching)
                        results_dict[con_num]['_api_data'] = {
                            "conNumber": con_num,
                            "consignmentKey": item.get('consignmentKey', ''),
                            "shipmentId": item.get('shipmentId', '')
                        }

                        # Extract Status
                        if item.get('status', {}).get('isDelivered', False):
                            results_dict[con_num]['Status'] = "Delivered"
                        else:
                            events = item.get('events', [])
                            if events:
                                results_dict[con_num]['Status'] = events[0].get('statusDescription', 'In Transit')
                        
                        # Extract Date
                        analytics = item.get('analytics', {})
                        dest_dates = analytics.get('destinationDateSources', {})
                        del_date_raw = dest_dates.get('delivered')
                        if not del_date_raw and item.get('events'):
                            del_date_raw = item['events'][0].get('date')
                        results_dict[con_num]['Delivery Date'] = parse_iso_date(del_date_raw)
                        
                        # Extract Ref
                        results_dict[con_num]['Customer Ref'] = item.get('customerReference', '')

        # Calculate Summary and prepare lists
        total = len(tracking_list)
        found = 0
        valid_shipments = []
        final_data = []

        for trk in tracking_list:
            row = results_dict[trk].copy()
            if row['_found']:
                found += 1
            if '_api_data' in row:
                valid_shipments.append(row['_api_data'])
                del row['_api_data']
            del row['_found']
            final_data.append(row)

        not_found = total - found

        # Generate Excel File in memory
        df = pd.DataFrame(final_data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Tracking Results')
            worksheet = writer.sheets['Tracking Results']
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = max_len

        # Store file in memory
        file_id = str(uuid.uuid4())
        output.seek(0)
        generated_files[file_id] = {'data': output, 'type': 'excel'}

        return jsonify({
            "summary": {
                "total": total,
                "found": found,
                "not_found": not_found
            },
            "file_id": file_id,
            "valid_shipments": valid_shipments
        })

    except requests.exceptions.ConnectionError as ce:
        print(f"Network Error: {str(ce)}")
        return jsonify({"error": "Failed to connect to TNT. Please check your network/VPN connection. DNS resolution failed."}), 503
    except requests.exceptions.Timeout as te:
        print(f"Timeout Error: {str(te)}")
        return jsonify({"error": "The connection to TNT timed out. The server might be busy."}), 504
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({"error": "An unexpected error occurred: " + str(e)}), 500


@app.route('/api/generate_pod_zip', methods=['POST'])
def generate_pod_zip():
    try:
        shipments = request.json.get('shipments', [])
        if not shipments:
            return jsonify({"zip_id": None, "count": 0})

        zip_buffer = io.BytesIO()
        pdf_count = 0
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://www.tnt.com/"
        }

        # Use our robust session with retries built-in
        session = get_robust_session()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for ship in shipments:
                con = ship.get('conNumber')
                key = ship.get('consignmentKey')
                sid = ship.get('shipmentId')
                
                if not (con and key and sid):
                    continue
                    
                # Helper function to attempt download with a specific account number
                def try_account(acc):
                    params = {
                        "conNumber": con,
                        "consignmentKey": key,
                        "securityQuestionType": "accountNumber",
                        "securityQuestionValue": acc,
                        "shipmentId": sid
                    }
                    url = "https://www.tnt.com/api/v1/shipment/confidentialDetails"
                    try:
                        # Using robust session
                        resp = session.get(url, params=params, headers=headers, timeout=15)
                        if resp.status_code == 200:
                            data = resp.json()
                            if "error" not in data.get("confidentialDetailsOutput", {}):
                                return data.get("confidentialDetailsOutput", {}).get("confidentialData", {}).get("podUrl")
                    except Exception as e:
                        print(f"Auth check failed for {con} with acc {acc}: {e}")
                    return None

                # Logic strictly matching the user request
                pod_url = try_account("0000196665")
                if not pod_url:
                    pod_url = try_account("000191600")
                    
                # If a valid PDF URL was extracted, follow it and save to zip
                if pod_url:
                    try:
                        # Using robust session for PDF fetch
                        pdf_resp = session.get(pod_url, headers=headers, allow_redirects=True, timeout=20)
                        if pdf_resp.status_code == 200:
                            zip_file.writestr(f"{con}.pdf", pdf_resp.content)
                            pdf_count += 1
                    except Exception as e:
                        print(f"Failed to fetch PDF for {con}: {e}")
                        
        if pdf_count > 0:
            zip_id = str(uuid.uuid4())
            zip_buffer.seek(0)
            generated_files[zip_id] = {'data': zip_buffer, 'type': 'zip'}
            return jsonify({"zip_id": zip_id, "count": pdf_count})
        else:
            return jsonify({"zip_id": None, "count": 0})

    except Exception as e:
        print(f"Error zipping PODs: {str(e)}")
        return jsonify({"zip_id": None, "count": 0, "error": str(e)})


@app.route('/download/<file_id>', methods=['GET'])
def download_file(file_id):
    if file_id in generated_files:
        file_info = generated_files[file_id]
        file_data = file_info['data']
        file_type = file_info['type']
        
        file_data.seek(0) # Ensure we read from start of bytes
        
        if file_type == 'excel':
            filename = f"Tracking_Results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            return send_file(
                file_data,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=filename
            )
        elif file_type == 'zip':
            filename = f"Proof_of_Delivery_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
            return send_file(
                file_data,
                mimetype="application/zip",
                as_attachment=True,
                download_name=filename
            )
            
    return "File not found or expired.", 404


if __name__ == '__main__':
    # Run the app on port 6767
    app.run(host='0.0.0.0', port=6767, debug=True)