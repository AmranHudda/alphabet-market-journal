import os
from flask import Flask, request, jsonify
import sqlite3
import uuid
import logging
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

app = Flask(__name__)

# Set up rate limiting
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"]
)

# Set up logging
logging.basicConfig(level=logging.INFO)

@app.route('/test', methods=['GET'])
def test():
    return jsonify({"message": "Server is working"}), 200

# Database setup
def init_db():
    conn = sqlite3.connect('licenses.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS licenses
                 (license_key TEXT PRIMARY KEY, device_id TEXT)''')
    conn.commit()
    conn.close()

init_db()

@app.route('/api/validate_license', methods=['POST'])
@limiter.limit("10/minute")
def validate_license():
    data = request.json
    license_key = data.get('license_key')
    device_id = data.get('device_id')

    logging.info(f"Validating license: {license_key} for device: {device_id}")

    if not license_key or not device_id:
        logging.warning("Missing license key or device ID")
        return jsonify({"valid": False, "message": "Missing license key or device ID"}), 400

    conn = sqlite3.connect('licenses.db')
    c = conn.cursor()
    c.execute("SELECT device_id FROM licenses WHERE license_key = ?", (license_key,))
    result = c.fetchone()

    if result is None:
        conn.close()
        logging.warning(f"Invalid license key attempt: {license_key}")
        return jsonify({"valid": False, "message": "Invalid license key"}), 200

    stored_device_id = result[0]

    if stored_device_id is None:
        # First time use, associate the device ID with the license key
        c.execute("UPDATE licenses SET device_id = ? WHERE license_key = ?", (device_id, license_key))
        conn.commit()
        conn.close()
        logging.info(f"License key activated: {license_key} for device: {device_id}")
        return jsonify({"valid": True, "message": "License key activated"}), 200
    elif stored_device_id == device_id:
        conn.close()
        logging.info(f"Valid license key used: {license_key} for device: {device_id}")
        return jsonify({"valid": True, "message": "License key valid"}), 200
    else:
        conn.close()
        logging.warning(f"License key reuse attempt: {license_key} for device: {device_id}")
        return jsonify({"valid": False, "message": "License key already in use on another device"}), 200
    
@app.route('/api/list_licenses', methods=['GET'])
def list_licenses():
    conn = sqlite3.connect('licenses.db')
    c = conn.cursor()
    c.execute("SELECT license_key, device_id FROM licenses")
    licenses = c.fetchall()
    conn.close()
    return jsonify({"licenses": [{"license_key": l[0], "device_id": l[1]} for l in licenses]}), 200

@app.route('/api/generate_license', methods=['POST'])
@limiter.limit("5/minute")
def generate_license():
    # In a real-world scenario, you would add authentication here
    # to ensure only authorized users can generate license keys
    new_license = str(uuid.uuid4())
    conn = sqlite3.connect('licenses.db')
    c = conn.cursor()
    c.execute("INSERT INTO licenses (license_key) VALUES (?)", (new_license,))
    conn.commit()
    conn.close()
    logging.info(f"New license key generated: {new_license}")
    return jsonify({"license_key": new_license}), 201

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=True, host='0.0.0.0', port=port)