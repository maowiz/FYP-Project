from flask import Flask, jsonify
from flask_cors import CORS
import subprocess
import os

app = Flask(__name__)
CORS(app)

@app.route('/start-auth', methods=['GET'])
def start_auth():
    try:
        # Get the directory of the current script
        current_dir = os.path.dirname(os.path.abspath(__file__))
        face_auth_script = os.path.join(current_dir, 'face_auth.py')
        
        # Run the face authentication script
        result = subprocess.run(['python', face_auth_script], capture_output=True, text=True)
        
        # Check if the script ran successfully
        if result.returncode == 0:
            return jsonify({'success': True, 'message': 'Authentication successful'})
        else:
            return jsonify({'success': False, 'message': 'Authentication failed: ' + result.stderr})
            
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(port=8000)
