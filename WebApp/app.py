import os
import traceback
import time
import zipfile
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from logic import ReportController

app = Flask(__name__)

# Configure a folder to store temporary uploaded files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    """Render the main HTML page."""
    return render_template('index.html')

@app.route('/generate-report', methods=['POST'])
def generate_report():
    try:
        
        if 'excelFile' not in request.files:
            return jsonify({"error": "No Excel file part in the request."}), 400
        
        excel_file = request.files['excelFile']
        if excel_file.filename == '':
            return jsonify({"error": "No selected Excel file."}), 400

    
        semester = request.form.get('semester')
        learner_type = request.form.get('learnerType')
        comment = request.form.get('comment')
        faculty_name = request.form.get('facultyName')
        slow_thresh_str = request.form.get('slowThreshold')
        fast_thresh_str = request.form.get('fastThreshold')
        format_choice = request.form.get('formatChoice')
        output_type = request.form.get('outputType')

        if not all([semester, learner_type, comment, slow_thresh_str, fast_thresh_str, format_choice, output_type]):
            return jsonify({"error": "Missing form data. Please fill out all fields."}), 400

        try:
            slow_thresh = float(slow_thresh_str)
            fast_thresh = float(fast_thresh_str)
        except ValueError:
             return jsonify({"error": "Thresholds must be valid numbers."}), 400

    
        filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_file.save(excel_path)

        
        cgpa_path = None
        cgpa_file = request.files.get('cgpaFile')
        if cgpa_file and cgpa_file.filename != '':
            cgpa_filename = secure_filename(f"CGPA_{cgpa_file.filename}")
            cgpa_path = os.path.join(app.config['UPLOAD_FOLDER'], cgpa_filename)
            cgpa_file.save(cgpa_path)

        
        img_path = None
        img_file = request.files.get('imageFile')
        if img_file and img_file.filename != '':
            img_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(img_file.filename))
            img_file.save(img_path)

        
        sign_info = {
            'should_sign': False, 
            'image_path': img_path 
        }

        
        if request.form.get('enableSigning') == 'on':
            key_file = request.files.get('keyFile')
            cert_file = request.files.get('certFile')
            password = request.form.get('keyPassword')

            if all([key_file, cert_file]):
                key_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(key_file.filename))
                cert_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(cert_file.filename))
                
                key_file.save(key_path)
                cert_file.save(cert_path)

                sign_info.update({
                    'should_sign': True,
                    'key_path': key_path,
                    'cert_path': cert_path,
                    'password': password
                })
            else:
                return jsonify({"error": "Digital Signing is enabled but missing keys or certificates."}), 400

        
        controller = ReportController(
            excel_path=excel_path,
            cgpa_path=cgpa_path,
            format_choice=format_choice,
            learner_type=learner_type,
            slow_thresh=slow_thresh,
            fast_thresh=fast_thresh,
            output_type=output_type,
            semester=semester,
            sign_info={'should_sign': False}, # Signing is not supported in this UI
            common_comment=comment
        )


        
        output_path = controller.run()
        
        
        if isinstance(output_path, list):
            output_path = [os.path.abspath(p) for p in output_path]
        else:
            output_path = os.path.abspath(output_path) if output_path else None

        
        if isinstance(output_path, list):
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], f"reports_{semester}_{learner_type}_{int(time.time())}.zip")
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for fpath in output_path:
                    if os.path.exists(fpath):
                        zf.write(fpath, arcname=os.path.basename(fpath))
            return send_file(os.path.abspath(zip_path), as_attachment=True)

        
        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
            return jsonify({"error": "Report generation failed."}), 500

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"An unexpected server error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)