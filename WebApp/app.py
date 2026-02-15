import os
import traceback
import re
from datetime import datetime
import time
import zipfile
import shutil
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from logic import ReportController

app = Flask(__name__)

# Configure a folder to store temporary uploaded files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def cleanup_uploads(folder, exclude=None):
    """Deletes all files in the upload folder to ensure data privacy."""
    if exclude is None:
        exclude = set()
    else:
        # Normalize paths to absolute for comparison
        exclude = set(os.path.abspath(p) for p in exclude)

    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            # Check if this file is in the exclude list
            if os.path.abspath(file_path) in exclude:
                continue
                
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

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

        # Extract form data
        semester = request.form.get('semester')
        learner_type = request.form.get('learnerType')
        comment = request.form.get('comment')
        faculty_name = request.form.get('facultyName')
        format_choice = request.form.get('formatChoice')
        output_type = request.form.get('outputType')

        # Thresholds are now conditional based on UI selection
        slow_thresh_str = request.form.get('slowThreshold')
        advanced_thresh_str = request.form.get('advancedThreshold')

        # Updated validation: We only check for the core fields. 
        # Thresholds are validated separately to handle the hidden UI state.
        if not all([semester, learner_type, comment, format_choice, output_type]):
            return jsonify({"error": "Missing form data. Please fill out all fields."}), 400

        try:
            # Provide defaults if the UI didn't send a value for the hidden field
            slow_thresh = float(slow_thresh_str) if slow_thresh_str else 40.0
            advanced_thresh = float(advanced_thresh_str) if advanced_thresh_str else 90.0
        except ValueError:
             return jsonify({"error": "Thresholds must be valid numbers."}), 400

        # Save main excel file
        filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_file.save(excel_path)

        # Handle CGPA file (Optional)
        cgpa_path = None
        cgpa_file = request.files.get('cgpaFile')
        if cgpa_file and cgpa_file.filename != '':
            cgpa_filename = secure_filename(f"CGPA_{cgpa_file.filename}")
            cgpa_path = os.path.join(app.config['UPLOAD_FOLDER'], cgpa_filename)
            cgpa_file.save(cgpa_path)

        # Handle Grade File (Optional)
        grade_path = None
        grade_file = request.files.get('gradeFile')
        if grade_file and grade_file.filename != '':
            grade_filename = secure_filename(f"Grade_{grade_file.filename}")
            grade_path = os.path.join(app.config['UPLOAD_FOLDER'], grade_filename)
            grade_file.save(grade_path)

        # Handle Visual Signature
        img_path = None
        img_file = request.files.get('imageFile')
        if img_file and img_file.filename != '':
            img_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(img_file.filename))
            img_file.save(img_path)

        sign_info = {'should_sign': False, 'image_path': img_path}

        # Handle Digital Signing
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
                return jsonify({"error": "Digital Signing is enabled but missing keys."}), 400

        controller = ReportController(
            excel_path=excel_path,
            cgpa_path=cgpa_path,
            grade_path=grade_path,
            format_choice=format_choice,
            learner_type=learner_type,
            slow_thresh=slow_thresh,
            advanced_thresh=advanced_thresh,
            output_type=output_type,
            semester=semester,
            sign_info=sign_info,
            common_comment=comment,
            faculty_name=faculty_name
        )

        output_path = controller.run()
        
        if not output_path:
            return jsonify({"error": "Report generation failed."}), 500

        # Handle ZIP for multiple files
        if isinstance(output_path, list):
            clean_subject = re.sub(r'[^\w\-]', '_', controller.subject).strip('_')
            timestamp = datetime.now().strftime('%d%m%Y_%H%M%S')
            zip_filename = f"{clean_subject}_{semester}_{learner_type.title()}_{timestamp}.zip"
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for fpath in output_path:
                    if os.path.exists(fpath):
                        zf.write(fpath, arcname=os.path.basename(fpath))
            
            response = send_file(os.path.abspath(zip_path), as_attachment=True)
            cleanup_uploads(app.config['UPLOAD_FOLDER'], exclude=[zip_path])
            return response

        # Send single file then cleanup
        response = send_file(os.path.abspath(output_path), as_attachment=True)
        cleanup_uploads(app.config['UPLOAD_FOLDER'], exclude=[output_path])
        return response

    except Exception as e:
        traceback.print_exc()
        cleanup_uploads(app.config['UPLOAD_FOLDER'])
        return jsonify({"error": f"Server error: {str(e)}"}), 500

if __name__ == '__main__':
    # Using port 5001 as per your original configuration
    app.run(host='0.0.0.0', debug=True, port=5001)