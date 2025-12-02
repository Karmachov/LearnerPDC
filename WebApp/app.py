import os
import traceback
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from logic import ReportController
import time

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    """Render the main HTML page."""
    return render_template('index.html')

@app.route('/generate-report', methods=['POST'])
def generate_report():
    """
    Handle the form submission, run the report generation logic,
    and return the generated file.
    """
    try:
        # 1. Basic File Validation
        if 'excelFile' not in request.files:
            return jsonify({"error": "No Excel file part in the request."}), 400
        
        excel_file = request.files['excelFile']
        if excel_file.filename == '':
            return jsonify({"error": "No selected Excel file."}), 400

        # 2. Retrieve Basic Form Data
        semester = request.form.get('semester')
        learner_type = request.form.get('learnerType')
        comment = request.form.get('comment')
        slow_thresh_str = request.form.get('slowThreshold')
        fast_thresh_str = request.form.get('fastThreshold')
        format_choice = request.form.get('formatChoice')
        output_type = request.form.get('outputType')

        # Validate required text fields
        if not all([semester, learner_type, comment, slow_thresh_str, fast_thresh_str, format_choice, output_type]):
            return jsonify({"error": "Missing form data. Please fill out all fields."}), 400

        # Convert numerical thresholds
        try:
            slow_thresh = float(slow_thresh_str)
            fast_thresh = float(fast_thresh_str)
        except ValueError:
             return jsonify({"error": "Thresholds must be valid numbers."}), 400

        # 3. Save the Excel File
        filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_file.save(excel_path)

        # 4. Handle Digital Signature Logic
        sign_info = {'should_sign': False}
        
        # Check if signing is enabled (checkbox in frontend sends 'on' if checked)
        if request.form.get('enableSigning') == 'on':
            # Retrieve signature assets
            key_file = request.files.get('keyFile')
            cert_file = request.files.get('certFile')
            img_file = request.files.get('imageFile')
            password = request.form.get('keyPassword')

            # Ensure all signing components are present
            if all([key_file, cert_file, img_file, password]):
                # Save assets securely
                key_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(key_file.filename))
                cert_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(cert_file.filename))
                img_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(img_file.filename))
                
                key_file.save(key_path)
                cert_file.save(cert_path)
                img_file.save(img_path)

                # Update sign_info dictionary
                sign_info = {
                    'should_sign': True,
                    'key_path': key_path,
                    'cert_path': cert_path,
                    'image_path': img_path,
                    'password': password
                }
            else:
                # If user checked the box but missed a file, you might want to warn them.
                # For now, we return an error to ensure they provide everything.
                return jsonify({"error": "Digital Signing is enabled but missing keys, certificates, or password."}), 400

        # 5. Initialize Controller
        controller = ReportController(
            excel_path=excel_path,
            format_choice=format_choice,
            learner_type=learner_type,
            slow_thresh=slow_thresh,
            fast_thresh=fast_thresh,
            output_type=output_type,
            semester=semester,
            sign_info=sign_info,  # Pass the dynamic sign_info here
            common_comment=comment
        )

        # 6. Run Generation
        output_path = controller.run()
        # Normalize controller output
        if isinstance(output_path, list):
            output_path = [os.path.abspath(p) for p in output_path]
        else:
            output_path = os.path.abspath(output_path) if output_path else None

        print("DEBUG: normalized output paths:", output_path)

# If list -> zip them
        if isinstance(output_path, list):
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], f"reports_{semester}_{learner_type}_{int(time.time())}.zip")
            import zipfile
            with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                for fpath in output_path:
                    if os.path.exists(fpath):
                        zf.write(fpath, arcname=os.path.basename(fpath))
                    else:
                        print("WARNING: expected file missing when zipping:", fpath)
            return send_file(os.path.abspath(zip_path), as_attachment=True)


# single file
        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
    # helpful debugging listing
            parent = os.path.dirname(output_path) if output_path else app.config['UPLOAD_FOLDER']
            listing = []
            try:
                listing = os.listdir(parent)
            except Exception as _:
                listing = ["(couldn't list parent dir)"]
            return jsonify({
                "error": "file_not_found",
                "expected_path": output_path,
                "parent_listing_sample": listing[:50]
            }), 500

    except ValueError as ve:
        traceback.print_exc()
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"An unexpected server error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001, threaded=False)