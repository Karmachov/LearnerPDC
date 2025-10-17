import os
import traceback
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from logic import ReportController # Import the controller from our logic file

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
    """
    Handle the form submission, run the report generation logic,
    and return the generated file.
    """
    try:
        if 'excelFile' not in request.files:
            return jsonify({"error": "No Excel file part in the request."}), 400
        
        excel_file = request.files['excelFile']
        if excel_file.filename == '':
            return jsonify({"error": "No selected Excel file."}), 400

        # Safely get form data
        semester = request.form.get('semester')
        learner_type = request.form.get('learnerType')
        comment = request.form.get('comment')
        slow_thresh_str = request.form.get('slowThreshold')
        fast_thresh_str = request.form.get('fastThreshold')
        format_choice = request.form.get('formatChoice')
        output_type = request.form.get('outputType')

        if not all([semester, learner_type, comment, slow_thresh_str, fast_thresh_str, format_choice, output_type]):
            return jsonify({"error": "Missing form data. Please fill out all fields."}), 400

        slow_thresh = float(slow_thresh_str)
        fast_thresh = float(fast_thresh_str)

        filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_file.save(excel_path)

        controller = ReportController(
            excel_path=excel_path,
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

        if output_path and os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)
        else:
            return jsonify({"error": "No students found for the criteria, or report could not be generated."}), 404

    except ValueError as ve:
        traceback.print_exc()
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"An unexpected server error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)

