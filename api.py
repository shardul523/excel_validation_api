from flask import Flask, request, jsonify
from flask_cors import CORS
from utils import validate_course_file

app = Flask(__name__)
CORS(app, resources={r"/upload": {"origins": "*"}})

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'course' not in request.files:
        return jsonify({
            'success': False,
            'errors': 'No file was uploaded'
        }), 400
    
    course_file = request.files['course']

    validation_status = validate_course_file(course_file)

    if not validation_status['success']:
        return jsonify(validation_status), 400

    return jsonify(validation_status), 200


if __name__ == '__main__':
    app.run(debug=True)