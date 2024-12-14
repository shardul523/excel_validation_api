import os
import pandas as pd

os.makedirs('uploads/', exist_ok=True)

UPLOAD_PATH = 'uploads/course.xlsx'

EXPECTED_EXCEL_CONFIGURATION = {
    'Course': ['Course ID', 'Course Name'],
    'Topic': ['Topic ID', 'Topic Name', 'Description'],
    'Resource': ['Resource ID', 'Resource Name', 'Resource Content', 'Module ID', 'Module Name', 'Sub Module ID'],
    'Learner': ['Learner ID', 'Name', 'Essay', 'Module ID', 'Submodule ID']
}

def validate_excel_file(file) -> bool:
    '''
    Validates whether the file is a valid excel file or not
    '''
    if not file:
        return False
    
    try:
        filename = file.filename
        file_ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''

        if file_ext != 'xlsx':
            return False
        
    except Exception as e:
        return False

    return True


def validate_course_file(file):
    '''
    Validates whether the file is a valid Course Excel File
    '''
    validation_status = {
        'success': False,
        'errors': ''
    }

    if not validate_excel_file(file):
        validation_status['errors'] = 'Valid Excel File Not Uploaded'
        return validation_status

    file.save(UPLOAD_PATH)

    excelFile = pd.ExcelFile(UPLOAD_PATH)
    sheet_names = excelFile.sheet_names
    missing_sheets = set(EXPECTED_EXCEL_CONFIGURATION.keys()) - set(sheet_names)

    if missing_sheets:
        print(missing_sheets)
        validation_status['errors'] = 'Missing Sheets in the Excel File'
        return validation_status
    
    for sheet_name, sheet_config in EXPECTED_EXCEL_CONFIGURATION.items():
        sheet = pd.read_excel(UPLOAD_PATH, sheet_name=sheet_name)

        if set(sheet_config) - set(sheet.columns):
            validation_status['errors'] = f'Missing Columns in {sheet_name} Sheet'
            return validation_status

        if not len(sheet):
            validation_status['errors'] = f'No Data Present in {sheet_name} Sheet'
            return validation_status
        
        for column in sheet_config:
            if sheet[column].isnull().any():
                validation_status['errors'] = f'Sheet {sheet_name} has rows with missing fields'
                return validation_status
            
    validation_status['success'] = True

    return validation_status