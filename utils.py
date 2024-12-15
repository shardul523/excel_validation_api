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

required_sheet_names = list(EXPECTED_EXCEL_CONFIGURATION.keys())


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


def validate_sheets():
    excel_file = pd.ExcelFile(UPLOAD_PATH)
    sheet_names = excel_file.sheet_names
    print(sheet_names, required_sheet_names)

    if len(sheet_names) != len(required_sheet_names):
        return False
    
    return sheet_names == required_sheet_names


def validate_sheet_format(sheet: pd.DataFrame, sheet_name: str) -> bool:
    required_columns = EXPECTED_EXCEL_CONFIGURATION[sheet_name]
    columns = sheet.columns.to_list()

    if len(columns) != len(required_columns):
        return False
    
    return columns == required_columns


def validate_sheet_data(sheet: pd.DataFrame):
    for column in sheet.columns:
        if sheet[column].isnull().any():
            return False
    
    return True


def validate_course_file(file):
    '''
    Validates whether the file is a valid Course Excel File
    '''
    validation_status = {
        'success': False,
        'errors': []
    }

    try:
        if not validate_excel_file(file):
            validation_status['errors'].append('Valid Excel File Not Uploaded')
            return validation_status

        file.save(UPLOAD_PATH)

        if not validate_sheets():
            validation_status['errors'].append('Missing / Extra Sheets in the Excel file')
            return validation_status
        

        for sheet_name in required_sheet_names:
            sheet = pd.read_excel(UPLOAD_PATH, sheet_name=sheet_name)

            if not validate_sheet_format(sheet, sheet_name):
                validation_status['errors'].append(f'Missing / Extra Columns in {sheet_name} Sheet')
            
            if not len(sheet):
                validation_status['errors'].append(f'Data is Missing from {sheet_name} Sheet')
            
            if not validate_sheet_data(sheet):
                validation_status['errors'].append(f'{sheet_name} Sheet has rows with missing columns')
        
        validation_status['success'] = True

    except Exception as e:
        print(e)
        validation_status['errors'].append('Some unexpected error occured')
    
    return validation_status