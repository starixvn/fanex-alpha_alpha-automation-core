from robot.libraries.BuiltIn import BuiltIn
from xlrd import open_workbook

def open_excel_without_keep_process(filename, useTempDir=False):
    """
        Opens the Excel file from the path provided in the file name parameter.
        If the boolean useTempDir is set to true, depending on the operating system of the computer running the test the file will be opened in the Temp directory if the operating system is Windows or tmp directory if it is not.

        Arguments:
                |  File Name (string)                      | The file name string value that will be used to open the excel file to perform tests upon.                                  |
                |  Use Temporary Directory (default=False) | The file will not open in a temporary directory by default. To activate and open the file in a temporary directory, pass 'True' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

    """
    
    excel = BuiltIn().get_library_instance('ExcelLibrary')
    if useTempDir is True:
        print 'Opening file at %s' % filename
        excel.wb = open_workbook(os.path.join("/", excel.tmpDir, filename), formatting_info=False, on_demand=False)
    else:
        excel.wb = open_workbook(filename, formatting_info=False, on_demand=False)
        excel.fileName = filename
        excel.sheetNames = excel.wb.sheet_names()