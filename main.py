# import packages
import os
import win32com.client as win32

# USE DYNAMIC FILE PATHS
# path where xml files are imported
XML_PATH="C:\\Users\\Dell\\Desktop\\Work\\WIP"
# path where tool resides
TOOL_PATH="C:\\Users\\Dell\\Desktop\\Work\\Emochb"
# path where final copy is saved
FINAL_PATH="C:\\Users\\Dell\\Desktop\\Work\\Emochb\\final"


def save_xml_to_excel(file):
    """ opens up xml file on predefined schemas and reformat excel file with macro vba and saves file at destination
    file[str]: name of xml file 
    """
    xml_file=f"{XML_PATH}\\{file}"
    tool_file=f"{TOOL_PATH}\\TOOL.xlsm"
    final_name=file.split(".")[-2]
    final_file=f"{FINAL_PATH}\\{final_name}.xlsx"

    # opening excel and workbook
    excel = win32.Dispatch('Excel.Application')
    try:
        wb=excel.Workbooks.Open(tool_file)
    except Exception:
        # for blockers due to already opened file
        workbooks=excel.Workbooks
        while workbooks.Count>0:
            workbooks(1).Close(False)
        wb=excel.Workbooks.Open(tool_file)
    try:
        wb.XmlImport(xml_file)
        wb.Application.Run("rpa_automate_emochb")
        excel.DisplayAlerts=False
        wb.SaveAs(final_file,FileFormat=51)
    except Exception as e:
        print(e)
    finally:
        wb.Close(False)
        excel.Quit()


# list xmls
xmls=[file for file in os.listdir(XML_PATH) if file[-3:]=="xml" ]
total=len(xmls)
for i,xml in enumerate(xmls):
    save_xml_to_excel(xml)
    print(f"Completed {i+1}/{total}")
