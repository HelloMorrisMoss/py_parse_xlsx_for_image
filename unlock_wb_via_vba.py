import subprocess

script_path = ""

subprocess.call(['cmd.exe', script_path])


#
# def set_password(excel_file_path, pw):
#     # this does the opposite, but it's neat and might be useful
#     from pathlib import Path
#
#     excel_file_path = Path(excel_file_path)
#
#     vbs_script = \
#     f"""' Save with password required upon opening
#
#     Set excel_object = CreateObject("Excel.Application")
#     Set workbook = excel_object.Workbooks.Open("{excel_file_path}")
#
#     excel_object.DisplayAlerts = False
#     excel_object.Visible = False
#
#     workbook.SaveAs "{excel_file_path}",, "{pw}"
#
#     excel_object.Application.Quit
#     """
#
#     # write
#     vbs_script_path = excel_file_path.parent.joinpath("set_pw.vbs")
#     with open(vbs_script_path, "w") as file:
#         file.write(vbs_script)
#
#     #execute
#     subprocess.call(['cscript.exe', str(vbs_script_path)])
#
#     # remove
#     vbs_script_path.unlink()
#
#     return None
#
#
# # this seems to actually open the wb and then save it without a pw
# # could pretty easily be reversed to encrypt the wb as well
# # this is ---> VBA <--- needs to be converted
# # https://stackoverflow.com/a/42103556/10941169
# """Sub OpenAndSaveWithoutPasswords()
#
#     Dim wb As Workbook
#
#     Application.DisplayAlerts = False
#
#     With Application.FileDialog(msoFileDialogFilePicker)
#       If .Show = False Then Exit Sub
#       strFolderPath = .SelectedItems(1)
#     End With
#
#     Set wb = Workbooks.Open(Filename:=strFolderPath, Password:="WKX35H", WriteResPassword:="MODIFY PASSWORD")
#     wb.SaveAs Filename:=strFolderPath, Password:="", WriteResPassword:=""
#
#     Application.DisplayAlerts = True
#
# End Sub"""
#
#
# def remove_password(excel_file_path, pw):
#     # this does the opposite, but it's neat and might be useful
#     from pathlib import Path
#
#     excel_file_path = Path(excel_file_path)
#
#     vbs_script = \
#         f"""Sub OpenAndSaveWithoutPasswords()
#
#             Dim wb 'As Workbook
#             Dim strFolderPath
#             Dim Application
#
#             Set Application = CreateObject("Excel.Application")
#
#             Application.DisplayAlerts = False
#
#             'With Application.FileDialog(msoFileDialogFilePicker)
#             ' If .Show = False Then Exit Sub
#             '  strFolderPath = .SelectedItems(1)
#             'End With
#
#             strFolderPath = "{excel_file_path}"
#
#             'Set wb = Workbooks.Open(Filename:=strFolderPath, Password:="{pw}", WriteResPassword:="MODIFY PASSWORD")
#             'Set wb = Workbooks.Open(strFolderPath, Null, Null, Null, {pw}, "MODIFY PASSWORD")
#              'Set wb = Application.Workbooks.Open(strFolderPath, Null, Null, Null, {pw}, "ModifyPassword", Null, Null, Null, Null, Null, Null, Null, Null, Null)
#             'wb.SaveAs strFolderPath, Password:="", WriteResPassword:=""
#             'wb.SaveAs strFolderPath, Null, "", ""
#             Set wb = Application.Workbooks.Open(strFolderPath, , , , "{pw}", "MODIFY PASSWORD") ' need to figure out the right number of empty arguments?
#             wb.SaveAs strFolderPath, , "", "", , , , , , , , ,
#             'wb.SaveAs strFolderPath, Null, "", "", Null, Null, Null, Null, Null, Null, Null, Null, Null
#
#             Application.DisplayAlerts = True
#
#         End Sub
#         """
#
#     # write
#     vbs_script_path = excel_file_path.parent.joinpath("set_pw.vbs")
#     with open(vbs_script_path, "w") as file:
#         file.write(vbs_script)
#
#     #execute
#     subprocess.call(['cscript.exe', str(vbs_script_path)])
#
#     # remove
#     vbs_script_path.unlink()
#
#     return None
#
#
# pw = "WKX35H"
# file_path = r"C:\my documents\vba unlock test\Operator Lookup Special Instructions v2.1 (2).xlsx"
# remove_password(file_path, pw)
