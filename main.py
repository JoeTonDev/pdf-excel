import os
import win32com.client

def transform_to_pdf(excel_file_path):
  # Create connection to MS Excel
  excel = win32com.Workbooks.Open("Excel.Application")
  
  # Open the Excel file
  workbook = excel.Workbooks.Open(excel_file_path)
  
  # Loop through each sheet in the workbook
  for sheet in workbook.Sheets:
    # Save the sheet as a PDF file
    sheet.ExportAsFixedFormat(0, f"{sheet.Name}.pdf")
    
    # Clsoe the workbook and exit the Excel Application
    workbook.Close()
    excel.Quit()
    
# Call the function and pass the path to the Excel file as an argument
transform_to_pdf("sample_excel_file.xlsx")