import win32com.client
import os

def export_excel_charts(xlsx_path, output_dir):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx_path)
    
    print(f"Opened workbook: {xlsx_path}")
    print(f"Number of sheets: {wb.Sheets.Count}")

    for i, sheet in enumerate(wb.Sheets):
        print(f"Processing sheet {i+1}: {sheet.Name}")
        chart_count = sheet.ChartObjects().Count
        print(f"  Charts found: {chart_count}")
        
        for j, chart in enumerate(sheet.ChartObjects()):
            chart_obj = chart.Chart
            image_path = f"{output_dir}/chart_{i}_{j}.png"
            chart_obj.Export(image_path)
            print(f"Saved: {image_path}")
    
    wb.Close(False)
    excel.Quit()


export_excel_charts(os.path.abspath("public/Book1.xlsx"), os.path.abspath("output"))