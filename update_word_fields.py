import win32com.client
import os
import shutil

src_path = os.path.abspath(r'C:\MAMP\htdocs\Riaya\ReportTerm5\FinalReportsDOCX\Riaya_Final_Report.docx')
dest_path = os.path.abspath(r'C:\MAMP\htdocs\Riaya\Report\Riaya_Final_Report.docx')

print('Starting Word application...')
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

try:
    print(f'Opening document {src_path}...')
    doc = word.Documents.Open(src_path, False, False, False)
    
    # Update single TOC
    print(f'Found {doc.TablesOfContents.Count} Table(s) of Contents.')
    for i in range(1, doc.TablesOfContents.Count + 1):
        doc.TablesOfContents(i).Update()
    
    # Update Table of Figures & Table of Tables
    print(f'Found {doc.TablesOfFigures.Count} Table(s) of Figures/Tables.')
    for i in range(1, doc.TablesOfFigures.Count + 1):
        doc.TablesOfFigures(i).Update()
    
    doc.Save()
    pages = doc.ComputeStatistics(2)
    print(f'Successfully updated fields. Total page count: {pages}')
    doc.Close(False)
finally:
    word.Quit()

# Copy updated file to Report directory
shutil.copy2(src_path, dest_path)
print(f'Successfully copied updated report to {dest_path}')
print('All Word fields and TOCs updated successfully!')
