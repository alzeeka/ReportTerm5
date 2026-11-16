import win32com.client
import os
import shutil

doc_paths = [
    os.path.abspath(r'C:\MAMP\htdocs\Riaya\ReportTerm5\FinalReportsDOCX\Riaya_Final_Report.docx'),
    os.path.abspath(r'C:\MAMP\htdocs\Riaya\Report\Riaya_Final_Report.docx')
]

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
try:
    for path in doc_paths:
        print(f'Updating fields in {path}...')
        doc = word.Documents.Open(path)
        for toc in doc.TablesOfContents:
            toc.Update()
        for f in doc.Fields:
            f.Update()
        doc.Save()
        print(f'Updated and saved {path}. Page count: {doc.ComputeStatistics(2)}') # 2 is wdStatisticPages
        doc.Close(False)
finally:
    word.Quit()
print('All Word fields and TOCs updated successfully!')
