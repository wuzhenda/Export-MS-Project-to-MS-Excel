Sub Export_to_Excel()

'Variables
Dim xlRange as Excel.Range
Dim xlApp As Excel.Application
Dim xlWkb As Excel.Workbook
Dim xlSheet As Excel.Worksheet

'Open Excel
Set xlApp = New Excel.Application
xlApp.Visible = True
AppActivate "Microsoft Excel"
Set xlWkb = xlApp.Workbooks.Add

'###################################
'EXPORTS RESOURCE USAGE - FIRST HALF
'###################################

'Copy first half of resource usage data
ViewApply Name:="Resource Usage"
SelectResourceColumn Column:="Name", Additional:=1
OutlineShowAllTasks
EditCopy

'Selects where to add new data
Set xlSheet = xlWkb.Worksheets("Sheet1")
xlSheet.Activate
Set xlRange = xlSheet.Range("A2:A2")
xlRange.Select

'Paste data into Excel
xlRange.PasteSpecial Paste:=xlPasteValues


'###################################
'EXPORTS RESOURCE USAGE - SECOND HALF - not working, need to resolve
'###################################

'Copy second half of resource usage data
ViewApply Name:="Resource Usage"
PaneNext
SelectTimescaleRange Row:=1, StartTime:="Mon 29/07/13 00:00", Width:=13, Height:=10
EditCopy

'Selects where to add new data
Set xlSheet = xlWkb.Worksheets("Sheet1")
xlSheet.Activate
Set xlRange = xlSheet.Range("G2:G2")
xlRange.Select

'Paste data into Excel
xlRange.PasteSpecial Paste:=xlPasteValues

End Sub
