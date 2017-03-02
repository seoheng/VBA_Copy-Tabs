# VBA_Copy-Tabs
Loop through listed spreadsheets, and copy them into tabs of stated file

Sub REL_BMK()

Dim FileRange As Range
Dim Count As Integer 'Count variable
Dim StartRow As Integer 'Starting row number to place data
Dim SourceFile As String
Dim FileName As Variant
Dim path As Variant

Application.StatusBar = "Please be patient - clearing the old data"
Application.ScreenUpdating = False

'*************************************************
'Loop for Factset Files to be opened and copied
    
    ThisWorkbook.Activate
    Worksheets("FilesConsol").Activate
    Set FileRange = Range("G15", Range("G15").End(xlDown))
    FileName = Cells(10, 7).Value
    
    Worksheets("Print").Activate
    path = Cells(8, 5).Value & "\"
    
    Workbooks.Add
    ActiveWorkbook.SaveAs FileName:=path & FileName, FileFormat _
        :=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:= _
        False, CreateBackup:=False

    'Initialise variables
    
    StartRow = 15
    Count = 0
            
    'Loop for tabs to clear
    
     For Each d In FileRange
        
        If d.Value <> "" Then
            
           ThisWorkbook.Activate
           Worksheets("FilesConsol").Activate
           SourceFile = Cells(StartRow + Count, 7).Value
           
           Workbooks.Open FileName:=path & SourceFile & ".xls", UpdateLinks:=1
           Workbooks(SourceFile & ".xls").Activate
           ActiveSheet.Copy Before:=Workbooks(FileName).Sheets(Count + 1)
           
           'Close Raw Data file

           Workbooks(SourceFile & ".xls").Activate
           Workbooks(SourceFile & ".xls").Close
           
           Count = Count + 1
            
        End If
        
      Next d
      
    'Save Final Report File
    Workbooks(FileName).Activate
    Workbooks(FileName).Save
    Workbooks(FileName).Close
      
    'Cursor Default Location
    ThisWorkbook.Activate
    Worksheets("FilesConsol").Activate
      
      
Application.StatusBar = False
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
        
End Sub





