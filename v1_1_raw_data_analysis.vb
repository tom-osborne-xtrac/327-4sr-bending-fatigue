'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' THIS MODULE INCLUDES ALL FUNCTIONS RELATING TO THE RAW DATA ANALYSIS WORKSHEET'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Directory Change Code

'Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
'#If VBA7 Then
'    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 '       (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
'        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
'#Else
'     Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
'        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
'        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'#End If

'Code to adjust "Raw Data Plot" chart scales

Sub SetRawDataPlotChartScales()

Dim t As Worksheet
Dim RawDataMin As String
Dim RawDataMax As String

RawDataMin = WorksheetFunction.RoundDown((Sheets("Data Processing").Range("AJ9").Value), -1) - 10
RawDataMax = WorksheetFunction.RoundUp((Sheets("Data Processing").Range("AJ8").Value * (35 / 32)), -1) + 10

Debug.Print "> Setting RawDataPlot graph scales..."
Debug.Print "Graph scale min: " & RawDataMin & " Graph scale max: " & RawDataMax

Set t = Sheets("Raw Data Analysis")

t.ChartObjects("RawDataPlot").Chart.Axes(xlValue).MinimumScale = RawDataMin
t.ChartObjects("RawDataPlot").Chart.Axes(xlValue).MaximumScale = RawDataMax

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Raw data import
' This prompts the user for a FileName and then calls LoadFromFile macro
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RawData_Import()
    Dim fileName As Variant
    Dim fso As New FileSystemObject
    Dim ShortfileName As String
    
    Debug.Print ">> Raw data import started..."
    
      
    'Changes the directory to the Data file folder - only works if RD drive is currently active
    Debug.Print "StartDir_" & CurDir
    'SetCurrentDirectory "\\UGSVR01\rd$\"
        ChDir "\\UGSVR01\rd$\R&D 4SQ\"
        check = CurDir
    Debug.Print "EndDir___" & CurDir
    
    'Open browser window for user to identify Raw Data file
    fileName = Application.GetOpenFilename(FileFilter:="CSV file (*.csv),*.csv")
    If fileName = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        ''''''''''''''''''''''''''
        Debug.Print "User cancelled!"
        Exit Sub
    End If
    'Debug.Print "FileName: " & fileName
    
    'Delete any data still in Raw Data worksheet and clear the "current file" cell
    ResetDocument
        
    'Run LoadFromFile sub to open the selected file and paste data into Raw Data worksheet
    RawData_LoadFromFile FName:=CStr(fileName)
    
    'Set Current File cell on Raw Data Analysis Page to filename just loaded
    ShortfileName = fso.GetFileName(fileName)
    Sheets("Raw Data Analysis").Range("Z15").Value = ShortfileName
    
    Calculate
         
    'Adjust chart scales
    SetRawDataPlotChartScales
            
    Debug.Print "Import Complete!"
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END DoTheImport
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DoTheExport
' This prompts the user for the FileName and then calls the ExportToTextFile procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DontyneExport()
    Dim fileName As Variant
    Dim Sep As String
    Dim r As Range
    'Dim TorqueRes As String
    'Dim FailGear As String
    Dim Material As String
    Dim GearSet As String
    Dim AENo As String
    Dim OutputName As String
    Dim DurHrs As String
    Dim AvgSpd As String
    Dim t As Worksheet
    Dim d As Worksheet
    Dim o As Worksheet
                    
    Debug.Print ">> Export Started..."
                    
    Set t = Sheets("Raw Data Analysis")
    Set d = Sheets("Data Processing")
    Set o = Sheets("Dontyne Output")
    
'This is automatically writing the filename for export i.e. "OutputName"
    AENo = t.Range("Z16").Text
    Material = t.Range("Z18").Text
    GearSet = t.Range("Z17").Text
    'TorqueRes = t.Range("B42").Text
    'FailGear = t.Range("H11").Text
    DurHrs = Left(Replace((d.Range("AJ5").Text), ".", "_"), 4)
    AvgSpd = Left(Replace((d.Range("AJ13").Text), ".", "_"), 4)
        
    OutputName = CStr(AENo) & " - " & CStr(GearSet) & " - " & CStr(Material) & " - " & CStr(AvgSpd) & "rpm" & " - " & CStr(DurHrs) & "hrs"
           
Application.ScreenUpdating = False

'Error checking - Checks that some Raw Data is actually present
CheckRawData
          
    'Opens a browser window for the user to select where to save data
    'Automatically set to the Dontyne Output folder on RD drive
    fileName = Application.GetSaveAsFilename(InitialFileName:="\\UGSVR01\users\rd\R&D 4SQ\" & OutputName, FileFilter:="Text Files (*.txt),*.txt")
    If fileName = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        '''''''''''''''''''''''''
        Debug.Print "User cancelled!"
        Exit Sub
    End If
    
    Sep = " "
    ' If Sep = vbNullString Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        ''''''''''''''''''''''''''
    '    Exit Sub
    'End If
            
    'Unhides and then selects the Dontyne output worksheet - not visible to user
    o.Visible = True
    o.Select
            
    'Runs ExportToTextFile sub
    Debug.Print "FileName:" & OutputName
    ExportToTextFile FName:=CStr(fileName), Sep:=CStr(Sep), _
       SelectionOnly:=False, AppendData:=False

'Hides the Dontyne output worksheet and selects the front sheet as active - not visible to user
Sheets("Raw Data Analysis").Select
'Sheets("Dontyne Output").Visible = False

'Add data to Processed data table
    table_CompileTable

Application.ScreenUpdating = True

Debug.Print "Export Complete!"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END DoTheExport
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RawData_Trim()
'Macro to trim last two rows of raw data to remove any torque drop offs

Debug.Print ">> Data Trim Started..."

'Error checking - Checks that some Raw Data is actually present
CheckRawData

'Turn off screen updating
Application.ScreenUpdating = False

'Open raw data worksheet
Sheets("Raw Data").Visible = True

'Find, select and delete the last two rows
Sheets("Raw Data").Select
With ActiveSheet
    Range("A53").End(xlDown).EntireRow.Offset(-4, 0).Resize(5).ClearContents
End With

'Close raw data worksheet
Sheets("Raw Data Analysis").Select
Sheets("Raw Data").Visible = False

'Turn screen updating back on
Application.ScreenUpdating = True

'Adjust chart scales
SetRawDataPlotChartScales

Calculate

RefreshCharts

Debug.Print "Raw data trim complete."

End Sub

Sub table_CompileTable()

'Dim Woksheets
Dim p As Worksheet
Dim t As Worksheet

'Set worksheets
Set p = Sheets("Raw Data")
Set t = Sheets("Raw Data Analysis")

'Dim items for compiler
Dim r As Range
Dim rCur As Range

Application.ScreenUpdating = False

'Set start point
Set rCur = t.Range("A39:Q39")
Set rStart = t.Range("A44:Q44")

Debug.Print ">> Compiling Table..."

'Error checking - Checks that some Raw Data is actually present
CheckRawData

rCur.Select
Selection.Copy

'check if start cell is empty if not find next empty row
If IsEmpty(r) Then
    rStart.Select
Else
    t.Range("A100:Q100").End(xlUp).Offset(1).Select
End If

'Fill out row with current data
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
        
Debug.Print "Table Compiled"

Application.ScreenUpdating = True


End Sub

Sub table_DeleteLastRow()

'Dim Objects
Dim t As Worksheet
Dim r As Range

'Set Objects
Set t = Sheets("Raw Data Analysis")
Set r = t.Range("A44")

'check if start cell is empty if not find next empty row
If IsEmpty(r) Then
    r.Select
    MsgBox "Table is empty!"
    Exit Sub
Else
    t.Range("A100").End(xlUp).Resize(1, 17).ClearContents
End If

End Sub

Sub table_SaveCompiledData()

Dim t As Worksheet
Dim r As Range
Dim fileName As Variant
Dim Sep As String
Dim OutputName As String


Set t = Sheets("Raw Data Analysis")
Set r = t.Range("A44")
OutputName = Left(Replace((t.Range("X53").Text), ".", "_"), 400)


Debug.Print ">> Saving compiled data..."

'check if table is empty
If IsEmpty(r) Then
    r.Select
    MsgBox "Table is empty!"
    Exit Sub
Else
 '''''SAVE CODE HERE'''''
 
    t.Range("A44:Q63").Select
    
    Sep = ","
    
    'Opens a browser window for the user to select where to save data
    'Automatically set to the Dontyne Output folder on RD drive
    fileName = Application.GetSaveAsFilename(InitialFileName:="\\UGSVR01\users\rd\R&D 4SQ\" & OutputName, FileFilter:="Text Files (*.txt),*.txt")
    If fileName = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        '''''''''''''''''''''''''
        Debug.Print "User cancelled!"
        Exit Sub
    End If
    

    
    'Runs ExportToTextFile sub
    Debug.Print "FileName:" & OutputName
    ExportToTextFile FName:=CStr(fileName), Sep:=CStr(Sep), _
       SelectionOnly:=True, AppendData:=False


End If

Debug.Print "Compiled data saved!"

End Sub

Sub table_import()
    Dim fileName As Variant
    Dim fso As New FileSystemObject
    Dim ShortfileName As String
    Dim testFilenameCell As Range
    
        
    Debug.Print ">> Opening compiled data..."
    
    'Changes the directory to the Data file folder - only works if RD drive is currently active
    Debug.Print "StartDir_" & CurDir
    'SetCurrentDirectory "\\UGSVR01\rd$\"
        ChDir "\\UGSVR01\rd$\R&D 4SQ\"
        check = CurDir
    Debug.Print "EndDir___" & CurDir
    
    'Open browser window for user to identify Raw Data file
    fileName = Application.GetOpenFilename(FileFilter:="TXT File (*.txt),*.txt")
    If fileName = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        ''''''''''''''''''''''''''
        Debug.Print "User cancelled!"
        Exit Sub
    End If
    

    Sheets("Raw Data Analysis").Range("A44:Q63").ClearContents
    Sheets("Raw Data Analysis").Range("X53").ClearContents
    
    'Run LoadFromFile sub to open the selected file and paste data into Raw Data worksheet
    table_LoadFromFile FName:=CStr(fileName)
    ShortfileName = Left(Left(Replace(fso.GetFileName(fileName), "_", "."), 400), Len(fso.GetFileName(fileName)) - 4)

    Sheets("Raw Data Analysis").Range("X53") = ShortfileName
    
    

End Sub

Sub table_Clear()

Dim r As Range
Dim t As Worksheet

Set t = Sheets("Raw Data Analysis")
Set r = t.Range("A44")

Debug.Print "Clearing 'Compiled Data' table..."

'check if table is empty
If IsEmpty(r) Then
    r.Select
    Debug.Print "Table is empty!!"
    MsgBox "Table is empty!"
    Exit Sub
Else

    Sheets("Raw Data Analysis").Range("A44:Q63").ClearContents
    Sheets("Raw Data Analysis").Range("X53").ClearContents

End If

Debug.Print "Table cleared!"

End Sub

Sub RefreshCharts()

Dim myChart As ChartObject
Dim myCharts As ChartObjects
Dim myChartname As String
 
Set myCharts = ActiveSheet.ChartObjects
 
For Each myChart In myCharts
    myChartname = myChart.Name
    ActiveSheet.ChartObjects(myChartname).Chart.Refresh
Next

Calculate

End Sub
