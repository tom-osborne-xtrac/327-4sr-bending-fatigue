'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' THIS MODULE INCLUDES ALL FUNCTIONS RELATING TO THE OVERALL FUNCTION OF THE WORKBOOK '
' These Functions can be called upon from any worksheet                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Auto_Open() 'Opens workbook up on "Raw Data Analysis" tab
    Sheets("Raw Data Analysis").Select
End Sub

Sub ResetDocument()
' ResetDocument Macro - Currently only deletes raw data and is only used on the "Raw Data Analysis" tab
    
    Sheets("Raw Data").Cells.Clear
    Sheets("Raw Data Analysis").Range("Z15").ClearContents
    
    Calculate
       
End Sub

Public Sub CheckRawData()

'Error checking - Checks that some Raw Data is actually present
Debug.Print "> Checking for raw data..."

    Set r = Sheets("Raw Data").Range("A1")
    
    If IsEmpty(r) Then
        MsgBox ("No raw data found! Use the 'Import Raw Data' button on the 'Data Analysis' menu")
        
        Debug.Print "Raw data missing. Operation aborted."
        
        End
    End If

Debug.Print "Raw data found."

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadFromFile
' This macro opens the file selected in the DoTheImport macro and loads the raw data into the worksheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RawData_LoadFromFile(FName As String)
      
    With Sheets("Raw Data").QueryTables _
        .Add(Connection:="TEXT;" & FName, Destination:=Range("'Raw Data'!$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Debug.Print "File imported..."
        
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LoadFromFile
' This macro opens the file selected in the DoTheImport macro and loads the raw data into the worksheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub table_LoadFromFile(FName As String)
      
    With Sheets("Raw Data Analysis").QueryTables _
        .Add(Connection:="TEXT;" & FName, Destination:=Range("'Raw Data Analysis'!$A$44"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Debug.Print "File imported..."
        
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExportToTextFile
' This exports the "Dontyne Output" worksheet to a text file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExportToTextFile(FName As String, _
    Sep As String, SelectionOnly As Boolean, _
    AppendData As Boolean)

Dim WholeLine As String
Dim FNum As Integer
Dim RowNdx As Long
Dim ColNdx As Integer
Dim StartRow As Long
Dim EndRow As Long
Dim StartCol As Integer
Dim EndCol As Integer
Dim CellValue As String


Application.ScreenUpdating = False
On Error GoTo EndMacro:
FNum = FreeFile

If SelectionOnly = True Then
    With Selection
        StartRow = .Cells(1).Row
        StartCol = .Cells(1).Column
        EndRow = .Cells(.Cells.Count).Row
        EndCol = .Cells(.Cells.Count).Column
    End With
Else
    With ActiveSheet.UsedRange
        StartRow = .Cells(1).Row
        StartCol = .Cells(1).Column
        EndRow = .Cells(.Cells.Count).Row
        EndCol = .Cells(.Cells.Count).Column
    End With
End If

Debug.Print "Rows " & EndRow & " Columns " & EndCol
If AppendData = True Then
    Open FName For Append Access Write As #FNum
Else
    Open FName For Output Access Write As #FNum
End If

For RowNdx = StartRow To EndRow
    WholeLine = ""
    For ColNdx = StartCol To EndCol
        If Cells(RowNdx, ColNdx).Value = "FALSE" Then
           GoTo SkipLine:
        Else
           CellValue = Cells(RowNdx, ColNdx).Value
        End If
        WholeLine = WholeLine & CellValue & Sep
    Next ColNdx
    WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))
    Print #FNum, WholeLine
SkipLine:
Next RowNdx

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True
Close #FNum

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END ExportTextFile
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

