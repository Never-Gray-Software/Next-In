Attribute VB_Name = "Compare"
Public DataArray() As String
Public FormIn As Variant

Sub ReadFilesv2(ColumnStart As Long, Optional ByVal unit_name As String)
On Error GoTo ErrorProc
    Dim MaxRowCount As Long    ' Do not use Integer, may be too small and cause overflow
    Set FormIn = Worksheets("Comparison")
    If unit_name = "" Then
        Call choosefile(Infile)
    Else
        Infile = unit_name
    End If
    If Infile = "" Then End
    FormIn.Cells(1, ColumnStart).Value2 = Infile
    With Worksheets("Comparison")
        MaxRowCount = .UsedRange.Rows.Count
        .Range(.Cells(4, ColumnStart), .Cells(MaxRowCount, ColumnStart + 7)).ClearContents
        .Range(.Cells(4, ColumnStart), .Cells(MaxRowCount, ColumnStart + 7)).Style = "Normal"
        '.Range(.Cells(Row, ColumnStart), .Cells(Row, ColumnStart).End(xlDown)).ClearContents
        '.Range(.Cells(Row, ColumnStart), .Cells(Row, ColumnStart).End(xlDown).Offset(0, 7)).Style = "Normal"
    End With
    Call TextFileToArray(Infile)
    Call Array2excel(ColumnStart)
    Exit Sub
ErrorProc:
  MsgBox "Error in procedure ReadFilesv2 : " & Err.Description
  'Call Speed(True)
  Err.Clear
End Sub


Public Sub choosefile(Infile)
  On Error GoTo ErrorProc
  'Declare a variable as a FileDialog object.
  Dim FD As FileDialog
  'Create a FileDialog object as a File Picker dialog box.
  Set FD = Application.FileDialog(msoFileDialogFilePicker)
  'Declare a variable to contain the path
  'of each selected item. Even though the path is a String,
  'the variable must be a Variant because For Each...Next
  'routines only work with Variants and Objects.
  Dim vrtSelectedItem As Variant
  'Use a With...End With block to reference the FileDialog object.
  With FD
    'Use the Show method to display the File Picker dialog box and return the user's action.
    'The user pressed the action button.
    .Filters.Add "SES Input Files", "*.SES; *.INP; *.SVS", 1
    If .Show = -1 Then
      For Each vrtSelectedItem In .SelectedItems
        'vrtSelectedItem is a String that contains the path of each selected item.
        'You can use any file I/O functions that you want to work with this path.
        'This example simply displays the path in a message box.
        'MsgBox "The path is: " & vrtSelectedItem
      Infile = vrtSelectedItem
      Next vrtSelectedItem
    Else: Infile = ""
    End If
  End With
  'Set the object variable to Nothing.
  Set FD = Nothing
  Exit Sub
  
ErrorProc:
  MsgBox "Error in procedure ChooseFile : " & Err.Description
  Err.Clear
End Sub



Sub TextFileToArray(ByVal FilePath As String)
'PURPOSE: Load an Array variable with data from a delimited text file
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FileContent As String
Dim LineArray() As String
Dim TempArray() As String
Dim rw As Long, col As Long


On Error GoTo ErrorProc
  
'Open the text file in a Read State
  TextFile = FreeFile
  Open FilePath For Input As TextFile
  
'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Close Text File
  Close TextFile
  
'Separate Out lines of data
  LineArray() = Split(FileContent, vbCrLf)
  If UBound(LineArray) < 1 Then 'if line divider is not vbCrlF
    LineArray() = Split(FileContent, Chr(10))
  End If
  
  
  

'Size DataArray for input file
Erase DataArray
ReDim Preserve DataArray(UBound(LineArray), 7)
  
  For x = LBound(LineArray) To UBound(LineArray)
    For y = 0 To 7
            DataArray(x, y) = Mid(LineArray(x), y * 10 + 1, 10)
    Next y
  Next x
  Exit Sub

ErrorProc:
  MsgBox "Error in procedure TextFileToArray: " & Err.Description
  Err.Clear
End Sub

Sub Array2excel(ColumnStart As Long)
    Dim RowStart As Long
    Dim LastDataRow As Long
    Dim Last
    RowStart = 4
    LastDataRow = UBound(DataArray, 1)
    With Worksheets("Comparison")
        .Range(.Cells(4, ColumnStart), .Cells(LastDataRow - 4, ColumnStart + 7)).Style = "Normal"
    End With
    OutRange2D(RowStart, ColumnStart, 8, LastDataRow, FormIn) = DataArray()
End Sub

Function OutRange2D(rindex As Long, Cindex As Long, ndata As Long, nlines As Long, sname As Variant) As Range 'Starting Line, Number of Data Points. Consider adding additional data
    With sname
        Set OutRange2D = .Range(.Cells(rindex, Cindex), .Cells(rindex + nlines - 1, Cindex + ndata - 1))
    End With
End Function

Sub CompareArray3()
    On Error GoTo ErrorProc
    Dim K As Integer, MaxRowCount As Integer
    Dim Difference As Boolean
    Dim WReport, WComp As Worksheet
    Dim tolerence As Double
    Dim percent_difference As Variant
    Set WReport = Worksheets("Report")
    Set WComp = Worksheets("Comparison")
    K = 0
    With Worksheets("Comparison")
        LastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        LastRow2 = .Cells(.Rows.Count, "L").End(xlUp).Row
        LastRow = WorksheetFunction.Min(LastRow, LastRow2)
        'check arrays will be the same size
        If LastRow <> LastRow2 Then
            MsgBox ("Length Files is different.")
        End If
        Set Original = Range(.Cells(4, 2), .Cells(LastRow, 9))
        File1 = Original.Value2
        Set NextIn = Range(.Cells(4, 12), .Cells(LastRow, 19))
        File2 = NextIn.Value2
        MaxRowCount = LastRow
        tolerence = .Range("AC2").Value2
    End With
    Call ClearFormatting(4, 2, MaxRowCount, 8)
    Call ClearFormatting(4, 12, MaxRowCount, 8)
    Call ClearFormatting(4, 22, MaxRowCount, 8)
    Call ClearFormatting(4, 27, MaxRowCount, 1)
    Column = 21
    Row = 4
    With WComp
        .Range(.Cells(4, Column), .Cells(LastRow - 4, Column + 8)).ClearContents
        .Range(.Cells(4, 31), .Cells(LastRow, 31)).ClearContents
    End With
    For i = LBound(File1, 1) To UBound(File1, 1)
        For J = LBound(File1, 2) To UBound(File1, 2)
            percent_difference = "X"
            Difference = False
            
            Dim value_left As Variant, value_right As Variant
            value_left = File1(i, J)
            value_right = File2(i, J)
            
            '===========================
            ' 1. ZERO-LIKE COMPARISON
            '===========================
            If Val(value_left) = 0 Or Val(value_right) = 0 Then
            
                'Both zero-like ? no difference
                If IsZeroLike(value_left) And IsZeroLike(value_right) Then
                    Difference = False
            
                'One zero-like, one not ? difference
                ElseIf IsZeroLike(value_left) Xor IsZeroLike(value_right) Then
                    Difference = True
                    percent_difference = "X"
            
                'Both numeric zero but text differs ? difference
                ElseIf Trim(value_left) <> Trim(value_right) Then
                    Difference = True
                    percent_difference = "X"
                End If
            
            '===========================
            ' 2. NUMERIC DIFFERENCE
            '===========================
            ElseIf IsNumeric(value_left) And IsNumeric(value_right) Then
            
                If CDbl(value_left) <> 0 Then
                    original_value = CDbl(value_left)
                    new_value = CDbl(value_right)
                    percent_difference = Abs((original_value - new_value) / original_value)
            
                    If percent_difference >= tolerence Then
                        Difference = True
                    End If
            
                Else
                    'Avoid divide-by-zero
                    If CDbl(value_right) <> 0 Then
                        Difference = True
                        percent_difference = "X"
                    End If
                End If
            
            '===========================
            ' 3. NON-NUMERIC DIFFERENCE
            '===========================
            Else
                If Trim(value_left) <> Trim(value_right) Then
                    Difference = True
                    percent_difference = "X"
                End If
            End If
            
            '===========================
            ' 4. WRITE DIFFERENCE
            '===========================
            If Difference Then
                K = K + 1
                CompRow = i + 3
                CompCol = J + 21
            
                WComp.Cells(CompRow, 21).Value2 = "Diff"
                WComp.Cells(CompRow, CompCol).Value2 = percent_difference
                WComp.Cells(CompRow, CompCol).Style = "Bad"
            
                File1Col = J + 1
                File2Col = J + 11
                WComp.Cells(CompRow, File1Col).Style = "Bad"
                WComp.Cells(CompRow, File2Col).Style = "Bad"
            
                WComp.Cells(Row, 31).Value2 = "R: " & i & " C: " & J
                Row = Row + 1
            End If
        Next J
    Next i
    If K = 0 Then
        WComp.Cells(1, 26).Value2 = "No Difference"
    Else
        WComp.Cells(1, 26).Value2 = K & " Difference"
    End If
    'Set PageSetup Area
    With WComp
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(LastRow, 31)).Address
    End With
Exit Sub
ErrorProc:
  MsgBox "Error in procedure Compare : " & Err.Description
  Err.Clear
End Sub

Sub ClearFormatting(StartRow As Integer, StartColumn As Integer, NumRow As Integer, NumCol As Integer)
    With Worksheets("Comparison")
        .Range(.Cells(StartRow, StartColumn), .Cells(StartRow + NumRow - 1, StartColumn + NumCol - 1)).Style = "Normal"
    End With
End Sub

Sub SavePDF(Optional unit_name As String)
    Dim FileName As String
    If unit_name = Null Then
        FileName = ActiveSheet.Range("B1").Value2
    Else
        FileName = unit_name
    End If
    Length = Len(FileName)
    FileName = Left(FileName, Length - 4) & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        FileName, Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub
Private Function IsZeroLike(ByVal v As Variant) As Boolean
    Dim s As String
    s = Trim(CStr(v))

    ' Empty string ? treat as zero-like
    If s = "" Then
        IsZeroLike = True
        Exit Function
    End If

    ' Pure numeric zero (including "-0", "0.", "-0.0", "0.000")
    If IsNumeric(s) Then
        If CDbl(s) = 0 Then
            IsZeroLike = True
            Exit Function
        End If
    End If

    ' Non-numeric but visually zero-like
    If s = "." Or s = "-." Then
        IsZeroLike = True
        Exit Function
    End If

    IsZeroLike = False
End Function
