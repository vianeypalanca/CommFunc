Imports Excel = Microsoft.Office.Interop.Excel

Public Module ExcelFunctions

    Private Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr,
              ByRef lpdwProcessId As Integer) As Integer

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function CreateNewWorkbook(ByRef ExcelApp As Excel.Application, ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelProc As Process)

        ExcelApp = CreateObject("Excel.Application")
        ExcelWorkbook = ExcelApp.Workbooks.Add

        Dim CurrentProcID As Integer = 0
        Dim xlHWND As Integer = ExcelApp.Hwnd
        Call GetWindowThreadProcessId(xlHWND, CurrentProcID)
        ExcelProc = Process.GetProcessById(CurrentProcID)

    End Function

    Public Function LoadWorkbook(ByRef ExcelApp As Excel.Application, ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelProc As Process, ByVal WorkbookPath As String) As Boolean

        LoadWorkbook = False

        ExcelApp = New Microsoft.Office.Interop.Excel.Application

        Dim CurrentProcID As Integer = 0
        Dim xlHWND As Integer = ExcelApp.Hwnd
        Call GetWindowThreadProcessId(xlHWND, CurrentProcID)
        ExcelProc = Process.GetProcessById(CurrentProcID)

        Try
            ExcelWorkbook = ExcelApp.Workbooks.Open(WorkbookPath)
            LoadWorkbook = True
        Catch

        End Try

    End Function

    Public Function SaveAsWorkbook(ByVal ExcelWorkbook As Excel.Workbook, ByVal FileName As String, Optional DeleteExisting As Boolean = False)

        If DeleteExisting = True Then
            If My.Computer.FileSystem.FileExists(FileName) Then IO.File.Delete(FileName)
        End If

        ExcelWorkbook.SaveAs(FileName)

    End Function

    Public Function CloseWorkbook(ByRef ExcelApp As Excel.Application, ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelSheet As Excel.Worksheet, ByRef ExcelProc As Process, ByVal WorkbookPath As String,optional IsWin10 as boolean = false) As Boolean

        Try
            ExcelApp.DisplayAlerts = False
            ExcelWorkbook.Close(WorkbookPath)
            ExcelApp.Quit()
            releaseObject(ExcelApp)
            releaseObject(ExcelWorkbook)
            releaseObject(ExcelSheet)
            ExcelApp = Nothing
            If Not ExcelProc.HasExited Then
                ExcelProc.Kill()
            End If
            CloseWorkbook = True
        Catch
            CloseWorkbook = False
        End Try

    End Function

    Public Function AddWorksheet(ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelSheet As Excel.Worksheet, ByVal SheetName As String, Optional IsFirst As Boolean = False) As String

        If SheetName.Length > 30 Then SheetName = SheetName.Remove(30).Trim

        If WorkSheetExisting(ExcelWorkbook, SheetName) = False Then

            Dim newWorksheet As Microsoft.Office.Interop.Excel.Worksheet

            newWorksheet = CType(ExcelWorkbook.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
            newWorksheet.Name = SheetName
            Dim totalSheets As Integer = ExcelWorkbook.Sheets.Count
            If IsFirst = False Then
                newWorksheet.Move(After:=ExcelWorkbook.Worksheets(totalSheets))
            Else
                newWorksheet.Move(Before:=ExcelWorkbook.Worksheets(1))
            End If

            ExcelSheet = newWorksheet

        End If

        Return SheetName

    End Function

    Public Function RenameWorksheet(ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelSheet As Excel.Worksheet, ByVal OldName As String, ByVal NewName As String)

        ExcelSheet = ExcelWorkbook.Worksheets(OldName)
        ExcelSheet.Name = NewName

    End Function

    Public Function DeleteWorksheet(ByRef ExcelWorkbook As Excel.Workbook, ByVal SheetName As String)

        ExcelWorkbook.Worksheets(SheetName).delete

    End Function

    Public Function GetLastColumn(ByVal ExcelSheet As Excel.Worksheet, ByVal RowNum As Long) As Long

        GetLastColumn = ExcelSheet.Cells(RowNum, ExcelSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

    End Function

    Public Function HideColumns(ByVal ExcelSheet As Excel.Worksheet, ByVal StartCol As Long, ByVal EndCol As Long) As Boolean

        HideColumns = False

        ExcelSheet.Range(ExcelSheet.Cells(1, StartCol), ExcelSheet.Cells(1, EndCol)).EntireColumn.Hidden = True

        HideColumns = True

    End Function

    Public Function GetLastRow(ByVal ExcelSheet As Excel.Worksheet, ByVal ColumnNum As Long) As Long

        GetLastRow = ExcelSheet.Cells(ExcelSheet.Rows.Count, ColumnNum).end(Excel.XlDirection.xlUp).Row

    End Function

    Public Function ConvertColNum(ByVal ExcelSheet As Excel.Worksheet, ByVal ColNum As Long) As String

        Dim Varr() As String

        Varr = Split(ExcelSheet.Cells(1, ColNum).Address(True, False), "$")
        ConvertColNum = Varr(0)

    End Function

    Public Function WorkSheetExisting(ByVal ExcelWorkbook As Excel.Workbook, ByVal SheetName As String) As Boolean

        WorkSheetExisting = False

        For Each ExcelSheet As Excel.Worksheet In ExcelWorkbook.Sheets
            If UCase(ExcelSheet.Name) = UCase(SheetName) Then
                WorkSheetExisting = True
                Exit For
            End If
        Next ExcelSheet

    End Function

    Public Function WorkSheetToArray(ByVal ExcelWorkbook As Excel.Workbook, ByVal SheetName As String)

        Dim RetArr() As String
        Dim TempString As String
        Dim RangeString As String
        Dim ExcelSheet As Excel.Worksheet
        Dim Range As Excel.Range

        ExcelSheet = ExcelWorkbook.Sheets(SheetName)
        Range = ExcelSheet.UsedRange

        RangeString = "A1:" & ConvertColNum(ExcelSheet, Range.Columns.Count) & Range.Rows.Count
        ExcelSheet.Range(RangeString).EntireColumn.Hidden = False
        ExcelSheet.Range(RangeString).EntireRow.Hidden = False
        ExcelSheet.Range(RangeString).Copy()

        TempString = My.Computer.Clipboard.GetText
        TempString = TempString.Replace(vbTab, ":")
        RetArr = Split(TempString, vbCrLf)

        Return RetArr

    End Function

    Public Function AddWorksheetList(ByVal FileName As String, ByVal NameList() As String) As Boolean

        AddWorksheetList = True

        Dim STApp As Excel.Application
        Dim STWB As Excel.Workbook
        Dim STWS As Excel.Worksheet
        Dim STProc As Process

        Call LoadWorkbook(STApp, STWB, STProc, FileName)

        For x As Long = 0 To UBound(NameList)
            Call AddWorksheet(STWB, STWS, NameList(x))
        Next x

        STWB.Save()

        Call CloseWorkbook(STApp, STWB, STWS, STProc, FileName)

        Exit Function
ErrHandler:
        AddWorksheetList = False
    End Function

    Public Function SelectWorkSheet(ByRef ExcelWorkbook As Excel.Workbook, ByRef ExcelSheet As Excel.Worksheet, ByVal SheetName As String) As Boolean

        SelectWorkSheet = True

        If WorkSheetExisting(ExcelWorkbook, SheetName) = True Then
            ExcelSheet = ExcelWorkbook.Worksheets(SheetName)
        Else
            SelectWorkSheet = False
        End If

    End Function

    Public Function SetCellValue(ByRef ExcelSheet As Excel.Worksheet, ByVal ColNum As Long, ByVal RowNum As Long, ByVal CellValue As String, ByVal IsFormula As Boolean) As Boolean

        On Error GoTo ErrHandler

        SetCellValue = False

        If IsFormula = False Then
            ExcelSheet.Cells(RowNum, ColNum).Value = CellValue
        Else
            ExcelSheet.Cells(RowNum, ColNum).formula = CellValue
        End If

        SetCellValue = True

        Exit Function

ErrHandler:

    End Function

    Public Function GetCellValue(ByVal ExcelSheet As Excel.Worksheet, ByVal RowNum As Long, ByVal ColNum As Long, Optional ErrString As String = "ERROR!!!") As String

        On Error GoTo ErrHandler

        GetCellValue = ExcelSheet.Cells(RowNum, ColNum).Value

        Exit Function

ErrHandler:

        GetCellValue = ErrString

    End Function

    Public Function GetWorkSheetsWithNoBlankRow(ByRef ExcelWorkbook As Excel.Workbook, ByVal ItemCount As Long, ByVal Offset As Long, ByVal RowNum As Long)

        Dim ExcelWS As Excel.Worksheet

        Dim RetArr() As String
        Dim RetCounter As Long = -1

        For Each ExcelWS In ExcelWorkbook.Worksheets

            Dim LastCol As Long = GetLastColumn(ExcelWS, 1)
            Dim NoBlank As Boolean = True
            LastCol = LastCol - Offset

            If LastCol <> -1 And LastCol > ItemCount + 5 Then

                For x = (LastCol - ItemCount) + 1 To LastCol
                    If ExcelWS.Cells(RowNum, x).Value = vbNullString Then
                        NoBlank = False
                        Exit For
                    End If
                Next x

                If NoBlank = True Then
                    RetCounter = RetCounter + 1
                    ReDim Preserve RetArr(0 To RetCounter)
                    RetArr(RetCounter) = ExcelWS.Name
                End If

            End If

        Next ExcelWS

        Return RetArr

    End Function

    Public Function GetRowValues(ByRef ExcelWB As Excel.Workbook, ByVal SheetName As String, ByVal RowNum As Long)

        Dim RetArr() As String
        Dim ThisSheet As Excel.Worksheet
        Dim ColCount As Long

        ThisSheet = ExcelWB.Worksheets(SheetName)

        ColCount = GetLastColumn(ThisSheet, RowNum)
        ReDim Preserve RetArr(0 To ColCount - 1)

        For ColCounter As Long = 1 To ColCount
            RetArr(ColCounter - 1) = GetCellValue(ThisSheet, RowNum, ColCounter)
        Next ColCounter

        Return RetArr

    End Function

    Public Function GetColValues(ByRef ExcelWB As Excel.Workbook, ByVal SheetName As String, ByVal ColNum As Long)

        Dim RetArr() As String
        Dim ThisSheet As Excel.Worksheet
        Dim RowCount As Long

        ThisSheet = ExcelWB.Worksheets(SheetName)

        RowCount = GetLastRow(ThisSheet, ColNum)
        ReDim Preserve RetArr(0 To RowCount - 1)

        For RowCounter As Long = 1 To RowCount
            RetArr(RowCounter - 1) = GetCellValue(ThisSheet, RowCounter, ColNum)
        Next RowCounter

        Return RetArr

    End Function

    Public Function ClearRowValues(ByRef ExcelWs As Excel.Worksheet, ByVal RowNum As Long)

        ExcelWs.Cells(RowNum, 1).EntireRow.clear()

    End Function

    Public Function DeleteRow(ByRef ExcelWs As Excel.Worksheet, ByVal RowNum As Long)

        ExcelWs.Cells(RowNum, 1).EntireRow.delete()

    End Function

End Module
