Option Explicit

Sub createInsertSql()
    Dim newbook As Workbook
    Dim currentCell As Range
    
    '前処理
    Dim srcSheet As Worksheet
    Set srcSheet = ActiveSheet
    
    Dim targetRange As Range
    Set targetRange = srcSheet.UsedRange

    '新しいSheet作成
    Dim ws As Worksheet, flag As Boolean
    Dim NewWorkSheet As Worksheet
    For Each ws In Worksheets
        If ws.Name = "CREATE文" Then
            flag = True
            Set NewWorkSheet = ws
        End If
    Next ws
    If flag = False Then
        Set NewWorkSheet = Worksheets.Add()
        NewWorkSheet.Name = "CREATE文"
    End If
    
    'INSERT文の前半
    Dim head As String
    head = "CREATE TABLE `" & srcSheet.Name & "` ("
    
    Dim first As Boolean
    first = True
    
    Dim currentColumnIndex As Integer
    Dim currentRowIndex As Long
    Dim col As Integer
    Dim maxLength As Integer
    col = 1
    For currentColumnIndex = 1 To targetRange.Columns.Count
        NewWorkSheet.Cells(currentColumnIndex, col).Value = head
        maxLength = 0
        head = ""
        If (first) Then
            first = False
            col = 2
            head = head & "  `"
        Else
            head = head & ", `"
        End If
        Set currentCell = srcSheet.Cells(1, currentColumnIndex)
        head = head & currentCell.Value
        For currentRowIndex = 2 To targetRange.Rows.Count
            If maxLength < Len(srcSheet.Cells(currentRowIndex, currentColumnIndex)) Then
                maxLength = Len(srcSheet.Cells(currentRowIndex, currentColumnIndex))
            End If
        Next
        head = head & "' varchar(" & maxLength & ") NULL"
    Next
    NewWorkSheet.Cells(currentColumnIndex, 1).Value = ")"
	MsgBox "CREATE TABLEが完成しました。", vbInformation
    
End Sub
