Attribute VB_Name = "Basic"
'Basic
'Last Updated:
'Last Updated By:
'Notes:

Option Explicit

Public Function genericEH(errNum As Long, errDesc As String)
'Common error handler
    On Error Resume Next
    Debug.Print "Error: " & errNum & " : " & errDesc
End Function

Public Function getLastRow(WS As Worksheet) As Long
'Returns the last used row from the passed worksheet
    On Error Resume Next
    getLastRow = WS.Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If getLastRow = 0 Then getLastRow = 1
End Function

Public Function getLastCol(WS As Worksheet) As Long
'Returns the last used row from the passed worksheet
    On Error Resume Next
    getLastCol = WS.Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If getLastCol = 0 Then getLastCol = 1
End Function

Public Sub renameWS(WSName As String)
'Rename a sheet
    On Error Resume Next
    ActiveSheet.Name = WSName
End Sub

Public Sub deleteWS(WSName)
'Delete a sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(WSName).Delete
    Application.DisplayAlerts = True
End Sub

Public Function getColLetters(colNumber As Integer)
'Return the Alpha name for a column from the column number
    On Error Resume Next
    'Logic for past ZZ, utilizes numeric value of true false
    If colNumber > 702 Then
        getColLetters = Left(Cells(1, colNumber).Address(False, False), 2 - (colNumber > 26))
    Else
        getColLetters = Left(Cells(1, colNumber).Address(False, False), 1 - (colNumber > 26))
    End If
End Function

