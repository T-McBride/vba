Attribute VB_Name = "DisectData"
'Disect Data
'Last Updated:
'Last Updated By:
'Notes:

Option Explicit

'Globals
Public WS As Worksheet

Public Sub disectData()

    Dim r As Range 'Active cell (start pos)
    Dim r2 As Range 'Data range (no header)
    Dim arrUnique 'Array of unique data (filter data)
    Dim strMessage
    Dim ans As Variant
    Dim value As Variant
    Dim LR As Long
    Dim LC As Long
    Dim WSName As String
    
    'Sanity check
    Set r = ActiveCell
    If IsNull(r) Or r = "" Then MsgBox ("Please select a cell containing data"): Exit Sub
    
    'Get sheet limits
    LR = getLastRow(ActiveSheet)
    LC = getLastCol(ActiveSheet)
    
    'Create a warning message
    strMessage = "You are about to create a worksheets based on the values found in column : " & _
    getColLetters(ActiveCell.Column) & vbNewLine & _
    "A new worksheet will be created for each value in this column and all matching data will be copied to this sheet" & _
    vbNewLine & "Do not continue if this is not your intent!" & vbNewLine & vbNewLine & "Do you wish to continue?"
    
    'Prompt the user
    ans = MsgBox(strMessage, vbYesNo, "Create Sheets from Data?")
    
    'Abort
    If ans <> vbYes Then Exit Sub
    
    'Logic to locate the top row
    'If row = 1 and the row below contains data
    If r.Row() = 1 And (r.Offset(1, 0) <> "" Or Not IsNull(r.Offset(1, 0))) Then
        Set r2 = Range(Selection.Offset(1, 0), Cells(LR, r.Column))
    'If row <> 1 and row above is empty
    ElseIf (Not IsNull(r) And Not r = "") And (IsNull(r.Offset(-1, 0)) Or r.Offset(-1, 0) = "") Then
        'Set the data range if the user is on the header row
        Set r2 = Range(Selection.Offset(1, 0), Cells(LR, r.Column))
    Else 'Active cell is in the data
        'Set thte data range if the user is in the data
        Set r2 = Range(Selection.End(xlUp).Offset(1, 0), Cells(LR, r.Column))
    End If
    
    'Create an array of unique values
    arrUnique = getArrayUnique(r2)

    'More sanity checks
    If IsNull(arrUnique) Then Exit Sub

    'Return the user to where they started
    Set WS = ActiveSheet
    
    'Add screen updating???....
        
    'Loop the arrray of unique values, filter, and call carveSheet
    For Each value In arrUnique
        'Enable filters and set to the value of the array
        WS.AutoFilterMode = False
        r2.Offset(-1, 0).AutoFilter Field:=1, Criteria1:=value
        'Set the name for the new worksheet
        If IsNull(value) Or value = "" Then
            WSName = "Temp"
        Else
            WSName = value
        End If
        'Call the sub to copy and paste all filtered data to a new sheet
        Call carveSheet(WSName, LR, LC)
    Next

    WS.AutoFilterMode = False
    'Return the cursor
    r.Activate
    
    'Always....
    'Screen Updating off....
    'Cleanup
    Set WS = Nothing: Set r = Nothing
End Sub



Public Sub carveSheet(WSName As String, LR As Long, LC As Long)
    'Don't like A1....
    Range("A1", Cells(LR, LC)).Copy
    'Add a new sheet after all the others
    Sheets.Add After:=Sheets(Sheets.Count)
    'Paste the values
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'Delete the sheet if it already exists
    deleteWS (WSName)
    'Rename the new sheet to the value of the filter
    renameWS (WSName)
    'Return to the data sheet
    WS.Activate
End Sub
