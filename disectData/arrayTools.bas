Attribute VB_Name = "ArrayTools"
'Array Tools
'Last Updated:
'Last Updated By:
'Notes:

Option Explicit

Public Function chkArray(arr() As String, val As String) As Boolean
'Returns true if the passed value is found in the passed array
    On Error Resume Next
    Dim l As Long
    For l = LBound(arr) To UBound(arr)
        If val = arr(l) Then
            chkArray = True
            Exit Function
        End If
    Next
End Function

Public Function getArrayUnique(rng As Range) As Variant
'Creates an array of unique values from the contents of the passed range
    
    'ToDo:  Ignore empty and space???
    
    On Error GoTo LEH:
    'Variables
    Dim b As Boolean
    Dim arr() As String
    Dim r As Range
    'Init the array
    ReDim arr(0 To 0) As String
    'Loop
    For Each r In rng
        If b = True Then
            'Check if the value already exists in the array
            If chkArray(arr(), r.value) = False Then
                'Expand the array
                ReDim Preserve arr(0 To UBound(arr) + 1)
                'Load the new value
                arr(UBound(arr)) = r.value
            End If
        Else
            'First pass only
            arr(UBound(arr)) = r.value
            b = True
        End If
    Next
    'Set the output
    getArrayUnique = arr
    Exit Function
LEH:
    Call genericEH(Err.Number, Err.Description)
End Function
