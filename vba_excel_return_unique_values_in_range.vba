
' Function returns number of distinct visible (selected with filter) items values in range in O(N) time
Function NumberOfDistinctFilteredValues(rng As Range) As Integer

    Dim dict As Object 'Declare a generic Object reference
    Set dict = CreateObject("Scripting.Dictionary") 'Late Binding of the Dictionary
    dict.RemoveAll

    'rng.SpecialCells(xlCellTypeVisible).Copy

    Dim cel As Range
    For Each cel In VisibleCells(rng).Cells
        With cel
        If .Value <> "" Then
           'Debug.Print .Address & ":" & .Value
           'Add item to VBA Dictionary
            If Not dict.Exists(.Value) Then
                dict.Add .Value, .Value
            End If
        End If
        End With
    Next cel

'    For Each Key In dict.Keys
'        Debug.Print Key
'    Next Key

    'Debug.Print dict.Count
    NumberOfDistinctFilteredValues = dict.Count
End Function

' Function returns number of distinct values in range
Function NumberOfDistinctValues(rng As Range) As Integer

    Dim dict As Object 'Declare a generic Object reference
    Set dict = CreateObject("Scripting.Dictionary") 'Late Binding of the Dictionary
    dict.RemoveAll

    'rng.SpecialCells(xlCellTypeVisible).Copy

    Dim cel As Range
    For Each cel In rng.Cells
        With cel
        If .Value <> "" Then
           'Debug.Print .Address & ":" & .Value
           'Add item to VBA Dictionary
            If Not dict.Exists(.Value) Then
                dict.Add .Value, .Value
            End If
        End If
        End With
    Next cel

'    For Each Key In dict.Keys
'        Debug.Print Key
'    Next Key

    'Debug.Print dict.Count
    NumberOfDistinctValues = dict.Count
End Function

Private Function VisibleCells(rng As Range) As Range
    Dim r As Range
    For Each r In rng
        If r.EntireRow.Hidden = False Then
            If VisibleCells Is Nothing Then
                Set VisibleCells = r
            Else
                Set VisibleCells = Union(VisibleCells, r)
            End If
        End If
    Next r
End Function


