' Function returns number of distinct visible (selected with filter) items values in range in O(N) time
Function NumberOfDistinctValues(rng As Range) As Integer
    Dim Ws As Worksheet
    Set Ws = rng.Worksheet
    Dim DbExtract, DuplicateRecords As Worksheet
    Set DbExtract = ThisWorkbook.Sheets(Ws.Name)
    Set DuplicateRecords = ThisWorkbook.Sheets(Ws.Name)

    Dim dict As Object 'Declare a generic Object reference
    Set dict = CreateObject("Scripting.Dictionary") 'Late Binding of the Dictionary

    rng.SpecialCells(xlCellTypeVisible).Copy

    Dim cel As Range
    For Each cel In rng.SpecialCells(xlCellTypeVisible).Cells
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


