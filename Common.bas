Attribute VB_Name = "Common"
Public Sub PrintArray(Data)

Dim DataRows As Integer
Dim DataCols As Integer

DataRows = UBound(Data, 1)
DataCols = UBound(Data, 2)

i = 0
j = 0

    For i = LBound(Data, 1) To UBound(Data, 1)
        For j = LBound(Data, 2) To UBound(Data, 2)
        If j = LBound(Data, 2) Then Debug.Print "[" & i & "]";
        Debug.Print "[" & j & "]" & Data(i, j) & "[" & j & "]" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab;

        Next j
        Debug.Print " "
    Next i

End Sub


