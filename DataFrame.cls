VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private columns As ArrayList
Private pRowCount As Long

Private Sub Class_Initialize()
    Set columns = New ArrayList
    pRowCount = 0
End Sub


'Properties

Public Property Get index(inRow As Long, inColumn As Long) As Variant
    index = columnByIndex(inColumn).index(inRow)
End Property

Public Property Let index(inRow As Long, inColumn As Long, inValue As Variant)
    columnByIndex(inColumn).index(inRow) = inValue
End Property

Public Property Set index(inRow As Long, inColumn As Long, inValue As Variant)
    columnByIndex(inColumn).index(inRow) = inValue
End Property

Public Property Get rowCount() As Long
    rowCount = pRowCount
End Property

Private Property Let rowCount(inRowCount As Long)
    pRowCount = inRowCount
End Property

Public Property Get columnCount() As Long
    columnCount = columns.size
End Property

'Column Getters

Public Function columnByName(inName As String) As DataFrameColumn
    Dim j As Long
    For j = 0 To columnCount - 1
        If columns.getItem(j).name = inName Then
            Set columnName = columns.getItem(j)
            Exit Function
        End If
    Next j
    Err.Raise -1, "DataFrame.columnByName", "Column with name: " & inName & " not found - DataFrame.columnByName"
End Function

Public Function columnByIndex(inIndex As Long) As DataFrameColumn
    Set columnByIndex = columns.getItem(inIndex)
End Function

Public Function column(inIdentifier As Variant)
    Set column = columnByIndex(columnIndex(inIdentifier))
End Function

Private Function columnIndex(inIdentifier)
    Dim vt As Long
    vt = VarType(inIdentifier)
    If vt = 8 Then
        columnIndex = columnIndexByName(CStr(inIdentifier))
        Exit Function
    ElseIf vt = 3 Or vt = 2 Then
        columnIndex = CLng(inIdentifier)
        Exit Function
    End If
    Err.Raise -1, "DataFrame.columnIndex", "Identifer has invalid type: " & vt & " - DataFrame.columnIndex"
End Function

Private Function columnIndexByName(inName As String) As Long
    Dim j As Long
    For j = 0 To columnCount - 1
        If columns.getItem(j).name = inName Then
            columnIndexByName = j
            Exit Function
        End If
    Next j
    Err.Raise -1, "DataFrame.columnIndexByName", "column with name: " & inName & " not found - DataFrame.columnIndexByName"
End Function


'Add to DataFrame

Public Function addColumn(inColumn As DataFrameColumn)
    columns.add inColumn
    rowCount = WorksheetFunction.max(rowCount, inColumn.size)
    ensureRowCount
End Function

Public Function appendColumn(inIdentifier As Variant, inColumn As DataFrameColumn)
    column(inIdentifier).appendColumn inColumn
    rowCount = rowCount + inColumn.size
    ensureRowCount
End Function

Public Function addDataFrame(inDataFrame As DataFrame)
    Dim j As Long
    For j = 0 To inDataFrame.columnCount
        addColumn inDataFrame.columnByIndex(j)
    Next j
    rowCount = worksheetfunctions.max(rowCount, inDataFrame.rowCount)
    ensureRowCount
End Function

Public Function appendDataFrame(inDataFrame As DataFrame)
    Dim j As Long
    If Not hasSameProfile(inDataFrame) Then
        Err.Raise -1, "DataFrame.appendDataFrame", "Other DataFrame must have the same profile - DataFrame.appendDataFrame"
    End If
    For j = 0 To columnCount - 1
        columnByIndex(j).append inDataFrame.columnByIndex(j)
    Next j
    rowCount = rowCount + inDataFrame.rowCount
End Function

Public Function addBlankRows(inBlankRowCount As Long)
    rowCount = rowCount + inBlankRowCount
    ensureRowCount
End Function

Public Function addRow(inArray As Variant)
    Dim j As Long
    If UBound(inArray) - LBound(inArray) + 1 <> columnCount Then
        Err.Raise -1, "DataFrame.addRow", "InArray as incorrect number of elements DataFrame.addRow"
    End If
    For j = 0 To columnCount - 1
        columnByIndex(j).add inArray(LBound(inArray) + j)
    Next j
End Function


'Data Manipulation

Public Function swapRows(rowIndex1 As Long, rowIndex2 As Long)
    Dim j As Long
    Dim hold As Variant
    For j = 0 To columnCount - 1
        hold = index(rowIndex1, j)
        index(rowIndex1, j) = index(rowIndex2, j)
        index(rowIndex2, j) = hold
    Next j
End Function

Public Function swapColumns(columnIdentifier1 As Variant, columnIdentifier2 As Variant)
    Dim colInd1 As Long, colInd2 As Long
    colInd1 = columnIndex(columnIdentifier1)
    colInd2 = columnIndex(columnIdentifier2)
    columns.setItem colInd1, columns.setItem(colInd2, columns.getItem(colInd1))
End Function

'Dataframe Cloning

Public Function clone() As DataFrame
    Dim j As Long
    Dim cl As DataFrame: cl = New DataFrame
    For j = 0 To columnCount - 1
        cl.addColumn columnByIndex(j).clone
    Next j
End Function

Public Function cloneProfile() As DataFrame
    Dim j As Long
    Dim cl As DataFrame: cl = New DataFrame
    For j = 0 To columnCount - 1
        cl.addColumn columnByIndex(j).cloneEmpty
    Next j
End Function



'Validation

Private Function ensureRowCount()
    Dim j As Long
    For j = 0 To columnCount - 1
        ensureRowCountHelper columnByIndex(j)
    Next j
End Function

Private Function ensureRowCountHelper(inColumn As DataFrameColumn)
    Dim blankArray As Variant
    Dim lenToAdd As Long: lenToAdd = (rowCount - inColumn.size) - 1
    If lenToAdd >= 0 Then
        ReDim blankArray(0 To lenToAdd)
        inColumn.addArray blankArray
    End If
End Function

Private Function hasSameProfile(inDataFrame As DataFrame) As Boolean
    Dim j As Long
    If columnCount <> inDataFrame.columnCount Then
        hasSameProfile = False
        Exit Function
    End If
    
    For j = 0 To columnCount - 1
        If columnByIndex(j).name <> inDataFrame.columnByIndex(j).name Then
            hasSameProfile = False
            Exit Function
        End If
    Next j
End Function

'Higher Order

Private Function eachColumn0(ByVal delegate As String)
    Dim j As Long
    For j = 0 To columnCount - 1
        Dim clm As DataFrameColumn: Set clm = columnByIndex(j)
        Application.Run delegate, clm
    Next j
End Function

Private Function eachColumn1(ByVal delegate As String, arg1 As Variant)
    Dim j As Long
    For j = 0 To columnCount - 1
        Application.Run delegate, columnByIndex(j), arg1
    Next j
End Function

Private Function eachColumn2(ByVal delegate As String, arg1 As Variant, arg2 As Variant)
    Dim j As Long
    For j = 0 To columnCount - 1
        Application.Run delegate, columnByIndex(j), arg1, arg2
    Next j
End Function

Private Function eachColumn3(ByVal delegate As String, arg1 As Variant, arg2 As Variant, arg3 As Variant)
    Dim j As Long
    For j = 0 To columnCount - 1
        Application.Run delegate, columnByIndex(j), arg1, arg2, arg3
    Next j
End Function


'Input

Public Function read_table(inTableName As String)
    read_range Range(inTableName & "[#All]")
End Function

Public Function read_range(inRange As Range)
    read_array inRange.Value2
End Function

Public Function read_array(inArray As Variant)
    Dim i As Long, j As Long
    Dim currentColumn As DataFrameColumn
    
    For j = 1 To UBound(inArray, 2)
        Set currentColumn = New DataFrameColumn
        currentColumn.construct inArray(1, j), UBound(inArray, 1) - 1
        For i = 2 To UBound(inArray, 1)
            currentColumn.add inArray(i, j)
        Next i
        columns.add currentColumn
    Next j
    rowCount = UBound(inArray, 1) - 1
End Function


'Output

Public Function to_array() As Variant
    Dim i As Long, j As Long
    Dim outputArray As Variant
    Dim currentColumn As DataFrameColumn
    
    ReDim outputArray(1 To rowCount + 1, 1 To columns.size)
    
    For j = 0 To columnCount - 1
        Set currentColumn = columnByIndex(j)
        outputArray(1, j + 1) = currentColumn.name
        For i = 0 To currentColumn.size - 1
            outputArray(i + 2, j + 1) = currentColumn.index(i)
        Next i
    Next j
    to_array = outputArray
End Function

Public Function to_range(inTopLeft As Range)
    Dim outputArray As Variant
    outputArray = Me.to_array
    inTopLeft.Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value2 = outputArray
End Function
