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

Public Property Get index(inRowIndex As Long, inColumnIndex As Long) As Variant
    index = columns.getItem(inColumnIndex).index(inRowIndex)
End Property

Public Property Let index(inRowIndex As Long, inColumnIndex As Long, inValue As Variant)
    columns.getItem(inColumnIndex).index(inRowIndex) = inValue
End Property

Public Property Set index(inRowIndex As Long, inColumnIndex As Long, inValue As Variant)
    Set columns.setItem(inColumnIndex).index(inRowIndex) = inValue
End Property

Public Property Get columnCount() As Long
    columnCount = columns.size
End Property

Public Property Get rowCount() As Long
    rowCount = pRowCount
End Property

Public Function clone(Optional makeEmpty As Boolean) As DataFrame
    Dim cloneDF As DataFrame: Set cloneDF = New DataFrame
    Dim i As Long
    For i = 0 To columnCount - 1
        cloneDF.addColumn columnByIndex(i).clone(makeEmpty)
    Next i
    Set clone = cloneDF
End Function

Public Function columnByIndex(inIndex As Long) As DataFrameColumn
    Set columnByIndex = columns.getItem(inIndex)
End Function


Public Function columnByName(inName As String) As DataFrameColumn
    Dim i As Long, tempDFC As DataFrameColumn
    For i = 0 To columnCount - 1
        Set tempDFC = columns.getItem(i)
        If tempDFC.name = inName Then
            Set columnByName = tempDFC
            Exit Function
        End If
    Next i
    Err.Raise -1, "DataFrame.columnByName()", "No column with name " & Chr(34) & inName & Chr(34) & " found in this DataFrame | DataFrame.columnByName()"
End Function

Public Function uniqueBy(inExpression As String, Optional ByVal inProgressFunctor As String) As DataFrame
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not inProgressFunctor = ""
    ProgressMod = Int(rowCount / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    
    
    Dim i As Long, j As Long, evaluation As String
    Dim al As ArrayList: Set al = New ArrayList
    Dim tempDF As DataFrame: Set tempDF = clone(True)
    
    For i = 0 To rowCount - 1
        If showProgress Then
            If i Mod ProgressMod = 0 Then DoEvents: Application.Run inProgressFunctor, i / ProgressMod
        End If
        
        evaluation = Evaluate(replaceLiterals(inExpression, i))
        If InStr(1, evaluation, "Error") <> 0 Then
            Err.Raise -1, "DataFrame.uniqueBy()", "Evaluation contained an error, Evaluation = " & Chr(34) & evaluation & Chr(34) & " - DataFrame.uniqueBy()"
        End If
        'If i Mod 5000 = 0 Then Stop
        If Not al.contains(evaluation) Then
            al.add evaluation
            For j = 0 To columnCount - 1
                tempDF.columnByIndex(j).add index(i, j)
            Next j
        End If
    Next i
    tempDF.ensureRowCountExternal
    Set uniqueBy = tempDF
End Function

Public Function KeepWhere(inExpression As String, Optional ByVal inProgressFunctor As String) As DataFrame
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not inProgressFunctor = ""
    ProgressMod = Int(rowCount / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    
    
    Dim i As Long, j As Long, evaluation As String
    Dim tempDF As DataFrame: Set tempDF = clone(True)
    For i = 0 To rowCount - 1
        If showProgress Then
            If i Mod ProgressMod = 0 Then DoEvents: Application.Run inProgressFunctor, i / ProgressMod
        End If
        
        evaluation = Evaluate(replaceLiterals(inExpression, i))
        If evaluation <> "True" And evaluation <> "False" Then
            Err.Raise "DataFrame.keepWhere", "evaluation at row " & Chr(34) & i & Chr(34) & " = " & Chr(34) & evaluation & Chr(34) & " which is not equal to a boolean - DataFrame.keepWhere"
        End If
        
        If evaluation = "True" Then
            For j = 0 To columnCount - 1
                tempDF.columnByIndex(j).add index(i, j)
            Next j
        End If
    Next i
    tempDF.ensureRowCountExternal
    Set KeepWhere = tempDF
End Function

Public Function CalculatedColumn(inName As String, inExpression As String, Optional ByVal inProgressFunctor As String) As DataFrameColumn
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not inProgressFunctor = ""
    ProgressMod = Int(rowCount / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    
    Dim i As Long
    Dim tempDFC As DataFrameColumn: Set tempDFC = New DataFrameColumn
    tempDFC.name = inName
    For i = 0 To rowCount - 1
        If showProgress Then
            If i Mod ProgressMod = 0 Then DoEvents: Application.Run inProgressFunctor, i / ProgressMod
        End If
        tempDFC.add Evaluate(replaceLiterals(inExpression, i))
    Next i
    Set CalculatedColumn = tempDFC
End Function

Public Function addBlankRows(Optional inCount As Long)
    Dim j As Long, count As Long
    
    If IsMissing(inCount) Then
        count = 1
    Else
        If inCount < 1 Then Err.Raise -1, "DataFrame.addBlankRow", "inCount parameter was " & inCount & " which is less than 1 - DataFrame.addBlankRow"
        count = inCount
    End If
    
    Dim blankArray As Variant
    ReDim blankArray(0 To count - 1)
    
    For j = 0 To columnCount - 1
            columnByIndex(j).addArray blankArray
    Next j
    rowCount = rowCount + count
    ensureRowCount count
End Function

Public Function addColumn(inColumn As DataFrameColumn, Optional inPosition As Variant)
    Dim tempDFC As DataFrameColumn: Set tempDFC = inColumn.clone
    If tempDFC.size > rowCount Then
        rowCount = tempDFC.size
        ensureRowCount rowCount
    ElseIf tempDFC.size < rowCount Then
        tempDFC.ensureRowCount rowCount
    End If
    
    If IsMissing(inPosition) Then
        columns.add tempDFC
    Else
        columns.addIndex columnIndexByIdentifier(inPosition), tempDFC
    End If
End Function

Public Function addDataFrame(inDataFrame As DataFrame, Optional inPosition As Variant)
    Dim i As Long
    If IsMissing(inPosition) Then
        For i = 0 To inDataFrame.columnCount - 1
            columns.add inDataFrame.columnByIndex(i).clone
        Next i
    Else
        Dim startIndex As Long: startIndex = columnIndexByIdentifier(inPosition)
        For i = 0 To inDataFrame.columnCount - 1
            columns.addIndex startIndex + i, inDataFrame.columnByIndex(i).clone
        Next i
    End If
    If rowCount < inDataFrame.rowCount Then rowCount = inDataFrame.rowCount
    ensureRowCount rowCount
End Function

Public Function appendColumn(inSourceColumn As DataFrameColumn, Optional inDestColumn As Variant)
    If inSourceColumn.size < 1 Then Exit Function
    If IsMissing(inDestColumn) Then
        columns.getItem(columnIndexByName(inSourceColumn.name)).append inSourceColumn
    Else
        If TypeName(inDestColumn) = "DataFrameColumn" Then
            inDestColumn.append insourcecolumjn
        Else
            columns.getItem(columnIndexByIdentifier(inDestColumn)).append inSourceColumn
        End If
    End If
    rowCount = rowCount + inSourceColumn.size
    ensureRowCount rowCount
End Function

Public Function appendDataFrame(inDataFrame As DataFrame)
    Dim i As Long
    Dim tempDFC As DataFrameColumn, tempDFC2 As DataFrameColumn
    
    If inDataFrame.columnCount = 0 Then Exit Function
    
    For i = 0 To inDataFrame.columnCount - 1
        Set tempDFC = inDataFrame.columnByIndex(i)
        If containsDFName(tempDFC.name) Then
            columnByName(tempDFC.name).append tempDFC
        Else
            Set tempDFC2 = New DataFrameColumn
            tempDFC2.name = tempDFC.name
            addColumn tempDFC2
            columnByName(tempDFC.name).append tempDFC
        End If
    Next i
    rowCount = rowCount + inDataFrame.rowCount
    ensureRowCount rowCount
End Function

Public Function ensureRowCountExternal()
    Dim max As Long
    max = maxRowCount
    rowCount = max
    ensureRowCount rowCount
End Function


Public Function read_table(inTableName As String, Optional ByVal inProgressFunctor As String)
    If inProgressFunctor = "" Then read_range Range(inTableName & "[#All]") Else read_range Range(inTableName & "[#All]"), inProgressFunctor
End Function

Public Function read_range(inRange As Range, Optional ByVal inProgressFunctor As String)
    If inProgressFunctor = "" Then read_array inRange.Value2 Else read_array inRange.Value2, inProgressFunctor
End Function

Public Function read_array(inArray As Variant, Optional ByVal inProgressFunctor As String)
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not inProgressFunctor = ""
    ProgressMod = Int(((UBound(inArray, 2) - LBound(inArray, 2) + 1) * (UBound(inArray, 1) - LBound(inArray, 1))) / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    'Above is Progress Functor Stuff
    
    Dim i As Long, j As Long, tempDFC As DataFrameColumn
    rowCount = UBound(inArray, 1) - LBound(inArray, 1)
    For j = LBound(inArray, 2) To UBound(inArray, 2)
        Set tempDFC = New DataFrameColumn
        tempDFC.construct CStr(inArray(LBound(inArray, 1), j)), UBound(inArray) - LBound(inArray)
        For i = LBound(inArray, 1) + 1 To UBound(inArray, 1)
            'Begin Progress Functor Stuff
            If showProgress Then
                If (((j - LBound(inArray, 2)) * rowCount) + i) Mod ProgressMod = 0 Then
                    DoEvents
                    Application.Run inProgressFunctor, (((j - LBound(inArray, 2)) * rowCount) + i) / ProgressMod
                End If
            End If
            'End Progress Functor Stuff
            tempDFC.add inArray(i, j)
        Next i
        columns.add tempDFC
    Next j
End Function

Public Function to_array(Optional rowLBound As Long, Optional columnLBound As Long, Optional ByVal inProgressFunctor As String) As Variant
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not inProgressFunctor = ""
    ProgressMod = Int((rowCount * columnCount) / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    'Above is progressFunctor stuff
    
    Dim i As Long, j As Long
    Dim tempDFC As DataFrameColumn
    Dim outputArray As Variant
    
    If (IsMissing(rowLBound) And Not IsMissing(columnLBound)) Or (Not IsMissing(rowLBound) And IsMissing(columnLBound)) Then
        Err.Raise -1, "DataFrame.to_array()", "One LBound defined but not the other! - DataFrame.to_array()"
    End If
    
    ReDim outputArray(rowLBound To rowLBound + rowCount, columnLBound To columnLBound + columnCount - 1)
        
    For j = LBound(outputArray, 2) To UBound(outputArray, 2)
        Set tempDFC = columnByIndex(j - LBound(outputArray, 2))
        outputArray(LBound(outputArray, 1), j) = tempDFC.name
        For i = LBound(outputArray, 1) + 1 To UBound(outputArray)
            'Start progressFunctor Stuff
            If showProgress Then
                If (((j - LBound(outputArray, 2)) * rowCount - 1) + i) Mod ProgressMod = 0 Then
                    DoEvents
                    Application.Run inProgressFunctor, (((j - LBound(outputArray, 2)) * rowCount - 1) + i) / ProgressMod
                End If
            End If
            outputArray(i, j) = tempDFC.index(i - LBound(outputArray, 1) - 1)
        Next i
    Next j
    to_array = outputArray
End Function

Public Function to_range(topLeft As Range, Optional ByVal inProgressFunctor As String) As Range
    Dim outputArray As Variant
    If inProgressFunctor = "" Then outputArray = to_array(1, 1) Else outputArray = to_array(1, 1, inProgressFunctor)
    topLeft.Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value2 = outputArray
    Set to_range = topLeft.Resize(UBound(outputArray, 1), UBound(outputArray, 2))
End Function


Public Function to_table(inTableName As String, Optional ByVal inProgressFunctor As String)
    Dim oldRange As Range: Set oldRange = Range(inTableName & "[#All]")
    oldRange.ClearContents
    oldRange.Worksheet.ListObjects.add(xlSrcRange, to_range(oldRange, inProgressFunctor), , xlYes).name = inTableName
End Function


'Private Functions

'CHECK ALIASING!!!
Private Function ensureRowCount(inRowCount)
    Dim i As Long, tempDFC As DataFrameColumn
    For i = 0 To columnCount - 1
        Set tempDFC = columns.getItem(i)
        If tempDFC.size < inRowCount Then
            tempDFC.ensureRowCount inRowCount
        ElseIf tempDFC.size > inRowCount Then
            Err.Raise -1, "DataFrame.ensureRowCount()", "Row " & Chr(34) & tempDFC.name & Chr(34) & " has more rows than parameter | DataFrame.ensureRowCount()"
        End If
    Next i
End Function

Private Function maxRowCount() As Long
    Dim j As Long, max As Long, tempDFC As DataFrameColumn
    max = 0
    For j = 0 To columnCount - 1
        Set tempDFC = columnByIndex(j)
        If max < tempDFC.size Then max = tempDFC.size
    Next j
    maxRowCount = max
End Function

Private Property Let rowCount(inRowCount As Long)
    pRowCount = inRowCount
End Property

Private Function columnIndexByIdentifier(inIdentifier As Variant) As Long
    Dim vt As Long: vt = VarType(inIdentifier)
    
    If vt = 2 Or vt = 3 Then
        indexByIdentifier = CLng(inIdentifier)
    ElseIf vt = 8 Then
        indexByIdentifier = columnIndexByName(CStr(inIdentifier))
    Else
        Err.Raise -1, "DataFrame.columnIndex", "Identifer has invalid type: " & vt & " - DataFrame.columnIndex"
    End If
End Function

Private Function columnIndexByName(inName As String) As Long
    Dim i As Long
    For i = 0 To columnCount - 1
        If columns.getItem(i).name = inName Then
            columnIndexByName = i
            Exit Function
        End If
    Next i
    Err.Raise -1, "DataFrame.columnIndexByName", "column with name: " & inName & " not found - DataFrame.columnIndexByName"
End Function

Private Function replaceLiterals(inExpression As String, inIndex As Long) As String
    Dim i As Long
    Dim expressionString As String
    expressionString = inExpression
    Dim tempDFC As DataFrameColumn
    For i = 0 To columnCount - 1
        Set tempDFC = columnByIndex(i)
        expressionString = Replace(expressionString, "{" & tempDFC.name & "}", CStr(tempDFC.index(inIndex)))
    Next i
    replaceLiterals = expressionString
End Function

Private Function containsDFName(inDataFrameName As String) As Boolean
    Dim i As Long
    For i = 0 To columnCount - 1
        If columnByIndex(i).name = inDataFrameName Then
            containsDFName = True
            Exit Function
        End If
    Next i
    containsDFName = False
End Function
