VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFrameColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private data As ArrayList
Private pName As String


Private Sub Class_Initialize()
    Set data = New ArrayList
End Sub

Public Function construct(inName As String, inStartingLength As Long)
    name = inName
    data.ensureCapacity inStartingLength
End Function

Public Property Get index(inIndex As Long) As Variant
    index = data.getItem(inIndex)
End Property

Public Property Let index(inIndex As Long, inValue As Variant)
    data.setItem inIndex, inValue
End Property

Public Property Set index(inIndex As Long, inValue As Variant)
    data.setItem inIndex, inValue
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(inName As String)
    pName = inName
End Property

Public Property Get size() As Long
    size = data.size
End Property

Public Function clone(Optional makeEmpty As Boolean) As DataFrameColumn
    Dim cloneDFC As DataFrameColumn: Set cloneDFC = New DataFrameColumn
    cloneDFC.name = name
    If (Not makeEmpty) And (Not data.isEmpty) Then cloneDFC.setData data.clone
    Set clone = cloneDFC
End Function

Public Function append(inDataFrameColumn As DataFrameColumn)
    If inDataFrameColumn.size > 0 Then data.addArray inDataFrameColumn.dataArray
End Function

Public Function ensureRowCount(inRowCount)
    If data.size < inRowCount Then
        Dim tempArray As Variant
        ReDim tempArray(0 To (inRowCount - data.size) - 1)
        data.addArray tempArray
    End If
End Function

Public Function add(inValue As Variant)
    data.add inValue
End Function

Public Function addArray(inArray As Variant)
    If UBound(inArray) - LBound(inArray) >= 0 Then data.addArray inArray
End Function

Public Function Map(ByVal delegate As String, Optional ByVal progressFunctor As String)
    Dim showProgress As Boolean, ProgressMod As Long
    showProgress = Not progressFunctor = ""
    ProgressMod = Int(size / 100)
    If ProgressMod < 1 Then ProgressMod = 1
    
    Dim i As Long
    For i = 0 To size - 1
        If showProgress Then
            If i Mod ProgressMod = 0 Then DoEvents: Application.Run progressFunctor, i / ProgressMod
        End If
        data.setItem i, Application.Run(delegate, data.getItem(i))
    Next i
End Function

Public Function mapped(ByVal delegate As String, Optional ByVal progressFunctor As String) As DataFrameColumn
    Dim cloneDFC As DataFrameColumn: Set cloneDFC = clone
    cloneDFC.name = delegate & "(" & cloneDFC.name & ")"
    If IsMissing(progressFunctor) Then cloneDFC.Map delegate Else cloneDFC.Map delegate, progressFunctor
    Set mapped = cloneDFC
End Function


'Boo bad functions
Public Function setData(inData As ArrayList)
    Set data = inData
End Function

Public Function dataArray() As Variant
    dataArray = data.toArray
End Function
