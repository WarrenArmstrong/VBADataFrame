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


'Constructor

Public Function construct(inColumnName, inLength As Long)
    data.ensureCapacity inLength
    pName = inColumnName
End Function


'Cloners

Public Function clone() As DataFrameColumn
    Dim cl As DataFrameColumn: Set cl = New DataFrameColumn
    cl.name = Me.name
    cl.setData data.clone
    Set clone = cl
End Function

Public Function cloneEmpty() As DataFrameColumn
    Dim cl As DataFrameColumn: Set cl = New DataFrameColumn
    cl.name = Me.name
    Set clone = cl
End Function


'Properties

Public Property Let name(inNewColumnName As String)
    pName = inNewColumnName
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get index(inIndex As Long) As Variant
    index = data.getItem(inIndex)
End Property

Public Property Let index(inIndex As Long, inValue As Variant)
    data.setItem inIndex, inValue
End Property

Public Property Set index(inIndex As Long, inValue As Variant)
    data.setItem inIndex, inValue
End Property


'Data Functions

Public Function setData(inData As ArrayList)
    Set data = inData
End Function

Public Function add(inItem As Variant)
    data.add inItem
End Function

Public Function size() As Long
    size = data.size
End Function

Public Function addArray(inArray As Variant)
    data.addArray inArray
End Function

Public Function appendColumn(inColumn As DataFrameColumn)
    data.addArray inColumn.data.toArray
End Function

'Higher Order

Public Function Map(ByVal delegate As String)
    Dim i As Long
    For i = 0 To data.size - 1
        data.setItem i, Application.Run(delegate, index(i))
    Next i
End Function

Public Function Mapped(ByVal delegate As String) As DataFrameColumn
    Dim cl As DataFrameColumn
    Set cl = Me.clone
    cl.name = delegate & "(" & cl.name & ")"
    cl.Map delegate
    Set Mapped = cl
End Function
