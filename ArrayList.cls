VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_CAPACITY As Long = 10
Private Const MAX_ARRAY_SIZE As Long = 2147483647 - 8

Private elementData As Variant
Private pSize As Long




Private Sub Class_Initialize()
    ReDim elementData(0 To DEFAULT_CAPACITY)
    pSize = 0
End Sub

Public Function trimToSize()
    If pSize < UBound(elementData) + 1 Then
        elementData = copyOf(elementData, pSize)
    End If
End Function

Public Function ensureCapacity(inMinCapacity As Long)
    Dim minExpand As Long
    
    minExpand = DEFAULT_CAPACITY
    
    If inMinCapacity > minExpand Then
        ensureExplicitCapacity inMinCapacity
    End If
End Function

Private Function ensureCapacityInternal(inMinCapacity As Long)
    inMinCapacity = WorksheetFunction.max(DEFAULT_CAPACITY, inMinCapacity)
    ensureExplicitCapacity inMinCapacity
End Function

Private Function ensureExplicitCapacity(inMinCapacity As Long)
    
    If inMinCapacity - (UBound(elementData) + 1) > 0 Then
        grow inMinCapacity
    End If
End Function

Private Function grow(inMinCapacity As Long)
    Dim oldCapacity As Long, newCapacity As Long
    oldCapacity = UBound(elementData) + 1
    newCapacity = oldCapacity + (shr(oldCapacity, 1))
    If newCapacity - inMinCapacity < 0 Then
        newCapacity = inMinCapacity
    End If
    
    If newCapacity - MAX_ARRAY_SIZE > 0 Then
        newCapacity = hugeCapacity(inMinCapacity)
    End If
    
    elementData = copyOf(elementData, newCapacity)
End Function

Private Function hugeCapacity(inMinCapacity As Long)
    If minCapacity < 0 Then
        Err.Raise -1, "ArrayList.hugeCapacity", "OutOfMemoryError ArrayList.hugeCapacity"
    End If
    
    If minCapacity > MAX_ARRAY_SIZE Then
        hugeCapacity = 2147483647
    Else
        hugeCapacity = MAX_ARRAY_SIZE
    End If
End Function

'bit shift function taken from http://www.excely.com/excel-vba/bit-shifting-function.shtml
Private Function shr(ByVal Value As Long, ByVal Shift As Byte) As Long
    Dim i As Byte
    shr = Value
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function

Public Function size() As Long
    size = pSize
End Function

Public Function isEmpty() As Boolean
    isEmpty = pSize = 0
End Function

Public Function contains(o As Variant) As Boolean
    contains = indexOf(o) >= 0
End Function

Public Function indexOf(o As Variant) As Long
    Dim i As Long
    
    If IsObject(o) Then
        For i = 0 To pSize - 1
            If IsObject(elementData(i)) Then
                If elementData(i) Is o Then
                    indexOf = i
                    Exit Function
                End If
            End If
        Next i
    Else
        For i = 0 To pSize - 1
            If Not IsObject(elementData(i)) Then
                If elementData(i) = o Then
                    indexOf = i
                    Exit Function
                End If
            End If
        Next i
    End If
    indexOf = -1
End Function

Public Function lastIndexOf(o As Variant) As Long
    Dim i As Long
    
    If IsObject(o) Then
        For i = pSize - 1 To 0 Step -1
            If IsObject(elementData(i)) Then
                If elementData(i) Is o Then
                    indexOf = i
                    Exit Function
                End If
            End If
        Next i
    Else
        For i = pSize - 1 To 0 Step -1
            If Not IsObject(elementData(i)) Then
                If elementData(i) = o Then
                    indexOf = i
                    Exit Function
                End If
            End If
        Next i
    End If
    indexOf = -1
End Function

Public Function clone() As ArrayList
    'Err.Raise -1, "ArrayList.clone()", "ArrayList.clone() must be re-Implemented"
    
    Dim v As ArrayList
    Set v = New ArrayList
    v.addArray toArray
    Set clone = v
End Function

Public Function toArray() As Variant
    toArray = copyOf(elementData, pSize)
End Function

Public Function getItem(inIndex As Long)
    rangeCheck inIndex
    
    If IsObject(elementData(inIndex)) Then Set getItem = elementData(inIndex) Else getItem = elementData(inIndex)
    
    'If IsObject(elementData(inIndex)) Then
    '    Set getItem = elementData(inIndex)
    'Else
    '    getItem = elementData(inIndex)
    'End If
End Function

Private Function rangeCheck(inIndex As Long)
    If inIndex >= pSize Then
        Err.Raise -1, "ArrayList.rangeCheck", "IndexOutOfBounds ArrayList.rangeCheck"
    End If
End Function

Private Function rangeCheckForAdd(inIndex As Long)
    If index > size Or index < 0 Then
        Err.Raise -1, "ArrayList.rangeCheckForAdd", "IndexOutOfBounds ArrayList.rangeCheckForAdd"
    End If
End Function

Public Function setItem(inIndex As Long, inElement As Variant) As Variant
    Dim oldValue As Variant
    
    rangeCheck inIndex
    
    If IsObject(elementData(inIndex)) Then Set oldValue = elementData(inIndex) Else oldValue = elementData(inIndex)
    If IsObject(inElement) Then Set elementData(inIndex) = inElement Else elementData(inIndex) = inElement
    If IsObject(oldValue) Then Set setItem = oldValue Else setItem = oldValue
    
    'oldValue = elementData(inIndex)
    'elementData(inIndex) = inElement
    'setItem = oldValue
End Function

Public Function add(e As Variant) As Boolean
    ensureCapacityInternal (pSize + 1)
    
    If IsObject(e) Then Set elementData(pSize) = e Else elementData(pSize) = e
    'If IsObject(e) Then
    '    Set elementData(pSize) = e
    'Else
    '    elementData(pSize) = e
    'End If
    pSize = pSize + 1
    add = True
End Function

Public Function addIndex(inIndex As Long, e As Variant)
    rangeCheckForAdd inIndex
    
    ensureCapacityInternal pSize + 1
    arrayCopy elementData, inIndex, elementData, inIndex + 1, pSize - index
    If IsObject(element) Then Set elementData(inIndex) = element Else elementData(inIndex) = element
    'elementData(inIndex) = element
    pSize = pSize + 1
End Function

Public Function remove(inIndex As Long) As Variant
    Dim oldValue As Variant
    Dim numMoved As Long
    rangeCheck inIndex
    
    If IsObject(elementData(inIndex)) Then Set oldValue = elementData(inIndex) Else oldValue = elementData(inIndex)
    'oldValue = elementData(inIndex)
    
    numMoved = pSize - inIndex - 1
    
    If numMoved > 0 Then
        arrayCopy elementData, inIndex + 1, elementData, inIndex, numMoved
    End If
    elementData(pSize - 1) = Null
    pSize = pSize - 1
    remove = oldValue
End Function

Public Function removeItem(o As Variant) As Boolean
    Dim oIndex As Long
    
    oIndex = indexOf(o)
    If oIndex = -1 Then
        removeItem = False
        Exit Function
    Else
        fastRemove oIndex
        removeItem = True
        Exit Function
    End If
        
    
    'Dim i As Long
   '
   '
   '
    'If IsObject(o) Then
    '    For i = 0 To pSize - 1
    '        If IsObject(elementData(i)) Then
    '            If elementData(i) Is o Then
    '                fastRemove i
    '                removeItem = True
    '                Exit Function
    '            End If
    '        End If
    '    Next i
    'Else
    '    For i = 0 To pSize - 1
    '        If Not IsObject(elementData(i)) Then
    '
   '
   ' End If
   '
   ' If o = Null Then
   '     For i = 0 To pSize - 1
   '         If elementData(i) = Null Then
   '             fastRemove inIndex
   '             removeItem = True
   '             End Function
   '         End If
   '     Next i
   ' Else
   '     For i = 0 To pSize - 1
   '         If elementData(i) = o Then
   '             fastRemove inIndex
   '             removeItem = True
   '             End Function
   '         End If
   '     Next i
   ' End If
End Function

Private Function fastRemove(inIndex As Long)
    Dim numMoved As Long
    numMoved = pSize - inIndex - 1
    If numMoved > 0 Then
        arrayCopy elementData, inIndex + 1, elementData, inIndex, numMoved
    End If
End Function

Public Function clear()
    Dim i As Long
    
    For i = 0 To pSize - 1
        elementData(i) = Null
    Next i
    pSize = 0
End Function

Public Function addArray(inArray As Variant) As Boolean
    Dim numNew As Long
    numNew = UBound(inArray) - LBound(inArray) + 1
    ensureCapacityInternal pSize + numNew
    arrayCopy inArray, LBound(inArray), elementData, pSize, numNew
    pSize = pSize + numNew
    addArray = numNew <> 0
End Function

Public Function addArrayIndex(inIndex As Long, inArray As Variant) As Boolean
    Dim numNew As Long
    Dim numMoved As Long
    rangeCheckForAdd inIndex
    
    numNew = UBound(inArray) - LBound(inArray) + 1
    ensureCapacityInternal pSize + numNew
    numMoved = pSize - inIndex
    If numMoved > 0 Then
        arrayCopy elementData, inIndex, elementData, inIndex + numNew, numMoved
    End If
    arrayCopy inArray, LBound(inArray), elementData, inIndex, numNew
    adarrayindex = numNew <> 0
End Function

Public Function removeRange(inFromIndex As Long, inToIndex As Long)
    Dim numMoved As Long, newSize As Long, i As Long
    numMoved = pSize - inToIndex
    arrayCopy elementData, inToIndex, elementData, inFromIndex, numMoved
    newSize = pSize - (inToIndex - inFromIndex)
    For i = newSize To pSize - 1
        elementData(i) = Null
    Next i
    pSize = newSize
End Function

Private Function outOfBoundsMsg(inIndex As Long) As String
    outOfBoundsMsg = "Index: " & inIndex & ", Size: " & pSize
End Function

Public Function replaceAll(inItem As Variant, inNewItem As Variant)
    Dim i As Long
    
    If IsObject(o) Then
        For i = 0 To pSize - 1
            If IsObject(elementData(i)) Then
                If elementData(i) Is o Then
                    setItem i, inNewItem
                End If
            End If
        Next i
    Else
        For i = 0 To pSize - 1
            If Not IsObject(elementData(i)) Then
                If elementData(i) = o Then
                    setItem i, inNewItem
                End If
            End If
        Next i
    End If
End Function

Public Function sort(Optional decreasing As Boolean)
    Err.Raise -1, "ArrayList.sort", "ArrayList.sort() not implemented yet"
End Function


Private Function copyOf(inArray As Variant, inSize As Long) As Variant
    Dim i As Long
    Dim ouputArray As Variant
    
    ReDim outputArray(0 To inSize - 1)
    
    If LBound(inArray) <> 0 Then
        Err.Raise -1, "ArrayList.copyOf", "inArray Parameter not indexed at 0"
    End If
    
    For i = 0 To WorksheetFunction.Min(inSize - 1, UBound(inArray))
        If IsObject(inArray(i)) Then
            Set outputArray(i) = inArray(i)
        Else
            outputArray(i) = inArray(i)
        End If
    Next i
    copyOf = outputArray
End Function

Private Function arrayCopy(sourceArr As Variant, sourcePos As Long, ByRef destArr As Variant, destPos As Long, inLen As Long)
    Dim i As Long
    
    For i = 0 To inLen - 1
        If IsObject(sourceArr(i + sourcePos)) Then
            Set destArr(i + destPos) = sourceArr(i + sourcePos)
        Else
            destArr(i + destPos) = sourceArr(i + sourcePos)
        End If
        'destArr(i + destPos) = sourceArr(i + sourcePos)
    Next i
End Function


