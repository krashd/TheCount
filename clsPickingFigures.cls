Option Explicit
'first attempt at a class collection.

Private mcolPickingFigures As Collection

Private Sub Class_Initialize()
    Set mcolPickingFigures = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolPickingFigures = Nothing
End Sub

Public Sub Add(cPickingFigure As clsPickingFigure)
'add a picking figure to the collection.
    mcolPickingFigures.Add cPickingFigure
End Sub

Public Property Get Count() As Long
'return the number of pickingfigures entered.
    Count = mcolPickingFigures.Count
End Property

Public Property Get Item(NameOrNumber As Variant) As clsPickingFigure
'return a particular picking figure.
    Set Item = mcolPickingFigures(NameOrNumber)
End Property

Property Get Items() As Collection
    Set Items = mcolPickingFigures
End Property
