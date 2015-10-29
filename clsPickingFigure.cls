Option Explicit

Private pPickDate As Date
Private pSheetNumber As String
Private pOperatorID As Long
Private pReasonCode As String
Private pSinglesQty As Long
Private pCasesQty As Long
Private pProductCode As String

Public Property Get pickDate() As Date
    pickDate = pPickDate
End Property

Public Property Let pickDate(Value As Date)
    pPickDate = Value
End Property

Public Property Get sheetNumber() As String
    sheetNumber = pSheetNumber
End Property

Public Property Let sheetNumber(Value As String)
    pSheetNumber = Value
End Property

Public Property Get operatorID() As Long
    operatorID = pOperatorID
End Property

Public Property Let operatorID(Value As Long)
    pOperatorID = Value
End Property

Public Property Get reasonCode() As String
    reasonCode = pReasonCode
End Property

Public Property Let reasonCode(Value As String)
    pReasonCode = Value
End Property

Public Property Get singlesQty() As Long
    singlesQty = pSinglesQty
End Property

Public Property Let singlesQty(Value As Long)
    pSinglesQty = Value
End Property

Public Property Get casesQty() As Long
    casesQty = pCasesQty
End Property

Public Property Let casesQty(Value As Long)
    pCasesQty = Value
End Property

Public Property Get productCode() As String
    productCode = pProductCode
End Property

Public Property Let productCode(Value As String)
    pProductCode = Value
End Property
