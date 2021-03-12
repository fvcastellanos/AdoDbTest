VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public OrderId As Long
Public ItemId As Long
Public UnitPrice As Double
Public Quantity As Integer
Public DiscountAmount As Double
Public DiscountTaxable As Double
Public ItemText As String
Public ItemDescription As String

