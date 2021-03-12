VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvoiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ItemType As String
Public UnitType As String
Public Quantity As String
Public Description As String
Public lineNumber As String
Public UnitPrice As String
Public Price As String
Public Discount As String
Public Total As String
Public TaxName As String
Public AffectedUnit As String
Public AffectedQuantity As String
Public BaseAmount As String
Public Tax As String

