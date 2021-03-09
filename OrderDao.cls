VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetOrderDetails(ByVal OrderId As Long) As Object

  Dim list As Object
  Set list = CreateObject("System.Collections.ArrayList")
  
  Set cn = New ADODB.Connection
  cn.Open BuildDatabaseConnectionString()

  Set cmd = New ADODB.Command
  
  With cmd
    .ActiveConnection = cn
    .Prepared = True
    .CommandText = BuildOrderDetailsQuery()
    .Parameters.Append .CreateParameter("OrderTransactions.OrderID", adInteger, adParamInput, 18, OrderId)
  End With
  
  Set rsQuery = cmd.Execute
  
  Do While Not rsQuery.EOF
  
    Debug.Print "MenuItemID: "; rsQuery!MenuItemId; " Item: "; rsQuery!MenuItemText
    
    Set Item = BuildOrderDetail(rsQuery)
    list.Add (Item)
    
    rsQuery.MoveNext
  Loop
  
  rsQuery.Close
  cn.Close
  
  Set rsQuery = Nothing
  Set cmd = Nothing
  Set db = Nothing
  
  Set GetOrderDetails = list
  
End Function


Private Function BuildOrderDetailsQuery() As String

  Dim query As String
  query = _
    "SELECT OrderTransactions.OrderID, OrderTransactions.MenuItemID, OrderTransactions.MenuItemUnitPrice, OrderTransactions.Quantity, " _
       + "   OrderTransactions.ExtendedPrice, OrderTransactions.DiscountAmount, OrderTransactions.DiscountBasis, OrderTransactions.DiscountTaxable," _
       + "   OrderTransactions.TransactionStatus, MenuItems.MenuItemText, MenuItems.MenuItemDescription" _
      + " FROM MenuItems INNER JOIN OrderTransactions ON MenuItems.MenuItemID = OrderTransactions.MenuItemID" _
      + " WHERE OrderTransactions.OrderID = ?"

  BuildOrderDetailsQuery = query

End Function

Private Function BuildDatabaseConnectionString() As String

  BuildDatabaseConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=" + GetDatabasePath()
                
End Function

Private Function GetDatabasePath() As String

  GetDatabasePath = Environ$("ADELO_DATABASE_PATH")
End Function

Private Function BuildOrderDetail(ByVal record As ADODB.Recordset) As OrderDetail

  Set Item = New OrderDetail
  
  Item.DiscountTaxable = 0
  Item.DiscountAmount = 0
  Item.ItemDescription = ""
  
  Item.OrderId = record!OrderId
  Item.ItemId = record!MenuItemId
  Item.Quantity = record!Quantity
  Item.UnitPrice = record!MenuItemUnitPrice
  
  If record!DiscountAmount <> Null Then
    Item.DiscountAmount = record!DiscountAmount
  End If
  
  If record!DiscountTaxable <> Null Then
    Item.DiscountTaxable = record!DiscountTaxable
  End If
  
  Item.ItemText = record!MenuItemText
  
  If record!MenuItemDescription <> Null Then
    Item.ItemDescription = record!MenuItemDescription
  End If
  
  Set BuildOrderDetail = Item
  
End Function


