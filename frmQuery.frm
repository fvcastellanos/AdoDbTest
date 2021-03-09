VERSION 5.00
Begin VB.Form frmQuery 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection

Private Sub Command1_Click()
  Set cn = New ADODB.Connection
  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
                & "Data Source=C:\AdeloDB\Standard\adResDemo.mdb"
                

  Dim query As String
  query = "SELECT OrderTransactions.OrderID, OrderTransactions.MenuItemID, OrderTransactions.MenuItemUnitPrice, OrderTransactions.Quantity, " _
  + "   OrderTransactions.ExtendedPrice, OrderTransactions.DiscountAmount, OrderTransactions.DiscountBasis, OrderTransactions.DiscountTaxable," _
  + "   OrderTransactions.TransactionStatus, MenuItems.MenuItemText, MenuItems.MenuItemDescription" _
  + " FROM MenuItems INNER JOIN OrderTransactions ON MenuItems.MenuItemID = OrderTransactions.MenuItemID" _
  + " WHERE OrderTransactions.OrderID = ?"

  Set cmd = New ADODB.Command
  
  With cmd
    .ActiveConnection = cn
    .Prepared = True
    .CommandText = query
    .Parameters.Append .CreateParameter("OrderTransactions.OrderID", adInteger, adParamInput, 18, 1)
  End With
  
  Set rsQuery = cmd.Execute
  
  Do While Not rsQuery.EOF
    Debug.Print "MenuItemID: "; rsQuery!MenuItemId; " Item: "; rsQuery!MenuItemText
    rsQuery.MoveNext
       
  Loop
  
  rsQuery.Close
  cn.Close
  
  Set rsQuery = Nothing
  Set cmd = Nothing
  Set db = Nothing
  
End Sub

Private Sub Command2_Click()

  Set OrderRepository = New OrderDao
  
  Set OrderDetails = OrderRepository.GetOrderDetails(1)
  
  Debug.Print OrderDetails.Count

  
End Sub
