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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
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

Private Sub Command2_Click()

  Set OrderRepository = New OrderDao
  
  Set orderDetails = OrderRepository.GetOrderDetails(1)
  
  Debug.Print "Total amount: "; OrderRepository.GetOrderTotalAmount(orderDetails)

  
End Sub

Private Sub Command3_Click()

  Set configurationHelper = New InFileConfigurationHelper

  Dim info As GeneralInformation
  Set info = configurationHelper.BuildGeneralInformation()

  Debug.Print "Currency Code: "; info.CurrencyCode
  
  Set generator = configurationHelper.BuildGeneratorInformation()
  
  Debug.Print "Name: "; generator.Name
  
  Set configurationHelper = Nothing
  
End Sub
