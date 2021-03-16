VERSION 5.00
Begin VB.Form frmQuery 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton edGenerarFactura 
      Caption         =   "Generar Factura"
      Height          =   615
      Left            =   5280
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox edCorreo 
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox edNombre 
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox edNit 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Text            =   "CF"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   600
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
   Begin VB.Label lbTexto 
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   1800
      TabIndex        =   12
      Top             =   4800
      Width           =   5895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbRespuesta 
      Caption         =   "Respueta:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Correo"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "NIT:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   735
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

  Set configurationHelper = New FileConfigurationHelper

  Dim info As GeneralInformation
  Set info = configurationHelper.BuildGeneralInformation()

  Debug.Print "Currency Code: "; info.CurrencyCode
  
  Set Generator = configurationHelper.BuildGeneratorInformation()
  
  Debug.Print "Name: "; Generator.name
  
  Set configurationHelper = Nothing
  
End Sub

Private Sub Command4_Click()
  Dim service As ElectronicInvoiceService
  Dim Response As String
  
  Set service = New ElectronicInvoiceService
  Response = service.GenerateInvoice(1, "1231232", "adelo@mailnator.com")
  
  MsgBox (Response)
  
End Sub

Private Sub edGenerarFactura_Click()

  Dim client As HttpClient
  Set client = New HttpClient
  
  lbTexto.Caption = client.InvoiceRequest(1, edNit.text, edNombre.text, edCorreo.text)
  
  Set client = Nothing
End Sub
