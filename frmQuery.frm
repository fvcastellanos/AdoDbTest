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
   Begin VB.TextBox edFecha 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   4800
      Width           =   3375
   End
   Begin VB.TextBox edNumero 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox edCorrelativo 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox edUuid 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ListBox lbErrors 
      Height          =   1620
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   7455
   End
   Begin VB.CommandButton edGenerarFactura 
      Caption         =   "Generar Factura"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox edCorreo 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox edNombre 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox edNit 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "CF"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Numero:"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Correlativo:"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "UUID:"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbRespuesta 
      Caption         =   "Processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lbRespuesta2 
      Caption         =   "Respueta:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Correo"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "NIT:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
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
  
  Set OrderDetails = OrderRepository.GetOrderDetails(1)
  
  Debug.Print "Total amount: "; OrderRepository.GetOrderTotalAmount(OrderDetails)

  
End Sub

Private Sub edGenerarFactura_Click()

  lbRespuesta.Caption = "Processing"

  Set OrderRepository = New OrderDao
  Set details = OrderRepository.GetOrderDetails(1)

  lbRespuesta.Caption = ""
  Dim client As AdeloFelClient
  Set client = New AdeloFelClient
  client.AdeloFelApi = "http://localhost:8080/invoices"
  
  Dim request As CertificationRequest
  Set request = New CertificationRequest
    
  request.OrderId = 1
  request.Name = edNombre.text
  request.TaxId = edNit.text
  request.Email = edCorreo.text
  Set request.OrderDetails = details

  Dim response As CertificationResponse
  Set response = client.InvoiceCertification(request)
  
  If response.Success = True Then
  
    lbRespuesta.Caption = "Generada!"
    edUuid.text = response.UUID
    edCorrelativo.text = response.Correlative
    edNumero.text = response.Number
    edFecha.text = response.CertificationDate
  
  Else
    
    lbRespuesta.Caption = "Error!"
    For index = 1 To response.Errors.Count
    
      lbErrors.AddItem (response.Errors(index))
    
    Next index
  End If
  
  Set client = Nothing
  Set OrderRepository = Nothing
  Set request = Nothing
  Set response = Nothing
  
End Sub

