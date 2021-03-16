VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public HttpResponse As String
Private myMSXML As Object

Public Function InvoiceRequest(ByVal orderId As Long, ByVal taxId As String, ByVal name As String, ByVal email As String) As String

  Dim payload As String
  payload = BuildRequestPayload(orderId, taxId, name, email)

  Dim request As WinHttp.WinHttpRequest
  Set request = New WinHttp.WinHttpRequest
  
  request.Open "POST", "http://localhost:8080/invoices"
  request.SetRequestHeader "Content-Type", "application/json"
  request.Send (payload)
  
  InvoiceRequest = request.ResponseText

End Function

Private Function BuildRequestPayload(ByVal orderId As Long, ByVal taxId As String, ByVal name As String, ByVal email As String) As String

  Dim payload As String
  payload = "{ ""orderId"": " + CStr(orderId) + ", ""taxId"": """ + taxId + """, ""name"": """ + name + """, ""email"": """ + email + """ }"
  
  Debug.Print "request: " + payload

  BuildRequestPayload = payload

End Function

Private Function OnReadyState()

  HttpResponse = myMSXML.ResponseText
  Debug.Print "Response: " + HttpResponse
  
End Function
