VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdeloFelClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public AdeloFelApi As String

Public Function InvoiceCertification(ByVal request As CertificationRequest) As CertificationResponse

  Dim response As CertificationResponse

  Dim payload As String
  payload = BuildRequestPayload(request)

  Dim httpRequest As WinHttp.WinHttpRequest
  Set httpRequest = New WinHttp.WinHttpRequest
  
  httpRequest.Open "POST", AdeloFelApi
  httpRequest.SetRequestHeader "Content-Type", "application/json"
  httpRequest.Send (payload)
  
  Dim responseText As String
  responseText = httpRequest.responseText
  
  If httpRequest.Status <> 200 Then
  
    Debug.Print "error"
    Set response = BuildErrorResponse(responseText)
    
  Else
  
    Set response = BuildResponse(responseText)
    
  End If
  
  
  Set InvoiceCertification = response

End Function

Private Function BuildErrorResponse(ByVal text As String) As CertificationResponse

  Set obj = JSON.parse(text)

  Dim response As CertificationResponse
  Set response = New CertificationResponse
  
  response.Success = False
  Set response.Errors = obj.item("errors")
  
  Set obj = Nothing
  
  Set BuildErrorResponse = response

End Function

Private Function BuildResponse(ByVal text As String) As CertificationResponse

  Set obj = JSON.parse(text)

  Dim response As CertificationResponse
  Set response = New CertificationResponse
  
  response.Success = True
  response.UUID = obj.item("uuid")
  response.Correlative = obj.item("correlative")
  response.Number = obj.item("number")
  response.Information = obj.item("information")
  response.Description = obj.item("description")
  response.CertificationDate = obj.item("date")
  response.Origin = obj.item("origin")
  
  Set obj = Nothing
  
  Set BuildResponse = response
  
End Function

Private Function BuildRequestPayload(ByVal request As CertificationRequest) As String

  Dim payload As String
  payload = "{ ""orderId"": " + CStr(request.OrderId) + _
            ", ""taxId"": """ + request.TaxId + """, ""name"": """ + request.Name + """, ""email"": """ + request.Email + """, " + _
            "  ""details"" : " + BuildOrderDetails(request.OrderDetails) + " }"
  
  Debug.Print "request: " + payload
  
  BuildRequestPayload = payload

End Function

Private Function BuildOrderDetails(ByVal details As Object) As String
    
  Dim detail As OrderDetail
  Dim text As String
  
  text = "[ "
   
  For index = 0 To details.Count - 1
  
    Set detail = details(index)
    text = text + BuildOrderDetail(detail)
    
    If (index < (details.Count - 1)) Then
      text = text + ", "
    End If
  
  Next index
  
  text = text + " ]"
  
  BuildOrderDetails = text
End Function

Private Function BuildOrderDetail(ByVal detail As OrderDetail) As String
 
  Dim text As String
 
  text = "{ ""unitPrice"" : " + CStr(detail.UnitPrice) + ", ""quantity"" : " + CStr(detail.Quantity) + ", ""discountAmount"" : " + CStr(detail.DiscountAmount) + ", " + _
        "  ""itemText"" : """ + detail.ItemText + """ }"
  
  BuildOrderDetail = text
End Function
