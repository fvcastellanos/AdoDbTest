VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElectronicInvoiceService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GenerateInvoice(ByVal OrderId As Long, ByVal RecipientTaxId As String, ByVal RecipientEmail As String) As String

  'Load FEL client
  Set connector = CreateObject("conectorNET.RequestCertificacionFel")
  
  'Load configurations
  Dim configurationHelper As FileConfigurationHelper
  Set configurationHelper = New FileConfigurationHelper
  
  Dim general As GeneralInformation
  Set general = configurationHelper.BuildGeneralInformation()
  
  Dim Generator As GeneratorInformation
  Set Generator = configurationHelper.BuildGeneratorInformation()
  
  Dim api As ApiInformation
  Set api = configurationHelper.BuildApiInformation()
  
  'Load Invoice details
  Dim items As Object
  Set items = BuildInvoiceItems(OrderId)
  
  'Generate Invoice Information
  Set GeneralInformation = connector.Datos_Generales(general.CurrencyCode, general.DocumentDate, general.DocumentType, general.AccessNumber, general.IsExporting)
  Set GeneratorInformaton = connector.Datos_emisor(Generator.SubscriptionType, Generator.Code, Generator.PostalCode, Generator.Email, Generator.Country, "", Generator.State, Generator.Address, Generator.TaxId, Generator.Name, Generator.CompanyName)
  Set RecipientInformation = connector.Datos_receptor(RecipientTaxId, "", "", RecipientEmail, "", "", "", "", "")
  
  'Phrases
  
  Set ps = Generator.Phrases
  If (ps <> Null) Then
  
    Dim index As Integer
      
    For index = 0 To p.Count
    
      Set p = ps(index)
      Set result = connector.Frases(p.PhraseType, p.Scenario, p.ResolutionNumber, p.ResolutionDate)
    
    Next index
      
  End If
  
  'Invoice Details
  
  Dim i As Integer
  
  For i = 0 To items.Count
  
    Set item = items(i)
    Set taxItem = connector.Item_un_impuesto(item.ItemType, item.UnitType, item.Quantity, item.Description, item.lineNumber, item.UnitPrice, item.Price, item.Discount, item.Total, item.TaxName, item.AffectedUnit, item.AffectedQuantity, item.BaseAmount, item.Tax)
  
  Next i
  
  'Invoice Totals
  
  Dim dao As OrderDao
  Set dao = New OrderDao
  
  Set details = dao.GetOrderDetails(OrderId)
  
  Dim totalInvoice As Double
  totalInvoice = dao.GetOrderTotalAmount(details)
  
  t = connector.Totales(totalInvoice)
  ti = connector.Total_impuestos(totalInvoice / 1.12)
  
  'Send Request
  Dim identifier As String
  identifier = GUID.GetGUID()
  
  response = connector.enviar_peticion(api.User, api.WebServiceToken, identifier, Generator.Email, api.Alias, api.Token)
  
  Set dao = Nothing
  Set connector = Nothing
  Set configurationHelper = Nothing
  Set general = Nothing
  Set Generator = Nothing
  Set items = Nothing

  GenerateInvoice = response

End Function

Private Function BuildInvoiceInformation(ByVal OrderId As Long) As InvoiceInformation

  Dim info As InvoiceInformation
  Set info = New InvoiceInformation
  
  Dim configurationHelper As FileConfigurationHelper
  Set configurationHelper = New FileConfigurationHelper
  
  info.general = configurationHelper.BuildGeneralInformation()
  info.Generator = configurationHelper.BuildGeneratorInformation()
  info.InvoiceItems = BuildInvoiceItems(OrderId)
  
  Set BuildInvoiceInformation = info

End Function

Private Function BuildInvoiceItems(ByVal OrderId As Long) As Object

  Dim dao As OrderDao
  Dim details As Object
  
  Dim list As Object
  Set list = CreateObject("System.Collections.ArrayList")
  
  Set dao = New OrderDao
  Set details = dao.GetOrderDetails(OrderId)
  
  For index = 0 To details.Count - 1
  
    Dim item As InvoiceItem
    Set item = ToInvoiceItem(details(index), index + 1)
    list.Add (item)
  
  Next index
  
  Set details = Nothing
  Set dao = Nothing
  
  BuildInvoiceItems = list

End Function

Private Function ToInvoiceItem(ByVal detail As OrderDetail, ByVal lineNumber As Integer) As InvoiceItem

  Dim item As InvoiceItem
  item = New InvoiceItem
  
  Dim Total As Double
  Total = detail.Quantity * detail.UnitPrice
  
  item.ItemType = "B"
  item.UnitType = "UND"
  item.lineNumber = lineNumber
  item.Description = detail.ItemText
  item.Quantity = detail.Quantity
  item.UnitPrice = detail.UnitPrice
  item.Price = Total
  item.Total = Total
  item.Tax = Total / 1.12
  item.BaseAmount = item.Total - item.Tax
  item.TaxName = "IVA"
  item.AffectedUnit = 1
  
  Set ToInvoiceItem = item

End Function
