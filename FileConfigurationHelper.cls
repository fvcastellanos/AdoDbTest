VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileConfigurationHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function BuildGeneralInformation() As GeneralInformation
  
  Dim info As GeneralInformation
  Set config = ReadConfiguration()
  
  Set info = New GeneralInformation
  info.CurrencyCode = config.item("currencyCode")
  info.DocumentDate = "" ' need to generate date
  info.DocumentType = config.item("documentType")
  info.IsExporting = ""
  info.AccessNumber = ""
  info.Person = ""
  
  Set config = Nothing
  
  Set BuildGeneralInformation = info

End Function

Public Function BuildGeneratorInformation() As GeneratorInformation
  Dim info As GeneratorInformation
  Set config = ReadConfiguration()
  Set generator = config.item("generator")
  
  Set info = New GeneratorInformation
  
  info.SubscriptionType = generator.item("subscriptionType")
  info.TaxId = generator.item("taxId")
  info.TaxType = generator.item("taxType")
  info.Code = generator.item("code")
  info.Name = generator.item("name")
  info.Email = generator.item("email")
  info.Country = generator.item("country")
  info.State = generator.item("state")
  info.Address = generator.item("address")
  info.PostalCode = generator.item("postalCode")
  info.CompanyName = generator.item("companyName")
  Set info.Phrases = BuildPhraseConfiguration(generator.item("phrases"))

  Set config = Nothing

  Set BuildGeneratorInformation = info
End Function

Public Function BuildApiInformation() As ApiInformation
  Dim info As ApiInformation
  Set config = ReadConfiguration()
  Set api = config.item("apiInformation")

  Set info = New ApiInformation
  info.User = api.item("user")
  info.WebServiceToken = api.item("wsToken")
  info.Alias = api.item("signatureAlias")
  info.Token = api.item("signatureToken")
  info.Salt = api.item("salt")
  
  Set config = Nothing
  Set BuildApiInformation = info

End Function

Private Function BuildPhraseConfiguration(ByVal Phrases As Object) As Object

  Dim list As Object
  Set list = CreateObject("System.Collections.ArrayList")

  Dim foo As String

  For index = 1 To Phrases.count
    Set p = New Phrase
    Set item = Phrases(index)
    
    p.PhraseType = item("type")
    p.Scenario = item("scenario")
    p.ResolutionNumber = item("resolutionNumber")
    p.ResolutionDate = item("resolutionDate")
    
    list.Add (p)
    
  Next index
  
  Set BuildPhraseConfiguration = list
End Function

Private Function ReadConfiguration() As Object
  
  Dim configText As String
  configText = ReadConfigurationFile()
  
  Set ReadConfiguration = JSON.parse(configText)

End Function

Private Function ReadConfigurationFile() As String

  Dim text As String
  Dim fso As FileSystemObject
  Dim ts As TextStream
    
  Set fso = New FileSystemObject
  
  'TODO: User environment variable to load configuration
  Set ts = fso.OpenTextFile("C:\Users\fvcg\Shared\App\config.json", ForReading)
  text = ts.ReadAll
  
  ts.Close
  Set ts = Nothing
  Set fso = Nothing
  
  ReadConfigurationFile = text

End Function
