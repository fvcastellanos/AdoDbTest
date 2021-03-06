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
  Set Generator = config.item("generator")
  
  Set info = New GeneratorInformation
  
  info.TaxId = Generator.item("taxId")
  info.TaxType = Generator.item("taxType")
  info.Code = Generator.item("code")
  info.Name = Generator.item("name")
  info.Email = Generator.item("email")
  info.Country = Generator.item("country")
  info.State = Generator.item("state")
  info.Address = Generator.item("address")
  info.PostalCode = Generator.item("postalCode")
  info.CompanyName = Generator.item("companyName")
  Set info.phrases = BuildPhraseConfiguration(Generator.item("phrases"))

  Set config = Nothing

  Set BuildGeneratorInformation = info
End Function

Private Function BuildPhraseConfiguration(ByVal phrases As Object) As Object

  Dim list As Object
  Set list = CreateObject("System.Collections.ArrayList")

  Dim foo As String

  For index = 1 To phrases.count
    Set p = New Phrase
    Set item = phrases(index)
    
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
