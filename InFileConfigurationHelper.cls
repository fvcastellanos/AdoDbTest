VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InFileConfigurationHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function BuildGeneralInformation() As GeneralInformation
  
  Dim info As GeneralInformation
  Set config = ReadConfiguration()
  
  Set info = New GeneralInformation
  info.CurrencyCode = config.Item("currencyCode")
  info.DocumentDate = "" ' need to generate date
  info.DocumentType = config.Item("documentType")
  info.IsExporting = ""
  info.AccessNumber = ""
  info.Person = ""
  
  Set BuildGeneralInformation = info

End Function

Public Function BuildGeneratorInformation() As GeneratorInformation
  Dim info As GeneratorInformation
  Set config = ReadConfiguration()
  Set generator = config.Item("generator")
  
  Set info = New GeneratorInformation
  
  info.TaxId = generator.Item("taxId")
  info.TaxType = generator.Item("taxType")
  info.Code = generator.Item("code")
  info.Name = generator.Item("name")
  info.Email = generator.Item("email")
  info.Country = generator.Item("country")
  info.State = generator.Item("state")
  info.Address = generator.Item("address")
  info.PostalCode = generator.Item("postalCode")
  info.CompanyName = generator.Item("companyName")

  Set BuildGeneratorInformation = info
End Function

Private Function ReadConfiguration() As Object
  
  Dim configText As String
  configText = ReadConfigurationFile()
  
  Set ReadConfiguration = JSON.parse(configText)

End Function

Private Function ReadConfigurationFile() As String

  Dim text As String
  Dim fso As fileSystemObject
  Dim ts As textStream
    
  Set fso = New fileSystemObject
  
  'TODO: User environment variable to load configuration
  Set ts = fso.OpenTextFile("C:\Users\fvcg\Shared\App\config.json", ForReading)
  text = ts.ReadAll
  
  ts.Close
  Set ts = Nothing
  Set fso = Nothing
  
  ReadConfigurationFile = text

End Function
