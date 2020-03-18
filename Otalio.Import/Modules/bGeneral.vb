Module bGeneral


  Public Property goConnection As New clsConnectionDetails
  Public Property goHTTPServer As New clsAPI
  Public Property goOpenDataImportTemplate As New clsDataImportTemplate
  Public Property goOpenWorkBook As New clsWorkbook
  Public Property goLookupTypes As New List(Of clsLookUpTypes)
  Public Property goHierarchies As New List(Of clsHierarchies)
  Public Property goJsonHelper As New clsJsonHelper

  Public Const gsSettingFileNameOld As String = "OtalioISF.Setting"

  Public Const gsSettingFileName As String = "OtalioImport.Setting"
  Public Property gsSelectedHierarchy As String = ""

  Public Enum UserControlWaitFormCommand
    SetSize
  End Enum

End Module
