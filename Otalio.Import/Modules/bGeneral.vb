Module bGeneral


     Public Property goConnection As New clsConnectionDetails
     Public Property goConnectionHistory As New clsConnectionHistory
     Public Property goHTTPServer As New clsAPI
     Public Property goOpenDataImportTemplate As New clsDataImportTemplateV2
     Public Property goOpenWorkBook As New clsWorkbook
     Public Property goLookupTypes As New List(Of clsLookUpTypes)
     Public Property goHierarchies As New List(Of clsHierarchies)

     Public Property goLanguages As New List(Of clsLanguages)
     Public Property goJsonHelper As New clsJsonHelper

     Public Const gsSettingFileNameOld As String = "OtalioISF.Setting"

     Public Const gsSettingFileName As String = "OtalioImport.Setting"

     Public Const gsWorkbookExtention As String = ".ditw2"

     Public Const gsImportTemplateHeaderExtention As String = ".dit2"
     Public Property gsSelectedHierarchy As String = ""
     Public Property gsSelectedHierarchyParent As String = ""
     Public Property gsSelectedHierarchyType As String = ""

     Public Property gbIgnoreArrays As Boolean = False


     Public mbCancel As Boolean = False

     Public Enum UserControlWaitFormCommand
          SetSize
     End Enum

     Public Property gbTranslationsEnabled As Boolean = False
End Module
