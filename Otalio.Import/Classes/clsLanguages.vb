Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports System.IO
<
  XmlInclude(GetType(clsLanguages)),
  Serializable()
  >
Public Class clsLanguages
     Public Property id As String
     Public Property code As String
     Public Property comments As String
     Public Property enabled As Boolean
     Public Property externalCode As String
     Public Property iso3 As String
     Public Property translationEnabled As Boolean
     Public Property translations As Translations
End Class

Public Class En
          Public Property description As String
     End Class

Public Class Translations
     Public Property en As En
End Class

