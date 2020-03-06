Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
<
  XmlInclude(GetType(clsLookUpTypes)),
  Serializable()
  >
Public Class clsLookUpTypes

  Public Property Id As String = ""
  Public Property Code As String
  Public Property Description As String = ""
  Public Property Enabled As String = ""
  Public Property LookupGroup As String = ""
  Public Property AllowCustomLookupValues As Boolean = False
  Public Property OverridableType As Boolean = False

End Class
