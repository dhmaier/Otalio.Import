Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
<
  XmlInclude(GetType(clsVariables)),
  Serializable()
  >
Public Class clsVariables
  Public Property Name As String = ""
  Public Property Value As String = ""
End Class
