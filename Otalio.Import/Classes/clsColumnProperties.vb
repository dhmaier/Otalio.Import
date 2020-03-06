Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports System.IO
<
  XmlInclude(GetType(clsColumnProperties)),
  Serializable()
  >
Public Class clsColumnProperties

  Public Property CellName As String = ""
  Public Property Format As String = ""
  Public Property IndexID As String = ""

End Class
