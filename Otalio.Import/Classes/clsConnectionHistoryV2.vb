Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsConnectionHistoryV2)),
  Serializable()
  >
Public Class clsConnectionHistoryV2

     Public Property ConnectionHistory As New List(Of clsConnectionDetailsV2)
     Public Property ActiveConnection As clsConnectionDetailsV2

     Public Property LoggedEvents As Boolean = True

     Public Property LastUsedExportFolder As String = ""

     Public Property LastUsedWorkbookFolder As String = ""

End Class
