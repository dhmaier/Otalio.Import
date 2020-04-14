﻿Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsConnectionHistory)),
  Serializable()
  >
Public Class clsConnectionHistory

  Public Property ConnectionHistory As New List(Of clsConnectionDetails)
  Public Property ActiveConnection As clsConnectionDetails

  Public Property LoggedEvents As Boolean = True

     Public Property LastUsedExportFolder As String = ""

     Public Property LastUsedWorkbookFolder As String = ""

End Class
