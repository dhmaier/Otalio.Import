Imports System
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Linq
Imports System.Collections.Generic

Public Class clsJsonHelper

  Public Function ParseJson(ByVal token As JToken, ByVal nodes As Dictionary(Of String, String), ByVal Optional parentLocation As String = "") As Boolean
    If token.HasValues Then

      For Each child As JToken In token.Children()

        If token.Type = JTokenType.[Property] Then

          If parentLocation = "" Then
            parentLocation = (CType(token, JProperty)).Name
          Else
            parentLocation += "." & (CType(token, JProperty)).Name
          End If
        End If

        ParseJson(child, nodes, parentLocation)
      Next

      Return True
    Else

      If nodes.ContainsKey(parentLocation) Then
        nodes(parentLocation) += "|" & token.ToString()
      Else
        nodes.Add(parentLocation, token.ToString())
      End If

      Return False
    End If
  End Function


End Class
