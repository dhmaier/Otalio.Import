Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports System.IO
<
  XmlInclude(GetType(clsColumnProperties)),
  Serializable()
  >
Public Class clsSelectors

     Public Property ID As String = System.Guid.NewGuid.ToString

     Public Property Priority As String = ""

     Public Property APIEndpoint As String = ""

     Public Property Query As String = ""

     Public Property Label As String = ""

     Public Property ValidationType As String = "LIST"

     Public Property ReturnNodeValue As String = ""

     Public Property ReturnCell As String = ""

     Public Property Variable As String = ""

     Public Property Enabled As String = "1"

     Public Property Formatted As String = ""

     Public Property LookUpType As New clsLookUpTypes

     Public Property Sort As String = ""

     Public Property Visibility As String = "1"

     Public Property GraphQLSourceNode As String = ""

     Public Property NodeID As String = ""
     Public Property NodeText As String = ""



     Public Function Clone() As clsSelectors
          Dim oObject As clsSelectors = DirectCast(Me.MemberwiseClone(), clsSelectors)
          oObject.ID = System.Guid.NewGuid.ToString
          Return oObject
     End Function




End Class
