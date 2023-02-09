Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
<
  XmlInclude(GetType(clsValidation)),
  Serializable()
  >
Public Class clsValidation

     Public Enum ValidationTypes
          ValidateObjectExists = 0
          FindAndReplace = 1
          Translation = 2
     End Enum

     Public Property ID As String = System.Guid.NewGuid.ToString

     Public Property Priority As String = ""

     Public Property APIEndpoint As String = ""

     Public Property Query As String = ""

     Public Property Headers As String = ""

     Public Property ValidationType As String = "1"

     Public Property ReturnNodeValue As String = ""

     Public Property ReturnCell As String = ""

     Public Property Comments As String = ""

     Public Property Enabled As String = "1"

     Public Property Formatted As String = ""

     Public Property LookUpType As New clsLookUpTypes

     Public Property Sort As String = ""

     Public Property Visibility As String = "1"

     Public Property GraphQLSourceNode As String = ""

     Public Property PreloadData As Boolean = False

     Public Function Clone() As clsValidation
          Dim oObject As clsValidation = DirectCast(Me.MemberwiseClone(), clsValidation)
          oObject.ID = System.Guid.NewGuid.ToString
          Return oObject
     End Function

     Public ReadOnly Property ReturnCellWithoutProperties As String
          Get
               Dim sColumnCode As String = ""
               If ReturnCell.Contains(":") Then
                    Dim sValues() As String = ReturnCell.Split(":")
                    If sValues IsNot Nothing And UBound(sValues) > 0 Then
                         sColumnCode = sValues(0)
                    End If
               Else
                    sColumnCode = ReturnCell
               End If

               Return sColumnCode

          End Get
     End Property

     Public ReadOnly Property DataObject As String
          Get
               Dim sAPIObject As String = ""

               Select Case ValidationType
                    Case "5", "6"

                         If LookUpType IsNot Nothing Then
                              sAPIObject = LookUpType.Description
                         End If

                    Case Else
                         If String.IsNullOrEmpty(APIEndpoint) = False Then
                              sAPIObject = APIEndpoint.Substring(APIEndpoint.LastIndexOf("/") + 1)
                         End If
               End Select

               If sAPIObject.Contains("?") Then
                    sAPIObject = sAPIObject.Substring(0, sAPIObject.IndexOf("?"))
               End If


               sAPIObject = StrConv(sAPIObject, VbStrConv.ProperCase)
               Return sAPIObject

          End Get
     End Property

     Public ReadOnly Property DetailedDescriptionFromQuery As String
          Get
               Dim sAPIObject As String = DataObject

               If String.IsNullOrEmpty(Query) = False Then
                    sAPIObject += String.Format(" ({0})", Query.Substring(0, Query.LastIndexOf("==")))
               End If

               sAPIObject = StrConv(sAPIObject, VbStrConv.ProperCase)
               Return sAPIObject
          End Get
     End Property


     Public ReadOnly Property DetailedDescriptionFromReturnValue As String
          Get
               Dim sAPIObject As String = DataObject

               If String.IsNullOrEmpty(ReturnNodeValue) = False Then
                    sAPIObject += String.Format(" ({0})", ReturnNodeValue.Substring(ReturnNodeValue.LastIndexOf(".") + 1))
               End If

               sAPIObject = StrConv(sAPIObject, VbStrConv.ProperCase)
               Return sAPIObject
          End Get

     End Property

End Class
