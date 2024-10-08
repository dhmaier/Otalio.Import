﻿Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Threading
Imports System.IO.Compression
<
  XmlInclude(GetType(clsDataImportTemplateV2)),
  Serializable()
  >
Public Class clsDataImportTemplateV2

     Public Property ID As String = System.Guid.NewGuid.ToString
     Public Property Name As String = ""
     Public Property APIEndpoint As String = ""
     Public Property APIEndpointSelect As String = ""
     Public Property DTOObject As String = ""
     Public Property Validators As New List(Of clsValidation)
     Public Property ImportColumns As New List(Of clsImportColum)
     Public Property Selectors As New List(Of clsSelectors)
     Public Property ReturnNodeValue As String = ""
     Public Property ReturnCell As String = ""
     Public Property ReturnCellDTO As String = ""
     Public Property UpdateQuery As String = ""
     Public Property SelectQuery As String = ""
     Public Property Variables As New List(Of clsVariables)
     Public Property GraphQLQuery As String = ""
     Public Property GraphQLRootNode As String = ""
     Public Property EntityColumn As String = ""
     Public Property FileLocationColumn As String = ""
     Public Property StatusCodeColumn As String = ""
     Public Property StatusDescriptionColumn As String = ""
     Public Property Group As String = ""
     Public Property Position As Integer = 0
     Public Property IsMaster As Boolean = False
     Public Property IsEnabled As Boolean = True
     Public Property ImportType As String = "2"
     Public Property IgnoreArray As Boolean = False
     Public Property RemoveEmptyAndNull As Boolean = False
     Public Property WorkbookSheetName As String = ""
     Public Property ImageResizeFormat As String = "0,0"


     Public ReadOnly Property DTOObjectFormated As String
          Get
               Return FindAndReplaceTranslations(Me.DTOObject)
          End Get
     End Property


     Public ReadOnly Property IsHierarchal As Boolean
          Get

               If String.IsNullOrEmpty(APIEndpoint) = False AndAlso APIEndpoint.Contains("@@HIERARCHYID@@") Then Return True
               If String.IsNullOrEmpty(DTOObject) = False AndAlso DTOObject.Contains("@@HIERARCHYID@@") Then Return True
               If String.IsNullOrEmpty(SelectQuery) = False AndAlso SelectQuery.Contains("@@HIERARCHYID@@") Then Return True
               If String.IsNullOrEmpty(UpdateQuery) = False AndAlso UpdateQuery.Contains("@@HIERARCHYID@@") Then Return True
               If String.IsNullOrEmpty(GraphQLQuery) = False AndAlso GraphQLQuery.Contains("@@HIERARCHYID@@") Then Return True

               If Validators IsNot Nothing Then
                    For Each oValidator As clsValidation In Validators
                         If oValidator IsNot Nothing Then
                              If String.IsNullOrEmpty(oValidator.APIEndpoint) = False AndAlso oValidator.APIEndpoint.Contains("@@HIERARCHYID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Query) = False AndAlso oValidator.Query.Contains("@@HIERARCHYID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Sort) = False AndAlso oValidator.Sort.Contains("@@HIERARCHYID@@") Then Return True
                         End If
                    Next
               End If

               Return False

          End Get
     End Property

     Public ReadOnly Property IsShipEntity As Boolean
          Get

               If String.IsNullOrEmpty(APIEndpoint) = False AndAlso APIEndpoint.Contains("@@SHIPID@@") Then Return True
               If String.IsNullOrEmpty(DTOObject) = False AndAlso DTOObject.Contains("@@SHIPID@@") Then Return True
               If String.IsNullOrEmpty(SelectQuery) = False AndAlso SelectQuery.Contains("@@SHIPID@@") Then Return True
               If String.IsNullOrEmpty(UpdateQuery) = False AndAlso UpdateQuery.Contains("@@SHIPID@@") Then Return True
               If String.IsNullOrEmpty(GraphQLQuery) = False AndAlso GraphQLQuery.Contains("@@SHIPID@@") Then Return True

               If Validators IsNot Nothing Then
                    For Each oValidator As clsValidation In Validators
                         If oValidator IsNot Nothing Then
                              If String.IsNullOrEmpty(oValidator.APIEndpoint) = False AndAlso oValidator.APIEndpoint.Contains("@@SHIPID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Query) = False AndAlso oValidator.Query.Contains("@@SHIPID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Sort) = False AndAlso oValidator.Sort.Contains("@@SHIPID@@") Then Return True
                         End If
                    Next
               End If

               Return False

          End Get
     End Property

     Public ReadOnly Property IsRVCEntity As Boolean
          Get

               If String.IsNullOrEmpty(APIEndpoint) = False AndAlso APIEndpoint.Contains("@@RVCID@@") Then Return True
               If String.IsNullOrEmpty(DTOObject) = False AndAlso DTOObject.Contains("@@RVCID@@") Then Return True
               If String.IsNullOrEmpty(SelectQuery) = False AndAlso SelectQuery.Contains("@@RVCID@@") Then Return True
               If String.IsNullOrEmpty(UpdateQuery) = False AndAlso UpdateQuery.Contains("@@RVCID@@") Then Return True
               If String.IsNullOrEmpty(GraphQLQuery) = False AndAlso GraphQLQuery.Contains("@@RVCID@@") Then Return True

               If Validators IsNot Nothing Then
                    For Each oValidator As clsValidation In Validators
                         If oValidator IsNot Nothing Then
                              If String.IsNullOrEmpty(oValidator.APIEndpoint) = False AndAlso oValidator.APIEndpoint.Contains("@@RVCID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Query) = False AndAlso oValidator.Query.Contains("@@RVCID@@") Then Return True
                              If String.IsNullOrEmpty(oValidator.Sort) = False AndAlso oValidator.Sort.Contains("@@RVCID@@") Then Return True
                         End If
                    Next
               End If

               Return False

          End Get
     End Property

     Public ReadOnly Property IsLookupList As Boolean
          Get
               If String.IsNullOrEmpty(APIEndpoint) = False AndAlso APIEndpoint.ToLower.Trim.Contains("metadata/v1/lookup-values") Then Return True
          End Get
     End Property


     Public ReadOnly Property DataObject As String
          Get
               Dim sAPIObject As String = ""


               If String.IsNullOrEmpty(APIEndpoint) = False Then
                    sAPIObject = APIEndpoint.Substring(APIEndpoint.LastIndexOf("/") + 1)
               End If

               If sAPIObject.Contains("?") Then
                    sAPIObject = sAPIObject.Substring(0, sAPIObject.IndexOf("?"))
               End If


               sAPIObject = StrConv(sAPIObject, VbStrConv.ProperCase)
               Return sAPIObject

          End Get
     End Property

     ' Read-only property to extract the width from the ImageResizeFormat
     Public ReadOnly Property Width As Integer
          Get
               If String.IsNullOrEmpty(ImageResizeFormat) Then
                    ImageResizeFormat = "0,0"
               End If
               Dim parts() As String = ImageResizeFormat.Split(","c)
               If parts.Length = 2 Then
                    Dim widthValue As Integer
                    If Integer.TryParse(parts(0).Trim(), widthValue) Then
                         Return widthValue
                    End If
               End If
               ' Return a default value or handle the error as needed
               Return 48
          End Get
     End Property

     ' Read-only property to extract the height from the ImageResizeFormat
     Public ReadOnly Property Height As Integer
          Get
               If String.IsNullOrEmpty(ImageResizeFormat) Then
                    ImageResizeFormat = "0,0"
               End If

               Dim parts() As String = ImageResizeFormat.Split(","c)
               If parts.Length = 2 Then
                    Dim heightValue As Integer
                    If Integer.TryParse(parts(1).Trim(), heightValue) Then
                         Return heightValue
                    End If
               End If
               ' Return a default value or handle the error as needed
               Return 48
          End Get
     End Property




End Class
