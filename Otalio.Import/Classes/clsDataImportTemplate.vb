Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsDataImportTemplate)),
  Serializable()
  >
Public Class clsDataImportTemplate

     Public Property ID As String = System.Guid.NewGuid.ToString
     Public Property ImportType As String = "2"
     Public Property Name As String = ""
     Public Property APIEndpoint As String = ""
     Public Property DTOObject As String = ""
     Public Property Validators As New List(Of clsValidation)
     Public Property ImportColumns As New List(Of clsImportColum)
     Public Property WorkbookSheetName As String = ""
     Public Property StatusCodeColumn As String = ""
     Public Property StatusDescirptionColumn As String = ""
     Public Property Notes As String = ""
     Public Property ReturnNodeValue As String = ""
     Public Property ReturnCell As String = ""
     Public Property ReturnCellDTO As String = ""
     Public Property UpdateQuery As String = ""
     Public Property Priority As Integer = 0
     Public Property SelectQuery As String = ""
     Public Property Variables As New List(Of clsVariables)
     Public Property GraphQLQuery As String = ""
     Public Property GraphQLRootNode As String = ""
     Public Property EntityColumn As String = ""
     Public Property FileLocationColumn As String = ""
     Public Property Group As String = ""



     Public ReadOnly Property SortedName As String
          Get
               Return String.Format("{0:000} - {1}", Priority, StrConv(Name, vbProperCase))
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

End Class
