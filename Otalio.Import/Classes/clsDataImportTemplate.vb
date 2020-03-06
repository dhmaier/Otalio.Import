Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsDataImportTemplate)),
  Serializable()
  >
Public Class clsDataImportTemplate

  Public Property ID As String = System.Guid.NewGuid.ToString
  Public Property ImportType As String = "1"
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


  Public ReadOnly Property SortedName As String
    Get
      Return String.Format("{0:000} - {1}", Priority, StrConv(Name, vbProperCase))
    End Get
  End Property

End Class
