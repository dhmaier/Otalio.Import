Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
<
  XmlInclude(GetType(clsWorkbook)),
  Serializable()
  >
Public Class clsWorkbook

     Private moTempaltes As New List(Of clsDataImportTemplate)

     Public Property Templates As List(Of clsDataImportTemplate)
          Get
               Return moTempaltes
          End Get
          Set(value As List(Of clsDataImportTemplate))
               moTempaltes = value
          End Set
     End Property

     Public Property WorkbookName As String = ""
     Public Property MajorVersion As Integer = 1
     Public Property MinorVersion As Integer = 0
     Public Property SaveVersion As Integer = 0
     Public Property SelectedHierarchy As String = ""

     Public ReadOnly Property WorkbookVersion As String
          Get
               Return String.Format("{0}.{1}.{2}", MajorVersion, MinorVersion, SaveVersion)
          End Get
     End Property

End Class
