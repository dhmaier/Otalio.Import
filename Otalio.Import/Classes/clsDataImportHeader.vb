Imports System.Xml.Serialization
Imports System.Linq

<
  XmlInclude(GetType(clsDataImportHeader)),
  Serializable()
  >
Public Class clsDataImportHeader


     Public Property ID As String = System.Guid.NewGuid.ToString
     Public Property Name As String = ""
     Public Property Group As String = ""
     Public Property WorkbookSheetName As String = ""
     Public Property Notes As String = ""
     Public Property Priority As Integer = 0
     Public Property Templates As New List(Of clsDataImportTemplateV2)
     Public Property StatusCodeColumn As String = ""
     Public Property StatusDescriptionColumn As String = ""

     Public ReadOnly Property SortedName As String
          Get
               Return String.Format("{0:000} - {1}", Priority, StrConv(Name, vbProperCase))
          End Get
     End Property

     Public ReadOnly Property IsHierarchal As Boolean
          Get
               If Templates IsNot Nothing AndAlso Templates.Count > 0 Then
                    For Each oTemplate As clsDataImportTemplateV2 In Templates
                         If oTemplate.IsHierarchal = True Then Return True
                    Next
               End If

               Return False

          End Get
     End Property

     Public ReadOnly Property IsShipEntity As Boolean
          Get
               If Templates IsNot Nothing AndAlso Templates.Count > 0 Then
                    For Each oTemplate As clsDataImportTemplateV2 In Templates
                         If oTemplate.IsShipEntity = True Then Return True
                    Next
               End If

               Return False
          End Get
     End Property

     Public ReadOnly Property IsRVCEntity As Boolean
          Get
               If Templates IsNot Nothing AndAlso Templates.Count > 0 Then
                    For Each oTemplate As clsDataImportTemplateV2 In Templates
                         If oTemplate.IsRVCEntity = True Then Return True
                    Next
               End If

               Return False
          End Get
     End Property

     Public ReadOnly Property IsLookupList As Boolean
          Get
               If Templates IsNot Nothing AndAlso Templates.Count > 0 Then
                    For Each oTemplate As clsDataImportTemplateV2 In Templates
                         If oTemplate.IsLookupList = True Then Return True
                    Next
               End If

               Return False
          End Get
     End Property


     Public ReadOnly Property ImportColumns As List(Of clsImportColum)
          Get
               Try
                    Dim oList As New List(Of clsImportColum)
                    If Templates IsNot Nothing Then
                         For Each oTemplate As clsDataImportTemplateV2 In Templates
                              For Each oColumn In oTemplate.ImportColumns
                                   If oList.Exists(Function(n) n.ColumnID = oColumn.ColumnID) = False Then
                                        oList.Add(oColumn)
                                   End If
                              Next
                         Next
                    End If

                    Return oList
               Catch ex As Exception
                    Return New List(Of clsImportColum)
               End Try
          End Get
     End Property

     Public ReadOnly Property Validators As List(Of clsValidation)
          Get
               Try
                    Dim oList As New List(Of clsValidation)
                    If Templates IsNot Nothing Then
                         For Each oTemplate As clsDataImportTemplateV2 In Templates.OrderBy(Function(n) n.Position).ToList
                              For Each oValidator In oTemplate.Validators.OrderBy(Function(n) n.Priority).ToList
                                   oList.Add(oValidator)
                              Next
                         Next
                    End If

                    Return oList
               Catch ex As Exception
                    Return New List(Of clsValidation)
               End Try
          End Get
     End Property

End Class
