Imports System.IO
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.Linq
Imports System.Net
Imports System.Threading
Imports System.IO.Compression


Public Class ucDataImportProperties

     Private moImportHeader As New clsDataImportHeader
     Private moSelectedTemplate As New clsDataImportTemplateV2

     Public msFileName As String = ""
     Public msFilePath As String = ""

     Public Property SelectedTemplate As clsDataImportTemplateV2
          Get
               Return moSelectedTemplate
          End Get
          Set(value As clsDataImportTemplateV2)
               moSelectedTemplate = value
          End Set
     End Property
     Public Property ImportHeader As clsDataImportHeader
          Get
               Return moImportHeader
          End Get
          Set(value As clsDataImportHeader)
               If value IsNot Nothing Then

                    moSelectedTemplate = Nothing
                    moImportHeader = value

                    With txtName
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "Name"))
                    End With

                    With txtWorkbookSheet
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "WorkbookSheetName"))
                    End With

                    With txtMemo
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "Notes"))
                    End With


                    With sePriority
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("EditValue", moImportHeader, "Priority"))
                    End With

                    With txtGroup
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "Group"))
                    End With

                    With txtGlobalStatusCode
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "StatusCodeColumn"))
                    End With

                    With txtGlobalStatusDescription
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "StatusDescriptionColumn"))
                    End With


                    Call LoadTemplates(value)

                    gridColumnListAll.DataSource = moImportHeader.ImportColumns
                    gdColumnsListAll.BestFitColumns()

                    gridValidatorsAll.DataSource = moImportHeader.Validators
                    gdValidatorsAll.BestFitColumns()

               End If
          End Set
     End Property

     Sub New()

          InitializeComponent()

     End Sub

     Private Sub gdValidators_DoubleClick(sender As Object, e As EventArgs) Handles gdValidators.DoubleClick
          With gdValidators
               If .SelectedRowsCount >= 0 Then
                    Dim oValidation As clsValidation = TryCast(.GetRow(.FocusedRowHandle), clsValidation)
                    If oValidation IsNot Nothing Then

                         Using oForm As New frmValidatorEdit
                              oForm._ucValidationProperties._Validation = oValidation
                              Select Case oForm.ShowDialog
                                   Case DialogResult.OK 'ok edited
                                   Case DialogResult.Cancel 'dont do nothing
                                   Case DialogResult.Abort 'use to flag item to be deleted
                                        Me.moSelectedTemplate.Validators.Remove(oForm._ucValidationProperties._Validation)

                              End Select

                              Me.gridValidators.RefreshDataSource()
                              gdValidators.BestFitColumns()

                         End Using
                    End If
               End If
          End With
     End Sub

     Private Sub icbType_EditValueChanged(sender As Object, e As EventArgs) Handles icbType.EditValueChanged

          Select Case moSelectedTemplate.ImportType

               Case "3" 'file import

                    lciEntityColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lciFileLocationColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lciReturnNodeColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
                    lciReturnNodeName.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always

               Case Else

                    lciEntityColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
                    lciFileLocationColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
                    lciReturnNodeColumn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lciReturnNodeName.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always

          End Select

     End Sub

     Public Sub SaveSettings(Optional pbSaveAs As Boolean = False)

          Dim fd As SaveFileDialog = New SaveFileDialog
          Me.gridImportColumns.Focus()

          If pbSaveAs Or String.IsNullOrEmpty(msFilePath) Then

               fd.Title = "Save a Otalio Dynamic import template."
               fd.Filter = String.Format("Dynamic import Template |*{0}", gsImportTemplateHeaderExtention)
               fd.FilterIndex = 1
               fd.FileName = moImportHeader.Name.ToString
               fd.RestoreDirectory = True
               If fd.ShowDialog() = DialogResult.OK Then

                    msFilePath = fd.FileName

                    If File.Exists(fd.FileName) = False Then
                         SaveFile(fd.FileName, moImportHeader)
                    Else
                         SaveFile(fd.FileName, moImportHeader)
                    End If

               End If
          Else

               SaveFile(msFilePath, moImportHeader)

          End If


          fd.Dispose()
          fd = Nothing

     End Sub

     Public Sub SaveSettings(psPath As String, poImportTemplate As clsDataImportHeader)

          SaveFile(psPath, poImportTemplate)

     End Sub


     Private Sub LoadTemplates(poImportHeader As clsDataImportHeader)

          'clear existing items



          If poImportHeader IsNot Nothing Then
               With poImportHeader
                    gridTemplates.DataSource = .Templates.OrderBy(Function(n) .Priority).ToList
               End With
          End If


          If moSelectedTemplate IsNot Nothing Then
               gdTemplates.SelectRow(moSelectedTemplate.Position - 1)
               BindTemplate(moSelectedTemplate)
          Else
               If poImportHeader.Templates IsNot Nothing AndAlso poImportHeader.Templates.Count > 0 Then
                    gdTemplates.SelectRow(0)
                    moSelectedTemplate = poImportHeader.Templates(0)
                    BindTemplate(moSelectedTemplate)
               End If
          End If

     End Sub

     Private Sub BindTemplate(poTemplate As clsDataImportTemplateV2)


          With poTemplate
               If .UpdateQuery = Nothing Then .UpdateQuery = ""
          End With

          With poTemplate
               If .SelectQuery = Nothing Then .SelectQuery = ""
          End With

          With poTemplate
               If .GraphQLQuery = Nothing Then .GraphQLQuery = ""
          End With

          With poTemplate
               If .GraphQLRootNode = Nothing Then .GraphQLRootNode = ""
          End With

          With icbType
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Value", poTemplate, "ImportType"))
          End With

          With txtTemplateName
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "Name"))
          End With

          With txtAPIEndpoint
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "APIEndpoint"))
          End With

          With txtAPIEndpointSelect
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "APIEndpointSelect"))
          End With

          With txtDataTransportObject
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "DTOObject"))
          End With



          With txtStatusCodeColumn
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "StatusCodeColumn"))
          End With

          With txtStatusDescriptionColumn
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "StatusDescriptionColumn"))
          End With

          With txtReturnValueName
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "ReturnNodeValue"))
          End With

          With txtReturnValueColumn
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "ReturnCell"))
          End With

          With txtReturnCellDTO
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "ReturnCellDTO"))
          End With

          With txtAPIQuery
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "UpdateQuery"))
          End With

          With txtSelectQuery
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "SelectQuery"))
          End With


          With txtGraphQLQuery
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "GraphQLQuery"))
          End With


          With txtGraphQLNode
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "GraphQLRootNode"))
          End With


          With txtEntityColumn
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "EntityColumn"))
          End With

          With txtFileLocationColumn
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "FileLocationColumn"))
          End With

          With chkMaster
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Checked", poTemplate, "IsMaster"))
          End With

          gridValidators.DataSource = poTemplate.Validators
          gdImportColumns.BestFitColumns()
          gdValidators.BestFitColumns()



     End Sub

     Public Sub AddTemplate()
          If moImportHeader.Templates Is Nothing Then
               moImportHeader.Templates = New List(Of clsDataImportTemplateV2)
          End If
          moImportHeader.Templates.Add(New clsDataImportTemplateV2 With {.Name = "New", .Position = moImportHeader.Templates.Count + 1})
          moSelectedTemplate = moImportHeader.Templates(moImportHeader.Templates.Count - 1)
          LoadTemplates(moImportHeader)
          BindTemplate(moSelectedTemplate)
     End Sub

     Public Sub RemoveTemplate()
          If moImportHeader IsNot Nothing AndAlso moImportHeader.Templates.Count > 0 Then
               If moSelectedTemplate IsNot Nothing Then
                    moImportHeader.Templates.Remove(moSelectedTemplate)
                    If moImportHeader.Templates.Count > 0 Then
                         moSelectedTemplate = moImportHeader.Templates(0)
                         BindTemplate(moSelectedTemplate)
                    End If

               End If
          End If
     End Sub

     Private Sub gdTemplates_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles gdTemplates.FocusedRowChanged
          With gdTemplates
               If .SelectedRowsCount > 0 Then
                    moSelectedTemplate = TryCast(.GetRow(e.FocusedRowHandle), clsDataImportTemplateV2)
                    BindTemplate(moSelectedTemplate)
               End If
          End With
     End Sub

     Public Sub Reload()
          Call LoadTemplates(moImportHeader)
     End Sub
End Class
