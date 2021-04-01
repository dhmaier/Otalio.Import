Imports System.IO

Public Class ucDataImportProperties

     Private moDataImportTemplate As New clsDataImportTemplate
     Public msFileName As String = ""
     Public msFilePath As String = ""

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
                                        Me._DataImportTemplate.Validators.Remove(oForm._ucValidationProperties._Validation)

                              End Select

                              Me.gridValidators.RefreshDataSource()
                              gdValidators.BestFitColumns()

                         End Using
                    End If
               End If
          End With
     End Sub

     Private Sub icbType_EditValueChanged(sender As Object, e As EventArgs) Handles icbType.EditValueChanged

          Select Case moDataImportTemplate.ImportType

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
               fd.Filter = "Dynamic import Template |*.dit"
               fd.FilterIndex = 1
               fd.FileName = moDataImportTemplate.Name.ToString
               fd.RestoreDirectory = True
               If fd.ShowDialog() = DialogResult.OK Then

                    msFilePath = fd.FileName

                    If File.Exists(fd.FileName) = False Then
                         SaveFile(fd.FileName, moDataImportTemplate)
                    Else
                         SaveFile(fd.FileName, moDataImportTemplate)
                    End If

               End If
          Else

               SaveFile(msFilePath, moDataImportTemplate)

          End If


          fd.Dispose()
          fd = Nothing

     End Sub

     Public Sub SaveSettings(psPath As String, poImportTemplate As clsDataImportTemplate)

          SaveFile(psPath, poImportTemplate)

     End Sub

     Public Property _DataImportTemplate As clsDataImportTemplate
          Get
               Return moDataImportTemplate
          End Get
          Set(value As clsDataImportTemplate)
               If value IsNot Nothing Then

                    moDataImportTemplate = value

                    With moDataImportTemplate
                         If .UpdateQuery = Nothing Then .UpdateQuery = ""
                    End With

                    With moDataImportTemplate
                         If .SelectQuery = Nothing Then .SelectQuery = ""
                    End With

                    With moDataImportTemplate
                         If .GraphQLQuery = Nothing Then .GraphQLQuery = ""
                    End With

                    With moDataImportTemplate
                         If .GraphQLRootNode = Nothing Then .GraphQLRootNode = ""
                    End With

                    With txtName
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "Name"))
                    End With

                    With txtAPIEndpoint
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "APIEndpoint"))
                    End With

                    With txtAPIEndpointSelect
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "APIEndpointSelect"))
                    End With

                    With txtDataTransportObject
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "DTOObject"))
                    End With

                    With txtWorkbookSheet
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "WorkbookSheetName"))
                    End With

                    With txtMemo
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "Notes"))
                    End With

                    With txtStatusCodeColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "StatusCodeColumn"))
                    End With

                    With txtStatusDescriptionColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "StatusDescirptionColumn"))
                    End With

                    With txtReturnValueName
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "ReturnNodeValue"))
                    End With

                    With txtReturnValueColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "ReturnCell"))
                    End With

                    With txtReturnCellDTO
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "ReturnCellDTO"))
                    End With

                    With txtAPIQuery
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "UpdateQuery"))
                    End With

                    With txtSelectQuery
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "SelectQuery"))
                    End With

                    With icbType
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moDataImportTemplate, "ImportType"))
                    End With

                    With sePriority
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("EditValue", moDataImportTemplate, "Priority"))
                    End With

                    With txtGraphQLQuery
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "GraphQLQuery"))
                    End With


                    With txtGraphQLNode
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "GraphQLRootNode"))
                    End With


                    With txtEntityColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "EntityColumn"))
                    End With

                    With txtFileLocationColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "FileLocationColumn"))
                    End With

                    With txtGroup
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moDataImportTemplate, "Group"))
                    End With






                    gridValidators.DataSource = moDataImportTemplate.Validators
                    gdImportColumns.BestFitColumns()
                    gdValidators.BestFitColumns()





               End If
          End Set
     End Property

End Class
