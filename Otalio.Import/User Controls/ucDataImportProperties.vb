Imports System.IO
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.Linq
Imports System.Net
Imports System.Threading
Imports System.IO.Compression
Imports DevExpress.XtraEditors
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Public Class ucDataImportProperties

     Private moImportHeader As New clsDataImportHeader
     Private moSelectedTemplate As New clsDataImportTemplateV2

     Public msFileName As String = ""
     Public msFilePath As String = ""
     Private mbLoading As Boolean = True
     Private isUpdatingDTO As Boolean = False
     Public linkedItemsCountDTO As Integer = 0
     Private isUpdatingSelect As Boolean = False
     Public linkedItemsCountSelect As Integer = 0
     Private isUpdatingUpdate As Boolean = False
     Public linkedItemsCountUpdate As Integer = 0

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

                    mbLoading = True
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

                    With txtHistory
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moImportHeader, "HistoryLog"))
                    End With

                    With chkReadOnly
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Checked", moImportHeader, "ReadOnlyImport"))
                    End With

                    Call LoadTemplates(value)

                    gridColumnListAll.DataSource = moImportHeader.ImportColumns
                    gdColumnsListAll.BestFitColumns()

                    gridValidatorsAll.DataSource = moImportHeader.Validators
                    gdValidatorsAll.BestFitColumns()

                    mbLoading = False


               End If
          End Set
     End Property

     Sub New()

          InitializeComponent()
          lcgColumns.Selected = True
          mbLoading = True



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
     Private Sub gdSelectors_DoubleClick(sender As Object, e As EventArgs) Handles gdSelectors.DoubleClick
          With gdSelectors
               If .SelectedRowsCount >= 0 Then
                    Dim oSelector As clsSelectors = TryCast(.GetRow(.FocusedRowHandle), clsSelectors)
                    If oSelector IsNot Nothing Then

                         Using oForm As New frmSelector
                              oForm.UcSelectors1._Selector = oSelector
                              Select Case oForm.ShowDialog
                                   Case DialogResult.OK 'ok edited
                                   Case DialogResult.Cancel 'dont do nothing
                                   Case DialogResult.Abort 'use to flag item to be deleted
                                        Me.moSelectedTemplate.Selectors.Remove(oForm.UcSelectors1._Selector)

                              End Select

                              Me.gridSelectors.RefreshDataSource()
                              gdSelectors.BestFitColumns()

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
          mbLoading = True
          linkedItemsCountDTO = 0

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
               If .SelectQuery = Nothing Then .SelectQuery = ""
               If .GraphQLQuery = Nothing Then .GraphQLQuery = ""
               If .Selectors Is Nothing Then .Selectors = New List(Of clsSelectors)
               If .UpdateQuery = Nothing Then .UpdateQuery = ""
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

          With recDTO
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

          With rtbAPIQuery
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Text", poTemplate, "UpdateQuery"))
          End With

          With rtbSelectQuery
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

          With ckIsEnabled
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Checked", poTemplate, "IsEnabled"))
          End With

          With ckIsMaster
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Checked", poTemplate, "IsMaster"))
          End With

          With chkIgnoreArray
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Checked", poTemplate, "IgnoreArray"))
          End With

          With chkRemoveEmptyAndNull
               .DataBindings.Clear()
               .DataBindings.Add(New Binding("Checked", poTemplate, "RemoveEmptyAndNull"))
          End With

          gridValidators.DataSource = poTemplate.Validators
          gridSelectors.DataSource = poTemplate.Selectors

          gdImportColumns.BestFitColumns()
          gdValidators.BestFitColumns()
          gdSelectors.BestFitColumns()


          mbLoading = False

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


     Private Sub ApplyFormatting(regex As Regex, color As Color, control As RichTextBox)
          Dim matches As System.Text.RegularExpressions.MatchCollection = regex.Matches(control.Text)
          For Each match As System.Text.RegularExpressions.Match In matches
               control.Select(match.Index, match.Length)
               control.SelectionColor = color
               control.SelectionFont = New Font(control.Font, FontStyle.Underline)
          Next
     End Sub

     Private Function CountElements(psElement As String) As Integer
          Dim text As String = ""
          Select Case psElement
               Case "DTO" : text = recDTO.Text
               Case "SELECT" : text = rtbSelectQuery.Text
               Case "UPDATE" : text = rtbAPIQuery.Text
          End Select


          ' Regex patterns to identify the values
          Dim regexShipId As New Regex("@@.*?@@")
          Dim regexSpecial As New Regex("<!.*?!>")

          ' Count total number of elements
          Dim shipIdCount As Integer = regexShipId.Matches(text).Count
          Dim specialCount As Integer = regexSpecial.Matches(text).Count

          Return shipIdCount + specialCount
     End Function

     Private Sub recDTO_TextChanged(sender As Object, e As EventArgs) Handles recDTO.TextChanged
          If isUpdatingDTO Then Return
          ' Save the current cursor position
          Dim cursorPosition As Integer = recDTO.SelectionStart

          ' Count current elements
          Dim currentLinkedItemsCount As Integer = CountElements("DTO")

          ' Only redraw if the number of elements has changed
          If currentLinkedItemsCount = linkedItemsCountDTO Then Return
          With recDTO
               Try
                    isUpdatingDTO = True


                    ' Suspend layout updates to prevent flickering
                    .SuspendLayout()

                    ' Clear all formatting
                    .SelectAll()
                    .SelectionColor = Color.Black
                    .SelectionFont = New Font(.Font, FontStyle.Regular)

                    ' Regex patterns to identify the values
                    Dim regexShipId As New Regex("@@.*?@@")
                    Dim regexSpecial As New Regex("<!.*?!>")

                    ' Apply formatting for @@...@@ values
                    ApplyFormatting(regexShipId, Color.DarkGreen, recDTO)

                    ' Apply formatting for <!...!> values
                    ApplyFormatting(regexSpecial, Color.DodgerBlue, recDTO)

                    ' Update the linked items count
                    linkedItemsCountDTO = currentLinkedItemsCount

               Finally
                    ' Restore the cursor position
                    .SelectionStart = cursorPosition
                    .SelectionLength = 0
                    .ScrollToCaret()

                    ' Resume layout updates
                    .ResumeLayout()
                    isUpdatingDTO = False
               End Try
          End With

     End Sub

     Private Sub rtbSelectQuery_TextChanged(sender As Object, e As EventArgs) Handles rtbSelectQuery.TextChanged
          If isUpdatingSelect Then Return

          With rtbSelectQuery
               ' Save the current cursor position
               Dim cursorPosition As Integer = .SelectionStart

               ' Count current elements
               Dim currentLinkedItemsCount As Integer = CountElements("SELECT")

               ' Only redraw if the number of elements has changed
               If currentLinkedItemsCount = linkedItemsCountSelect Then Return

               Try
                    isUpdatingSelect = True


                    ' Suspend layout updates to prevent flickering
                    .SuspendLayout()

                    ' Clear all formatting
                    .SelectAll()
                    .SelectionColor = Color.Black
                    .SelectionFont = New Font(.Font, FontStyle.Regular)

                    ' Regex patterns to identify the values
                    Dim regexShipId As New Regex("@@.*?@@")
                    Dim regexSpecial As New Regex("<!.*?!>")

                    ' Apply formatting for @@...@@ values
                    ApplyFormatting(regexShipId, Color.DarkGreen, rtbSelectQuery)

                    ' Apply formatting for <!...!> values
                    ApplyFormatting(regexSpecial, Color.DodgerBlue, rtbSelectQuery)

                    ' Update the linked items count
                    linkedItemsCountSelect = currentLinkedItemsCount

               Finally
                    ' Restore the cursor position
                    .SelectionStart = cursorPosition
                    .SelectionLength = 0
                    .ScrollToCaret()

                    ' Resume layout updates
                    .ResumeLayout()
                    isUpdatingSelect = False
               End Try
          End With

     End Sub

     Private Sub rtbAPIQuery_TextChanged(sender As Object, e As EventArgs) Handles rtbAPIQuery.TextChanged
          If isUpdatingUpdate Then Return
          With rtbAPIQuery
               ' Save the current cursor position
               Dim cursorPosition As Integer = .SelectionStart

               ' Count current elements
               Dim currentLinkedItemsCount As Integer = CountElements("UPDATE")

               ' Only redraw if the number of elements has changed
               If currentLinkedItemsCount = linkedItemsCountUpdate Then Return

               Try
                    isUpdatingUpdate = True

                    ' Suspend layout updates to prevent flickering
                    .SuspendLayout()

                    ' Clear all formatting
                    .SelectAll()
                    .SelectionColor = Color.Black
                    .SelectionFont = New Font(.Font, FontStyle.Regular)

                    ' Regex patterns to identify the values
                    Dim regexShipId As New Regex("@@.*?@@")
                    Dim regexSpecial As New Regex("<!.*?!>")

                    ' Apply formatting for @@...@@ values
                    ApplyFormatting(regexShipId, Color.DarkGreen, rtbAPIQuery)

                    ' Apply formatting for <!...!> values
                    ApplyFormatting(regexSpecial, Color.DodgerBlue, rtbAPIQuery)

                    ' Update the linked items count
                    linkedItemsCountUpdate = currentLinkedItemsCount

               Finally
                    ' Restore the cursor position
                    .SelectionStart = cursorPosition
                    .SelectionLength = 0
                    .ScrollToCaret()

                    ' Resume layout updates
                    .ResumeLayout()
                    isUpdatingUpdate = False
               End Try
          End With

     End Sub
End Class
