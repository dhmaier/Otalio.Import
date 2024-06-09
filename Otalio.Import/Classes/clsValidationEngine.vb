Public Class clsValidationEngine
     Private mbContainsHierarchies As Boolean
     Private mbCancel As Boolean

     ' Add references to external dependencies
     Private UcProperties1 As UcProperties
     Private spreadsheetControl As SpreadsheetControl
     Private UcConnectionDetails1 As ucConnectionDetails

     Public Sub New(UcProperties As UcProperties, spreadsheetCtrl As SpreadsheetControl, UcConnectionDetails As ucConnectionDetails)
          mbContainsHierarchies = False
          mbCancel = False
          UcProperties1 = UcProperties
          spreadsheetControl = spreadsheetCtrl
          UcConnectionDetails1 = UcConnectionDetails
     End Sub

     Public Sub ValidateDataTemplate(oImportHeader As clsDataImportHeader)
          Try
               Dim oTotalVariables As New List(Of clsVariables)
               Dim oTotalActiveColumns As New List(Of clsImportColum)

               For Each oTemplate As clsDataImportTemplateV2 In oImportHeader.Templates
                    InitializeTemplate(oTemplate, oImportHeader)
                    Dim oColumns = ExtractAndFormatColumns(oTemplate)
                    Dim oActiveColumns = ExtractActiveColumns(oTemplate, oColumns)
                    Dim oVariables = ExtractVariables(oColumns)

                    UpdateTemplateVariables(oTemplate, oVariables)
                    CheckHierarchies(oTemplate)

                    ValidateTranslationsValidators(oTemplate, gbTranslationsEnabled)
                    ValidateTemplateValidators(oTemplate, oActiveColumns)

                    AddSpecialColumns(oTemplate, oImportHeader, oActiveColumns)

                    oTemplate.ImportColumns = oActiveColumns.OrderBy(Function(n) n.No).ToList()
                    oTotalActiveColumns.AddRange(oActiveColumns)
                    oTotalVariables.AddRange(oVariables)
               Next

               UpdateUIWithImportData(oTotalActiveColumns, oTotalVariables, oImportHeader)

               LoadLogs(oImportHeader.ID)
          Catch ex As Exception
               UcProperties1.gridImportColumns.DataSource = Nothing
          End Try
     End Sub

     Public Function ValidateData(poImportHeader As clsDataImportHeader) As Boolean
          If Not ValidateHierarchiesIsSelected() Then Return False

          ClearStatusColumns(poImportHeader, poImportHeader.Templates.Count = 1)
          FormatExcelSheet(poImportHeader)
          Application.DoEvents()

          Dim bHasErrors As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim dStartDateTime As DateTime = Now

          Try
               goHTTPServer.LogEvent("User has started validation of Template", "Validation", poImportHeader.Name, poImportHeader.ID)
               SetEditExcelMode(True, poImportHeader)

               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates
                    If oTemplate IsNot Nothing AndAlso oTemplate.Validators IsNot Nothing Then
                         ValidateTemplateValidators(poImportHeader, oTemplate, bHasErrors, bAutoColumnWidth)
                    End If
               Next

               If bAutoColumnWidth Then spreadsheetControl.ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)
               HideShowColumns(poImportHeader, mbHideCalulationColumns)
               UpdateProgressStatus()

               Return bHasErrors
          Catch ex As Exception
               UpdateProgressStatus()
               MsgBox($"Error code {ex.HResult} - {ex.Message}{vbNewLine}{ex.StackTrace}")
          Finally
               SetEditExcelMode(False, poImportHeader)
               EnableCancelButton(False)
               mbCancel = False
               UpdateProgressStatus()
               goHTTPServer.LogEvent($"Completed validation in {DateDiff(DateInterval.Minute, dStartDateTime, Now)} minutes", "Validation", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)
          End Try
     End Function

     Private Sub InitializeTemplate(oTemplate As clsDataImportTemplateV2, oImportHeader As clsDataImportHeader)
          If oTemplate.Selectors Is Nothing Then oTemplate.Selectors = New List(Of clsSelectors)()
          If oImportHeader.Templates.Count = 1 Then
               oTemplate.IsMaster = True
               oTemplate.IsEnabled = True
          End If
          mbContainsHierarchies = False
          oTemplate.GraphQLQuery = If(oTemplate.GraphQLQuery, "")
          oImportHeader.ID = If(oImportHeader.ID, GenerateGUID())
     End Sub

     Private Function ExtractAndFormatColumns(oTemplate As clsDataImportTemplateV2) As List(Of clsImportColum)
          Return ExactExcelColumns(oTemplate.DTOObjectFormated).OrderBy(Function(n) n.No).ToList()
     End Function

     Private Function ExtractActiveColumns(oTemplate As clsDataImportTemplateV2, oColumns As List(Of clsImportColum)) As List(Of clsImportColum)
          Dim oActiveColumns = oColumns.Where(Function(n) n.ColumnID <> String.Empty).OrderBy(Function(n) n.No).ToList()
          For Each oColumn In oActiveColumns
               oColumn.Name = $"{oTemplate.DataObject} ({oColumn.Name})"
               oColumn.DataType = "Data Source"
          Next
          Return oActiveColumns
     End Function

     Private Function ExtractVariables(oColumns As List(Of clsImportColum)) As List(Of clsVariables)
          Dim oVariables = New List(Of clsVariables)()
          For Each oColumn In oColumns
               If Not String.IsNullOrEmpty(oColumn.VariableName) Then
                    If Not oVariables.Exists(Function(n) n.Name = oColumn.VariableName) Then
                         oVariables.Add(New clsVariables With {.Name = oColumn.VariableName, .Value = String.Empty})
                    End If
               End If
          Next
          Return oVariables
     End Function

     Private Sub UpdateTemplateVariables(oTemplate As clsDataImportTemplateV2, oVariables As List(Of clsVariables))
          If oTemplate.Variables Is Nothing Then oTemplate.Variables = New List(Of clsVariables)()
          oTemplate.Variables.RemoveAll(Function(v) Not oVariables.Exists(Function(n) n.Name = v.Name))
          For Each oVariable In oVariables
               If Not oTemplate.Variables.Exists(Function(n) n.Name = oVariable.Name) Then
                    oTemplate.Variables.Add(oVariable)
               End If
          Next
     End Sub

     Private Sub CheckHierarchies(oTemplate As clsDataImportTemplateV2)
          Dim queriesToCheck = {oTemplate.SelectQuery, oTemplate.UpdateQuery, oTemplate.APIEndpoint, oTemplate.GraphQLQuery, oTemplate.DTOObject}
          For Each query In queriesToCheck
               If query.ToUpper().Contains("@@HIERARCHYID@@") Then
                    mbContainsHierarchies = True
                    Exit For
               End If
          Next
     End Sub

     Private Sub ValidateTemplateValidators(oTemplate As clsDataImportTemplateV2, oActiveColumns As List(Of clsImportColum))
          If oTemplate.Validators Is Nothing Then Return
          For Each oValidation As clsValidation In oTemplate.Validators
               If String.IsNullOrEmpty(oValidation.Visibility) Then oValidation.Visibility = "1"
               If Not String.IsNullOrEmpty(oValidation.Query) Then
                    oValidation.Query = oValidation.Query.Replace("  ", "").Replace(vbNewLine, "")
                    CheckHierarchiesInValidation(oValidation)
               End If
               ProcessValidationColumns(oValidation, oActiveColumns)
               If String.IsNullOrEmpty(oValidation.ReturnCell) = False Then
                    ProcessReturnCell(oValidation, oActiveColumns)
               End If
          Next
     End Sub

     Private Sub CheckHierarchiesInValidation(oValidation As clsValidation)
          Dim queriesToCheck = {oValidation.Query, oValidation.APIEndpoint}
          For Each query In queriesToCheck
               If query.ToUpper().Contains("@@HIERARCHYID@@") Then
                    mbContainsHierarchies = True
                    Exit For
               End If
          Next
     End Sub

     Private Sub ProcessValidationColumns(oValidation As clsValidation, oActiveColumns As List(Of clsImportColum))
          If oValidation.ValidationType <> "2" Then
               Dim oCols As List(Of clsImportColum) = ExtractListOfColumnsFromString(oValidation.Query)
               For Each oCol In oCols
                    If Not oActiveColumns.Exists(Function(n) n.ColumnName = oCol.ColumnName) Then
                         oCol.DataType = "Validator Query"
                         SetColumnParentAndDTOField(oValidation, oCol)
                         oActiveColumns.Add(oCol)
                    Else
                         Dim existingCol = oActiveColumns.First(Function(n) n.ColumnName = oCol.ColumnName)
                         existingCol.DataType = "Calculated Validator & DTO"
                    End If
               Next
          End If
     End Sub

     Private Sub SetColumnParentAndDTOField(oValidation As clsValidation, oCol As clsImportColum)
          If Not String.IsNullOrEmpty(oValidation.GraphQLSourceNode) Then
               oCol.Parent = oValidation.GraphQLSourceNode.Trim().Replace("!", "")
               oCol.DTOField = oCol.Name
          Else
               oCol.Name = $"{oValidation.DataObject} ({oCol.Name})"
          End If
     End Sub

     Private Sub ProcessReturnCell(oValidation As clsValidation, oActiveColumns As List(Of clsImportColum))
          Dim sCol As String = oValidation.ReturnCell
          Dim oCol As New clsImportColum With {
              .ColumnName = sCol,
              .Name = oValidation.Comments,
              .Type = JTokenType.None,
              .ColumnID = sCol,
              .DataType = "Calculated Validator"
          }
          Select Case oValidation.ValidationType
               Case "5", "6"
                    oCol.Name = oValidation.DetailedDescriptionFromReturnValue.ToString()
               Case "1", "4"
                    oCol.Name = oValidation.DetailedDescriptionFromQuery.ToString()
          End Select
          If sCol.Contains(":") Then
               Dim oValues As String() = sCol.Split(":"c)
               If oValues.Length > 1 Then oCol.ColumnName = oValues(0)
               If oValues.Length > 4 Then oCol.ArrayID = oValues(4)
          End If
          If Not oActiveColumns.Exists(Function(n) n.ColumnName = oCol.ColumnName) Then
               oActiveColumns.Add(oCol)
          Else
               Dim existingCol = oActiveColumns.First(Function(n) n.ColumnName = oCol.ColumnName)
               existingCol.DataType = "Calculated Validator & DTO"
          End If
     End Sub

     Private Sub AddSpecialColumns(oTemplate As clsDataImportTemplateV2, oImportHeader As clsDataImportHeader, oActiveColumns As List(Of clsImportColum))
          AddColumnIfNotExists(oTemplate.ReturnCell, oActiveColumns, "DTO Return Value", oTemplate.DataObject, oTemplate.ReturnNodeValue)
          AddColumnIfNotExists(oTemplate.StatusCodeColumn, oActiveColumns, "Response", $"Response Status Code {oTemplate.Name}")
          AddColumnIfNotExists(oTemplate.StatusDescriptionColumn, oActiveColumns, "Response", $"Response Status Description {oTemplate.Name}")
          AddColumnIfNotExists(oTemplate.ReturnCellDTO, oActiveColumns, "Response DTO", $"Response DTO {oTemplate.Name}")
          AddColumnIfNotExists(oImportHeader.StatusCodeColumn, oActiveColumns, "Response", $"Final Response Status Code {oImportHeader.Name}")
          AddColumnIfNotExists(oImportHeader.StatusDescriptionColumn, oActiveColumns, "Response", $"Final Status Description {oImportHeader.Name}")
     End Sub

     Private Sub AddColumnIfNotExists(columnID As String, oActiveColumns As List(Of clsImportColum), dataType As String, name As String, Optional returnValue As String = "")
          If String.IsNullOrEmpty(columnID) Then Return
          If Not oActiveColumns.Exists(Function(n) n.ColumnName = columnID) Then
               oActiveColumns.Add(New clsImportColum With {
                   .ColumnID = columnID,
                   .ColumnName = columnID,
                   .DataType = dataType,
                   .Name = If(String.IsNullOrEmpty(returnValue), name, $"{name} ({returnValue.Substring(returnValue.LastIndexOf(".") + 1)})")
               })
          End If
     End Sub

     Private Sub UpdateUIWithImportData(oTotalActiveColumns As List(Of clsImportColum), oTotalVariables As List(Of clsVariables), oImportHeader As clsDataImportHeader)
          UcProperties1.gridImportColumns.DataSource = oTotalActiveColumns
          UcProperties1.gridImportColumns.RefreshDataSource()
          UcProperties1.gridVariables.DataSource = oTotalVariables
          UcProperties1.gridVariables.RefreshDataSource()
          UpdateTemplatePositions(oImportHeader)
          UcProperties1.gridColumnListAll.DataSource = oImportHeader.ImportColumns
          UcProperties1.gridColumnListAll.RefreshDataSource()
          UcProperties1.gridValidatorsAll.DataSource = oImportHeader.Validators
          UcProperties1.gridValidatorsAll.RefreshDataSource()
          UcProperties1.gridValidators.RefreshDataSource()
     End Sub

     Private Sub UpdateTemplatePositions(oImportHeader As clsDataImportHeader)
          Dim nCounter As Integer = 0
          For Each oTemplate In oImportHeader.Templates
               nCounter += 1
               oTemplate.Position = nCounter
          Next
     End Sub

     Private Sub ValidateTemplateValidators(poImportHeader As clsDataImportHeader, oTemplate As clsDataImportTemplateV2, ByRef bHasErrors As Boolean, ByRef bAutoColumnWidth As Boolean)
          For Each oValidation In oTemplate.Validators.Where(Function(n) n.Enabled = "1").OrderBy(Function(n) n.Priority)
               If oValidation IsNot Nothing Then
                    goHTTPServer.LogEvent($"Validating {oValidation.Comments}", "Validation", poImportHeader.Name, poImportHeader.ID)

                    If Not ValidateWorksheetExists(poImportHeader) Then Return

                    EnableCancelButton(True)
                    mbCancel = False

                    Dim sColumns = ExtractColumnDetails(oValidation.Query)
                    If sColumns IsNot Nothing AndAlso sColumns.Count >= 0 Then
                         ValidateColumns(oValidation, poImportHeader, sColumns, bHasErrors, bAutoColumnWidth)
                    End If

                    If mbCancel Then
                         goHTTPServer.LogEvent("User requested Cancel of validation", "Validation", poImportHeader.Name, poImportHeader.ID)
                         Exit For
                    End If
               End If
          Next
     End Sub

     Private Function ValidateWorksheetExists(poImportHeader As clsDataImportHeader) As Boolean
          With spreadsheetControl
               If Not .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) Then
                    UpdateProgressStatus()
                    MsgBox($"Cannot find worksheet {poImportHeader.WorkbookSheetName}")
                    Return False
               End If
          End With
          Return True
     End Function

     Private Sub ValidateColumns(oValidation As clsValidation, poImportHeader As clsDataImportHeader, sColumns As List(Of String), ByRef bHasErrors As Boolean, ByRef bAutoColumnWidth As Boolean)
          With spreadsheetControl.ActiveWorksheet
               Dim nCountRows = .Rows.LastUsedIndex
               Dim nCountColumns = .Columns.LastUsedIndex

               If nCountColumns < 20 AndAlso nCountRows < 10000 Then bAutoColumnWidth = True
               If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then
                    ShowWaitDialogWithCancelCaption($"Validating {oValidation.Priority} - {oValidation.Comments}")
                    Dim eStatus = goHTTPServer.LogIntoIAM()

                    Dim bhadError As Boolean = False
                    Select Case oValidation.ValidationType
                         Case "0" 'Exits
                              bhadError = Exists(oValidation, spreadsheetControl, sColumns, oTemplate)
                         Case "1", "5", "7" 'FindAndReplace
                              bhadError = FindAndReplace(oValidation, spreadsheetControl, sColumns)
                         Case "2" 'Translate
                              bhadError = Translate(oValidation, spreadsheetControl, sColumns)
                         Case "3" 'Delete
                              bhadError = Deletes(oValidation, spreadsheetControl, sColumns)
                         Case "4", "6" 'FindAndReplaceList
                              bhadError = FindAndReplaceList(oValidation, spreadsheetControl, sColumns)
                         Case "8" ' Array of values
                              bhadError = ArryOfValues(oValidation, spreadsheetControl, sColumns)
                    End Select

                    If bhadError Then bHasErrors = True
               End If
          End With
     End Sub
End Class
