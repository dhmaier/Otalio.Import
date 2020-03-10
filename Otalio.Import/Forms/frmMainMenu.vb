﻿Imports System.ComponentModel
Imports System.Text
Imports DevExpress.Skins
Imports DevExpress.LookAndFeel
Imports DevExpress.UserSkins
Imports DevExpress.XtraBars.Helpers
Imports DevExpress.XtraBars.Ribbon
Imports System.IO
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.Linq
Imports DevExpress.Spreadsheet
Imports System.Net
Imports RestSharp
Imports DevExpress.XtraSpreadsheet
Imports DevExpress.XtraEditors


Public Class frmMainMenu

  Private mbCancel As Boolean = False


#Region "Form"

  Private Const EventLogHistory As Integer = 1000000
  Private msFilePath As String = ""
  Private msFileName As String
  Private moCopyObject As New List(Of clsValidation)
  Private mbForceValidate As Boolean = True
  Private mbHideCalulationColumns As Boolean = True
  Private mbContainsHierarchies As Boolean = False


  Sub New()
    InitSkins()
    InitializeComponent()
    Me.InitSkinGallery()

  End Sub

  Sub InitSkins()
    DevExpress.Skins.SkinManager.EnableFormSkins()
    DevExpress.UserSkins.BonusSkins.Register()
  End Sub

  Private Sub InitSkinGallery()
    SkinHelper.InitSkinGallery(rgbiSkins, True)
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

    UcProperties1._DataImportTemplate = New clsDataImportTemplate
    UcConnectionDetails1.LoadSettingFile()
    siInfo.Caption = String.Format("Version:{0}", My.Application.Info.Version)

    AddHandler goHTTPServer.APICallEvent, AddressOf APICallEvent
    AddHandler goHTTPServer.ErrorEvent, AddressOf ErrorEvent



  End Sub

#End Region

#Region "Buttons"

  Private Sub bbiImportOpen_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportOpen.ItemClick


    Dim fd As OpenFileDialog = New OpenFileDialog
    fd.Multiselect = False
    fd.Title = "Select a Otalio Dynamic import template."
    fd.Filter = "Dynamic import Template |*.dit"
    fd.FilterIndex = 1
    fd.FileName = ""
    fd.RestoreDirectory = True
    If fd.ShowDialog() = DialogResult.OK Then

      Dim sFilename As String = fd.FileName

      If File.Exists(sFilename) Then

        Dim oDataimportTemplate As clsDataImportTemplate = TryCast(LoadFile(sFilename), clsDataImportTemplate)

        If oDataimportTemplate IsNot Nothing Then

          UcProperties1.msFilePath = sFilename
          UcProperties1.msFileName = fd.SafeFileName

          If oDataimportTemplate.Validators Is Nothing Then
            oDataimportTemplate.Validators = New List(Of clsValidation)
          End If

          goOpenWorkBook.Templates.Add(oDataimportTemplate)
          loadWorkbook()



        Else
          MsgBox("Failed to load Data import template file")
        End If

      End If


    End If

  End Sub

  Private Sub bbiSaveImport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveImport.ItemClick
    Call ValidateDataTempalte(UcProperties1._DataImportTemplate)
    Call UcProperties1.SaveSettings()
  End Sub

  Private Sub bbiSaveImportAs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveImportAs.ItemClick
    Call ValidateDataTempalte(UcProperties1._DataImportTemplate)
    Call UcProperties1.SaveSettings(True)
  End Sub

  Private Sub bbiCreateEmptyExcelDocument_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCreateEmptyExcelDocument.ItemClick

    For Each oItem In lbxImportTemplates.CheckedItems
      Call createNewExcelSheet(TryCast(oItem, clsDataImportTemplate))
      Application.DoEvents()
    Next


  End Sub

  Private Sub bbiImportExcel_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportExcel.ItemClick

    If lbxImportTemplates.CheckedItemsCount = 0 Then
      MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
      Exit Sub
    End If

    If MsgBox("You are about to import data from the Excel sheet.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then

      For Each oItem In lbxImportTemplates.CheckedItems
        ImportData(TryCast(oItem, clsDataImportTemplate))
        Application.DoEvents()
      Next

    End If


  End Sub

  Private Sub bbiValidateData_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidateData.ItemClick

    If lbxImportTemplates.CheckedItemsCount = 0 Then
      MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
      Exit Sub
    End If

    For Each oItem In lbxImportTemplates.CheckedItems
      Call ValidateDataTempalte(TryCast(oItem, clsDataImportTemplate))
      Call ValidateData(TryCast(oItem, clsDataImportTemplate))
      Application.DoEvents()
    Next


  End Sub

  Private Sub bbiFormatJsonCode_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiFormatJsonCode.ItemClick
    Try
      UcProperties1.txtDataTransportObject.Text = JValue.Parse(UcProperties1.txtDataTransportObject.Text).ToString(Newtonsoft.Json.Formatting.Indented)
    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try
  End Sub

  Private Sub bbiValidatorAdd_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidatorAdd.ItemClick

    Using oForm As New frmValidatorEdit
      oForm._ucValidationProperties._Validation = New clsValidation
      Select Case oForm.ShowDialog
        Case DialogResult.OK 'ok add

          UcProperties1._DataImportTemplate.Validators.Add(oForm._ucValidationProperties._Validation)
        Case DialogResult.Cancel 'dont add
        Case DialogResult.Abort 'use to flag item to be deleted
          UcProperties1._DataImportTemplate.Validators.Remove(oForm._ucValidationProperties._Validation)

      End Select

      UcProperties1.gridValidators.RefreshDataSource()
    End Using

  End Sub

  Private Sub bbiDuplicate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiDuplicate.ItemClick

    If moCopyObject IsNot Nothing Then

      For Each oValidation As clsValidation In moCopyObject
        If oValidation IsNot Nothing Then
          Dim oNewValidation = oValidation.Clone
          UcProperties1._DataImportTemplate.Validators.Add(oNewValidation)
        End If
      Next
      UcProperties1.gridValidators.RefreshDataSource()
    End If

  End Sub

  Private Sub bbiCopyValidator_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCopyValidator.ItemClick
    'clear the existing cash of validations
    moCopyObject = New List(Of clsValidation)

    Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdValidators.GetSelectedRows.ToList, List(Of Integer))
    If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then

      For Each orow As Integer In oSelectedRows
        Dim oValidation As clsValidation = TryCast(UcProperties1.gdValidators.GetRow(orow), clsValidation)
        If oValidation IsNot Nothing Then
          moCopyObject.Add(oValidation)
        End If
        bbiDuplicate.Enabled = True
      Next

    End If
  End Sub

  Private Sub bbiValidatorDelete_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidatorDelete.ItemClick

    Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdValidators.GetSelectedRows.ToList, List(Of Integer))

    If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then
      Dim oValidation As clsValidation = TryCast(UcProperties1.gdValidators.GetRow(oSelectedRows(0)), clsValidation)
      If oValidation IsNot Nothing Then
        UcProperties1._DataImportTemplate.Validators.Remove(oValidation)
        UcProperties1.gridValidators.RefreshDataSource()
      End If
    End If

  End Sub

  Private Sub bbiOpenWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiOpenWorkbook.ItemClick

    Dim fd As OpenFileDialog = New OpenFileDialog
    fd.Multiselect = False
    fd.Title = "Select a Otalio Dynamic import workbook."
    fd.Filter = "Dynamic import workbook |*.ditw"
    fd.FilterIndex = 1
    fd.FileName = txtWorkbookName.Text
    fd.RestoreDirectory = True
    If fd.ShowDialog() = DialogResult.OK Then

      Dim sFilename As String = fd.FileName

      If File.Exists(sFilename) Then

        msFilePath = Path.GetDirectoryName(sFilename)
        msFileName = fd.SafeFileName
        txtWorkbookName.EditValue = msFileName


        Dim oWorkbook As clsWorkbook = TryCast(LoadFile(sFilename), clsWorkbook)

        If oWorkbook IsNot Nothing Then

          goOpenWorkBook = oWorkbook
          goOpenWorkBook.WorkbookName = msFileName

          loadWorkbook()

          'check if there is a data excel sheet in the same folder with the same name
          If File.Exists(Replace(sFilename, ".ditw", ".xlsx")) Then
            spreadsheetControl.LoadDocument(Replace(sFilename, ".ditw", ".xlsx"))
          End If


        Else
          MsgBox("Failed to load Data import workbook file",, "Warning...")
        End If

      End If


    End If
  End Sub

  Private Sub bbiSaveWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveWorkbook.ItemClick
    SaveWorkbook(False)
  End Sub

  Private Sub bbiSaveAsWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveAsWorkbook.ItemClick
    SaveWorkbook(True)
  End Sub

  Private Sub bbiAddTemplate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiAddTemplate.ItemClick
    goOpenWorkBook.Templates.Add(New clsDataImportTemplate)
    loadWorkbook()
  End Sub

  Private Sub bbiRemoveTemplate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiRemoveTemplate.ItemClick
    With lbxImportTemplates

      If .SelectedIndex >= 0 Then
        Dim oItem = .GetItem(.SelectedIndex)
        If oItem IsNot Nothing Then
          goOpenWorkBook.Templates.Remove(TryCast(oItem, clsDataImportTemplate))
        End If
      End If

      loadWorkbook()

    End With
  End Sub

  Private Sub bbiCancel_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCancel.ItemClick
    If mbCancel = False Then mbCancel = True
  End Sub

  Private Sub bbiSaveAll_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveAll.ItemClick

    If String.IsNullOrEmpty(msFilePath) = False Then
      lbxImportTemplates.Focus()
      Call ValidateDataTempalte(UcProperties1._DataImportTemplate)

      Dim sPath As String = Path.GetDirectoryName(msFilePath)
      Dim sFileName As String = Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName)
      Dim sExtention As String = Path.GetExtension(goOpenWorkBook.WorkbookName)
      Dim sVersion As String = goOpenWorkBook.WorkbookVersion

      'create backups
      SaveWorkbook(False, Path.GetDirectoryName(msFilePath), "BACKUP",
                 String.Format("{0} {1}", Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName), sVersion),
                 Path.GetExtension(goOpenWorkBook.WorkbookName), False)



      SaveWorkbook(False)
      Dim sCurrentFilename As String = spreadsheetControl.Document.Path
      If String.IsNullOrEmpty(sCurrentFilename) = True Then
        sCurrentFilename = String.Format("{0}\{1}.{2}", sPath, sFileName, "xlsx")
      End If

      Dim sExcelExtention As String = Path.GetExtension(sCurrentFilename)
      spreadsheetControl.SaveDocument(String.Format("{0}\{1}\{2} {3}{4}", sPath, "BACKUP", sFileName, sVersion, sExcelExtention))

      spreadsheetControl.SaveDocument(sCurrentFilename)
    End If

  End Sub

  Private Sub bbiQuery_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiQuery.ItemClick

    If lbxImportTemplates.CheckedItemsCount = 0 Then
      MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
      Exit Sub
    End If

    If MsgBox("Running the query will remove all existing data from the excel worksheet.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then

      For Each oItem In lbxImportTemplates.CheckedItems
        QueryandLoad(TryCast(oItem, clsDataImportTemplate))
        Application.DoEvents()
      Next


    End If


  End Sub

#End Region

#Region "Functions"

  Public Sub APICallEvent(psRequest As RestRequest, psResponse As IRestResponse)
    Try

      If Me.UcConnectionDetails1.chkEnableLoging.Checked Then

        If psResponse IsNot Nothing Then

          Dim sText As String = ""

          For Each oPar In psResponse.Headers
            sText = sText + String.Format("Header:{0}", oPar.ToString) & vbNewLine
          Next

          If psResponse.Content IsNot Nothing Then


            Dim jsServerResponse As JObject = JObject.Parse(psResponse.Content)
            If jsServerResponse IsNot Nothing Then

              sText = sText + String.Format("Body:{0}", jsServerResponse.ToString) & vbNewLine & vbNewLine & vbNewLine
            End If

          End If

          txtLogs.EditValue = String.Format("Response:{0}{0}{1}{0}{0}{2}", vbNewLine, sText.ToString, txtLogs.EditValue)
        End If

        If txtLogs.EditValue.ToString.Length > EventLogHistory Then
          txtLogs.EditValue = txtLogs.EditValue.ToString.Substring(0, EventLogHistory)
        End If
      End If

    Catch ex As Exception

    End Try

  End Sub

  Public Sub ErrorEvent(poExcetpion As Exception)
    If poExcetpion IsNot Nothing Then

      Dim sText As String = ""

      sText = sText + String.Format("Code:{0} - Message: {1} {3}Stack:{2}", poExcetpion.HResult, poExcetpion.Message, poExcetpion.StackTrace, vbNewLine) & vbNewLine & vbNewLine & vbNewLine

      txtErrors.EditValue = sText & txtErrors.EditValue & vbNewLine

    End If

    If txtErrors.EditValue.ToString.Length > EventLogHistory Then
      txtErrors.EditValue = txtErrors.EditValue.ToString.Substring(0, EventLogHistory)
    End If

  End Sub

  Public Sub createNewExcelSheet(poImportTemplate As clsDataImportTemplate)

    If poImportTemplate IsNot Nothing Then
      ValidateDataTempalte(poImportTemplate)

      If String.IsNullOrEmpty(poImportTemplate.WorkbookSheetName.ToString) = False Then


        With spreadsheetControl
          If .Document.Worksheets.Contains(poImportTemplate.WorkbookSheetName) = False Then
            .Document.Worksheets.Add(StrConv(poImportTemplate.WorkbookSheetName, vbProperCase))
            .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
          Else
            'if its found make it the active worksheet
            .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
          End If

          With .ActiveWorksheet
            For Each oImportColumn In poImportTemplate.ImportColumns.OrderBy(Function(n) n.ColumnID)

              Dim sColumnName As String = ""
              If String.IsNullOrEmpty(oImportColumn.Parent) = False Then
                sColumnName = String.Format("{0}.{1}", oImportColumn.Parent, oImportColumn.Name)
              Else
                sColumnName = oImportColumn.Name
              End If

              .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value = sColumnName

              If oImportColumn.Parent = "translations.en" Then
                .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value = "English"
              End If
            Next

            For Each oValadation In poImportTemplate.Validators
              If String.IsNullOrEmpty(oValadation.Comments) = False Then
                Dim oCell As String = oValadation.ReturnCell
                If oValadation.ReturnCell.Contains(":") Then
                  oCell = oValadation.ReturnCell.Substring(0, oValadation.ReturnCell.IndexOf(":"))
                End If
                If String.IsNullOrEmpty(oCell) = False Then
                  .Cells(String.Format("{0}1", oCell)).Value = String.Format("({0})", oValadation.Comments)
                End If

              End If
            Next

          End With



        End With

      End If
    End If

  End Sub

  Private Sub ValidateDataTempalte(poImportTemplate As clsDataImportTemplate)
    Try

      mbContainsHierarchies = False

      'extract list of columns from the Json code
      Dim oColumns As List(Of clsImportColum) = ExactExcelColumns(poImportTemplate.DTOObject).OrderBy(Function(n) n.No).ToList

      'extract only column used for the data import
      Dim oActiveColumns As List(Of clsImportColum) = oColumns.Where(Function(n) n.ColumnID <> "").OrderBy(Function(n) n.No).ToList

      Dim oVariables As New List(Of clsVariables)

      'look for variables
      For Each oColumn As clsImportColum In oColumns
        If oColumn.VariableName IsNot Nothing AndAlso String.IsNullOrEmpty(oColumn.VariableName) = False Then
          With oVariables
            If .Exists(Function(n) n.Name = oColumn.Name) = False Then
              .Add(New clsVariables With {.Name = oColumn.VariableName, .Value = ""})
            End If
          End With
        End If
      Next

      'remove any unused ones
      With poImportTemplate
        If .Variables IsNot Nothing AndAlso .Variables.Count > 0 Then
          For Each oVariable As clsVariables In .Variables
            If oVariables.Exists(Function(n) n.Name = oVariable.Name) = False Then
              .Variables.Remove(oVariable)
            End If
          Next
        Else
          .Variables = New List(Of clsVariables)
        End If
      End With

      With oVariables
        If oVariables IsNot Nothing AndAlso .Count > 0 Then
          For Each oVariable As clsVariables In oVariables
            If poImportTemplate.Variables.Exists(Function(n) n.Name = oVariable.Name) = False Then
              .Add(oVariable)
            End If
          Next
        End If
      End With

      If mbContainsHierarchies = False Then mbContainsHierarchies = poImportTemplate.SelectQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
      If mbContainsHierarchies = False Then mbContainsHierarchies = poImportTemplate.UpdateQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
      If mbContainsHierarchies = False Then mbContainsHierarchies = poImportTemplate.APIEndpoint.ToString.ToUpper.Contains("@@HIERARCHYID@@")
      If mbContainsHierarchies = False Then mbContainsHierarchies = poImportTemplate.GraphQLQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
      If mbContainsHierarchies = False Then mbContainsHierarchies = poImportTemplate.DTOObject.ToString.ToUpper.Contains("@@HIERARCHYID@@")


      If poImportTemplate.Validators IsNot Nothing Then

        For Each oValidation As clsValidation In poImportTemplate.Validators
          If String.IsNullOrEmpty(oValidation.Visibility) Then
            oValidation.Visibility = "1"
          End If
          If String.IsNullOrEmpty(oValidation.Query) = False Then
            oValidation.Query = Replace(oValidation.Query, " ", "")
            oValidation.Query = Replace(oValidation.Query, vbNewLine, "")

            If mbContainsHierarchies = False Then mbContainsHierarchies = oValidation.Query.ToString.ToUpper.Contains("@@HIERARCHYID@@")
            If mbContainsHierarchies = False Then mbContainsHierarchies = oValidation.APIEndpoint.ToString.ToUpper.Contains("@@HIERARCHYID@@")

          End If



          If oValidation.ValidationType <> "2" Then
            Dim oCols As List(Of clsImportColum) = ExtractListOfColumnsFromString(oValidation.Query)

            If oCols IsNot Nothing AndAlso oCols.Count > 0 Then
              For Each oCol In oCols.ToList
                If oActiveColumns.Exists(Function(n) n.ColumnName = oCol.ColumnName) = False Then

                  If String.IsNullOrEmpty(oValidation.GraphQLSourceNode) = False Then
                    oCol.Parent = oValidation.GraphQLSourceNode.ToString.Trim
                  Else
                    oCol.Name = String.Format("{0}-{1}", oCol.Name, oValidation.Priority)
                  End If

                  oActiveColumns.Add(oCol)
                End If
              Next
            End If
          End If

        Next
      End If


      'pass the data to the template 
      With UcProperties1
        ._DataImportTemplate.ImportColumns = TryCast(oActiveColumns.OrderBy(Function(n) n.No).ToList, List(Of clsImportColum))
        .gridImportColumns.DataSource = oActiveColumns
      End With

      'pass the data to the template 
      With UcProperties1
        .gridVariables.DataSource = ._DataImportTemplate.Variables.ToList
      End With

    Catch ex As Exception
      With UcProperties1
        ._DataImportTemplate.ImportColumns = Nothing
        .gridImportColumns.DataSource = Nothing
      End With

    End Try
  End Sub

  Public Function ExactExcelColumns(poJsonObject As Object) As List(Of clsImportColum)
    Try

      If poJsonObject IsNot Nothing AndAlso String.IsNullOrEmpty(poJsonObject) = False Then
        Dim oColumns As New List(Of clsImportColum)

        Dim nodes As Dictionary(Of String, String) = New Dictionary(Of String, String)()
        Dim rootObject As JObject = JObject.Parse(poJsonObject)
        goJsonHelper.ParseJson(rootObject, nodes)
        For Each key As String In nodes.Keys
          ExtractColumnDetails2(nodes(key), key, oColumns)
        Next

        For Each oCol As clsImportColum In oColumns
          Dim oNode = rootObject.SelectToken(IIf(oCol.Parent = "", oCol.Name, String.Format("{0}", oCol.Parent, oCol.Name)))
          If oNode IsNot Nothing Then
            oCol.Type = oNode.Type
          End If
        Next

        Return oColumns



        If poJsonObject IsNot Nothing Then
          Select Case poJsonObject.GetType.Name.ToString.ToUpper

            Case "JOBJECT"
              Dim oJsonObject As JObject = JObject.Parse(poJsonObject)

              If oJsonObject IsNot Nothing Then
                For Each oNode In oJsonObject

                  Select Case oNode.Value.Type

                    Case JTokenType.Array
                      Dim oNodes = oNode.Value.Children
                      ExactExcelColumns(oNode.Value.Children)

                    Case JTokenType.Object

                      Dim oNodes = oNode.Value.Children
                      Dim oChildColumns As New List(Of clsImportColum)

                      For Each oChild In oNodes
                        oChildColumns.AddRange(ExtractExcelColumnChildern(oChild))
                      Next

                      oColumns.AddRange(oChildColumns)

                    Case Else

                      oColumns.Add(ExtractColumnDetails(oNode, oNode.Value.Type))

                  End Select
                Next
              End If
            Case ""
              Dim oJsonObject As JEnumerable(Of JToken) = poJsonObject
              If oJsonObject.Count > 0 Then
                For Each oNode As JToken In oJsonObject

                  Select Case oNode.Type

                    Case JTokenType.Array
                      Dim oNodes = oNode.Children
                      ExactExcelColumns(oNode.Children)

                    Case JTokenType.Object

                      Dim oNodes = oNode.Children
                      Dim oChildColumns As New List(Of clsImportColum)

                      For Each oChild In oNodes
                        oChildColumns.AddRange(ExtractExcelColumnChildern(oChild))
                      Next

                      oColumns.AddRange(oChildColumns)

                    Case Else

                      oColumns.Add(ExtractColumnDetails(oNode, oNode.Type))

                  End Select
                Next
              End If
          End Select
        End If





        Return oColumns

      End If

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try

  End Function

  Public Function ExtractExcelColumnChildern(poJsonChild As Object) As List(Of clsImportColum)
    Try
      Dim oColumns As New List(Of clsImportColum)
      Dim oObject As JProperty = TryCast(poJsonChild, JProperty)

      If oObject IsNot Nothing AndAlso oObject.Value.Type = JTokenType.String Then
        oColumns.Add(ExtractColumnDetails(oObject, oObject.Type))

      Else
        For Each oChild In poJsonChild
          Select Case oChild.Type
            Case JTokenType.Object
              oColumns.AddRange(ExtractExcelColumnChildern(oChild))
            Case Else
              oColumns.Add(ExtractColumnDetails(oChild, oChild.type))
          End Select
        Next
      End If
      Return oColumns

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try


  End Function

  Public Function ExtractColumnDetails(poObject As Object, pnJkokenType As JTokenType) As clsImportColum

    Try

      Dim sColumn As String = ""
      Dim sColumnName As String = ""
      Dim sVariable As String = ""
      Dim sFormat As String = ""
      If InStr(poObject.Value.ToString, "<!") > 0 Then
        sColumn = (Replace(Replace(poObject.Value.ToString, "<!", ""), "!>", ""))
      End If

      sColumn = sColumn.Replace("[", "")
      sColumn = sColumn.Replace("]", "")
      sColumn = sColumn.Replace(" ", "")
      sColumn = sColumn.Replace("""", "")
      sColumn = sColumn.Replace(vbNewLine, "")

      If sColumn.Contains(":") Then
        Dim sValues As List(Of String) = sColumn.Split(":").ToList
        If sValues IsNot Nothing AndAlso sValues.Count > 1 Then
          sColumnName = sValues(0)
          sFormat = sValues(1)
          If sValues.Count > 2 Then
            sVariable = sValues(2)
          End If
        End If
      Else
        sColumnName = sColumn
      End If


      Select Case pnJkokenType
        Case JTokenType.Property : Return (New clsImportColum With {.Name = poObject.Name, .ColumnID = sColumn,
            .Type = poObject.Type, .Parent = poObject.Path, .Formatted = sFormat, .ColumnName = sColumnName,
            .VariableName = sVariable})
        Case Else : Return (New clsImportColum With {.Name = poObject.Key, .ColumnID = sColumn,
            .Type = poObject.Value.Type, .Parent = poObject.value.path, .Formatted = sFormat, .ColumnName = sColumnName,
            .VariableName = sVariable})
      End Select

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try


  End Function

  Public Sub ExtractColumnDetails2(poObject As String, psParent As String, poListOfColumn As List(Of clsImportColum))

    Try

      Dim sColumnData As String = ""
      Dim sColumnHeader As String = ""
      Dim sColumnFieldName As String = ""
      Dim sVariable As String = ""
      Dim sFormat As String = ""
      Dim sChild As String = ""
      Dim sParent As String = ""

      If InStr(poObject, "<!") > 0 Then
        sColumnData = (Replace(Replace(poObject, "<!", ""), "!>", ""))
      End If

      'sColumnData = sColumnData.Replace("[", "")
      'sColumnData = sColumnData.Replace("]", "")
      'sColumnData = sColumnData.Replace(" ", "")
      'sColumnData = sColumnData.Replace("""", "")
      'sColumnData = sColumnData.Replace(vbNewLine, "")

      If sColumnData.Contains(":") Then
        Dim sValues As List(Of String) = sColumnData.Split(":").ToList
        If sValues IsNot Nothing AndAlso sValues.Count > 1 Then
          'columnn name
          sColumnHeader = sValues(0)
          'column formating
          sFormat = sValues(1)
          'column arry objects
          If sValues.Count > 2 Then
            sChild = sValues(2)
          End If

          If sValues.Count > 3 Then
            sVariable = sValues(3)
          End If

        End If
      Else
        sColumnHeader = sColumnData
      End If

      'if this is a sub noode get properites name
      If psParent.Contains(".") Then
        Dim sNodes As String() = psParent.ToString.Split(".")
        If sNodes IsNot Nothing Then
          Dim nCount As Integer = 1
          For Each snode As String In sNodes
            If nCount = sNodes.Count Then
              sColumnFieldName = snode
            Else
              sParent = sParent & snode & "."
            End If
            nCount += 1
          Next

          If sParent.EndsWith(".") Then
            sParent = sParent.Substring(0, sParent.Length - 1)
          End If

        End If
      Else
        sColumnFieldName = psParent
        sParent = ""
      End If

      Dim oColumn As New clsImportColum
      With oColumn
        .Name = sColumnFieldName
        .ColumnID = sColumnData
        .Parent = sParent
        .Formatted = sFormat
        .ColumnName = sColumnHeader
        .VariableName = sVariable
        If String.IsNullOrEmpty(sChild) = False Then
          .ChildNode = sChild
          .Type = JTokenType.Array
        Else
          .Type = JTokenType.String
        End If

      End With

      poListOfColumn.Add(oColumn)


    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try


  End Sub

  Public Function ExtractColumnDetails(psString As String) As List(Of String)

    Try

      Dim oDataSource As List(Of String) = psString.Split(New Char() {" ", ",", ".", ";", ControlChars.Quote, ":", "=", "=="}, StringSplitOptions.RemoveEmptyEntries).ToList

      Dim sColumns As New List(Of String)

      For Each sItem As String In oDataSource
        Dim nStart As Integer = InStr(sItem, "<!")
        Dim nend As Integer = InStr(sItem, "!>")

        If nStart > 0 And nend > 0 Then
          sColumns.Add(sItem.Substring(nStart + 1, nend - nStart - 2))
        End If
      Next

      Return sColumns

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try


  End Function

  Public Function ExtractColumnDetails2(psString As String) As List(Of String)

    Try

      Dim oDataSource As List(Of String) = psString.Split(New Char() {" ", ",", ".", ";", ControlChars.Quote, "=", "=="}, StringSplitOptions.RemoveEmptyEntries).ToList

      Dim sColumns As New List(Of String)

      For Each sItem As String In oDataSource
        Dim nStart As Integer = InStr(sItem, "<!")
        Dim nend As Integer = InStr(sItem, "!>")

        If nStart > 0 And nend > 0 Then
          sColumns.Add(sItem.Substring(nStart + 1, nend - nStart - 2))
        End If
      Next

      Return sColumns

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try


  End Function

  Public Function ExtractColumnDataField(psQuery As String, psColumn As String) As String
    Try

      Dim oDataSource As List(Of String) = psQuery.Split(New Char() {" ", ",", ";", "=", "==", "?", ControlChars.Quote}, StringSplitOptions.RemoveEmptyEntries).ToList

      Dim sColumns As New List(Of String)
      Dim sPreviousItem As String = ""
      For Each sItem As String In oDataSource
        sItem = Replace(sItem, ControlChars.Quote, "")
        If sItem = String.Format("<!{0}!>", psColumn) Then
          Return sPreviousItem
        End If
        sPreviousItem = sItem
      Next


    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try

    Return ""

  End Function

  Public Function ExtractListOfColumnsFromString(psString As String) As List(Of clsImportColum)

    Dim oColumns As New List(Of clsImportColum)

    For Each oItem As String In ExtractColumnDetails(psString)
      If oItem IsNot Nothing Then
        Dim sField As String = ExtractColumnDataField(psString, oItem)
        If sField IsNot Nothing AndAlso String.IsNullOrEmpty(sField) = False Then

          Dim oColProp As clsColumnProperties = ExtractColumnProperties(oItem)

          oColumns.Add(New clsImportColum With {.ColumnID = oItem, .ColumnName = oColProp.CellName, .Name = sField, .Type = JTokenType.String, .Parent = ""})
        End If
      End If
    Next

    Return oColumns

  End Function

  Public Function ExtractRootFromNodes(psString As String) As String

    If String.IsNullOrEmpty(psString) = False AndAlso psString.Contains(".") Then

      Dim sRoot As String = psString.Substring(0, psString.LastIndexOf(".") + 1)

      Return sRoot

    End If

  End Function

  Public Function ValidateData(poImportTemplate As clsDataImportTemplate) As Boolean

    If ValidateHierarchiesIsSelected() = False Then Return False

    'set focus to the excel spreed sheet
    tcgTabs.SelectedTabPage = lcgSpreedsheet
    Application.DoEvents()

    Dim bHasErrors As Boolean = False
    Dim bFormatted As Boolean = False
    Dim bAutoColumnWidth As Boolean = False
    Try


      Dim oTemplate As clsDataImportTemplate = poImportTemplate

      If oTemplate IsNot Nothing Then
        If oTemplate.Validators IsNot Nothing AndAlso oTemplate.Validators.Count > 0 Then
          For Each oValidation In oTemplate.Validators.Where(Function(n) n.Enabled = "1").OrderBy(Function(n) n.Priority).ToList

            If oValidation IsNot Nothing Then

              Dim sTemplateDTO As String = poImportTemplate.DTOObject
              sTemplateDTO = Replace(sTemplateDTO, vbNewLine, "")

              'validate that the specified worksheet exists
              With spreadsheetControl
                If .Document.Worksheets.Contains(poImportTemplate.WorkbookSheetName) = False Then
                  MsgBox(String.Format("Cannot find worksheet {0}", poImportTemplate.WorkbookSheetName))
                  Return False
                Else
                  'if its found make it the active worksheet
                  .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                End If
              End With


              bbiCancel.Enabled = True
              mbCancel = False

              Dim sColumns As List(Of String) = ExtractColumnDetails(oValidation.Query)

              If sColumns IsNot Nothing AndAlso sColumns.Count >= 0 Then


                With spreadsheetControl.ActiveWorksheet


                  Dim nCountRows As Integer = .Rows.LastUsedIndex
                  Dim nCountColumns As Integer = .Columns.LastUsedIndex

                  If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True

                  If bFormatted = False Then
                    .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, 1)).Value = "Validation Code"
                    .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, 1)).Value = "Validation Description"

                    Dim oRangeError As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(.Columns(poImportTemplate.StatusCodeColumn).Index, 1, .Columns(poImportTemplate.StatusCodeColumn).Index, nCountRows)
                    oRangeError.Value = ""
                    oRangeError = .Range.FromLTRB(.Columns(poImportTemplate.StatusDescirptionColumn).Index, 1, .Columns(poImportTemplate.StatusDescirptionColumn).Index, nCountRows)
                    oRangeError.Value = ""

                    If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)

                    Dim oRange As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(0, 0, .Columns.LastUsedIndex, .Rows.LastUsedIndex)

                    'get the data as a excel table
                    If .Tables.Count = 0 Then
                      Dim oTable As Table = .Tables.Add(oRange, True)
                    Else

                      If oRange.RightColumnIndex <> .Columns.LastUsedIndex Or oRange.BottomRowIndex <> .Rows.LastUsedIndex Then

                        Dim oRangeNew As CellRange = .Tables(0).Range
                        oRangeNew = .Range.FromLTRB(oRange.LeftColumnIndex,
                                                   oRange.TopRowIndex,
                                                    .Columns.LastUsedIndex,
                                                   .Rows.LastUsedIndex)

                      End If

                    End If

                    ' Access the workbook's collection of table styles.
                    Dim tableStyles As TableStyleCollection = spreadsheetControl.Document.TableStyles

                    ' Apply the table style to the existing table.
                    .Tables(0).Style = tableStyles("TableStyleMedium7")


                    Application.DoEvents()
                    bFormatted = True
                  End If


                  If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                    spreadsheetControl.BeginUpdate()

                    Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()

                    Select Case oValidation.ValidationType
                      Case "0" 'Exits

                        Dim bhadError As Boolean = Exists(oValidation, spreadsheetControl, sColumns, oTemplate)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True

                      Case "1", "5" 'FindAndReplace

                        Dim bhadError As Boolean = FindAndReplace(oValidation, spreadsheetControl, sColumns)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True

                      Case "2" 'Translate

                        Dim bhadError As Boolean = Translate(oValidation, spreadsheetControl, sColumns)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True


                      Case "3" 'Delete

                        Dim bhadError As Boolean = Deletes(oValidation, spreadsheetControl, sColumns)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True

                      Case "4", "6" 'FindAndReplaceList

                        Dim bhadError As Boolean = FindAndReplaceList(oValidation, spreadsheetControl, sColumns)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True

                      Case "8" ' Array of values

                        Dim bhadError As Boolean = ArryOfValues(oValidation, spreadsheetControl, sColumns)
                        If bhadError = True And bHasErrors = False Then bHasErrors = True

                    End Select

                    spreadsheetControl.EndUpdate()
                    Application.DoEvents()

                  End If
                End With

              End If
            End If

            If mbCancel Then
              Exit For
            End If

          Next
          If bAutoColumnWidth Then spreadsheetControl.ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

        End If

      End If
      Call HideShowColumns(poImportTemplate, mbHideCalulationColumns)

      Call UpdateProgressStatus()
      Return bHasErrors
    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    Finally
      spreadsheetControl.EndUpdate()
      bbiCancel.Enabled = False
      mbCancel = False
    End Try
  End Function

  Private Function Exists(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String), poTemplate As clsDataImportTemplate) As Boolean

    Dim oReturnValuesStore As New Dictionary(Of String, String)
    Dim sQuery As String = ""
    Dim bHasErrors As Boolean = False

    With spreadsheetControl.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex

      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1

        Call UpdateProgressStatus(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

        'get copy of query
        sQuery = poValidation.Query

        'find all colunms that are reference from the excel sheet to do the query
        For Each sColumn As String In psColumns
          Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
          sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
        Next


        'check if query has already been process to void duplicate calls
        If oReturnValuesStore.ContainsKey(sQuery) = False Then

          'call the end point to get the query results
          Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint, poValidation.Headers, sQuery)
          Dim json As JObject = JObject.Parse(oResponse.Content)
          If json IsNot Nothing Then

            Dim sRecordsEffeted As String = json.SelectToken("responsePayload.numberOfElements").ToString
            Dim sMessage As String = String.Format("{0} records found.", sRecordsEffeted)
            Dim sCode As String = IIf(CInt(sRecordsEffeted) = 1, "OK", "Failed")

            'add item to dictionary 
            oReturnValuesStore.Add(sQuery, String.Format("{0}:{1}", sCode, sMessage))
            .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sCode
            .Cells(String.Format("{0}{1}", poTemplate.StatusDescirptionColumn, nRow)).Value = sMessage
          End If
        Else

          'item was already validated, post same results
          Dim sData() As String = oReturnValuesStore.Item(sQuery).ToString.Split(":")
          .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sData(0)
          .Cells(String.Format("{0}{1}", poTemplate.StatusDescirptionColumn, nRow)).Value = sData(1)

        End If


        If mbCancel Then
          Exit For
        End If

      Next

      If oReturnValuesStore.ContainsKey("Failed") And bHasErrors = False Then
        bHasErrors = True
      End If

    End With

    Return bHasErrors

  End Function

  Private Function FindAndReplace(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

    Dim oReturnValuesStore As New Dictionary(Of String, String)
    Dim sQuery As String = ""
    Dim sAPIEndpoint As String = ""
    Dim bHasErrors As Boolean = False

    With spreadsheetControl.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex

      .Cells(String.Format("{0}{1}", poValidation.ReturnCell, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)


      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1
        If nRow.ToString.EndsWith("00") = True Then Application.DoEvents()

        Dim bSourceData As Boolean = True

        Call UpdateProgressStatus(String.Format("Validating {3} Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1, poValidation.Comments))
        Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
        If (oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True) Or mbForceValidate Then

          'get copy of query
          sQuery = poValidation.Query
          sAPIEndpoint = poValidation.APIEndpoint

          'find all colunms that are reference from the excel sheet to do the query
          For Each sColumn As String In psColumns
            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
            If String.IsNullOrEmpty(.Cells(sCellAddress).Value.ToString.Trim) = False Then
              sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
            Else
              bSourceData = False
            End If
          Next

          Dim olist As List(Of String) = ExtractColumnDetails(sAPIEndpoint)
          For Each sColumn As String In olist
            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
            sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
          Next

          'add the Lookup type UUID to the query
          If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
            If poValidation.LookUpType IsNot Nothing Then
              sQuery = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sQuery, ControlChars.Quote)
            End If
          End If

          If bSourceData = True Then
            'check if query has already been process to void duplicate calls
            If oReturnValuesStore.ContainsKey(sQuery) = False Then

              'call the end point to get the query results
              Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet(sAPIEndpoint, poValidation.Headers, sQuery.Trim, poValidation.Sort)
              Dim json As JObject = JObject.Parse(oResponse.Content)
              If json IsNot Nothing Then

                Dim sReturnValue As String = ""

                Try

                  If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then
                    Dim oOjbect As JToken = TryCast(json.SelectToken("responsePayload.content"), JToken)
                    If oOjbect IsNot Nothing Then
                      If oOjbect.Count > 0 Then

                        'Extract the value
                        sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString

                        'check if there is formating
                        If poValidation.Formatted IsNot Nothing Then
                          Select Case poValidation.Formatted.ToString
                            Case "U" : sReturnValue = sReturnValue.ToUpper
                            Case "L" : sReturnValue = sReturnValue.ToLower
                            Case "P" : sReturnValue = StrConv(sReturnValue, VbStrConv.ProperCase)
                          End Select
                        End If
                      End If

                    End If

                  Else

                  End If

                Catch ex As Exception
                  sReturnValue = String.Format("Error: {0}", ex.Message)
                  If bHasErrors = False Then
                    bHasErrors = True
                  End If
                End Try

                'add item to dictionary 
                oReturnValuesStore.Add(sQuery, sReturnValue)

                oCell.Value = sReturnValue.ToString

              End If
            Else

              'item was already validated, post same results
              oCell.Value = oReturnValuesStore.Item(sQuery)
            End If
          End If
        End If


        If mbCancel Then
          Exit For
        End If

      Next

      If String.IsNullOrEmpty(poValidation.ReturnCell) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCell).AutoFit()


    End With

    Return bHasErrors

  End Function

  Private Function Translate(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

    Dim oReturnValuesStore As New Dictionary(Of String, String)
    Dim sQuery As String = ""
    Dim bHasErrors As Boolean = False

    With spreadsheetControl.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex

      .Cells(String.Format("{0}{1}", poValidation.ReturnCell.Trim, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)

      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1

        Call UpdateProgressStatus(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

        Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
        If oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True Then

          'get copy of query
          sQuery = poValidation.Query

          'find all colunms that are reference from the excel sheet to do the query
          For Each sColumn As String In psColumns
            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
            sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
          Next

          'check if query has already been process to void duplicate calls
          If oReturnValuesStore.ContainsKey(sQuery) = False Then

            'call the end point to get the query results
            Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingPost(poValidation.APIEndpoint, sQuery)
            Dim json As JObject = JObject.Parse(oResponse.Content)
            If json IsNot Nothing Then

              Dim sReturnValue As String = ""

              Try
                sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString
                sReturnValue = Replace(sReturnValue, "&#39;", "'")
                sReturnValue = Replace(sReturnValue, "&Amp;", "&")
                sReturnValue = Replace(sReturnValue, "&amp;", "&")

                'check if there is formating
                Select Case poValidation.Formatted.ToString
                  Case "U" : sReturnValue = sReturnValue.ToUpper
                  Case "L" : sReturnValue = sReturnValue.ToLower
                  Case "P" : sReturnValue = StrConv(sReturnValue, VbStrConv.ProperCase)
                End Select

              Catch ex As Exception
                sReturnValue = ""
                If bHasErrors = False Then
                  bHasErrors = True
                End If
              End Try

              'add item to dictionary 
              oReturnValuesStore.Add(sQuery, sReturnValue)

              .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = sReturnValue.ToString

            End If
          Else

            'item was already validated, post same results
            .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = oReturnValuesStore.Item(sQuery)

          End If
        End If


        If mbCancel Then
          Exit For
        End If

      Next

      If String.IsNullOrEmpty(poValidation.ReturnCell) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCell).AutoFit()


    End With

    Return bHasErrors

  End Function

  Private Function Deletes(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

    Dim oReturnValuesStore As New Dictionary(Of String, String)
    Dim sQuery As String = ""
    Dim sAPIEndpoint As String = ""
    Dim bHasErrors As Boolean = False

    With spreadsheetControl.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex

      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1

        Call UpdateProgressStatus(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

        'get copy of query
        sQuery = poValidation.Query
        sAPIEndpoint = poValidation.APIEndpoint

        'find all colunms that are reference from the excel sheet to do the query
        For Each sColumn As String In psColumns
          Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
          sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
        Next

        Dim olist As List(Of String) = ExtractColumnDetails(sAPIEndpoint)
        For Each sColumn As String In olist
          Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
          sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
        Next

        'check if query has already been process to void duplicate calls
        If oReturnValuesStore.ContainsKey(String.Format("{0}-{1}", sAPIEndpoint, sQuery)) = False Then

          'call the end point to get the query results
          Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingDelete(sAPIEndpoint, sQuery)
          Dim json As JObject = JObject.Parse(oResponse.Content)
          'If json IsNot Nothing Then

          '  Dim sRecordsEffeted As String = json.SelectToken("responsePayload.numberOfElements").ToString
          '  Dim sMessage As String = String.Format("{0} records found.", sRecordsEffeted)
          '  Dim sCode As String = IIf(CInt(sRecordsEffeted) = 1, "OK", "Failed")

          '  'add item to dictionary 
          oReturnValuesStore.Add(String.Format("{0}-{1}", sAPIEndpoint, sQuery), String.Format("{0}-{1}", sAPIEndpoint, sQuery))
          .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = json.SelectToken(poValidation.ReturnNodeValue).ToString

          '  .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sCode
          '  .Cells(String.Format("{0}{1}", poTemplate.StatusDescirptionColumn, nRow)).Value = sMessage
          'End If

        End If


        If mbCancel Then
          Exit For
        End If

      Next

      If oReturnValuesStore.ContainsKey("Failed") And bHasErrors = False Then
        bHasErrors = True
      End If

    End With

    Return bHasErrors

  End Function

  Private Function FindAndReplaceList(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

    Dim oReturnValuesStore As New Dictionary(Of String, String)
    Dim sQuery As String = ""
    Dim sAPIEndpoint As String = ""
    Dim bHasErrors As Boolean = False

    With poSpreedSheet.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex


      .Cells(String.Format("{0}{1}", poValidation.ReturnCellWithoutProperties, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments),
                             poValidation.APIEndpoint, String.Format("({0})", poValidation.Comments)), VbStrConv.ProperCase)

      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1

        Call UpdateProgressStatus(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

        Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCellWithoutProperties, nRow))
        If (oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True) Or mbForceValidate Then

          'get copy of query
          sQuery = poValidation.Query
          sAPIEndpoint = poValidation.APIEndpoint


          Dim sColumnListAddress As String = ""
          Dim sColumnOrginalValues As String = ""
          Dim sListValues As New List(Of String)

          Dim sReturnCell As String = ""
          Dim sReturnIndex As String = ""
          If poValidation.ReturnCell.Contains(":") Then
            Dim sValues As String() = poValidation.ReturnCell.Split(":")
            If sValues IsNot Nothing And sValues.Count >= 2 Then
              sReturnCell = sValues(0)
              sReturnIndex = sValues(1)
            End If
          Else
            sReturnCell = poValidation.ReturnCell
          End If

          'find all colunms that are reference from the excel sheet to do the query
          For Each sColumnfromList As String In psColumns

            Dim sColumn As String = ""

            If sColumnfromList.Contains(":") Then
              sColumn = sColumnfromList.Substring(0, sColumnfromList.IndexOf(":"))
            Else
              sColumn = sColumnfromList
            End If

            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
            sColumnOrginalValues = .Cells(sCellAddress).Value.ToString

            'If String.IsNullOrEmpty(sColumnOrginalValues) = False AndAlso sColumnOrginalValues.Length > 0 Then
            '  If sColumnOrginalValues.ToString.Trim.EndsWith(",") = False Then
            '    sColumnOrginalValues = sColumnOrginalValues.ToString.Trim & ","
            '  End If
            'End If

            If sColumnOrginalValues.ToString.Contains(",") Then

              sColumnListAddress = sColumnfromList
              sListValues = .Cells(sCellAddress).Value.ToString.Split(",").ToList



            Else
              sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumnfromList), .Cells(sCellAddress).Value.ToString.Trim)
            End If

          Next

          Dim olist As List(Of String) = ExtractColumnDetails(sAPIEndpoint)
          For Each sColumn As String In olist
            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
            sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
          Next


          Dim oResponseRow As String = ""


          For Each oValue As String In sListValues

            Dim sQueryUpdated As String = sQuery
            If String.IsNullOrEmpty(oValue.ToString.Trim) = False Then

              Dim sNewValue As String = ""
              Dim sSourceCell As String = ""
              Dim sSourceIndex As String = ""
              If sColumnListAddress.Contains(":") Then
                Dim sValues As String() = sColumnListAddress.Split(":")
                If sValues IsNot Nothing And sValues.Count >= 2 Then
                  sSourceCell = sValues(0)
                  sSourceIndex = sValues(1)
                End If
              Else
                sReturnCell = sColumnListAddress
              End If


              If sSourceIndex <> "" Then
                Dim oArray As String() = oValue.Split(":").ToArray
                If oArray IsNot Nothing AndAlso oArray.Count >= CInt(sSourceIndex) Then
                  sNewValue = oArray(CInt(sSourceIndex) - 1)
                End If
              Else
                sNewValue = oValue
              End If


              sQueryUpdated = Replace(sQueryUpdated, String.Format("<!{0}!>", sColumnListAddress), sNewValue.ToString.Trim)

              'add the Lookup type UUID to the query
              If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                If poValidation.LookUpType IsNot Nothing Then
                  sQueryUpdated = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sQueryUpdated, ControlChars.Quote)
                End If
              End If


              'check if query has already been process to void duplicate calls
              If oReturnValuesStore.ContainsKey(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated)) = False Then

                'call the end point to get the query results
                Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet(sAPIEndpoint, poValidation.Headers, sQueryUpdated)
                Dim json As JObject = JObject.Parse(oResponse.Content)
                If json IsNot Nothing Then

                  Dim sReturnValue As String = ""

                  Try
                    sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString

                    'check if there is formating
                    If poValidation.Formatted IsNot Nothing Then
                      Select Case poValidation.Formatted.ToString
                        Case "U" : sReturnValue = sReturnValue.ToUpper
                        Case "L" : sReturnValue = sReturnValue.ToLower
                        Case "P" : sReturnValue = StrConv(sReturnValue, VbStrConv.ProperCase)
                      End Select
                    End If
                  Catch ex As Exception
                    sReturnValue = String.Format("Error:{0}", ex.Message)
                    If bHasErrors = False Then
                      bHasErrors = True
                    End If
                  End Try


                  'add item to dictionary 
                  oReturnValuesStore.Add(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated), sReturnValue)

                  If sReturnIndex <> "" Then
                    sReturnValue = Replace(oValue, sNewValue, sReturnValue)
                  End If


                  If String.IsNullOrEmpty(oResponseRow) Then
                    oResponseRow = sReturnValue
                  Else
                    oResponseRow = String.Format("{0},{1}", oResponseRow, sReturnValue)
                  End If

                End If
              Else

                Dim sValueMemory As String = oReturnValuesStore.Item(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated)).ToString

                If sReturnIndex <> "" Then
                  sValueMemory = Replace(oValue, sNewValue, sValueMemory)
                End If

                If String.IsNullOrEmpty(oResponseRow) Then
                  oResponseRow = sValueMemory
                Else
                  oResponseRow = String.Format("{0},{1}", oResponseRow, sValueMemory)
                End If



              End If

            Else


            End If

          Next

          .Cells(String.Format("{0}{1}", poValidation.ReturnCellWithoutProperties, nRow)).Value = oResponseRow


        End If

        If mbCancel Then
          Exit For
        End If

      Next

      If String.IsNullOrEmpty(poValidation.ReturnCellWithoutProperties) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCellWithoutProperties.ToString).AutoFit()

    End With

    Return bHasErrors

  End Function

  Private Function ArryOfValues(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

    Dim bHasErrors As Boolean = False

    With poSpreedSheet.ActiveWorksheet

      Dim nCountColumns As Integer = .Columns.LastUsedIndex
      Dim nCountRows As Integer = .Rows.LastUsedIndex


      .Cells(String.Format("{0}{1}", poValidation.ReturnCell, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)

      'loop though all rows in the excel sheet
      For nRow As Integer = 2 To nCountRows + 1

        Call UpdateProgressStatus(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

        Dim oCellDestination As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
        If (oCellDestination.Value.IsEmpty Or String.IsNullOrEmpty(oCellDestination.Value.ToString) = True) Or mbForceValidate Then

          Dim sSourceColumnName As String = ""
          Dim olist As List(Of String) = ExtractColumnDetails(poValidation.Query)
          If olist IsNot Nothing And olist.Count > 0 Then

            sSourceColumnName = olist(0).ToString
            Dim oCellSource As Cell = .Cells(String.Format("{0}{1}", sSourceColumnName, nRow))

            If String.IsNullOrEmpty(oCellSource.Value.ToString) = False And oCellSource.Value.ToString.Length > 0 Then
              If oCellSource.Value.ToString.Trim.EndsWith(",") = False Then
                oCellSource.Value = oCellSource.Value.ToString.Trim & ","
              End If
            End If


            Dim sColumnOrginalValues As String = oCellDestination.Value.ToString
            Dim sListValues As New List(Of String)
            Dim sReturnValue As String = ""

            'find all colunms that are reference from the excel sheet to do the query
            If oCellSource.Value.ToString.Contains(",") Then
              sListValues = oCellSource.Value.ToString.Split(",").ToList
            End If

            If sListValues IsNot Nothing Then
              For Each sValue In sListValues
                If sValue.ToString.Trim.Length > 0 Then
                  sReturnValue = String.Format("{0}{1}{2}{1},", sReturnValue, ControlChars.Quote, sValue)
                End If
              Next
            End If

            If sReturnValue.EndsWith(",") Then
              sReturnValue = sReturnValue.Substring(0, sReturnValue.Length - 1)
            End If

            oCellDestination.Value = sReturnValue

          End If

        End If

        If mbCancel Then
          Exit For
        End If


      Next

      If String.IsNullOrEmpty(poValidation.ReturnCell) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCell).AutoFit()

    End With

    Return bHasErrors

  End Function

  Public Sub ImportData(poImportTemplate As clsDataImportTemplate)
    Try

      If ValidateHierarchiesIsSelected() = False Then Exit Sub

      Dim mbIgnore As Boolean = False
      Dim bAutoColumnWidth As Boolean = False

      'set focus to the excel spreed sheet
      tcgTabs.SelectedTabPage = lcgSpreedsheet
      Application.DoEvents()

      Dim sTemplateDTO As String = poImportTemplate.DTOObject
      sTemplateDTO = Replace(sTemplateDTO, vbNewLine, "")

      'validate that the specified worksheet exists
      With spreadsheetControl
        If .Document.Worksheets.Contains(poImportTemplate.WorkbookSheetName) = False Then
          MsgBox(String.Format("Cannot find worksheet {0}", poImportTemplate.WorkbookSheetName))
          Exit Sub
        Else
          'if its found make it the active worksheet
          .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
        End If
      End With

      If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

        Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()

        With spreadsheetControl.ActiveWorksheet
          Dim nCountRows As Integer = .Rows.LastUsedIndex
          Dim nCountColumns As Integer = .Columns.LastUsedIndex
          If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True

          spreadsheetControl.BeginUpdate()

          .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, 1)).Value = "Result Code"
          .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, 1)).Value = "Result Description"

          If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)

          Dim oRangeError As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(.Columns(poImportTemplate.StatusCodeColumn).Index, 1, .Columns(poImportTemplate.StatusCodeColumn).Index, nCountRows)
          oRangeError.Value = ""
          oRangeError = .Range.FromLTRB(.Columns(poImportTemplate.StatusDescirptionColumn).Index, 1, .Columns(poImportTemplate.StatusDescirptionColumn).Index, nCountRows)
          oRangeError.Value = ""


          'get the data as a excel table
          If .Tables.Count = 0 Then

            Dim oRange As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(0, 0, .Columns.LastUsedIndex, .Rows.LastUsedIndex)
            Dim oTable As Table = .Tables.Add(oRange, True)

          End If

          ' Access the workbook's collection of table styles.
          Dim tableStyles As TableStyleCollection = spreadsheetControl.Document.TableStyles

          ' Access the built-in table style from the collection by its name.
          Dim tableStyle As TableStyle = tableStyles("TableStyleMedium2")

          ' Apply the table style to the existing table.
          .Tables(0).Style = tableStyle

          Application.DoEvents()

          bbiCancel.Enabled = True
          mbCancel = False

          For nRow = 2 To nCountRows + 1

            'Call UpdateProgressStatus(String.Format("Importing row {0} of {1}", nRow, nCountRows + 2))
            If .Rows.Item(nRow).Visible = True Then

              Dim oDTO As String = sTemplateDTO

              Dim jsnDTO As JObject = JObject.Parse(sTemplateDTO)
              For Each oProperty As JProperty In jsnDTO.Children
                CheckChildrenNodes(oProperty, nRow)
              Next

              oDTO = jsnDTO.ToString

              Dim sQuery As String = poImportTemplate.UpdateQuery.ToString

              Dim sColumns As List(Of String) = ExtractColumnDetails(sQuery)


              'find all colunms that are reference from the excel sheet to do the query
              For Each sColumn As String In sColumns
                Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
              Next


              If String.IsNullOrEmpty(poImportTemplate.ReturnCellDTO) = False Then
                .Cells(String.Format("{0}{1}", poImportTemplate.ReturnCellDTO, nRow)).Value = oDTO.ToString
              End If

              Dim oResponse As IRestResponse

              Select Case poImportTemplate.ImportType
                Case "1"
                  oResponse = goHTTPServer.CallWebEndpointUsingPost(poImportTemplate.APIEndpoint, oDTO)
                Case "2"
                  If String.IsNullOrEmpty(.Cells(String.Format("{0}{1}", poImportTemplate.ReturnCell, nRow)).Value.ToString) Then
                    'add
                    oResponse = goHTTPServer.CallWebEndpointUsingPost(poImportTemplate.APIEndpoint, oDTO)
                  Else
                    'edit
                    oResponse = goHTTPServer.CallWebEndpointUsingPut(poImportTemplate.APIEndpoint,
                                                                   .Cells(String.Format("{0}{1}", poImportTemplate.ReturnCell, nRow)).Value.ToString,
                                                                    sQuery, oDTO)
                  End If
                Case Else
                  Exit Sub
              End Select

              Call UpdateProgressStatus(String.Format("Importing row {0} of {1} with {2}", nRow, nCountRows + 2, oResponse.StatusCode))


              Select Case oResponse.StatusCode

                Case HttpStatusCode.Created, HttpStatusCode.OK

                  Dim json As JObject = JObject.Parse(oResponse.Content)
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, nRow)).Value = json.SelectToken("message").ToString

                  If String.IsNullOrEmpty(poImportTemplate.ReturnNodeValue.ToString) = False Then
                    ' If json.ContainsKey(poImportTemplate.ReturnNodeValue) Then
                    .Cells(String.Format("{0}{1}", poImportTemplate.ReturnCell, nRow)).Value = json.SelectToken(poImportTemplate.ReturnNodeValue).ToString
                    'End If
                  End If

                Case HttpStatusCode.BadRequest

                  Dim json As JObject = JObject.Parse(oResponse.Content)
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, nRow)).Value = json.SelectToken("message").ToString

                  If mbIgnore = False Then
                    'located the beging of the stack trace
                    Dim sMessage As String = oResponse.Content.ToString.Substring(0, oResponse.Content.ToString.IndexOf("stackTrace"))

                    Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, sMessage), MsgBoxStyle.AbortRetryIgnore)
                      Case MsgBoxResult.Abort
                        Clipboard.SetText(oDTO)
                        Exit Sub
                      Case MsgBoxResult.Retry
                      Case MsgBoxResult.Ignore
                        mbIgnore = True
                    End Select
                  End If
                Case HttpStatusCode.NotFound
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, nRow)).Value = "Record not Found"


                Case Else
                  Dim json As JObject = JObject.Parse(oResponse.Content)
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                  .Cells(String.Format("{0}{1}", poImportTemplate.StatusDescirptionColumn, nRow)).Value = json.SelectToken("message").ToString

              End Select

            End If

            If mbCancel Then
              Exit For
            End If

          Next
          If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)
        End With
      End If

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    Finally

      Call HideShowColumns(poImportTemplate, mbHideCalulationColumns)
      spreadsheetControl.EndUpdate()
      Call UpdateProgressStatus("Import Completed.")
      bbiCancel.Enabled = False
      mbCancel = False
    End Try


  End Sub

  Public Sub CheckChildrenNodes(poNode As JProperty, pnRow As Integer)

    Try


      Select Case poNode.Value.Type
        Case JTokenType.Object

          For Each oSubProperty As JProperty In poNode.Value
            CheckChildrenNodes(oSubProperty, pnRow)
          Next

        Case JTokenType.String
          Try


            Dim sColumn As List(Of String) = ExtractColumnDetails2(poNode.Value.ToString)
            If sColumn IsNot Nothing AndAlso sColumn.Count > 0 Then
              Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumn(0))

              If oColumnPropeties IsNot Nothing Then

                Dim sValue As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim

                If oColumnPropeties.Format <> "" Then
                  sValue = FormatString(sValue, oColumnPropeties.Format)
                End If

                If sValue = "" Then
                  poNode.Value = Nothing
                Else
                  poNode.Value = sValue
                End If

              End If
            End If
          Catch ex As Exception
            MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))

          End Try
        Case JTokenType.Array

          For Each oChild In poNode.Value.Children
            If oChild IsNot Nothing Then

              Select Case oChild.Type

                Case JTokenType.Property
                Case JTokenType.String
                  Try
                    Dim sColumn As List(Of String) = ExtractColumnDetails(poNode.Value.ToString)
                    Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumn(0))

                    Dim sNewValue As String = ""
                    If oColumnPropeties IsNot Nothing Then

                      Dim sValue As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim
                      Dim jArray As New JArray
                      If sValue IsNot Nothing AndAlso String.IsNullOrEmpty(sValue) = False Then
                        Dim sArray As List(Of String) = sValue.Split(",").ToList
                        For Each sObject As String In sArray
                          If sObject.ToString.Trim.Length > 0 Then
                            jArray.Add(sObject)
                          End If
                        Next

                        poNode.Value = jArray
                      Else
                        poNode.Value = Nothing
                      End If
                    End If
                  Catch ex As Exception
                    MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))

                  End Try
                Case JTokenType.Array

                Case JTokenType.Object
                  Try
                    Dim oToken As JToken = poNode.First.DeepClone

                    Dim sColumns As List(Of String) = ExtractColumnDetails(oToken.ToString)
                    If sColumns.Count > 0 Then
                      Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumns(0))
                      Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim

                      Dim sValues As New List(Of String)
                      If sCellData IsNot Nothing And String.IsNullOrEmpty(sCellData) = False Then
                        sValues = sCellData.Split(",").ToList

                        For Each oVal As String In sValues
                          If oVal IsNot Nothing And String.IsNullOrEmpty(oVal) = False Then
                            Dim oNewToken As JToken = oToken.DeepClone
                            Dim oArray As New List(Of String)
                            If oVal.Contains(":") Then
                              oArray = oVal.Split(":").ToList
                            End If

                            For Each oProp As JProperty In oNewToken.Values
                              Dim sCol As List(Of String) = ExtractColumnDetails(oProp)
                              Dim oCP As clsColumnProperties = ExtractColumnProperties(sCol(0))

                              If oCP.IndexID <> "" Then
                                If oArray.Count >= CInt(oCP.IndexID) Then

                                  If oColumnPropeties.Format <> "" Then
                                    oArray(CInt(oCP.IndexID) - 1) = FormatString(oArray(CInt(oCP.IndexID) - 1), oColumnPropeties.Format)
                                  End If

                                  oProp.Value = oArray(CInt(oCP.IndexID) - 1).ToString
                                End If
                              Else
                                oProp.Value = oVal
                              End If
                            Next

                            poNode.Values.First.AddAfterSelf(oNewToken.First)

                          End If
                        Next

                        'now remove the first value
                        If poNode.Values.Count > 0 Then
                          poNode.Values.First.Remove()
                        End If

                        Exit For
                      Else

                        If poNode.Values.Count > 0 Then
                          poNode.Values.First.Remove()
                        End If

                      End If
                    End If
                  Catch ex As Exception
                    MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))

                  End Try
              End Select
            End If
          Next


      End Select

    Catch ex As Exception
      '     MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    End Try

  End Sub

  Public Sub UpdateProgressStatus(Optional psStatus As String = "")

    siStatus.Caption = psStatus
    siStatus.Refresh()
    Application.DoEvents()

  End Sub

  Public Function SaveWorkbook(Optional pbSaveAs As Boolean = False, Optional psPath As String = "", Optional psFolder As String = "",
                          Optional psFileName As String = "", Optional psExtention As String = "", Optional pbIncreaseCounter As Boolean = True) As Boolean

    Dim sTempPath As String = Path.GetTempFileName
    goOpenWorkBook.SelectedHierarchy = gsSelectedHierarchy

    Try

      Dim sPath As String = IIf(String.IsNullOrEmpty(psPath), msFilePath, psPath)
      Dim sFileName As String = IIf(String.IsNullOrEmpty(psFileName), msFileName, psFileName)

      If String.IsNullOrEmpty(psFolder) = False Then
        sPath = Path.Combine(sPath, psFolder)
      End If


      If String.IsNullOrEmpty(sFileName) Or String.IsNullOrEmpty(sPath) Or pbSaveAs Then
        Using fd As SaveFileDialog = New SaveFileDialog

          fd.Title = "Save a Otalio Dynamic import template."
          fd.Filter = "Dynamic import Template |*.ditw"
          fd.FilterIndex = 1
          fd.FileName = txtWorkbookName.Text
          fd.RestoreDirectory = True
          If fd.ShowDialog() = DialogResult.OK Then

            sPath = Path.GetDirectoryName(fd.FileName)
            sFileName = Path.GetFileName(fd.FileName)

          Else

            Return False

          End If
        End Using
      End If


      'create a temporary file
      SaveFile(sTempPath, goOpenWorkBook)

      If Directory.Exists(sPath) = False Then
        Directory.CreateDirectory(sPath)
      End If


      If String.IsNullOrEmpty(psExtention) = False Then
        If Path.GetExtension(sFileName) = "" Then
          sFileName = Path.Combine(sFileName, psExtention)
        Else
          sFileName = Replace(sFileName, Path.GetExtension(sFileName), psExtention)
        End If
      End If


      If File.Exists(Path.Combine(sPath, sFileName)) Then
        'check if file are different
        If pbIncreaseCounter Then
          If CompareFiles(sTempPath, Path.Combine(sPath, sFileName)) = False Then
            If pbIncreaseCounter Then goOpenWorkBook.SaveVersion += 1
            labVersion.Text = goOpenWorkBook.WorkbookVersion
          End If
        End If
      End If

      If File.Exists(Path.Combine(sPath, sFileName)) Then
        File.Delete(Path.Combine(sPath, sFileName))
      End If

      SaveFile(Path.Combine(sPath, sFileName), goOpenWorkBook)


      If pbSaveAs Then
        'copy also the excel file
        sFileName = Path.GetFileNameWithoutExtension(sFileName)

        Dim sCurrentFilename As String = spreadsheetControl.Document.Path
        If String.IsNullOrEmpty(sCurrentFilename) = True Then
          sCurrentFilename = String.Format("{0}\{1}.{2}", sPath, sFileName, "xlsx")
        End If

        Dim sExcelExtention As String = Path.GetExtension(sCurrentFilename)
        spreadsheetControl.SaveDocument(String.Format("{0}\{1}{2}", sPath, sFileName, sExcelExtention))

      End If

    Catch ex As Exception

    Finally
      File.Delete(sTempPath)
    End Try

  End Function

  Public Sub loadWorkbook()

    If goOpenWorkBook IsNot Nothing Then

      lbxImportTemplates.DataSource = goOpenWorkBook.Templates.OrderBy(Function(x) x.Priority).ToList
      lbxImportTemplates.Refresh()
      labVersion.Text = goOpenWorkBook.WorkbookVersion
      txtWorkbookName.EditValue = goOpenWorkBook.WorkbookName

      With txtMajor
        .DataBindings.Clear()
        .DataBindings.Add(New Binding("EditValue", goOpenWorkBook, "MajorVersion"))
      End With

      With txtMinor
        .DataBindings.Clear()
        .DataBindings.Add(New Binding("EditValue", goOpenWorkBook, "MinorVersion"))
      End With

      With txtBuild
        .DataBindings.Clear()
        .DataBindings.Add(New Binding("EditValue", goOpenWorkBook, "SaveVersion"))
      End With

      lueHierarchies.Properties.DataSource = goHierarchies
      gsSelectedHierarchy = goOpenWorkBook.SelectedHierarchy
      lueHierarchies.EditValue = gsSelectedHierarchy

    End If

  End Sub

  Public Sub QueryandLoad(poImportTemplate As clsDataImportTemplate)


    If ValidateHierarchiesIsSelected() = False Then Exit Sub

    'set focus to the excel spreed sheet
    tcgTabs.SelectedTabPage = lcgSpreedsheet
    bbiQuery.Enabled = False

    Application.DoEvents()

    Dim bHasErrors As Boolean = False
    Dim bFormatted As Boolean = False
    Dim bAutoColumnWidth As Boolean = False
    Dim oListUsedColumns As New List(Of String)
    Dim bIgnore As Boolean = False
    Try

      '
      Dim oTemplate As clsDataImportTemplate = TryCast(poImportTemplate, clsDataImportTemplate)
      If oTemplate IsNot Nothing Then



        If oTemplate.ImportColumns IsNot Nothing AndAlso oTemplate.ImportColumns.Count >= 0 Then

          With spreadsheetControl

            If .Document.Worksheets.Contains(oTemplate.WorkbookSheetName) = False Then
              createNewExcelSheet(oTemplate)
              .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(oTemplate.WorkbookSheetName)
            Else
              'if its found make it the active worksheet
              .Document.Worksheets.Remove(.Document.Worksheets(poImportTemplate.WorkbookSheetName))
              createNewExcelSheet(poImportTemplate)
              .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(oTemplate.WorkbookSheetName)
            End If

            Call UpdateProgressStatus(String.Format("Requesting Data from Server {0} on API {1} ", goConnection._ServerAddress, oTemplate.APIEndpoint))


            If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

              Dim oResponse As IRestResponse
              Dim sNodeLoad As String = ""

              If oTemplate.GraphQLQuery IsNot Nothing AndAlso String.IsNullOrEmpty(oTemplate.GraphQLQuery.ToString) = False Then
                oResponse = goHTTPServer.CallGraphQL("apollo-server", oTemplate.GraphQLQuery.ToString)
                sNodeLoad = String.Format("data.{0}", oTemplate.GraphQLRootNode)
                If sNodeLoad.EndsWith(".") Then
                  sNodeLoad = sNodeLoad.Substring(0, sNodeLoad.Length - 1)
                End If
              Else
                oResponse = goHTTPServer.CallWebEndpointUsingGet(oTemplate.APIEndpoint, "", oTemplate.SelectQuery.ToString)
                sNodeLoad = "responsePayload.content"
              End If

              Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
              If jsServerResponse IsNot Nothing Then

                If oResponse.StatusCode = HttpStatusCode.OK Then
                  Dim oObject As JToken = TryCast(jsServerResponse.SelectToken(sNodeLoad), JToken)

                  If oObject IsNot Nothing Then
                    If oObject.Count > 0 Then

                      bbiCancel.Enabled = True
                      mbCancel = False


                      .BeginUpdate()

                      If .ActiveWorksheet.Rows.LastUsedIndex = 0 Then
                        createNewExcelSheet(oTemplate)
                      Else

                        If .ActiveWorksheet.Tables.Count > 0 Then
                          .ActiveWorksheet.Tables.Clear()
                        End If

                        If .ActiveWorksheet.Rows.LastUsedIndex > 0 Then
                          .ActiveWorksheet.Rows.Remove(2, .ActiveWorksheet.Rows.LastUsedIndex)
                        End If

                        createNewExcelSheet(oTemplate)

                      End If


                      Dim nCountRows As Integer = .ActiveWorksheet.Rows.LastUsedIndex
                      Dim nCountColumns As Integer = .ActiveWorksheet.Columns.LastUsedIndex

                      Dim nRow As Integer = 1
                      Dim nCounter As Integer = 0

                      For Each oRow In oObject

                        Try

                          If oRow IsNot Nothing Then
                            nRow += 1
                            nCounter += 1

                            Call UpdateProgressStatus(String.Format("Downloading object {0} of {1} objects ", nCounter, oObject.Count))

                            For Each ocolumn In oTemplate.ImportColumns.OrderBy(Function(n) n.No).ToList

                              Try


                                If mbCancel = True Then Exit For

                                Dim sValue As String = ""
                                Dim sColumnName As String = ""

                                Dim oCell As Cell = .ActiveWorksheet.Cells(String.Format("{0}{1}", ocolumn.ColumnName, nRow))

                                'check if this is a multi node array
                                Dim sNodeName As String = ""
                                If String.IsNullOrEmpty(ocolumn.Parent) = False Then
                                  sNodeName = ocolumn.Parent & "." & ocolumn.Name
                                Else
                                  sNodeName = ocolumn.Name
                                End If

                                Dim oJsn As JToken = oRow
                                Dim oNodes As List(Of String) = ocolumn.Parent.Split(".").ToList

                                Select Case ocolumn.Type

                                  Case JTokenType.Object

                                    oJsn = oRow.SelectToken(sNodeName)
                                    If oJsn IsNot Nothing Then
                                      oCell.Value = oJsn.ToString
                                      oCell.NumberFormat = "Text"


                                      With oListUsedColumns
                                        If .Contains(ocolumn.ColumnName) = False Then
                                          .Add(ocolumn.ColumnName)
                                        End If
                                      End With

                                    End If

                                  Case JTokenType.String

                                    oJsn = oRow.SelectToken(sNodeName)
                                    If oJsn IsNot Nothing Then
                                      oCell.Value = oJsn.ToString
                                      oCell.NumberFormat = "Text"


                                      With oListUsedColumns
                                        If .Contains(ocolumn.ColumnName) = False Then
                                          .Add(ocolumn.ColumnName)
                                        End If
                                      End With
                                    End If

                                  Case JTokenType.Array

                                    Dim sValues As String = ""
                                    oJsn = oRow.SelectToken(IIf(ocolumn.Parent = "", ocolumn.Name, ocolumn.Parent))
                                    If oJsn IsNot Nothing Then
                                      For Each oChildren In oJsn.Children
                                        If oChildren IsNot Nothing Then
                                          Select Case oChildren.Type
                                            Case JTokenType.String

                                              sValues = sValues & String.Format("{0},", TryCast(oChildren, JValue).Value)

                                            Case JTokenType.Object

                                              Dim sUsedProperties As New List(Of String)
                                              If ocolumn.ChildNode <> "" Then
                                                Dim oCols As List(Of clsImportColum) = oTemplate.ImportColumns.Where(Function(n) n.ColumnName = ocolumn.ColumnName).ToList
                                                If oCols IsNot Nothing AndAlso oCols.Count > 0 Then
                                                  For Each oC As clsImportColum In oCols
                                                    sUsedProperties.Add(oC.Name)
                                                  Next
                                                End If
                                              Else
                                                sUsedProperties.Add(ocolumn.Name)
                                              End If

                                              For Each oProperty As JProperty In oChildren
                                                If oProperty IsNot Nothing Then
                                                  If sUsedProperties.Contains(oProperty.Name) Then
                                                    sValues = sValues & String.Format("{1}:", oProperty.Name, oProperty.Value)
                                                  End If
                                                End If
                                              Next

                                              If sValues.EndsWith(":") Then
                                                sValues = sValues.Substring(0, sValues.Length - 1)
                                              End If

                                              sValues = sValues & ","

                                          End Select
                                        End If


                                      Next



                                      oCell.Value = sValues
                                      oCell.NumberFormat = "Text"


                                      With oListUsedColumns
                                        If .Contains(ocolumn.ColumnName) = False Then
                                          .Add(ocolumn.ColumnName)
                                        End If
                                      End With

                                    End If

                                End Select




                                'format cell data to clean up some issues
                                If oCell.Value.ToString.Contains("[]") Then oCell.Value = Replace(oCell.Value.ToString, "[]", "").ToString


                              Catch ex As Exception
                                If bIgnore = False Then
                                  Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, ex.Message), MsgBoxStyle.AbortRetryIgnore)
                                    Case MsgBoxResult.Abort
                                      Clipboard.SetText(ex.ToString)
                                      Exit Sub
                                    Case MsgBoxResult.Retry
                                    Case MsgBoxResult.Ignore
                                      bIgnore = True
                                  End Select
                                End If
                              End Try
                            Next

                            'add the last ID
                            If oTemplate.ReturnNodeValue IsNot Nothing Then
                              If jsServerResponse.ContainsKey(oTemplate.ReturnNodeValue) Then
                                Dim sNode As String = oTemplate.ReturnNodeValue.Substring(oTemplate.ReturnNodeValue.LastIndexOf(".") + 1)
                                If String.IsNullOrEmpty(sNode) = False Then
                                  Dim oCellID As Cell = .ActiveWorksheet.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow))
                                  oCellID.Value = oRow.SelectToken(sNode.ToString).ToString

                                  If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, 1)).Value.ToString) Then
                                    .ActiveWorksheet.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, 1)).Value = sNode
                                  End If
                                End If
                              End If
                            End If
                          End If

                        Catch ex As Exception
                          MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
                        End Try

                      Next



                      Call UpdateProgressStatus(String.Format("Formating document"))


                      Dim oRange As DevExpress.Spreadsheet.CellRange = .ActiveWorksheet.Range.FromLTRB(0, 0, .ActiveWorksheet.Columns.LastUsedIndex, .ActiveWorksheet.Rows.LastUsedIndex)

                      'get the data as a excel table
                      If .ActiveWorksheet.Tables.Count = 0 Then
                        Dim oTable As Table = .ActiveWorksheet.Tables.Add(oRange, True)
                      Else

                        If oRange.RightColumnIndex <> .ActiveWorksheet.Columns.LastUsedIndex Or oRange.BottomRowIndex <> .ActiveWorksheet.Rows.LastUsedIndex Then

                          Dim oRangeNew As CellRange = .ActiveWorksheet.Tables(0).Range
                          oRangeNew = .ActiveWorksheet.Range.FromLTRB(oRange.LeftColumnIndex,
                                               oRange.TopRowIndex,
                                                .ActiveWorksheet.Columns.LastUsedIndex,
                                               .ActiveWorksheet.Rows.LastUsedIndex)

                        End If

                      End If


                      oResponse = Nothing
                      jsServerResponse = Nothing
                      oObject = Nothing

                      .EndUpdate()
                      Application.DoEvents()
                      .BeginUpdate()

                      Dim oListReferencedColumns As New Dictionary(Of String, clsValidation)
                      For Each oValidation As clsValidation In oTemplate.Validators.Where(Function(n) n.Enabled = "1" Or n.Visibility = "1").ToList
                        If oValidation.Query IsNot Nothing AndAlso String.IsNullOrEmpty(oValidation.Query) = False Then
                          Dim olist As List(Of String) = ExtractColumnDetails(oValidation.Query)
                          If olist IsNot Nothing And olist.Count > 0 Then
                            For Each sItem As String In olist
                              'check if it in DTO
                              If oListUsedColumns.Contains(sItem) = False Then
                                If oListReferencedColumns.ContainsKey(sItem) = False Then
                                  oListReferencedColumns.Add(sItem, oValidation)
                                End If
                              End If
                            Next
                          End If
                        End If
                      Next

                      If oListReferencedColumns IsNot Nothing And oListReferencedColumns.Count > 0 Then
                        For Each sKey As String In oListReferencedColumns.Keys
                          Dim oValidation As clsValidation = TryCast(oListReferencedColumns.Item(sKey), clsValidation)

                          Call UpdateProgressStatus(String.Format("Validating Key {0}", sKey))


                          If oValidation IsNot Nothing Then

                            Select Case oValidation.ValidationType

                              Case "1", "5" 'FindAndReplace

                                Try

                                  'If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value.ToString) = True Then
                                  .ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value = String.Format("{0}({1})", oValidation.Comments, oValidation.Priority)
                                  'End If

                                  Dim sQuery As String = oValidation.Query
                                  Dim sDataSource As String = ExtractColumnDataField(sQuery, sKey)
                                  Dim sRootNode As String = ExtractRootFromNodes(oValidation.ReturnNodeValue)
                                  Dim oListValues As New Dictionary(Of String, String)

                                  sQuery = Replace(sQuery, sDataSource, Replace(oValidation.ReturnNodeValue, sRootNode, ""))

                                  Dim oReturnNode As String = String.Format("{1}{0}", sDataSource, sRootNode)

                                  For nRow = 2 To .ActiveWorksheet.Rows.LastUsedIndex + 1
                                    'only extract if the destination cell are empty
                                    If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, nRow)).Value.ToString) = True Then
                                      If mbCancel = True Then Exit Sub
                                      Call UpdateProgressStatus(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

                                      If .ActiveWorksheet.Cells(String.Format("{0}{1}", oValidation.ReturnCell, nRow)).Value IsNot Nothing AndAlso
                                     String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", oValidation.ReturnCell, nRow)).Value.ToString) = False Then

                                        Dim sNewQuery As String = sQuery
                                        Dim sColumns As List(Of String) = ExtractColumnDetails(oValidation.Query)

                                        For Each sColumn As String In sColumns
                                          If sColumn <> sKey Then
                                            Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                            If String.IsNullOrEmpty(.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim) = False Then
                                              sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                            End If
                                          Else
                                            Dim sCellAddress As String = String.Format("{0}{1}", oValidation.ReturnCell, nRow)
                                            sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sKey), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                          End If
                                        Next

                                        Dim sAPIendpoint As String = oValidation.APIEndpoint
                                        Dim olist As List(Of String) = ExtractColumnDetails(oValidation.APIEndpoint)
                                        For Each sColumn As String In olist
                                          Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                          sAPIendpoint = Replace(sAPIendpoint, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                        Next


                                        Dim sValue As String = ""
                                        If oListValues.ContainsKey(sNewQuery) Then
                                          sValue = oListValues(sNewQuery)
                                        Else
                                          sValue = goHTTPServer.GetValueFromEndpoint(sAPIendpoint, sNewQuery, oReturnNode, sRootNode)
                                          If String.IsNullOrEmpty(sValue) = False Then
                                            oListValues.Add(sNewQuery, sValue)
                                          End If
                                        End If

                                        .ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, nRow)).Value = sValue

                                      End If
                                    End If
                                  Next nRow

                                Catch ex As Exception
                                  If bIgnore = False Then
                                    Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, ex.Message), MsgBoxStyle.AbortRetryIgnore)
                                      Case MsgBoxResult.Abort
                                        Clipboard.SetText(ex.ToString)
                                        Exit For
                                      Case MsgBoxResult.Retry
                                      Case MsgBoxResult.Ignore
                                        bIgnore = True
                                    End Select
                                  End If
                                End Try


                              Case "4", "6" 'FindAndReplaceList
                                Try

                                  Dim sColumnCode As String = ""
                                  Dim sIndex As String = ""
                                  If sKey.Contains(":") Then
                                    Dim sValues As String() = sKey.Split(":")
                                    If sValues IsNot Nothing And sValues.Count >= 2 Then
                                      sColumnCode = sValues(0)
                                      sIndex = sValues(1)
                                    End If
                                  Else
                                    sColumnCode = sKey
                                  End If

                                  Dim sReturnCell As String = ""
                                  Dim sReturnIndex As String = ""
                                  If oValidation.ReturnCell.Contains(":") Then
                                    Dim sValues As String() = oValidation.ReturnCell.Split(":")
                                    If sValues IsNot Nothing And sValues.Count >= 2 Then
                                      sReturnCell = sValues(0)
                                      sReturnIndex = sValues(1)
                                    End If
                                  Else
                                    sReturnCell = oValidation.ReturnCell
                                  End If

                                  ' If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value.ToString) = True Then
                                  .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, 1)).Value = oValidation.Comments
                                  'End If

                                  Dim sQuery As String = oValidation.Query
                                  Dim sDataSource As String = ExtractColumnDataField(sQuery, sKey)
                                  Dim sRootNode As String = ExtractRootFromNodes(oValidation.ReturnNodeValue)
                                  Dim oListValues As New Dictionary(Of String, String)

                                  sQuery = Replace(sQuery, sDataSource, Replace(oValidation.ReturnNodeValue, sRootNode, ""))

                                  Dim oReturnNode As String = String.Format("{1}{0}", sDataSource, sRootNode)

                                  For nRow = 2 To .ActiveWorksheet.Rows.LastUsedIndex + 1
                                    Call UpdateProgressStatus(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

                                    'only extract if the destination cell are empty
                                    If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value.ToString) = True Then

                                      If mbCancel = True Then Exit Sub

                                      If .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value IsNot Nothing AndAlso
                                     String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.ToString) = False Then

                                        Dim sArray As String = ""

                                        'check if destination cell have alreay been populated with values
                                        If .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value.IsEmpty = False Then
                                          sArray = .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value.ToString
                                        Else
                                          sArray = .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.ToString
                                        End If

                                        If sArray IsNot Nothing AndAlso String.IsNullOrEmpty(sArray) = False Then
                                          Dim oListofKeys As List(Of String) = sArray.Split(",").ToList
                                          Dim oListofValues As New List(Of String)
                                          Dim oListofQueries As New List(Of String)

                                          If oListofKeys IsNot Nothing And oListofKeys.Count > 0 Then
                                            For Each sValueIDs As String In oListofKeys
                                              If String.IsNullOrEmpty(sValueIDs) = False Then
                                                Dim sValueID As String = ""

                                                If sReturnIndex <> "" Then
                                                  Dim oVal As String() = sValueIDs.Split(":").ToArray
                                                  If oVal IsNot Nothing And oVal.Count >= CInt(sReturnIndex) Then
                                                    sValueID = oVal(CInt(sReturnIndex) - 1)
                                                  End If
                                                Else
                                                  sValueID = sValueIDs
                                                End If

                                                Dim sNewQuery As String = sQuery
                                                Dim sColumns As List(Of String) = ExtractColumnDetails(sNewQuery)

                                                For Each sColumn As String In sColumns
                                                  If sColumn <> sKey Then
                                                    Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                                    If String.IsNullOrEmpty(.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim) = False Then
                                                      sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                    End If
                                                  Else
                                                    Dim sCellAddress As String = String.Format("{0}{1}", sReturnCell, nRow)
                                                    sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sKey), sValueID)
                                                  End If

                                                Next

                                                Dim sAPIendpoint As String = oValidation.APIEndpoint
                                                Dim olist As List(Of String) = ExtractColumnDetails(oValidation.APIEndpoint)
                                                For Each sColumn As String In olist
                                                  Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                                  sAPIendpoint = Replace(sAPIendpoint, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                Next

                                                Dim sValue As String = ""

                                                If oListValues.ContainsKey(sNewQuery) Then
                                                  sValue = oListValues(sNewQuery)
                                                Else
                                                  sValue = goHTTPServer.GetValueFromEndpoint(sAPIendpoint, sNewQuery, oReturnNode, sRootNode)

                                                  If String.IsNullOrEmpty(sValue) = False Then
                                                    oListValues.Add(sNewQuery, sValue)
                                                  End If

                                                End If

                                                If sReturnIndex <> "" Then
                                                  sValue = Replace(sValueIDs, sValueID, sValue)
                                                End If

                                                oListofValues.Add(sValue)

                                              End If
                                            Next

                                          End If

                                          Dim sNewArray As String = ""

                                          For Each sValue In oListofValues
                                            sNewArray = String.Format("{0}{1},", sNewArray, sValue)
                                          Next

                                          'If sNewArray.StartsWith(",") Then
                                          '  sNewArray = sNewArray.Remove(0, 1)
                                          'End If

                                          .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value = sNewArray
                                          .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value = (sArray)
                                        End If
                                      End If
                                    End If
                                  Next nRow

                                Catch ex As Exception
                                  If bIgnore = False Then
                                    Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, ex.Message), MsgBoxStyle.AbortRetryIgnore)
                                      Case MsgBoxResult.Abort
                                        Clipboard.SetText(ex.ToString)
                                        Exit For
                                      Case MsgBoxResult.Retry
                                      Case MsgBoxResult.Ignore
                                        bIgnore = True
                                    End Select
                                  End If
                                End Try
                              Case "8"

                                .ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value = oValidation.Comments

                                For nRow = 2 To .ActiveWorksheet.Rows.LastUsedIndex + 1
                                  Call UpdateProgressStatus(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

                                  If .ActiveWorksheet.Cells(String.Format("{0}{1}", oValidation.ReturnCell, nRow)).Value IsNot Nothing AndAlso
                                     String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", oValidation.ReturnCell, nRow)).Value.ToString) = False Then

                                    Dim sArray As String = .ActiveWorksheet.Cells(String.Format("{0}{1}", oValidation.ReturnCell, nRow)).Value.ToString

                                    'sArray = Replace(sArray, ",", "|")
                                    'sArray = Replace(sArray, "[", "")
                                    'sArray = Replace(sArray, "]", "")
                                    'sArray = Replace(sArray, vbNewLine, "")
                                    'sArray = Replace(sArray, " ", "")
                                    'sArray = Replace(sArray, ControlChars.Quote, "")

                                    .ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, nRow)).Value = sArray

                                  End If

                                Next nRow

                            End Select

                          End If
                        Next sKey

                      End If


                      .ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

                      'Hide unused columns
                      Call HideShowColumns(oTemplate, True)

                      For nColumn = 1 To .ActiveWorksheet.Columns.LastUsedIndex
                        Dim oColumn As Column = .ActiveWorksheet.Columns.Item(nColumn)
                        If .ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.Heading, 1)).Value IsNot Nothing Then

                          If mbHideCalulationColumns Then
                            If .ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.Heading, 1)).Value.ToString.StartsWith("translations.en.description") Then
                              .ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.Heading, 1)).Value = "English"
                            End If
                          End If

                        End If
                      Next

                    Else
                      MsgBox("No data found.", vbOKOnly)
                    End If
                  End If
                Else
                  'located the beging of the stack trace
                  Dim sMessage As String = oResponse.Content.ToString.Substring(0, oResponse.Content.ToString.IndexOf("stackTrace"))

                  Select Case MsgBox(String.Format("Error in select query.{0}{0}{1}{0}{0}Error message will be copied to your clipboard.", vbNewLine, sMessage), MsgBoxStyle.OkOnly)
                    Case MsgBoxResult.Ok
                      Clipboard.SetText(sMessage)
                  End Select
                End If

              End If

            End If

            If .ActiveWorksheet.Tables.Count > 0 Then

              ' Access the workbook's collection of table styles.
              Dim tableStyles As TableStyleCollection = spreadsheetControl.Document.TableStyles

              ' Apply the table style to the existing table.
              .ActiveWorksheet.Tables(0).Style = tableStyles("TableStyleMedium7")

            End If


            Application.DoEvents()


          End With
        End If
      End If

    Catch ex As Exception
      MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
    Finally
      spreadsheetControl.EndUpdate()
      bbiCancel.Enabled = False
      mbCancel = False
      Call UpdateProgressStatus("")
      bbiQuery.Enabled = True
    End Try

  End Sub

  Private Sub HideShowColumns(poImportTemplate As clsDataImportTemplate, psHide As Boolean)
    Try

      Dim oTemplate As clsDataImportTemplate = poImportTemplate

      If oTemplate.ImportColumns IsNot Nothing AndAlso oTemplate.ImportColumns.Count >= 0 Then

        With spreadsheetControl

          .BeginUpdate()
          .ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

          'Hide unused columns
          For nColumn = 1 To .ActiveWorksheet.Columns.LastUsedIndex

            If psHide Then
              Dim oColumn As Column = .ActiveWorksheet.Columns.Item(nColumn)
              If .ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.Heading, 1)).Value IsNot Nothing Then
                If .ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.Heading, 1)).Value.ToString.StartsWith("Column") Then
                  .ActiveWorksheet.Columns.Hide(nColumn, nColumn)
                End If
              End If

              If oTemplate.Validators.Where(Function(n) n.ReturnCellWithoutProperties = oColumn.Heading And n.Visibility = 0).Count > 0 Then
                .ActiveWorksheet.Columns.Hide(nColumn, nColumn)
              End If
            Else
              .ActiveWorksheet.Columns.Unhide(nColumn, nColumn)
            End If

          Next

          If String.IsNullOrEmpty(oTemplate.ReturnCell) = False Then
            Dim nCol As Integer = .ActiveWorksheet.Columns(oTemplate.ReturnCell).Index
            .ActiveWorksheet.Columns.Unhide(nCol, nCol)
          End If

        End With
      End If
    Catch ex As Exception

    Finally
      spreadsheetControl.EndUpdate()
    End Try




  End Sub

  Private Function ValidateHierarchiesIsSelected() As Boolean
    If mbContainsHierarchies And gsSelectedHierarchy Is Nothing AndAlso String.IsNullOrEmpty(gsSelectedHierarchy) = True Then
      MsgBox("This template requires a Hierarchy to be selected.", MsgBoxStyle.OkOnly)
      Return False
    Else
      Return True
    End If
  End Function



#End Region

#Region "Objects"
  Private Sub lbxImportTemplates_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxImportTemplates.SelectedIndexChanged

    With lbxImportTemplates
      If .SelectedIndex >= 0 Then
        Dim oItem = .GetItem(.SelectedIndex)
        If oItem IsNot Nothing Then
          UcProperties1._DataImportTemplate = oItem
          Call ValidateDataTempalte(UcProperties1._DataImportTemplate)

          If String.IsNullOrEmpty(UcProperties1._DataImportTemplate.WorkbookSheetName) = False Then
            With spreadsheetControl
              If .Document.Worksheets.Contains(UcProperties1._DataImportTemplate.WorkbookSheetName) = False Then
                .Document.Worksheets.Add(StrConv(UcProperties1._DataImportTemplate.WorkbookSheetName, vbProperCase))
                .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1._DataImportTemplate.WorkbookSheetName)
              Else
                'if its found make it the active worksheet
                .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1._DataImportTemplate.WorkbookSheetName)
              End If
            End With
          End If

        End If
      End If
    End With

  End Sub

  Private Sub bbiVersions_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiVersions.ItemClick
    If bbiVersions.Down = True Then
      LcgVersions.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
    Else
      LcgVersions.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
    End If

  End Sub

  Private Sub lbxImportTemplates_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lbxImportTemplates.MouseDoubleClick

    With lbxImportTemplates
      If .SelectedIndex >= 0 Then
        Dim oItem = .GetItem(.SelectedIndex)
        If oItem IsNot Nothing Then
          UcProperties1._DataImportTemplate = oItem
          Call ValidateDataTempalte(UcProperties1._DataImportTemplate)

          With spreadsheetControl
            If .Document.Worksheets.Contains(UcProperties1._DataImportTemplate.WorkbookSheetName) = False Then
              .Document.Worksheets.Add(StrConv(UcProperties1._DataImportTemplate.WorkbookSheetName, vbProperCase))
              .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1._DataImportTemplate.WorkbookSheetName)
            Else
              'if its found make it the active worksheet
              .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1._DataImportTemplate.WorkbookSheetName)
            End If
          End With

          tcgTabs.SelectedTabPage = lcgSpreedsheet

        End If
      End If
    End With
  End Sub

  Private Sub lbxImportTemplates_MouseDown(sender As Object, e As MouseEventArgs) Handles lbxImportTemplates.MouseDown

    If e.Button = MouseButtons.Right Then
      pumImportTemplates.ShowPopup(Me.MousePosition)
    End If
  End Sub

  Private Sub bbiCheckAll_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCheckAll.ItemClick

    lbxImportTemplates.CheckAll()


  End Sub

  Private Sub bbiUncheckAll_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiUncheckAll.ItemClick
    lbxImportTemplates.UnCheckAll()
  End Sub

  Private Sub lueHierarchies_EditValueChanged(sender As Object, e As EventArgs) Handles lueHierarchies.EditValueChanged
    If lueHierarchies.EditValue IsNot Nothing Then
      gsSelectedHierarchy = lueHierarchies.EditValue.ToString
    Else
      gsSelectedHierarchy = Nothing
    End If

  End Sub


#End Region

End Class
