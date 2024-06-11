Imports DevExpress.Spreadsheet
Imports DevExpress.XtraBars.Helpers
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraSplashScreen
Imports DevExpress.XtraSpreadsheet
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports RestSharp
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Threading
Imports System.IO.Compression
Imports System.ComponentModel

Public Class frmMainMenu


#Region "Form"
     Private Const csLayoutFileName As String = "gridLayouts.xml"

     Private Const EventLogHistory As Integer = 10000
     Private msFilePath As String = String.Empty
     Private msFileName As String
     Private moCopyObjectValidators As New List(Of clsValidation)
     Private moCopyObjectSelectors As New List(Of clsSelectors)
     Private mbForceValidate As Boolean = True
     Private mbHideCalulationColumns As Boolean = True
     Private mbContainsHierarchies As Boolean = False
     Private moOverlayHandle As IOverlaySplashScreenHandle
     Private miButtonImage As Image
     Private miHotButtonImage As Image
     Private moOverlayButton As OverlayImagePainter
     Private moOverlayLabel As OverlayTextPainter
     Private mnCounter As Integer = 0
     Private mbWaitFormShown As Boolean = False
     Private WithEvents moTimer As New System.Windows.Forms.Timer



     Sub New()

          ' Open a Splash Screen
          SplashScreenManager.ShowForm(Me, GetType(frmSplashScreen), True, True, False)
          SplashScreenManager.Default.SendCommand(frmSplashScreen.SplashScreenCommand.SetProgress, "Initializing")


          Me.miButtonImage = CreateButtonImage()
          Me.miHotButtonImage = CreateHotButtonImage()
          Me.moOverlayLabel = New OverlayTextPainter()
          Me.moOverlayButton = New OverlayImagePainter(miButtonImage, miHotButtonImage, AddressOf OnCancelButtonClick)

          InitSkins()
          InitializeComponent()
          Me.InitSkinGallery()

          bbiPublish.Visibility = If(Debugger.IsAttached, DevExpress.XtraBars.BarItemVisibility.Always, DevExpress.XtraBars.BarItemVisibility.Never)

          moTimer.Interval = 1000
          moTimer.Start()

          SplashScreenManager.Default.SendCommand(frmSplashScreen.SplashScreenCommand.SetProgress, "Loading Data Connections")


          UcProperties1.ImportHeader = New clsDataImportHeader
          UcConnectionDetails1.LoadSettingFile(True, False)
          SplashScreenManager.Default.SendCommand(frmSplashScreen.SplashScreenCommand.SetProgress, "Connecting to last Server")
          UcConnectionDetails1.TestConnection(goConnection, True)

          siInfo.Caption = String.Format("Version:{0}", My.Application.Info.Version)
          lueHierarchies.Properties.DataSource = goHierarchies

          SplashScreenManager.Default.SendCommand(frmSplashScreen.SplashScreenCommand.SetProgress, "Adding hooks")

          AddHandler goHTTPServer.APICallEvent, AddressOf APICallEvent
          AddHandler goHTTPServer.ErrorEvent, AddressOf ErrorEvent


          With labelSystem
               .Text = goConnection.Key
          End With

          AddHandler UcConnectionDetails1.ChangeOfConnection, AddressOf UpdateSystemLabel

     End Sub

     Sub InitSkins()
          DevExpress.Skins.SkinManager.EnableFormSkins()
          DevExpress.UserSkins.BonusSkins.Register()
     End Sub

     Private Sub InitSkinGallery()
          SkinHelper.InitSkinGallery(rgbiSkins, True)
     End Sub

     Private Function CreateButtonImage() As Image
          Return CreateImage(My.Resources.cancel_normal)
     End Function

     Private Function CreateHotButtonImage() As Image
          Return CreateImage(My.Resources.cancel_active)
     End Function

     Private Sub frmMainMenu_Shown(sender As Object, e As EventArgs) Handles Me.Shown
          ' Close the Splash Screen.
          SplashScreenManager.Default.SendCommand(frmSplashScreen.SplashScreenCommand.SetProgress, "Initialization Completed")
          SplashScreenManager.CloseForm(False)

          If String.IsNullOrEmpty(ReadRegistry("Last Accessed Workbook")) = False Then
               OpenWorkBook(ReadRegistry("Last Accessed Workbook", ""), ReadRegistry("Last Accessed Workbook Safe Name", ""))
               UpdateProgressStatus()
          End If

          LoadGridLayouts(Me, csLayoutFileName)


     End Sub

     Private Sub frmMainMenu_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

          SaveGridLayouts(Me, csLayoutFileName)


     End Sub

#End Region

#Region "Buttons"

     Private Sub bbiImportOpen_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportOpen.ItemClick

          Try


               Dim fd As OpenFileDialog = New OpenFileDialog
               fd.Multiselect = False
               fd.Title = "Select a Otalio Dynamic import template."
               fd.Filter = String.Format("Dynamic import Template V1|*.dit|Dynamic import Template V2|*{0}", ".dit2")
               fd.FilterIndex = 1
               fd.FileName = String.Empty
               fd.RestoreDirectory = True
               fd.Multiselect = True
               If fd.ShowDialog() = DialogResult.OK Then

                    UpdateProgressStatus("Opening Import", False)

                    For Each sFilename As String In fd.FileNames
                         OpenTempalte(sFilename, fd.SafeFileName)
                    Next


               End If
          Catch ex As Exception

          Finally

               UpdateProgressStatus("")

          End Try
     End Sub

     Private Sub bbiSaveImport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveImport.ItemClick

          Try


               If gdWorkbook.SelectedRowsCount = 0 Then
                    MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
                    Exit Sub
               End If
               If MsgBox("You are about to save one or more import templates.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then


                    Using oFolder As New FolderBrowserDialog

                         oFolder.Description = "Please select a folder to save template in"

                         If String.IsNullOrEmpty(goConnectionHistory.LastUsedExportFolder) = False Then
                              oFolder.SelectedPath = goConnectionHistory.LastUsedExportFolder
                         End If

                         If oFolder.ShowDialog = DialogResult.OK Then
                              Dim sFolderName As String = oFolder.SelectedPath

                              goConnectionHistory.LastUsedExportFolder = sFolderName

                              For Each nRowID As Integer In gdWorkbook.GetSelectedRows
                                   If nRowID >= 0 Then
                                        Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                                        If oItem IsNot Nothing Then
                                             Try

                                                  Dim oImportHeader As clsDataImportHeader = TryCast(oItem, clsDataImportHeader)
                                                  If oImportHeader IsNot Nothing Then
                                                       Call ValidateDataTempalte(oImportHeader)
                                                       UcProperties1.SaveSettings(String.Format("{0}\{1}{2}", sFolderName, oImportHeader.Name, gsImportTemplateHeaderExtention), oImportHeader)
                                                  End If

                                             Catch ex As Exception

                                             End Try

                                             Application.DoEvents()
                                        End If
                                   End If

                              Next
                         End If
                    End Using

               End If



          Catch ex As Exception

          End Try




     End Sub

     Private Sub bbiSaveImportAs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveImportAs.ItemClick
          Call ValidateDataTempalte(UcProperties1.ImportHeader)
          Call UcProperties1.SaveSettings(True)
     End Sub

     Private Sub bbiCreateEmptyExcelDocument_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCreateEmptyExcelDocument.ItemClick

          For Each nRowID As Integer In gdWorkbook.GetSelectedRows
               If nRowID >= 0 Then
                    Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                    Call createNewExcelSheet(TryCast(oItem, clsDataImportHeader))
                    Application.DoEvents()
               End If
          Next

     End Sub

     Private Sub bbiImportExcel_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportExcel.ItemClick

          VerifySelectedHierarchy()

          If gdWorkbook.SelectedRowsCount = 0 Then
               MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
               Exit Sub
          End If

          If MsgBox("You are about to import data from the Excel sheet.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then
               Dim bIgnore As Boolean = False

               For Each nRowID As Integer In gdWorkbook.GetSelectedRows
                    If nRowID >= 0 Then
                         Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)

                         If oItem.ReadOnlyImport = False Then
                              Dim bApply As Boolean = True

                              If VerifyHierarchSelection(oItem) = False And bIgnore = False Then
                                   Select Case MsgBox(String.Format("This import template {0} requires a valid {1} selected.{2}{2}Select Abort to exit, retry to move to next template and ignore to ignore all errors.", oItem.Name, IIf(oItem.IsShipEntity, "Ship", "Hierarchy"), vbNewLine), vbAbortRetryIgnore, "Warning...")
                                        Case MsgBoxResult.Abort : Exit Sub
                                        Case MsgBoxResult.Ignore : bIgnore = True : bApply = False
                                        Case MsgBoxResult.Retry : bApply = False
                                   End Select
                              End If

                              If bApply Then
                                   '   Select Case TryCast(oItem, clsDataImportHeader).ImportType
                                   'Case "XX"
                                   '          ImportFiles(TryCast(oItem, clsDataImportHeader))
                                   '          Application.DoEvents()
                                   '     Case "XXX"
                                   '          DeleteRecords(TryCast(oItem, clsDataImportHeader))
                                   '          Application.DoEvents()
                                   '     Case Else
                                   ImportData(TryCast(oItem, clsDataImportHeader))
                                   Application.DoEvents()

                                   ' End Select
                              End If

                         End If

                         MsgBox(String.Format("This import template {0} is Read Only and cannot be imported.", oItem.Name), vbAbortRetryIgnore, "Warning...")
                    End If

               Next

          End If


     End Sub

     Private Sub bbiValidateData_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidateData.ItemClick

          Try


               VerifySelectedHierarchy()

               If gdWorkbook.SelectedRowsCount = 0 Then
                    MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
                    Exit Sub
               End If

               Dim bIgnore As Boolean = False

               For Each nRowID As Integer In gdWorkbook.GetSelectedRows
                    If nRowID >= 0 Then
                         Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                         If oItem IsNot Nothing AndAlso oItem.ReadOnlyImport = False Then



                              Try

                                   Dim bApply As Boolean = True

                                   If VerifyHierarchSelection(oItem) = False And bIgnore = False Then
                                        Select Case MsgBox(String.Format("This import template {0} requires a valid {1} selected.{2}{2}Select Abort to exit, retry to move to next template and ignore to ignore all errors.", oItem.Name, IIf(oItem.IsShipEntity, "Ship", "Hierarchy"), vbNewLine), vbAbortRetryIgnore, "Warning...")
                                             Case MsgBoxResult.Abort : Exit Sub
                                             Case MsgBoxResult.Ignore : bIgnore = True : bApply = False
                                             Case MsgBoxResult.Retry : bApply = False
                                        End Select
                                   End If


                                   If bApply Then
                                        UpdateProgressStatus("Validating...")
                                        Call ValidateDataTempalte(TryCast(oItem, clsDataImportHeader))
                                        Call ValidateData(TryCast(oItem, clsDataImportHeader))
                                   End If

                              Catch ex As Exception

                              End Try

                              Application.DoEvents()
                         End If
                    End If
               Next
          Catch ex As Exception
          Finally
               UpdateProgressStatus("")
          End Try

     End Sub

     Private Sub bbiFormatJsonCode_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiFormatJsonCode.ItemClick
          Try

               Dim sResult As String = ValidateJson(UcProperties1.recDTO.Text)
               If sResult <> "OK" Then
                    ShowErrorForm(sResult)
                    Exit Sub
               End If
               UcProperties1.linkedItemsCountDTO = 0
               UcProperties1.recDTO.Text = JValue.Parse(UcProperties1.recDTO.Text).ToString(Newtonsoft.Json.Formatting.Indented)

          Catch ex As Exception
               ShowErrorForm(ex)
          End Try
     End Sub

     Private Sub bbiValidatorAdd_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidatorAdd.ItemClick

          Using oForm As New frmValidatorEdit
               oForm._ucValidationProperties._Validation = New clsValidation
               Select Case oForm.ShowDialog
                    Case DialogResult.OK 'ok add

                         UcProperties1.SelectedTemplate.Validators.Add(oForm._ucValidationProperties._Validation)
                    Case DialogResult.Cancel 'dont add
                    Case DialogResult.Abort 'use to flag item to be deleted
                         UcProperties1.SelectedTemplate.Validators.Remove(oForm._ucValidationProperties._Validation)

               End Select

               UcProperties1.gridValidators.RefreshDataSource()
          End Using

     End Sub

     Private Sub bbiDuplicate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiDuplicate.ItemClick

          Select Case UcProperties1.tcgTemplateObjects.SelectedTabPage.Text.ToString
               Case "Validation Rules"
                    If moCopyObjectValidators IsNot Nothing Then
                         For Each oValidation As clsValidation In moCopyObjectValidators
                              If oValidation IsNot Nothing Then
                                   Dim oNewValidation = oValidation.Clone
                                   UcProperties1.SelectedTemplate.Validators.Add(oNewValidation)
                              End If
                         Next
                         UcProperties1.gridValidators.RefreshDataSource()
                    End If
               Case "Selectors"
                    If moCopyObjectSelectors IsNot Nothing Then
                         For Each oSelector As clsSelectors In moCopyObjectSelectors
                              If oSelector IsNot Nothing Then
                                   Dim oNewValidation = oSelector.Clone
                                   UcProperties1.SelectedTemplate.Selectors.Add(oNewValidation)
                              End If
                         Next
                         UcProperties1.gridSelectors.RefreshDataSource()
                    End If
          End Select



     End Sub

     Private Sub bbiCopyObject_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCopyValidator.ItemClick

          Select Case UcProperties1.tcgTemplateObjects.SelectedTabPage.Text.ToString
               Case "Validation Rules"
                    'clear the existing cash of validations
                    moCopyObjectValidators = New List(Of clsValidation)

                    Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdValidators.GetSelectedRows.ToList, List(Of Integer))
                    If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then

                         For Each orow As Integer In oSelectedRows
                              Dim oValidation As clsValidation = TryCast(UcProperties1.gdValidators.GetRow(orow), clsValidation)
                              If oValidation IsNot Nothing Then
                                   moCopyObjectValidators.Add(oValidation)
                              End If
                              bbiDuplicate.Enabled = True
                         Next

                    End If
               Case "Selectors"
                    'clear the existing cash of validations
                    moCopyObjectSelectors = New List(Of clsSelectors)

                    Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdSelectors.GetSelectedRows.ToList, List(Of Integer))
                    If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then

                         For Each orow As Integer In oSelectedRows
                              Dim oSelector As clsSelectors = TryCast(UcProperties1.gdSelectors.GetRow(orow), clsSelectors)
                              If oSelector IsNot Nothing Then
                                   moCopyObjectSelectors.Add(oSelector)
                              End If
                              bbiDuplicate.Enabled = True
                         Next

                    End If
          End Select


     End Sub

     Private Sub bbiValidatorDelete_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiValidatorDelete.ItemClick

          Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdValidators.GetSelectedRows.ToList, List(Of Integer))
          Dim oValidations As New List(Of clsValidation)
          If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then
               For Each orow As Integer In oSelectedRows

                    Dim oValidation As clsValidation = TryCast(UcProperties1.gdValidators.GetRow(orow), clsValidation)
                    If oValidation IsNot Nothing Then
                         oValidations.Add(oValidation)
                    End If
               Next

               If oValidations.Count > 0 Then
                    For Each oValidation As clsValidation In oValidations
                         UcProperties1.SelectedTemplate.Validators.Remove(oValidation)
                    Next
               End If

               UcProperties1.gridValidators.RefreshDataSource()

          End If

     End Sub

     Private Sub bbiOpenWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiOpenWorkbook.ItemClick


          Try



               Dim fd As OpenFileDialog = New OpenFileDialog
               fd.Multiselect = False
               fd.Title = "Select a Otalio Dynamic import workbook."
               fd.Filter = String.Format("Dynamic import workbook |*{0}", gsWorkbookExtention)
               fd.FilterIndex = 1
               fd.FileName = txtWorkbookName.Text
               fd.RestoreDirectory = True
               If fd.ShowDialog() = DialogResult.OK Then

                    Dim sFilename As String = fd.FileName

                    OpenWorkBook(sFilename, fd.SafeFileName)


               End If

          Catch ex As Exception
          Finally
               UpdateProgressStatus()
          End Try

     End Sub

     Private Sub bbiSaveWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveWorkbook.ItemClick
          SaveWorkbook(False)
     End Sub

     Private Sub bbiSaveAsWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveAsWorkbook.ItemClick
          SaveWorkbook(True)
     End Sub

     Private Sub bbiAddTemplate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiAddTemplate.ItemClick
          goOpenWorkBook.Templates.Add(New clsDataImportHeader)
          loadWorkbook()
     End Sub

     Private Sub bbiRemoveTemplate_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiRemoveTemplate.ItemClick


          If gdWorkbook.SelectedRowsCount = 0 Then
               MsgBox("Please select one or more import templates to remove ",, "Warning...")
               Exit Sub
          End If
          If MsgBox("You are about to remove one or more import templates.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then

               For Each nRowID As Integer In gdWorkbook.GetSelectedRows
                    If nRowID >= 0 Then
                         Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                         goOpenWorkBook.Templates.Remove(TryCast(oItem, clsDataImportHeader))
                    End If
               Next

               loadWorkbook()
          End If

     End Sub

     Private Sub bbiCancel_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCancel.ItemClick
          If mbCancel = False Then mbCancel = True
     End Sub

     Private Sub bbiSaveAll_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSaveAll.ItemClick

          UpdateProgressStatus("Saving Files", False)

          If String.IsNullOrEmpty(msFilePath) Then

               Using oFolder As New FolderBrowserDialog

                    oFolder.Description = "Please select a folder to save template in"

                    If String.IsNullOrEmpty(goConnectionHistory.LastUsedWorkbookFolder) = False Then
                         oFolder.SelectedPath = goConnectionHistory.LastUsedWorkbookFolder
                    End If

                    If oFolder.ShowDialog = DialogResult.OK Then
                         msFilePath = oFolder.SelectedPath
                         goConnectionHistory.LastUsedWorkbookFolder = msFileName
                    End If
               End Using
          End If

          If String.IsNullOrEmpty(msFilePath) = False Then

               txtWorkbookName.Focus()

               Call ValidateDataTempalte(UcProperties1.ImportHeader)

               Dim sPath As String = Path.GetDirectoryName(msFilePath)
               Dim sFileName As String = Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName)
               Dim sExtention As String = Path.GetExtension(goOpenWorkBook.WorkbookName)
               Dim sVersion As String = goOpenWorkBook.WorkbookVersion
               sVersion = Replace(sVersion, ".", "-")

               UpdateProgressStatus("Saving import workbook", False)
               'create backups
               SaveWorkbook(False, Path.GetDirectoryName(msFilePath), "BACKUP",
                          String.Format("{0} {1}", Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName), sVersion),
                          Path.GetExtension(goOpenWorkBook.WorkbookName), False)


               SortWorksheets()

               SaveWorkbook(False)

               UpdateProgressStatus("Saving excel workbook", False)
               Dim sCurrentFilename As String = spreadsheetControl.Document.Path
               If String.IsNullOrEmpty(sCurrentFilename) = True Then
                    sCurrentFilename = String.Format("{0}\{1}.{2}", msFilePath, sFileName, "xlsx")
               End If

               Dim sExcelExtention As String = Path.GetExtension(sCurrentFilename)
               spreadsheetControl.SaveDocument(String.Format("{0}\{1}\{2} {3}{4}", sPath, "BACKUP", sFileName, sVersion, sExcelExtention))
               spreadsheetControl.SaveDocument(sCurrentFilename)

          End If



          UpdateProgressStatus()

     End Sub

     Private Sub bbiQuery_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiQuery.ItemClick




          Dim oLastSelectedTemplate As clsDataImportHeader = Nothing
          Call VerifySelectedHierarchy()

          If gdWorkbook.SelectedRowsCount = 0 Then
               MsgBox("Please select one or more import templates by checking them from the main list ",, "Warning...")
               Exit Sub
          End If

          If MsgBox("Running the query will remove all existing data from the excel worksheet.  Are you sure?", vbYesNoCancel, "Confirm...") = vbYes Then
               Dim bIgnore As Boolean = False

               For Each nRowID As Integer In gdWorkbook.GetSelectedRows
                    If nRowID >= 0 Then
                         Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                         If oItem IsNot Nothing Then
                              oLastSelectedTemplate = oItem
                              Dim bApply As Boolean = True

                              If VerifyHierarchSelection(oItem) = False And bIgnore = False Then
                                   Select Case MsgBox(String.Format("This import template {0} requires a valid {1} selected.{2}{2}Select Abort to exit, retry to move to next template and ignore to ignore all errors.", oItem.Name, IIf(oItem.IsShipEntity, "Ship", "Hierarchy"), vbNewLine), vbAbortRetryIgnore, "Warning...")
                                        Case MsgBoxResult.Abort : Exit Sub
                                        Case MsgBoxResult.Ignore : bIgnore = True : bApply = False
                                        Case MsgBoxResult.Retry : bApply = False
                                   End Select
                              End If

                              If bApply Then
                                   Try


                                        If TryCast(oItem, clsDataImportHeader).Templates.ToList.Where(Function(n) n.IsMaster = True).Count > 1 Then
                                             MsgBox("Query of Data is not supported in multiple template imports.")
                                             Exit Sub
                                        End If

                                        QueryandLoad(TryCast(oItem, clsDataImportHeader))
                                        Application.DoEvents()
                                   Catch ex As Exception

                                   End Try


                              End If
                         End If
                    End If


               Next

               If oLastSelectedTemplate IsNot Nothing Then
                    SetEditExcelMode(False, oLastSelectedTemplate, True)
               End If
          End If


     End Sub

     Private Sub beAction_ButtonPressed(sender As Object, e As ButtonPressedEventArgs) Handles beAction.ButtonPressed

          Dim otemplate As clsDataImportHeader = CType(gdWorkbook.GetRow(gdWorkbook.FocusedRowHandle), clsDataImportHeader)
          If otemplate IsNot Nothing Then

               Select Case e.Button.Caption.ToUpper
                    Case "VIEW"
                         Call ViewExcelDocument(otemplate)
                    Case "VALIDATE"
                    Case "IMPORT"
                    Case "QUERY"
               End Select
          End If
     End Sub

     Private Sub butNewWorkbook_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles butNewWorkbook.ItemClick
          If MsgBox("You are about to create a new workbook, are you sure?", vbYesNoCancel) = vbYes Then

               Dim sWorkbookName As String = InputBox("Workbook Name", "User Input", "New Workbook")

               Dim oWorkbook As New clsWorkbook
               oWorkbook.WorkbookName = sWorkbookName

               'make one last backup of current file
               If String.IsNullOrEmpty(msFilePath) = False Then

                    txtWorkbookName.Focus()

                    Call ValidateDataTempalte(UcProperties1.ImportHeader)

                    Dim sPath As String = Path.GetDirectoryName(msFilePath)
                    Dim sFileName As String = Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName)
                    Dim sExtention As String = Path.GetExtension(goOpenWorkBook.WorkbookName)
                    Dim sVersion As String = goOpenWorkBook.WorkbookVersion
                    sVersion = Replace(sVersion, ".", "-")

                    'create backups
                    SaveWorkbook(False, Path.GetDirectoryName(msFilePath), "BACKUP",
                               String.Format("{0} {1}", Path.GetFileNameWithoutExtension(goOpenWorkBook.WorkbookName), sVersion),
                               Path.GetExtension(goOpenWorkBook.WorkbookName), False)


                    Dim sCurrentFilename As String = spreadsheetControl.Document.Path
                    If String.IsNullOrEmpty(sCurrentFilename) = True Then
                         sCurrentFilename = String.Format("{0}\{1}.{2}", sPath, sFileName, "xlsx")
                    End If

                    Dim sExcelExtention As String = Path.GetExtension(sCurrentFilename)
                    spreadsheetControl.SaveDocument(String.Format("{0}\{1}\{2} {3}{4}", sPath, "BACKUP", sFileName, sVersion, sExcelExtention))


               End If

               goOpenWorkBook = oWorkbook
               spreadsheetControl.CreateNewDocument()
               msFilePath = ""
               msFileName = sWorkbookName & gsWorkbookExtention
               UcProperties1.ImportHeader = New clsDataImportHeader
               loadWorkbook()

          End If
     End Sub

     Private Sub bbiPublish_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPublish.ItemClick

          Dim sPath As String = ReadRegistry("Last Published Folder", "")
          Dim sFileName As String = String.Format("{0} {1}", msFileName, goOpenWorkBook.WorkbookVersion)



          'first save everything
          bbiSaveAll_ItemClick(Me, e)

          Using oFolder As New FolderBrowserDialog

               oFolder.Description = "Please select a folder to save template in"
               oFolder.SelectedPath = sPath


               If oFolder.ShowDialog = DialogResult.OK Then
                    sPath = oFolder.SelectedPath
               End If
          End Using

          Dim sFullFileName As String = sPath & "\" & (sFileName.Replace(gsWorkbookExtention, "")) & ".zip"

          If File.Exists(sFullFileName) Then
               File.Delete(sFullFileName)
          End If

          Dim oZipFile As ZipArchive = ZipFile.Open(sFullFileName, ZipArchiveMode.Create)
          oZipFile.CreateEntryFromFile(msFilePath & "\" & msFileName, msFileName)
          oZipFile.Dispose()

          WriteRegistry("Last Published Folder", sPath)

     End Sub

     Private Sub rbiSortTabs_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles rbiSortTabs.ItemClick
          SortWorksheets()
     End Sub

     Private Sub bbiExecutionPlan_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExecutionPlan.ItemClick
          UcProperties1.AddTemplate()
     End Sub

     Private Sub bbiExecutionPlanRemove_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExecutionPlanRemove.ItemClick
          UcProperties1.RemoveTemplate()
     End Sub

     Private Sub bbiImportIntoTemplateHeader_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportIntoTemplateHeader.ItemClick

          Try


               Dim fd As OpenFileDialog = New OpenFileDialog
               fd.Multiselect = False
               fd.Title = "Select a Otalio Dynamic import template."
               fd.Filter = String.Format("Dynamic import Template V1|*.dit|Dynamic import Template V2|*{0}", ".dit2")
               fd.FilterIndex = 1
               fd.FileName = String.Empty
               fd.RestoreDirectory = True
               fd.Multiselect = True
               If fd.ShowDialog() = DialogResult.OK Then

                    UpdateProgressStatus("Opening Import", False)

                    For Each sFilename As String In fd.FileNames

                         If File.Exists(sFilename) Then

                              Dim oImportHeader As New clsDataImportHeader

                              If sFilename.EndsWith(".dit") Then
                                   Dim oImportTemplate As clsDataImportTemplate = TryCast(LoadFile(sFilename), clsDataImportTemplate)
                                   If oImportTemplate IsNot Nothing Then
                                        With oImportHeader

                                             .Name = oImportTemplate.Name
                                             .ID = oImportTemplate.ID

                                             .Notes = oImportTemplate.Notes
                                             .Priority = oImportTemplate.Priority
                                             .Group = oImportTemplate.Group
                                             .WorkbookSheetName = oImportTemplate.WorkbookSheetName

                                             Dim oTemplate As New clsDataImportTemplateV2
                                             With oTemplate
                                                  .ImportType = oImportTemplate.ImportType
                                                  .Name = oImportTemplate.Name
                                                  .APIEndpoint = oImportTemplate.APIEndpoint
                                                  .APIEndpointSelect = oImportTemplate.APIEndpointSelect
                                                  .DTOObject = oImportTemplate.DTOObject
                                                  .EntityColumn = oImportTemplate.EntityColumn
                                                  .GraphQLQuery = oImportTemplate.GraphQLQuery
                                                  .GraphQLRootNode = oImportTemplate.GraphQLRootNode
                                                  .ImportColumns = oImportTemplate.ImportColumns
                                                  .ReturnCell = oImportTemplate.ReturnCell
                                                  .ReturnCellDTO = oImportTemplate.ReturnCellDTO
                                                  .ReturnNodeValue = oImportTemplate.ReturnNodeValue
                                                  .SelectQuery = oImportTemplate.SelectQuery
                                                  .StatusCodeColumn = oImportTemplate.StatusCodeColumn
                                                  .StatusDescriptionColumn = oImportTemplate.StatusDescirptionColumn
                                                  .UpdateQuery = oImportTemplate.UpdateQuery
                                                  .Validators = oImportTemplate.Validators
                                                  .Variables = oImportTemplate.Variables
                                                  .Position = oImportHeader.Templates.Count + 1
                                             End With

                                             .Templates.Add(oTemplate)
                                        End With
                                   End If
                              Else
                                   oImportHeader = TryCast(LoadFile(sFilename), clsDataImportHeader)

                              End If


                              If oImportHeader IsNot Nothing Then



                                   For Each oTemplate As clsDataImportTemplateV2 In oImportHeader.Templates
                                        If oTemplate.Validators Is Nothing Then
                                             oTemplate.Validators = New List(Of clsValidation)
                                        End If
                                   Next



                                   UcProperties1.ImportHeader.Templates.AddRange(oImportHeader.Templates)
                                   UcProperties1.Reload()
                                   Call ValidateDataTempalte(UcProperties1.ImportHeader)


                              Else
                                   UpdateProgressStatus()
                                   MsgBox("Failed to load Data import template file")
                              End If
                         End If
                    Next


               End If
          Catch ex As Exception

          Finally

               UpdateProgressStatus("")

          End Try
     End Sub

     Private Sub bbiFormatExcelSheet_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiFormatExcelSheet.ItemClick

          For Each nRowID As Integer In gdWorkbook.GetSelectedRows
               If nRowID >= 0 Then
                    Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                    ValidateDataTempalte(TryCast(oItem, clsDataImportHeader))
                    FormatExcelSheet(TryCast(oItem, clsDataImportHeader))
                    Application.DoEvents()
               End If
          Next
     End Sub

     Private Sub bbiShowAllColumns_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiShowAllColumns.ItemClick
          For Each nRowID As Integer In gdWorkbook.GetSelectedRows
               If nRowID >= 0 Then
                    Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(nRowID), clsDataImportHeader)
                    HideShowColumns(oItem, False)
                    Application.DoEvents()
               End If
          Next
     End Sub

     Private Sub bbiClearFormatting_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClearFormatting.ItemClick
          ApplyExcelStyleTemplate(True)

     End Sub

     Private Sub bbiTranslations_CheckedChanged(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiTranslations.CheckedChanged
          gbTranslationsEnabled = bbiTranslations.Checked

          Call ValidateDataTempalte(UcProperties1.ImportHeader)
     End Sub

     Private Sub bbiSelectorAdd_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSelectorAdd.ItemClick

          Using oForm As New frmSelector
               oForm.UcSelectors1._Selector = New clsSelectors
               Select Case oForm.ShowDialog
                    Case DialogResult.OK 'ok add

                         UcProperties1.SelectedTemplate.Selectors.Add(oForm.UcSelectors1._Selector)
                    Case DialogResult.Cancel 'dont add
                    Case DialogResult.Abort 'use to flag item to be deleted
                         UcProperties1.SelectedTemplate.Selectors.Remove(oForm.UcSelectors1._Selector)

               End Select

               UcProperties1.gridSelectors.RefreshDataSource()
          End Using

     End Sub

     Private Sub bbiSelectorDelete_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSelectorDelete.ItemClick


          Dim oSelectedRows As List(Of Integer) = TryCast(UcProperties1.gdSelectors.GetSelectedRows.ToList, List(Of Integer))
          Dim oSelectors As New List(Of clsSelectors)
          If oSelectedRows IsNot Nothing And oSelectedRows.Count > 0 Then
               For Each orow As Integer In oSelectedRows

                    Dim oSelector As clsSelectors = TryCast(UcProperties1.gdSelectors.GetRow(orow), clsSelectors)
                    If oSelector IsNot Nothing Then
                         oSelectors.Add(oSelector)
                    End If
               Next

               If oSelectors.Count > 0 Then
                    For Each oSelector As clsSelectors In oSelectors
                         UcProperties1.SelectedTemplate.Selectors.Remove(oSelector)
                    Next
               End If

               UcProperties1.gridValidators.RefreshDataSource()

          End If
     End Sub

     Private Sub bbiRestoreGridsLayouts_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiRestoreGridsLayouts.ItemClick

          Try
               Dim xmlFilePath As String = csLayoutFileName

               If File.Exists(xmlFilePath) Then
                    File.Delete(xmlFilePath)
                    Dim result As DialogResult = MessageBox.Show("The layout file has been deleted. To restore the layout, the application needs to be restarted. Do you want to restart now?", "Restart Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                    If result = DialogResult.Yes Then
                         Application.Restart()
                    End If
               Else
                    ShowErrorForm($"File not found: {xmlFilePath}", New FileNotFoundException())
               End If

          Catch ex As Exception
               ShowErrorForm("Error deleting layout file", ex)
          End Try
     End Sub

#End Region

#Region "Core Function"
     Private Sub ValidateDataTempalte(oImportHeader As clsDataImportHeader)
          Try


               'extract list of columns from the Json code
               Dim oTotalVariables As New List(Of clsVariables)

               'extract only column used for the data import
               Dim oTotalActiveColumns As New List(Of clsImportColum)

               For Each oTempalte As clsDataImportTemplateV2 In oImportHeader.Templates

                    If oTempalte.Selectors Is Nothing Then
                         oTempalte.Selectors = New List(Of clsSelectors)
                    End If

                    If oImportHeader.Templates.Count = 1 Then
                         oTempalte.IsMaster = True
                         oTempalte.IsEnabled = True
                    End If



                    mbContainsHierarchies = False
                    If oTempalte.GraphQLQuery Is Nothing Then oTempalte.GraphQLQuery = ""
                    If oImportHeader.ID Is Nothing Then oImportHeader.ID = GenerateGUID()

                    'extract list of columns from the Json code
                    Dim oColumns As List(Of clsImportColum) = ExactExcelColumns(oTempalte.DTOObjectFormated).OrderBy(Function(n) n.No).ToList

                    'extract only column used for the data import
                    Dim oActiveColumns As New List(Of clsImportColum)
                    For Each oColumn As clsImportColum In oColumns.Where(Function(n) n.ColumnID <> String.Empty).OrderBy(Function(n) n.No).ToList
                         oColumn.Name = String.Format("{0} ({1})", oTempalte.DataObject, oColumn.Name)
                         oColumn.DataType = "Data Source"
                         oActiveColumns.Add(oColumn)
                    Next

                    Dim oVariables As New List(Of clsVariables)

                    'look for variables
                    For Each oColumn As clsImportColum In oColumns
                         If oColumn.VariableName IsNot Nothing AndAlso String.IsNullOrEmpty(oColumn.VariableName) = False Then
                              With oVariables
                                   If .Exists(Function(n) n.Name = oColumn.Name) = False Then
                                        .Add(New clsVariables With {.Name = oColumn.VariableName, .Value = String.Empty})
                                   End If
                              End With
                         End If
                    Next

                    'remove any unused ones
                    With oTempalte
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
                                   If oTempalte.Variables.Exists(Function(n) n.Name = oVariable.Name) = False Then
                                        .Add(oVariable)
                                   End If
                              Next
                         End If
                    End With

                    If mbContainsHierarchies = False Then mbContainsHierarchies = oTempalte.SelectQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
                    If mbContainsHierarchies = False Then mbContainsHierarchies = oTempalte.UpdateQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
                    If mbContainsHierarchies = False Then mbContainsHierarchies = oTempalte.APIEndpoint.ToString.ToUpper.Contains("@@HIERARCHYID@@")
                    If mbContainsHierarchies = False Then mbContainsHierarchies = oTempalte.GraphQLQuery.ToString.ToUpper.Contains("@@HIERARCHYID@@")
                    If mbContainsHierarchies = False Then mbContainsHierarchies = oTempalte.DTOObject.ToString.ToUpper.Contains("@@HIERARCHYID@@")


                    ValidateTranslationsValidators(oTempalte, gbTranslationsEnabled)

                    If oTempalte.Validators IsNot Nothing Then

                         For Each oValidation As clsValidation In oTempalte.Validators

                              If String.IsNullOrEmpty(oValidation.Visibility) Then
                                   oValidation.Visibility = "1"
                              End If
                              If String.IsNullOrEmpty(oValidation.Query) = False Then
                                   oValidation.Query = Replace(oValidation.Query, "  ", String.Empty)
                                   oValidation.Query = Replace(oValidation.Query, vbNewLine, String.Empty)

                                   If mbContainsHierarchies = False Then mbContainsHierarchies = oValidation.Query.ToString.ToUpper.Contains("@@HIERARCHYID@@")
                                   If mbContainsHierarchies = False Then mbContainsHierarchies = oValidation.APIEndpoint.ToString.ToUpper.Contains("@@HIERARCHYID@@")

                              End If

                              If oValidation.ValidationType <> "2" Then
                                   Dim oCols As List(Of clsImportColum) = ExtractListOfColumnsFromString(oValidation.Query)

                                   If oCols IsNot Nothing AndAlso oCols.Count > 0 Then
                                        For Each oCol In oCols.ToList
                                             If oActiveColumns.Exists(Function(n) n.ColumnName = oCol.ColumnName) = False Then
                                                  oCol.DataType = "Validator Query"
                                                  If String.IsNullOrEmpty(oValidation.GraphQLSourceNode) = False Then
                                                       If oValidation.GraphQLSourceNode.Contains("!") Then
                                                            oCol.Parent = Replace(oValidation.GraphQLSourceNode.ToString.Trim, "!", "")
                                                            oCol.DTOField = oCol.Name
                                                       Else
                                                            oCol.Parent = Replace(oValidation.GraphQLSourceNode.ToString.Trim, "!", "")
                                                            oCol.DTOField = oCol.Name
                                                       End If

                                                  Else
                                                       oCol.Name = String.Format("{0} ({1})", oValidation.DataObject, oCol.Name)
                                                  End If

                                                  oActiveColumns.Add(oCol)
                                             End If
                                        Next



                                   End If

                                   If String.IsNullOrEmpty(oValidation.ReturnCell) = False Then
                                        Dim sCol As String = oValidation.ReturnCell
                                        Dim oCol As New clsImportColum
                                        With oCol

                                             .ColumnName = oValidation.ReturnCell
                                             .Name = oValidation.Comments
                                             .Type = JTokenType.None
                                             .ColumnID = oValidation.ReturnCell
                                             .DataType = "Calculated Validator"
                                             Select Case oValidation.ValidationType
                                                  Case "5", "6"
                                                       .Name = oValidation.DetailedDescriptionFromReturnValue.ToString

                                                  Case "1", "4"
                                                       .Name = oValidation.DetailedDescriptionFromQuery.ToString

                                             End Select


                                             If sCol.Contains(":") Then
                                                  Dim oValues As String() = sCol.Split(":")
                                                  If oValues.Count > 1 Then
                                                       .ColumnName = oValues(0)
                                                  End If

                                                  If oValues.Count > 4 Then
                                                       .ArrayID = oValues(4)
                                                  End If
                                             End If


                                             'If oActiveColumns.Exists(Function(n) n.ColumnName = .ColumnName) = True Then
                                             '     oActiveColumns.Remove(oActiveColumns.Where(Function(n) n.ColumnName = .ColumnName).First)
                                             'End If

                                             If oActiveColumns.Exists(Function(n) n.ColumnName = .ColumnName) = False Then
                                                  oActiveColumns.Add(oCol)
                                             Else
                                                  Dim oCol1 As clsImportColum = oActiveColumns.Where(Function(n) n.ColumnName = .ColumnName).First
                                                  oCol1.DataType = "Calculated Validator & DTO"
                                             End If

                                        End With


                                   End If

                              End If

                              If oValidation.Formatted Is Nothing Then oValidation.Formatted = ""

                         Next
                    End If

                    If String.IsNullOrEmpty(oTempalte.ReturnCell) = False Then
                         If oActiveColumns.Exists(Function(n) n.ColumnName = oTempalte.ReturnCell) = False Then
                              oActiveColumns.Add(New clsImportColum With {.ColumnID = oTempalte.ReturnCell,
                                                 .ColumnName = oTempalte.ReturnCell, .DataType = "DTO Return Value",
                                                 .Name = String.Format("{1} ({0})", oTempalte.ReturnNodeValue.Substring(oTempalte.ReturnNodeValue.LastIndexOf(".") + 1), oTempalte.DataObject)})
                         End If
                    End If

                    If String.IsNullOrEmpty(oTempalte.StatusCodeColumn) = False Then
                         oActiveColumns.Add(New clsImportColum With {.ColumnID = oTempalte.StatusCodeColumn,
                                           .ColumnName = oTempalte.StatusCodeColumn, .DataType = "Response",
                                           .Name = String.Format("Response Status Code {0}", oTempalte.Name)})
                    End If

                    If String.IsNullOrEmpty(oTempalte.StatusDescriptionColumn) = False Then
                         oActiveColumns.Add(New clsImportColum With {.ColumnID = oTempalte.StatusDescriptionColumn,
                                           .ColumnName = oTempalte.StatusDescriptionColumn, .DataType = "Response",
                                           .Name = String.Format("Response Status Description {0}", oTempalte.Name)})
                    End If

                    If String.IsNullOrEmpty(oTempalte.ReturnCellDTO) = False Then
                         oActiveColumns.Add(New clsImportColum With {.ColumnID = oTempalte.ReturnCellDTO,
                                           .ColumnName = oTempalte.ReturnCellDTO, .DataType = "Response DTO",
                                           .Name = String.Format("Response DTO {0}", oTempalte.Name)})
                    End If

                    If String.IsNullOrEmpty(oImportHeader.StatusCodeColumn) = False Then
                         If oActiveColumns.Exists(Function(n) n.ColumnName = oImportHeader.StatusCodeColumn) = False Then
                              oActiveColumns.Add(New clsImportColum With {.ColumnID = oImportHeader.StatusCodeColumn,
                                           .ColumnName = oImportHeader.StatusCodeColumn, .DataType = "Response",
                                           .Name = String.Format("Final Response Status Code {0}", oImportHeader.Name)})
                         End If
                    End If

                    If String.IsNullOrEmpty(oImportHeader.StatusDescriptionColumn) = False Then
                         If oActiveColumns.Exists(Function(n) n.ColumnName = oImportHeader.StatusDescriptionColumn) = False Then
                              oActiveColumns.Add(New clsImportColum With {.ColumnID = oImportHeader.StatusDescriptionColumn,
                                           .ColumnName = oImportHeader.StatusDescriptionColumn, .DataType = "Response",
                                           .Name = String.Format("Final Status Description {0}", oImportHeader.Name)})
                         End If
                    End If


                    oTempalte.ImportColumns = TryCast(oActiveColumns.OrderBy(Function(n) n.No).ToList, List(Of clsImportColum))
                    oTotalActiveColumns.AddRange(oActiveColumns)
                    oTotalVariables.AddRange(oVariables)

               Next



               'pass the data to the template 
               With UcProperties1
                    .gridImportColumns.DataSource = oTotalActiveColumns
                    .gridImportColumns.RefreshDataSource()
               End With

               'pass the data to the template 
               With UcProperties1
                    .gridVariables.DataSource = oTotalVariables
                    .gridVariables.RefreshDataSource()
               End With

               Dim nCounter As Integer = 0
               For Each oTemplate In oImportHeader.Templates
                    nCounter += 1
                    oTemplate.Position = nCounter
               Next

               UcProperties1.gridColumnListAll.DataSource = oImportHeader.ImportColumns
               UcProperties1.gridColumnListAll.RefreshDataSource()
               UcProperties1.gridValidatorsAll.DataSource = oImportHeader.Validators
               UcProperties1.gridValidatorsAll.RefreshDataSource()
               UcProperties1.gridValidators.RefreshDataSource()

               LoadLogs(oImportHeader.ID)

          Catch ex As Exception
               With UcProperties1
                    '._ImportHeader.ImportColumns = Nothing
                    .gridImportColumns.DataSource = Nothing
               End With

          End Try
     End Sub

     Public Function ValidateData(poImportHeader As clsDataImportHeader) As Boolean

          If ValidateHierarchiesIsSelected() = False Then Return False

          ClearStatusColumns(poImportHeader, poImportHeader.Templates.Count = 1)
          FormatExcelSheet(poImportHeader)

          Application.DoEvents()


          Dim bHasErrors As Boolean = False
          Dim bFormatted As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim dStartDateTime As DateTime = Now

          Try

               goHTTPServer.LogEvent("User has started validation of Template", "Validation", poImportHeader.Name, poImportHeader.ID)
               SetEditExcelMode(True, poImportHeader)


               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates

                    If oTemplate IsNot Nothing Then
                         If oTemplate.Validators IsNot Nothing AndAlso oTemplate.Validators.Count > 0 Then



                              For Each oValidation In oTemplate.Validators.Where(Function(n) n.Enabled = "1").OrderBy(Function(n) n.Priority).ToList

                                   If oValidation IsNot Nothing Then

                                        goHTTPServer.LogEvent(String.Format("Validating {0}", oValidation.Comments), "Validation", poImportHeader.Name, poImportHeader.ID)


                                        Dim sTemplateDTO As String = oTemplate.DTOObjectFormated
                                        sTemplateDTO = Replace(sTemplateDTO, vbNewLine, String.Empty)

                                        'validate that the specified worksheet exists
                                        With spreadsheetControl
                                             If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                                                  UpdateProgressStatus()
                                                  MsgBox(String.Format("Cannot find worksheet {0}", poImportHeader.WorkbookSheetName))
                                                  Return False
                                             Else
                                                  'if its found make it the active worksheet
                                                  '  .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                                             End If
                                        End With


                                        EnableCancelButton(True)

                                        mbCancel = False

                                        Dim sColumns As List(Of String) = ExtractColumnDetails(oValidation.Query)

                                        If sColumns IsNot Nothing AndAlso sColumns.Count >= 0 Then


                                             With spreadsheetControl.ActiveWorksheet


                                                  Dim nCountRows As Integer = .Rows.LastUsedIndex
                                                  Dim nCountColumns As Integer = .Columns.LastUsedIndex

                                                  If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True

                                                  If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                                                       ShowWaitDialogWithCancelCaption(String.Format("Validating {0} - {1}", oValidation.Priority, oValidation.Comments))
                                                       Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()

                                                       Select Case oValidation.ValidationType
                                                            Case "0" 'Exits

                                                                 Dim bhadError As Boolean = Exists(oValidation, spreadsheetControl, sColumns, oTemplate)
                                                                 If bhadError = True And bHasErrors = False Then bHasErrors = True

                                                            Case "1", "5", "7" 'FindAndReplace

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


                                                  End If
                                             End With

                                        End If
                                   End If

                                   If mbCancel Then
                                        goHTTPServer.LogEvent(String.Format("User requested Cancel of validation"), "Validation", poImportHeader.Name, poImportHeader.ID)
                                        Exit For
                                   End If

                              Next


                         End If

                    End If


               Next

               If bAutoColumnWidth Then spreadsheetControl.ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

               Call HideShowColumns(poImportHeader, mbHideCalulationColumns)
               Call UpdateProgressStatus()

               Return bHasErrors
          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally
               SetEditExcelMode(False, poImportHeader)
               EnableCancelButton(False)
               mbCancel = False
               Call UpdateProgressStatus()
               goHTTPServer.LogEvent(String.Format("Completed validation in {0} minutes", DateDiff(DateInterval.Minute, dStartDateTime, Now)), "Validation", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)
          End Try
     End Function

     Public Sub ImportData(poImportHeader As clsDataImportHeader)



          Dim mbIgnore As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim dStartDateTime As DateTime = Now
          Dim oErrorRows As New List(Of Integer)

          Try

               If ValidateHierarchiesIsSelected() = False Then Exit Sub
               UpdateProgressStatus("Import started, calculating effort")

               SetEditExcelMode(True, poImportHeader)

               'set focus to the excel spreadsheet
               Application.DoEvents()


               'validate that the specified worksheet exists
               With spreadsheetControl
                    If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                         UpdateProgressStatus()
                         ShowErrorForm(String.Format("Cannot find worksheet {0}", poImportHeader.WorkbookSheetName))
                         Exit Sub
                    End If
               End With

               If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                    Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()

                    goHTTPServer.LogEvent(String.Format("User has start import of template"), "Import", poImportHeader.Name, poImportHeader.ID)

                    ClearStatusColumns(poImportHeader, False)

                    With spreadsheetControl.ActiveWorksheet

                         Dim nCountRows As Integer = .Rows.LastUsedIndex
                         Dim nCountColumns As Integer = .Columns.LastUsedIndex
                         If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True


                         .Cells(String.Format("{0}{1}", poImportHeader.StatusCodeColumn, 1)).Value = "Result Code"
                         .Cells(String.Format("{0}{1}", poImportHeader.StatusDescriptionColumn, 1)).Value = "Result Description"

                         If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)

                         Dim oRangeError As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(.Columns(poImportHeader.StatusCodeColumn).Index, 1, .Columns(poImportHeader.StatusCodeColumn).Index, nCountRows)
                         oRangeError.Value = String.Empty
                         oRangeError = .Range.FromLTRB(.Columns(poImportHeader.StatusDescriptionColumn).Index, 1, .Columns(poImportHeader.StatusDescriptionColumn).Index, nCountRows)
                         oRangeError.Value = String.Empty


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


                         EnableCancelButton(True)
                         mbCancel = False

                         Dim nTotalRowsToImport As Integer = 0
                         For nRow = 2 To nCountRows + 1
                              If .Rows.Item(nRow - 1).Visible = True Then
                                   nTotalRowsToImport += 1
                              End If
                         Next

                         goHTTPServer.LogEvent(String.Format("Total of {0} found to import", nTotalRowsToImport), "Import", poImportHeader.Name, poImportHeader.ID)


                         For nRow = 2 To nCountRows + 1
                              If .Rows.Item(nRow - 1).Visible = True Then
                                   Dim oListResponses As New List(Of clsImportResponse)

                                   For Each oTemplate In poImportHeader.Templates.Where(Function(n) n.IsEnabled = True).ToList

                                        Call ShowWaitDialogWithCancelCaption(String.Format("Template {0} - {1}", oTemplate.Position, oTemplate.Name))

                                        Dim sTemplateDTO As String = oTemplate.DTOObjectFormated

                                        sTemplateDTO = RemoveCommands(sTemplateDTO)
                                        sTemplateDTO = Replace(sTemplateDTO, vbNewLine, String.Empty)

                                        sTemplateDTO = UpdateFormattingConditions(sTemplateDTO, "code", beiCodeFormat.EditValue.ToString, False)
                                        sTemplateDTO = UpdateFormattingConditions(sTemplateDTO, "translations", beiDescriptionFormat.EditValue.ToString, True)
                                        sTemplateDTO = UpdateFormattingConditions(sTemplateDTO, "description", beiDescriptionFormat.EditValue.ToString, False)



                                        Dim oDTO As String = sTemplateDTO
                                        Dim sAPIEndpoint As String = oTemplate.APIEndpoint
                                        Dim sUnusedProperties As New List(Of String)
                                        If sTemplateDTO IsNot Nothing Then

                                             If sTemplateDTO.ToString.Trim.StartsWith("{") Then

                                                  Dim jsnDTO As JObject = JObject.Parse(sTemplateDTO)

                                                  gbIgnoreArrays = oTemplate.IgnoreArray
                                                  For Each oProperty As JProperty In jsnDTO.Children
                                                       If oProperty.Value.ToString.Contains("<!") Then
                                                            If CheckChildrenNodes(oProperty, nRow) = False Then
                                                                 If oProperty.Value.ToString.Contains("@@") = False Then
                                                                      sUnusedProperties.Add(oProperty.Name)
                                                                 End If
                                                            End If
                                                       End If
                                                  Next

                                                  For Each sProperty In sUnusedProperties
                                                       jsnDTO.Remove(sProperty)
                                                  Next


                                                  oDTO = jsnDTO.ToString

                                                  'remove empty properties
                                                  oDTO = CleanJson(oDTO)

                                             Else
                                                  oDTO = sTemplateDTO
                                             End If




                                             gbIgnoreArrays = False

                                        End If

                                        Dim sQuery As String = oTemplate.UpdateQuery.ToString

                                        Dim sColumns As List(Of String) = ExtractColumnDetails(sQuery)


                                        'find all columns that are reference from the excel sheet to do the query
                                        For Each sColumn As String In sColumns
                                             Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                             sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                        Next

                                        sColumns = ExtractColumnDetails(sAPIEndpoint)
                                        For Each sColumn As String In sColumns
                                             Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                             sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                        Next

                                        oDTO = goHTTPServer.ExtractSystemVariables(oDTO)

                                        If String.IsNullOrEmpty(oDTO) = False AndAlso oTemplate.RemoveEmptyAndNull = True Then
                                             oDTO = RemoveEmptyOrNullProperties(oDTO)
                                        End If

                                        If String.IsNullOrEmpty(oTemplate.ReturnCellDTO) = False AndAlso oDTO IsNot Nothing Then
                                             .Cells(String.Format("{0}{1}", oTemplate.ReturnCellDTO, nRow)).Value = goHTTPServer.ExtractSystemVariables(oDTO.ToString)
                                        End If

                                        sColumns = ExtractColumnDetails(oDTO)
                                        For Each sColumn As String In sColumns
                                             Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                             oDTO = Replace(oDTO, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                        Next

                                        Dim oResponse As RestResponse
                                        Dim bIgnored As Boolean = False

                                        Select Case oTemplate.ImportType
                                             Case "1"

                                                  'only import not update, and only if it does not already exist
                                                  If .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.IsEmpty Then
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTO)
                                                  Else
                                                       bIgnored = True
                                                  End If

                                             Case "2"
                                                  If String.IsNullOrEmpty(.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString) Then
                                                       'add
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTO)
                                                  Else
                                                       'edit
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPut(sAPIEndpoint,
                                                                                               .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString,
                                                                                                sQuery, oDTO)

                                                       'check if its was found, if not then do an insert
                                                       If oResponse IsNot Nothing AndAlso oResponse.StatusCode = HttpStatusCode.NotFound Then

                                                            If oDTO.Contains("""id""") Then
                                                                 Dim oDTOExcludingID As String = Replace(oDTO, .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString, "")

                                                                 'Try an add
                                                                 oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTOExcludingID)
                                                            Else
                                                                 oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTO)
                                                            End If

                                                       End If


                                                  End If
                                             Case "3"

                                                  oResponse = goHTTPServer.CallWebEndpointUploadFile(sAPIEndpoint,
                                                                                            .Cells(String.Format("{0}{1}", oTemplate.EntityColumn, nRow)).Value.ToString,
                                                                                            .Cells(String.Format("{0}{1}", oTemplate.FileLocationColumn, nRow)).Value.ToString,
                                                                                            oTemplate.ReturnNodeValue)

                                             Case "4"

                                                  oResponse = goHTTPServer.CallWebEndpointUsingPut(sAPIEndpoint,
                                                                                          .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString,
                                                                                           sQuery, oDTO)
                                             Case "5"

                                                  oResponse = goHTTPServer.CallWebEndpointUsingDeleteByID(sAPIEndpoint,
                                                                                        .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString, sQuery)
                                             Case "6"

                                                  If String.IsNullOrEmpty(.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString) Then
                                                       'add
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTO)
                                                  Else
                                                       'edit
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPatch(sAPIEndpoint,
                                                                                               .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString,
                                                                                                sQuery, oDTO)

                                                       'check if its was found, if not then do an insert
                                                       If oResponse IsNot Nothing AndAlso oResponse.StatusCode = HttpStatusCode.NotFound Then

                                                            If oDTO.Contains("""id""") Then
                                                                 Dim oDTOExcludingID As String = Replace(oDTO, .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString, "")

                                                                 'Try an add
                                                                 oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTOExcludingID)
                                                            Else
                                                                 oResponse = goHTTPServer.CallWebEndpointUsingPost(sAPIEndpoint, oDTO)
                                                            End If

                                                       End If


                                                  End If
                                             Case Else
                                                  Exit Sub
                                        End Select


                                        If oResponse IsNot Nothing Or bIgnored = True Then

                                             Dim sReplyMessage As String = String.Empty
                                             Dim sStatusCode As String = ""
                                             If bIgnored = False Then
                                                  Select Case oResponse.StatusCode

                                                       Case HttpStatusCode.Created, HttpStatusCode.OK

                                                            Dim json As JObject = JObject.Parse(oResponse.Content)

                                                            oListResponses.Add(AddImportResponse(False, oTemplate.Name, oResponse.StatusCode.ToString, json.SelectToken("message").ToString))

                                                            sReplyMessage = json.SelectToken("message").ToString

                                                            Select Case oTemplate.ImportType
                                                                 Case "1", "2"

                                                                      If String.IsNullOrEmpty(oTemplate.ReturnNodeValue.ToString) = False Then
                                                                           ' If json.ContainsKey(oTemplate.ReturnNodeValue) Then
                                                                           .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value = json.SelectToken(oTemplate.ReturnNodeValue).ToString
                                                                           'End If
                                                                      End If

                                                            End Select

                                                            FormatCells(False, nRow - 1, nCountColumns)



                                                       Case HttpStatusCode.BadRequest

                                                            oErrorRows.Add(nRow)

                                                            Dim json As JObject = JObject.Parse(oResponse.Content)
                                                            oListResponses.Add(AddImportResponse(True, oTemplate.Name, oResponse.StatusCode.ToString, json.SelectToken("message").ToString))
                                                            FormatCells(True, nRow - 1, nCountColumns)

                                                            If mbIgnore = False Then
                                                                 UpdateProgressStatus()
                                                                 'located the beginning of the stack trace
                                                                 Dim sMessage As String = json.SelectToken("message").ToString


                                                                 Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, sMessage), MsgBoxStyle.AbortRetryIgnore)
                                                                      Case MsgBoxResult.Abort
                                                                           Clipboard.SetText(oDTO)
                                                                           mbCancel = True
                                                                           GoTo ExitProces
                                                                      Case MsgBoxResult.Retry
                                                                      Case MsgBoxResult.Ignore
                                                                           mbIgnore = True
                                                                 End Select
                                                            End If

                                                            sReplyMessage = json.SelectToken("message").ToString

                                                       Case HttpStatusCode.NotFound

                                                            oErrorRows.Add(nRow)
                                                            Dim json As JObject = JObject.Parse(oResponse.Content)
                                                            oListResponses.Add(AddImportResponse(True, oTemplate.Name, oResponse.StatusCode.ToString, json.SelectToken("message").ToString))
                                                            sReplyMessage = "Record Not Found"

                                                            FormatCells(True, nRow - 1, nCountColumns)
                                                       Case HttpStatusCode.Unauthorized
                                                            oErrorRows.Add(nRow)
                                                            Dim json As JObject = JObject.Parse(oResponse.Content)
                                                            oListResponses.Add(AddImportResponse(True, oTemplate.Name, oResponse.StatusCode.ToString, json.SelectToken("message").ToString))
                                                            FormatCells(True, nRow - 1, nCountColumns)
                                                       Case Else
                                                            oErrorRows.Add(nRow)
                                                            Try
                                                                 Dim json As JObject = JObject.Parse(oResponse.Content)
                                                                 sReplyMessage = json.SelectToken("message").ToString
                                                            Catch ex As Exception
                                                                 sReplyMessage = oResponse.StatusDescription
                                                            End Try


                                                            oListResponses.Add(AddImportResponse(True, oTemplate.Name, oResponse.StatusCode.ToString, sReplyMessage))
                                                            FormatCells(True, nRow - 1, nCountColumns)
                                                  End Select

                                                  sStatusCode = oResponse.StatusCode

                                                  If String.IsNullOrEmpty(oTemplate.StatusCodeColumn) = False Then .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode
                                                  If String.IsNullOrEmpty(oTemplate.StatusDescriptionColumn) = False Then .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = sReplyMessage




                                             Else
                                                  sReplyMessage = "Object already exists can cannot be created again."
                                                  oListResponses.Add(AddImportResponse(False, oTemplate.Name, "OK", sReplyMessage))
                                             End If

                                             Call ShowWaitDialogWithCancelProgress(String.Format("{5} row {0} of {1} from Template {4} with {2} - {3}", nRow - 1, nCountRows, sStatusCode, sReplyMessage, oTemplate.Name, IIf(oTemplate.ImportType = "6", "Deleted", "Processed")))

                                        Else

                                             oListResponses.Add(AddImportResponse(True, oTemplate.Name, "Unknown Error", "Unknown Error has occurred"))
                                             Call ShowWaitDialogWithCancelProgress(String.Format("Importing row {0} of {1} with {2} - {3}", nRow - 1, nCountRows, "Error", "Unknown Error has occurred"))
                                             FormatCells(True, nRow - 1, nCountColumns)

                                        End If





                                        If mbCancel Then
                                             goHTTPServer.LogEvent(String.Format("User requested Cancel of Import"), "Import", poImportHeader.Name, poImportHeader.ID)
                                             Exit For
                                        End If
                                   Next

ExitProces:

                                   If oListResponses IsNot Nothing AndAlso oListResponses.Count > 0 Then

                                        Dim sCode As String = ""
                                        Dim sDescritpion As String = ""


                                        If oListResponses.Where(Function(n) n.HasError = True).Count = 0 Then
                                             'has error
                                             sCode = "OK"
                                             sDescritpion = "Record processed."


                                        Else
                                             'no error
                                             For Each oResponse As clsImportResponse In oListResponses.Where(Function(n) n.HasError = True).ToList
                                                  sCode += String.Format("{0} = {1}, ", oResponse.TemplateName, oResponse.Code)
                                                  sDescritpion += String.Format("{0} = {1}, ", oResponse.TemplateName, oResponse.Description)
                                             Next
                                        End If

                                        .Cells(String.Format("{0}{1}", poImportHeader.StatusCodeColumn, nRow)).Value = sCode
                                        .Cells(String.Format("{0}{1}", poImportHeader.StatusDescriptionColumn, nRow)).Value = sDescritpion

                                   End If

                                   If mbCancel Then
                                        goHTTPServer.LogEvent(String.Format("User requested Cancel of Import"), "Import", poImportHeader.Name, poImportHeader.ID)
                                        Exit For
                                   End If

                              End If
                         Next
                         If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)
                    End With
               End If



          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally

               SetEditExcelMode(False, poImportHeader)
               FormatExcelSheet(poImportHeader)

               Call UpdateProgressStatus(String.Empty)
               EnableCancelButton(False)
               mbCancel = False
               goHTTPServer.LogEvent(String.Format("Completed import in {0} minutes", DateDiff(DateInterval.Minute, dStartDateTime, Now)), "Import", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)

          End Try


     End Sub

     Public Sub QueryandLoad(poImportHeader As clsDataImportHeader)




          Dim dStartDateTime As DateTime = Now
          If ValidateHierarchiesIsSelected() = False Then Exit Sub

          'set focus to the excel spreed sheet
          ' tcgTabs.SelectedTabPage = lcgSpreedsheet
          bbiQuery.Enabled = False

          Application.DoEvents()

          Dim bHasErrors As Boolean = False
          Dim bFormatted As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim oListUsedColumns As New List(Of String)
          Dim bIgnore As Boolean = False
          Try
               goHTTPServer.LogEvent("User has started Query of Template", "Query", poImportHeader.Name, poImportHeader.ID)
               SetEditExcelMode(True, poImportHeader)

               If poImportHeader.Templates.Where(Function(n) n.IsMaster = True).Count = 1 Then

                    Dim oTemplate As clsDataImportTemplateV2 = poImportHeader.Templates.Where(Function(n) n.IsMaster = True).FirstOrDefault

                    If oTemplate IsNot Nothing Then

                         Dim sEndpointSelect As String = IIf(String.IsNullOrEmpty(oTemplate.APIEndpointSelect), oTemplate.APIEndpoint, oTemplate.APIEndpointSelect)


                         If oTemplate.ImportColumns IsNot Nothing AndAlso oTemplate.ImportColumns.Count >= 0 Then

                              With spreadsheetControl

                                   If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                                        createNewExcelSheet(poImportHeader)
                                        '  .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(oTemplate.WorkbookSheetName)
                                   Else
                                        'if its found make it the active worksheet
                                        .Document.Worksheets.Remove(.Document.Worksheets(poImportHeader.WorkbookSheetName))
                                        createNewExcelSheet(poImportHeader)
                                        ' .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(oTemplate.WorkbookSheetName)
                                   End If




                                   If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                                        Dim oResponse As RestResponse
                                        Dim sNodeLoad As String = String.Empty
                                        Dim nPages As Integer = 0
                                        Dim bPaged As Boolean = False
                                        Dim bGraphQL As Boolean = False
                                        Dim nItems As Integer = 0
                                        Dim sSelectQuery As String = oTemplate.SelectQuery

                                        If oTemplate.GraphQLQuery IsNot Nothing AndAlso String.IsNullOrEmpty(oTemplate.GraphQLQuery.ToString) = False Then

                                             bGraphQL = True

                                             Dim sQuery As String = oTemplate.GraphQLQuery.ToString


                                             If sQuery.Contains("@@PAGE@@") Then
                                                  sQuery = Replace(sQuery, "@@PAGE@@", "0")
                                                  bPaged = True
                                             End If

                                             If sQuery.Contains("@@SEARCH@@") Then
                                                  If AddSelectorQueries(oTemplate, sSelectQuery) = False Then Exit Sub
                                                  sQuery = Replace(sQuery, "@@SEARCH@@", sSelectQuery)
                                             End If



                                             Call UpdateProgressStatus(String.Format("Requesting Data from Server {0} on API {1} ", goConnection._ServerAddress, sEndpointSelect, vbNewLine))

                                             oResponse = goHTTPServer.CallGraphQL("apollo-server", sQuery)
                                             sNodeLoad = String.Format("data.{0}.content", oTemplate.GraphQLRootNode)
                                             If sNodeLoad.EndsWith(".") Then
                                                  sNodeLoad = sNodeLoad.Substring(0, sNodeLoad.Length - 1)
                                             End If

                                             Try
                                                  Dim jsServerResponse1 As JObject = JObject.Parse(oResponse.Content)
                                                  nPages = CInt(jsServerResponse1.SelectToken(String.Format("data.{0}.totalPages", oTemplate.GraphQLRootNode)))
                                                  If nPages > 0 Then
                                                       bPaged = True
                                                  End If
                                             Catch ex As Exception
                                                  bPaged = False
                                             End Try


                                        Else
                                             If AddSelectorQueries(oTemplate, sSelectQuery) = False Then Exit Sub
                                             Call UpdateProgressStatus(String.Format("Requesting Data from Server {0} on API {1} ", goConnection._ServerAddress, sEndpointSelect, vbNewLine))

                                             If sEndpointSelect.StartsWith("POST:") = True Then

                                                  goHTTPServer.LogEvent($"User is using search query {sSelectQuery}", "Query", poImportHeader.Name, poImportHeader.ID)

                                                  Dim sUpdatedEndpoint As String = sEndpointSelect.Remove(0, 5).ToString.Trim

                                                  oResponse = goHTTPServer.CallWebEndpointUsingPost(sUpdatedEndpoint, String.Empty, sSelectQuery,,,, -2)
                                             Else
                                                  oResponse = goHTTPServer.CallWebEndpointUsingGet(sEndpointSelect, String.Empty, sSelectQuery,,,, -2)
                                             End If

                                             sNodeLoad = "responsePayload.content"
                                             bPaged = True

                                        End If


                                        Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
                                        If jsServerResponse IsNot Nothing Then

                                             If oResponse.StatusCode = HttpStatusCode.OK Then
                                                  Dim oObject As JToken = jsServerResponse.SelectToken(sNodeLoad)

                                                  If bGraphQL Then
                                                       nItems = CInt(jsServerResponse.SelectToken(String.Format("data.{0}.totalElements", oTemplate.GraphQLRootNode)))
                                                  Else
                                                       nItems = CInt(jsServerResponse.SelectToken("responsePayload.totalElements"))
                                                  End If

                                                  If bPaged Then
                                                       If bGraphQL Then
                                                            nPages = CInt(jsServerResponse.SelectToken(String.Format("data.{0}.totalPages", oTemplate.GraphQLRootNode)))
                                                       Else
                                                            nPages = CInt(jsServerResponse.SelectToken("responsePayload.totalPages")) - 1
                                                       End If

                                                  End If



                                                  If oObject IsNot Nothing Then
                                                       If oObject.Count > 0 Then

                                                            EnableCancelButton(True)
                                                            mbCancel = False




                                                            If .ActiveWorksheet.Rows.LastUsedIndex = 0 Then
                                                                 createNewExcelSheet(poImportHeader)
                                                            Else

                                                                 If .ActiveWorksheet.Tables.Count > 0 Then
                                                                      .ActiveWorksheet.Tables.Clear()
                                                                 End If

                                                                 If .ActiveWorksheet.Rows.LastUsedIndex > 0 Then
                                                                      .ActiveWorksheet.Rows.Remove(2, .ActiveWorksheet.Rows.LastUsedIndex)
                                                                 End If

                                                                 createNewExcelSheet(poImportHeader)

                                                            End If


                                                            Dim nCountRows As Integer = .ActiveWorksheet.Rows.LastUsedIndex
                                                            Dim nCountColumns As Integer = .ActiveWorksheet.Columns.LastUsedIndex

                                                            Dim nRow As Integer = 1
                                                            Dim nCounter As Integer = 0

                                                            For nPage = 0 To IIf(bGraphQL, nPages - 1, nPages)

                                                                 goHTTPServer.LogEvent(String.Format("Loading Page {0} of {1}", IIf(bGraphQL, nPage + 1, nPage), IIf(bGraphQL, nPages, nPages)), "Query", poImportHeader.Name, poImportHeader.ID)

                                                                 If nPage > IIf(bGraphQL, 0, 0) Then

                                                                      If bGraphQL Then


                                                                           Dim sQuery As String = oTemplate.GraphQLQuery.ToString

                                                                           If sQuery.Contains("@@PAGE@@") Then
                                                                                sQuery = Replace(sQuery, "@@PAGE@@", nPage.ToString)
                                                                                bPaged = True
                                                                           End If

                                                                           If sQuery.Contains("@@SEARCH@@") Then
                                                                                sQuery = Replace(sQuery, "@@SEARCH@@", oTemplate.SelectQuery)
                                                                           End If

                                                                           oResponse = goHTTPServer.CallGraphQL("apollo-server", sQuery)

                                                                      Else

                                                                           'call the end point to get the query results

                                                                           If sEndpointSelect.StartsWith("POST:") = True Then
                                                                                Dim sUpdatedEndpoint As String = sEndpointSelect.Remove(0, 5).ToString.Trim

                                                                                oResponse = goHTTPServer.CallWebEndpointUsingPost(sUpdatedEndpoint, String.Empty, sSelectQuery, , , nPage, -2)
                                                                           Else
                                                                                oResponse = goHTTPServer.CallWebEndpointUsingGet(sEndpointSelect, String.Empty, sSelectQuery, , , nPage, -2)

                                                                           End If

                                                                      End If

                                                                      jsServerResponse = JObject.Parse(oResponse.Content)
                                                                      oObject = jsServerResponse.SelectToken(sNodeLoad)

                                                                 End If

                                                                 If oObject IsNot Nothing Then
                                                                      For Each oRow In oObject

                                                                           Try

                                                                                If oRow IsNot Nothing Then
                                                                                     nRow += 1
                                                                                     nCounter += 1

                                                                                     Call ShowWaitDialogWithCancelProgress(String.Format("Downloading object {0}  page {1} of {2} pages for {3} Items ", nCounter, IIf(bGraphQL, nPage + 1, nPage), IIf(bGraphQL, nPages, nPages), nItems))

                                                                                     For Each ocolumn In oTemplate.ImportColumns.Where(Function(n) String.IsNullOrEmpty(n.DTOField) = False).OrderBy(Function(n) n.No).ToList

                                                                                          Try


                                                                                               If mbCancel = True Then
                                                                                                    goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                                    Exit For
                                                                                               End If

                                                                                               Dim sValue As String = String.Empty
                                                                                               Dim sColumnName As String = String.Empty

                                                                                               Dim oCell As Cell = .ActiveWorksheet.Cells(String.Format("{0}{1}", ocolumn.ColumnName, nRow))

                                                                                               'check if this is a multi node array
                                                                                               Dim sNodeName As String = String.Empty
                                                                                               If String.IsNullOrEmpty(ocolumn.Parent) = False Then
                                                                                                    sNodeName = ocolumn.Parent & "." & ocolumn.DTOField
                                                                                               Else
                                                                                                    sNodeName = ocolumn.DTOField
                                                                                               End If

                                                                                               Dim oJsn As JToken = oRow
                                                                                               Dim oNodes As List(Of String) = ocolumn.Parent.Split(".").ToList

                                                                                               Debug.Print(sNodeName)

                                                                                               Select Case ocolumn.Type
                                                                                                    Case JTokenType.Property

                                                                                                         If String.IsNullOrEmpty(ocolumn.ArrayID) = False Then

                                                                                                              Dim sRootNode As String = ""
                                                                                                              Dim sSubNodes As String = ""
                                                                                                              If ocolumn.Parent.Contains(".") Then
                                                                                                                   Dim sNodes As List(Of String) = ocolumn.Parent.Split(".").ToList
                                                                                                                   sRootNode = sNodes(0)

                                                                                                                   If sNodes.Count > 1 Then
                                                                                                                        sSubNodes = Replace(ocolumn.Parent, sRootNode & ".", "")
                                                                                                                   End If
                                                                                                              Else
                                                                                                                   sRootNode = ocolumn.Parent
                                                                                                              End If


                                                                                                              oJsn = oRow.SelectToken(sRootNode)
                                                                                                              If oJsn IsNot Nothing And String.IsNullOrEmpty(ocolumn.ArrayID) = False AndAlso oJsn.Children.Count > CInt(ocolumn.ArrayID) Then

                                                                                                                   oJsn = oJsn(CInt(ocolumn.ArrayID))

                                                                                                                   For Each sNode As String In oNodes
                                                                                                                        If sNode <> sRootNode Then
                                                                                                                             oJsn = oJsn.SelectToken(sNode)
                                                                                                                        End If

                                                                                                                        If oJsn.GetType = GetType(JArray) Then
                                                                                                                             oJsn = oJsn(0)
                                                                                                                        End If
                                                                                                                   Next

                                                                                                                   oJsn = oJsn.SelectToken(ocolumn.DTOField)
                                                                                                                   oCell.Value = oJsn.ToString


                                                                                                                   With oListUsedColumns
                                                                                                                        If .Contains(ocolumn.ColumnName) = False Then
                                                                                                                             .Add(ocolumn.ColumnName)
                                                                                                                        End If
                                                                                                                   End With

                                                                                                              End If
                                                                                                         End If

                                                                                                    Case JTokenType.Object


                                                                                                         If String.IsNullOrEmpty(ocolumn.ChildNode) = False Then
                                                                                                              sNodeName = sNodeName + "." + ocolumn.ChildNode
                                                                                                         End If

                                                                                                         oJsn = oRow.SelectToken(sNodeName)
                                                                                                         If oJsn IsNot Nothing Then
                                                                                                              oCell.Value = oJsn.ToString


                                                                                                              With oListUsedColumns
                                                                                                                   If .Contains(ocolumn.ColumnName) = False Then
                                                                                                                        .Add(ocolumn.ColumnName)
                                                                                                                   End If
                                                                                                              End With

                                                                                                         End If

                                                                                                    Case JTokenType.String

                                                                                                         If String.IsNullOrEmpty(ocolumn.ChildNode) = False Then
                                                                                                              sNodeName = sNodeName + "." + ocolumn.ChildNode
                                                                                                         End If

                                                                                                         oJsn = oRow.SelectToken(sNodeName)
                                                                                                         If oJsn IsNot Nothing Then

                                                                                                              If IsDate(oJsn.ToString) Then
                                                                                                                   oCell.NumberFormat = "@"
                                                                                                              End If

                                                                                                              oCell.Value = oJsn.ToString



                                                                                                              With oListUsedColumns
                                                                                                                   If .Contains(ocolumn.ColumnName) = False Then
                                                                                                                        .Add(ocolumn.ColumnName)
                                                                                                                   End If
                                                                                                              End With
                                                                                                         End If

                                                                                                    Case JTokenType.Array

                                                                                                         Dim sValues As String = String.Empty

                                                                                                         oJsn = oRow.SelectToken(IIf(ocolumn.Parent = String.Empty, ocolumn.DTOField, ocolumn.Parent))
                                                                                                         If oJsn IsNot Nothing Then
                                                                                                              If String.IsNullOrEmpty(ocolumn.ArrayID) = False Then
                                                                                                                   Dim nRowID As Integer = CInt(ocolumn.ArrayID)
                                                                                                                   If oJsn.Count >= nRowID + 1 Then
                                                                                                                        oJsn = oJsn(CInt(ocolumn.ArrayID))
                                                                                                                        oJsn = oJsn.SelectToken(ocolumn.DTOField)
                                                                                                                        sValues = oJsn.ToString
                                                                                                                   Else
                                                                                                                        Debug.Print("")
                                                                                                                   End If

                                                                                                              Else




                                                                                                                   For Each oChildren In oJsn.Children
                                                                                                                        If oChildren IsNot Nothing Then
                                                                                                                             Select Case oChildren.Type
                                                                                                                                  Case JTokenType.String

                                                                                                                                       sValues = sValues & String.Format("{0},", TryCast(oChildren, JValue).Value)

                                                                                                                                  Case JTokenType.Object

                                                                                                                                       Dim sUsedProperties As New List(Of clsImportColum)
                                                                                                                                       If ocolumn.ChildNode <> String.Empty Then
                                                                                                                                            Dim oCols As List(Of clsImportColum) = oTemplate.ImportColumns.Where(Function(n) n.ColumnName = ocolumn.ColumnName).OrderBy(Function(n) n.ChildNode).ToList
                                                                                                                                            If oCols IsNot Nothing AndAlso oCols.Count > 0 Then
                                                                                                                                                 For Each oC As clsImportColum In oCols
                                                                                                                                                      sUsedProperties.Add(oC)
                                                                                                                                                 Next
                                                                                                                                            End If
                                                                                                                                       Else
                                                                                                                                            sUsedProperties.Add(ocolumn)
                                                                                                                                       End If

                                                                                                                                       For Each oCol As clsImportColum In sUsedProperties


                                                                                                                                            For Each oProperty As JProperty In oChildren
                                                                                                                                                 If oProperty IsNot Nothing Then
                                                                                                                                                      If oProperty.Name = oCol.DTOField Then
                                                                                                                                                           sValues = sValues & String.Format("{0}|", oProperty.Value)
                                                                                                                                                      End If
                                                                                                                                                 End If
                                                                                                                                            Next

                                                                                                                                            'If TryCast(oChildren.SelectToken(oCol.DTOField), JProperty) IsNot Nothing Then
                                                                                                                                            '     sValues = sValues & String.Format("{0}:", TryCast(oChildren.SelectToken(oCol.DTOField), JProperty).Value)
                                                                                                                                            'End If
                                                                                                                                       Next

                                                                                                                                       'For Each oProperty As JProperty In oChildren
                                                                                                                                       '     If oProperty IsNot Nothing Then
                                                                                                                                       '          If sUsedProperties.Contains(oProperty.Name) Then
                                                                                                                                       '               sValues = sValues & String.Format("{1}:", oProperty.Name, oProperty.Value)
                                                                                                                                       '          End If
                                                                                                                                       '     End If
                                                                                                                                       'Next

                                                                                                                                       If sValues.EndsWith("|") Then
                                                                                                                                            sValues = sValues.Substring(0, sValues.Length - 1)
                                                                                                                                       End If

                                                                                                                                       sValues = sValues & ","

                                                                                                                             End Select
                                                                                                                        End If


                                                                                                                   Next
                                                                                                              End If

                                                                                                              oCell.Value = sValues



                                                                                                              With oListUsedColumns
                                                                                                                   If .Contains(ocolumn.ColumnName) = False Then
                                                                                                                        .Add(ocolumn.ColumnName)
                                                                                                                   End If
                                                                                                              End With

                                                                                                         End If

                                                                                               End Select





                                                                                               'format cell data to clean up some issues
                                                                                               If oCell.Value.ToString.Contains("[]") Then oCell.Value = Replace(oCell.Value.ToString, "[]", String.Empty).ToString


                                                                                          Catch ex As Exception
                                                                                               If bIgnore = False Then
                                                                                                    UpdateProgressStatus()
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
                                                                                UpdateProgressStatus()
                                                                                ShowErrorForm(ex)
                                                                           End Try

                                                                           If mbCancel Then
                                                                                goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                Exit For
                                                                           End If



                                                                      Next
                                                                 End If

                                                                 If mbCancel Then
                                                                      goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                      Exit For
                                                                 End If

                                                            Next

                                                            goHTTPServer.LogEvent(String.Format("Downloaded total of {0} records", nRow), "Query", poImportHeader.Name, poImportHeader.ID)


                                                            If mbCancel Then
                                                                 goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                 Exit Sub
                                                            End If


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



                                                            Dim oListReferencedColumns As New Dictionary(Of String, clsValidation)
                                                            For Each oValidation As clsValidation In oTemplate.Validators.Where(Function(n) n.Enabled = "1" Or n.Visibility = "1").OrderByDescending(Function(n) n.Priority).ToList
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
                                                                 If String.IsNullOrEmpty(oValidation.ReturnCell) = False Then
                                                                      If oListUsedColumns.Contains(oValidation.ReturnCell) = False Then
                                                                           If oListReferencedColumns.ContainsKey(oValidation.ReturnCell) = False Then
                                                                                oListReferencedColumns.Add(oValidation.ReturnCell, oValidation)
                                                                           End If
                                                                      End If
                                                                 End If

                                                            Next

                                                            If oListReferencedColumns IsNot Nothing And oListReferencedColumns.Count > 0 Then
                                                                 For Each sKey As String In oListReferencedColumns.Keys
                                                                      Dim oValidation As clsValidation = oListReferencedColumns.Item(sKey)
                                                                      Dim sDataIndex As String = ""
                                                                      Call UpdateProgressStatus(String.Format("Validating Key {0}", sKey))


                                                                      If oValidation IsNot Nothing Then

                                                                           Select Case oValidation.ValidationType

                                                                                Case "1", "5" 'FindAndReplace

                                                                                     Try
                                                                                          'check if is a array of data sources
                                                                                          If sKey.Contains(":") Then
                                                                                               Dim sValues As String() = sKey.Split(":")
                                                                                               If sValues IsNot Nothing And sValues.Count > 1 Then
                                                                                                    sKey = sValues(0)
                                                                                                    sDataIndex = sValues(1)
                                                                                               End If
                                                                                          End If

                                                                                          'If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value.ToString) = True Then
                                                                                          .ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, 1)).Value = String.Format("{0}({1})", oValidation.Comments, oValidation.Priority)
                                                                                          'End If

                                                                                          Dim sQuery As String = oValidation.Query
                                                                                          Dim sDataSource As String = ExtractColumnDataField(sQuery, IIf(String.IsNullOrEmpty(sDataIndex), sKey, String.Format("{0}:{1}", sKey, sDataIndex)))
                                                                                          Dim oQueryParameters As List(Of String) = ExtractQueryParameters(oValidation.Query)
                                                                                          Dim sRootNode As String = ExtractRootFromNodes(oValidation.ReturnNodeValue)
                                                                                          Dim oListValues As New Dictionary(Of String, String)


                                                                                          Dim oReturnNode As String = ""
                                                                                          If sKey = oValidation.ReturnCell Then
                                                                                               oReturnNode = oValidation.ReturnNodeValue

                                                                                          Else
                                                                                               oReturnNode = String.Format("{1}{0}", sDataSource, sRootNode)
                                                                                          End If

                                                                                          sQuery = String.Format("{0}==<!{1}!>", Replace(oValidation.ReturnNodeValue, sRootNode, String.Empty), oValidation.ReturnCell)

                                                                                          Dim sRequiredQuery As String = ""
                                                                                          If oValidation.Query.Contains("?") Then
                                                                                               sRequiredQuery = oValidation.Query.Substring(oValidation.Query.IndexOf("?"))
                                                                                          End If

                                                                                          goHTTPServer.LogEvent(String.Format("Reverse engineering records {0}", oValidation.Comments), "Query", poImportHeader.Name, poImportHeader.ID)


                                                                                          For nRow = 2 To .ActiveWorksheet.Rows.LastUsedIndex + 1


                                                                                               'only extract if the destination cell are empty
                                                                                               If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sKey, nRow)).Value.ToString) = True Then

                                                                                                    If mbCancel = True Then
                                                                                                         goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                                         Exit Sub
                                                                                                    End If

                                                                                                    Dim sReturnCell As String = String.Empty
                                                                                                    Dim sReturnIndex As String = String.Empty
                                                                                                    If oValidation.ReturnCell.Contains(":") Then
                                                                                                         Dim sValues As String() = oValidation.ReturnCell.Split(":")
                                                                                                         If sValues IsNot Nothing And sValues.Count >= 2 Then
                                                                                                              sReturnCell = sValues(0)
                                                                                                              sReturnIndex = sValues(1)
                                                                                                         End If
                                                                                                    Else
                                                                                                         sReturnCell = oValidation.ReturnCell
                                                                                                    End If


                                                                                                    If .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value IsNot Nothing AndAlso
                                                                                                   String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.ToString) = False Or sDataSource = "" Then



                                                                                                         Dim sNewQuery As String = sQuery
                                                                                                         Dim sColumns As List(Of String) = ExtractColumnDetails(oValidation.Query)

                                                                                                         'For Each sColumn As String In sColumns


                                                                                                         '     If sColumn <> IIf(String.IsNullOrEmpty(sDataIndex), sKey, String.Format("{0}:{1}", sKey, sDataIndex)) Then
                                                                                                         '          Dim sCellAddress As String = String.Format("{0}{1}", sReturnCell, nRow)
                                                                                                         '          If String.IsNullOrEmpty(.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim) = False Then
                                                                                                         '               sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                                                                         '          End If
                                                                                                         '     Else
                                                                                                         '          Dim sCellAddress As String = String.Format("{0}{1}", sReturnCell, nRow)
                                                                                                         '          Dim sValue = .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim
                                                                                                         '          If String.IsNullOrEmpty(sReturnIndex) = False Then
                                                                                                         '               Dim oValues As String() = sValue.Split(":")
                                                                                                         '               If oValues IsNot Nothing AndAlso oValues.Count >= CInt(sReturnIndex) Then
                                                                                                         '                    sValue = oValues(CInt(sReturnIndex - 1))
                                                                                                         '               End If
                                                                                                         '          End If
                                                                                                         '          sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", IIf(String.IsNullOrEmpty(sDataIndex), sKey, String.Format("{0}:{1}", sKey, sDataIndex))), sValue)
                                                                                                         '     End If
                                                                                                         'Next

                                                                                                         sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sReturnCell), .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.ToString.Trim)


                                                                                                         Dim sAPIendpoint As String = oValidation.APIEndpoint
                                                                                                         Dim olist As List(Of String) = ExtractColumnDetails(oValidation.APIEndpoint)
                                                                                                         For Each sColumn As String In olist
                                                                                                              Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                                                                                              sAPIendpoint = Replace(sAPIendpoint, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                                                                         Next

                                                                                                         For Each sParameter As String In oQueryParameters
                                                                                                              Dim sColCells As List(Of String) = ExtractColumnDetails(sParameter)
                                                                                                              If sColCells.Count > 0 Then
                                                                                                                   Dim sColDTOEntity As String = ExtractColumnDataField(sParameter, sColCells(0))
                                                                                                                   If sNewQuery.Contains("<!") = False Then


                                                                                                                        Dim sValue As String = String.Empty
                                                                                                                        If oListValues.ContainsKey(sNewQuery & sParameter) Then
                                                                                                                             sValue = oListValues(sNewQuery & sParameter)
                                                                                                                        Else
                                                                                                                             Call ShowWaitDialogWithCancelProgress(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

                                                                                                                             sValue = goHTTPServer.GetValueFromEndpoint(sAPIendpoint, sNewQuery & sRequiredQuery, sColDTOEntity, sRootNode)

                                                                                                                             If String.IsNullOrEmpty(sValue) = False Then
                                                                                                                                  oListValues.Add(sNewQuery & sParameter, sValue)
                                                                                                                             End If
                                                                                                                        End If

                                                                                                                        If .ActiveWorksheet.Cells(String.Format("{0}{1}", sColCells(0), nRow)).Value.IsEmpty Then
                                                                                                                             .ActiveWorksheet.Cells(String.Format("{0}{1}", sColCells(0), nRow)).Value = sValue
                                                                                                                        End If

                                                                                                                   End If
                                                                                                              End If
                                                                                                         Next
                                                                                                    End If
                                                                                               End If


                                                                                               If mbCancel Then
                                                                                                    goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                                    Exit Sub
                                                                                               End If





                                                                                          Next nRow

                                                                                     Catch ex As Exception
                                                                                          If bIgnore = False Then
                                                                                               UpdateProgressStatus()
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

                                                                                          Dim sColumnCode As String = String.Empty
                                                                                          Dim sIndex As String = String.Empty
                                                                                          If sKey.Contains(":") Then
                                                                                               Dim sValues As String() = sKey.Split(":")
                                                                                               If sValues IsNot Nothing And sValues.Count >= 2 Then
                                                                                                    sColumnCode = sValues(0)
                                                                                                    sIndex = sValues(1)
                                                                                               End If
                                                                                          Else
                                                                                               sColumnCode = sKey
                                                                                          End If

                                                                                          Dim sReturnCell As String = String.Empty
                                                                                          Dim sReturnIndex As String = String.Empty
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

                                                                                          sQuery = Replace(sQuery, sDataSource, Replace(oValidation.ReturnNodeValue, sRootNode, String.Empty))

                                                                                          Dim sRequiredQuery As String = ""
                                                                                          If oValidation.Query.Contains("?") Then
                                                                                               sRequiredQuery = oValidation.Query.Substring(oValidation.Query.IndexOf("?"))
                                                                                          End If

                                                                                          Dim oReturnNode As String = String.Format("{1}{0}", sDataSource, sRootNode)

                                                                                          goHTTPServer.LogEvent(String.Format("Reverse engineering records {0}", oValidation.Comments), "Query", poImportHeader.Name, poImportHeader.ID)


                                                                                          For nRow = 2 To .ActiveWorksheet.Rows.LastUsedIndex + 1

                                                                                               'only extract if the destination cell are empty
                                                                                               If String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value.ToString) = True Then

                                                                                                    If mbCancel = True Then
                                                                                                         goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                                         Exit Sub
                                                                                                    End If

                                                                                                    If .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value IsNot Nothing AndAlso
                                                                                                   String.IsNullOrEmpty(.ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.ToString) = False Then


                                                                                                         Dim sArray As String = String.Empty

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
                                                                                                                             Dim sValueID As String = String.Empty


                                                                                                                             Dim oQueryParameters As List(Of String) = ExtractQueryParameters(oValidation.Query)
                                                                                                                             Dim oListCellValues As List(Of String) = sValueIDs.Split("|").ToList

                                                                                                                             If sReturnIndex <> String.Empty Then
                                                                                                                                  Dim oVal As String() = sValueIDs.Split("|").ToArray
                                                                                                                                  If oVal IsNot Nothing And oVal.Count >= CInt(sReturnIndex) Then
                                                                                                                                       sValueID = oVal(CInt(sReturnIndex) - 1)
                                                                                                                                  End If
                                                                                                                             Else
                                                                                                                                  sValueID = sValueIDs
                                                                                                                             End If

                                                                                                                             Dim sNewQuery As String = sQuery
                                                                                                                             Dim sColumns As List(Of String) = ExtractColumnDetails(sNewQuery)
                                                                                                                             Dim nCount As Integer = 0

                                                                                                                             For Each sColumn As String In sColumns
                                                                                                                                  If Extractcelladdress(sColumn) <> Extractcelladdress(sKey) Then
                                                                                                                                       Dim sCellAddress As String = String.Format("{0}{1}", Extractcelladdress(sColumn), nRow)
                                                                                                                                       If String.IsNullOrEmpty(.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim) = False Then
                                                                                                                                            sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                                                                                                       End If
                                                                                                                                  Else
                                                                                                                                       Dim sCellAddress As String = String.Format("{0}{1}", sReturnCell, nRow)
                                                                                                                                       If ExtractValidatorQueryArrayPosition(sKey) <> 0 Then
                                                                                                                                            sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sColumn), oListCellValues(ExtractValidatorQueryArrayPosition(sColumn) - 1))
                                                                                                                                       Else
                                                                                                                                            sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", sKey), sValueID)
                                                                                                                                       End If

                                                                                                                                  End If
                                                                                                                             Next


                                                                                                                             Dim sAPIendpoint As String = oValidation.APIEndpoint
                                                                                                                             Dim olist As List(Of String) = ExtractColumnDetails(oValidation.APIEndpoint)
                                                                                                                             For Each sColumn As String In olist
                                                                                                                                  Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                                                                                                                  sAPIendpoint = Replace(sAPIendpoint, String.Format("<!{0}!>", sColumn), .ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim)
                                                                                                                             Next

                                                                                                                             Dim sConvertedArray As String = sValueIDs

                                                                                                                             For Each sParameter As String In oQueryParameters
                                                                                                                                  Dim sColCells As List(Of String) = ExtractColumnDetails(sParameter)
                                                                                                                                  If sColCells.Count > 0 Then
                                                                                                                                       Dim sColDTOEntity As String = ExtractColumnDataField(sParameter, sColCells(0))
                                                                                                                                       If sNewQuery.Contains("<!") = False Then


                                                                                                                                            Dim sValue As String = String.Empty
                                                                                                                                            If oListValues.ContainsKey(sNewQuery & sParameter) Then
                                                                                                                                                 sValue = oListValues(sNewQuery & sParameter)
                                                                                                                                            Else
                                                                                                                                                 Call ShowWaitDialogWithCancelProgress(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

                                                                                                                                                 sValue = goHTTPServer.GetValueFromEndpoint(sAPIendpoint, sNewQuery & sRequiredQuery, sColDTOEntity, sRootNode)

                                                                                                                                                 If String.IsNullOrEmpty(sValue) = False Then
                                                                                                                                                      oListValues.Add(sNewQuery & sParameter, sValue)
                                                                                                                                                 Else
                                                                                                                                                      oListValues.Add(sNewQuery & sParameter, "")
                                                                                                                                                      sValue = String.Empty
                                                                                                                                                 End If
                                                                                                                                            End If

                                                                                                                                            sConvertedArray = Replace(sConvertedArray, sValueID, sValue)

                                                                                                                                       End If
                                                                                                                                  End If
                                                                                                                             Next
                                                                                                                             oListofValues.Add(sConvertedArray)



                                                                                                                             'If sNewQuery.Contains("<!") = False Then
                                                                                                                             '     Dim sValue As String = String.Empty

                                                                                                                             '     If oListValues.ContainsKey(sNewQuery) Then
                                                                                                                             '          sValue = oListValues(sNewQuery)
                                                                                                                             '     Else
                                                                                                                             '          Call ShowWaitDialogWithCancelProgress(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))
                                                                                                                             '          sValue = goHTTPServer.GetValueFromEndpoint(sAPIendpoint, sNewQuery, sDataSource, sRootNode)

                                                                                                                             '          If String.IsNullOrEmpty(sValue) = False Then
                                                                                                                             '               oListValues.Add(sNewQuery, sValue)
                                                                                                                             '          End If

                                                                                                                             '     End If

                                                                                                                             '     If sReturnIndex <> String.Empty Then
                                                                                                                             '          sValue = Replace(sValueIDs, sValueID, sValue)
                                                                                                                             '     End If

                                                                                                                             '     oListofValues.Add(sValue)
                                                                                                                             'End If
                                                                                                                        End If
                                                                                                                   Next

                                                                                                              End If

                                                                                                              Dim sNewArray As String = String.Empty

                                                                                                              For Each sValue In oListofValues
                                                                                                                   sNewArray = String.Format("{0}{1},", sNewArray, sValue)
                                                                                                              Next


                                                                                                              If .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value.IsEmpty Then .ActiveWorksheet.Cells(String.Format("{0}{1}", sColumnCode, nRow)).Value = sNewArray
                                                                                                              If .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value.IsEmpty Then .ActiveWorksheet.Cells(String.Format("{0}{1}", sReturnCell, nRow)).Value = (sArray)
                                                                                                         End If
                                                                                                    End If
                                                                                               End If


                                                                                               If mbCancel Then
                                                                                                    goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)

                                                                                                    Exit Sub
                                                                                               End If

                                                                                          Next nRow

                                                                                     Catch ex As Exception
                                                                                          If bIgnore = False Then
                                                                                               UpdateProgressStatus()
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
                                                                                          Call ShowWaitDialogWithCancelProgress(String.Format("Linked Object {2} with Priority {3} - {0} of {1} objects ", nRow, .ActiveWorksheet.Rows.LastUsedIndex + 1, oValidation.Comments, oValidation.Priority))

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


                                                                                          If mbCancel Then
                                                                                               goHTTPServer.LogEvent(String.Format("User requested Cancel of Query"), "Query", poImportHeader.Name, poImportHeader.ID)
                                                                                               Exit Sub
                                                                                          End If

                                                                                     Next nRow

                                                                           End Select

                                                                      End If
                                                                 Next sKey

                                                            End If


                                                            .ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

                                                            Call UpdateProgressStatus(String.Format("Formating Sheet"))
                                                            'Hide unused columns
                                                            Call HideShowColumns(poImportHeader, True)

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

                                                       End If
                                                  End If
                                             Else
                                                  'located the beging of the stack trace
                                                  Dim nIndex As Integer = oResponse.Content.ToString.IndexOf("stackTrace")
                                                  Dim sMessage As String = oResponse.Content.ToString.Substring(0, IIf(nIndex > 0, nIndex, oResponse.Content.ToString.Length))
                                                  UpdateProgressStatus()

                                                  ShowErrorForm(String.Format("Error in select query.{0}{0}{1}{0}{0}", vbNewLine, sMessage))

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

               End If

               goHTTPServer.LogEvent(String.Format("Completed export in {0} minutes", DateDiff(DateInterval.Minute, dStartDateTime, Now)), "Query", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally
               SetEditExcelMode(False, poImportHeader)
               EnableCancelButton(False)
               mbCancel = False
               Call UpdateProgressStatus(String.Empty)
               bbiQuery.Enabled = True


          End Try

     End Sub

     Private Function AddSelectorQueries(poTemplate As clsDataImportTemplateV2, ByRef psQuery As String, Optional pbUseMemory As Boolean = False) As Boolean

          If pbUseMemory Then
               For Each pair As KeyValuePair(Of String, String) In goQueryMemory

                    psQuery = Replace(psQuery, pair.Key, pair.Value)
                    Return True
               Next
          End If


          If poTemplate.Selectors IsNot Nothing AndAlso poTemplate.Selectors.Count > 0 Then

               Dim bContainsSelector As Boolean = False
               For Each oSelector As clsSelectors In poTemplate.Selectors
                    If poTemplate.SelectQuery.Contains(oSelector.Variable) Then

                         bContainsSelector = True
                         Exit For
                    End If
               Next

               If bContainsSelector Then
                    'display selector form
                    ShowWaitDialog("Loading Queries... ")

                    Using oForm As New frmSelectorQuery
                         oForm.ListOfSelectors = poTemplate.Selectors
                         oForm.QueryText = psQuery


                         oForm.LoadQueries()
                         ShowWaitDialog()

                         Select Case oForm.ShowDialog
                              Case DialogResult.OK 'ok add
                                   psQuery = oForm.QueryText
                                   Return True
                              Case DialogResult.Cancel 'dont add
                                   Return False
                              Case DialogResult.Abort 'use to flag item to be deleted
                                   Return False

                         End Select
                    End Using
               Else
                    'no change required
                    Return True
               End If
          Else
               Return True
          End If
     End Function

     Public Sub ImportFiles(poImportHeader As clsDataImportHeader)


          Dim mbIgnore As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim dStartDateTime As DateTime = Now


          Try

               If ValidateHierarchiesIsSelected() = False Then Exit Sub
               'set focus to the excel spreed sheet
               ' tcgTabs.SelectedTabPage = lcgSpreedsheet
               Application.DoEvents()

               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates



                    Dim sTemplateDTO As String = oTemplate.DTOObjectFormated
                    sTemplateDTO = Replace(sTemplateDTO, vbNewLine, String.Empty)

                    'validate that the specified worksheet exists
                    With spreadsheetControl
                         If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                              UpdateProgressStatus()
                              MsgBox(String.Format("Cannot find worksheet {0}", poImportHeader.WorkbookSheetName))
                              Exit Sub
                         Else
                              'if its found make it the active worksheet
                              ' .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                         End If
                    End With

                    If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                         Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()
                         goHTTPServer.LogEvent(String.Format("User has start import of template"), "Import Files", poImportHeader.Name, poImportHeader.ID)

                         With spreadsheetControl.ActiveWorksheet
                              Dim nCountRows As Integer = .Rows.LastUsedIndex
                              Dim nCountColumns As Integer = .Columns.LastUsedIndex
                              If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True

                              SetEditExcelMode(True, poImportHeader)

                              .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, 1)).Value = "Result Code"
                              .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, 1)).Value = "Result Description"

                              If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)

                              Dim oRangeError As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(.Columns(oTemplate.StatusCodeColumn).Index, 1, .Columns(oTemplate.StatusCodeColumn).Index, nCountRows)
                              oRangeError.Value = String.Empty
                              oRangeError = .Range.FromLTRB(.Columns(oTemplate.StatusDescriptionColumn).Index, 1, .Columns(oTemplate.StatusDescriptionColumn).Index, nCountRows)
                              oRangeError.Value = String.Empty


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

                              EnableCancelButton(True)
                              mbCancel = False
                              Dim nTotalRowsToImport As Integer = 0
                              For nRow = 2 To nCountRows + 1
                                   If .Rows.Item(nRow - 1).Visible = True Then
                                        nTotalRowsToImport += 1
                                   End If
                              Next

                              goHTTPServer.LogEvent(String.Format("Total of {0} found to import", nTotalRowsToImport), "Import Files", poImportHeader.Name, poImportHeader.ID)
                              ShowWaitDialogWithCancelCaption(String.Format("Total of {0} found to import", nTotalRowsToImport))

                              For nRow = 2 To nCountRows + 1

                                   If .Rows.Item(nRow - 1).Visible = True Then

                                        Dim oDTO As String = sTemplateDTO

                                        Dim jsnDTO As JObject = JObject.Parse(sTemplateDTO)
                                        For Each oProperty As JProperty In jsnDTO.Children
                                             CheckChildrenNodes(oProperty, nRow)
                                        Next

                                        oDTO = jsnDTO.ToString

                                        Dim sQuery As String = oTemplate.UpdateQuery.ToString

                                        Dim sColumns As List(Of String) = ExtractColumnDetails(sQuery)


                                        'find all colunms that are reference from the excel sheet to do the query
                                        For Each sColumn As String In sColumns
                                             Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                             sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                        Next


                                        If String.IsNullOrEmpty(oTemplate.ReturnCellDTO) = False Then
                                             .Cells(String.Format("{0}{1}", oTemplate.ReturnCellDTO, nRow)).Value = oDTO.ToString
                                        End If

                                        oDTO = goHTTPServer.ExtractSystemVariables(oDTO)

                                        Dim oResponse As RestResponse

                                        Select Case oTemplate.ImportType
                                             Case "1"
                                                  oResponse = goHTTPServer.CallWebEndpointUsingPost(oTemplate.APIEndpoint, oDTO)
                                             Case "2"
                                                  If String.IsNullOrEmpty(.Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString) Then
                                                       'add
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPost(oTemplate.APIEndpoint, oDTO)
                                                  Else
                                                       'edit
                                                       oResponse = goHTTPServer.CallWebEndpointUsingPut(oTemplate.APIEndpoint,
                                                                                                      .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString,
                                                                                                       sQuery, oDTO)
                                                  End If
                                             Case "3" 'File import

                                                  oResponse = goHTTPServer.CallWebEndpointUploadFile(oTemplate.APIEndpoint,
                                                                                                 .Cells(String.Format("{0}{1}", oTemplate.EntityColumn, nRow)).Value.ToString,
                                                                                                 .Cells(String.Format("{0}{1}", oTemplate.FileLocationColumn, nRow)).Value.ToString,
                                                                                                 oTemplate.ReturnNodeValue)
                                             Case Else
                                                  Exit Sub
                                        End Select

                                        Call ShowWaitDialogWithCancelProgress(String.Format("Importing row {0} of {1} with {2}", nRow, nCountRows + 2, oResponse.StatusCode))


                                        Select Case oResponse.StatusCode

                                             Case HttpStatusCode.Created, HttpStatusCode.OK

                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  If String.IsNullOrEmpty(oTemplate.ReturnNodeValue.ToString) = False Then
                                                       ' If json.ContainsKey(oTemplate.ReturnNodeValue) Then
                                                       .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value = json.SelectToken(oTemplate.ReturnNodeValue).ToString
                                                       'End If
                                                  End If

                                                  FormatCells(False, nRow - 1, nCountColumns)

                                             Case HttpStatusCode.BadRequest

                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  If mbIgnore = False Then
                                                       UpdateProgressStatus()
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


                                                  FormatCells(True, nRow - 1, nCountColumns)

                                             Case HttpStatusCode.NotFound

                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = "Record not Found"

                                                  FormatCells(True, nRow - 1, nCountColumns)

                                             Case Else
                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  FormatCells(True, nRow - 1, nCountColumns)


                                        End Select

                                   End If

                                   If mbCancel Then
                                        goHTTPServer.LogEvent(String.Format("User requested Cancel of Import"), "Import File", poImportHeader.Name, poImportHeader.ID)
                                        Exit For
                                   End If

                              Next
                              If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)
                         End With
                    End If

               Next
          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally

               Call HideShowColumns(poImportHeader, mbHideCalulationColumns)
               SetEditExcelMode(False, poImportHeader)
               Call UpdateProgressStatus(String.Empty)
               EnableCancelButton(False)
               mbCancel = False
               goHTTPServer.LogEvent(String.Format("Completed import in {0} minutes", DateDiff(DateInterval.Minute, dStartDateTime, Now)), "Import Files", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)
          End Try


     End Sub

     Public Sub DeleteRecords(poImportHeader As clsDataImportHeader)


          Dim mbIgnore As Boolean = False
          Dim bAutoColumnWidth As Boolean = False
          Dim dStartDateTime As DateTime = Now


          Try

               If ValidateHierarchiesIsSelected() = False Then Exit Sub
               'set focus to the excel spreed sheet
               ' tcgTabs.SelectedTabPage = lcgSpreedsheet
               Application.DoEvents()

               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates

                    Dim sTemplateDTO As String = oTemplate.DTOObjectFormated
                    sTemplateDTO = Replace(sTemplateDTO, vbNewLine, String.Empty)

                    'validate that the specified worksheet exists
                    With spreadsheetControl
                         If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                              UpdateProgressStatus()
                              MsgBox(String.Format("Cannot find worksheet {0}", poImportHeader.WorkbookSheetName))
                              Exit Sub
                         Else
                              'if its found make it the active worksheet
                              ' .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                         End If
                    End With

                    If goHTTPServer.TestConnection(True, UcConnectionDetails1._Connection) Then

                         Dim eStatus As HttpStatusCode = goHTTPServer.LogIntoIAM()
                         goHTTPServer.LogEvent(String.Format("User has start import of template"), "Import Files", poImportHeader.Name, poImportHeader.ID)

                         With spreadsheetControl.ActiveWorksheet
                              Dim nCountRows As Integer = .Rows.LastUsedIndex
                              Dim nCountColumns As Integer = .Columns.LastUsedIndex
                              If nCountColumns < 20 And nCountRows < 10000 Then bAutoColumnWidth = True

                              SetEditExcelMode(True, poImportHeader)

                              .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, 1)).Value = "Result Code"
                              .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, 1)).Value = "Result Description"

                              If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)

                              Dim oRangeError As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(.Columns(oTemplate.StatusCodeColumn).Index, 1, .Columns(oTemplate.StatusCodeColumn).Index, nCountRows)
                              oRangeError.Value = String.Empty
                              oRangeError = .Range.FromLTRB(.Columns(oTemplate.StatusDescriptionColumn).Index, 1, .Columns(oTemplate.StatusDescriptionColumn).Index, nCountRows)
                              oRangeError.Value = String.Empty


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

                              EnableCancelButton(True)
                              mbCancel = False
                              Dim nTotalRowsToImport As Integer = 0
                              For nRow = 2 To nCountRows + 1
                                   If .Rows.Item(nRow - 1).Visible = True Then
                                        nTotalRowsToImport += 1
                                   End If
                              Next

                              goHTTPServer.LogEvent(String.Format("Total of {0} found to import", nTotalRowsToImport), "Import Files", poImportHeader.Name, poImportHeader.ID)
                              ShowWaitDialogWithCancelCaption(String.Format("Total of {0} found to import", nTotalRowsToImport))


                              For nRow = 2 To nCountRows + 1

                                   If .Rows.Item(nRow - 1).Visible = True Then

                                        Dim sID As String = .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value.ToString


                                        Dim sQuery As String = oTemplate.UpdateQuery.ToString

                                        Dim sColumns As List(Of String) = ExtractColumnDetails(sQuery)


                                        'find all columns that are reference from the excel sheet to do the query
                                        For Each sColumn As String In sColumns
                                             Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                             sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                        Next




                                        Dim oResponse As RestResponse
                                        oResponse = goHTTPServer.CallWebEndpointUsingDeleteByID(oTemplate.APIEndpoint, sID, sQuery)


                                        Call ShowWaitDialogWithCancelProgress(String.Format("Deleting row {0} of {1} with {2}", nRow, nCountRows + 2, oResponse.StatusCode))


                                        Select Case oResponse.StatusCode

                                             Case HttpStatusCode.OK

                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  If String.IsNullOrEmpty(oTemplate.ReturnNodeValue.ToString) = False Then
                                                       ' If json.ContainsKey(oTemplate.ReturnNodeValue) Then
                                                       .Cells(String.Format("{0}{1}", oTemplate.ReturnCell, nRow)).Value = json.SelectToken(oTemplate.ReturnNodeValue).ToString
                                                       'End If
                                                  End If

                                                  FormatCells(False, nRow - 1, nCountColumns)

                                             Case HttpStatusCode.BadRequest

                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  If mbIgnore = False Then
                                                       UpdateProgressStatus()
                                                       'located the beging of the stack trace
                                                       Dim sMessage As String = oResponse.Content.ToString.Substring(0, oResponse.Content.ToString.IndexOf("stackTrace"))

                                                       Select Case MsgBox(String.Format("Error, do you which to continue?{0}{0}{1}{0}{0}If you exit, error message will be copied to clipboard.", vbNewLine, sMessage), MsgBoxStyle.AbortRetryIgnore)
                                                            Case MsgBoxResult.Abort
                                                                 Clipboard.SetText(sID)
                                                                 Exit Sub
                                                            Case MsgBoxResult.Retry
                                                            Case MsgBoxResult.Ignore
                                                                 mbIgnore = True
                                                       End Select
                                                  End If


                                                  FormatCells(True, nRow - 1, nCountColumns)

                                             Case HttpStatusCode.NotFound

                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = "Record not Found"

                                                  FormatCells(True, nRow - 1, nCountColumns)

                                             Case Else
                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = oResponse.StatusCode.ToString
                                                  .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = json.SelectToken("message").ToString

                                                  FormatCells(True, nRow - 1, nCountColumns)


                                        End Select

                                   End If

                                   If mbCancel Then
                                        goHTTPServer.LogEvent(String.Format("User requested Cancel of Import"), "Import File", poImportHeader.Name, poImportHeader.ID)
                                        Exit For
                                   End If

                              Next
                              If bAutoColumnWidth Then .Columns.AutoFit(0, .Columns.LastUsedIndex)
                         End With
                    End If
               Next

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally

               Call HideShowColumns(poImportHeader, mbHideCalulationColumns)
               SetEditExcelMode(False, poImportHeader)
               Call UpdateProgressStatus(String.Empty)
               EnableCancelButton(False)
               mbCancel = False
               goHTTPServer.LogEvent(String.Format("Completed import in {0} minutes", DateDiff(DateInterval.Minute, dStartDateTime, Now)), "Import Files", poImportHeader.Name, poImportHeader.ID)
               LoadLogs(poImportHeader.ID)
          End Try


     End Sub

#End Region

#Region "Functions"

     Public Sub ClearStatusColumns(poImportHeader As clsDataImportHeader, pbClearlinkedData As Boolean)
          With spreadsheetControl.ActiveWorksheet

               spreadsheetControl.BeginUpdate()

               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates

                    If oTemplate IsNot Nothing Then
                         If oTemplate.Validators IsNot Nothing AndAlso oTemplate.Validators.Count > 0 Then


                              If pbClearlinkedData Then
                                   For Each oValidation In oTemplate.Validators.Where(Function(n) n.DataObject <> "Translate" And n.Enabled = True).OrderBy(Function(n) n.Priority).ToList

                                        If oValidation IsNot Nothing Then

                                             Dim oCol As clsColumnProperties = ExtractColumnProperties(oValidation.ReturnCell)
                                             For nRow As Integer = 2 To .Rows.LastUsedIndex + 1
                                                  If .Rows.Item(nRow - 1).Visible = True Then
                                                       If String.IsNullOrEmpty(oCol.CellName) = False Then
                                                            Try
                                                                 .Cells(String.Format("{0}{1}", oCol.CellName, nRow)).Value = ""
                                                            Catch ex As Exception

                                                            End Try

                                                       End If
                                                  End If
                                             Next

                                        End If
                                   Next
                              End If

                         End If

                         For nRow As Integer = 2 To .Rows.LastUsedIndex + 1
                              If .Rows.Item(nRow - 1).Visible = True Then
                                   If String.IsNullOrEmpty(oTemplate.StatusCodeColumn) = False Then
                                        .Cells(String.Format("{0}{1}", oTemplate.StatusCodeColumn, nRow)).Value = ""
                                   End If

                                   If String.IsNullOrEmpty(oTemplate.StatusDescriptionColumn) = False Then
                                        .Cells(String.Format("{0}{1}", oTemplate.StatusDescriptionColumn, nRow)).Value = ""
                                   End If

                                   If pbClearlinkedData Then
                                        If String.IsNullOrEmpty(oTemplate.ReturnCellDTO) = False Then
                                             .Cells(String.Format("{0}{1}", oTemplate.ReturnCellDTO, nRow)).Value = ""
                                        End If
                                   End If
                              End If

                         Next

                    End If

                    For nRow As Integer = 2 To .Rows.LastUsedIndex + 1
                         If .Rows.Item(nRow - 1).Visible = True Then
                              If String.IsNullOrEmpty(poImportHeader.StatusCodeColumn) = False Then
                                   .Cells(String.Format("{0}{1}", poImportHeader.StatusCodeColumn, nRow)).Value = ""
                              End If

                              If String.IsNullOrEmpty(poImportHeader.StatusDescriptionColumn) = False Then
                                   .Cells(String.Format("{0}{1}", poImportHeader.StatusDescriptionColumn, nRow)).Value = ""
                              End If
                         End If

                    Next
               Next

               spreadsheetControl.EndUpdate()

          End With

     End Sub

     Public Function AddImportResponse(pbHasError As Boolean, psTemplateName As String, psCode As String, psDescription As String) As clsImportResponse

          Dim oResponse As New clsImportResponse
          With oResponse
               .HasError = pbHasError
               .TemplateName = psTemplateName
               .Code = psCode
               .Description = psDescription
          End With

          Return oResponse

     End Function

     Public Sub APICallEvent(psRequest As RestRequest, psResponse As RestResponse)
          Try

               If Me.UcConnectionDetails1.chkEnableLoging.Checked Then

                    If psResponse IsNot Nothing Then

                         Dim sText As String = String.Empty

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

               Dim sText As String = String.Empty

               sText = sText + String.Format("Code:{0} - Message: {1} {3}Stack:{2}", poExcetpion.HResult, poExcetpion.Message, poExcetpion.StackTrace, vbNewLine) & vbNewLine & vbNewLine & vbNewLine

               txtErrors.EditValue = sText & txtErrors.EditValue & vbNewLine

          End If

          If txtErrors.EditValue.ToString.Length > EventLogHistory Then
               txtErrors.EditValue = txtErrors.EditValue.ToString.Substring(0, EventLogHistory)
          End If

     End Sub

     Public Sub createNewExcelSheet(poImportHeader As clsDataImportHeader)

          Dim oColumns As New List(Of clsImportColum)
          Dim oValidators As New List(Of clsValidation)
          For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates
               oColumns.AddRange(oTemplate.ImportColumns)
               oValidators.AddRange(oTemplate.Validators)
          Next


          If poImportHeader IsNot Nothing Then
               ValidateDataTempalte(poImportHeader)

               If String.IsNullOrEmpty(poImportHeader.WorkbookSheetName.ToString) = False Then


                    With spreadsheetControl
                         If .Document.Worksheets.Contains(poImportHeader.WorkbookSheetName) = False Then
                              .Document.Worksheets.Add(StrConv(poImportHeader.WorkbookSheetName, vbProperCase))
                              ' .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                         Else
                              'if its found make it the active worksheet
                              '.Document.Worksheets.ActiveWorksheet = .Document.Worksheets(poImportTemplate.WorkbookSheetName)
                         End If

                         FormatExcelSheet(poImportHeader)

                    End With

               End If
          End If

     End Sub

     Public Sub FormatExcelSheet(poImportHeader As clsDataImportHeader)
          Try



               With spreadsheetControl.ActiveWorksheet

                    spreadsheetControl.BeginUpdate()
                    Dim oColumnas As New List(Of String)
                    For nCol = 0 To .Columns.LastUsedIndex
                         Dim oColumn As Column = .Columns.Item(nCol)
                         Debug.Print(String.Format("ID : {0} column : {1}", nCol, oColumn.Heading))



                         Dim oColumnText As String = .Cells(String.Format("{0}1", oColumn.Heading)).Value.TextValue

                         If String.IsNullOrEmpty(oColumnText) = False Then
                              If String.IsNullOrEmpty(oColumnText) = False AndAlso oColumnas.Contains(oColumnText) = False Then
                                   oColumnas.Add(oColumnText)
                              End If

                              If poImportHeader.ImportColumns.Where(Function(n) n.ColumnName = oColumn.Heading).Count = 0 Then
                                   If .Cells(String.Format("{0}1", oColumn.Heading)).Tag Is Nothing Then
                                        .Cells(String.Format("{0}1", oColumn.Heading)).Tag = .Cells(String.Format("{0}1", oColumn.Heading)).Value.TextValue
                                        Debug.Print(.Cells(String.Format("{0}1", oColumn.Heading)).Tag)
                                   End If
                              End If
                         End If
                    Next

                    spreadsheetControl.EndUpdate()

                    ApplyExcelStyleTemplate(False)

                    spreadsheetControl.BeginUpdate()

                    For Each oImportColumn In poImportHeader.ImportColumns.OrderBy(Function(n) n.No)

                         Dim sColumnName As String = String.Empty

                         If String.IsNullOrEmpty(oImportColumn.Parent) = False Then
                              sColumnName = String.Format("{0}.{1}", oImportColumn.Parent, oImportColumn.Name)
                         Else
                              sColumnName = oImportColumn.Name
                         End If

                         If .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value <> sColumnName Then
                              'backup exsting column name
                              .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Tag = .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value


                              If oColumnas.Where(Function(value) value = sColumnName).Count > 1 Then
                                   sColumnName += oColumnas.Where(Function(value) value = sColumnName).Count.ToString
                              End If

                              Dim nAttempt As Integer = 0
RetryUpdate:

                              Try
                                   'rename
                                   .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value = String.Format("{0}{1}", sColumnName, IIf(nAttempt > 0, String.Format("({0})", nAttempt), ""))
                              Catch ex As Exception
                                   nAttempt += 1
                                   GoTo RetryUpdate
                              End Try

                         End If
                         If oImportColumn.Parent = "translations.en" Then
                              .Cells(String.Format("{0}1", oImportColumn.ColumnName)).Value = "English"
                         End If
                    Next


               End With

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally
               spreadsheetControl.EndUpdate()
               HideShowColumns(poImportHeader, True)
          End Try
     End Sub

     Public Function ExactExcelColumns(poJsonObject As Object) As List(Of clsImportColum)

          Dim oColumns As New List(Of clsImportColum)

          Try

               If poJsonObject IsNot Nothing AndAlso String.IsNullOrEmpty(poJsonObject) = False AndAlso poJsonObject.ToString.Trim.StartsWith("{") Then

                    Dim nodes As Dictionary(Of String, String) = New Dictionary(Of String, String)()
                    Dim rootObject As JObject = JObject.Parse(poJsonObject)
                    goJsonHelper.ParseJson(rootObject, nodes)

                    For Each key As String In nodes.Keys
                         ExtractColumnDetails2(nodes(key), key, oColumns)
                    Next

                    For Each oCol As clsImportColum In oColumns
                         Dim oNode = rootObject.SelectToken(IIf(oCol.Parent = String.Empty, oCol.Name, String.Format("{0}", oCol.Parent, oCol.Name)))
                         If oNode IsNot Nothing Then
                              oCol.Type = oNode.Type
                         End If
                    Next



               End If

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          Finally

          End Try
          Return oColumns
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
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Function

     Public Function ExtractColumnDetails(poObject As Object, pnJkokenType As JTokenType) As clsImportColum

          Try

               Dim sColumn As String = String.Empty
               Dim sColumnName As String = String.Empty
               Dim sVariable As String = String.Empty
               Dim sFormat As String = String.Empty

               If InStr(poObject.Value.ToString, "<!") > 0 Then
                    sColumn = (Replace(Replace(poObject.Value.ToString, "<!", String.Empty), "!>", String.Empty))
               End If

               sColumn = sColumn.Replace("[", String.Empty)
               sColumn = sColumn.Replace("]", String.Empty)
               sColumn = sColumn.Replace(" ", String.Empty)
               sColumn = sColumn.Replace("""", String.Empty)
               sColumn = sColumn.Replace(vbNewLine, String.Empty)



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
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Function

     Public Function RemoveCommands(psDTO As String) As String


          Dim oDataSource As List(Of String) = psDTO.Split(New Char() {"(", ")"}, StringSplitOptions.RemoveEmptyEntries).ToList
          Dim sReturnValue As String = ""
          Dim bIsCommand As Boolean = False
          For Each sString As String In oDataSource
               If bIsCommand = False Then
                    sReturnValue = sReturnValue & sString
                    bIsCommand = True
               Else
                    bIsCommand = False
               End If

          Next

          Return sReturnValue


     End Function

     Public Sub ExtractColumnDetails2(poObject As String, psParent As String, poListOfColumn As List(Of clsImportColum))

          Try


               Dim sColumnDataList As New List(Of String)


               If InStr(poObject, "<!") > 0 Then
                    poObject = (Replace(Replace(poObject, "<!", String.Empty), "!>", String.Empty))


                    If poObject.Contains("|") = True Then
                         sColumnDataList = poObject.Split("|").ToList
                    Else
                         sColumnDataList.Add(poObject)
                    End If

                    For Each sColumnData As String In sColumnDataList

                         Dim sColumnHeader As String = String.Empty
                         Dim sColumnFieldName As String = String.Empty
                         Dim sVariable As String = String.Empty
                         Dim sFormat As String = String.Empty
                         Dim sChild As String = String.Empty
                         Dim sParent As String = String.Empty
                         Dim sChildNode As String = String.Empty
                         Dim sCommands As String = String.Empty
                         Dim sArrayID As String = String.Empty

                         If sColumnData.Contains("(") Then
                              sCommands = sColumnData.Substring(sColumnData.IndexOf("("), sColumnData.IndexOf(")") - sColumnData.IndexOf("(") + 1)
                              sColumnData = Replace(sColumnData, sCommands, "")
                              sCommands = sCommands.Substring(sCommands.IndexOf("(") + 1, sCommands.IndexOf(")") - 1)


                              For Each sCommnad As String In sCommands.Split(",")
                                   If sCommnad IsNot Nothing Then
                                        If sCommnad.Contains("=") Then
                                             Dim oObjects() As String = sCommnad.Split("=")

                                             If oObjects(0) IsNot Nothing Then
                                                  Select Case oObjects(0).ToString
                                                       Case "parent" : sParent = oObjects(1)
                                                       Case "childnode" : sChildNode = oObjects(1)
                                                       Case "format" : sFormat = oObjects(1).ToString.ToUpper.Substring(0, 1)
                                                  End Select
                                             End If

                                        End If
                                   End If
                              Next



                         End If


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

                                   If sValues.Count > 4 Then
                                        sArrayID = sValues(4)
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
                         End If

                         Dim oColumn As New clsImportColum
                         With oColumn
                              .Name = sColumnFieldName
                              .DTOField = sColumnFieldName
                              .ColumnID = sColumnData
                              .Parent = sParent
                              .Formatted = sFormat
                              .ColumnName = sColumnHeader
                              .VariableName = sVariable
                              .Commands = sCommands
                              .Type = JTokenType.String
                              .ChildNode = sChildNode
                              .ArrayID = sArrayID
                              If String.IsNullOrEmpty(sChild) = False Then
                                   .ChildNode = sChild
                                   .Type = JTokenType.Array
                              End If

                              If String.IsNullOrEmpty(.ArrayID) = False Then
                                   .Type = JTokenType.Property
                              End If

                         End With

                         poListOfColumn.Add(oColumn)
                    Next
               End If

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Sub

     Public Sub ExtractColumnDetails3(poObject As String, psParent As String, poListOfColumn As List(Of clsImportColum))

          Try


               Dim sColumnDataList As New List(Of String)


               poObject = (Replace(Replace(poObject, "<!", String.Empty), "!>", String.Empty))


               If poObject.Contains("|") = True Then
                    sColumnDataList = poObject.Split("|").ToList
               Else
                    sColumnDataList.Add(poObject)
               End If

               For Each sColumnData As String In sColumnDataList

                    Dim sColumnHeader As String = String.Empty
                    Dim sColumnFieldName As String = String.Empty
                    Dim sVariable As String = String.Empty
                    Dim sFormat As String = String.Empty
                    Dim sChild As String = String.Empty
                    Dim sParent As String = String.Empty
                    Dim sChildNode As String = String.Empty
                    Dim sCommands As String = String.Empty
                    Dim sArrayID As String = String.Empty


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

                              If sValues.Count > 4 Then
                                   sArrayID = sValues(4)
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
                    End If

                    Dim oColumn As New clsImportColum
                    With oColumn
                         .Name = sColumnFieldName
                         .DTOField = sColumnFieldName
                         .ColumnID = sColumnData
                         .Parent = sParent
                         .Formatted = sFormat
                         .ColumnName = sColumnHeader
                         .VariableName = sVariable
                         .Commands = sCommands
                         .Type = JTokenType.String
                         .ChildNode = sChildNode
                         .ArrayID = sArrayID
                         If String.IsNullOrEmpty(sChild) = False Then
                              .ChildNode = sChild
                              .Type = JTokenType.Array
                         End If

                         If String.IsNullOrEmpty(.ArrayID) = False Then
                              .Type = JTokenType.Property
                         End If

                    End With

                    poListOfColumn.Add(oColumn)
               Next


          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Sub

     Public Function ExtractColumnDetails(psString As String, Optional pbColumnAddressOnly As Boolean = False) As List(Of String)

          Try

               Dim oDataSource As List(Of String) = psString.Split(New Char() {" ", ",", ".", ";", ControlChars.Quote, "=", "=="}, StringSplitOptions.RemoveEmptyEntries).ToList

               Dim sColumns As New List(Of String)

               For Each sItem As String In oDataSource
                    Dim nStart As Integer = InStr(sItem, "<!")
                    Dim nend As Integer = InStr(sItem, "!>")

                    If nStart > 0 And nend > 0 Then

                         Dim sCellText As String = sItem.Substring(nStart + 1, nend - nStart - 2)

                         If pbColumnAddressOnly Then
                              If sCellText.Contains(":") Then
                                   sCellText = sCellText.Substring(0, sCellText.IndexOf(":"))
                              End If
                         End If
                         sColumns.Add(sCellText)


                    End If
               Next

               Return sColumns

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Function
     Public Function Extractcelladdress(psString As String) As String

          Try
               If psString.Contains(":") Then
                    psString = psString.Substring(0, psString.IndexOf(":"))
               End If

               Return psString

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Function

     Public Function ExtractValidatorQueryArrayPosition(psString As String) As Integer

          Try
               If psString.Contains(":") Then
                    psString = psString.Substring(psString.IndexOf(":") + 1)
               End If

               If IsNumeric(psString) Then
                    Return CInt(psString)
               Else
                    Return 0
               End If


          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
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
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Function

     Public Function ExtractColumnDataField(psQuery As String, psColumn As String) As String
          Try

               Dim oDataSource As List(Of String) = psQuery.Split(New Char() {" ", ",", ";", "=", "==", "?", "'", ControlChars.Quote}, StringSplitOptions.RemoveEmptyEntries).ToList

               Dim sColumns As New List(Of String)
               Dim sPreviousItem As String = String.Empty
               For Each sItem As String In oDataSource
                    sItem = Replace(sItem, ControlChars.Quote, String.Empty)
                    If sItem = String.Format("<!{0}!>", psColumn) Then
                         Return sPreviousItem
                    End If
                    sPreviousItem = sItem
               Next


          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try

          Return String.Empty

     End Function

     Public Function ExtractListOfColumnsFromString(psString As String) As List(Of clsImportColum)

          Dim oColumns As New List(Of clsImportColum)

          For Each oItem As String In ExtractColumnDetails(psString)
               If oItem IsNot Nothing Then
                    Dim sField As String = ExtractColumnDataField(psString, oItem)
                    If sField IsNot Nothing AndAlso String.IsNullOrEmpty(sField) = False Then

                         Dim oColProp As clsColumnProperties = ExtractColumnProperties(oItem)

                         oColumns.Add(New clsImportColum With {.ColumnID = oItem, .ColumnName = oColProp.CellName, .Name = sField, .Type = JTokenType.String, .Parent = String.Empty})
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



     Private Function Exists(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String), poTemplate As clsDataImportTemplateV2) As Boolean

          Dim oReturnValuesStore As New Dictionary(Of String, String)
          Dim sQuery As String = String.Empty
          Dim bHasErrors As Boolean = False

          With spreadsheetControl.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex

               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then

                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

                         'get copy of query
                         sQuery = poValidation.Query

                         'find all columns that are reference from the excel sheet to do the query
                         For Each sColumn As String In psColumns
                              Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                              sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                         Next


                         'check if query has already been process to void duplicate calls
                         If oReturnValuesStore.ContainsKey(sQuery) = False Then

                              'call the end point to get the query results
                              Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint, poValidation.Headers, sQuery)
                              Dim json As JObject = JObject.Parse(oResponse.Content)
                              If json IsNot Nothing Then

                                   Dim sRecordsEffeted As String = json.SelectToken("responsePayload.numberOfElements").ToString
                                   Dim sMessage As String = String.Format("{0} records found.", sRecordsEffeted)
                                   Dim sCode As String = IIf(CInt(sRecordsEffeted) = 1, "OK", "Failed")

                                   'add item to dictionary 
                                   oReturnValuesStore.Add(sQuery, String.Format("{0}:{1}", sCode, sMessage))
                                   .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sCode
                                   .Cells(String.Format("{0}{1}", poTemplate.StatusDescriptionColumn, nRow)).Value = sMessage
                              End If
                         Else

                              'item was already validated, post same results
                              Dim sData() As String = oReturnValuesStore.Item(sQuery).ToString.Split(":")
                              .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sData(0)
                              .Cells(String.Format("{0}{1}", poTemplate.StatusDescriptionColumn, nRow)).Value = sData(1)

                         End If
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
          Dim sQuery As String = String.Empty
          Dim sAPIEndpoint As String = String.Empty
          Dim bHasErrorInLoop As Boolean = False

          With spreadsheetControl.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex

               .Cells(String.Format("{0}{1}", poValidation.ReturnCell, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)

               If poValidation.PreloadData Then

                    Dim sQuerySelect As String = ""
                    Dim sColumnFilter As String = "responsePayload,content,totalPages," & poValidation.ReturnNodeValue.Substring(poValidation.ReturnNodeValue.LastIndexOf(".") + 1) & ","

                    If poValidation.Query.Contains("?") Then
                         sQuerySelect = poValidation.Query.Substring(poValidation.Query.IndexOf("?"))
                    End If
                    Dim sProperties As New Dictionary(Of String, String)

                    For Each ocolumn In psColumns
                         sProperties.Add(ocolumn, ExtractColumnDataField(poValidation.Query, ocolumn).ToString)
                         sColumnFilter = sColumnFilter & ExtractColumnDataField(poValidation.Query, ocolumn).ToString & ","
                    Next


                    'add the Lookup type UUID to the query
                    If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                         If poValidation.LookUpType IsNot Nothing Then
                              sQuerySelect = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sQuerySelect, ControlChars.Quote)
                         End If
                    End If

                    If sQuerySelect.EndsWith(";") Then
                         sQuerySelect = sQuerySelect.Substring(0, sQuerySelect.Length - 1)
                    End If


                    Call UpdateProgressStatus(String.Format("Preloading data on {0}", poValidation.APIEndpoint))

                    'call the end point to get the query results
                    Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint, poValidation.Headers, sQuerySelect, poValidation.Sort, sColumnFilter)
                    Dim json As JObject = JObject.Parse(oResponse.Content)
                    If json IsNot Nothing Then

                         Dim sReturnValue As String = String.Empty

                         If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then

                              Dim nPages As Integer = CInt(json.SelectToken("responsePayload.totalPages"))

                              For nPage As Integer = 1 To nPages
                                   Call UpdateProgressStatus(String.Format("Preloading data on {0} page {1} of {2}", poValidation.APIEndpoint, nPage, nPages))

                                   If nPage > 1 Then

                                        'call the end point to get the query results
                                        oResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint & String.Format("?page={0}", nPage), poValidation.Headers, sQuerySelect, poValidation.Sort, sColumnFilter)
                                        json = JObject.Parse(oResponse.Content)

                                   End If
                                   If json IsNot Nothing Then
                                        Dim oOjbect As JToken = json.SelectToken("responsePayload.content")
                                        If oOjbect IsNot Nothing Then
                                             If oOjbect.Count > 0 Then
                                                  For Each jnode As JObject In oOjbect.Children

                                                       Dim sReturnKey As String = poValidation.ReturnNodeValue.Substring(poValidation.ReturnNodeValue.LastIndexOf("]") + 2)


                                                       'Extract the value
                                                       sReturnValue = jnode.SelectToken(sReturnKey).ToString

                                                       'check if there is formating
                                                       sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)

                                                       Dim sNewQuery As String = poValidation.Query

                                                       'add the Lookup type UUID to the query
                                                       If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                                                            If poValidation.LookUpType IsNot Nothing Then
                                                                 sNewQuery = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sNewQuery, ControlChars.Quote)
                                                            End If
                                                       End If

                                                       For Each oitem In sProperties

                                                            'special handling for the translation properties
                                                            If oitem.Value = "description" And jnode.ContainsKey("description") = False Then
                                                                 Dim sValue As String = ""
                                                                 Try
                                                                      sValue = jnode.SelectToken("translations.en.description").ToString
                                                                 Catch ex As Exception

                                                                 End Try

                                                                 If String.IsNullOrEmpty(sValue) = False Then
                                                                      sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", oitem.Key), sValue)
                                                                 End If

                                                            Else
                                                                 sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", oitem.Key), jnode.SelectToken(oitem.Value).ToString)
                                                            End If
                                                       Next

                                                       'add item to dictionary 
                                                       If oReturnValuesStore.ContainsKey(sNewQuery.ToUpper) = False Then
                                                            oReturnValuesStore.Add(sNewQuery.ToUpper, sReturnValue)
                                                       End If

                                                  Next

                                             End If
                                        End If
                                   End If
                              Next nPage

                         End If
                    End If
               End If

               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then
                         Dim bSourceData As Boolean = True
                         Dim bHasError As Boolean = False

                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating {3} Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1, poValidation.Comments))
                         Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
                         oCell.Value = ""
                         FormatCell(False, oCell, Color.Transparent)
                         If (oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True) Or mbForceValidate Then

                              'get copy of query
                              sQuery = poValidation.Query
                              sAPIEndpoint = poValidation.APIEndpoint
                              FormatCell(False, oCell, Color.Transparent)


                              'find all columns that are reference from the excel sheet to do the query
                              For Each sColumn As String In psColumns

                                   Dim oColumn As New List(Of clsImportColum)
                                   ExtractColumnDetails3(sColumn, "", oColumn)

                                   Dim sCellAddress As String = String.Format("{0}{1}", oColumn(0).ColumnName, nRow)


                                   If String.IsNullOrEmpty(.Cells(sCellAddress).Value.ToString.Trim) = False Then

                                        Dim sValue As String = spreadsheetControl.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim

                                        If String.IsNullOrEmpty(oColumn(0).Formatted) = False Then
                                             sValue = FormatCase(oColumn(0).Formatted, sValue)
                                        End If

                                        sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), sValue.ToString.Trim)


                                   Else
                                        bSourceData = False
                                        FormatCell(True, oCell, Color.Transparent)
                                   End If
                              Next

                              Dim olist As List(Of String) = ExtractColumnDetails(sAPIEndpoint)
                              For Each sColumn As String In olist

                                   Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)

                                   Dim oColumn As New List(Of clsImportColum)
                                   ExtractColumnDetails3(sColumn, "", oColumn)


                                   If String.IsNullOrEmpty(.Cells(sCellAddress).Value.ToString.Trim) = False Then

                                        Dim sValue As String = spreadsheetControl.ActiveWorksheet.Cells(sCellAddress).Value.ToString.Trim


                                        If String.IsNullOrEmpty(oColumn(0).Formatted) = False Then
                                             sValue = FormatCase(oColumn(0).Formatted, sValue)
                                        End If


                                        sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), sValue.ToString.Trim)

                                   End If
                                   If sAPIEndpoint.EndsWith(",") Then sAPIEndpoint = RemoveLastComma(sAPIEndpoint)
                              Next

                              'add the Lookup type UUID to the query
                              If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                                   If poValidation.LookUpType IsNot Nothing Then
                                        sQuery = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sQuery, ControlChars.Quote)
                                   End If
                              End If

                              If bSourceData = True Then

                                   'check if query has already been process to void duplicate calls
                                   If oReturnValuesStore.ContainsKey(sQuery.ToUpper) = False Then

                                        If poValidation.PreloadData = False Then
                                             'call the end p6oint to get the query results
                                             Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet(sAPIEndpoint, poValidation.Headers, sQuery.Trim, poValidation.Sort)
                                             Dim json As JObject = JObject.Parse(oResponse.Content)
                                             If json IsNot Nothing Then

                                                  Dim sReturnValue As String = String.Empty

                                                  Try

                                                       If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then
                                                            Dim oOjbect As JToken = json.SelectToken("responsePayload.content")
                                                            If oOjbect IsNot Nothing Then
                                                                 If oOjbect.Count > 0 Then

                                                                      If poValidation.ReturnNodeValue.ToString.Split("[0]").Count > 2 Then
                                                                           'get the last instance
                                                                           Dim oDataNode As String = poValidation.ReturnNodeValue.Substring(poValidation.ReturnNodeValue.ToString.LastIndexOf("[0]") + 4)
                                                                           Dim oRootNode As String = poValidation.ReturnNodeValue.Substring(0, poValidation.ReturnNodeValue.ToString.LastIndexOf("[0]"))
                                                                           Dim oRootGroup As String = oRootNode.Substring(oRootNode.LastIndexOf("[0]") + 4)
                                                                           If sQuery.ToUpper.Contains(oRootGroup.ToUpper) Then
                                                                                Dim oSearchforvalue As String = sQuery.Substring(sQuery.IndexOf(oRootGroup) + oRootGroup.Length + 1)


                                                                                Dim oSearchFilter As String = oSearchforvalue.Substring(0, oSearchforvalue.IndexOf("=="))
                                                                                oSearchforvalue = oSearchforvalue.Substring(oSearchforvalue.IndexOf("==") + 2)

                                                                                Dim retrunNodes As JArray = json.SelectToken(oRootNode)

                                                                                For Each oChild In retrunNodes.Children
                                                                                     If oChild.SelectToken(oSearchFilter).ToString = oSearchforvalue.Replace(ControlChars.Quote, "") Then
                                                                                          sReturnValue = oChild.SelectToken(oDataNode).ToString
                                                                                     End If
                                                                                Next
                                                                           Else
                                                                                sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString
                                                                           End If

                                                                      Else

                                                                           'Extract the value
                                                                           sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString
                                                                      End If



                                                                 Else
                                                                      bHasError = True
                                                                 End If
                                                            Else
                                                                 'Extract the value
                                                                 sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString
                                                            End If

                                                            'check if there is formating
                                                            sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)
                                                       Else

                                                            sReturnValue = String.Format("Error:{0}", json.SelectToken("message"))
                                                            bHasError = True
                                                            If bHasErrorInLoop = False Then
                                                                 bHasErrorInLoop = True
                                                            End If
                                                       End If

                                                  Catch ex As Exception
                                                       sReturnValue = String.Format("Error: {0}", ex.Message)
                                                       bHasError = True
                                                       If bHasErrorInLoop = False Then
                                                            bHasErrorInLoop = True
                                                       End If
                                                  End Try


                                                  'add item to dictionary 
                                                  oReturnValuesStore.Add(sQuery.ToUpper, sReturnValue.ToString)

                                                  oCell.Value = sReturnValue.ToString

                                                  If bHasError Then FormatCell(True, oCell, Color.LightSalmon)
                                                  If String.IsNullOrEmpty(sReturnValue) Then FormatCell(True, oCell, Color.Transparent)

                                             End If
                                        Else
                                             oCell.Value = ""
                                             FormatCell(True, oCell, Color.Transparent)
                                        End If
                                   Else

                                        'item was already validated, post same results
                                        oCell.Value = oReturnValuesStore.Item(sQuery.ToUpper)
                                        If String.IsNullOrEmpty(oCell.Value.ToString) Then FormatCell(True, oCell, Color.Transparent)


                                   End If
                              End If

                         End If
                    End If

                    If mbCancel Then
                         Exit For
                    End If



               Next

               If String.IsNullOrEmpty(poValidation.ReturnCell) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCell).AutoFit()


          End With

          Return bHasErrorInLoop

     End Function

     Private Function Translate(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

          Dim oReturnValuesStore As New Dictionary(Of String, String)
          Dim sQuery As String = String.Empty
          Dim bHasErrors As Boolean = False

          With spreadsheetControl.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex

               .Cells(String.Format("{0}{1}", poValidation.ReturnCell.Trim, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)

               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then
                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

                         Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
                         If oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True Then

                              'get copy of query
                              sQuery = poValidation.Query
                              Dim sTranslationtext As String = ""

                              'find all colunms that are reference from the excel sheet to do the query
                              For Each sColumn As String In psColumns
                                   Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                   sTranslationtext = .Cells(sCellAddress).Value.ToString.Trim
                                   sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim.ToLower)
                              Next

                              'check if query has already been process to void duplicate calls
                              If oReturnValuesStore.ContainsKey(sQuery) = False Then

                                   'call the end point to get the query results
                                   Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingPost(poValidation.APIEndpoint, sQuery)
                                   Dim json As JObject = JObject.Parse(oResponse.Content)
                                   If json IsNot Nothing Then

                                        If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then

                                             Dim sReturnValue As String = String.Empty

                                             Try

                                                  sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString
                                                  sReturnValue = Replace(sReturnValue, "&#39;", "'")
                                                  sReturnValue = Replace(sReturnValue, "&Amp;", "&")
                                                  sReturnValue = Replace(sReturnValue, "&amp;", "&")

                                                  'check if there is formating
                                                  sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)

                                             Catch ex As Exception
                                                  sReturnValue = String.Empty
                                                  If bHasErrors = False Then
                                                       bHasErrors = True
                                                  End If
                                             End Try

                                             'add item to dictionary 
                                             oReturnValuesStore.Add(sQuery, sReturnValue)

                                             .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = sReturnValue.ToString

                                        Else

                                             'check if there is formating
                                             sTranslationtext = FormatCase(poValidation.Formatted.ToString, sTranslationtext)


                                             .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = sTranslationtext.ToString

                                        End If

                                   End If
                              Else

                                   'item was already validated, post same results
                                   .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = oReturnValuesStore.Item(sQuery)

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

     Private Function Deletes(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

          Dim oReturnValuesStore As New Dictionary(Of String, String)
          Dim sQuery As String = String.Empty
          Dim sAPIEndpoint As String = String.Empty
          Dim bHasErrors As Boolean = False

          With spreadsheetControl.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex

               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then
                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

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
                              Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingDelete(sAPIEndpoint, sQuery)
                              Dim json As JObject = JObject.Parse(oResponse.Content)
                              'If json IsNot Nothing Then

                              '  Dim sRecordsEffeted As String = json.SelectToken("responsePayload.numberOfElements").ToString
                              '  Dim sMessage As String = String.Format("{0} records found.", sRecordsEffeted)
                              '  Dim sCode As String = IIf(CInt(sRecordsEffeted) = 1, "OK", "Failed")

                              '  'add item to dictionary 
                              oReturnValuesStore.Add(String.Format("{0}-{1}", sAPIEndpoint, sQuery), String.Format("{0}-{1}", sAPIEndpoint, sQuery))
                              If String.IsNullOrEmpty(poValidation.ReturnNodeValue) = False Then
                                   .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow)).Value = json.SelectToken(poValidation.ReturnNodeValue).ToString
                              End If
                              '  .Cells(String.Format("{0}{1}", poTemplate.StatusCodeColumn, nRow)).Value = sCode
                              '  .Cells(String.Format("{0}{1}", poTemplate.StatusDescirptionColumn, nRow)).Value = sMessage
                              'End If

                         End If
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
          Dim sQuery As String = String.Empty
          Dim sAPIEndpoint As String = String.Empty
          Dim bHasErrorsInLoop As Boolean = False

          With poSpreedSheet.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex


               .Cells(String.Format("{0}{1}", poValidation.ReturnCellWithoutProperties, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments),
                                      poValidation.APIEndpoint, String.Format("({0})", poValidation.Comments)), VbStrConv.ProperCase)


               If poValidation.PreloadData Then

                    Dim sQuerySelect As String = ""
                    Dim sColumnFilter As String = "responsePayload,content,totalPages," & poValidation.ReturnNodeValue.Substring(poValidation.ReturnNodeValue.LastIndexOf(".") + 1) & ","


                    If poValidation.Query.Contains("?") Then
                         sQuerySelect = poValidation.Query.Substring(poValidation.Query.IndexOf("?"))
                    End If
                    Dim sProperties As New Dictionary(Of String, String)
                    For Each ocolumn In psColumns
                         Dim sField As String = ExtractColumnDataField(poValidation.Query, ocolumn).ToString
                         sProperties.Add(ocolumn, sField)
                         sColumnFilter = sColumnFilter & sField & ","
                         If sField = "description" Then
                              sColumnFilter = sColumnFilter & "translations,en,"
                         End If

                    Next

                    'add the Lookup type UUID to the query
                    If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                         If poValidation.LookUpType IsNot Nothing Then
                              sQuerySelect = String.Format("lookupType.id=={2}{0}{2}", poValidation.LookUpType.Id, sQuerySelect, ControlChars.Quote)
                         End If
                    End If

                    sAPIEndpoint = poValidation.APIEndpoint

                    Call UpdateProgressStatus(String.Format("Preloading data on {0}", poValidation.APIEndpoint))

                    'call the end point to get the query results
                    Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint, poValidation.Headers, sQuerySelect, poValidation.Sort, sColumnFilter)
                    Dim json As JObject = JObject.Parse(oResponse.Content)
                    If json IsNot Nothing Then

                         Dim sReturnValue As String = String.Empty

                         If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then

                              Dim nPages As Integer = CInt(json.SelectToken("responsePayload.totalPages"))

                              For nPage As Integer = 1 To nPages
                                   Call UpdateProgressStatus(String.Format("Preloading data on {0} page {1} of {2}", poValidation.APIEndpoint, nPage, nPages))

                                   If nPage > 1 Then

                                        'call the end point to get the query results
                                        oResponse = goHTTPServer.CallWebEndpointUsingGet(poValidation.APIEndpoint & String.Format("?page={0}", nPage), poValidation.Headers, sQuerySelect, poValidation.Sort, sColumnFilter)
                                        json = JObject.Parse(oResponse.Content)

                                   End If
                                   If json IsNot Nothing Then


                                        Dim oOjbect As JToken = json.SelectToken("responsePayload.content")
                                        If oOjbect IsNot Nothing Then
                                             If oOjbect.Count > 0 Then
                                                  For Each jnode As JObject In oOjbect.Children

                                                       Dim sReturnKey As String = poValidation.ReturnNodeValue.Substring(poValidation.ReturnNodeValue.LastIndexOf("]") + 2)


                                                       'Extract the value
                                                       sReturnValue = jnode.SelectToken(sReturnKey).ToString

                                                       'check if there is formating
                                                       sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)

                                                       Dim sNewQuery As String = poValidation.Query
                                                       For Each oitem In sProperties

                                                            'special handling for the translation properties
                                                            If oitem.Value = "description" And json.ContainsKey("description") = False Then
                                                                 Dim sValue As String = ""
                                                                 Try
                                                                      sValue = jnode.SelectToken("translations.en.description").ToString
                                                                 Catch ex As Exception

                                                                 End Try

                                                                 If String.IsNullOrEmpty(sValue) = False Then
                                                                      sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", oitem.Key), sValue)
                                                                 End If

                                                            Else
                                                                 sNewQuery = Replace(sNewQuery, String.Format("<!{0}!>", oitem.Key), jnode.SelectToken(oitem.Value).ToString)
                                                            End If
                                                       Next

                                                       If String.IsNullOrEmpty(sQuerySelect) = False Then
                                                            sNewQuery = sQuerySelect.ToUpper & ";" & sNewQuery.ToUpper
                                                       End If

                                                       'add item to dictionary 
                                                       If oReturnValuesStore.ContainsKey(String.Format("{0}-{1}", sAPIEndpoint, sNewQuery.ToUpper)) = False Then
                                                            oReturnValuesStore.Add(String.Format("{0}-{1}", sAPIEndpoint, sNewQuery.ToUpper), sReturnValue)
                                                       End If

                                                  Next
                                             End If

                                        End If
                                   End If
                              Next nPage
                         End If
                    End If
               End If


               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then
                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

                         Dim oCell As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCellWithoutProperties, nRow))
                         oCell.Value = ""

                         If (oCell.Value.IsEmpty Or String.IsNullOrEmpty(oCell.Value.ToString) = True) Or mbForceValidate Then

                              'get copy of query
                              sQuery = poValidation.Query
                              sAPIEndpoint = poValidation.APIEndpoint


                              Dim sColumnListAddress As New List(Of String) '= String.Empty
                              Dim sColumnOrginalValues As String = String.Empty
                              Dim sListValues As New List(Of String)

                              Dim sReturnCell As String = String.Empty
                              Dim sReturnIndex As String = String.Empty
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

                                   Dim sColumn As String = String.Empty

                                   If sColumnfromList.Contains(":") Then
                                        sColumn = sColumnfromList.Substring(0, sColumnfromList.IndexOf(":"))
                                   Else
                                        sColumn = sColumnfromList
                                   End If

                                   Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                   sColumnOrginalValues = .Cells(sCellAddress).Value.ToString



                                   If sColumnOrginalValues.ToString.Contains(",") Then

                                        sColumnListAddress.Add(sColumnfromList)
                                        sListValues = .Cells(sCellAddress).Value.ToString.Split({","}, StringSplitOptions.RemoveEmptyEntries).ToList

                                   Else
                                        sQuery = Replace(sQuery, String.Format("<!{0}!>", sColumnfromList), .Cells(sCellAddress).Value.ToString.Trim)
                                        sListValues.Add(.Cells(sCellAddress).Value.ToString)
                                   End If

                              Next

                              Dim olist As List(Of String) = ExtractColumnDetails(sAPIEndpoint)
                              For Each sColumn As String In olist
                                   Dim sCellAddress As String = String.Format("{0}{1}", sColumn, nRow)
                                   sAPIEndpoint = Replace(sAPIEndpoint, String.Format("<!{0}!>", sColumn), .Cells(sCellAddress).Value.ToString.Trim)
                                   If sAPIEndpoint.EndsWith(",") Then sAPIEndpoint = RemoveLastComma(sAPIEndpoint)
                              Next


                              Dim oResponseRow As String = String.Empty
                              Dim bHasErrors As Boolean = False

                              For Each oValue As String In sListValues

                                   Dim sQueryUpdated As String = sQuery
                                   If String.IsNullOrEmpty(oValue.ToString.Trim) = False Then

                                        Dim sNewValue As New List(Of String) ' String.Empty

                                        For Each sString As String In sColumnListAddress

                                             If String.IsNullOrEmpty(sString) = False Then

                                                  Dim sSourceCell As String = String.Empty
                                                  Dim sSourceIndex As String = String.Empty
                                                  If sString.Contains(":") Then
                                                       Dim sValues As String() = sString.Split(":")
                                                       If sValues IsNot Nothing And sValues.Count >= 2 Then
                                                            sSourceCell = sValues(0)
                                                            sSourceIndex = sValues(1)
                                                       End If
                                                  Else
                                                       sReturnCell = sString
                                                  End If


                                                  If sSourceIndex <> String.Empty Then
                                                       Dim oArray As String() = oValue.Split("|").ToArray
                                                       If oArray IsNot Nothing AndAlso oArray.Count >= CInt(sSourceIndex) Then
                                                            sNewValue.Add(oArray(CInt(sSourceIndex) - 1))
                                                       End If
                                                  Else
                                                       sNewValue.Add(oValue)
                                                  End If


                                                  sQueryUpdated = Replace(sQueryUpdated, String.Format("<!{0}!>", sString), sNewValue(sNewValue.Count - 1).ToString.Trim)
                                             End If
                                        Next
                                        'add the Lookup type UUID to the query
                                        If poValidation.ValidationType = "5" Or poValidation.ValidationType = "6" Then
                                             If poValidation.LookUpType IsNot Nothing Then
                                                  sQueryUpdated = String.Format("lookupType.id=={2}{0}{2};{1}", poValidation.LookUpType.Id, sQueryUpdated, ControlChars.Quote)
                                             End If
                                        End If


                                        'check if query has already been process to void duplicate calls
                                        If oReturnValuesStore.ContainsKey(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated.ToUpper)) = False Then
                                             If poValidation.PreloadData = False Then

                                                  'call the end point to get the query results
                                                  Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet(sAPIEndpoint, poValidation.Headers, sQueryUpdated)
                                                  Dim json As JObject = JObject.Parse(oResponse.Content)
                                                  If json IsNot Nothing Then

                                                       Dim sReturnValue As String = String.Empty
                                                       Dim stype As String = json.SelectToken("messageType")

                                                       If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then

                                                            If poValidation.ReturnNodeValue.Contains("[") AndAlso CInt(json.SelectToken("responsePayload.numberOfElements").ToString) > 0 Then

                                                                 Try
                                                                      sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString

                                                                      'check if there is formating
                                                                      sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)

                                                                 Catch ex As Exception
                                                                      sReturnValue = String.Format("Error:{0}", ex.Message)
                                                                      If bHasErrorsInLoop = False Then bHasErrorsInLoop = True
                                                                      If bHasErrors = False Then bHasErrors = True
                                                                 End Try
                                                            Else

                                                                 Try
                                                                      If CInt(json.SelectToken("responsePayload.numberOfElements").ToString) > 0 Then
                                                                           sReturnValue = json.SelectToken(poValidation.ReturnNodeValue.ToString).ToString

                                                                           'check if there is formating
                                                                           sReturnValue = FormatCase(poValidation.Formatted.ToString, sReturnValue)
                                                                      End If

                                                                 Catch ex As Exception
                                                                      sReturnValue = String.Format("Error:{0}", ex.Message)
                                                                      If bHasErrorsInLoop = False Then bHasErrorsInLoop = True
                                                                      If bHasErrors = False Then bHasErrors = True
                                                                 End Try
                                                            End If

                                                            'add item to dictionary 
                                                            oReturnValuesStore.Add(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated.ToUpper), sReturnValue)
                                                       Else
                                                            sReturnValue = String.Format("Error:{0}", json.SelectToken("message"))
                                                            If bHasErrorsInLoop = False Then bHasErrorsInLoop = True
                                                            If bHasErrors = False Then bHasErrors = True
                                                       End If

                                                       If sReturnIndex <> String.Empty Then
                                                            sReturnValue = Replace(oValue, sNewValue(sReturnIndex - 1), sReturnValue)
                                                       End If


                                                       If String.IsNullOrEmpty(oResponseRow) Then
                                                            oResponseRow = sReturnValue
                                                       Else
                                                            oResponseRow = String.Format("{0},{1}", oResponseRow, sReturnValue)
                                                       End If


                                                  End If

                                             Else
                                                  oCell.Value = ""
                                             End If

                                        Else

                                             Dim sValueMemory As String = oReturnValuesStore.Item(String.Format("{0}-{1}", sAPIEndpoint, sQueryUpdated.ToUpper)).ToString

                                             If sReturnIndex <> String.Empty Then
                                                  sValueMemory = Replace(oValue, sNewValue(sReturnIndex - 1), sValueMemory)
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


                              If bHasErrors Then FormatCell(True, oCell, Color.LightSalmon)
                              If String.IsNullOrEmpty(oResponseRow) Then FormatCell(True, oCell, Color.Transparent)


                         End If
                    End If
                    If mbCancel Then
                         Exit For
                    End If

               Next

               If String.IsNullOrEmpty(poValidation.ReturnCellWithoutProperties) = False Then poSpreedSheet.ActiveWorksheet.Columns(poValidation.ReturnCellWithoutProperties.ToString).AutoFit()

          End With

          Return bHasErrorsInLoop

     End Function

     Private Function ArryOfValues(poValidation As clsValidation, poSpreedSheet As DevExpress.XtraSpreadsheet.SpreadsheetControl, psColumns As List(Of String)) As Boolean

          Dim bHasErrors As Boolean = False

          With poSpreedSheet.ActiveWorksheet

               Dim nCountColumns As Integer = .Columns.LastUsedIndex
               Dim nCountRows As Integer = .Rows.LastUsedIndex


               .Cells(String.Format("{0}{1}", poValidation.ReturnCell, 1)).Value = StrConv(IIf(String.IsNullOrEmpty(poValidation.Comments), poValidation.APIEndpoint, poValidation.Comments), VbStrConv.ProperCase)

               'loop though all rows in the excel sheet
               For nRow As Integer = 2 To nCountRows + 1
                    If .Rows.Item(nRow - 1).Visible = True Then
                         Call ShowWaitDialogWithCancelProgress(String.Format("Validating Priority {0} of row {1} of {2} rows ", poValidation.Priority, nRow, nCountRows + 1))

                         Dim oCellDestination As Cell = .Cells(String.Format("{0}{1}", poValidation.ReturnCell, nRow))
                         If (oCellDestination.Value.IsEmpty Or String.IsNullOrEmpty(oCellDestination.Value.ToString) = True) Or mbForceValidate Then

                              Dim sSourceColumnName As String = String.Empty
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
                                   Dim sReturnValue As String = String.Empty

                                   'find all colunms that are reference from the excel sheet to do the query
                                   If oCellSource.Value.ToString.Contains(",") Then
                                        sListValues = oCellSource.Value.ToString.Split(",").ToList
                                   End If

                                   If sListValues IsNot Nothing Then
                                        For Each sValue In sListValues
                                             If sValue.ToString.Trim.Length > 0 Then
                                                  sReturnValue = String.Format("{0}{1}{2}{1},", sReturnValue, "", sValue)
                                             End If
                                        Next
                                   End If

                                   If sReturnValue.EndsWith(",") Then
                                        sReturnValue = sReturnValue.Substring(0, sReturnValue.Length - 1)
                                   End If

                                   oCellDestination.Value = sReturnValue

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

     Private Sub FormatCells(pbError As Boolean, pnRow As Integer, pnLastColumn As Integer)

          With spreadsheetControl.ActiveWorksheet

               With .Range.FromLTRB(0, pnRow, pnLastColumn, pnRow)
                    If pbError = False Then
                         .Font.Bold = False
                         .Font.Italic = False
                         .Borders.RemoveBorders()
                         .FillColor = Color.Transparent

                    Else
                         .Font.Italic = True
                         .Font.Bold = True
                         .Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Thin)
                         .FillColor = Color.LightPink

                    End If
               End With

          End With

     End Sub

     Private Sub FormatCell(pbError As Boolean, poCell As Cell, poColour As Color)

          With poCell
               If pbError = False Then
                    .Font.Bold = False
                    .Font.Italic = False
                    .Borders.RemoveBorders()
                    .FillColor = Color.Transparent

               Else
                    .Font.Italic = True
                    .Font.Bold = True
                    .Borders.SetOutsideBorders(Color.Red, BorderLineStyle.Medium)
                    .FillColor = poColour

               End If
          End With
     End Sub

     Public Function CheckChildrenNodes(ByRef poNode As JProperty, pnRow As Integer) As Boolean

          Try


               Select Case poNode.Value.Type
                    Case JTokenType.Boolean, JTokenType.Integer, JTokenType.Float, JTokenType.Comment

                         Return ChildNodeAsString(poNode, pnRow)

                    Case JTokenType.Object
                         Dim bhasValue As Boolean = False
                         For Each oSubProperty As JProperty In poNode.Value
                              If CheckChildrenNodes(oSubProperty, pnRow) And bhasValue = False Then
                                   bhasValue = True
                              End If
                         Next
                         Return bhasValue
                    Case JTokenType.String

                         Return ChildNodeAsString(poNode, pnRow)

                    Case JTokenType.Array

                         If ChildNodeAsArray(poNode, pnRow) = False Then
                              'poNode.Value = ""
                              Return False
                         Else
                              Return True
                         End If

                    Case Else

               End Select

          Catch ex As Exception
               '     ShowErrorForm(ex)
          End Try

     End Function

     Private Function ChildNodeAsString(ByRef poNode As JProperty, pnRow As Integer) As Boolean
          Try
               Dim sColumn As List(Of String) = ExtractColumnDetails2(poNode.Value.ToString)
               If sColumn IsNot Nothing AndAlso sColumn.Count > 0 Then
                    Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumn(0))

                    If String.IsNullOrEmpty(oColumnPropeties.IndexID) Then
                         If oColumnPropeties IsNot Nothing Then

                              Dim sValue As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim

                              If oColumnPropeties.Format <> String.Empty Then
                                   sValue = FormatCase(oColumnPropeties.Format, sValue)
                              End If

                              If sValue = String.Empty Then
                                   poNode.Value = Nothing
                                   Return False
                              Else
                                   poNode.Value = sValue
                                   Return True
                              End If
                         End If

                    Else

                         Dim oToken As JToken = poNode
                         Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim
                         Dim sValues As New List(Of String)
                         If sCellData IsNot Nothing And String.IsNullOrEmpty(sCellData) = False Then

                              sValues = sCellData.Split(",").ToList

                              For Each oVal As String In sValues
                                   If oVal IsNot Nothing And String.IsNullOrEmpty(oVal) = False Then
                                        Dim oNewToken As JToken = oToken.DeepClone
                                        Dim oArray As New List(Of String)
                                        If oVal.Contains("|") Then
                                             oArray = oVal.Split("|").ToList
                                        End If

                                        If oColumnPropeties.IndexID <> String.Empty Then
                                             If oArray.Count >= CInt(oColumnPropeties.IndexID) Then

                                                  If oColumnPropeties.Format <> String.Empty Then
                                                       oArray(CInt(oColumnPropeties.IndexID) - 1) = FormatCase(oColumnPropeties.Format, oArray(CInt(oColumnPropeties.IndexID) - 1))
                                                  End If

                                                  poNode.Value = oArray(CInt(oColumnPropeties.IndexID) - 1).ToString
                                             End If
                                        Else
                                             poNode.Value = oVal
                                        End If


                                   End If
                              Next
                              Return True
                         Else
                              poNode.Parent.Remove()
                              Return False
                         End If


                    End If


               End If
          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)

          End Try
     End Function

     Private Function ChildNodeAsArray(ByRef poNode As JProperty, pnRow As Integer) As Boolean
          Try
               Dim bHasValues As Boolean = False
               For Each oChild In poNode.Value.Children
                    If oChild IsNot Nothing Then


                         Select Case oChild.Type

                              Case JTokenType.Property
                                   Return CheckChildrenNodes(poNode, pnRow)
                              Case JTokenType.String
                                   Try
                                        Dim sColumn As List(Of String) = ExtractColumnDetails(poNode.Value.ToString)
                                        If sColumn.Count > 0 Then
                                             Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumn(0))

                                             Dim sNewValue As String = String.Empty
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
                                                       If bHasValues = False Then bHasValues = True

                                                  Else
                                                       poNode.Value = "@@EMPTYARRAY@@"
                                                  End If
                                             End If
                                        End If

                                   Catch ex As Exception
                                        UpdateProgressStatus()
                                        ShowErrorForm(ex)

                                   End Try

                              Case JTokenType.Array

                              Case JTokenType.Object

                                   ' Dim oToken As JToken = oChild
                                   Dim bIsArrayOfArray As Boolean = False

                                   For Each oToken In oChild.Children
                                        Debug.Print(oToken.ToString)
                                        Debug.Print(oToken.Type.ToString)
                                        Dim sColumns As List(Of String) = ExtractColumnDetails(oToken.ToString)

                                        Select Case oToken.Type
                                             Case JTokenType.Array
                                                  bHasValues = IIf(ChildNodeAsArray(oToken, pnRow), True, bHasValues)
                                             Case Else
                                                  ' oToken = oChild.First

                                                  If sColumns.Count > 0 Then
                                                       Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumns(0))
                                                       Debug.Print(spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString)
                                                       If spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Contains("|") Or
                                                             spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Contains(",") Then
                                                            bIsArrayOfArray = True
                                                       End If
                                                  End If

                                        End Select

                                        If gbIgnoreArrays = False Then
                                             Return ChildArrayOfObject(poNode, pnRow)
                                        Else
                                             Return ChildReplaceNodeValues(poNode, pnRow)
                                        End If


                                        'If bIsArrayOfArray = False Then
                                        '     Try


                                        '          Select Case oToken.Type
                                        '               Case JTokenType.Array
                                        '                    bHasValues = IIf(ChildNodeAsArray(oToken, pnRow), True, bHasValues)
                                        '               Case JTokenType.Property
                                        '                    bHasValues = IIf(CheckChildrenNodes(oToken, pnRow), True, bHasValues)
                                        '               Case Else
                                        '          End Select

                                        '          'If oToken.Type = JTokenType.Property Then
                                        '          '     CheckChildrenNodes(oToken, pnRow)
                                        '          'Else

                                        '          '     For Each oChildlist In oToken.Children
                                        '          '          If oChildlist.Type = JTokenType.Property Then
                                        '          '               If CheckChildrenNodes(oChildlist, pnRow) And bHasValues = False Then
                                        '          '                    bHasValues = True
                                        '          '               End If
                                        '          '          Else
                                        '          '               '  CheckChildrenNodes(oToken, pnRow)
                                        '          '          End If
                                        '          '     Next

                                        '          '     If bHasValues = False Then

                                        '          '     End If
                                        '          'End If


                                        '     Catch ex As Exception
                                        '          UpdateProgressStatus()
                                        '          ShowErrorForm(ex)
                                        '     End Try
                                        'Else

                                        '     Try


                                        '          oToken = oChild.First
                                        '          If sColumns.Count > 0 Then

                                        '               Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumns(0))
                                        '               Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim

                                        '               Dim sValues As New List(Of String)
                                        '               If sCellData IsNot Nothing And String.IsNullOrEmpty(sCellData) = False Then
                                        '                    sValues = sCellData.Split(",").ToList
                                        '                    Dim nCounter As Integer = 0

                                        '                    For Each oVal As String In sValues
                                        '                         If oVal IsNot Nothing And String.IsNullOrEmpty(oVal) = False Then
                                        '                              Dim oNewToken As JToken = oToken.DeepClone
                                        '                              Dim oArray As New List(Of String)
                                        '                              If oVal.Contains("|") Then
                                        '                                   oArray = oVal.Split("|").ToList
                                        '                              End If



                                        '                              If oNewToken.Type = JTokenType.Property Then
                                        '                                   Dim oObject As JObject = oChild.DeepClone

                                        '                                   Dim oToken1 As JToken = TryCast(oObject, JToken)
                                        '                                   For Each oProperty1 As JProperty In oToken1.Children
                                        '                                        Dim sCol As List(Of String) = ExtractColumnDetails(oProperty1.ToString)
                                        '                                        If sCol.Count > 0 Then
                                        '                                             Dim oCP As clsColumnProperties = ExtractColumnProperties(sCol(0))

                                        '                                             If oCP.IndexID <> String.Empty Then
                                        '                                                  If oArray.Count >= CInt(oCP.IndexID) Then

                                        '                                                       If oColumnPropeties.Format <> String.Empty Then
                                        '                                                            oArray(CInt(oCP.IndexID) - 1) = FormatCase(oColumnPropeties.Format, oArray(CInt(oCP.IndexID) - 1))
                                        '                                                       End If

                                        '                                                       oProperty1.Value = oArray(CInt(oCP.IndexID) - 1).ToString
                                        '                                                  End If
                                        '                                             Else
                                        '                                                  oProperty1.Value = oVal
                                        '                                             End If
                                        '                                        End If
                                        '                                   Next

                                        '                                   'Dim oProperty As JProperty = TryCast(oToken1.First, JProperty)

                                        '                                   'If oProperty IsNot Nothing Then
                                        '                                   '     oProperty.Value = oVal
                                        '                                   'End If
                                        '                                   poNode.First.First.AddAfterSelf(oObject)
                                        '                                   If bHasValues = False Then bHasValues = True
                                        '                              Else
                                        '                                   For Each oProp As JProperty In oNewToken.Values
                                        '                                        Dim sCol As List(Of String) = ExtractColumnDetails(oProp)
                                        '                                        If sCol.Count > 0 Then
                                        '                                             Dim oCP As clsColumnProperties = ExtractColumnProperties(sCol(0))

                                        '                                             If oCP.IndexID <> String.Empty Then
                                        '                                                  If oArray.Count >= CInt(oCP.IndexID) Then

                                        '                                                       If oColumnPropeties.Format <> String.Empty Then
                                        '                                                            oArray(CInt(oCP.IndexID) - 1) = FormatCase(oColumnPropeties.Format, oArray(CInt(oCP.IndexID) - 1))
                                        '                                                       End If

                                        '                                                       oProp.Value = oArray(CInt(oCP.IndexID) - 1).ToString
                                        '                                                  End If
                                        '                                             Else
                                        '                                                  oProp.Value = oVal
                                        '                                             End If
                                        '                                        End If
                                        '                                   Next
                                        '                                   poNode.Values.First.AddAfterSelf(oNewToken.First)
                                        '                                   If bHasValues = False Then bHasValues = True
                                        '                              End If


                                        '                         End If
                                        '                    Next



                                        '                    'now remove the first value
                                        '                    If poNode.Values.Count > 0 Then
                                        '                         poNode.Values.First.Remove()
                                        '                    End If

                                        '                    Exit For
                                        '               Else

                                        '                    If poNode.Values.Count > 0 Then
                                        '                         poNode.Values.First.Remove()
                                        '                    End If

                                        '               End If
                                        '          End If


                                        '     Catch ex As Exception
                                        '          UpdateProgressStatus()
                                        '          ShowErrorForm(ex)

                                        '     End Try
                                        'End If

                                   Next


                              Case Else

                         End Select
                    End If
               Next

               Return bHasValues

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try
     End Function

     Private Function ChildArrayOfObject(ByRef poNode As JProperty, pnRow As Integer) As Boolean

          Dim bHasValues As Boolean = False
          Try
               'Node contains an array of an object
               Dim sColumns As List(Of String) = ExtractColumnDetails(poNode.ToString)
               Dim oColumns As New List(Of clsColumnProperties)
               Dim sValues As New Dictionary(Of String, String)
               Dim nCountOfRecords As Integer = 0
               Dim nCountOfNotNullValue As Integer = 0
               For Each sColumn In sColumns
                    Dim oColumn As clsColumnProperties = ExtractColumnProperties(sColumn)
                    Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.CellName, pnRow)).Value.ToString.Trim

                    If IsPropertyArray(poNode) Then
                         If nCountOfRecords < sCellData.Split(",", 9999, StringSplitOptions.RemoveEmptyEntries).Count Then
                              nCountOfRecords = sCellData.Split(",", 9999, StringSplitOptions.RemoveEmptyEntries).Count
                         End If
                    Else

                         If nCountOfRecords = 0 Then
                              nCountOfRecords = 1
                         End If
                    End If


                    If sValues.ContainsKey(sColumn) = False Then
                         oColumns.Add(oColumn)

                         If String.IsNullOrEmpty(sCellData) = False Then
                              nCountOfNotNullValue += 1
                         End If
                         sValues.Add(sColumn, sCellData)
                    End If

               Next

               If nCountOfNotNullValue = 0 Then
                    Return False
               End If

               Dim oNewArray As New JArray
               Dim oNewNode As JToken
               For nCount = 1 To nCountOfRecords

                    'get the first object in the array to build values
                    oNewNode = poNode.Children.First.DeepClone
                    For Each oNewNodeChild As JToken In oNewNode.Children
                         For Each oProp As JProperty In oNewNodeChild

                              If oProp.Value.Type = JTokenType.Array Then
                                   Dim bValue = ChildArrayOfObject(oProp, pnRow)
                                   If bValue = False And bHasValues = True Then
                                   Else
                                        bHasValues = bValue
                                   End If
                              Else

                                   Dim sCol As List(Of String) = ExtractColumnDetails(oProp.ToString)
                                   If sCol.Count > 0 Then

                                        Dim oCP As clsColumnProperties = ExtractColumnProperties(sCol(0))

                                        If oCP.IndexID <> String.Empty Then

                                             If sValues(oCP.SourceText).Split(",").Count > 0 AndAlso sValues(oCP.SourceText).Split(",").Count >= nCount Then
                                                  Dim oArray As New List(Of String)
                                                  If sValues(oCP.SourceText).Split(",")(nCount - 1).ToString.Contains("|") Then
                                                       oArray = sValues(oCP.SourceText).Split(",")(nCount - 1).ToString.Split("|").ToList
                                                  End If

                                                  If oArray.Count >= CInt(oCP.IndexID) Then

                                                       If oCP.Format <> String.Empty Then
                                                            oArray(CInt(oCP.IndexID) - 1) = FormatCase(oCP.Format, oArray(CInt(oCP.IndexID) - 1))
                                                       End If

                                                       oProp.Value = oArray(CInt(oCP.IndexID) - 1).ToString
                                                  End If
                                             Else
                                                  oProp.Value = sValues(oCP.SourceText).ToString.Split("|")(oCP.IndexID - 1)
                                             End If

                                        Else

                                             If sValues(oCP.SourceText).Split(",").Count > 0 AndAlso sValues(oCP.SourceText).Split(",").Count >= nCount Then
                                                  oProp.Value = sValues(oCP.SourceText).Split(",")(nCount - 1).ToString
                                             Else
                                                  oProp.Value = sValues(oCP.SourceText).ToString
                                             End If

                                        End If


                                   End If

                              End If
                         Next
                         oNewArray.Add(oNewNodeChild)
                    Next




               Next

               poNode.Value = oNewArray
               bHasValues = oNewArray.Count > 0

               Return bHasValues

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try

     End Function

     Private Function ChildReplaceNodeValues(ByRef poNode As JProperty, pnRow As Integer) As Boolean

          Dim bHasValues As Boolean = True
          Try
               'Node contains an array of an object
               Dim sCount As Integer = 0
               Dim oNewArray As New JArray
               For Each oNode As JObject In poNode.First.Children
                    Dim sColumns As List(Of String) = ExtractColumnDetails(oNode.ToString)
                    Dim sNodeAsString As String = oNode.ToString
                    Dim bHasValue As Boolean = False
                    For Each sColumn In sColumns
                         Dim oColumn As clsColumnProperties = ExtractColumnProperties(sColumn)
                         Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumn.CellName, pnRow)).Value.ToString.Trim
                         sNodeAsString = Replace(sNodeAsString, String.Format("<!{0}!>", sColumn), sCellData)

                         If String.IsNullOrEmpty(sCellData) = False Then
                              bHasValue = True
                         End If

                    Next

                    If bHasValue Then
                         oNewArray.Add(JObject.Parse(sNodeAsString))
                    End If

               Next

               poNode.Value = oNewArray

               Return bHasValues

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try

     End Function

     Public Function ExactChildObjects(psToken As JToken, ByRef poNode As JProperty, pnRow As Integer) As String

          If psToken.Children.Count = 1 Then
               For Each oChildToken As JToken In psToken.Children
                    Dim sColumns As List(Of String) = ExtractColumnDetails(oChildToken)
                    If sColumns.Count > 0 Then
                         Dim oColumnPropeties As clsColumnProperties = ExtractColumnProperties(sColumns(0))
                         Dim sCellData As String = spreadsheetControl.ActiveWorksheet.Cells(String.Format("{0}{1}", oColumnPropeties.CellName, pnRow)).Value.ToString.Trim

                         Dim sValues As New List(Of String)
                         If sCellData IsNot Nothing And String.IsNullOrEmpty(sCellData) = False Then
                              sValues = sCellData.Split(",").ToList

                              For Each oVal As String In sValues
                                   If oVal IsNot Nothing And String.IsNullOrEmpty(oVal) = False Then
                                        Dim oNewToken As JToken = oChildToken.DeepClone
                                        Dim oArray As New List(Of String)
                                        If oVal.Contains(":") Then
                                             oArray = oVal.Split(":").ToList
                                        End If

                                        For Each oProp As JProperty In oNewToken.Values
                                             Dim sCol As List(Of String) = ExtractColumnDetails(oProp)
                                             If sCol.Count > 0 Then
                                                  Dim oCP As clsColumnProperties = ExtractColumnProperties(sCol(0))

                                                  If oCP.IndexID <> String.Empty Then
                                                       If oArray.Count >= CInt(oCP.IndexID) Then

                                                            If oColumnPropeties.Format <> String.Empty Then
                                                                 oArray(CInt(oCP.IndexID) - 1) = FormatCase(oColumnPropeties.Format, oArray(CInt(oCP.IndexID) - 1))
                                                            End If

                                                            oProp.Value = oArray(CInt(oCP.IndexID) - 1).ToString
                                                       End If
                                                  Else
                                                       oProp.Value = oVal
                                                  End If
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
               Next
          Else
               For Each oChild In psToken.Children
                    ExactChildObjects(oChild, poNode, pnRow)
               Next
          End If

     End Function

     Public Sub ShowWaitDialog(Optional psStatus As String = "")
          If psStatus = String.Empty Then
               SplashScreenManager.CloseForm(False)
          Else
               SplashScreenManager.ShowForm(Me, GetType(frmWait), True, True, False)
               SplashScreenManager.Default.SetWaitFormDescription(psStatus)
          End If


          mnCounter += 1
          If mnCounter >= 100 Then
               Application.DoEvents()
               mnCounter = 0
          End If

     End Sub

     Public Sub ShowWaitDialogWithCancelProgress(Optional psStatus As String = "")
          Try

               psStatus = Replace(psStatus, vbNewLine, "")

               If psStatus = String.Empty Then
                    siStatus.Caption = ""
                    SplashScreenManager.CloseForm(False)
                    mbWaitFormShown = False
               Else
                    If mbWaitFormShown = False Then
                         SplashScreenManager.ShowForm(Me, GetType(frmWaitWithCancel), False, False, False)
                         mbWaitFormShown = True
                         SplashScreenManager.Default.SetWaitFormDescription("")
                    End If

                    siStatus.Caption = psStatus
                    siStatus.Refresh()
               End If

          Catch ex As Exception
               ShowWaitDialogWithCancelProgress("")
          End Try


     End Sub

     Public Sub ShowWaitDialogWithCancelCaption(Optional psStatus As String = "")
          Try
               If psStatus = String.Empty Then
                    siStatus.Caption = ""
                    SplashScreenManager.CloseForm(False)
                    mbWaitFormShown = False
                    siStatus.Caption = ""
                    siStatus.Refresh()
               Else
                    If mbWaitFormShown = False Then
                         SplashScreenManager.ShowForm(Me, GetType(frmWaitWithCancel), False, False, False)
                         mbWaitFormShown = True
                         SplashScreenManager.Default.SetWaitFormDescription("")
                    End If

                    SplashScreenManager.Default.SetWaitFormDescription(psStatus)

               End If

          Catch ex As Exception
               ShowWaitDialogWithCancelCaption("")
          End Try


     End Sub

     Public Sub UpdateProgressStatus(Optional psStatus As String = "", Optional mbCacelEnabled As Boolean = True)

          Try
               Call ShowWaitDialogWithCancelCaption(Replace(psStatus, vbNewLine, ""))

               Exit Sub
          Catch ex As Exception

          End Try


          Try

               Me.miButtonImage = CreateButtonImage()
               Me.miHotButtonImage = CreateHotButtonImage()


               If psStatus = String.Empty Then

                    Try
                         If moOverlayHandle IsNot Nothing Then
                              SplashScreenManager.CloseOverlayForm(moOverlayHandle)
                              moOverlayHandle = Nothing
                         End If
                    Catch ex As Exception

                    End Try
                    moOverlayHandle = Nothing

               Else


                    If moOverlayHandle Is Nothing Then
                         Try

                              If mbCacelEnabled Then
                                   moOverlayHandle = SplashScreenManager.ShowOverlayForm(LayoutControl1, customPainter:=New OverlayWindowCompositePainter(moOverlayLabel, moOverlayButton), opacity:=220)
                              Else
                                   moOverlayHandle = SplashScreenManager.ShowOverlayForm(LayoutControl1, customPainter:=New OverlayWindowCompositePainter(moOverlayLabel), opacity:=220)
                              End If

                         Catch ex As Exception

                         End Try

                         mnCounter = 0

                    End If

                    moOverlayLabel.Text = Replace(psStatus, vbNewLine, "")

               End If

               siStatus.Caption = Replace(psStatus, vbNewLine, "")

               mnCounter += 1
               Application.DoEvents()

          Catch ex As Exception

          End Try
     End Sub

     Private Sub OnCancelButtonClick()
          If mbCancel = False Then mbCancel = True

          moOverlayLabel.Text = "Cancelling"

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
                         fd.Filter = String.Format("Dynamic import Template |*{0}", gsWorkbookExtention)
                         fd.FilterIndex = 1
                         fd.FileName = txtWorkbookName.Text
                         fd.RestoreDirectory = True
                         If fd.ShowDialog() = DialogResult.OK Then

                              sPath = Path.GetDirectoryName(fd.FileName)
                              sFileName = Path.GetFileName(fd.FileName)

                              If sFileName.EndsWith(gsWorkbookExtention) = False Then
                                   sFileName = sFileName & gsWorkbookExtention
                              End If

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
                    If Path.GetExtension(sFileName) = String.Empty Then
                         sFileName = sFileName & psExtention
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

               gridWorkbook.DataSource = goOpenWorkBook.Templates.OrderBy(Function(x) x.Priority).ToList
               labVersion.Text = goOpenWorkBook.WorkbookVersion
               txtWorkbookName.EditValue = goOpenWorkBook.WorkbookName

               With goOpenWorkBook
                    If String.IsNullOrEmpty(.CodeFormat) Then
                         .CodeFormat = "U"
                    End If

                    If String.IsNullOrEmpty(.DescriptionFormat) Then
                         .DescriptionFormat = "P"
                    End If
               End With

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

               With beiCodeFormat
                    .DataBindings.Clear()
                    .DataBindings.Add(New Binding("EditValue", goOpenWorkBook, "CodeFormat"))
               End With

               With beiDescriptionFormat
                    .DataBindings.Clear()
                    .DataBindings.Add(New Binding("EditValue", goOpenWorkBook, "DescriptionFormat"))
               End With

               lueHierarchies.Properties.DataSource = goHierarchies
               gsSelectedHierarchy = goOpenWorkBook.SelectedHierarchy
               lueHierarchies.EditValue = gsSelectedHierarchy

          End If

     End Sub

     Sub SortWorksheets()

          Try
               ' sort worksheets in a workbook in ascending order
               Dim sCount As Integer, i As Integer, j As Integer

               If spreadsheetControl IsNot Nothing AndAlso spreadsheetControl.Document IsNot Nothing Then

                    With spreadsheetControl.Document

                         sCount = spreadsheetControl.Document.Worksheets.Count - 1

                         If sCount = 0 Then Exit Sub

                         For i = 0 To sCount

                              For j = i To sCount

                                   If .Worksheets(j).Name < .Worksheets(i).Name Then

                                        .Worksheets(j).Move(i)

                                   End If

                              Next j

                         Next i
                    End With
               End If

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)

          End Try


     End Sub

     Private Sub HideShowColumns(poImportHeader As clsDataImportHeader, psHide As Boolean)
          Try

               Dim oImportColumns As New List(Of clsImportColum)
               Dim oValidators As New List(Of clsValidation)
               For Each oTemplate As clsDataImportTemplateV2 In poImportHeader.Templates
                    oImportColumns.AddRange(oTemplate.ImportColumns)
                    oValidators.AddRange(oTemplate.Validators)
               Next

               If oImportColumns IsNot Nothing AndAlso oImportColumns.Count >= 0 Then

                    With spreadsheetControl

                         .BeginUpdate()
                         .ActiveWorksheet.Columns.AutoFit(0, spreadsheetControl.ActiveWorksheet.Columns.LastUsedIndex)

                         'Hide unused columns
                         For nColumn = 0 To .ActiveWorksheet.Columns.LastUsedIndex
                              Dim oColumn As Column = .ActiveWorksheet.Columns.Item(nColumn)

                              If psHide Then

                                   If oImportColumns.Where(Function(n) n.ColumnName = oColumn.Heading).Count > 0 Then
                                        Dim oImpCol As clsImportColum = oImportColumns.Where(Function(n) n.ColumnName = oColumn.Heading).First
                                        .ActiveWorksheet.Columns.Unhide(nColumn, nColumn)
                                   Else
                                        .ActiveWorksheet.Columns.Hide(nColumn, nColumn)
                                   End If

                              Else
                                   .ActiveWorksheet.Columns.Unhide(nColumn, nColumn)
                                   Debug.Print(.ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Value.TextValue)
                                   If .ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Value.ToString.StartsWith("Column") And .ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Tag IsNot Nothing Then
                                        .ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Value = .ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Tag.ToString
                                        Debug.Print(.ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Tag.ToString)
                                        Debug.Print(.ActiveWorksheet.Cells(String.Format("{0}1", oColumn.Heading)).Value.TextValue)
                                   End If
                              End If

                         Next


                    End With
               End If
          Catch ex As Exception

          Finally
               spreadsheetControl.EndUpdate()
          End Try




     End Sub

     Private Function ValidateHierarchiesIsSelected() As Boolean
          If mbContainsHierarchies And gsSelectedHierarchy Is Nothing AndAlso String.IsNullOrEmpty(gsSelectedHierarchy) = True Then
               UpdateProgressStatus()
               MsgBox("This template requires a Hierarchy to be selected.", MsgBoxStyle.OkOnly)
               Return False
          Else
               Return True
          End If
     End Function

     Private Sub ViewExcelDocument(poImportHeader As clsDataImportHeader)
          If poImportHeader IsNot Nothing Then
               UcProperties1.ImportHeader = poImportHeader
               Call ValidateDataTempalte(poImportHeader)

               With spreadsheetControl
                    If .Document.Worksheets.Contains(UcProperties1.ImportHeader.WorkbookSheetName) = False Then
                         .Document.Worksheets.Add(StrConv(UcProperties1.ImportHeader.WorkbookSheetName, vbProperCase))
                         .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1.ImportHeader.WorkbookSheetName)
                    Else
                         'if its found make it the active worksheet
                         .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1.ImportHeader.WorkbookSheetName)
                    End If
               End With

               tcgTabs.SelectedTabPage = lcgSpreedsheet

          End If
     End Sub

     Private Function VerifyHierarchSelection(poImportHeder As clsDataImportHeader) As Boolean

          If poImportHeder.IsHierarchal Then
               Dim oHierarchy As clsHierarchies = TryCast(lueHierarchies.GetSelectedDataRow, clsHierarchies)
               If oHierarchy Is Nothing Then Return False
          End If

          If poImportHeder.IsShipEntity Then
               Dim oHierarchy As clsHierarchies = TryCast(lueHierarchies.GetSelectedDataRow, clsHierarchies)
               If oHierarchy Is Nothing Then Return False
               If oHierarchy.Level <> "SHIP" Then Return False
          End If

          If poImportHeder.IsRVCEntity Then
               Dim oHierarchy As clsHierarchies = TryCast(lueHierarchies.GetSelectedDataRow, clsHierarchies)
               If oHierarchy Is Nothing Then Return False
               If oHierarchy.Level <> "REVENUECENTER" Then Return False
          End If

          Return True

     End Function

     Private Sub LoadLogs(psTemplateID As String)

          Try



               If tcgImport.SelectedTabPage.Name = lcgLogs.Name Then

                    Call UpdateProgressStatus(String.Format("Downloading Logs"))


                    Dim oResponse As RestResponse = goHTTPServer.CallWebEndpointUsingGet("metadata/v1/logs", "", String.Format("eventType=={0}{1}{0};logEntities.objectId=={0}{2}{0}", ControlChars.Quote, "DATA_IMPORT", Replace(psTemplateID, "-", "")), "utcDatetime,desc", "", 0, 2000)
                    Dim oListLogs As New List(Of clsLogs)


                    If oResponse IsNot Nothing Then
                         Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
                         If jsServerResponse IsNot Nothing Then

                              If oResponse.StatusCode = HttpStatusCode.OK Then


                                   Dim oObject As JToken = jsServerResponse.SelectToken("responsePayload.content")

                                   Call UpdateProgressStatus(String.Format("Populating logs with {0} records", oObject.Children.Count))


                                   For Each oToken As JObject In oObject.Children

                                        Dim oObjectChild As JArray = oToken.SelectToken("logEntities")

                                        oListLogs.Add(New clsLogs(oToken.SelectToken("id"), oToken.SelectToken("microserviceName"), oToken.SelectToken("microserviceVersion"),
                                                             oToken.SelectToken("microserviceInstanceId"), oToken.SelectToken("currentHierarchyId"),
                                                             oToken.SelectToken("currentLocation"), oToken.SelectToken("datetime"), oToken.SelectToken("utcDatetime"),
                                                             oToken.SelectToken("hostName"), oToken.SelectToken("hostIpAddress"), oToken.SelectToken("user"),
                                                             oToken.SelectToken("eventType"), oToken.SelectToken("logGroup"),
                                                             oObjectChild.First.SelectToken("message"), oObjectChild.First.SelectToken("className"),
                                                             oObjectChild.First.SelectToken("objectId"), oObjectChild.First.SelectToken("objectFriendlyName")))
                                   Next


                                   oObject = Nothing

                              End If
                         End If
                         jsServerResponse = Nothing
                    End If

                    oResponse = Nothing

                    gridLogs.DataSource = oListLogs
                    gridLogs.RefreshDataSource()
                    gdLogs.BestFitColumns()
               End If
          Catch ex As Exception
               gridLogs.DataSource = Nothing
               gridLogs.RefreshDataSource()
               gdLogs.BestFitColumns()
          Finally
               Call UpdateProgressStatus()
          End Try

     End Sub

     Private Sub SetEditExcelMode(pbStart As Boolean, poHeader As clsDataImportHeader, Optional pbSetFocus As Boolean = False)

          Try

               If spreadsheetControl.Document.Worksheets.Contains(poHeader.WorkbookSheetName) = False Then
                    createNewExcelSheet(poHeader)
               End If

               spreadsheetControl.Document.Worksheets.ActiveWorksheet = spreadsheetControl.Document.Worksheets(poHeader.WorkbookSheetName)

               If pbStart Then
                    tcgTabs.SelectedTabPage = LcgImportProperties
                    Me.Refresh()
                    Application.DoEvents()
                    lcgSpreedsheet.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
                    spreadsheetControl.BeginUpdate()

               Else
                    If spreadsheetControl.IsUpdateLocked Then spreadsheetControl.EndUpdate()
                    lcgSpreedsheet.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    Me.Refresh()
                    Application.DoEvents()
               End If

               If pbSetFocus Then

                    tcgTabs.SelectedTabPage = lcgSpreedsheet
                    Me.Refresh()
                    Application.DoEvents()
               End If

          Catch ex As Exception
               UpdateProgressStatus()
               ShowErrorForm(ex)
          End Try


     End Sub

     Public Sub EnableCancelButton(pbEnabled As Boolean)

          bbiCancel.Enabled = pbEnabled
          Application.DoEvents()

     End Sub

     Public Sub VerifySelectedHierarchy()
          lueHierarchies_EditValueChanged(Nothing, Nothing)
     End Sub

     Public Sub ApplyExcelStyleTemplate(pbResetFormating As Boolean)

          Try

               With spreadsheetControl.ActiveWorksheet
                    spreadsheetControl.BeginUpdate()

                    If .Columns.LastUsedIndex > 0 And .Rows.LastUsedIndex > 0 Then
                         Dim oRange As DevExpress.Spreadsheet.CellRange = .Range.FromLTRB(0, 0, .Columns.LastUsedIndex, .Rows.LastUsedIndex)

                         'get the data as a excel table
                         If .Tables.Count = 0 Then
                              Dim oTable As Table = .Tables.Add(oRange, True)
                         Else

                              If .Tables(0).Range.RightColumnIndex <> .Columns.LastUsedIndex Or .Tables(0).Range.BottomRowIndex <> .Rows.LastUsedIndex Then
                                   .Tables(0).Range = .Range.FromLTRB(0, 0, .Columns.LastUsedIndex, .Rows.LastUsedIndex)
                              End If

                         End If

                         If pbResetFormating Then
                              oRange.Borders.RemoveBorders()
                              oRange.Font.Bold = False
                              oRange.Font.Italic = False
                              oRange.FillColor = Color.Transparent
                         End If

                         oRange.AutoFitColumns
                    End If

                    If .Tables.Count > 0 Then

                         ' Access the workbook's collection of table styles.
                         Dim tableStyles As TableStyleCollection = spreadsheetControl.Document.TableStyles

                         If pbResetFormating Then
                              ' Apply the table style to the existing table.
                              .Tables(0).Style = tableStyles("TableStyleMedium6")

                         Else
                              ' Apply the table style to the existing table.
                              .Tables(0).Style = tableStyles("TableStyleMedium7")

                         End If


                    End If


               End With

          Catch ex As Exception
          Finally
               spreadsheetControl.EndUpdate()
          End Try
     End Sub

     Private Sub OpenTempalte(psFileName As String, psSafeFileName As String)
          Try
               If File.Exists(psFileName) Then

                    Dim oImportHeader As New clsDataImportHeader

                    If psFileName.EndsWith(".dit") Then
                         Dim oImportTemplate As clsDataImportTemplate = TryCast(LoadFile(psFileName), clsDataImportTemplate)
                         If oImportTemplate IsNot Nothing Then
                              With oImportHeader

                                   .Name = oImportTemplate.Name
                                   .ID = oImportTemplate.ID
                                   .Notes = oImportTemplate.Notes
                                   .Priority = oImportTemplate.Priority
                                   .Group = oImportTemplate.Group
                                   .WorkbookSheetName = oImportTemplate.WorkbookSheetName
                                   .StatusCodeColumn = oImportTemplate.StatusCodeColumn
                                   .StatusDescriptionColumn = oImportTemplate.StatusDescirptionColumn


                                   Dim oTemplate As New clsDataImportTemplateV2
                                   With oTemplate
                                        .ImportType = oImportTemplate.ImportType
                                        .Name = oImportTemplate.Name
                                        .APIEndpoint = oImportTemplate.APIEndpoint
                                        .APIEndpointSelect = oImportTemplate.APIEndpointSelect
                                        .DTOObject = oImportTemplate.DTOObject
                                        .EntityColumn = oImportTemplate.EntityColumn
                                        .GraphQLQuery = oImportTemplate.GraphQLQuery
                                        .GraphQLRootNode = oImportTemplate.GraphQLRootNode
                                        .ImportColumns = oImportTemplate.ImportColumns
                                        .ReturnCell = oImportTemplate.ReturnCell
                                        .ReturnCellDTO = oImportTemplate.ReturnCellDTO
                                        .ReturnNodeValue = oImportTemplate.ReturnNodeValue
                                        .SelectQuery = oImportTemplate.SelectQuery
                                        .StatusCodeColumn = oImportTemplate.StatusCodeColumn
                                        .StatusDescriptionColumn = oImportTemplate.StatusDescirptionColumn
                                        .UpdateQuery = oImportTemplate.UpdateQuery
                                        .Validators = oImportTemplate.Validators
                                        .Variables = oImportTemplate.Variables
                                        .Position = oImportHeader.Templates.Count + 1
                                   End With

                                   .Templates.Add(oTemplate)
                              End With
                         End If
                    Else
                         oImportHeader = TryCast(LoadFile(psFileName), clsDataImportHeader)

                    End If


                    If oImportHeader IsNot Nothing Then

                         UcProperties1.msFilePath = psFileName
                         UcProperties1.msFileName = psSafeFileName

                         For Each oTemplate As clsDataImportTemplateV2 In oImportHeader.Templates
                              If oTemplate.Validators Is Nothing Then
                                   oTemplate.Validators = New List(Of clsValidation)
                              End If
                         Next


                         If goOpenWorkBook.Templates IsNot Nothing AndAlso goOpenWorkBook.Templates.Count > 0 Then

                              oImportHeader.Priority = goOpenWorkBook.Templates.Max(Function(n) n.Priority) + 1

                         Else

                              oImportHeader.Priority = 1

                         End If

                         goOpenWorkBook.Templates.Add(oImportHeader)
                         loadWorkbook()



                    Else
                         UpdateProgressStatus()
                         MsgBox("Failed to load Data import template file")
                    End If
               End If
          Catch ex As Exception
               MsgBox(String.Format("Failed to load Data import template file: {0}", ex.Message), "Warning...")
          End Try
     End Sub

     Private Sub OpenWorkBook(psFileName As String, psSafeFileName As String)
          Try
               If File.Exists(psFileName) Then

                    msFilePath = Path.GetDirectoryName(psFileName)
                    msFileName = psSafeFileName
                    txtWorkbookName.EditValue = msFileName

                    UpdateProgressStatus("Opening Document", False)

                    Dim oWorkbook As clsWorkbook = TryCast(LoadFile(psFileName), clsWorkbook)

                    If oWorkbook IsNot Nothing Then

                         goOpenWorkBook = oWorkbook
                         goOpenWorkBook.WorkbookName = msFileName

                         loadWorkbook()

                         'check if there is a data excel sheet in the same folder with the same name
                         If File.Exists(Replace(psFileName, gsWorkbookExtention, ".xlsx")) Then
                              spreadsheetControl.LoadDocument(Replace(psFileName, gsWorkbookExtention, ".xlsx"))
                         End If

                         WriteRegistry("Last Accessed Workbook", psFileName)
                         WriteRegistry("Last Accessed Workbook Safe Name", psSafeFileName)
                    Else
                         UpdateProgressStatus()
                         MsgBox("Failed to load Data import workbook file",, "Warning...")
                    End If

               End If
          Catch ex As Exception
               MsgBox(String.Format("Failed to load Data import workbook file: {0}", ex.Message), "Warning...")
          End Try
     End Sub

     Private Sub UpdateSystemLabel(psConnection As clsConnectionDetails)
          labelSystem.Text = psConnection.Key
     End Sub
#End Region

#Region "Objects"


     Private Sub bbiVersions_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiVersions.ItemClick
          If bbiVersions.Down = True Then
               LcgVersions.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
          Else
               LcgVersions.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never
          End If

     End Sub

     Private Sub lueHierarchies_EditValueChanged(sender As Object, e As EventArgs) Handles lueHierarchies.EditValueChanged

          If lueHierarchies.EditValue IsNot Nothing Then

               gsSelectedHierarchy = lueHierarchies.EditValue.ToString

               Dim oHierarchy As clsHierarchies = TryCast(lueHierarchies.GetSelectedDataRow, clsHierarchies)
               If oHierarchy IsNot Nothing Then
                    gsSelectedHierarchyParent = oHierarchy.ParentHierarchyId
               End If

          Else
               gsSelectedHierarchy = Nothing
          End If

     End Sub

     Private Sub gdWorkbook_FocusedRowChanged(sender As Object, e As FocusedRowChangedEventArgs) Handles gdWorkbook.FocusedRowChanged

          If e.FocusedRowHandle >= 0 Then

               Dim oItem As clsDataImportHeader = TryCast(gdWorkbook.GetRow(e.FocusedRowHandle), clsDataImportHeader)
               If oItem IsNot Nothing Then

                    UcProperties1.ImportHeader = oItem

                    Call ValidateDataTempalte(UcProperties1.ImportHeader)

                    If String.IsNullOrEmpty(UcProperties1.ImportHeader.WorkbookSheetName) = False Then
                         With spreadsheetControl
                              If .Document.Worksheets.Contains(UcProperties1.ImportHeader.WorkbookSheetName) = False Then
                                   '.Document.Worksheets.Add(StrConv(UcProperties1._DataImportTemplate.WorkbookSheetName, vbProperCase))
                                   '.Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1._DataImportTemplate.WorkbookSheetName)
                              Else
                                   'if its found make it the active worksheet
                                   .Document.Worksheets.ActiveWorksheet = .Document.Worksheets(UcProperties1.ImportHeader.WorkbookSheetName)
                              End If
                         End With
                    End If

                    bbiImportExcel.Enabled = oItem.ReadOnlyImport = 0
                    bbiValidateData.Enabled = oItem.ReadOnlyImport = 0
               End If
          End If


     End Sub

     Private Sub tcgImport_SelectedPageChanged(sender As Object, e As DevExpress.XtraLayout.LayoutTabPageChangedEventArgs) Handles tcgImport.SelectedPageChanged
          Call ValidateDataTempalte(UcProperties1.ImportHeader)
     End Sub

     Private Sub LcgImportProperties_Shown(sender As Object, e As EventArgs) Handles LcgImportProperties.Shown
          lueHierarchies.Properties.DataSource = goHierarchies
     End Sub

     Private Sub beiCodeFormat_EditValueChanged(sender As Object, e As EventArgs) Handles beiCodeFormat.EditValueChanged
          goOpenWorkBook.CodeFormat = beiCodeFormat.EditValue.ToString
     End Sub

     Private Sub beiDescriptionFormat_EditValueChanged(sender As Object, e As EventArgs) Handles beiDescriptionFormat.EditValueChanged
          goOpenWorkBook.DescriptionFormat = beiDescriptionFormat.EditValue.ToString
     End Sub


#End Region

#Region "Timmers"

     Private Sub moTimer_Tick(sender As Object, e As EventArgs) Handles moTimer.Tick


          bsiTokenRefesh.Caption = String.Format("Token Refresh: {0:n0}(s)", goHTTPServer.RefreshInSeconds)
          bsiLastLogIn.Caption = String.Format("Last Login: {0}", goHTTPServer.LastLogIn.ToString("hh:mm:ss"))
          bsiLastToken.Caption = String.Format("Last Refresh: {0}", goHTTPServer.LastRefreshToken.ToString("hh:mm:ss"))



     End Sub


#End Region

     Public Sub TestErrorFormWithException()
          Try
               ' Intentionally cause a divide by zero exception
               Dim zero As Integer = 0
               Dim result As Integer = 1 / zero
          Catch ex As Exception
               ' Display the frmError form with the exception details
               Dim errorForm As New frmError(ex)
               errorForm.ShowDialog()
          End Try
     End Sub


End Class
