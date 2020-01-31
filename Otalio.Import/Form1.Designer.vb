Imports DevExpress.Skins
Imports DevExpress.LookAndFeel
Imports DevExpress.UserSkins
Imports DevExpress.XtraBars.Helpers
Imports DevExpress.XtraBars.Ribbon


<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits RibbonForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

        Me.ribbonControl = New DevExpress.XtraBars.Ribbon.RibbonControl
Me.iExit = New DevExpress.XtraBars.BarButtonItem
Me.siStatus = New DevExpress.XtraBars.BarStaticItem
Me.siInfo = New DevExpress.XtraBars.BarStaticItem
Me.ribbonPageSkins = New DevExpress.XtraBars.Ribbon.RibbonPage
Me.skinsRibbonPageGroup = New DevExpress.XtraBars.Ribbon.RibbonPageGroup
Me.rgbiSkins = New DevExpress.XtraBars.RibbonGalleryBarItem
Me.helpRibbonPage = New DevExpress.XtraBars.Ribbon.RibbonPage
Me.helpRibbonPageGroup = New DevExpress.XtraBars.Ribbon.RibbonPageGroup
Me.iHelp = New DevExpress.XtraBars.BarButtonItem
Me.iAbout = New DevExpress.XtraBars.BarButtonItem
Me.appMenu = New DevExpress.XtraBars.Ribbon.ApplicationMenu(Me.components)
Me.popupControlContainer1 = New DevExpress.XtraBars.PopupControlContainer(Me.components)
Me.someLabelControl2 = New DevExpress.XtraEditors.LabelControl
Me.someLabelControl1 = New DevExpress.XtraEditors.LabelControl
Me.popupControlContainer2 = New DevExpress.XtraBars.PopupControlContainer(Me.components)
Me.buttonEdit = New DevExpress.XtraEditors.ButtonEdit
Me.ribbonStatusBar = New DevExpress.XtraBars.Ribbon.RibbonStatusBar
Me.ribbonImageCollection = New DevExpress.Utils.ImageCollection(Me.components)
Me.ribbonImageCollectionLarge = New DevExpress.Utils.ImageCollection(Me.components)
Me.spreadsheetFormulaBarPanel = New System.Windows.Forms.Panel()
Me.spreadsheetControl = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
Me.splitterControl = New DevExpress.XtraEditors.SplitterControl()
Me.formulaBarNameBoxSplitContainerControl = New DevExpress.XtraEditors.SplitContainerControl()
Me.spreadsheetNameBoxControl = New DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl()
Me.spreadsheetFormulaBarControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl()


        CType(Me.ribbonControl, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.appMenu, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.popupControlContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.popupControlContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.buttonEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.ribbonImageCollection, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.ribbonImageCollectionLarge, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.formulaBarNameBoxSplitContainerControl, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.spreadsheetNameBoxControl.Properties, System.ComponentModel.ISupportInitialize).BeginInit()


        Me.SuspendLayout()

        Me.ribbonControl.ApplicationButtonDropDownControl = Me.appMenu
Me.ribbonControl.ApplicationButtonText = Nothing
Me.ribbonControl.ExpandCollapseItem.Id = 0
Me.ribbonControl.ExpandCollapseItem.Name = ""
Me.ribbonControl.Images = Me.ribbonImageCollection
Me.ribbonControl.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.ribbonControl.ExpandCollapseItem, Me.iExit, Me.iHelp, Me.iAbout, Me.siStatus, Me.siInfo, Me.rgbiSkins})
Me.ribbonControl.LargeImages = Me.ribbonImageCollectionLarge
Me.ribbonControl.Location = New System.Drawing.Point(0, 0)
Me.ribbonControl.MaxItemId = 62
Me.ribbonControl.Name = "ribbonControl"
Me.ribbonControl.PageHeaderItemLinks.Add(Me.iAbout)
Me.ribbonControl.Pages.AddRange(New DevExpress.XtraBars.Ribbon.RibbonPage() {Me.ribbonPageSkins, Me.helpRibbonPage})
Me.ribbonControl.RibbonStyle = DevExpress.XtraBars.Ribbon.RibbonControlStyle.Office2010
Me.ribbonControl.Size = New System.Drawing.Size(750, 144)
Me.ribbonControl.StatusBar = Me.ribbonStatusBar
Me.ribbonControl.Toolbar.ItemLinks.Add(Me.iHelp)
Me.iExit.Caption = "Exit"
Me.iExit.Description = "Closes this program after prompting you to save unsaved data."
Me.iExit.Hint = "Closes this program after prompting you to save unsaved data"
Me.iExit.Id = 20
Me.iExit.ImageIndex = 6
Me.iExit.LargeImageIndex = 6
Me.iExit.Name = "iExit"
Me.siStatus.Caption = "Some Status Info"
Me.siStatus.Id = 31
Me.siStatus.Name = "siStatus"
Me.siStatus.TextAlignment = System.Drawing.StringAlignment.Near
Me.siInfo.Caption = "Some Info"
Me.siInfo.Id = 32
Me.siInfo.Name = "siInfo"
Me.siInfo.TextAlignment = System.Drawing.StringAlignment.Near
Me.ribbonPageSkins.Groups.AddRange(New DevExpress.XtraBars.Ribbon.RibbonPageGroup() {Me.skinsRibbonPageGroup})
Me.ribbonPageSkins.Name = "RibbonPageSkins"
Me.ribbonPageSkins.Text = "Skins"
Me.skinsRibbonPageGroup.ItemLinks.Add(Me.rgbiSkins)
Me.skinsRibbonPageGroup.Name = "skinsRibbonPageGroup"
Me.skinsRibbonPageGroup.ShowCaptionButton = False
Me.skinsRibbonPageGroup.Text = "Skins"
Me.rgbiSkins.Caption = "Skins"
Me.rgbiSkins.Gallery.AllowHoverImages = True
Me.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.Options.UseFont = True
Me.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.Options.UseTextOptions = True
Me.rgbiSkins.Gallery.Appearance.ItemCaptionAppearance.Normal.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
Me.rgbiSkins.Gallery.ColumnCount = 4
Me.rgbiSkins.Gallery.FixedHoverImageSize = False
Me.rgbiSkins.Gallery.ImageSize = New System.Drawing.Size(32, 17)
Me.rgbiSkins.Gallery.ItemImageLocation = DevExpress.Utils.Locations.Top
Me.rgbiSkins.Gallery.RowCount = 4
Me.rgbiSkins.Id = 60
Me.rgbiSkins.Name = "rgbiSkins"
Me.helpRibbonPage.Groups.AddRange(New DevExpress.XtraBars.Ribbon.RibbonPageGroup() {Me.helpRibbonPageGroup})
Me.helpRibbonPage.Name = "helpRibbonPage"
Me.helpRibbonPage.Text = "Help"
Me.helpRibbonPageGroup.ItemLinks.Add(Me.iHelp)
Me.helpRibbonPageGroup.ItemLinks.Add(Me.iAbout)
Me.helpRibbonPageGroup.Name = "helpRibbonPageGroup"
Me.helpRibbonPageGroup.Text = "Help"
Me.iHelp.Caption = "Help"
Me.iHelp.Description = "Start the program help system."
Me.iHelp.Hint = "Start the program help system"
Me.iHelp.Id = 22
Me.iHelp.ImageIndex = 7
Me.iHelp.LargeImageIndex = 7
Me.iHelp.Name = "iHelp"
Me.iAbout.Caption = "About"
Me.iAbout.Description = "Displays general program information."
Me.iAbout.Hint = "Displays general program information"
Me.iAbout.Id = 24
Me.iAbout.ImageIndex = 8
Me.iAbout.LargeImageIndex = 8
Me.iAbout.Name = "iAbout"
Me.appMenu.BottomPaneControlContainer = Me.popupControlContainer2
Me.appMenu.ItemLinks.Add(Me.iExit)
Me.appMenu.Name = "appMenu"
Me.appMenu.Ribbon = Me.ribbonControl
Me.appMenu.RightPaneControlContainer = Me.popupControlContainer1
Me.appMenu.ShowRightPane = True
Me.popupControlContainer1.SuspendLayout()
Me.popupControlContainer1.Appearance.BackColor = System.Drawing.Color.Transparent
Me.popupControlContainer1.Appearance.Options.UseBackColor = True
Me.popupControlContainer1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
Me.popupControlContainer1.Controls.Add(Me.someLabelControl2)
Me.popupControlContainer1.Controls.Add(Me.someLabelControl1)
Me.popupControlContainer1.Location = New System.Drawing.Point(111, 197)
Me.popupControlContainer1.Name = "popupControlContainer1"
Me.popupControlContainer1.Ribbon = Me.ribbonControl
Me.popupControlContainer1.Size = New System.Drawing.Size(76, 70)
Me.popupControlContainer1.TabIndex = 6
Me.popupControlContainer1.Visible = False
Me.popupControlContainer1.ResumeLayout(False)
Me.popupControlContainer1.PerformLayout()
Me.someLabelControl2.Location = New System.Drawing.Point(3, 57)
Me.someLabelControl2.Name = "someLabelControl2"
Me.someLabelControl2.Size = New System.Drawing.Size(49, 13)
Me.someLabelControl2.TabIndex = 0
Me.someLabelControl2.Text = "Some Info"
Me.someLabelControl1.Location = New System.Drawing.Point(3, 3)
Me.someLabelControl1.Name = "someLabelControl1"
Me.someLabelControl1.Size = New System.Drawing.Size(49, 13)
Me.someLabelControl1.TabIndex = 0
Me.someLabelControl1.Text = "Some Info"
Me.popupControlContainer2.SuspendLayout()
Me.popupControlContainer2.Appearance.BackColor = System.Drawing.Color.Transparent
Me.popupControlContainer2.Appearance.Options.UseBackColor = True
Me.popupControlContainer2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
Me.popupControlContainer2.Controls.Add(Me.buttonEdit)
Me.popupControlContainer2.Location = New System.Drawing.Point(238, 289)
Me.popupControlContainer2.Name = "popupControlContainer2"
Me.popupControlContainer2.Ribbon = Me.ribbonControl
Me.popupControlContainer2.Size = New System.Drawing.Size(118, 28)
Me.popupControlContainer2.TabIndex = 7
Me.popupControlContainer2.Visible = False
Me.popupControlContainer2.ResumeLayout(False)
Me.buttonEdit.EditValue = "Some Text"
Me.buttonEdit.Location = New System.Drawing.Point(3, 5)
Me.buttonEdit.MenuManager = Me.ribbonControl
Me.buttonEdit.Name = "buttonEdit"
Me.buttonEdit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton})
Me.buttonEdit.Size = New System.Drawing.Size(100, 20)
Me.buttonEdit.TabIndex = 0
Me.ribbonStatusBar.ItemLinks.Add(Me.siStatus)
Me.ribbonStatusBar.ItemLinks.Add(Me.siInfo)
Me.ribbonStatusBar.Location = New System.Drawing.Point(0, 367)
Me.ribbonStatusBar.Name = "ribbonStatusBar"
Me.ribbonStatusBar.Ribbon = Me.ribbonControl
Me.ribbonStatusBar.Size = New System.Drawing.Size(750, 31)
Me.ribbonImageCollection.ImageStream = CType(resources.GetObject("ribbonImageCollection.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
Me.ribbonImageCollection.Images.SetKeyName(0, "Ribbon_New_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(1, "Ribbon_Open_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(2, "Ribbon_Close_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(3, "Ribbon_Find_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(4, "Ribbon_Save_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(5, "Ribbon_SaveAs_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(6, "Ribbon_Exit_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(7, "Ribbon_Content_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(8, "Ribbon_Info_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(9, "Ribbon_Bold_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(10, "Ribbon_Italic_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(11, "Ribbon_Underline_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(12, "Ribbon_AlignLeft_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(13, "Ribbon_AlignCenter_16x16.png")
Me.ribbonImageCollection.Images.SetKeyName(14, "Ribbon_AlignRight_16x16.png")
Me.ribbonImageCollectionLarge.ImageSize = New System.Drawing.Size(32, 32)
Me.ribbonImageCollectionLarge.ImageStream = CType(resources.GetObject("ribbonImageCollectionLarge.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
Me.ribbonImageCollectionLarge.Images.SetKeyName(0, "Ribbon_New_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(1, "Ribbon_Open_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(2, "Ribbon_Close_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(3, "Ribbon_Find_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(4, "Ribbon_Save_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(5, "Ribbon_SaveAs_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(6, "Ribbon_Exit_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(7, "Ribbon_Content_32x32.png")
Me.ribbonImageCollectionLarge.Images.SetKeyName(8, "Ribbon_Info_32x32.png")
Me.spreadsheetFormulaBarPanel.SuspendLayout()
Me.spreadsheetFormulaBarPanel.Controls.Add(Me.spreadsheetControl)
Me.spreadsheetFormulaBarPanel.Controls.Add(Me.splitterControl)
Me.spreadsheetFormulaBarPanel.Controls.Add(Me.formulaBarNameBoxSplitContainerControl)
Me.spreadsheetFormulaBarPanel.Dock = System.Windows.Forms.DockStyle.Fill
Me.spreadsheetFormulaBarPanel.Location = New System.Drawing.Point(0, 0)
Me.spreadsheetFormulaBarPanel.Name = "spreadsheetFormulaBarPanel"
Me.spreadsheetFormulaBarPanel.Size = New System.Drawing.Size(499, 375)
Me.spreadsheetFormulaBarPanel.TabIndex = 3
Me.spreadsheetFormulaBarPanel.ResumeLayout(False)
Me.spreadsheetControl.Dock = System.Windows.Forms.DockStyle.Fill
Me.spreadsheetControl.Location = New System.Drawing.Point(0, 25)
Me.spreadsheetControl.Name = "spreadsheetControl"
Me.spreadsheetControl.Size = New System.Drawing.Size(499, 350)
Me.spreadsheetControl.TabIndex = 1
Me.spreadsheetControl.Text = "spreadsheetControl1"
Me.splitterControl.Dock = System.Windows.Forms.DockStyle.Top
Me.splitterControl.Location = New System.Drawing.Point(0, 20)
Me.splitterControl.MinSize = 20
Me.splitterControl.Name = "splitterControl"
Me.splitterControl.Size = New System.Drawing.Size(499, 5)
Me.splitterControl.TabIndex = 2
Me.splitterControl.TabStop = False
Me.formulaBarNameBoxSplitContainerControl.SuspendLayout()
Me.formulaBarNameBoxSplitContainerControl.Dock = System.Windows.Forms.DockStyle.Top
Me.formulaBarNameBoxSplitContainerControl.Location = New System.Drawing.Point(0, 0)
Me.formulaBarNameBoxSplitContainerControl.Name = "formulaBarNameBoxSplitContainerControl"
Me.formulaBarNameBoxSplitContainerControl.Panel1.Controls.Add(Me.spreadsheetNameBoxControl)
Me.formulaBarNameBoxSplitContainerControl.Panel2.Controls.Add(Me.spreadsheetFormulaBarControl1)
Me.formulaBarNameBoxSplitContainerControl.Size = New System.Drawing.Size(499, 20)
Me.formulaBarNameBoxSplitContainerControl.SplitterPosition = 145
Me.formulaBarNameBoxSplitContainerControl.TabIndex = 3
Me.formulaBarNameBoxSplitContainerControl.ResumeLayout(False)
Me.spreadsheetNameBoxControl.Dock = System.Windows.Forms.DockStyle.Fill
Me.spreadsheetNameBoxControl.EditValue = "A1"
Me.spreadsheetNameBoxControl.Location = New System.Drawing.Point(0, 0)
Me.spreadsheetNameBoxControl.Name = "spreadsheetNameBoxControl"
Me.spreadsheetNameBoxControl.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() { New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
Me.spreadsheetNameBoxControl.Size = New System.Drawing.Size(145, 20)
Me.spreadsheetNameBoxControl.SpreadsheetControl = Me.spreadsheetControl
Me.spreadsheetNameBoxControl.TabIndex = 0
Me.spreadsheetFormulaBarControl1.Dock = System.Windows.Forms.DockStyle.Fill
Me.spreadsheetFormulaBarControl1.Location = New System.Drawing.Point(0, 0)
Me.spreadsheetFormulaBarControl1.MinimumSize = New System.Drawing.Size(0, 20)
Me.spreadsheetFormulaBarControl1.Name = "spreadsheetFormulaBarControl1"
Me.spreadsheetFormulaBarControl1.Size = New System.Drawing.Size(349, 20)
Me.spreadsheetFormulaBarControl1.SpreadsheetControl = Me.spreadsheetControl
Me.spreadsheetFormulaBarControl1.TabIndex = 0


        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ClientSize = New System.Drawing.Size(1100, 700)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

        Me.AllowFormGlass = DevExpress.Utils.DefaultBoolean.False        
        Me.Controls.Add(Me.spreadsheetFormulaBarPanel)

        Me.Controls.Add(Me.ribbonControl)
Me.Controls.Add(Me.popupControlContainer1)
Me.Controls.Add(Me.popupControlContainer2)
Me.Controls.Add(Me.ribbonStatusBar)


        CType(Me.ribbonControl, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.appMenu, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.popupControlContainer1, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.popupControlContainer2, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.buttonEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.ribbonImageCollection, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.ribbonImageCollectionLarge, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.formulaBarNameBoxSplitContainerControl, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.spreadsheetNameBoxControl.Properties, System.ComponentModel.ISupportInitialize).EndInit()


        Me.ResumeLayout(False)
    End Sub
    Private WithEvents ribbonControl As DevExpress.XtraBars.Ribbon.RibbonControl
Private WithEvents iExit As DevExpress.XtraBars.BarButtonItem
Private WithEvents siStatus As DevExpress.XtraBars.BarStaticItem
Private WithEvents siInfo As DevExpress.XtraBars.BarStaticItem
Private WithEvents ribbonPageSkins As DevExpress.XtraBars.Ribbon.RibbonPage
Private WithEvents skinsRibbonPageGroup As DevExpress.XtraBars.Ribbon.RibbonPageGroup
Private WithEvents rgbiSkins As DevExpress.XtraBars.RibbonGalleryBarItem
Private WithEvents helpRibbonPage As DevExpress.XtraBars.Ribbon.RibbonPage
Private WithEvents helpRibbonPageGroup As DevExpress.XtraBars.Ribbon.RibbonPageGroup
Private WithEvents iHelp As DevExpress.XtraBars.BarButtonItem
Private WithEvents iAbout As DevExpress.XtraBars.BarButtonItem
Private WithEvents appMenu As DevExpress.XtraBars.Ribbon.ApplicationMenu
Private WithEvents popupControlContainer1 As DevExpress.XtraBars.PopupControlContainer
Private WithEvents someLabelControl2 As DevExpress.XtraEditors.LabelControl
Private WithEvents someLabelControl1 As DevExpress.XtraEditors.LabelControl
Private WithEvents popupControlContainer2 As DevExpress.XtraBars.PopupControlContainer
Private WithEvents buttonEdit As DevExpress.XtraEditors.ButtonEdit
Private WithEvents ribbonStatusBar As DevExpress.XtraBars.Ribbon.RibbonStatusBar
Private WithEvents ribbonImageCollection As DevExpress.Utils.ImageCollection
Private WithEvents ribbonImageCollectionLarge As DevExpress.Utils.ImageCollection
Private WithEvents spreadsheetFormulaBarPanel As System.Windows.Forms.Panel
Private WithEvents spreadsheetControl As DevExpress.XtraSpreadsheet.SpreadsheetControl
Private WithEvents splitterControl As DevExpress.XtraEditors.SplitterControl
Private WithEvents formulaBarNameBoxSplitContainerControl As DevExpress.XtraEditors.SplitContainerControl
Private WithEvents spreadsheetNameBoxControl As DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl
Private WithEvents spreadsheetFormulaBarControl1 As DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl
        
End Class
