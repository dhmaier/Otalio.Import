<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSelector
     Inherits DevExpress.XtraEditors.XtraForm

     'Form overrides dispose to clean up the component list.
     <System.Diagnostics.DebuggerNonUserCode()>
     Protected Overrides Sub Dispose(ByVal disposing As Boolean)
          If disposing AndAlso components IsNot Nothing Then
               components.Dispose()
          End If
          MyBase.Dispose(disposing)
     End Sub

     'Required by the Windows Form Designer
     Private components As System.ComponentModel.IContainer

     'NOTE: The following procedure is required by the Windows Form Designer
     'It can be modified using the Windows Form Designer.  
     'Do not modify it using the code editor.
     <System.Diagnostics.DebuggerStepThrough()>
     Private Sub InitializeComponent()
          Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelector))
          Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
        Me.butLoadFromJson = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.butCancel = New DevExpress.XtraEditors.SimpleButton()
        Me.butOk = New DevExpress.XtraEditors.SimpleButton()
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup()
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem()
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.UcSelectors1 = New Otalio.Import.ucSelectors()
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.butExportToJson = New DevExpress.XtraEditors.SimpleButton()
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.butExportToJson)
        Me.LayoutControl1.Controls.Add(Me.butLoadFromJson)
        Me.LayoutControl1.Controls.Add(Me.UcSelectors1)
        Me.LayoutControl1.Controls.Add(Me.SimpleButton1)
        Me.LayoutControl1.Controls.Add(Me.butCancel)
        Me.LayoutControl1.Controls.Add(Me.butOk)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(836, 393, 250, 350)
        Me.LayoutControl1.OptionsView.UseSkinIndents = False
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(753, 591)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'butLoadFromJson
        '
        Me.butLoadFromJson.Location = New System.Drawing.Point(119, 564)
        Me.butLoadFromJson.Name = "butLoadFromJson"
        Me.butLoadFromJson.Size = New System.Drawing.Size(102, 22)
        Me.butLoadFromJson.StyleController = Me.LayoutControl1
        Me.butLoadFromJson.TabIndex = 9
        Me.butLoadFromJson.Text = "Load Json"
        '
        'SimpleButton1
        '
        Me.SimpleButton1.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.SimpleButton1.Location = New System.Drawing.Point(5, 564)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(104, 22)
        Me.SimpleButton1.StyleController = Me.LayoutControl1
        Me.SimpleButton1.TabIndex = 7
        Me.SimpleButton1.Text = "Delete"
        '
        'butCancel
        '
        Me.butCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butCancel.Location = New System.Drawing.Point(661, 564)
        Me.butCancel.Name = "butCancel"
        Me.butCancel.Size = New System.Drawing.Size(87, 22)
        Me.butCancel.StyleController = Me.LayoutControl1
        Me.butCancel.TabIndex = 6
        Me.butCancel.Text = "Cancel"
        '
        'butOk
        '
        Me.butOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.butOk.Location = New System.Drawing.Point(565, 564)
        Me.butOk.Name = "butOk"
        Me.butOk.Size = New System.Drawing.Size(86, 22)
        Me.butOk.StyleController = Me.LayoutControl1
        Me.butOk.TabIndex = 5
        Me.butOk.Text = "Ok"
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
        Me.LayoutControlGroup1.GroupBordersVisible = False
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem3, Me.LayoutControlItem2, Me.EmptySpaceItem1, Me.LayoutControlItem4, Me.LayoutControlItem5, Me.LayoutControlItem1, Me.LayoutControlItem6})
        Me.LayoutControlGroup1.Name = "Root"
        Me.LayoutControlGroup1.OptionsItemText.TextToControlDistance = 5
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(753, 591)
        Me.LayoutControlGroup1.TextVisible = False
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.butCancel
        Me.LayoutControlItem3.Location = New System.Drawing.Point(656, 559)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(97, 32)
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem3.TextVisible = False
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.butOk
        Me.LayoutControlItem2.Location = New System.Drawing.Point(560, 559)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(96, 32)
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem2.TextVisible = False
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.AllowHotTrack = False
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(342, 559)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(218, 32)
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.SimpleButton1
        Me.LayoutControlItem4.Location = New System.Drawing.Point(0, 559)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(114, 32)
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem4.TextVisible = False
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.butLoadFromJson
        Me.LayoutControlItem1.Location = New System.Drawing.Point(114, 559)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(112, 32)
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem1.TextVisible = False
        '
        'UcSelectors1
        '
        Me.UcSelectors1._Selector = CType(resources.GetObject("UcSelectors1._Selector"), Otalio.Import.clsSelectors)
        Me.UcSelectors1.Location = New System.Drawing.Point(5, 5)
        Me.UcSelectors1.Name = "UcSelectors1"
        Me.UcSelectors1.Size = New System.Drawing.Size(743, 549)
        Me.UcSelectors1.TabIndex = 8
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.UcSelectors1
        Me.LayoutControlItem5.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Size = New System.Drawing.Size(753, 559)
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem5.TextVisible = False
        '
        'butExportToJson
        '
        Me.butExportToJson.Location = New System.Drawing.Point(231, 564)
        Me.butExportToJson.Name = "butExportToJson"
        Me.butExportToJson.Size = New System.Drawing.Size(106, 22)
        Me.butExportToJson.StyleController = Me.LayoutControl1
        Me.butExportToJson.TabIndex = 10
        Me.butExportToJson.Text = "Export Json"
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.butExportToJson
        Me.LayoutControlItem6.Location = New System.Drawing.Point(226, 559)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Size = New System.Drawing.Size(116, 32)
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem6.TextVisible = False
        '
        'frmSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(753, 591)
        Me.Controls.Add(Me.LayoutControl1)
        Me.Name = "frmSelector"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Selector Editor"
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
     Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
     Friend WithEvents butCancel As DevExpress.XtraEditors.SimpleButton
     Friend WithEvents butOk As DevExpress.XtraEditors.SimpleButton
     Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
     Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
     Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
     Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
     Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
     Friend WithEvents UcSelectors1 As ucSelectors
     Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents butLoadFromJson As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents butExportToJson As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
End Class
