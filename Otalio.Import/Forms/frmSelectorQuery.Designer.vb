<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSelectorQuery
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
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
        Me.butCancel = New DevExpress.XtraEditors.SimpleButton()
        Me.butOk = New DevExpress.XtraEditors.SimpleButton()
        Me.lcgMain = New DevExpress.XtraLayout.LayoutControlGroup()
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem()
        Me._ucSelectorQuery = New Otalio.Import.ucSelectorQuery()
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.lcgMain, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.butCancel)
        Me.LayoutControl1.Controls.Add(Me.butOk)
        Me.LayoutControl1.Controls.Add(Me._ucSelectorQuery)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(836, 393, 250, 350)
        Me.LayoutControl1.OptionsView.UseSkinIndents = False
        Me.LayoutControl1.Root = Me.lcgMain
        Me.LayoutControl1.Size = New System.Drawing.Size(579, 429)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'butCancel
        '
        Me.butCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.butCancel.Location = New System.Drawing.Point(488, 402)
        Me.butCancel.Name = "butCancel"
        Me.butCancel.Size = New System.Drawing.Size(86, 22)
        Me.butCancel.StyleController = Me.LayoutControl1
        Me.butCancel.TabIndex = 6
        Me.butCancel.Text = "Cancel"
        '
        'butOk
        '
        Me.butOk.Cursor = System.Windows.Forms.Cursors.Default
        Me.butOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.butOk.Location = New System.Drawing.Point(393, 402)
        Me.butOk.Name = "butOk"
        Me.butOk.Size = New System.Drawing.Size(85, 22)
        Me.butOk.StyleController = Me.LayoutControl1
        Me.butOk.TabIndex = 5
        Me.butOk.Text = "Ok"
        '
        'lcgMain
        '
        Me.lcgMain.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
        Me.lcgMain.GroupBordersVisible = False
        Me.lcgMain.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem1, Me.LayoutControlItem3, Me.LayoutControlItem2, Me.EmptySpaceItem1})
        Me.lcgMain.Name = "lcgMain"
        Me.lcgMain.OptionsItemText.TextToControlDistance = 5
        Me.lcgMain.Size = New System.Drawing.Size(579, 429)
        Me.lcgMain.TextVisible = False
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.butCancel
        Me.LayoutControlItem3.Location = New System.Drawing.Point(483, 397)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(96, 32)
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem3.TextVisible = False
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.butOk
        Me.LayoutControlItem2.Location = New System.Drawing.Point(388, 397)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(95, 32)
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem2.TextVisible = False
        '
        'EmptySpaceItem1
        '
        Me.EmptySpaceItem1.AllowHotTrack = False
        Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 397)
        Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
        Me.EmptySpaceItem1.Size = New System.Drawing.Size(388, 32)
        Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
        '
        '_ucSelectorQuery
        '
        Me._ucSelectorQuery._Query = Nothing
        Me._ucSelectorQuery._Selectors = Nothing
        Me._ucSelectorQuery.Location = New System.Drawing.Point(5, 5)
        Me._ucSelectorQuery.Name = "_ucSelectorQuery"
        Me._ucSelectorQuery.Size = New System.Drawing.Size(569, 387)
        Me._ucSelectorQuery.TabIndex = 4
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me._ucSelectorQuery
        Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(579, 397)
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem1.TextVisible = False
        '
        'frmSelectorQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 429)
        Me.Controls.Add(Me.LayoutControl1)
        Me.Name = "frmSelectorQuery"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Selector Query"
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.lcgMain, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
    Friend WithEvents lcgMain As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents _ucSelectorQuery As ucSelectorQuery
    Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents butCancel As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents butOk As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
End Class
