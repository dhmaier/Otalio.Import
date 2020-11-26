<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ucConnectionDetails
  Inherits System.Windows.Forms.UserControl

  'UserControl overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()>
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
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
        Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
        Me.txtHTTPSPassword = New DevExpress.XtraEditors.TextEdit()
        Me.txtHTTPSUserName = New DevExpress.XtraEditors.TextEdit()
        Me.gridHistory = New DevExpress.XtraGrid.GridControl()
        Me.gdHistory = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.chkEnableLoging = New DevExpress.XtraEditors.CheckEdit()
        Me.btnSaveConnection = New DevExpress.XtraEditors.SimpleButton()
        Me.btnTestConnection = New DevExpress.XtraEditors.SimpleButton()
        Me.txtPassword = New DevExpress.XtraEditors.TextEdit()
        Me.txtUserName = New DevExpress.XtraEditors.TextEdit()
        Me.txtServerName = New DevExpress.XtraEditors.TextEdit()
        Me.LayoutControlGroup1 = New DevExpress.XtraLayout.LayoutControlGroup()
        Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem4 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem5 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup()
        Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem7 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem8 = New DevExpress.XtraLayout.LayoutControlItem()
        Me.LayoutControlItem9 = New DevExpress.XtraLayout.LayoutControlItem()
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.LayoutControl1.SuspendLayout()
        CType(Me.txtHTTPSPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtHTTPSUserName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gdHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkEnableLoging.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUserName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtServerName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LayoutControl1
        '
        Me.LayoutControl1.Controls.Add(Me.txtHTTPSPassword)
        Me.LayoutControl1.Controls.Add(Me.txtHTTPSUserName)
        Me.LayoutControl1.Controls.Add(Me.gridHistory)
        Me.LayoutControl1.Controls.Add(Me.chkEnableLoging)
        Me.LayoutControl1.Controls.Add(Me.btnSaveConnection)
        Me.LayoutControl1.Controls.Add(Me.btnTestConnection)
        Me.LayoutControl1.Controls.Add(Me.txtPassword)
        Me.LayoutControl1.Controls.Add(Me.txtUserName)
        Me.LayoutControl1.Controls.Add(Me.txtServerName)
        Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControl1.Name = "LayoutControl1"
        Me.LayoutControl1.Root = Me.LayoutControlGroup1
        Me.LayoutControl1.Size = New System.Drawing.Size(465, 726)
        Me.LayoutControl1.TabIndex = 0
        Me.LayoutControl1.Text = "LayoutControl1"
        '
        'txtHTTPSPassword
        '
        Me.txtHTTPSPassword.Location = New System.Drawing.Point(105, 108)
        Me.txtHTTPSPassword.Name = "txtHTTPSPassword"
        Me.txtHTTPSPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtHTTPSPassword.Size = New System.Drawing.Size(348, 20)
        Me.txtHTTPSPassword.StyleController = Me.LayoutControl1
        Me.txtHTTPSPassword.TabIndex = 12
        '
        'txtHTTPSUserName
        '
        Me.txtHTTPSUserName.Location = New System.Drawing.Point(105, 84)
        Me.txtHTTPSUserName.Name = "txtHTTPSUserName"
        Me.txtHTTPSUserName.Size = New System.Drawing.Size(348, 20)
        Me.txtHTTPSUserName.StyleController = Me.LayoutControl1
        Me.txtHTTPSUserName.TabIndex = 11
        '
        'gridHistory
        '
        Me.gridHistory.Location = New System.Drawing.Point(12, 158)
        Me.gridHistory.MainView = Me.gdHistory
        Me.gridHistory.Name = "gridHistory"
        Me.gridHistory.Size = New System.Drawing.Size(441, 487)
        Me.gridHistory.TabIndex = 10
        Me.gridHistory.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gdHistory})
        '
        'gdHistory
        '
        Me.gdHistory.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn3, Me.GridColumn2})
        Me.gdHistory.GridControl = Me.gridHistory
        Me.gdHistory.GroupCount = 1
        Me.gdHistory.Name = "gdHistory"
        Me.gdHistory.OptionsBehavior.AutoExpandAllGroups = True
        Me.gdHistory.SortInfo.AddRange(New DevExpress.XtraGrid.Columns.GridColumnSortInfo() {New DevExpress.XtraGrid.Columns.GridColumnSortInfo(Me.GridColumn1, DevExpress.Data.ColumnSortOrder.Ascending)})
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "Server"
        Me.GridColumn1.FieldName = "_ServerAddress"
        Me.GridColumn1.Name = "GridColumn1"
        Me.GridColumn1.OptionsColumn.AllowEdit = False
        Me.GridColumn1.Visible = True
        Me.GridColumn1.VisibleIndex = 0
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "User"
        Me.GridColumn3.FieldName = "_UserName"
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.OptionsColumn.AllowEdit = False
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 1
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Date"
        Me.GridColumn2.DisplayFormat.FormatString = "d"
        Me.GridColumn2.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.GridColumn2.FieldName = "_DateLastUsed"
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.OptionsColumn.AllowEdit = False
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 0
        '
        'chkEnableLoging
        '
        Me.chkEnableLoging.Location = New System.Drawing.Point(117, 682)
        Me.chkEnableLoging.Name = "chkEnableLoging"
        Me.chkEnableLoging.Properties.Caption = "Enabled"
        Me.chkEnableLoging.Size = New System.Drawing.Size(324, 20)
        Me.chkEnableLoging.StyleController = Me.LayoutControl1
        Me.chkEnableLoging.TabIndex = 9
        '
        'btnSaveConnection
        '
        Me.btnSaveConnection.Location = New System.Drawing.Point(234, 132)
        Me.btnSaveConnection.Name = "btnSaveConnection"
        Me.btnSaveConnection.Size = New System.Drawing.Size(219, 22)
        Me.btnSaveConnection.StyleController = Me.LayoutControl1
        Me.btnSaveConnection.TabIndex = 8
        Me.btnSaveConnection.Text = "Save Connection"
        '
        'btnTestConnection
        '
        Me.btnTestConnection.Location = New System.Drawing.Point(12, 132)
        Me.btnTestConnection.Name = "btnTestConnection"
        Me.btnTestConnection.Size = New System.Drawing.Size(218, 22)
        Me.btnTestConnection.StyleController = Me.LayoutControl1
        Me.btnTestConnection.TabIndex = 7
        Me.btnTestConnection.Text = "Test Connection"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(105, 60)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(348, 20)
        Me.txtPassword.StyleController = Me.LayoutControl1
        Me.txtPassword.TabIndex = 6
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(105, 36)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(348, 20)
        Me.txtUserName.StyleController = Me.LayoutControl1
        Me.txtUserName.TabIndex = 5
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(105, 12)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(348, 20)
        Me.txtServerName.StyleController = Me.LayoutControl1
        Me.txtServerName.TabIndex = 4
        '
        'LayoutControlGroup1
        '
        Me.LayoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
        Me.LayoutControlGroup1.GroupBordersVisible = False
        Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem2, Me.LayoutControlItem1, Me.LayoutControlItem3, Me.LayoutControlItem4, Me.LayoutControlItem5, Me.LayoutControlGroup2, Me.LayoutControlItem7, Me.LayoutControlItem8, Me.LayoutControlItem9})
        Me.LayoutControlGroup1.Name = "LayoutControlGroup1"
        Me.LayoutControlGroup1.Size = New System.Drawing.Size(465, 726)
        Me.LayoutControlGroup1.TextVisible = False
        '
        'LayoutControlItem2
        '
        Me.LayoutControlItem2.Control = Me.txtUserName
        Me.LayoutControlItem2.Location = New System.Drawing.Point(0, 24)
        Me.LayoutControlItem2.Name = "LayoutControlItem2"
        Me.LayoutControlItem2.Size = New System.Drawing.Size(445, 24)
        Me.LayoutControlItem2.Text = "User Name:"
        Me.LayoutControlItem2.TextSize = New System.Drawing.Size(90, 13)
        '
        'LayoutControlItem1
        '
        Me.LayoutControlItem1.Control = Me.txtServerName
        Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem1.Name = "LayoutControlItem1"
        Me.LayoutControlItem1.Size = New System.Drawing.Size(445, 24)
        Me.LayoutControlItem1.Text = "Otalio Server:"
        Me.LayoutControlItem1.TextSize = New System.Drawing.Size(90, 13)
        '
        'LayoutControlItem3
        '
        Me.LayoutControlItem3.Control = Me.txtPassword
        Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 48)
        Me.LayoutControlItem3.Name = "LayoutControlItem3"
        Me.LayoutControlItem3.Size = New System.Drawing.Size(445, 24)
        Me.LayoutControlItem3.Text = "Password:"
        Me.LayoutControlItem3.TextSize = New System.Drawing.Size(90, 13)
        '
        'LayoutControlItem4
        '
        Me.LayoutControlItem4.Control = Me.btnTestConnection
        Me.LayoutControlItem4.Location = New System.Drawing.Point(0, 120)
        Me.LayoutControlItem4.Name = "LayoutControlItem4"
        Me.LayoutControlItem4.Size = New System.Drawing.Size(222, 26)
        Me.LayoutControlItem4.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem4.TextVisible = False
        '
        'LayoutControlItem5
        '
        Me.LayoutControlItem5.Control = Me.btnSaveConnection
        Me.LayoutControlItem5.Location = New System.Drawing.Point(222, 120)
        Me.LayoutControlItem5.Name = "LayoutControlItem5"
        Me.LayoutControlItem5.Size = New System.Drawing.Size(223, 26)
        Me.LayoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem5.TextVisible = False
        '
        'LayoutControlGroup2
        '
        Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem6})
        Me.LayoutControlGroup2.Location = New System.Drawing.Point(0, 637)
        Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
        Me.LayoutControlGroup2.Size = New System.Drawing.Size(445, 69)
        Me.LayoutControlGroup2.Text = "Options"
        '
        'LayoutControlItem6
        '
        Me.LayoutControlItem6.Control = Me.chkEnableLoging
        Me.LayoutControlItem6.Location = New System.Drawing.Point(0, 0)
        Me.LayoutControlItem6.Name = "LayoutControlItem6"
        Me.LayoutControlItem6.Size = New System.Drawing.Size(421, 24)
        Me.LayoutControlItem6.Text = "Enable Logging:"
        Me.LayoutControlItem6.TextSize = New System.Drawing.Size(90, 13)
        '
        'LayoutControlItem7
        '
        Me.LayoutControlItem7.Control = Me.gridHistory
        Me.LayoutControlItem7.Location = New System.Drawing.Point(0, 146)
        Me.LayoutControlItem7.Name = "LayoutControlItem7"
        Me.LayoutControlItem7.Size = New System.Drawing.Size(445, 491)
        Me.LayoutControlItem7.TextSize = New System.Drawing.Size(0, 0)
        Me.LayoutControlItem7.TextVisible = False
        '
        'LayoutControlItem8
        '
        Me.LayoutControlItem8.Control = Me.txtHTTPSUserName
        Me.LayoutControlItem8.Location = New System.Drawing.Point(0, 72)
        Me.LayoutControlItem8.Name = "LayoutControlItem8"
        Me.LayoutControlItem8.Size = New System.Drawing.Size(445, 24)
        Me.LayoutControlItem8.Text = "HTTPS User Name:"
        Me.LayoutControlItem8.TextSize = New System.Drawing.Size(90, 13)
        '
        'LayoutControlItem9
        '
        Me.LayoutControlItem9.Control = Me.txtHTTPSPassword
        Me.LayoutControlItem9.Location = New System.Drawing.Point(0, 96)
        Me.LayoutControlItem9.Name = "LayoutControlItem9"
        Me.LayoutControlItem9.Size = New System.Drawing.Size(445, 24)
        Me.LayoutControlItem9.Text = "HTTPS Password:"
        Me.LayoutControlItem9.TextSize = New System.Drawing.Size(90, 13)
        '
        'ucConnectionDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LayoutControl1)
        Me.Name = "ucConnectionDetails"
        Me.Size = New System.Drawing.Size(465, 726)
        CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.LayoutControl1.ResumeLayout(False)
        CType(Me.txtHTTPSPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtHTTPSUserName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gdHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkEnableLoging.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUserName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtServerName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LayoutControlItem9, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
  Friend WithEvents LayoutControlGroup1 As DevExpress.XtraLayout.LayoutControlGroup
  Friend WithEvents btnTestConnection As DevExpress.XtraEditors.SimpleButton
  Friend WithEvents txtPassword As DevExpress.XtraEditors.TextEdit
  Friend WithEvents txtUserName As DevExpress.XtraEditors.TextEdit
  Friend WithEvents txtServerName As DevExpress.XtraEditors.TextEdit
  Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
  Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
  Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
  Friend WithEvents LayoutControlItem4 As DevExpress.XtraLayout.LayoutControlItem
  Friend WithEvents btnSaveConnection As DevExpress.XtraEditors.SimpleButton
  Friend WithEvents LayoutControlItem5 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents chkEnableLoging As DevExpress.XtraEditors.CheckEdit
    Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
    Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents gridHistory As DevExpress.XtraGrid.GridControl
    Friend WithEvents gdHistory As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents LayoutControlItem7 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents txtHTTPSPassword As DevExpress.XtraEditors.TextEdit
    Friend WithEvents txtHTTPSUserName As DevExpress.XtraEditors.TextEdit
    Friend WithEvents LayoutControlItem8 As DevExpress.XtraLayout.LayoutControlItem
    Friend WithEvents LayoutControlItem9 As DevExpress.XtraLayout.LayoutControlItem
End Class
