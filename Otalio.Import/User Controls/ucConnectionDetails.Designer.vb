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
    Me.EmptySpaceItem1 = New DevExpress.XtraLayout.EmptySpaceItem()
    Me.chkEnableLoging = New DevExpress.XtraEditors.CheckEdit()
    Me.LayoutControlItem6 = New DevExpress.XtraLayout.LayoutControlItem()
    Me.EmptySpaceItem2 = New DevExpress.XtraLayout.EmptySpaceItem()
    Me.LayoutControlGroup2 = New DevExpress.XtraLayout.LayoutControlGroup()
    CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.LayoutControl1.SuspendLayout()
    CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.txtUserName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.txtServerName.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.chkEnableLoging.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'LayoutControl1
    '
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
    'btnSaveConnection
    '
    Me.btnSaveConnection.Location = New System.Drawing.Point(234, 84)
    Me.btnSaveConnection.Name = "btnSaveConnection"
    Me.btnSaveConnection.Size = New System.Drawing.Size(219, 22)
    Me.btnSaveConnection.StyleController = Me.LayoutControl1
    Me.btnSaveConnection.TabIndex = 8
    Me.btnSaveConnection.Text = "Save Connection"
    '
    'btnTestConnection
    '
    Me.btnTestConnection.Location = New System.Drawing.Point(12, 84)
    Me.btnTestConnection.Name = "btnTestConnection"
    Me.btnTestConnection.Size = New System.Drawing.Size(218, 22)
    Me.btnTestConnection.StyleController = Me.LayoutControl1
    Me.btnTestConnection.TabIndex = 7
    Me.btnTestConnection.Text = "Test Connection"
    '
    'txtPassword
    '
    Me.txtPassword.Location = New System.Drawing.Point(91, 60)
    Me.txtPassword.Name = "txtPassword"
    Me.txtPassword.Properties.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
    Me.txtPassword.Size = New System.Drawing.Size(362, 20)
    Me.txtPassword.StyleController = Me.LayoutControl1
    Me.txtPassword.TabIndex = 6
    '
    'txtUserName
    '
    Me.txtUserName.Location = New System.Drawing.Point(91, 36)
    Me.txtUserName.Name = "txtUserName"
    Me.txtUserName.Size = New System.Drawing.Size(362, 20)
    Me.txtUserName.StyleController = Me.LayoutControl1
    Me.txtUserName.TabIndex = 5
    '
    'txtServerName
    '
    Me.txtServerName.Location = New System.Drawing.Point(91, 12)
    Me.txtServerName.Name = "txtServerName"
    Me.txtServerName.Size = New System.Drawing.Size(362, 20)
    Me.txtServerName.StyleController = Me.LayoutControl1
    Me.txtServerName.TabIndex = 4
    '
    'LayoutControlGroup1
    '
    Me.LayoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
    Me.LayoutControlGroup1.GroupBordersVisible = False
    Me.LayoutControlGroup1.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem2, Me.LayoutControlItem1, Me.LayoutControlItem3, Me.LayoutControlItem4, Me.LayoutControlItem5, Me.EmptySpaceItem1, Me.EmptySpaceItem2, Me.LayoutControlGroup2})
    Me.LayoutControlGroup1.Location = New System.Drawing.Point(0, 0)
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
    Me.LayoutControlItem2.TextSize = New System.Drawing.Size(76, 13)
    '
    'LayoutControlItem1
    '
    Me.LayoutControlItem1.Control = Me.txtServerName
    Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 0)
    Me.LayoutControlItem1.Name = "LayoutControlItem1"
    Me.LayoutControlItem1.Size = New System.Drawing.Size(445, 24)
    Me.LayoutControlItem1.Text = "Otalio Server:"
    Me.LayoutControlItem1.TextSize = New System.Drawing.Size(76, 13)
    '
    'LayoutControlItem3
    '
    Me.LayoutControlItem3.Control = Me.txtPassword
    Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 48)
    Me.LayoutControlItem3.Name = "LayoutControlItem3"
    Me.LayoutControlItem3.Size = New System.Drawing.Size(445, 24)
    Me.LayoutControlItem3.Text = "Password:"
    Me.LayoutControlItem3.TextSize = New System.Drawing.Size(76, 13)
    '
    'LayoutControlItem4
    '
    Me.LayoutControlItem4.Control = Me.btnTestConnection
    Me.LayoutControlItem4.Location = New System.Drawing.Point(0, 72)
    Me.LayoutControlItem4.Name = "LayoutControlItem4"
    Me.LayoutControlItem4.Size = New System.Drawing.Size(222, 26)
    Me.LayoutControlItem4.TextSize = New System.Drawing.Size(0, 0)
    Me.LayoutControlItem4.TextVisible = False
    '
    'LayoutControlItem5
    '
    Me.LayoutControlItem5.Control = Me.btnSaveConnection
    Me.LayoutControlItem5.Location = New System.Drawing.Point(222, 72)
    Me.LayoutControlItem5.Name = "LayoutControlItem5"
    Me.LayoutControlItem5.Size = New System.Drawing.Size(223, 26)
    Me.LayoutControlItem5.TextSize = New System.Drawing.Size(0, 0)
    Me.LayoutControlItem5.TextVisible = False
    '
    'EmptySpaceItem1
    '
    Me.EmptySpaceItem1.AllowHotTrack = False
    Me.EmptySpaceItem1.Location = New System.Drawing.Point(0, 173)
    Me.EmptySpaceItem1.Name = "EmptySpaceItem1"
    Me.EmptySpaceItem1.Size = New System.Drawing.Size(445, 533)
    Me.EmptySpaceItem1.TextSize = New System.Drawing.Size(0, 0)
    '
    'chkEnableLoging
    '
    Me.chkEnableLoging.Location = New System.Drawing.Point(103, 150)
    Me.chkEnableLoging.Name = "chkEnableLoging"
    Me.chkEnableLoging.Properties.Caption = "Enabled"
    Me.chkEnableLoging.Size = New System.Drawing.Size(338, 19)
    Me.chkEnableLoging.StyleController = Me.LayoutControl1
    Me.chkEnableLoging.TabIndex = 9
    '
    'LayoutControlItem6
    '
    Me.LayoutControlItem6.Control = Me.chkEnableLoging
    Me.LayoutControlItem6.Location = New System.Drawing.Point(0, 0)
    Me.LayoutControlItem6.Name = "LayoutControlItem6"
    Me.LayoutControlItem6.Size = New System.Drawing.Size(421, 23)
    Me.LayoutControlItem6.Text = "Enable Logging:"
    Me.LayoutControlItem6.TextSize = New System.Drawing.Size(76, 13)
    '
    'EmptySpaceItem2
    '
    Me.EmptySpaceItem2.AllowHotTrack = False
    Me.EmptySpaceItem2.Location = New System.Drawing.Point(0, 98)
    Me.EmptySpaceItem2.Name = "EmptySpaceItem2"
    Me.EmptySpaceItem2.Size = New System.Drawing.Size(445, 10)
    Me.EmptySpaceItem2.TextSize = New System.Drawing.Size(0, 0)
    '
    'LayoutControlGroup2
    '
    Me.LayoutControlGroup2.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem6})
    Me.LayoutControlGroup2.Location = New System.Drawing.Point(0, 108)
    Me.LayoutControlGroup2.Name = "LayoutControlGroup2"
    Me.LayoutControlGroup2.Size = New System.Drawing.Size(445, 65)
    Me.LayoutControlGroup2.Text = "Options"
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
    CType(Me.txtPassword.Properties, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.txtUserName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.txtServerName.Properties, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlGroup1, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem4, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem5, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.EmptySpaceItem1, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.chkEnableLoging.Properties, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlItem6, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.EmptySpaceItem2, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.LayoutControlGroup2, System.ComponentModel.ISupportInitialize).EndInit()
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
  Friend WithEvents EmptySpaceItem1 As DevExpress.XtraLayout.EmptySpaceItem
  Friend WithEvents chkEnableLoging As DevExpress.XtraEditors.CheckEdit
  Friend WithEvents EmptySpaceItem2 As DevExpress.XtraLayout.EmptySpaceItem
  Friend WithEvents LayoutControlGroup2 As DevExpress.XtraLayout.LayoutControlGroup
  Friend WithEvents LayoutControlItem6 As DevExpress.XtraLayout.LayoutControlItem
End Class
