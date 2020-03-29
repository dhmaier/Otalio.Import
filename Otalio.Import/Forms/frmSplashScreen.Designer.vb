<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSplashScreen
  Inherits DevExpress.XtraSplashScreen.SplashScreen

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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSplashScreen))
        Me.peImage = New DevExpress.XtraEditors.PictureEdit()
        Me.peLogo = New DevExpress.XtraEditors.PictureEdit()
        Me.labelStatus = New DevExpress.XtraEditors.LabelControl()
        Me.labelCopyright = New DevExpress.XtraEditors.LabelControl()
        Me.progressBarControl = New DevExpress.XtraEditors.MarqueeProgressBarControl()
        CType(Me.peImage.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.peLogo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.progressBarControl.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'peImage
        '
        Me.peImage.Cursor = System.Windows.Forms.Cursors.Default
        Me.peImage.EditValue = CType(resources.GetObject("peImage.EditValue"), Object)
        Me.peImage.Location = New System.Drawing.Point(12, 12)
        Me.peImage.Name = "peImage"
        Me.peImage.Properties.AllowFocused = False
        Me.peImage.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.peImage.Properties.Appearance.Options.UseBackColor = True
        Me.peImage.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.peImage.Properties.ShowMenu = False
        Me.peImage.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Stretch
        Me.peImage.Size = New System.Drawing.Size(715, 424)
        Me.peImage.TabIndex = 14
        '
        'peLogo
        '
        Me.peLogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.peLogo.EditValue = CType(resources.GetObject("peLogo.EditValue"), Object)
        Me.peLogo.Location = New System.Drawing.Point(547, 478)
        Me.peLogo.Name = "peLogo"
        Me.peLogo.Properties.AllowFocused = False
        Me.peLogo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.peLogo.Properties.Appearance.Options.UseBackColor = True
        Me.peLogo.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.peLogo.Properties.ShowMenu = False
        Me.peLogo.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Squeeze
        Me.peLogo.Size = New System.Drawing.Size(170, 66)
        Me.peLogo.TabIndex = 13
        '
        'labelStatus
        '
        Me.labelStatus.Appearance.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.labelStatus.Appearance.Options.UseFont = True
        Me.labelStatus.Location = New System.Drawing.Point(23, 407)
        Me.labelStatus.Name = "labelStatus"
        Me.labelStatus.Size = New System.Drawing.Size(81, 19)
        Me.labelStatus.TabIndex = 12
        Me.labelStatus.Text = "Starting..."
        '
        'labelCopyright
        '
        Me.labelCopyright.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelCopyright.Appearance.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.labelCopyright.Appearance.Options.UseFont = True
        Me.labelCopyright.Appearance.Options.UseTextOptions = True
        Me.labelCopyright.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        Me.labelCopyright.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal
        Me.labelCopyright.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.labelCopyright.Location = New System.Drawing.Point(23, 503)
        Me.labelCopyright.Name = "labelCopyright"
        Me.labelCopyright.Size = New System.Drawing.Size(79, 19)
        Me.labelCopyright.TabIndex = 11
        Me.labelCopyright.Text = "Copyright"
        '
        'progressBarControl
        '
        Me.progressBarControl.EditValue = 0
        Me.progressBarControl.Location = New System.Drawing.Point(12, 455)
        Me.progressBarControl.Name = "progressBarControl"
        Me.progressBarControl.Size = New System.Drawing.Size(725, 17)
        Me.progressBarControl.TabIndex = 10
        '
        'frmSplashScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(739, 567)
        Me.Controls.Add(Me.peLogo)
        Me.Controls.Add(Me.labelStatus)
        Me.Controls.Add(Me.labelCopyright)
        Me.Controls.Add(Me.progressBarControl)
        Me.Controls.Add(Me.peImage)
        Me.Name = "frmSplashScreen"
        Me.Text = "Form1"
        Me.TopMost = True
        CType(Me.peImage.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.peLogo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.progressBarControl.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents peImage As DevExpress.XtraEditors.PictureEdit
  Private WithEvents peLogo As DevExpress.XtraEditors.PictureEdit
  Private WithEvents labelStatus As DevExpress.XtraEditors.LabelControl
  Private WithEvents labelCopyright As DevExpress.XtraEditors.LabelControl
  Private WithEvents progressBarControl As DevExpress.XtraEditors.MarqueeProgressBarControl
End Class
