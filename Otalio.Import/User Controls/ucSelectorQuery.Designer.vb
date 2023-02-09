<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ucSelectorQuery
     Inherits DevExpress.XtraEditors.XtraUserControl

     'UserControl overrides dispose to clean up the component list.
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
          Me.Root = New DevExpress.XtraLayout.LayoutControlGroup()
          CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
          CType(Me.Root, System.ComponentModel.ISupportInitialize).BeginInit()
          Me.SuspendLayout()
          '
          'LayoutControl1
          '
          Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
          Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
          Me.LayoutControl1.Name = "LayoutControl1"
          Me.LayoutControl1.Root = Me.Root
          Me.LayoutControl1.Size = New System.Drawing.Size(626, 418)
          Me.LayoutControl1.TabIndex = 0
          Me.LayoutControl1.Text = "LayoutControl1"
          '
          'Root
          '
          Me.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
          Me.Root.GroupBordersVisible = False
          Me.Root.Name = "Root"
          Me.Root.Size = New System.Drawing.Size(626, 418)
          Me.Root.TextVisible = False
          '
          'ucSelectorQuery
          '
          Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
          Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
          Me.Controls.Add(Me.LayoutControl1)
          Me.Name = "ucSelectorQuery"
          Me.Size = New System.Drawing.Size(626, 418)
          CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
          CType(Me.Root, System.ComponentModel.ISupportInitialize).EndInit()
          Me.ResumeLayout(False)

     End Sub

     Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
     Friend WithEvents Root As DevExpress.XtraLayout.LayoutControlGroup
End Class
