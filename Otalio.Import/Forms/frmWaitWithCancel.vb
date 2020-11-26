Public Class frmWaitWithCancel



     Sub New()
          InitializeComponent()
          Me.progressPanel1.AutoHeight = True
     End Sub

     Public Overrides Sub SetCaption(ByVal caption As String)
          MyBase.SetCaption(caption)
          Me.progressPanel1.Caption = caption
     End Sub

     Public Overrides Sub SetDescription(ByVal description As String)
          MyBase.SetDescription(description)
          Me.progressPanel1.Description = description
     End Sub

     Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
          MyBase.ProcessCommand(cmd, arg)
     End Sub

     Public Enum WaitFormCommand
          SomeCommandId
     End Enum

     Private Sub butCancel_Click(sender As Object, e As EventArgs) Handles butCancel.Click
          mbCancel = True
     End Sub
End Class
