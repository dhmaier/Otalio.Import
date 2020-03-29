Imports DevExpress.XtraEditors

Public Class frmSplashScreen
  Sub New()
    InitializeComponent()
    Me.labelCopyright.Text = "Copyright © 2018 - " & DateTime.Now.Year.ToString()
  End Sub

  Public Overrides Sub ProcessCommand(ByVal cmd As System.Enum, ByVal arg As Object)
    MyBase.ProcessCommand(cmd, arg)
    Dim oCommand As SplashScreenCommand = CType(cmd, SplashScreenCommand)
    Select Case oCommand
      Case SplashScreenCommand.SetProgress
        labelStatus.Text = arg.ToString
        Application.DoEvents()

    End Select
  End Sub

  Public Enum SplashScreenCommand
    SetProgress
  End Enum


End Class
