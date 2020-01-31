Imports System.ComponentModel
Imports System.Text
Imports DevExpress.Skins
Imports DevExpress.LookAndFeel
Imports DevExpress.UserSkins
Imports DevExpress.XtraBars.Helpers
Imports DevExpress.XtraBars.Ribbon


Public Class Form1
  Sub New()
    InitSkins()
    InitializeComponent()
    Me.InitSkinGallery()

  End Sub
  Sub InitSkins()
    DevExpress.Skins.SkinManager.EnableFormSkins()
    DevExpress.UserSkins.BonusSkins.Register()

  End Sub
  Private Sub InitSkinGallery()
    SkinHelper.InitSkinGallery(rgbiSkins, True)
  End Sub

End Class
