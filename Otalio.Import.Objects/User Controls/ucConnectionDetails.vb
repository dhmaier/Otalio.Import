Imports System.Net
Imports System.IO


Public Class ucConnectionDetails

  Private moConnection As New clsConnectionDetails

  Public Property _Connection As clsConnectionDetails
    Get
      Return moConnection
    End Get
    Set(value As clsConnectionDetails)
      If value IsNot Nothing Then
        moConnection = value
        txtServerName.DataBindings.Clear()
        txtServerName.DataBindings.Add(New Binding("Text", moConnection, "_ServerAddress"))
        txtUserName.DataBindings.Clear()
        txtUserName.DataBindings.Add(New Binding("Text", moConnection, "_UserName"))
        txtPassword.DataBindings.Clear()
        txtPassword.DataBindings.Add(New Binding("Text", moConnection, "_UserPwd"))
      End If
    End Set
  End Property


  Private Sub btnSaveConnection_Click(sender As Object, e As EventArgs) Handles btnSaveConnection.Click

    'create it
    If moConnection IsNot Nothing Then
      SaveFile(GetAppPath(gsSettingFileName), moConnection)
    End If

  End Sub

  Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click

    goHTTPServer.TestConnection(False)

  End Sub

  Public Function LoadSettingFile() As Boolean

    Dim sFullPath As String = GetAppPath(gsSettingFileName)

    If File.Exists(sFullPath) Then
      'load it
      Dim oConnnection As clsConnectionDetails = TryCast(LoadFile(sFullPath), clsConnectionDetails)
      If oConnnection IsNot Nothing Then
        moConnection = oConnnection
        txtServerName.DataBindings.Clear()
        txtServerName.DataBindings.Add(New Binding("Text", moConnection, "_ServerAddress"))
        txtUserName.DataBindings.Clear()
        txtUserName.DataBindings.Add(New Binding("Text", moConnection, "_UserName"))
        txtPassword.DataBindings.Clear()
        txtPassword.DataBindings.Add(New Binding("Text", moConnection, "_UserPwd"))

        goConnection = oConnnection
        goHTTPServer.TestConnection()

      Else
        MsgBox("Failed to load connection file")
      End If
    Else
      'create it
      If goConnection IsNot Nothing Then
        SaveFile(sFullPath, moConnection)
      End If
    End If


  End Function

End Class
