Imports System.Net
Imports System.IO
Imports System.Linq




Public Class ucConnectionDetails

  Private moConnection As New clsConnectionDetails
  Private moConnectionHistory As New clsConnectionHistory
  Private moLastUsedConnection As New clsConnectionDetails

  Public Property _Connection As clsConnectionDetails
    Get
      Return moConnection
    End Get
    Set(value As clsConnectionDetails)
      If value IsNot Nothing Then

        moConnection = value

        If moConnection._HTTPUserName Is Nothing Then moConnection._HTTPUserName = ""
        If moConnection._HTTPPassword Is Nothing Then moConnection._HTTPPassword = ""

        BindConnection()

      End If
    End Set
  End Property


  Private Sub btnSaveConnection_Click(sender As Object, e As EventArgs) Handles btnSaveConnection.Click

    'create it

    'If moLastUsedConnection.Key <> moConnection.Key Then

    If moConnectionHistory.ConnectionHistory IsNot Nothing AndAlso moConnectionHistory.ConnectionHistory.Count > 0 Then
      'check if connection already exist in History
      Dim oConnection As clsConnectionDetails = TryCast(moConnectionHistory.ConnectionHistory.Where(Function(n) n.Key = moConnection.Key).FirstOrDefault, clsConnectionDetails)
      If oConnection IsNot Nothing Then
        moConnectionHistory.ConnectionHistory.Remove(oConnection)
        'update it
        moConnection._DateLastUsed = Now
        moConnectionHistory.ConnectionHistory.Add(moConnection.Clone)

      Else
        'add it.
        moConnection._DateLastUsed = Now
        moConnectionHistory.ConnectionHistory.Add(moConnection.Clone)
      End If
      End If

    'End If


    moLastUsedConnection = moConnection
    goConnection = moConnection

    SaveFile(GetAppPath(gsSettingFileName), moConnectionHistory)
    gridHistory.RefreshDataSource()
    gdHistory.BestFitColumns()
    gdHistory.ExpandAllGroups()

  End Sub

  Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click

    goHTTPServer.TestConnection(False,, True)

  End Sub

  Public Function LoadSettingFile(Optional pbSilent As Boolean = False) As Boolean

    Dim sFullPath As String = GetAppPath(gsSettingFileNameOld)

    ''check if old setting file exists
    If File.Exists(sFullPath) Then
      'convert it new format
      Dim oConnnection As clsConnectionDetails = TryCast(LoadFile(sFullPath), clsConnectionDetails)
      If oConnnection IsNot Nothing Then

        If oConnnection._ID Is Nothing Then
          oConnnection._ID = System.Guid.NewGuid.ToString
        End If

        Dim oNewFormat As New clsConnectionHistory
        oNewFormat.ActiveConnection = oConnnection
        If SaveFile(GetAppPath(gsSettingFileName), oNewFormat) Then
          File.Delete(GetAppPath(gsSettingFileNameOld))
        End If

      End If

    End If

    sFullPath = GetAppPath(gsSettingFileName)


    If File.Exists(sFullPath) Then
      'load it
      moConnectionHistory = TryCast(LoadFile(sFullPath), clsConnectionHistory)
      If moConnectionHistory IsNot Nothing Then

        If moConnectionHistory.ActiveConnection Is Nothing Then
          If moConnectionHistory.ConnectionHistory IsNot Nothing AndAlso moConnectionHistory.ConnectionHistory.Count > 0 Then
            moConnectionHistory.ActiveConnection = moConnectionHistory.ConnectionHistory(moConnectionHistory.ConnectionHistory.Count - 1)
          Else
            moConnectionHistory.ActiveConnection = New clsConnectionDetails
          End If
        End If

        If moConnectionHistory.ActiveConnection._ID Is Nothing Then
          moConnectionHistory.ActiveConnection._ID = System.Guid.NewGuid.ToString
        End If

        moConnection = moConnectionHistory.ActiveConnection
        goConnection = moConnectionHistory.ActiveConnection

        BindConnection()

        'check if connection already exist in History
        Dim oConnection As clsConnectionDetails = TryCast(moConnectionHistory.ConnectionHistory.Where(Function(n) n.Key = moConnection.Key).FirstOrDefault, clsConnectionDetails)

        If oConnection IsNot Nothing Then
          oConnection._DateLastUsed = Now
        Else
          moConnection._DateLastUsed = Now
          moConnectionHistory.ConnectionHistory.Add(moConnection.Clone)
        End If

        gridHistory.DataSource = moConnectionHistory.ConnectionHistory
        gdHistory.BestFitColumns()

        goHTTPServer.TestConnection(pbSilent,, True)

        moLastUsedConnection = moConnectionHistory.ActiveConnection.Clone


      Else
        MsgBox("Failed to load connection file")
      End If
    Else
      'create it
      If goConnection IsNot Nothing Then
        SaveFile(sFullPath, moConnectionHistory)
      End If
    End If


  End Function

  Private Sub gridHistory_Click(sender As Object, e As EventArgs) Handles gridHistory.Click

  End Sub

  Private Sub gridHistory_DoubleClick(sender As Object, e As EventArgs) Handles gridHistory.DoubleClick

    If gdHistory.SelectedRowsCount > 0 Then
      Dim oConnection As clsConnectionDetails = gdHistory.GetRow(gdHistory.GetSelectedRows(0))
      If oConnection IsNot Nothing Then

        moConnectionHistory.ActiveConnection = oConnection.Clone
        moConnection = moConnectionHistory.ActiveConnection
        goConnection = moConnection

        BindConnection()

      End If

    End If

  End Sub

  Private Sub BindConnection()

    txtServerName.DataBindings.Clear()
    txtServerName.DataBindings.Add(New Binding("Text", moConnection, "_ServerAddress"))

    txtUserName.DataBindings.Clear()
    txtUserName.DataBindings.Add(New Binding("Text", moConnection, "_UserName"))

    txtPassword.DataBindings.Clear()
    txtPassword.DataBindings.Add(New Binding("Text", moConnection, "_UserPwd"))

    With txtHTTPSUserName
      .DataBindings.Clear()
      .DataBindings.Add(New Binding("Text", moConnection, "_HTTPUserName"))
    End With

    With txtHTTPSPassword
      .DataBindings.Clear()
      .DataBindings.Add(New Binding("Text", moConnection, "_HTTPPassword"))
    End With

    With chkEnableLoging
      .DataBindings.Clear()
      .DataBindings.Add(New Binding("Checked", moConnectionHistory, "LoggedEvents"))
    End With

  End Sub
End Class
