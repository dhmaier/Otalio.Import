Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsConnectionDetails)),
  Serializable()
  >
Public Class clsConnectionDetails
  Public Property _ID As String = System.Guid.NewGuid.ToString

  Public Property _ServerAddress As String = ""

  Public Property _UserName As String = ""

  Public Property _UserPwd As String = ""

  Public Property _DateLastUsed As DateTime = Now

  Public Property _HTTPUserName As String = ""

  Public Property _HTTPPassword As String = ""


  Public ReadOnly Property Key As String
    Get
      Return String.Format("{0}@{1}", _UserName, _ServerAddress)
    End Get
  End Property

  Public Function Clone() As clsConnectionDetails
    Dim oObject As clsConnectionDetails = DirectCast(Me.MemberwiseClone(), clsConnectionDetails)
    oObject._ID = System.Guid.NewGuid.ToString
    Return oObject
  End Function

End Class
