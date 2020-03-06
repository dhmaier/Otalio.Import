Imports System.Xml.Serialization

<
  XmlInclude(GetType(clsConnectionDetails)),
  Serializable()
  >
Public Class clsConnectionDetails

  Public Property _ServerAddress As String = ""
  Public Property _UserName As String = ""
  Public Property _UserPwd As String = ""

End Class
