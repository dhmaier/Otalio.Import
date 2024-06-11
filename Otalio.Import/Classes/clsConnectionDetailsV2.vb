Imports System.Xml.Serialization
Imports System.ComponentModel

<XmlInclude(GetType(clsConnectionDetailsV2)), Serializable()>
Public Class clsConnectionDetailsV2
     Implements INotifyPropertyChanged

     ' Implement the INotifyPropertyChanged event
     Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

     ' Helper method to raise the PropertyChanged event
     Protected Sub OnPropertyChanged(propertyName As String)
          RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
     End Sub

     ' Properties
     Private ID As String = System.Guid.NewGuid.ToString
     Public Property _ID As String
          Get
               Return ID
          End Get
          Set(value As String)
               If Not ID.Equals(value) Then
                    ID = value
                    OnPropertyChanged(NameOf(_ID))
               End If
          End Set
     End Property

     Private ServerAddress As String = ""
     Public Property _ServerAddress As String
          Get
               Return ServerAddress
          End Get
          Set(value As String)
               If Not ServerAddress.Equals(value) Then
                    ServerAddress = value
                    OnPropertyChanged(NameOf(_ServerAddress))
               End If
          End Set
     End Property

     Private UserName As String = ""
     Public Property _UserName As String
          Get
               Return UserName
          End Get
          Set(value As String)
               If Not UserName.Equals(value) Then
                    UserName = value
                    OnPropertyChanged(NameOf(_UserName))
               End If
          End Set
     End Property

     Private UserPwd As String = ""
     Public Property _UserPwd As String
          Get
               Return UserPwd
          End Get
          Set(value As String)
               If Not UserPwd.Equals(value) Then
                    UserPwd = value
                    OnPropertyChanged(NameOf(_UserPwd))
               End If
          End Set
     End Property

     Private DateLastUsed As DateTime = Now
     Public Property _DateLastUsed As DateTime
          Get
               Return DateLastUsed
          End Get
          Set(value As DateTime)
               If Not DateLastUsed.Equals(value) Then
                    DateLastUsed = value
                    OnPropertyChanged(NameOf(_DateLastUsed))
               End If
          End Set
     End Property

     Private HTTPUserName As String = ""
     Public Property _HTTPUserName As String
          Get
               Return HTTPUserName
          End Get
          Set(value As String)
               If Not HTTPUserName.Equals(value) Then
                    HTTPUserName = value
                    OnPropertyChanged(NameOf(_HTTPUserName))
               End If
          End Set
     End Property

     Private HTTPPassword As String = ""
     Public Property _HTTPPassword As String
          Get
               Return HTTPPassword
          End Get
          Set(value As String)
               If Not HTTPPassword.Equals(value) Then
                    HTTPPassword = value
                    OnPropertyChanged(NameOf(_HTTPPassword))
               End If
          End Set
     End Property

     ' ReadOnly Property
     Public ReadOnly Property Key As String
          Get
               Return String.Format("{0}@{1}", UserName, ServerAddress)
          End Get
     End Property

     ' Method to clone the object
     Public Function Clone() As clsConnectionDetailsV2
          Dim oObject As clsConnectionDetailsV2 = DirectCast(Me.MemberwiseClone(), clsConnectionDetailsV2)
          oObject.ID = System.Guid.NewGuid.ToString
          Return oObject
     End Function

End Class
