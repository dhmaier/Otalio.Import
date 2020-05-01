Public Class clsLogs

     Public Property Id As String = ""
     Public Property User As String = ""
     Public Property Message As String = ""
     Public Property ClassName As String = ""
     Public Property ObjectId As String = ""
     Public Property ObjectFriendlyName As String = ""
     Public Property MicroserviceName As String = ""
     Public Property MicroserviceVersion As String = ""
     Public Property MicroserviceInstanceId As String = ""
     Public Property CurrentHierarchyId As String = ""
     Public Property CurrentLocation As String = ""
     Public Property Datetime As String = ""
     Public Property UtcDatetime As String = ""
     Public Property HostName As String = ""
     Public Property HostIpAddress As String = ""
     Public Property EventType As String = ""
     Public Property LogGroup As String = ""



     Public Sub New(psId As String, psMicroserviceName As String, psMicroserviceVersion As String, psMicroserviceInstanceId As String, psCurrentHierarchyId As String,
                    psCurrentLocation As String, psDatetime As String, psUtcDatetime As String, psHostName As String, psHostIpAddress As String,
                    psUser As String, psEventType As String, psLogGroup As String, psMessage As String, psClassName As String, psObjectId As String,
                    psObjectFriendlyName As String)


          Me.id = psId
          Me.microserviceName = psMicroserviceName
          Me.microserviceVersion = psMicroserviceVersion
          Me.microserviceInstanceId = psMicroserviceInstanceId
          Me.currentHierarchyId = psCurrentHierarchyId
          Me.currentLocation = psCurrentLocation
          Me.datetime = psDatetime
          Me.utcDatetime = psUtcDatetime
          Me.hostName = psHostName
          Me.hostIpAddress = psHostIpAddress
          Me.user = psUser
          Me.eventType = psEventType
          Me.logGroup = psLogGroup
          Me.message = psMessage
          Me.className = psClassName
          Me.objectId = psObjectId
          Me.objectFriendlyName = psObjectFriendlyName


     End Sub

End Class



