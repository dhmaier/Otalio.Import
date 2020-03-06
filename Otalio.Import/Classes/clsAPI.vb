Option Strict Off

Imports RestSharp
Imports Newtonsoft.Json
Imports System.Web.Script.Serialization
Imports System.Dynamic
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports System.Linq



Public Class clsAPI

  Private msAccessToken As String
  Private msRefreshToken As String

  Private moResponseDataLoginResponse As Object
  Private moLoginResponse As Object
  Private moRefreshResponse As Object
  Private mdResponseDateTime As DateTime
  Private mnRefreshTokenCount As Long = 0
  Private mnTokenExpires As Long? = 0

  Private moTimmer As New Timer

  Public Event APICallEvent(psRequest As RestRequest, psResponse As IRestResponse)

  Public Sub New()
    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
  End Sub

  Private Function ExecuteAPI(ByVal oClient As RestClient, ByVal oRequest As RestRequest, ByVal isLogin As Boolean) As IRestResponse
    If goConnection._UserName <> "" Then
      oRequest.AddHeader("Authorization", "Basic " & Base64Encode(goConnection._UserName & ":" + goConnection._UserPwd))

      If Not isLogin Then
        oRequest.AddHeader("X-API-Token", "Bearer " & msAccessToken)
      End If
    Else

      If Not isLogin Then
        oRequest.AddHeader("Authorization", "Bearer " & msAccessToken)
      End If
    End If

    Return oClient.Execute(oRequest)
  End Function

  Public Function TestConnection(Optional pbSilent As Boolean = True, Optional poConnection As clsConnectionDetails = Nothing) As Boolean

    If poConnection IsNot Nothing Then
      goConnection = poConnection
    End If

    Dim eStatus As HttpStatusCode = LogIntoIAM(True)

    If eStatus = HttpStatusCode.OK Then
      If pbSilent = False Then MsgBox("System sucessfully tested connection to server")
      Call LoadLookupTypes()
      Return True
    Else
      MsgBox(String.Format("Server returned following error code {0}", eStatus))
      Return False
    End If
  End Function

  Public Function LogIntoIAM(Optional pbTestOnly As Boolean = False) As HttpStatusCode

    Dim url As String = String.Format("{0}iam/v1/sso/login", goConnection._ServerAddress)
    Dim oColumns = New Dictionary(Of String, String)
    With oColumns
      .Add("login", goConnection._UserName)
      .Add("password", goConnection._UserPwd)
      .Add("rememberMe", "false")
    End With


    Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
    Dim serialized = jsSerializer.Serialize(oColumns)
    Dim oClient = New RestClient(url)


    Dim oRequest = New RestRequest(Method.POST)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
    oRequest.AddParameter("application/json", serialized, ParameterType.RequestBody)
    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    Select Case oResponse.StatusCode
      Case HttpStatusCode.OK

        Dim json As JObject = JObject.Parse(oResponse.Content)
        moLoginResponse = json.SelectToken("responsePayload")
        moResponseDataLoginResponse = json.SelectToken("message")
        mdResponseDateTime = DateTime.Now
        mnTokenExpires = json.SelectToken("message").SelectToken("expires_in")
        msAccessToken = json.SelectToken("message").SelectToken("access_token")
        msRefreshToken = json.SelectToken("message").SelectToken("refresh_token")

        If pbTestOnly = False Then
          moTimmer.Interval = 1000
          moTimmer.Start()
        End If

      Case Else
        MsgBox(String.Format("Error connecting, request was {0} - {1}", oResponse.StatusCode, oResponse.StatusDescription))
    End Select

    ' RaiseEvent APICallEvent(oRequest, oResponse)

    Return oResponse.StatusCode

  End Function

  Private Sub RefreshIAMConnection()


    If moLoginResponse IsNot Nothing Then



      If (mnTokenExpires - DateTime.Now.Subtract(mdResponseDateTime).TotalSeconds <= 10) Then

        Try
          Dim url As String = String.Format("{0}iam/v1/sso/refresh", goConnection._ServerAddress)
          Dim columns = New Dictionary(Of String, String) From {{"refreshToken", msRefreshToken}}

          Dim jsSerializer = New JavaScriptSerializer()
          Dim serialized = jsSerializer.Serialize(columns)
          Dim oClient = New RestClient(url)
          Dim oRequest = New RestRequest(Method.POST)

          oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
          oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
          oRequest.AddParameter("application/json", serialized, ParameterType.RequestBody)

          Dim oResponse = ExecuteAPI(oClient, oRequest, True)


          If oResponse.StatusCode = HttpStatusCode.OK Then



            Dim json As JObject = JObject.Parse(oResponse.Content)
            moRefreshResponse = json.SelectToken("responsePayload")
            mdResponseDateTime = DateTime.Now
            mnRefreshTokenCount += 1
            mnTokenExpires = json.SelectToken("message").SelectToken("expires_in")
            msAccessToken = json.SelectToken("message").SelectToken("access_token")
            msRefreshToken = json.SelectToken("message").SelectToken("refresh_token")

          End If

          ' RaiseEvent APICallEvent(oRequest, oResponse)

        Catch ex As Exception
          Dim index As Integer = ex.Message.IndexOf("{")

          If index >= 0 Then
            moTimmer.Stop()
            Dim jsonString As String = ex.Message.Substring(index)
            MessageBox.Show(jsonString)
          Else
            moTimmer.Stop()
            MessageBox.Show(ex.Message)
          End If
        End Try
      Else
      End If
    End If


  End Sub

  Public Function CallWebEndpointUsingPost(psEndPoint As String, psDTOJson As String) As IRestResponse

    Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

    Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
    Dim serialized = jsSerializer.Serialize(psDTOJson)
    Dim oClient = New RestClient(url)


    Dim oRequest = New RestRequest(Method.POST)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
    oRequest.AddParameter("application/json", psDTOJson, ParameterType.RequestBody)
    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    RaiseEvent APICallEvent(oRequest, oResponse)

    Return oResponse

  End Function

  Public Function CallWebEndpointUsingGet(psEndPoint As String, psHeader As String, psQuery As String, Optional psSort As String = "") As IRestResponse

    Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

    If Not String.IsNullOrEmpty(psQuery) Then
      If psQuery.Contains("??") Then
        url = url + Replace(psQuery, "??", "?")
      Else
        If url.Contains("?") Then
          url = url + String.Format("&search={0}", System.Web.HttpUtility.UrlEncode(psQuery))
        Else
          url = url + String.Format("?search={0}", System.Web.HttpUtility.UrlEncode(psQuery))
        End If
      End If

    End If

    If psSort <> "" Then
      url = url + "&sort=" & psSort
    End If

    Dim oClient = New RestClient(url)


    Dim oRequest = New RestRequest(Method.GET)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

    If psHeader IsNot Nothing AndAlso String.IsNullOrEmpty(psHeader) = False Then
      Dim oJObject As JObject = JObject.Parse(psHeader)
      For Each oProperty As JProperty In oJObject.Properties
        If oProperty IsNot Nothing Then
          oRequest.AddParameter(oProperty.Name, oProperty.Value, ParameterType.HttpHeader)
        End If
      Next
    End If


    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    RaiseEvent APICallEvent(oRequest, oResponse)


    Return oResponse

  End Function

  Public Function CallWebEndpointUsingDelete(psEndPoint As String, psQuery As String) As IRestResponse

    Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

    If Not String.IsNullOrEmpty(psQuery) Then
      url = url + String.Format("?{0}", (psQuery))
    End If

    Dim oClient = New RestClient(url)


    Dim oRequest = New RestRequest(Method.DELETE)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    RaiseEvent APICallEvent(oRequest, oResponse)

    Return oResponse

  End Function

  Public Function CallWebEndpointUsingPut(psEndPoint As String, psEntityID As String, psQuery As String, psDTOJson As String) As IRestResponse

    Dim url As String = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID)

    If Not String.IsNullOrEmpty(psQuery) Then
      url = url + String.Format("?{0}", (psQuery.Trim))
    End If

    Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
    Dim serialized = jsSerializer.Serialize(psDTOJson)
    Dim oClient = New RestClient(url)


    Dim oRequest = New RestRequest(Method.PUT)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
    oRequest.AddParameter("application/json", psDTOJson, ParameterType.RequestBody)

    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    RaiseEvent APICallEvent(oRequest, oResponse)

    Return oResponse

  End Function

  Public Function GetValueFromEndpoint(psEndPoint As String, psQuery As String, psNodeName As String) As String

    Dim sReturnValue As String = ""

    Dim oResponse As IRestResponse = CallWebEndpointUsingGet(psEndPoint, "", psQuery)
    Dim json As JObject = JObject.Parse(oResponse.Content)
    If json IsNot Nothing Then


      If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then
        Dim oObject As JToken = TryCast(json.SelectToken("responsePayload.content"), JToken)
        If oObject IsNot Nothing Then
          If oObject.Count > 0 Then

            'Extract the value
            sReturnValue = json.SelectToken(psNodeName).ToString


          End If
        End If
      Else
        sReturnValue = oResponse.ErrorMessage
      End If


    End If


    Return sReturnValue

  End Function

  Private Sub LoadLookupTypes()

    If goLookupTypes IsNot Nothing AndAlso goLookupTypes.Count = 0 Then
      Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet("metadata/v1/lookup-types", "", "")
      If oResponse IsNot Nothing Then

        Dim json As JObject = JObject.Parse(oResponse.Content)
        Dim oOjbect As JToken = TryCast(json.SelectToken("responsePayload.content"), JToken)

        For Each oNode In oOjbect
          If oNode IsNot Nothing Then
            goLookupTypes.Add(New clsLookUpTypes With {.Id = oNode("id"), .Code = oNode("code"), .Description = oNode("description"), .LookupGroup = oNode("lookupGroup")})
          End If
        Next

        json = Nothing
        oOjbect = Nothing

        If goLookupTypes IsNot Nothing Then
          goLookupTypes = goLookupTypes.OrderBy(Function(n) n.Description).ToList
        End If

      End If
      oResponse = Nothing
    End If
  End Sub

  Public Function CallGraphQL(psEndPoint As String, psDTOJson As String) As IRestResponse

    Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

    Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
    Dim serialized = jsSerializer.Serialize(psDTOJson)
    Dim oClient = New RestClient(url)
    Dim sBody As String = Replace(String.Format("{1}{3}query{3}:{3}{0}{3}{2}", psDTOJson, "{", "}", ControlChars.Quote), vbNewLine, "")

    Dim oRequest = New RestRequest(Method.POST)
    oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
    oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
    oRequest.AddParameter("Accept-Encoding", "gzip,deflate,br", ParameterType.HttpHeader)
    oRequest.AddParameter("Content-Type", "application/json", ParameterType.HttpHeader)

    oRequest.AddParameter("application/json", sBody, ParameterType.RequestBody)
    Dim oResponse = ExecuteAPI(oClient, oRequest, True)

    RaiseEvent APICallEvent(oRequest, oResponse)

    Return oResponse

  End Function


End Class
