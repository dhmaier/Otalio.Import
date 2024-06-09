Option Strict Off

Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Web.Script.Serialization




Public Class clsAPI1

     Private Const DefaultPageSize As Integer = 100


     Private msAccessToken As String
     Private msRefreshToken As String

     Private moResponseDataLoginResponse As Object
     Private moLoginResponse As Object
     Private moRefreshResponse As Object
     Private mdResponseDateTime As DateTime
     Private mnRefreshTokenCount As Long = 0
     Private mnTokenExpires As Long? = 0
     Private mnDefaultGetSize As Integer = 100
     Private mnTokenRefeshTime As Integer = 10
     Private mbLastLogIn As DateTime = Date.MinValue
     Private mbLastRefreshToken As DateTime = Date.MinValue

     Private WithEvents moTimmer As New Timer

     Public Event APICallEvent(psRequest As RestRequest, psResponse As IRestResponse)
     Public Event ErrorEvent(psExcetpion As Exception)

     Public ReadOnly Property RefreshInSeconds As Integer
          Get
               Try
                    Return ((mnTokenExpires - mnTokenRefeshTime) - CInt(DateTime.Now.Subtract(mdResponseDateTime).TotalSeconds))
               Catch ex As Exception
                    Return 0
               End Try

          End Get
     End Property
     Public ReadOnly Property LastLogIn As DateTime
          Get
               Return mbLastLogIn
          End Get
     End Property

     Public ReadOnly Property LastRefreshToken As DateTime
          Get
               Return mbLastRefreshToken
          End Get
     End Property


     Public Sub New()

          'need for connection to our webservices
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls

     End Sub


     Private Sub moTimmer_Tick(sender As Object, e As EventArgs) Handles moTimmer.Tick
          RefreshIAMConnection()
     End Sub

     Private Function ExecuteAPI(ByVal oClient As RestClient, ByVal oRequest As RestRequest, ByVal isLogin As Boolean) As IRestResponse
          Dim nRetryAttempts = 0
          Try

RetryAttempt:
               If isLogin Then
                    oRequest.AddHeader("Authorization", "Basic " & Base64Encode(goConnection._UserName & ":" + goConnection._UserPwd))
               Else
                    'Clipboard.SetText("Bearer " & msAccessToken)
                    oRequest.AddHeader("Authorization", "Bearer " & msAccessToken)

               End If

               If goConnection._HTTPUserName <> String.Empty Then
                    oClient.Authenticator = New RestSharp.Authenticators.HttpBasicAuthenticator(goConnection._HTTPUserName, goConnection._HTTPPassword)
               End If

               Dim oResponse As IRestResponse = oClient.Execute(oRequest)
               If oResponse.IsSuccessful = False And nRetryAttempts < 1 Then
                    Select Case oResponse.StatusCode
                         Case HttpStatusCode.Forbidden, HttpStatusCode.Unauthorized
                              RefreshIAMConnection()
                              nRetryAttempts += 1
                              GoTo RetryAttempt
                    End Select
               End If

               Return oResponse

          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function TestConnection(Optional pbSilent As Boolean = True, Optional poConnection As clsConnectionDetails = Nothing, Optional psLoadData As Boolean = False) As Boolean
          Try
               If poConnection IsNot Nothing Then
                    goConnection = poConnection
               End If

               Dim eStatus As HttpStatusCode = LogIntoIAM(True, pbSilent)

               If eStatus = HttpStatusCode.OK Then
                    If pbSilent = False Then MsgBox("System successfully tested connection to server")
                    If psLoadData Then
                         Call LoadLookupTypes()
                         Call LoadHierarchies()
                         Call LoadTranslateableLanguages()
                    End If

                    Return True
               Else
                    If pbSilent = False Then MsgBox(String.Format("Server returned following error code {0}", eStatus))
                    Return False
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return False

     End Function

     Public Function LogIntoIAM(Optional pbTestOnly As Boolean = False, Optional pbSilent As Boolean = True) As HttpStatusCode
          Try
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

                         mnTokenExpires = CInt(moLoginResponse.SelectToken("expires_in"))
                         msAccessToken = moLoginResponse.SelectToken("access_token")
                         msRefreshToken = moLoginResponse.SelectToken("refresh_token")
                         mnRefreshTokenCount = 0
                         mbLastLogIn = Now


                         If moTimmer.Enabled = False Then
                              moTimmer.Interval = 1000
                              moTimmer.Start()
                         End If

                    Case Else

                         If pbSilent = False Then
                              MsgBox(String.Format("Error connecting, request was {0} - {1}", oResponse.StatusCode, oResponse.StatusDescription))
                         End If

               End Select

               Return oResponse.StatusCode
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Private Sub RefreshIAMConnection()
          Try

               If moLoginResponse IsNot Nothing Then



                    If (mnTokenExpires - DateTime.Now.Subtract(mdResponseDateTime).TotalSeconds <= mnTokenRefeshTime) Then

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
                                   mnTokenExpires = CInt(moRefreshResponse.SelectToken("expires_in"))
                                   msAccessToken = moRefreshResponse.SelectToken("access_token")
                                   msRefreshToken = moRefreshResponse.SelectToken("refresh_token")
                                   mbLastRefreshToken = Now


                                   If RefreshInSeconds <= 0 Then
                                        'try to login again
                                        LogIntoIAM()
                                   End If

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

                              'try to login again
                              LogIntoIAM()

                         End Try
                    Else
                    End If
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try


     End Sub

     Public Function CallWebEndpointUsingPost(psEndPoint As String, psDTOJson As String) As IRestResponse
          Try
               ExtractSystemVariables(psEndPoint)
               ExtractSystemVariables(psDTOJson)

               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
               Dim serialized = jsSerializer.Serialize(psDTOJson)
               Dim oClient = New RestClient(url)


               Dim oRequest = New RestRequest(Method.POST)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
               oRequest.AddParameter("application/json", psDTOJson, ParameterType.RequestBody)
               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUsingGet(psEndPoint As String, psHeader As String, psQuery As String, Optional psSort As String = "", Optional psColumnsInclude As String = "", Optional pnPage As Integer = -1, Optional pnSize As Integer = -1) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psHeader = ExtractSystemVariables(psHeader)
               psQuery = ExtractSystemVariables(psQuery)
               psSort = ExtractSystemVariables(psSort)


               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)
               Dim sQuery As String = psQuery.ToString.Trim

               If Not String.IsNullOrEmpty(sQuery) Then

                    If sQuery.Contains("??") Then
                         Dim sQueryValues As List(Of String) = sQuery.Split(New Char() {"?"}, StringSplitOptions.RemoveEmptyEntries).ToList
                         If sQueryValues IsNot Nothing Then
                              If sQueryValues.Count = 1 Then
                                   url = url + "?" & sQueryValues(0)
                                   sQuery = String.Empty
                              End If

                              If sQueryValues.Count = 2 Then
                                   url = url + "?" & sQueryValues(1)
                                   sQuery = sQueryValues(0)
                              End If
                         End If
                    End If



                    If String.IsNullOrEmpty(sQuery) = False Then
                         If url.Contains("?") Then
                              url = url + String.Format("&search={0}", System.Web.HttpUtility.UrlEncode(sQuery))
                         Else
                              url = url + String.Format("?search={0}", System.Web.HttpUtility.UrlEncode(sQuery))
                         End If
                    End If



               End If


               If pnPage > -1 Then
                    If url.Contains("?") Then
                         url = url & String.Format("&page={0}", pnPage)
                    Else
                         url = url & String.Format("?page={0}", pnPage)
                    End If
               End If

               If url.ToLower.Contains("size=") = False Then
                    If pnSize > -1 Then
                         If url.Contains("?") Then
                              url = url & String.Format("&size={0}", pnSize)
                         Else
                              url = url & String.Format("?size={0}", pnSize)
                         End If
                    End If
                    If pnSize = -2 Then
                         If url.Contains("?") Then
                              url = url & String.Format("&size={0}", mnDefaultGetSize)
                         Else
                              url = url & String.Format("?size={0}", mnDefaultGetSize)
                         End If
                    End If
               End If


               If psSort <> String.Empty Then
                    If url.Contains("?") Then
                         url = url + "&sort=" & psSort
                    Else
                         url = url + "?sort=" & psSort
                    End If
               End If

               Dim oClient = New RestClient(url)


               Dim oRequest = New RestRequest(Method.GET)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

               'include fields
               If String.IsNullOrEmpty(psColumnsInclude) = False Then
                    oRequest.AddParameter("include-fields", psColumnsInclude, ParameterType.HttpHeader)

               End If


               If psHeader IsNot Nothing AndAlso String.IsNullOrEmpty(psHeader) = False Then
                    Dim oJObject As JObject = JObject.Parse(psHeader)
                    For Each oProperty As JProperty In oJObject.Properties
                         If oProperty IsNot Nothing Then
                              oRequest.AddParameter(oProperty.Name, oProperty.Value, ParameterType.HttpHeader)
                         End If
                    Next
               End If


               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)


               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUsingPost(psEndPoint As String, psHeader As String, psQuery As String, Optional psSort As String = "", Optional psColumnsInclude As String = "", Optional pnPage As Integer = -1, Optional pnSize As Integer = -1) As IRestResponse
          Try
               ' Extract system variables from input parameters
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psHeader = ExtractSystemVariables(psHeader)
               psQuery = ExtractSystemVariables(psQuery)
               psSort = ExtractSystemVariables(psSort)

               ' Construct URL
               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               ' Create the payload with the query
               Dim payload As New JObject()
               If Not String.IsNullOrEmpty(psQuery) Then
                    payload = JObject.Parse(psQuery)
               Else
                    payload = New JObject()
               End If

               If pnPage > -1 Then
                    payload("page") = pnPage
               End If
               If pnSize > -1 Then
                    payload("size") = pnSize
               ElseIf pnSize = -2 Then
                    payload("size") = mnDefaultGetSize
               End If

               ' Convert psSort to JSON array and add to payload
               If Not String.IsNullOrEmpty(psSort) Then
                    Dim sortArray As New JArray()
                    Dim sortItems As String() = psSort.Split(","c)

                    For i As Integer = 0 To sortItems.Length - 1 Step 2
                         If i + 1 < sortItems.Length Then
                              Dim sortObject As New JObject()
                              sortObject("property") = sortItems(i).Trim()
                              sortObject("direction") = sortItems(i + 1).Trim()
                              sortArray.Add(sortObject)
                         End If
                    Next

                    payload("sort") = sortArray
               End If


               ' Create REST client and request
               Dim oClient = New RestClient(url)
               Dim oRequest = New RestRequest(Method.POST)
               oRequest.AddParameter("application/json", payload.ToString(), ParameterType.RequestBody)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

               ' Include fields
               If Not String.IsNullOrEmpty(psColumnsInclude) Then
                    oRequest.AddParameter("include-fields", psColumnsInclude, ParameterType.HttpHeader)
               End If

               ' Add headers
               If Not String.IsNullOrEmpty(psHeader) Then
                    Dim oJObject As JObject = JObject.Parse(psHeader)
                    For Each oProperty As JProperty In oJObject.Properties
                         oRequest.AddParameter(oProperty.Name, oProperty.Value, ParameterType.HttpHeader)
                    Next
               End If

               ' Execute API call
               Dim oResponse = ExecuteAPI(oClient, oRequest, False)
               RaiseEvent APICallEvent(oRequest, oResponse)
               Return oResponse

          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing
     End Function


     Public Function CallWebEndpointUsingDelete(psEndPoint As String, psQuery As String) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psQuery = ExtractSystemVariables(psQuery)


               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               If Not String.IsNullOrEmpty(psQuery) Then
                    url = url + String.Format("?{0}", (psQuery))
               End If

               Dim oClient = New RestClient(url)


               Dim oRequest = New RestRequest(Method.DELETE)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUsingDeleteByID(psEndPoint As String, psID As String, psQuery As String) As IRestResponse
          Try
               psQuery = ExtractSystemVariables(psQuery)
               psEndPoint = ExtractSystemVariables(psEndPoint)


               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               If Not String.IsNullOrEmpty(psID) Then
                    url = url + String.Format("/{0}", (psID))
               End If

               If Not String.IsNullOrEmpty(psQuery) Then
                    url = url + String.Format("?{0}", (psQuery))
               End If


               Dim oClient = New RestClient(url)


               Dim oRequest = New RestRequest(Method.DELETE)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUsingPut(psEndPoint As String, psEntityID As String, psQuery As String, psDTOJson As String) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psQuery = ExtractSystemVariables(psQuery)
               psDTOJson = ExtractSystemVariables(psDTOJson)

               Dim url As String = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID)

               If Not String.IsNullOrEmpty(psQuery) Then
                    If psQuery.Contains("??") Then
                         url = url + String.Format("{0}", Replace((psQuery.Trim), "??", String.Empty))
                    Else
                         url = url + String.Format("?{0}", (psQuery.Trim))
                    End If


               End If

               Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
               Dim serialized = jsSerializer.Serialize(psDTOJson)
               Dim oClient = New RestClient(url)



               Dim oRequest = New RestRequest(Method.PUT)

               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
               If psDTOJson IsNot Nothing Then
                    oRequest.AddParameter("application/json", psDTOJson, ParameterType.RequestBody)
               End If

               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUsingPatch(psEndPoint As String, psEntityID As String, psQuery As String, psDTOJson As String) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psQuery = ExtractSystemVariables(psQuery)
               psDTOJson = ExtractSystemVariables(psDTOJson)

               Dim url As String = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID)

               If Not String.IsNullOrEmpty(psQuery) Then
                    If psQuery.Contains("??") Then
                         url = url + String.Format("{0}", Replace((psQuery.Trim), "??", String.Empty))
                    Else
                         url = url + String.Format("?{0}", (psQuery.Trim))
                    End If


               End If

               Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
               Dim serialized = jsSerializer.Serialize(psDTOJson)
               Dim oClient = New RestClient(url)



               Dim oRequest = New RestRequest(Method.PATCH)

               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
               If psDTOJson IsNot Nothing Then
                    oRequest.AddParameter("application/json", psDTOJson, ParameterType.RequestBody)
               End If

               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function CallWebEndpointUploadFile(psEndPoint As String, psEntityID As String, psFileName As String, psObject As String) As IRestResponse

          Dim bFileContent As Byte() = Nothing

          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               If File.Exists(psFileName) Then

                    Select Case Path.GetExtension(psFileName).ToString.ToLower
                         Case ".jpg" : bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Jpeg)
                         Case ".png" : bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Png)
                         Case ".bmp" : bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Bmp)
                         Case Else : bFileContent = System.IO.File.ReadAllBytes(psFileName)
                    End Select

                    If bFileContent IsNot Nothing Then

                         Dim url As String = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID, psObject)
                         Dim oClient = New RestClient(url)

                         Dim oRequest = New RestRequest(psObject, Method.POST)
                         ' oRequest.AddHeader("Content-Type", "multipart/form-data")
                         oRequest.AddHeader("Accept-Encoding", ": gzip, deflate, br")
                         oRequest.AddFile("file", bFileContent, Path.GetFileName(psFileName), String.Format("image/{0}", Path.GetExtension(psFileName)))
                         Dim oResponse = ExecuteAPI(oClient, oRequest, False)

                         RaiseEvent APICallEvent(oRequest, oResponse)

                         Return oResponse
                    End If
               End If

          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          Finally
               bFileContent = Nothing
          End Try

          Return Nothing

     End Function

     Public Function GetValueFromEndpoint(psEndPoint As String, psQuery As String, psNodeName As String, psRootNode As String) As String
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psQuery = ExtractSystemVariables(psQuery)

               Dim sReturnValue As String = String.Empty

               Dim oResponse As IRestResponse = CallWebEndpointUsingGet(psEndPoint, String.Empty, psQuery)
               Dim json As JObject = JObject.Parse(oResponse.Content)
               If json IsNot Nothing Then

                    Dim sRoot As String = String.Empty
                    Dim sNode As String = String.Empty

                    If String.IsNullOrEmpty(psRootNode) = False AndAlso psRootNode.Contains("[0]") Then
                         sRoot = psRootNode.Substring(0, psRootNode.LastIndexOf("[0]"))
                         sNode = psNodeName ' psRootNode.Substring(psRootNode.LastIndexOf("]") + 2)
                    End If

                    If oResponse.StatusCode = HttpStatusCode.OK And json.ContainsKey("responsePayload") = True Then
                         Dim oObject As JToken = json.SelectToken(sRoot)
                         If oObject IsNot Nothing Then
                              If oObject.Count > 0 Then

                                   Try
                                        If sNode = "description" Then
                                             'try to exact it from the object
                                             Dim oToken As JToken = json.SelectToken(String.Format("{0}[0].translations.en.description", sRoot, sNode))
                                             If oToken IsNot Nothing Then
                                                  Return oToken.ToString
                                             End If
                                        End If

                                   Catch ex As Exception

                                   End Try


                                   Try
                                        If oObject.HasValues Then
                                             If oObject.First IsNot Nothing Then
                                                  sReturnValue = oObject.First.SelectToken(sNode).Value(Of String)
                                             End If
                                        End If
                                   Catch ex As Exception

                                   End Try

                                   'Extract the value
                                   'If json.ContainsKey(psNodeName) Then
                                   '     sReturnValue = json.SelectToken(psNodeName).ToString
                                   'End If

                              End If
                         End If
                    Else

                         sReturnValue = oResponse.ErrorMessage

                    End If


               End If


               Return sReturnValue
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Private Sub LoadLookupTypes()
          Try
               goLookupTypes.Clear()

               If goLookupTypes IsNot Nothing AndAlso goLookupTypes.Count = 0 Then
                    Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet("metadata/v1/lookup-types", String.Empty, String.Empty, "code")
                    If oResponse IsNot Nothing Then

                         Dim json As JObject = JObject.Parse(oResponse.Content)
                         Dim oOjbect As JToken = json.SelectToken("responsePayload.content")
                         If oOjbect IsNot Nothing Then
                              For Each oNode In oOjbect
                                   If oNode IsNot Nothing Then
                                        goLookupTypes.Add(New clsLookUpTypes With {.Id = oNode("id"), .Code = oNode("code"), .Description = oNode("description"), .LookupGroup = oNode("lookupGroup")})
                                   End If
                              Next
                         End If
                         json = Nothing
                         oOjbect = Nothing

                         If goLookupTypes IsNot Nothing Then
                              goLookupTypes = goLookupTypes.OrderBy(Function(n) n.Description).ToList
                         End If

                    End If

                    oResponse = Nothing

               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

     End Sub

     Private Sub LoadTranslateableLanguages()
          goLanguages.Clear()

          If goLanguages IsNot Nothing AndAlso goLanguages.Count = 0 Then
               Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet("metadata/v1/languages", String.Empty, "translationEnabled==true", "code")
               If oResponse IsNot Nothing Then

                    Dim json As JObject = JObject.Parse(oResponse.Content)
                    Dim oOjbect As JToken = json.SelectToken("responsePayload.content")
                    If oOjbect IsNot Nothing Then
                         For Each oNode In oOjbect
                              If oNode IsNot Nothing Then
                                   goLanguages.Add(New clsLanguages With {.id = oNode("id"), .code = oNode("code"), .translations = GetTranslation(oNode.SelectToken("translations.en.description"))})
                              End If
                         Next
                    End If
                    json = Nothing
                    oOjbect = Nothing

               End If

               oResponse = Nothing

          End If

     End Sub
     Public Function CallGraphQL(psEndPoint As String, psDTOJson As String) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psDTOJson = ExtractSystemVariables(psDTOJson)


               psDTOJson = Replace(psDTOJson, vbNewLine, "")
               psDTOJson = Replace(psDTOJson, ControlChars.Quote, "\" & ControlChars.Quote)


               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               Dim jsSerializer As JavaScriptSerializer = New JavaScriptSerializer()
               Dim serialized = jsSerializer.Serialize(psDTOJson)
               Dim oClient = New RestClient(url)
               Dim sBody As String = Replace(String.Format("{1}{3}query{3}:{3}{0}{3}{2}", psDTOJson, "{", "}", ControlChars.Quote), vbNewLine, String.Empty)

               Dim oRequest = New RestRequest(Method.POST)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)
               oRequest.AddParameter("Accept-Encoding", "gzip,deflate,br", ParameterType.HttpHeader)
               oRequest.AddParameter("Content-Type", "application/json", ParameterType.HttpHeader)

               oRequest.AddParameter("application/json", sBody, ParameterType.RequestBody)
               Dim oResponse = ExecuteAPI(oClient, oRequest, False)

               RaiseEvent APICallEvent(oRequest, oResponse)

               Return oResponse
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Private Sub LoadHierarchies()
          Try

               goHierarchies.Clear()

               If goHierarchies IsNot Nothing AndAlso goHierarchies.Count = 0 Then
                    Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet("metadata/v1/hierarchies", String.Empty, "")
                    If oResponse IsNot Nothing Then

                         Dim json As JObject = JObject.Parse(oResponse.Content)
                         Dim oOjbect As JToken = json.SelectToken("responsePayload")

                         For Each oNode In oOjbect
                              If oNode IsNot Nothing Then

                                   Dim sDescription As JToken = oNode.SelectToken("translations.en.description")
                                   If oNode.SelectToken("deleted").ToString.ToLower = "false" Then
                                        goHierarchies.Add(New clsHierarchies With {.Id = oNode("id"), .Code = oNode("code"), .Description = sDescription, .ParentHierarchyId = oNode("parentHierarchyId"),
                                                  .Enabled = IIf(oNode("enabled").ToString.ToUpper = "TRUE", True, False), .Level = oNode("level"), .Priority = oNode("priority")})
                                   End If

                              End If
                         Next

                         json = Nothing
                         oOjbect = Nothing

                         'check for children
                         For Each oH As clsHierarchies In goHierarchies

                              oH.ParentName = LookforParentHierarchies(oH, String.Empty)

                         Next


                         If goHierarchies IsNot Nothing Then
                              goHierarchies = goHierarchies.OrderBy(Function(n) n.ParentName).ThenBy(Function(n) n.Description).ToList
                         End If



                    End If
                    oResponse = Nothing
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try


     End Sub

     Private Function LookforParentHierarchies(poHierarchy As clsHierarchies, psParentString As String) As String
          Try

               If poHierarchy.ParentHierarchyId IsNot Nothing AndAlso poHierarchy.ParentHierarchyId.ToString <> String.Empty Then

                    Dim oParent As clsHierarchies = goHierarchies.Where(Function(n) n.Id = poHierarchy.ParentHierarchyId).FirstOrDefault
                    If oParent IsNot Nothing Then
                         If oParent.Id <> "1" Then
                              If oParent IsNot Nothing AndAlso oParent.Description <> String.Empty Then
                                   If psParentString = String.Empty Then
                                        psParentString = oParent.Description
                                   Else
                                        psParentString = oParent.Description & "\" & psParentString
                                   End If
                              End If

                              If oParent.ParentHierarchyId IsNot Nothing AndAlso oParent.ParentHierarchyId.ToString <> String.Empty Then
                                   psParentString = LookforParentHierarchies(oParent, psParentString)
                              End If
                         Else
                              If psParentString = String.Empty Then
                                   psParentString = "Enterprise"
                              Else
                                   psParentString = "Enterprise" & " \ " & psParentString
                              End If
                         End If
                    End If
               End If

               Return psParentString
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return Nothing

     End Function

     Public Function ExtractSystemVariables(psObject As String) As String

          Try


               'check for Hierarchy
               If psObject IsNot Nothing AndAlso String.IsNullOrEmpty(psObject) = False Then

                    If psObject.ToString.ToUpper.Contains("@@HIERARCHYID@@") = True Then psObject = Replace(psObject, "@@HIERARCHYID@@", gsSelectedHierarchy.ToString)
                    If psObject.ToString.ToUpper.Contains("@@EMPTYARRAY@@") = True Then psObject = Replace(psObject, ControlChars.Quote & "@@EMPTYARRAY@@" & ControlChars.Quote, "[]")
                    If psObject.ToString.ToUpper.Contains("@@SHIPID@@") = True Then psObject = Replace(psObject, "@@SHIPID@@", gsSelectedHierarchy.ToString)
                    If psObject.ToString.ToUpper.Contains("@@RVCID@@") = True Then psObject = Replace(psObject, "@@RVCID@@", gsSelectedHierarchy.ToString)
                    If psObject.ToString.ToUpper.Contains("@@PARENTID@@") = True Then psObject = Replace(psObject, "@@PARENTID@@", gsSelectedHierarchyParent.ToString)

               End If

               Return psObject

          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return String.Empty

     End Function

     Public Function LogEvent(psMessage As String, psClassName As String, psObjectFriendlyName As String, psObjectID As String)


          Try
               Dim oJsn As New JObject
               oJsn.Add(New JProperty("eventType", "DATA_IMPORT"))

               Dim oLog As New JObject
               oLog.Add(New JProperty("className", psClassName))
               oLog.Add(New JProperty("message", psMessage))
               oLog.Add(New JProperty("objectFriendlyName", psObjectFriendlyName))
               oLog.Add(New JProperty("objectId", Replace(psObjectID, "-", "")))

               Dim oLogs As New JArray
               oLogs.Add(oLog)
               oJsn.Add(New JProperty("logEntities", oLogs))

               Dim oDTO As New JArray
               oDTO.Add(oJsn)

               Call CallWebEndpointUsingPost("metadata/v1/logs", oDTO.ToString)
          Catch ex As Exception

          End Try



     End Function

End Class
