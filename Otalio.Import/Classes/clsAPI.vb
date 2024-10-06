
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Timers

Public Class clsAPI

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
     Private mnTokenRefreshTime As Integer = 10
     Private mbLastLogIn As DateTime = Date.MinValue
     Private mbLastRefreshToken As DateTime = Date.MinValue

     Private WithEvents moTimer As New Timer

     Public Event APICallEvent(psRequest As RestRequest, psResponse As IRestResponse)
     Public Event ErrorEvent(psException As Exception)

     Public ReadOnly Property RefreshInSeconds As Integer
          Get
               Try
                    Return CInt((mnTokenExpires - mnTokenRefreshTime) - DateTime.Now.Subtract(mdResponseDateTime).TotalSeconds)
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
          ' Need for connection to our web services
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
     End Sub

     Private Sub moTimer_Tick(sender As Object, e As EventArgs) Handles moTimer.Elapsed
          RefreshIAMConnection()
     End Sub

     Private Function ExecuteAPI(ByVal oClient As RestClient, ByVal oRequest As RestRequest, ByVal isLogin As Boolean) As IRestResponse
          Dim nRetryAttempts = 0
          Try
RetryAttempt:
               If isLogin Then
                    oRequest.AddHeader("Authorization", "Basic " & Base64Encode(goConnection._UserName & ":" & goConnection._UserPwd))
               Else
                    oRequest.AddHeader("Authorization", "Bearer " & msAccessToken)
               End If

               If Not String.IsNullOrEmpty(goConnection._HTTPUserName) Then
                    oClient.Authenticator = New RestSharp.Authenticators.HttpBasicAuthenticator(goConnection._HTTPUserName, goConnection._HTTPPassword)
               End If

               Dim oResponse As IRestResponse = oClient.Execute(oRequest)
               If Not oResponse.IsSuccessful AndAlso nRetryAttempts < 1 Then
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
                    If Not pbSilent Then MsgBox("System successfully tested connection to server")
                    If psLoadData Then
                         LoadLookupTypes()
                         LoadHierarchies()
                         LoadTranslatableLanguages()
                    End If

                    Return True
               Else
                    If Not pbSilent Then MsgBox(String.Format("Server returned following error code {0}", eStatus))
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
               Dim oColumns = New Dictionary(Of String, String) From {
                   {"login", goConnection._UserName},
                   {"password", goConnection._UserPwd},
                   {"rememberMe", "false"}
               }

               Dim jsSerializer As New JavaScriptSerializer()
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

                         If Not moTimer.Enabled Then
                              moTimer.Interval = 1000
                              moTimer.Start()
                         End If

                    Case Else
                         If Not pbSilent Then
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
                    If mnTokenExpires - DateTime.Now.Subtract(mdResponseDateTime).TotalSeconds <= mnTokenRefreshTime Then
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
                                        ' Try to login again
                                        LogIntoIAM()
                                   End If
                              End If
                         Catch ex As Exception
                              HandleException(ex)
                              ' Try to login again
                              LogIntoIAM()
                         End Try
                    End If
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try
     End Sub

     Private Sub HandleException(ex As Exception)
          Dim index As Integer = ex.Message.IndexOf("{")

          If index >= 0 Then
               moTimer.Stop()
               Dim jsonString As String = ex.Message.Substring(index)
               MessageBox.Show(jsonString)
          Else
               moTimer.Stop()
               MessageBox.Show(ex.Message)
          End If
     End Sub

     Public Function CallWebEndpointUsingPost(psEndPoint As String, psDTOJson As String) As IRestResponse
          Try
               ExtractSystemVariables(psEndPoint)
               ExtractSystemVariables(psDTOJson)

               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)
               Dim jsSerializer As New JavaScriptSerializer()
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
               Dim sQuery As String = psQuery.Trim()

               If Not String.IsNullOrEmpty(sQuery) Then
                    If sQuery.Contains("??") Then
                         Dim sQueryValues As List(Of String) = sQuery.Split(New Char() {"?"}, StringSplitOptions.RemoveEmptyEntries).ToList()
                         If sQueryValues IsNot Nothing Then
                              If sQueryValues.Count = 1 Then
                                   url += "?" & sQueryValues(0)
                                   sQuery = String.Empty
                              ElseIf sQueryValues.Count = 2 Then
                                   url += "?" & sQueryValues(1)
                                   sQuery = sQueryValues(0)
                              End If
                         End If
                    End If

                    If Not String.IsNullOrEmpty(sQuery) Then
                         url += If(url.Contains("?"), "&", "?") & "search=" & WebUtility.UrlEncode(sQuery)
                    End If
               End If

               If pnPage > -1 Then
                    url += If(url.Contains("?"), "&", "?") & String.Format("page={0}", pnPage)
               End If

               If Not url.ToLower().Contains("size=") Then
                    If pnSize > -1 Then
                         url += If(url.Contains("?"), "&", "?") & String.Format("size={0}", pnSize)
                    ElseIf pnSize = -2 Then
                         url += If(url.Contains("?"), "&", "?") & String.Format("size={0}", mnDefaultGetSize)
                    End If
               End If

               If Not String.IsNullOrEmpty(psSort) Then
                    url += If(url.Contains("?"), "&", "?") & "sort=" & psSort
               End If

               Dim oClient = New RestClient(url)
               Dim oRequest = New RestRequest(Method.GET)
               oRequest.AddParameter("format", "json", ParameterType.UrlSegment)
               oRequest.AddParameter("Accepts", "application/json;charset=UTF-8", ParameterType.HttpHeader)

               If Not String.IsNullOrEmpty(psColumnsInclude) Then
                    oRequest.AddParameter("include-fields", psColumnsInclude, ParameterType.HttpHeader)
               End If

               If Not String.IsNullOrEmpty(psHeader) Then
                    Dim oJObject As JObject = JObject.Parse(psHeader)
                    For Each oProperty As JProperty In oJObject.Properties
                         oRequest.AddParameter(oProperty.Name, oProperty.Value, ParameterType.HttpHeader)
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
                    url += String.Format("?{0}", psQuery)
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
                    url += String.Format("/{0}", psID)
               End If

               If Not String.IsNullOrEmpty(psQuery) Then
                    url += String.Format("?{0}", psQuery)
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

               Dim url As String = ""
               If String.IsNullOrEmpty(psEntityID) = False Then
                    url = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID)
               Else
                    url = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)
               End If


               If Not String.IsNullOrEmpty(psQuery) Then
                    If psQuery.Contains("??") Then
                         url += String.Format("{0}", Replace(psQuery.Trim(), "??", String.Empty))
                    Else
                         url += String.Format("?{0}", psQuery.Trim())
                    End If
               End If

               Dim jsSerializer As New JavaScriptSerializer()
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
                         url += String.Format("{0}", Replace(psQuery.Trim(), "??", String.Empty))
                    Else
                         url += String.Format("?{0}", psQuery.Trim())
                    End If
               End If

               Dim jsSerializer As New JavaScriptSerializer()
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

     Public Function CallWebEndpointUploadFile(psEndPoint As String, psEntityID As String, psFileName As String, psObject As String, pnWidth As Integer, pnHeight As Integer) As IRestResponse
          Dim bFileContent As Byte() = Nothing
          Dim validImageTypes As String() = {".jpg", ".jpeg", ".png", ".bmp"}

          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)

               ' Validate if the file exists
               If Not File.Exists(psFileName) Then
                    Throw New FileNotFoundException("The specified file does not exist: " & psFileName)
               End If

               ' Validate if the file is a valid image type
               Dim fileExtension As String = Path.GetExtension(psFileName).ToLower()

               ' Process the file based on its extension
               Select Case fileExtension
                    Case ".jpg", ".jpeg"
                         bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Jpeg, pnWidth, pnHeight)
                    Case ".png"
                         bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Png, pnWidth, pnHeight)
                    Case ".bmp"
                         bFileContent = LoadAndResizeImageAsBytes(psFileName, Imaging.ImageFormat.Bmp, pnWidth, pnHeight)
                    Case Else
                         bFileContent = File.ReadAllBytes(psFileName)
               End Select

               If bFileContent IsNot Nothing Then
                    Dim url As String = String.Format("{0}{1}/{2}", goConnection._ServerAddress, psEndPoint, psEntityID, psObject)
                    Dim oClient = New RestClient(url)
                    Dim oRequest = New RestRequest(psObject, Method.POST)
                    oRequest.AddHeader("Accept-Encoding", "gzip, deflate, br")
                    oRequest.AddFile("file", bFileContent, Path.GetFileName(psFileName), String.Format("{0}/{1}", psObject, fileExtension.TrimStart("."c)))
                    Dim oResponse = ExecuteAPI(oClient, oRequest, False)
                    RaiseEvent APICallEvent(oRequest, oResponse)
                    Return oResponse
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

                    If Not String.IsNullOrEmpty(psRootNode) AndAlso psRootNode.Contains("[0]") Then
                         sRoot = psRootNode.Substring(0, psRootNode.LastIndexOf("[0]"))
                         sNode = psNodeName
                    End If

                    If oResponse.StatusCode = HttpStatusCode.OK AndAlso json.ContainsKey("responsePayload") Then
                         Dim oObject As JToken = json.SelectToken(sRoot)
                         If oObject IsNot Nothing AndAlso oObject.Count > 0 Then
                              Try
                                   If sNode = "description" Then
                                        ' Try to extract it from the object
                                        Dim oToken As JToken = json.SelectToken(String.Format("{0}[0].translations.en.description", sRoot))
                                        If oToken IsNot Nothing Then
                                             Return oToken.ToString()
                                        End If
                                   End If
                              Catch ex As Exception
                                   ' Handle exception
                              End Try

                              Try
                                   If oObject.HasValues AndAlso oObject.First IsNot Nothing Then
                                        sReturnValue = oObject.First.SelectToken(sNode).Value(Of String)()
                                   End If
                              Catch ex As Exception
                                   ' Handle exception
                              End Try
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
                    Dim oResponse As IRestResponse = CallWebEndpointUsingGet("metadata/v1/lookup-types", String.Empty, String.Empty, "code")
                    If oResponse IsNot Nothing Then
                         Dim json As JObject = JObject.Parse(oResponse.Content)
                         Dim oObject As JToken = json.SelectToken("responsePayload.content")
                         If oObject IsNot Nothing Then
                              For Each oNode In oObject
                                   If oNode IsNot Nothing Then
                                        goLookupTypes.Add(New clsLookUpTypes With {.Id = oNode("id"), .Code = oNode("code"), .Description = oNode("description"), .LookupGroup = oNode("lookupGroup")})
                                   End If
                              Next
                         End If

                         If goLookupTypes IsNot Nothing Then
                              goLookupTypes = goLookupTypes.OrderBy(Function(n) n.Description).ToList()
                         End If
                    End If
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try
     End Sub

     Private Sub LoadTranslatableLanguages()
          Try
               ' Remove the clear operation to retain existing languages
               ' goLanguages.Clear()

               If goLanguages IsNot Nothing Then
                    Dim oResponse As IRestResponse = CallWebEndpointUsingGet("metadata/v1/languages", String.Empty, "translationEnabled==true", "code")
                    If oResponse IsNot Nothing Then
                         Dim json As JObject = JObject.Parse(oResponse.Content)
                         Dim oObject As JToken = json.SelectToken("responsePayload.content")
                         If oObject IsNot Nothing Then
                              For Each oNode In oObject
                                   If oNode IsNot Nothing Then
                                        Dim code As String = oNode("code").ToString()
                                        ' Check if the code already exists in goLanguages
                                        Dim exists As Boolean = goLanguages.Any(Function(lang) lang.code = code)
                                        If Not exists Then
                                             goLanguages.Add(New clsLanguages With {
                                                   .id = oNode("id"),
                                                   .code = code,
                                                   .translations = GetTranslation(oNode.SelectToken("translations.en.description"))
                                              })
                                        End If
                                   End If
                              Next
                         End If
                    End If
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try
     End Sub


     Public Function CallGraphQL(psEndPoint As String, psDTOJson As String) As IRestResponse
          Try
               psEndPoint = ExtractSystemVariables(psEndPoint)
               psDTOJson = ExtractSystemVariables(psDTOJson)

               psDTOJson = psDTOJson.Replace(vbNewLine, "").Replace(ControlChars.Quote, "\" & ControlChars.Quote)

               Dim url As String = String.Format("{0}{1}", goConnection._ServerAddress, psEndPoint)

               Dim jsSerializer As New JavaScriptSerializer()
               Dim serialized = jsSerializer.Serialize(psDTOJson)
               Dim oClient = New RestClient(url)
               Dim sBody As String = String.Format("{{""query"":""{0}""}}", psDTOJson).Replace(vbNewLine, String.Empty)

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
                    Dim oResponse As IRestResponse = CallWebEndpointUsingGet("metadata/v1/hierarchies", String.Empty, "")
                    If oResponse IsNot Nothing Then
                         Dim json As JObject = JObject.Parse(oResponse.Content)
                         Dim oObject As JToken = json.SelectToken("responsePayload")

                         For Each oNode In oObject
                              If oNode IsNot Nothing AndAlso oNode.SelectToken("deleted").ToString().ToLower() = "false" Then
                                   Dim sDescription As JToken = oNode.SelectToken("translations.en.description")
                                   goHierarchies.Add(New clsHierarchies With {
                                       .Id = oNode("id"),
                                       .Code = oNode("code"),
                                       .Description = sDescription,
                                       .ParentHierarchyId = oNode("parentHierarchyId"),
                                       .Enabled = oNode("enabled").ToString().ToUpper() = "TRUE",
                                       .Level = oNode("level"),
                                       .Priority = oNode("priority")
                                   })
                              End If
                         Next

                         ' Check for children
                         For Each oH As clsHierarchies In goHierarchies
                              oH.ParentName = LookForParentHierarchies(oH, String.Empty)
                         Next

                         If goHierarchies IsNot Nothing Then
                              goHierarchies = goHierarchies.OrderBy(Function(n) n.ParentName).ThenBy(Function(n) n.Description).ToList()
                         End If
                    End If
               End If
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try
     End Sub

     Private Function LookForParentHierarchies(poHierarchy As clsHierarchies, psParentString As String) As String
          Try
               If Not String.IsNullOrEmpty(poHierarchy.ParentHierarchyId) Then
                    Dim oParent As clsHierarchies = goHierarchies.FirstOrDefault(Function(n) n.Id = poHierarchy.ParentHierarchyId)
                    If oParent IsNot Nothing AndAlso oParent.Id <> "1" Then
                         If Not String.IsNullOrEmpty(oParent.Description) Then
                              psParentString = If(String.IsNullOrEmpty(psParentString), oParent.Description, oParent.Description & "\" & psParentString)
                         End If

                         If Not String.IsNullOrEmpty(oParent.ParentHierarchyId) Then
                              psParentString = LookForParentHierarchies(oParent, psParentString)
                         End If
                    Else
                         psParentString = If(String.IsNullOrEmpty(psParentString), "Enterprise", "Enterprise" & " \ " & psParentString)
                    End If
               End If

               Return psParentString
          Catch ex As Exception
               RaiseEvent ErrorEvent(ex)
          End Try

          Return String.Empty
     End Function

     Public Function ExtractSystemVariables(psObject As String) As String
          Try
               ' Check for Hierarchy
               If Not String.IsNullOrEmpty(psObject) Then
                    If psObject.ToUpper().Contains("@@HIERARCHYID@@") Then psObject = psObject.Replace("@@HIERARCHYID@@", gsSelectedHierarchy.ToString())
                    If psObject.ToUpper().Contains("@@EMPTYARRAY@@") Then psObject = psObject.Replace(ControlChars.Quote & "@@EMPTYARRAY@@" & ControlChars.Quote, "[]")
                    If psObject.ToUpper().Contains("@@SHIPID@@") Then psObject = psObject.Replace("@@SHIPID@@", gsSelectedHierarchy.ToString())
                    If psObject.ToUpper().Contains("@@RVCID@@") Then psObject = psObject.Replace("@@RVCID@@", gsSelectedHierarchy.ToString())
                    If psObject.ToUpper().Contains("@@PARENTID@@") Then psObject = psObject.Replace("@@PARENTID@@", gsSelectedHierarchyParent.ToString())
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
               oLog.Add(New JProperty("objectId", psObjectID.Replace("-", "")))

               Dim oLogs As New JArray
               oLogs.Add(oLog)
               oJsn.Add(New JProperty("logEntities", oLogs))

               Dim oDTO As New JArray
               oDTO.Add(oJsn)

               Call CallWebEndpointUsingPost("metadata/v1/logs", oDTO.ToString())
          Catch ex As Exception
               ' Handle exception
          End Try
     End Function

End Class
