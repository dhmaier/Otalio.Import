Imports System.IO
Imports System.Configuration
Imports System.Linq
Imports System.Drawing.Drawing2D
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.LookAndFeel
Imports DevExpress.Skins
Imports DevExpress.Utils
Imports DevExpress.Utils.Drawing
Imports DevExpress.Utils.Svg
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Globalization
Imports System.Text.RegularExpressions

Module bFunctions

     Public Function IsPropertyArray(ByVal jsonString As String, ByVal propertyName As String) As Boolean
          Dim jsonObject As JObject = JObject.Parse(jsonString)
          Dim propertyNode As JToken = jsonObject.SelectToken(propertyName)

          If propertyNode IsNot Nothing AndAlso propertyNode.Type = JTokenType.Array Then
               Return True
          End If

          Return False
     End Function
     Public Function IsPropertyArray(ByVal propertyNode As JProperty) As Boolean
          If propertyNode IsNot Nothing AndAlso propertyNode.Value.Type = JTokenType.Array Then
               Return True
          End If

          Return False
     End Function

     Public Function GetAppPath(psFile As String, Optional psSystemLocation As String = "") As String

          If psSystemLocation = "" Then
               psSystemLocation = Application.LocalUserAppDataPath
          End If

          Dim sPath As String = Path.GetDirectoryName(psSystemLocation)
          Dim sFullPath As String = sPath & "\" & psFile

          Return sFullPath

     End Function



     Public Function LoadFile(psFullPath As String) As Object


    Try

      Using oFileStream = New FileStream(psFullPath, FileMode.Open, FileAccess.Read, FileShare.Read)
        Dim bytes() As Byte = New Byte((oFileStream.Length) - 1) {}
        Dim numBytesToRead As Integer = CType(oFileStream.Length, Integer)
        Dim numBytesRead As Integer = 0

        While (numBytesToRead > 0)
          ' Read may return anything from 0 to numBytesToRead. 
          Dim n As Integer = oFileStream.Read(bytes, numBytesRead,
              numBytesToRead)
          ' Break when the end of the file is reached. 
          If (n = 0) Then
            Exit While
          End If
          numBytesRead = (numBytesRead + n)
          numBytesToRead = (numBytesToRead - n)

        End While

        numBytesToRead = bytes.Length

        Return ByteArrayToObject(bytes)


        oFileStream.Close()

        'cleanup
        bytes = Nothing
      End Using

      Return True

    Catch ex As Exception
      Return False

    End Try
  End Function

  Public Function LoadFileJson(psFullPath As String) As Object


    Try

      Dim oObject = JObject.Parse(File.ReadAllText(psFullPath))

      If oObject IsNot Nothing Then
        Return oObject
      End If

      Return True

    Catch ex As Exception
      Return False

    End Try
  End Function

     Public Function SaveFile(psFullPath As String, poObject As Object) As Boolean
          'Save it

          Try

               Dim oPackageBytes() As Byte = ObjectToByteArray(poObject)
               Using oFileStream = New FileStream(psFullPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Write)

                    oFileStream.Write(oPackageBytes, 0, oPackageBytes.Length)
                    oFileStream.Close()

                    oPackageBytes = Nothing
               End Using

               Return True
          Catch ex As Exception
               Return False
          Finally

          End Try
     End Function

     Public Function SaveFileJson(psFullPath As String, poObject As Object) As Boolean
          Try
               ' Validate the file path
               If String.IsNullOrEmpty(psFullPath) Then
                    Throw New ArgumentException("File path cannot be null or empty.", NameOf(psFullPath))
               End If

               ' Serialize the object to JSON and write to the file
               Dim json As String = JsonConvert.SerializeObject(poObject, Formatting.Indented)
               File.WriteAllText(psFullPath, json)

               Return True
          Catch ex As ArgumentException
               ' Log specific argument exception details (example: to a log file or system log)
               ShowErrorForm("Invalid argument: " & ex.Message)
               Return False
          Catch ex As IOException
               ' Log I/O exception details
               ShowErrorForm("I/O error while writing file: " & ex.Message)
               Return False
          Catch ex As UnauthorizedAccessException
               ' Log unauthorized access exception details
               ShowErrorForm("Unauthorized access: " & ex.Message)
               Return False
          Catch ex As JsonException
               ' Log JSON serialization exception details
               ShowErrorForm("JSON serialization error: " & ex.Message)
               Return False
          Catch ex As Exception
               ' Log general exception details
               ShowErrorForm("An unexpected error occurred: " & ex.Message)
               Return False
          End Try
     End Function


     'The Following functions should be moved to MiscUtils dll
     Public Function ObjectToByteArray(ByVal Obj As Object) As Byte()
    Dim Value() As Byte = Nothing
    Dim MS As MemoryStream = Nothing
    Dim BF As Runtime.Serialization.Formatters.Binary.BinaryFormatter = Nothing
    Try
      MS = New MemoryStream
      BF = New Runtime.Serialization.Formatters.Binary.BinaryFormatter
      BF.Serialize(MS, Obj)
      Value = MS.ToArray()
      Return Value
    Finally
      If Not MS Is Nothing Then
        MS.Close()
        MS = Nothing
      End If
      BF = Nothing
      Value = Nothing

    End Try

  End Function

  Public Function ByteArrayToObject(ByVal ByteArray() As Byte) As Object
    Dim Value As Object = Nothing
    Dim MS As MemoryStream = Nothing
    Dim BF As Runtime.Serialization.Formatters.Binary.BinaryFormatter = Nothing
    Try
      MS = New MemoryStream(ByteArray)
      BF = New Runtime.Serialization.Formatters.Binary.BinaryFormatter
      Value = BF.Deserialize(MS)
      Return Value
    Finally
      If Not MS Is Nothing Then
        MS.Close()
        MS = Nothing
      End If
      BF = Nothing
      Value = Nothing
    End Try

  End Function

  Public Function Base64Encode(ByVal plainText As String) As String
    Dim plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText)
    Return System.Convert.ToBase64String(plainTextBytes)
  End Function

  Public Function __Assign(Of T)(ByRef target As T, value As T) As T
    target = value
    Return value
  End Function

     Public Function GenerateGUID() As String


          Return Replace(System.Guid.NewGuid.ToString(), "-", "")


     End Function



     Public Function ExtractColumnProperties(psNode As String) As clsColumnProperties
          Try



               If psNode IsNot Nothing Then

                    Dim oCP As New clsColumnProperties
                    oCP.SourceText = psNode

                    If psNode.Contains(":") Then
                         Dim oValues As List(Of String) = psNode.Split(":").ToList

                         If oValues.Count >= 1 Then oCP.CellName = oValues(0).ToString
                         If oValues.Count >= 2 Then oCP.Format = oValues(1).ToString
                         If oValues.Count >= 3 Then oCP.IndexID = oValues(2).ToString

                    Else
                         oCP.CellName = psNode
                    End If

                    Return oCP
               Else
                    Return Nothing
               End If

          Catch ex As Exception

          End Try

     End Function


     Public Function CompareFilesOld(ByVal file1FullPath As String, ByVal file2FullPath As String) As Boolean

    If Not File.Exists(file1FullPath) Or Not File.Exists(file2FullPath) Then
      'One or both of the files does not exist.
      Return False
    End If

    If file1FullPath = file2FullPath Then
      ' fileFullPath1 and fileFullPath2 points to the same file...
      Return True
    End If

    Try
      Dim file1Hash As String = hashFile(file1FullPath)
      Dim file2Hash As String = hashFile(file2FullPath)

      If file1Hash = file2Hash Then
        Return True
      Else
        Return False
      End If

    Catch ex As Exception
      Return False
    End Try
  End Function

  Private Function hashFile(ByVal filepath As String) As String
    Using reader As New System.IO.FileStream(filepath, IO.FileMode.Open, IO.FileAccess.Read)
      Using md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim hash() As Byte = md5.ComputeHash(reader)
        Return System.Text.Encoding.Unicode.GetString(hash)
      End Using
    End Using
  End Function

     Public Function CompareFiles(ByVal FileFullPath1 As String,
    ByVal FileFullPath2 As String) As Boolean

          'returns true if two files passed to is are identical, false
          'otherwise

          'does byte comparison; works for both text and binary files

          'Throws exception on errors; you can change to just return 
          'false if you prefer


          Dim objMD5 As New System.Security.Cryptography.MD5CryptoServiceProvider
          Dim objEncoding As New System.Text.ASCIIEncoding()

          Dim aFile1() As Byte, aFile2() As Byte
          Dim strContents1, strContents2 As String
          Dim objReader As StreamReader
          Dim objFS As FileStream
          Dim bAns As Boolean
          If Not File.Exists(FileFullPath1) Then _
            Throw New Exception(FileFullPath1 & " doesn't exist")
          If Not File.Exists(FileFullPath2) Then _
         Throw New Exception(FileFullPath2 & " doesn't exist")

          Try

               objFS = New FileStream(FileFullPath1, FileMode.Open)
               objReader = New StreamReader(objFS)
               aFile1 = objEncoding.GetBytes(objReader.ReadToEnd)
               strContents1 =
              objEncoding.GetString(objMD5.ComputeHash(aFile1))
               objReader.Close()
               objFS.Close()


               objFS = New FileStream(FileFullPath2, FileMode.Open)
               objReader = New StreamReader(objFS)
               aFile2 = objEncoding.GetBytes(objReader.ReadToEnd)
               strContents2 =
             objEncoding.GetString(objMD5.ComputeHash(aFile2))

               bAns = strContents1 = strContents2
               objReader.Close()
               objFS.Close()
               aFile1 = Nothing
               aFile2 = Nothing

          Catch ex As Exception
               Throw ex

          End Try

          Return bAns
     End Function

     Public Function IsDateOrTime(psString As String) As Boolean
          If IsDate(psString) Then Return True
     End Function


#Region "Drawing Functions"
     Public Function ResizeImage(img As Image, width As Integer, height As Integer) As Image
    Dim newImage = New Bitmap(width, height)
    Using gr = Graphics.FromImage(newImage)
      gr.SmoothingMode = SmoothingMode.HighQuality
      gr.InterpolationMode = InterpolationMode.HighQualityBicubic
      gr.PixelOffsetMode = PixelOffsetMode.HighQuality
      gr.DrawImage(img, New Rectangle(0, 0, width, height))
    End Using
    Return newImage
  End Function

  Public Function ResizeImage(img As Image, size As Size) As Image
    Return ResizeImage(img, size.Width, size.Height)
  End Function

  Public Function ResizeImage(bmp As Bitmap, width As Integer, height As Integer) As Image
    Return ResizeImage(DirectCast(bmp, Image), width, height)
  End Function

  Public Function ResizeImage(bmp As Bitmap, size As Size) As Image
    Return ResizeImage(DirectCast(bmp, Image), size.Width, size.Height)
  End Function

     Public Function LoadAndResizeImageAsBytes(psFileName As String, pnFormat As System.Drawing.Imaging.ImageFormat, pnWidth As Integer, pnHeight As Integer) As Byte()
          Try

               If File.Exists(psFileName) Then
                    Dim oImage As Image = Image.FromFile(psFileName)
                    If oImage IsNot Nothing Then
                         oImage = ResizeImage(oImage, pnWidth, pnHeight)

                         Using ms = New MemoryStream()
                              oImage.Save(ms, pnFormat) ' Use appropriate format here
                              Return ms.ToArray()
                         End Using

                    End If
                    oImage.Dispose()

               End If
          Catch ex As Exception

          End Try
          Return Nothing

     End Function

     Public Function CreateImage(ByVal data() As Byte, Optional ByVal skinProvider As ISkinProvider = Nothing) As Image
          Dim svgBitmap As New SvgBitmap(data)
          Return svgBitmap.Render(SvgPaletteHelper.GetSvgPalette(If(skinProvider, UserLookAndFeel.Default), ObjectState.Normal), ScaleUtils.GetScaleFactor().Height)
     End Function


#End Region

     Public Function RemoveLastComma(psString As String) As String
          If psString.EndsWith(",") Then
               Return psString.Substring(0, psString.Length - 1)
          Else
               Return psString
          End If
     End Function

     Function GetLastValue(textValue As String, separator As String) As String
          If textValue.Contains(separator) Then
               Dim values() As String = textValue.Split(separator)
               Return values(values.Length - 1)
          Else
               Return textValue
          End If
     End Function


     Function RemoveConditionFromRsqlQuery(query As String, valueToRemove As String) As String

          Dim separators As String() = {" OR ", " AND ", ";", " or ", " and "}
          Dim pattern As String = String.Join("|", separators.Select(Function(x) "(" + Regex.Escape(x) + ")"))

          ' The Regex.Split method splits the string at the separators but includes the separators in the result
          Dim parts = Regex.Split(query, pattern).ToList()

          Dim conditions As New List(Of String)
          For i = 0 To parts.Count - 2 Step 2
               If Not parts(i).Contains(valueToRemove) Then
                    conditions.Add(parts(i) & parts(i + 1))
               End If
          Next
          ' Last part does not have a trailing separator, add it separately if it does not contain the value to remove
          If parts.Count Mod 2 = 1 AndAlso Not parts.Last().Contains(valueToRemove) Then
               conditions.Add(parts.Last())
          End If

          ' Reassemble the RSQL query
          Dim newQuery As String = String.Join("", conditions)

          ' Remove trailing logical operation
          For Each oOperator In separators
               If newQuery.EndsWith(oOperator) Then
                    newQuery = newQuery.Substring(0, newQuery.Length - oOperator.Length)
                    Exit For
               End If
          Next

          Return newQuery.Trim()


     End Function



     Public Function FormatCase(psFormat As String, psString As String) As String
          If String.IsNullOrEmpty(psString) = False Then
               Select Case psFormat
                    Case "U" : Return psString.ToUpper
                    Case "L" : Return psString.ToLower
                    Case "P" : Return ToProperCase(psString)
                    Case "T" : Return ToTitleCase(psString)
                    Case "FD"
                         If IsDate(psString) Then
                              'If CDate(psString).TimeOfDay.TotalSeconds = 0 And CDate(psString).Date <> Date.MinValue Then
                              Return CDate(psString).ToString("yyyy-MM-dd")
                              'Else

                              'Return CDate(psString).ToString("yyyy-MM-dd HH:mm")
                              'End If
                         Else
                              Return psString
                         End If
                    Case "FT"
                         If IsDateOrTime(psString) Then
                              Return CDate(psString).ToString("HH:mm:ss")
                         End If
                    Case "FDT"
                         If IsDateOrTime(psString) Then
                              Return CDate(psString).ToString("yyyy-MM-dd HH:mm")
                         End If

                    Case Else
                         Return psString
               End Select
          Else
               Return ""
          End If

     End Function

     ' Function to convert a sentence to Proper Case
     Public Function ToProperCase(sentence As String) As String
          Return StrConv(sentence, VbStrConv.ProperCase)
     End Function

     ' Function to convert a sentence to Title Case
     Public Function ToTitleCase(sentence As String) As String
          Dim textInfo As TextInfo = CultureInfo.CurrentCulture.TextInfo
          Dim words As String() = sentence.ToLower().Split(" "c)
          Dim smallWords As String() = {
        "a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet",
        "at", "by", "for", "in", "of", "on", "to", "up", "with",
        "about", "across", "after", "against", "along", "among", "around", "as", "before", "behind", "below",
        "beneath", "beside", "between", "beyond", "during", "except", "from", "inside", "into", "near",
        "off", "onto", "out", "outside", "over", "past", "through", "under", "until", "within", "without"
    }

          For i As Integer = 0 To words.Length - 1
               If i = 0 OrElse i = words.Length - 1 OrElse Not smallWords.Contains(words(i)) Then
                    words(i) = textInfo.ToTitleCase(words(i))
               Else
                    words(i) = words(i)
               End If
          Next

          Return String.Join(" ", words)
     End Function




     Public Function ExtractQueryParameters(psQuery As String) As List(Of String)

          psQuery = Replace(psQuery, ";", " and ")
          psQuery = Replace(psQuery, ",", " or ")

          Dim oDataSource As List(Of String) = psQuery.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries).ToList
          Dim oList As New List(Of String)
          Dim sCurrentParameter As String = ""
          For Each sString As String In oDataSource
               If sString.ToUpper = "AND" Or sString.ToUpper = "OR" Or sString.ToUpper = "," Or sString.ToUpper = ";" Then
                    oList.Add(sCurrentParameter)
                    sCurrentParameter = ""
               Else
                    sCurrentParameter += sString & " "

               End If

          Next

          If String.IsNullOrEmpty(sCurrentParameter) = False Then
               oList.Add(sCurrentParameter)
          End If

          Return oList

     End Function

     Public Function GetTranslation(psDescription As String, Optional psLength As Integer = 0) As Translations

          If psLength > 0 Then
               If psDescription.Length > psLength Then
                    psDescription = psDescription.Substring(0, psLength - 1)
               End If
          End If

          Dim oTransaltion As New Translations With {.en = New En With {.description = psDescription}}
          Return oTransaltion

     End Function

     Public Function GetTranslationsArray(pbEnableTransaltion As Boolean, psNodes As List(Of Array)) As String

          Dim oTranslation As String = ""
          oTranslation += "  " & String.Format("{0}{1}{0}: {2}{4}{3},", ControlChars.Quote, "en", "{", "}", GetTranslationsObjects(psNodes, 0))

          If pbEnableTransaltion Then
               Dim nColumnNumber As Integer = 0
               For Each oLanguage As clsLanguages In goLanguages.Where(Function(n) n.code.ToLower <> "en").ToList
                    nColumnNumber += 1
                    oTranslation += vbNewLine & "  " & String.Format("{0}{1}{0}: {2}{4}{3},", ControlChars.Quote, oLanguage.code.ToLower, "{", "}", GetTranslationsObjects(psNodes, nColumnNumber))
               Next
          End If

          If oTranslation.EndsWith(",") Then
               oTranslation = oTranslation.Substring(0, oTranslation.Length - 1)
          End If


          Return "{" & oTranslation & "}"

     End Function

     Public Function GetTranslationsObjects(psNodes As List(Of Array), pnColumnNumber As Integer) As String
          Dim sString As String = ""
          For Each oArray As Array In psNodes
               Dim sCell As String = ConvertToLetter(ConvertToNumber(oArray(0)) + pnColumnNumber)
               sString += String.Format("{0}{1}{0}: {0}<!{2}:{3}!>{0},", ControlChars.Quote, oArray(2), sCell, oArray(1))
          Next


          If sString.EndsWith(",") Then
               sString = sString.Substring(0, sString.Length - 1)
          End If

          Return sString
     End Function

     Public Function FindAndReplaceTranslations(psObject As String) As String

          Dim nPOS As Integer = 0
          Do While psObject.Substring(nPOS).Contains("@@T!")
               Dim sColumn As String = ""
               Dim sFormat As String = ""
               Dim sNode As String = "description"
               Dim nStart As Integer = InStr(nPOS + 1, psObject, "@@T!")
               Dim nEnd As Integer = InStr(nPOS + 1, psObject, "!T@@")
               Dim sColumnText As String = psObject.Substring(nStart + 3, nEnd - nStart - 4)
               Dim oListOfProperties As New List(Of Array)
               For Each oTranslatedObject As String In sColumnText.Split("|")
                    If oTranslatedObject.Contains(":") Then
                         Dim sProperties As String() = oTranslatedObject.Split(":").ToArray
                         Select Case sProperties.Count
                              Case 1
                                   sColumn = sProperties(0)
                              Case 2
                                   sColumn = sProperties(0)
                                   sFormat = sProperties(1)
                              Case 3
                                   sColumn = sProperties(0)
                                   sFormat = sProperties(1)
                                   sNode = sProperties(2)

                         End Select

                         oListOfProperties.Add({sColumn, sFormat, sNode})
                    End If
               Next

               psObject = Replace(psObject, String.Format("{1}@@T!{0}!T@@{1}", sColumnText, ControlChars.Quote), GetTranslationsArray(gbTranslationsEnabled, oListOfProperties))


               nPOS = nEnd
          Loop

          Return psObject

     End Function
     Public Function ConvertToLetter(iCol As Long) As String
          Dim a As Long
          Dim b As Long
          a = iCol
          ConvertToLetter = ""
          Do While iCol > 0
               a = Int((iCol - 1) / 26)
               b = (iCol - 1) Mod 26
               ConvertToLetter = Chr(b + 65) & ConvertToLetter
               iCol = a
          Loop
     End Function

     Public Function ConvertToNumber(psColumnLetters As String) As Integer
          Dim nNumber As Integer = 0
          Dim sCharList As List(Of String) = psColumnLetters.ToUpper.Split().ToList
          If sCharList IsNot Nothing Then
               For Each sChar As String In sCharList
                    nNumber += (Asc(sChar) - 64)
               Next
          End If
          Return nNumber
     End Function


     Public Sub ValidateTranslationsValidators(ByRef poTemplate As clsDataImportTemplateV2, pbEnableTransaltion As Boolean)

          If poTemplate IsNot Nothing Then
               'check if there is the translations function included in the DTO stacte
               If poTemplate.DTOObject.Contains("@@T!") Then

                    Dim sDTOObject As String = poTemplate.DTOObject.ToString
                    Dim nPOS As Integer = 0
                    Dim sListOfColumns As New List(Of String)
                    Dim sListOfFormats As New List(Of String)
                    Do While sDTOObject.Substring(nPOS).Contains("@@T!")

                         Dim sColumn As String = ""
                         Dim sFormat As String = ""

                         Dim nStart As Integer = InStr(nPOS + 1, sDTOObject, "@@T!")
                         Dim nEnd As Integer = InStr(nPOS + 1, sDTOObject, "!T@@")
                         Dim oTranslatedObject As String = sDTOObject.Substring(nStart + 3, nEnd - nStart - 4)
                         For Each sColumnText As String In oTranslatedObject.Split("|")
                              If sColumnText.Contains(":") Then
                                   Dim sProperties As String() = sColumnText.Split(":").ToArray
                                   If sProperties.Count() > 0 Then
                                        sListOfColumns.Add(sProperties(0))
                                   End If

                                   If sProperties.Count >= 1 Then
                                        sListOfFormats.Add(sProperties(1))
                                   Else
                                        sListOfFormats.Add("")
                                   End If

                              End If
                         Next

                         nPOS = nEnd
                    Loop

                    poTemplate.Validators.RemoveAll(Function(n) n.ValidationType = 2)


                    If sListOfColumns.Count > 0 Then
                         Dim nCountOfTranslations As Integer = 0
                         For Each oColumnAddress As String In sListOfColumns
                              Dim nColumnId As Integer = ConvertToNumber(oColumnAddress)
                              Dim nCounter As Integer = 0

                              poTemplate.Validators.RemoveAll(Function(n) n.ValidationType = 2 And n.Query.Contains(String.Format("<!{0}!>", oColumnAddress)))

                              For Each oLanguage As clsLanguages In goLanguages.Where(Function(n) n.code.ToLower <> "en").OrderBy(Function(n) n.code).ToList
                                   nCounter += 1
                                   poTemplate.Validators.Add(CreateTranslationValidation(oLanguage, oColumnAddress, ConvertToLetter(nColumnId + nCounter), ((nCountOfTranslations + 1) * 99) + nCounter, pbEnableTransaltion, sListOfFormats(nCountOfTranslations)))
                              Next

                              nCountOfTranslations += 1
                         Next

                    End If

               End If
          End If

     End Sub

     Public Function CreateTranslationValidation(psLanguage As clsLanguages, psCellAddressSource As String, psCellAddressDestination As String, pnPriority As Integer, pbEnableTransaltion As Boolean, psFormat As String) As clsValidation

          Dim oValidation As New clsValidation
          With oValidation
               .ID = GenerateGUID()
               .Priority = pnPriority.ToString
               .ValidationType = "2"
               .APIEndpoint = "metadata/v1/languages/translate"
               .Comments = psLanguage.translations.en.description
               .Query = String.Format("{0}{2}source{2}:{2}EN{2},{2}targetCodes{2}:[{2}{3}{2}],{2}text{2}:{2}<!{4}!>{2}{1}", "{", "}", ControlChars.Quote, psLanguage.code.ToUpper, psCellAddressSource)
               .ReturnCell = psCellAddressDestination
               .ReturnNodeValue = "responsePayload.translationMap.success.de.translatedText"
               .Enabled = IIf(pbEnableTransaltion, "1", "0")
               .Visibility = IIf(pbEnableTransaltion, "1", "0")
               .Formatted = psFormat
          End With

          Return oValidation

     End Function


     Public Function RemoveEmptyPropertiesRecursively(ByVal jToken As JToken) As JToken
          If jToken.Type = JTokenType.Object Then
               Dim copy As JObject = New JObject()
               For Each prop As JProperty In jToken.Children(Of JProperty)()
                    Dim childToken As JToken = RemoveEmptyPropertiesRecursively(prop.Value)
                    If Not (childToken.Type = JTokenType.String AndAlso String.IsNullOrEmpty(childToken.ToString())) Then
                         copy.Add(prop.Name, childToken)
                    End If
               Next
               Return copy
          ElseIf jToken.Type = JTokenType.Array Then
               Dim newArray As JArray = New JArray()
               For Each child As JToken In jToken.Children()
                    Dim childToken As JToken = RemoveEmptyPropertiesRecursively(child)
                    newArray.Add(childToken)
               Next
               Return newArray
          Else
               Return jToken
          End If
     End Function

     Public Function CleanJson(ByVal json As String) As String
          Dim parsedJson As JToken = JToken.Parse(json)
          Dim cleanedJson As JToken = RemoveEmptyPropertiesRecursively(parsedJson)
          Return cleanedJson.ToString()
     End Function



     Public Sub ShowErrorForm(errorMessage As String)
          Dim errorForm As New frmError(errorMessage)
          errorForm.ShowDialog()
     End Sub

     Public Sub ShowErrorForm(errorMessage As Exception)
          Dim errorForm As New frmError(errorMessage)
          errorForm.ShowDialog()
     End Sub

     Public Sub ShowErrorForm(errorMessagetext As String, errorMessage As Exception)
          Dim errorForm As New frmError(errorMessagetext, errorMessage)
          errorForm.ShowDialog()
     End Sub

End Module
