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

Module bFunctions



  Public Function GetAppPath(psFile As String) As String


    Dim sPath As String = Path.GetDirectoryName(Application.CommonAppDataPath)
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
    'Save it

    Try

      File.WriteAllText(psFullPath, JsonConvert.SerializeObject(poObject, Newtonsoft.Json.Formatting.Indented))

      Return True
    Catch ex As Exception
      Return False
    Finally

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

  Private Function GenerateGUID() As String


    Return System.Guid.NewGuid.ToString()

  End Function

  Public Function FormatString(psString As String, psFormat As String) As String


    Select Case psFormat.ToString
      Case "U" : Return psString.ToUpper
      Case "L" : Return psString.ToLower
      Case "P" : Return StrConv(psString, VbStrConv.ProperCase)
      Case Else : Return psString
    End Select
  End Function

  Public Function ExtractColumnProperties(psNode As String) As clsColumnProperties

    If psNode IsNot Nothing Then

      Dim oCP As New clsColumnProperties

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

  Public Function LoadAndResizeImageAsBytes(psFileName As String, pnFormat As System.Drawing.Imaging.ImageFormat) As Byte()
    Try

      If File.Exists(psFileName) Then
        Dim oImage As Image = Image.FromFile(psFileName)
        If oImage IsNot Nothing Then
          oImage = ResizeImage(oImage, 256, 256)

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

End Module
