
Imports DevExpress.XtraGrid
Imports System.Windows.Forms
Imports System.Xml
Imports System.IO

Module bGridHelper



     Public Sub SaveGridLayouts(ByVal parent As Control, ByVal xmlFilePath As String)
          Try
               Dim doc As New XmlDocument()
               Dim root As XmlElement = doc.CreateElement("GridLayouts")
               doc.AppendChild(root)

               SaveGridLayoutsRecursive(parent, doc, root)

               doc.Save(xmlFilePath)
          Catch ex As Exception
               ShowErrorForm("Error saving grid layouts", ex)
          End Try
     End Sub

     Private Sub SaveGridLayoutsRecursive(ByVal parent As Control, ByVal doc As XmlDocument, ByVal root As XmlElement)
          For Each control As Control In parent.Controls
               Try
                    If TypeOf control Is GridControl Then
                         Dim gridControl As GridControl = DirectCast(control, GridControl)
                         Dim view = gridControl.MainView
                         If view IsNot Nothing Then
                              Dim gridElement As XmlElement = doc.CreateElement("Grid")
                              gridElement.SetAttribute("Name", gridControl.Name)

                              Using ms As New MemoryStream()
                                   view.SaveLayoutToStream(ms)
                                   ms.Seek(0, SeekOrigin.Begin)
                                   Dim reader As New StreamReader(ms)
                                   gridElement.InnerText = reader.ReadToEnd()
                              End Using

                              root.AppendChild(gridElement)
                         End If
                    End If

                    If control.HasChildren Then
                         SaveGridLayoutsRecursive(control, doc, root)
                    End If
               Catch ex As Exception
                    ShowErrorForm($"Error saving layout for control: {control.Name}", ex)
               End Try
          Next
     End Sub


     Public Sub LoadGridLayouts(ByVal parent As Control, ByVal xmlFilePath As String)
          Try
               If Not File.Exists(xmlFilePath) Then Return

               Dim doc As New XmlDocument()
               doc.Load(xmlFilePath)

               Dim root As XmlElement = doc.DocumentElement
               LoadGridLayoutsRecursive(parent, root)
          Catch ex As Exception
               ShowErrorForm("Error loading grid layouts", ex)
          End Try
     End Sub

     Private Sub LoadGridLayoutsRecursive(ByVal parent As Control, ByVal root As XmlElement)
          For Each control As Control In parent.Controls
               Try
                    If TypeOf control Is GridControl Then
                         Dim gridControl As GridControl = DirectCast(control, GridControl)
                         Dim gridElement As XmlElement = TryCast(root.SelectSingleNode($"Grid[@Name='{gridControl.Name}']"), XmlElement)
                         If gridElement IsNot Nothing Then
                              Using ms As New MemoryStream()
                                   Dim writer As New StreamWriter(ms)
                                   writer.Write(gridElement.InnerText)
                                   writer.Flush()
                                   ms.Seek(0, SeekOrigin.Begin)
                                   gridControl.MainView.RestoreLayoutFromStream(ms)
                              End Using
                         End If
                    End If

                    If control.HasChildren Then
                         LoadGridLayoutsRecursive(control, root)
                    End If
               Catch ex As Exception
                    ShowErrorForm($"Error loading layout for control: {control.Name}", ex)
               End Try
          Next
     End Sub



End Module
