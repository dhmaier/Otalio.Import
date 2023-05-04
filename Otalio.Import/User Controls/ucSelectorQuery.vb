Imports System.Net
Imports System.IO
Imports System.Web
Imports RestSharp
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.Linq
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraLayout
Imports System.ComponentModel

Public Class ucSelectorQuery

     Private moSelector As List(Of clsSelectors)
     Public Property _Selectors As List(Of clsSelectors)
          Get
               Return moSelector
          End Get
          Set(value As List(Of clsSelectors))
               moSelector = value
               Load()
          End Set
     End Property
     Public Property _Query As String
     Private listOfContorls As New List(Of Control)
     Public Sub Load()

          If moSelector IsNot Nothing AndAlso moSelector.Count > 0 Then
               For Each oSelector In moSelector.OrderBy(Function(n) n.Priority).ToList

                    Select Case oSelector.ValidationType
                         Case "TEXT"
                              Dim oEditor As New TextEdit
                              Dim oLayoutControlItem As New LayoutControlItem
                              With oLayoutControlItem
                                   .Text = oSelector.Label
                                   .Control = oEditor
                                   .Location = New System.Drawing.Point(0, 0)
                                   .Size = New System.Drawing.Size(606, 24)
                                   .TextSize = New System.Drawing.Size(96, 13)
                              End With
                              Me.Root.AddItem(oLayoutControlItem)
                              oEditor.Tag = oSelector
                              listOfContorls.Add(oEditor)
                         Case "DATE"
                              Dim oEditor As New DateEdit
                              oEditor.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom
                              oEditor.Properties.EditFormat.FormatString = "yyyy-MM-dd"

                              ' Set the display format to "yyyy-MM-dd"
                              oEditor.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
                              oEditor.Properties.DisplayFormat.FormatString = "yyyy-MM-dd"

                              Dim oLayoutControlItem As New LayoutControlItem
                              With oLayoutControlItem
                                   .Text = oSelector.Label
                                   .Control = oEditor
                                   .Location = New System.Drawing.Point(0, 0)
                                   .Size = New System.Drawing.Size(606, 24)
                                   .TextSize = New System.Drawing.Size(96, 13)
                              End With
                              Me.Root.AddItem(oLayoutControlItem)
                              oEditor.Tag = oSelector
                              listOfContorls.Add(oEditor)
                         Case "LIST", "LOOKUP"

                              Dim oEditor As New ImageComboBoxEdit
                              oEditor.Properties.Sorted = True
                              Dim oLayoutControlItem As New LayoutControlItem
                              With oLayoutControlItem
                                   .Text = oSelector.Label
                                   .Control = oEditor
                                   .Location = New System.Drawing.Point(0, 0)
                                   .Size = New System.Drawing.Size(606, 24)
                                   .TextSize = New System.Drawing.Size(96, 13)
                              End With
                              Me.Root.AddItem(oLayoutControlItem)
                              LoadComboBox(oEditor, oSelector, oSelector.Query)
                              AddHandler oEditor.EditValueChanged, AddressOf LoadLinkedComboBox
                              oEditor.Tag = oSelector
                              listOfContorls.Add(oEditor)


                    End Select
               Next
          End If

     End Sub

     Private Sub LoadComboBox(poEditor As ImageComboBoxEdit, poSelector As clsSelectors, psQuery As String)
          Try

               'DONT populate if linked to a selector
               If psQuery.Contains("<<!") = False Then

                    Dim oResponse As IRestResponse
                    Dim sNodeLoad As String = "responsePayload.content"
                    oResponse = goHTTPServer.CallWebEndpointUsingGet(poSelector.APIEndpoint, String.Empty, psQuery)

                    Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
                    jsServerResponse = JObject.Parse(oResponse.Content)
                    Dim oObject As JToken = jsServerResponse.SelectToken(sNodeLoad)



                    Dim sID As String = poSelector.NodeID.Substring(poSelector.NodeID.IndexOf("]"c) + 2)
                    Dim sText As String = poSelector.NodeText.Substring(poSelector.NodeID.IndexOf("]"c) + 2)

                    poEditor.Properties.Items.Clear()

                    If oObject IsNot Nothing Then
                         For Each oRow As JObject In oObject
                              poEditor.Properties.Items.Add(New DevExpress.XtraEditors.Controls.ImageComboBoxItem With {.Description = FormatCase("T", oRow.SelectToken(sText)).Trim, .Value = oRow.SelectToken(sID)})
                         Next
                    End If
               End If

          Catch ex As Exception
               MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try

     End Sub

     Private Sub LoadLinkedComboBox(sender As Object, e As EventArgs)
          Dim oSelectedSelector As clsSelectors = TryCast(sender.tag, clsSelectors)
          Dim oComboBox As ImageComboBoxEdit = TryCast(sender, ImageComboBoxEdit)

          If oSelectedSelector IsNot Nothing And oComboBox IsNot Nothing Then
               For Each oControl As ImageComboBoxEdit In listOfContorls
                    Dim oSelector As clsSelectors = TryCast(oControl.Tag, clsSelectors)
                    If oSelector IsNot Nothing Then
                         If oSelector.Query.Contains(String.Format("<<!{0}!>>", oSelectedSelector.Priority)) = True Then
                              Dim sQuery As String = Replace(oSelector.Query, String.Format("<<!{0}!>>", oSelectedSelector.Priority), oComboBox.EditValue.ToString)
                              LoadComboBox(oControl, oSelector, sQuery)
                         End If
                    End If
               Next
          End If

     End Sub
     Private Sub ucSelectorQuery_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating

          Try

               For Each oControl As Control In listOfContorls
                    Dim oSelector As clsSelectors = TryCast(oControl.Tag, clsSelectors)
                    Dim svalue As String = ""

                    Select Case oControl.GetType.Name
                         Case "DateEdit"
                              svalue = TryCast(oControl, DateEdit).Text
                         Case "textEdit"
                              svalue = TryCast(oControl, TextEdit).Text
                         Case "ImageComboBoxEdit"
                              If TryCast(oControl, ImageComboBoxEdit).EditValue IsNot Nothing Then
                                   svalue = TryCast(oControl, ImageComboBoxEdit).EditValue.ToString
                              End If
                    End Select

                    _Query = Replace(_Query, oSelector.Variable, svalue)


               Next

          Catch ex As Exception
               MsgBox(String.Format("Error code {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try

     End Sub
End Class
