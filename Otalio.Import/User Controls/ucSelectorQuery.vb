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

''' <summary>
''' User control for loading Developer Express editors dynamically based on a list of clsSelectors.
''' </summary>
Public Class ucSelectorQuery

     Private moSelector As List(Of clsSelectors)
     Private listOfControls As New List(Of Control)

     ''' <summary>
     ''' Gets or sets the list of selectors.
     ''' </summary>
     Public Property _Selectors As List(Of clsSelectors)
          Get
               Return moSelector
          End Get
          Set(value As List(Of clsSelectors))
               moSelector = value
               LoadSelectors()
          End Set
     End Property

     ''' <summary>
     ''' Gets or sets the query string.
     ''' </summary>
     Public Property _Query As String

     ''' <summary>
     ''' Loads the selectors and creates corresponding controls.
     ''' </summary>
     Public Sub LoadSelectors()
          Try
               ' Reset memory
               goQueryMemory = New Dictionary(Of String, String)

               If moSelector IsNot Nothing AndAlso moSelector.Count > 0 Then
                    For Each oSelector In moSelector.OrderBy(Function(n) n.Priority).ToList()
                         Select Case oSelector.ValidationType
                              Case "TEXT"
                                   CreateTextEdit(oSelector)
                              Case "DATE"
                                   CreateDateEdit(oSelector)
                              Case "LIST", "LOOKUP"
                                   CreateImageComboBoxEdit(oSelector)
                              Case "CHECKBOXLIST"
                                   CreateCheckedComboBoxEdit(oSelector)
                              Case "BOOLEAN"
                                   CreateCheckEdit(oSelector)
                              Case "CHECKBOXPREDEFINEDLIST"
                                   CreateCheckBoxPredefinedList(oSelector)
                         End Select
                    Next
               End If
          Catch ex As Exception
               MsgBox(String.Format("Error loading selectors: {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try
     End Sub

     ''' <summary>
     ''' Creates a TextEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateTextEdit(oSelector As clsSelectors)
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
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Creates a DateEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateDateEdit(oSelector As clsSelectors)
          Dim oEditor As New DateEdit
          oEditor.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom
          oEditor.Properties.EditFormat.FormatString = "yyyy-MM-dd"
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
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Creates an ImageComboBoxEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateImageComboBoxEdit(oSelector As clsSelectors)
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
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Creates a CheckedComboBoxEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateCheckedComboBoxEdit(oSelector As clsSelectors)
          Dim oEditor As New CheckedComboBoxEdit
          Dim oLayoutControlItem As New LayoutControlItem
          With oLayoutControlItem
               .Text = oSelector.Label
               .Control = oEditor
               .Location = New System.Drawing.Point(0, 0)
               .Size = New System.Drawing.Size(606, 24)
               .TextSize = New System.Drawing.Size(96, 13)
          End With
          oEditor.Properties.SelectAllItemVisible = True
          Me.Root.AddItem(oLayoutControlItem)
          LoadCheckComboBox(oEditor, oSelector, oSelector.Query)
          AddHandler oEditor.EditValueChanged, AddressOf LoadLinkedCheckComboBox
          oEditor.Tag = oSelector
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Creates a CheckEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateCheckEdit(oSelector As clsSelectors)
          Dim oEditor As New CheckEdit()
          Dim oLayoutControlItem As New LayoutControlItem()
          With oLayoutControlItem
               .Text = oSelector.Label
               .Control = oEditor
               .Location = New System.Drawing.Point(0, 0)
               .Size = New System.Drawing.Size(606, 24)
               .TextSize = New System.Drawing.Size(96, 13)
               .TextVisible = True
          End With
          oEditor.Text = "Enabled"
          oEditor.EditValue = Nothing
          Me.Root.AddItem(oLayoutControlItem)
          oEditor.Tag = oSelector
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Creates a predefined list in an ImageComboBoxEdit control for the specified selector.
     ''' </summary>
     ''' <param name="oSelector">The selector.</param>
     Private Sub CreateCheckBoxPredefinedList(oSelector As clsSelectors)
          Dim oEditor As New ImageComboBoxEdit()
          oEditor.Properties.Sorted = True
          Dim oLayoutControlItem As New LayoutControlItem()
          With oLayoutControlItem
               .Text = oSelector.Label
               .Control = oEditor
               .Location = New System.Drawing.Point(0, 0)
               .Size = New System.Drawing.Size(606, 24)
               .TextSize = New System.Drawing.Size(96, 13)
          End With
          Me.Root.AddItem(oLayoutControlItem)
          LoadPredefinedList(oEditor, oSelector)
          oEditor.Tag = oSelector
          listOfControls.Add(oEditor)
     End Sub

     ''' <summary>
     ''' Loads data into an ImageComboBoxEdit control.
     ''' </summary>
     ''' <param name="poEditor">The editor control.</param>
     ''' <param name="poSelector">The selector.</param>
     ''' <param name="psQuery">The query string.</param>
     Private Sub LoadComboBox(poEditor As ImageComboBoxEdit, poSelector As clsSelectors, psQuery As String)
          Try
               ' DONT populate if linked to a selector
               If Not psQuery.Contains("<<!") Then
                    Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet(poSelector.APIEndpoint, String.Empty, psQuery)
                    Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
                    Dim oObject As JToken = jsServerResponse.SelectToken("responsePayload.content")

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
               MsgBox(String.Format("Error loading combo box: {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try
     End Sub

     ''' <summary>
     ''' Loads data into a CheckedComboBoxEdit control.
     ''' </summary>
     ''' <param name="poEditor">The editor control.</param>
     ''' <param name="poSelector">The selector.</param>
     ''' <param name="psQuery">The query string.</param>
     Private Sub LoadCheckComboBox(poEditor As CheckedComboBoxEdit, poSelector As clsSelectors, psQuery As String)
          Try
               ' DONT populate if linked to a selector
               If Not psQuery.Contains("<<!") Then
                    Dim sID As String = poSelector.NodeID.Substring(poSelector.NodeID.IndexOf("]"c) + 2)
                    Dim sText As String = poSelector.NodeText.Substring(poSelector.NodeID.IndexOf("]"c) + 2)
                    Dim sSort As String = If(sText.Contains("."), GetLastValue(sText, "."), sText)

                    Dim oResponse As IRestResponse = goHTTPServer.CallWebEndpointUsingGet(poSelector.APIEndpoint, String.Empty, psQuery, sSort)
                    Dim jsServerResponse As JObject = JObject.Parse(oResponse.Content)
                    Dim oObject As JToken = jsServerResponse.SelectToken("responsePayload.content")

                    poEditor.Properties.Items.Clear()

                    If oObject IsNot Nothing Then
                         For Each oRow As JObject In oObject
                              poEditor.Properties.Items.Add(New DevExpress.XtraEditors.Controls.CheckedListBoxItem With {.Description = FormatCase("T", oRow.SelectToken(sText)).Trim, .Value = oRow.SelectToken(sID)})
                         Next
                    End If
               End If
          Catch ex As Exception
               MsgBox(String.Format("Error loading checked combo box: {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try
     End Sub

     ''' <summary>
     ''' Loads a predefined list into an ImageComboBoxEdit control.
     ''' </summary>
     ''' <param name="poEditor">The editor control.</param>
     ''' <param name="poSelector">The selector.</param>
     Private Sub LoadPredefinedList(poEditor As ImageComboBoxEdit, poSelector As clsSelectors)
          Try
               Dim predefinedValues As String = poSelector.Query
               Dim defaultValue As String = Nothing

               ' Check if there's a :DEFAULT=value instruction
               If predefinedValues.Contains(":DEFAULT=") Then
                    Dim parts As String() = predefinedValues.Split(New String() {":DEFAULT="}, StringSplitOptions.None)
                    predefinedValues = parts(0).Trim()
                    defaultValue = parts(1).Trim()
               End If

               Dim values As String() = predefinedValues.Split(","c)

               poEditor.Properties.Items.Clear()

               For Each value In values
                    poEditor.Properties.Items.Add(New DevExpress.XtraEditors.Controls.ImageComboBoxItem(value.Trim(), value.Trim()))
               Next

               ' Preselect the default value if specified
               If defaultValue IsNot Nothing Then
                    poEditor.SelectedItem = poEditor.Properties.Items.Cast(Of DevExpress.XtraEditors.Controls.ImageComboBoxItem)().
                FirstOrDefault(Function(item) item.Value.ToString() = defaultValue)
               End If

          Catch ex As Exception
               MsgBox(String.Format("Error loading predefined list: {0} - {1}{2}{3}", ex.HResult, ex.Message, vbNewLine, ex.StackTrace))
          End Try
     End Sub


     ''' <summary>
     ''' Handles the EditValueChanged event for linked combo boxes.
     ''' </summary>
     Private Sub LoadLinkedComboBox(sender As Object, e As EventArgs)
          Dim oSelectedSelector As clsSelectors = TryCast(sender.tag, clsSelectors)
          Dim oComboBox As ImageComboBoxEdit = TryCast(sender, ImageComboBoxEdit)

          If oSelectedSelector IsNot Nothing And oComboBox IsNot Nothing Then
               For Each oControl As Control In listOfControls
                    Dim oSelector As clsSelectors = TryCast(oControl.Tag, clsSelectors)
                    If oSelector IsNot Nothing AndAlso oSelector.Query.Contains(String.Format("<<!{0}!>>", oSelectedSelector.Priority)) Then
                         Dim sQuery As String = Replace(oSelector.Query, String.Format("<<!{0}!>>", oSelectedSelector.Priority), oComboBox.EditValue.ToString())
                         Select Case oSelector.ValidationType
                              Case "LIST"
                                   LoadComboBox(oControl, oSelector, sQuery)
                              Case "CHECKBOXLIST"
                                   LoadCheckComboBox(oControl, oSelector, sQuery)
                         End Select
                    End If
               Next
          End If
     End Sub

     ''' <summary>
     ''' Handles the EditValueChanged event for linked checked combo boxes.
     ''' </summary>
     Private Sub LoadLinkedCheckComboBox(sender As Object, e As EventArgs)
          Dim oSelectedSelector As clsSelectors = TryCast(sender.Tag, clsSelectors)
          Dim oComboBox As CheckedComboBoxEdit = TryCast(sender, CheckedComboBoxEdit)

          If oSelectedSelector IsNot Nothing And oComboBox IsNot Nothing Then
               For Each oControl As Control In listOfControls
                    Dim oSelector As clsSelectors = TryCast(oControl.Tag, clsSelectors)
                    If oSelector IsNot Nothing AndAlso oSelector.Query.Contains(String.Format("<<!{0}!>>", oSelectedSelector.Priority)) Then
                         Dim sQuery As String = Replace(oSelector.Query, String.Format("<<!{0}!>>", oSelectedSelector.Priority), GetCheckedItemsValueString(oComboBox))
                         sQuery = RemoveConditionFromRsqlQuery(sQuery, "!!ALLITEMS!!")
                         Select Case oSelector.ValidationType
                              Case "LIST"
                                   LoadComboBox(oControl, oSelector, sQuery)
                              Case "CHECKBOXLIST"
                                   LoadCheckComboBox(oControl, oSelector, sQuery)
                         End Select
                    End If
               Next
          End If
     End Sub

     ''' <summary>
     ''' Gets the selected items from a CheckedComboBoxEdit control as a comma-separated string.
     ''' </summary>
     ''' <param name="oComboBox">The checked combo box.</param>
     ''' <returns>The selected items as a comma-separated string.</returns>
     Private Function GetCheckedItemsValueString(oComboBox As CheckedComboBoxEdit) As String
          Dim selectedItems As String = String.Empty
          Dim sCheckItems As String = oComboBox.Properties.GetCheckedItems().ToString()

          If sCheckItems.Split(","c).Count < oComboBox.Properties.Items.Count Then
               For Each item In sCheckItems.Split(","c).ToList()
                    If String.IsNullOrEmpty(selectedItems) Then
                         selectedItems = item.ToString()
                    Else
                         selectedItems += "," & item.ToString()
                    End If
               Next
               selectedItems = RemoveLastComma(selectedItems)
          Else
               selectedItems = "!!ALLITEMS!!"
          End If

          Return selectedItems
     End Function

     ''' <summary>
     ''' Handles the Validating event for the user control.
     ''' </summary>
     Public Sub ucSelectorQuery_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating
          Try
               For Each oControl As Control In listOfControls
                    Dim oSelector As clsSelectors = TryCast(oControl.Tag, clsSelectors)
                    Dim svalue As String = String.Empty

                    Select Case oControl.GetType().Name
                         Case "DateEdit"
                              svalue = TryCast(oControl, DateEdit).Text
                              If Not IsDate(svalue) Then
                                   _Query = RemoveProperty(_Query, oSelector.Variable)
                              End If
                         Case "TextEdit"
                              svalue = TryCast(oControl, TextEdit).Text
                              If String.IsNullOrEmpty(svalue) Then
                                   _Query = RemoveProperty(_Query, oSelector.Variable)
                              End If
                         Case "ImageComboBoxEdit"
                              If TryCast(oControl, ImageComboBoxEdit).EditValue IsNot Nothing Then
                                   svalue = TryCast(oControl, ImageComboBoxEdit).EditValue.ToString()
                              Else
                                   _Query = RemoveProperty(_Query, oSelector.Variable)
                              End If
                         Case "CheckedComboBoxEdit"
                              If TryCast(oControl, CheckedComboBoxEdit).EditValue IsNot Nothing Then
                                   Dim selectedItems As String = TryCast(oControl, CheckedComboBoxEdit).Properties.GetCheckedItems().ToString()
                                   If selectedItems.Split(","c).Count = TryCast(oControl, CheckedComboBoxEdit).Properties.Items.Count OrElse String.IsNullOrEmpty(selectedItems) Then
                                        _Query = RemoveProperty(_Query, oSelector.Variable)
                                   Else
                                        svalue = RemoveLastComma(selectedItems)
                                   End If
                              End If
                         Case "CheckEdit"
                              If TryCast(oControl, CheckEdit).EditValue IsNot Nothing Then
                                   svalue = TryCast(oControl, CheckEdit).EditValue.ToString()
                              End If
                              If String.IsNullOrEmpty(svalue) Then
                                   _Query = RemoveProperty(_Query, oSelector.Variable)
                              End If
                    End Select

                    If Not String.IsNullOrEmpty(svalue) Then
                         _Query = Replace(_Query, oSelector.Variable, svalue)
                         If Not goQueryMemory.ContainsKey(oSelector.Variable) Then
                              goQueryMemory.Add(oSelector.Variable, svalue)
                         End If
                    End If
               Next
          Catch ex As Exception
               ShowErrorForm(ex)
          End Try
     End Sub

     ''' <summary>
     ''' Removes a property from the query string if its value matches the specified value.
     ''' </summary>
     ''' <param name="psQuery">The query string.</param>
     ''' <param name="valueToRemove">The value to remove.</param>
     ''' <returns>The updated query string.</returns>
     Private Function RemoveProperty(psQuery As String, valueToRemove As String) As String
          Dim sValue As String = String.Empty
          If IsJsonObject(psQuery) Then
               sValue = RemovePropertyIfValueMatches(psQuery, valueToRemove)
          Else
               sValue = RemoveConditionFromRsqlQuery(psQuery, valueToRemove)
          End If
          Return sValue
     End Function

     ''' <summary>
     ''' Calculates the height of the user control based on the controls it contains.
     ''' </summary>
     ''' <returns>The total height of the user control.</returns>
     Public Function CalculateUserControlHeight() As Integer
          Const padding As Integer = 20
          Const controlSpacing As Integer = 5

          Dim totalHeight As Integer = padding ' Initial padding for the user control

          If listOfControls IsNot Nothing AndAlso listOfControls.Count > 0 Then
               For Each control As Control In listOfControls
                    totalHeight += control.Height + controlSpacing
               Next
          End If

          totalHeight += padding ' Add padding at the bottom

          Return totalHeight
     End Function

End Class
