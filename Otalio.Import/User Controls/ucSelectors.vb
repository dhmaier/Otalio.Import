Imports System.Net
Imports System.IO
Imports System.Web
Imports RestSharp
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.Linq
Imports DevExpress.Spreadsheet


Public Class ucSelectors

     Private moSelector As New clsSelectors
     Private mbLoading As Boolean = False

     Public Property _Selector As clsSelectors
          Get
               Return moSelector
          End Get
          Set(value As clsSelectors)
               If value IsNot Nothing Then
                    moSelector = Nothing
                    moSelector = value

                    If moSelector.Enabled = "" Then moSelector.Enabled = "1"
                    If moSelector.LookUpType Is Nothing Then moSelector.LookUpType = New clsLookUpTypes
                    If moSelector.Visibility = "" Then moSelector.Visibility = "1"

                    With txtPriority
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "Priority"))
                    End With

                    With txtAPIEndpoint
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "APIEndpoint"))
                    End With


                    With txtHeader
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "Label"))
                    End With

                    With txtAPIQuery
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "Query"))
                    End With


                    With icbType
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moSelector, "ValidationType"))
                    End With

                    With txtComments
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "Variable"))
                    End With


                    With txtUUID
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "ID"))
                    End With

                    With icbEnabled
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moSelector, "Enabled"))
                    End With


                    With icbVisibility
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moSelector, "Visibility"))
                    End With

                    With lueLookupTypes
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("EditValue", moSelector.LookUpType, "Id"))
                    End With

                    With cbNodeID
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "NodeID"))
                    End With

                    With cbNodeText
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moSelector, "NodeText"))
                    End With

                    If moSelector.LookUpType IsNot Nothing Then
                         Dim olookup As clsLookUpTypes = goLookupTypes.Where(Function(n) n.Id = moSelector.LookUpType.Id).FirstOrDefault
                         If olookup IsNot Nothing Then
                              moSelector.LookUpType = olookup
                              lueLookupTypes.EditValue = olookup
                         End If
                    End If



               End If
          End Set
     End Property

     Private Sub txtAPIQuery_EditValueChanged(sender As Object, e As EventArgs) Handles txtAPIQuery.EditValueChanged


     End Sub


     Private Sub icbType_EditValueChanged(sender As Object, e As EventArgs) Handles icbType.EditValueChanged
          mbLoading = True

          Select Case icbType.EditValue.ToString
               Case "LOOKUP"

                    lciLookUpType.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lueLookupTypes.Properties.DataSource = goLookupTypes

                    If moSelector.LookUpType IsNot Nothing Then
                         lueLookupTypes.EditValue = moSelector.LookUpType
                    End If

                    moSelector.APIEndpoint = String.Format("metadata/v1/lookup-values?hierarchyId=@@HIERARCHYID@@")

                    If String.IsNullOrEmpty(txtAPIQuery.EditValue) = True Then
                         moSelector.Query = String.Format("code=={0}<!XxX!>{0}", ControlChars.Quote)
                    End If


               Case Else
                    lciLookUpType.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never

          End Select



          mbLoading = False
     End Sub

     Private Sub lueLookupTypes_EditValueChanged(sender As Object, e As EventArgs) Handles lueLookupTypes.EditValueChanged


          If icbType.EditValue = "5" Or icbType.EditValue = "6" Or icbType.EditValue = "7" Then
               If moSelector.LookUpType IsNot Nothing Then
                    '    If String.IsNullOrEmpty(txtComments.EditValue) = True Then
                    Dim oLookup As clsLookUpTypes = TryCast(lueLookupTypes.GetSelectedDataRow, clsLookUpTypes)
                    If oLookup IsNot Nothing Then

                         moSelector.LookUpType = oLookup

                         moSelector.Variable = String.Format("{0}", oLookup.Code, oLookup.Description)
                         txtComments.EditValue = moSelector.Variable

                         If icbType.EditValue = "7" Then
                              moSelector.Query = String.Format("lookupGroup=={2}{0}{2}", oLookup.LookupGroup, oLookup.Code, ControlChars.Quote)
                              txtAPIQuery.EditValue = moSelector.Query
                         End If
                    End If
                    '   End If
               End If
          End If
     End Sub

     Private Sub cbReturnNode_EditValueChanged(sender As Object, e As EventArgs)
          'If String.IsNullOrEmpty(cbReturnNode.EditValue) = True Then
          '     moValidation.ReturnNodeValue = "responsePayload.content[0].entityId"
          'Else
          '     moValidation.ReturnNodeValue = cbReturnNode.EditValue
          'End If
     End Sub

     Public Sub LoadFromJsonfile()
          moSelector.LoadSelectorFromJsonFile()
     End Sub

     Public Sub ExportToJsonFile()
          moSelector.ExportToJson()
     End Sub
End Class
