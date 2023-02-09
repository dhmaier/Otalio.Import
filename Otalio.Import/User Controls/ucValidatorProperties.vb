Imports System.Net
Imports System.IO
Imports System.Web
Imports RestSharp
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.Linq
Imports DevExpress.Spreadsheet


Public Class ucValidatorProperties

     Private moValidation As New clsValidation
     Private mbLoading As Boolean = False

     Public Property _Validation As clsValidation
          Get
               Return moValidation
          End Get
          Set(value As clsValidation)
               If value IsNot Nothing Then
                    moValidation = Nothing
                    moValidation = value

                    If moValidation.Enabled = "" Then moValidation.Enabled = "1"
                    If moValidation.LookUpType Is Nothing Then moValidation.LookUpType = New clsLookUpTypes
                    If moValidation.Visibility = "" Then moValidation.Visibility = "1"

                    With txtPriority
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "Priority"))
                    End With

                    With txtAPIEndpoint
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "APIEndpoint"))
                    End With


                    With txtHeader
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "Headers"))
                    End With

                    With txtAPIQuery
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "Query"))
                    End With


                    With icbType
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moValidation, "ValidationType"))
                    End With

                    'With txtReturnNode
                    '     .DataBindings.Clear()
                    '     .DataBindings.Add(New Binding("Text", moValidation, "ReturnNodeValue"))
                    'End With


                    With txtReturnValueColumn
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "ReturnCell"))
                    End With

                    With txtComments
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "Comments"))
                    End With


                    With txtUUID
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "ID"))
                    End With

                    With icbEnabled
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moValidation, "Enabled"))
                    End With

                    With icbFormat
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moValidation, "Formatted"))
                    End With

                    With icbVisibility
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Value", moValidation, "Visibility"))
                    End With

                    With lueLookupTypes
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("EditValue", moValidation.LookUpType, "Id"))
                    End With

                    With txtGraphQLNode
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "GraphQLSourceNode"))
                    End With

                    With cbReturnNode
                         .DataBindings.Clear()
                         .DataBindings.Add(New Binding("Text", moValidation, "ReturnNodeValue"))
                    End With

                    If moValidation.LookUpType IsNot Nothing Then
                         Dim olookup As clsLookUpTypes = goLookupTypes.Where(Function(n) n.Id = moValidation.LookUpType.Id).FirstOrDefault
                         If olookup IsNot Nothing Then
                              moValidation.LookUpType = olookup
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

          Select Case icbType.EditValue
               Case "5", "6"

                    lciLookUpType.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lueLookupTypes.Properties.DataSource = goLookupTypes

                    If moValidation.LookUpType IsNot Nothing Then
                         lueLookupTypes.EditValue = moValidation.LookUpType
                    End If

                    moValidation.APIEndpoint = String.Format("metadata/v1/lookup-values?hierarchyId=@@HIERARCHYID@@")

                    If String.IsNullOrEmpty(txtAPIQuery.EditValue) = True Then
                         moValidation.Query = String.Format("code=={0}<!XxX!>{0}", ControlChars.Quote)
                    End If

                    moValidation.ReturnNodeValue = "responsePayload.content[0].entityId"


               Case "7"

                    lciLookUpType.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always
                    lueLookupTypes.Properties.DataSource = goLookupTypes

                    If moValidation.LookUpType IsNot Nothing Then
                         lueLookupTypes.EditValue = moValidation.LookUpType.Id
                    End If

                    If String.IsNullOrEmpty(txtAPIEndpoint.EditValue) = True Then
                         moValidation.APIEndpoint = "metadata/v1/lookup-types"
                    End If


                    If String.IsNullOrEmpty(cbReturnNode.EditValue) = True Then
                         moValidation.ReturnNodeValue = "responsePayload.content[0].entityId"

                    End If
               Case Else
                    lciLookUpType.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never

          End Select



          mbLoading = False
     End Sub

     Private Sub lueLookupTypes_EditValueChanged(sender As Object, e As EventArgs) Handles lueLookupTypes.EditValueChanged


          If icbType.EditValue = "5" Or icbType.EditValue = "6" Or icbType.EditValue = "7" Then
               If moValidation.LookUpType IsNot Nothing Then
                    '    If String.IsNullOrEmpty(txtComments.EditValue) = True Then
                    Dim oLookup As clsLookUpTypes = TryCast(lueLookupTypes.GetSelectedDataRow, clsLookUpTypes)
                    If oLookup IsNot Nothing Then

                         moValidation.LookUpType = oLookup

                         moValidation.Comments = String.Format("{0}", oLookup.Code, oLookup.Description)
                         txtComments.EditValue = moValidation.Comments

                         If icbType.EditValue = "7" Then
                              moValidation.Query = String.Format("lookupGroup=={2}{0}{2};code=={2}{1}{2}", oLookup.LookupGroup, oLookup.Code, ControlChars.Quote)
                              txtAPIQuery.EditValue = moValidation.Query
                         End If
                    End If
                    '   End If
               End If
          End If
     End Sub

     Private Sub cbReturnNode_EditValueChanged(sender As Object, e As EventArgs) Handles cbReturnNode.EditValueChanged
          'If String.IsNullOrEmpty(cbReturnNode.EditValue) = True Then
          '     moValidation.ReturnNodeValue = "responsePayload.content[0].entityId"
          'Else
          '     moValidation.ReturnNodeValue = cbReturnNode.EditValue
          'End If
     End Sub
End Class
