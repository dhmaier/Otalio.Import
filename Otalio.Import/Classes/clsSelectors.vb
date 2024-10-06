Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Imports System.IO
Imports System.Windows.Forms

<
  XmlInclude(GetType(clsColumnProperties)),
  Serializable()
  >
Public Class clsSelectors

     Public Property ID As String = System.Guid.NewGuid.ToString

     Public Property Priority As String = ""

     Public Property APIEndpoint As String = ""

     Public Property Query As String = ""

     Public Property Label As String = ""

     Public Property ValidationType As String = "LIST"

     Public Property ReturnNodeValue As String = ""

     Public Property ReturnCell As String = ""

     Public Property Variable As String = ""

     Public Property Enabled As String = "1"

     Public Property Formatted As String = ""

     Public Property LookUpType As New clsLookUpTypes

     Public Property Sort As String = ""

     Public Property Visibility As String = "1"

     Public Property GraphQLSourceNode As String = ""

     Public Property NodeID As String = ""
     Public Property NodeText As String = ""




     Public Function Clone() As clsSelectors
          Dim oObject As clsSelectors = DirectCast(Me.MemberwiseClone(), clsSelectors)
          oObject.ID = System.Guid.NewGuid.ToString
          Return oObject
     End Function

     Private Sub loadFromJson(jsonNode As JObject)
          Try
               If jsonNode Is Nothing Then
                    Throw New ArgumentNullException(NameOf(jsonNode), "JSON node cannot be null.")
               End If

               Me.ID = jsonNode("ID")?.ToString()
               Me.Priority = jsonNode("Priority")?.ToString()
               Me.APIEndpoint = jsonNode("APIEndpoint")?.ToString()
               Me.Query = jsonNode("Query")?.ToString()
               Me.Label = jsonNode("Label")?.ToString()
               Me.ValidationType = jsonNode("ValidationType")?.ToString()
               Me.ReturnNodeValue = jsonNode("ReturnNodeValue")?.ToString()
               Me.ReturnCell = jsonNode("ReturnCell")?.ToString()
               Me.Variable = jsonNode("Variable")?.ToString()
               Me.Enabled = jsonNode("Enabled")?.ToString()
               Me.Formatted = jsonNode("Formatted")?.ToString()
               Me.Sort = jsonNode("Sort")?.ToString()
               Me.Visibility = jsonNode("Visibility")?.ToString()
               Me.GraphQLSourceNode = jsonNode("GraphQLSourceNode")?.ToString()
               Me.NodeID = jsonNode("NodeID")?.ToString()
               Me.NodeText = jsonNode("NodeText")?.ToString()
          Catch ex As ArgumentNullException
               MessageBox.Show(ex.Message, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          Catch ex As Exception
               MessageBox.Show(String.Format("Unexpected error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          End Try
     End Sub

     Public Sub LoadSelectorFromJsonFile()
          Dim openFileDialog As New OpenFileDialog With {
              .Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
              .Title = "Select a JSON file"
          }

          Try
               If openFileDialog.ShowDialog() = DialogResult.OK Then
                    Dim filePath As String = openFileDialog.FileName
                    Dim jsonContent As String = File.ReadAllText(filePath)

                    ' Validate JSON content
                    If String.IsNullOrWhiteSpace(jsonContent) Then
                         MessageBox.Show("The selected file is empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                         Return
                    End If

                    Dim jsonObject As JObject = JObject.Parse(jsonContent)
                    loadFromJson(jsonObject)

               End If
          Catch ex As JsonReaderException
               MessageBox.Show(String.Format("Invalid JSON format: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          Catch ex As FileNotFoundException
               MessageBox.Show(String.Format("File not found: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          Catch ex As IOException
               MessageBox.Show(String.Format("IO error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          Catch ex As Exception
               MessageBox.Show(String.Format("Unexpected error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
          End Try
     End Sub

     Public Sub ExportToJson()
          Dim defaultFileName As String = If(Not String.IsNullOrWhiteSpace(Me.Label), Me.Label, "default")

          Dim saveFileDialog As New SaveFileDialog With {
            .Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            .Title = "Select a location to save the JSON file",
            .DefaultExt = "json",
            .FileName = defaultFileName & ".json"
        }

          If saveFileDialog.ShowDialog() = DialogResult.OK Then
               Dim filePath As String = saveFileDialog.FileName
               Try
                    Dim jsonContent As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
                    File.WriteAllText(filePath, jsonContent)
                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
               Catch ex As IOException
                    MessageBox.Show(String.Format("IO error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
               Catch ex As Exception
                    MessageBox.Show(String.Format("Unexpected error: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
               End Try
          End If
     End Sub


End Class
