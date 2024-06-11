Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module bJson

     Function ValidateJson(jsonString As String) As String
          Try
               ' Parse the JSON string
               Dim json As JObject = JObject.Parse(jsonString)
               Return "OK"
          Catch ex As JsonReaderException
               ' Add line numbers to the JSON string
               Dim numberedJsonString As String = AddLineNumbers(jsonString)

               ' Return detailed error message if JSON is not valid
               Dim errorMessage As String = $"Invalid JSON detected. Error details:" & Environment.NewLine &
                                         $"Line: {ex.LineNumber}" & Environment.NewLine &
                                         $"Position: {ex.LinePosition}" & Environment.NewLine &
                                         $"Message: {ex.Message}" & Environment.NewLine &
                                         $"Error Path: {ex.Path}" & Environment.NewLine &
                                         $"Complete JSON string with line numbers:" & Environment.NewLine &
                                         $"{numberedJsonString}"
               Return errorMessage
          Catch ex As Exception
               ' Add line numbers to the JSON string
               Dim numberedJsonString As String = AddLineNumbers(jsonString)

               ' Return general error message for any other exception
               Dim errorMessage As String = $"An unexpected error occurred while validating the JSON." & Environment.NewLine &
                                         $"Message: {ex.Message}" & Environment.NewLine &
                                         $"Complete JSON string with line numbers:" & Environment.NewLine &
                                         $"{numberedJsonString}"
               Return errorMessage
          End Try
     End Function

     Private Function AddLineNumbers(jsonString As String) As String
          Dim lines As String() = jsonString.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
          For i As Integer = 0 To lines.Length - 1
               lines(i) = $"{i + 1}: {lines(i)}"
          Next
          Return String.Join(Environment.NewLine, lines)
     End Function


     Public Function RemovePropertyIfValueMatches(jsonString As String, valueToMatch As String) As String
          Try
               Dim originalJsonObject As JObject = JObject.Parse(jsonString)
               Dim jsonObject As JObject = JObject.Parse(jsonString)

               ' Recursive function to remove properties with the specified value
               RemoveMatchingProperties(jsonObject, valueToMatch)

               ' Log the differences between the original and the processed JSON strings
               LogDifferences(originalJsonObject, jsonObject)

               Return jsonObject.ToString()
          Catch ex As Exception
               Throw New Exception("An error occurred while processing the JSON.", ex)
          End Try
     End Function

     Private Sub RemoveMatchingProperties(jToken As JToken, valueToMatch As String)
          If jToken.Type = JTokenType.Object Then
               Dim propertiesToRemove As New List(Of JProperty)

               ' Collect properties to remove
               For Each prop As JProperty In jToken.Children(Of JProperty)()
                    If prop.Value.Type = JTokenType.String AndAlso prop.Value.ToString() = valueToMatch Then
                         Console.WriteLine($"Marking property '{prop.Name}' for removal")
                         propertiesToRemove.Add(prop)
                    ElseIf prop.Value.Type = JTokenType.Array Then
                         RemoveMatchingProperties(prop.Value, valueToMatch)
                         ' If the array is now empty after removing elements, mark it for removal
                         If Not prop.Value.HasValues Then
                              propertiesToRemove.Add(prop)
                         End If
                    Else
                         RemoveMatchingProperties(prop.Value, valueToMatch)
                    End If
               Next

               ' Remove collected properties
               For Each prop As JProperty In propertiesToRemove
                    Console.WriteLine($"Removing property '{prop.Name}'")
                    prop.Remove()
               Next
          ElseIf jToken.Type = JTokenType.Array Then
               Dim itemsToRemove As New List(Of JToken)

               For Each item As JToken In jToken.Children()
                    If item.Type = JTokenType.String AndAlso item.ToString() = valueToMatch Then
                         itemsToRemove.Add(item)
                    Else
                         RemoveMatchingProperties(item, valueToMatch)
                    End If
               Next

               ' Remove collected items
               For Each item As JToken In itemsToRemove
                    item.Remove()
               Next
          End If
     End Sub

     Private Sub LogDifferences(originalJsonObject As JObject, processedJsonObject As JObject)
          Console.WriteLine("Original JSON:")
          Console.WriteLine(originalJsonObject.ToString())

          Console.WriteLine("Processed JSON:")
          Console.WriteLine(processedJsonObject.ToString())

          Dim differences As New List(Of String)
          GetDifferences(originalJsonObject, processedJsonObject, differences, "")

          Console.WriteLine("Differences:")
          For Each difference As String In differences
               Console.WriteLine(difference)
          Next
     End Sub

     Private Sub GetDifferences(originalToken As JToken, processedToken As JToken, differences As List(Of String), currentPath As String)
          If originalToken.Type = JTokenType.Object Then
               Dim originalObject As JObject = CType(originalToken, JObject)
               Dim processedObject As JObject = CType(processedToken, JObject)

               For Each prop As JProperty In originalObject.Properties()
                    Dim newPath As String = If(String.IsNullOrEmpty(currentPath), prop.Name, $"{currentPath}.{prop.Name}")
                    If Not processedObject.ContainsKey(prop.Name) Then
                         differences.Add($"Removed: {newPath}")
                    Else
                         GetDifferences(prop.Value, processedObject(prop.Name), differences, newPath)
                    End If
               Next
          ElseIf originalToken.Type = JTokenType.Array Then
               Dim originalArray As JArray = CType(originalToken, JArray)
               Dim processedArray As JArray = CType(processedToken, JArray)

               For i As Integer = 0 To originalArray.Count - 1
                    Dim newPath As String = $"{currentPath}[{i}]"
                    If i >= processedArray.Count Then
                         differences.Add($"Removed: {newPath}")
                    Else
                         GetDifferences(originalArray(i), processedArray(i), differences, newPath)
                    End If
               Next
          End If
     End Sub

     Public Function RemoveEmptyOrNullProperties(jsonString As String) As String
          Try
               Dim originalJsonObject As JObject = JObject.Parse(jsonString)
               Dim jsonObject As JObject = JObject.Parse(jsonString)

               ' Recursive function to remove empty or null properties
               RemoveEmptyOrNullPropertiesFromToken(jsonObject)

               ' Log the differences between the original and the processed JSON strings
               LogDifferences(originalJsonObject, jsonObject)

               Return jsonObject.ToString()
          Catch ex As Exception
               Throw New Exception("An error occurred while processing the JSON.", ex)
          End Try
     End Function

     Private Sub RemoveEmptyOrNullPropertiesFromToken(jToken As JToken)
          If jToken.Type = JTokenType.Object Then
               Dim propertiesToRemove As New List(Of JProperty)

               ' Collect properties to remove
               For Each prop As JProperty In jToken.Children(Of JProperty)()
                    If prop.Value.Type = JTokenType.Null OrElse
                (prop.Value.Type = JTokenType.String AndAlso String.IsNullOrEmpty(prop.Value.ToString())) OrElse
                (prop.Value.Type = JTokenType.Object AndAlso Not prop.Value.HasValues) OrElse
                (prop.Value.Type = JTokenType.Array AndAlso Not prop.Value.HasValues) Then

                         propertiesToRemove.Add(prop)
                    Else
                         RemoveEmptyOrNullPropertiesFromToken(prop.Value)
                         ' Re-check if the object became empty after recursive call
                         If (prop.Value.Type = JTokenType.Object AndAlso Not prop.Value.HasValues) OrElse
                   (prop.Value.Type = JTokenType.Array AndAlso Not prop.Value.HasValues) Then
                              propertiesToRemove.Add(prop)
                         End If
                    End If
               Next

               ' Remove collected properties
               For Each prop As JProperty In propertiesToRemove
                    prop.Remove()
               Next
          ElseIf jToken.Type = JTokenType.Array Then
               Dim itemsToRemove As New List(Of JToken)

               For Each item As JToken In jToken.Children()
                    RemoveEmptyOrNullPropertiesFromToken(item)
                    If item.Type = JTokenType.Null OrElse
               (item.Type = JTokenType.String AndAlso String.IsNullOrEmpty(item.ToString())) OrElse
               (item.Type = JTokenType.Object AndAlso Not item.HasValues) OrElse
               (item.Type = JTokenType.Array AndAlso Not item.HasValues) Then
                         itemsToRemove.Add(item)
                    End If
               Next

               ' Remove collected items
               For Each item As JToken In itemsToRemove
                    item.Remove()
               Next
          End If
     End Sub


     Public Function IsJsonObject(jsonString As String) As Boolean
          Try
               ' Parse the string into a JObject
               Dim jsonObject As JObject = JObject.Parse(jsonString)
               Return True
          Catch ex As Exception
               ' If parsing fails, return False
               Return False
          End Try
     End Function




     Public Function UpdateFormattingConditions(jsonString As String, propertyType As String, newFormatType As String, updateAllSubProperties As Boolean) As String
          Try
               Dim jsonObject As JObject = JObject.Parse(jsonString)
               ' Recursive function to update formatting conditions
               UpdateFormattingConditionsInToken(jsonObject, propertyType, newFormatType, updateAllSubProperties)
               Return jsonObject.ToString()
          Catch ex As Exception
               Throw New Exception("An error occurred while processing the JSON.", ex)
          End Try
     End Function

     Private Sub UpdateFormattingConditionsInToken(jToken As JToken, propertyType As String, newFormatType As String, updateAllSubProperties As Boolean)
          If jToken.Type = JTokenType.Object Then
               For Each prop As JProperty In jToken.Children(Of JProperty)()
                    If prop.Name = propertyType Then
                         UpdateFormattingInValue(prop, newFormatType, updateAllSubProperties)
                    End If
                    UpdateFormattingConditionsInToken(prop.Value, propertyType, newFormatType, updateAllSubProperties)
               Next
          ElseIf jToken.Type = JTokenType.Array Then
               For Each item As JToken In jToken.Children()
                    UpdateFormattingConditionsInToken(item, propertyType, newFormatType, updateAllSubProperties)
               Next
          End If
     End Sub

     Private Sub UpdateFormattingInValue(jProperty As JProperty, newFormatType As String, updateAllSubProperties As Boolean)
          If jProperty.Value.Type = JTokenType.String Then
               Dim value As String = jProperty.Value.ToString()
               ' Pattern for regular formatting
               Dim match As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(value, "<!(\w+):(\w)!>")
               ' Pattern for special formatting in translations
               Dim specialMatch As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(value, "@@T!B:(\w+):([^:]+)!T@@")


               If match.Success Then
                    ' Regular formatting
                    Dim originalValue As String = match.Groups(1).Value
                    Dim newValue As String = $"<!{originalValue}:{newFormatType}!>"
                    jProperty.Value = newValue
               ElseIf specialMatch.Success Then
                    ' Special formatting in translations
                    Dim beforeColon As String = specialMatch.Groups(1).Value
                    Dim afterColon As String = specialMatch.Groups(2).Value
                    Dim newValue As String = $"@@T!B:{beforeColon}:{newFormatType}:{afterColon}!T@@"
                    jProperty.Value = newValue
               End If
          ElseIf jProperty.Value.Type = JTokenType.Object AndAlso updateAllSubProperties Then
               For Each subProp As JProperty In jProperty.Value.Children(Of JProperty)()
                    UpdateFormattingInValue(subProp, newFormatType, updateAllSubProperties)
               Next
          End If
     End Sub

     Private Function IsFormattingValue(value As String) As Boolean
          Dim match As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(value, "<!\w+:(\w)!>")
          Dim specialMatch As Text.RegularExpressions.Match = Text.RegularExpressions.Regex.Match(value, "@@T!B:(\w):[^:]+!T@@")
          Return match.Success OrElse specialMatch.Success
     End Function


     ''' <summary>
     ''' Adds or updates "page" and "size" nodes in a JSON object.
     ''' </summary>
     ''' <param name="jsonString">The JSON string.</param>
     ''' <param name="pageValue">The value to set for the "page" node.</param>
     ''' <param name="sizeValue">The value to set for the "size" node.</param>
     ''' <returns>The updated JSON string.</returns>
     Public Function AddOrUpdateNodesInJsonPageAndSize(jsonString As String, pageValue As Integer, sizeValue As Integer) As String
          Try
               ' Parse the JSON string into a JObject
               Dim jsonObject As JObject = JObject.Parse(jsonString)

               ' Add or update the "page" node
               If jsonObject("page") Is Nothing Then
                    jsonObject.Add("page", pageValue)
               Else
                    jsonObject("page") = pageValue
               End If

               ' Add or update the "size" node
               If jsonObject("size") Is Nothing Then
                    jsonObject.Add("size", sizeValue)
               Else
                    jsonObject("size") = sizeValue
               End If

               ' Return the updated JSON string
               Return jsonObject.ToString()
          Catch ex As Exception
               Throw New Exception("An error occurred while adding or updating nodes in the JSON.", ex)
          End Try
     End Function


End Module
