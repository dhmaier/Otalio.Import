Imports System.Linq

Module bRSQL

     Public Function FormatRSQL(query As String) As String

          ' Define logical connectors that should be kept in the formatted query
          Dim logicalConnectors As String() = {";", "and ", ",", "or ", "AND ", "OR "}

          ' Start with the original query
          Dim formattedQuery As String = query

          formattedQuery = formattedQuery.Replace(";", " and ")
          formattedQuery = formattedQuery.Replace(",", " or ")

          ' Remove any existing line breaks
          formattedQuery = formattedQuery.Replace(vbCrLf, " ").Replace(vbLf, " ")

          ' Replace logical connectors with a new line and the connector
          For Each connector As String In logicalConnectors
               formattedQuery = formattedQuery.Replace(connector, vbCrLf & connector.Trim() & " ")
          Next

          ' Split the formatted query into lines
          Dim lines As String() = formattedQuery.Split(New String() {vbCrLf}, StringSplitOptions.None)

          If lines.Count > 1 Then
               ' Define the column where variables should start (4th column)
               Dim variableStartColumn As Integer = 5
               Dim padding As String = New String(" "c, variableStartColumn)

               ' Adjust each line to align conditions properly
               For i As Integer = 0 To lines.Length - 1
                    If i = 0 Then
                         lines(i) = padding & lines(i).Trim()
                    Else
                         If lines(i).Trim().ToLower.StartsWith("and ") OrElse lines(i).Trim().ToLower.StartsWith("or ") Then
                              lines(i) = " " & lines(i).Substring(0, 3).ToLower() & " " & lines(i).Substring(3).Trim()
                         End If
                    End If
               Next

               ' Join the lines back into a single string
               formattedQuery = String.Join(vbCrLf, lines)
          Else
               formattedQuery = formattedQuery.Trim
          End If

          ' Return the formatted query
          Return formattedQuery

     End Function

End Module
