Imports System.Xml.Serialization
Imports Newtonsoft.Json.Linq
Imports System.IO
<
  XmlInclude(GetType(clsImportColum)),
  Serializable()
  >
Public Class clsImportColum

  Public ReadOnly Property No As Integer
    Get
      Try


        If String.IsNullOrEmpty(ColumnID) = False Then

          Dim sString As String = ColumnID

          'check if column id as additonal properties attached
          If ColumnID.Contains(":") Then
            Dim oList As String() = sString.Split(":")
            If oList IsNot Nothing AndAlso oList(0) IsNot Nothing Then
              sString = oList(0)
            End If
          End If


          sString = sString.ToUpper()
          Dim sum As Integer = 0

          For i As Integer = 0 To sString.Length - 1
            sum *= 26
            Dim charA As Double = Char.ConvertToUtf32("A", 0) - 64
            Dim charColLetter As Double = Char.ConvertToUtf32(sString(i), 0) - 64
            sum += (charColLetter - charA) + 1
          Next
          Return sum

        Else
          Return 0
        End If
      Catch ex As Exception
        Return 0
      End Try
    End Get
  End Property
  Public Property ColumnName As String = ""
  Public Property Type As JTokenType
  Public Property Name As String = ""
  Public Property Parent As String = ""
  Public Property Formatted As String = ""
  Public Property ColumnID As String = ""
  Public Property VariableName As String = ""
  Public Property ChildNode As String = ""

End Class
