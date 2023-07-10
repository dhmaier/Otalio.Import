Imports System.ComponentModel

Public Class frmSelectorQuery

     Public Property ListOfSelectors As List(Of clsSelectors)
     Public Property QueryText As String

     Private Sub frmSelectorQuery_Load(sender As Object, e As EventArgs) Handles Me.Load


     End Sub

     Private Sub frmSelectorQuery_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
          Try
               _ucSelectorQuery.ucSelectorQuery_Validating(Nothing, Nothing)
               Me.QueryText = _ucSelectorQuery._Query
          Catch ex As Exception

          End Try

     End Sub

     Public Sub LoadQueries()

          Me._ucSelectorQuery._Selectors = ListOfSelectors
          Me._ucSelectorQuery._Query = QueryText

     End Sub
End Class