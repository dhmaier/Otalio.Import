Imports System.Text

Public Class frmError
     Public Sub New(errorMessage As String)
          ' This call is required by the designer.
          InitializeComponent()

          ' Add any initialization after the InitializeComponent() call.
          txtErrorMessage.Text = errorMessage

          ' Change the font of the txtErrorMessage to Lucida Console
          txtErrorMessage.Properties.Appearance.Font = New Font("Lucida Console", txtErrorMessage.Properties.Appearance.Font.Size)


          ' Set focus to the OK button
          Me.ActiveControl = btnOK

     End Sub

     ' Overloaded constructor to accept an exception
     Public Sub New(ex As Exception)
          ' This call is required by the designer.
          InitializeComponent()

          ' Build a detailed error message from the exception
          Dim errorMessage As String = BuildErrorMessage(ex)

          ' Set the detailed error message in the txtErrorMessage control
          txtErrorMessage.Text = errorMessage

          ' Change the font of the txtErrorMessage to Lucida Console
          txtErrorMessage.Properties.Appearance.Font = New Font("Lucida Console", txtErrorMessage.Properties.Appearance.Font.Size)

          ' Set focus to the OK button
          Me.ActiveControl = btnOK
     End Sub

     ' Overloaded constructor to accept an exception
     Public Sub New(errorMessage As String, ex As Exception)
          ' This call is required by the designer.
          InitializeComponent()

          ' Build a detailed error message from the exception
          errorMessage += BuildErrorMessage(ex)

          ' Set the detailed error message in the txtErrorMessage control
          txtErrorMessage.Text = errorMessage

          ' Change the font of the txtErrorMessage to Lucida Console
          txtErrorMessage.Properties.Appearance.Font = New Font("Lucida Console", txtErrorMessage.Properties.Appearance.Font.Size)

          ' Set focus to the OK button
          Me.ActiveControl = btnOK
     End Sub
     Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
          Me.Close()
     End Sub

     Private Sub btnCopyToClipboard_Click(sender As Object, e As EventArgs) Handles btnCopyToClipboard.Click
          Clipboard.SetText(txtErrorMessage.Text)
          MessageBox.Show("Error message copied to clipboard.", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
     End Sub

     ' Method to build a detailed error message from an exception
     Private Function BuildErrorMessage(ex As Exception) As String
          Dim sb As New StringBuilder()

          sb.AppendLine("Message: " & ex.Message)
          sb.AppendLine()
          sb.AppendLine("Source: " & ex.Source)
          sb.AppendLine()
          sb.AppendLine("Stack Trace: " & ex.StackTrace)

          ' Include inner exceptions if any
          Dim innerEx As Exception = ex.InnerException
          While innerEx IsNot Nothing
               sb.AppendLine()
               sb.AppendLine("Inner Exception:")
               sb.AppendLine("Message: " & innerEx.Message)
               sb.AppendLine("Source: " & innerEx.Source)
               sb.AppendLine("Stack Trace: " & innerEx.StackTrace)
               innerEx = innerEx.InnerException
          End While

          sb.AppendLine()


          Return sb.ToString()
     End Function

End Class
