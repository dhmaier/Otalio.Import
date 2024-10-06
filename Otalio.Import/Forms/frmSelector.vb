Public Class frmSelector

     Private Sub butLoadFromJson_Click(sender As Object, e As EventArgs)
          Me.UcSelectors1.LoadFromJsonfile()
     End Sub

     Private Sub butExportToJson_Click(sender As Object, e As EventArgs)
          Me.UcSelectors1.ExportToJsonFile()
     End Sub
End Class