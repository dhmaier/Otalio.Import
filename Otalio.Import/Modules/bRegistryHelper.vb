Imports Microsoft.Win32

Module bRegistryHelper

     Private Const RootKey As String = "SOFTWARE\Otalio\Import"


     Public Function ReadRegistry(psKey As String, Optional psDefault As String = "") As String
          Dim value As String = Nothing
          Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(RootKey)
               If key IsNot Nothing Then
                    value = key.GetValue(psKey)?.ToString()
               Else
                    value = psDefault
               End If
          End Using
          Return value
     End Function

     Public Sub WriteRegistry(psKey As String, psValue As String)
          Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(RootKey)
               If key IsNot Nothing Then
                    key.SetValue(psKey, psValue)
               End If
          End Using
     End Sub
End Module
