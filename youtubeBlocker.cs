'
'
Dim MyURL As String
'
'
Private Sub Command1_Click()
MyURL = "https://www.youtube.com"
WebBrowser1.Navigate2 MyURL
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If Instr(URL, MyURL) > 0 Then
 Cancel = True ' Do not allow
End If
End Sub
