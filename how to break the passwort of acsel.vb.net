Sub PasswordBreaker()
'Breaks worksheet password protection.
Dim T As Integer, A As Integer, R As Integer
Dim E As Integer, K As Integer
Dim T1 As Integer, A1 As Integer, R1 As Integer
Dim E1 As Integer, K1 As Integer, S As Integer, D As Integer
On Error Resume Next
For T = 99 To 100: For A = 99 To 100: For R = 99 To 100
For E = 99 To 100: For K = 99 To 100: For T1 = 99 To 100
For A1 = 99 To 100: For R1 = 99 To 100: For E1 = 99 To 100
For K1 = 99 To 100: For S = 99 To 100: For D = 1 To 100
ActiveSheet.Unprotect Chr(T) & Chr(A) & Chr(R) & _
Chr(E) & Chr(K) & Chr(T1) & Chr(A1) & Chr(R1) & _
Chr(E1) & Chr(K1) & Chr(S) & Chr(D)
If ActiveSheet.ProtectContents = False Then
MsgBox "done"
Exit Sub
End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub

------------------------------------------------------------------------------------------------------

'(:) Means (to)
