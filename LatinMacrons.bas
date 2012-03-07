' Created by James Hodgkinson in 2012
' latest version probably available here: https://github.com/yaleman/Latin-Macron-Macros-for-Word-2003
Sub smallO()
WriteMacron 333
End Sub

Sub bigO()
WriteMacron 332
End Sub

Sub smallA()
WriteMacron 257
End Sub

Sub bigA()
WriteMacron 256
End Sub

Sub bigE()
WriteMacron 274
End Sub

Sub smallE()
WriteMacron 275
End Sub

Sub bigI()
WriteMacron 298
End Sub

Sub smallI()
WriteMacron 299
End Sub


Sub bigU()
WriteMacron 362
End Sub

Sub smallU()
WriteMacron 363
End Sub

Sub bigY()
WriteMacron 562
End Sub

Sub smallY()
WriteMacron 563
End Sub

Sub bigAENoTop()
WriteMacron 198
End Sub

Sub smallAENoTop()
WriteMacron 230
End Sub

Sub bigAE()
WriteMacron 482
End Sub

Sub smallAE()
WriteMacron 483
End Sub


' A helper function 
Sub WriteMacron(n)
Dim MyText As String
MacronString = ChrW(n)
Selection.TypeText (MacronString)
End Sub
