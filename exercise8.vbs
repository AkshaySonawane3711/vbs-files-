'ption explicit
'on error resume next


enterNums = InputBox("enter the number between 1 to 25 :  ")

If enterNums > 25 Or enterNums < 1 Then
	MsgBox "the number you entered is outside the specified range !!!"
	Else
	For enterNums = 1 To enterNums
	MsgBox "the number you entered is within range !!!"
	Next
End If