'option explicit 
'on error resume next

enterNums = InputBox("enter number between 1 to 100:  ")

If enterNums > 100 Or enterNums < 1 then
	MsgBox "the number you entered is outside the specified range!!!!"
	'Else
	'MsgBox "the number you entered is within the specified range!!!!"
	
Else	
	for x = 1 TO enterNums 
		reult = x mod 10
		if result = 0 then
		MsgBox x,1, "result"
		end If
		
	next
	
End If
	
