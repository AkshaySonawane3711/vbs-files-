'ption explicit 
'on error resume next

' Thompson - Presisent
' Rooter - Sr. Vice President
' Cooper - Vice President
' Parker - Manager
lastname = InputBox("enter your name: ")
companyName = "ABC_inc"

select case lastname
case"Thompson"
	MsgBox " Thompson is president of"& companyName
case"Rooter"
	MsgBox " Rooter is Sr. Vice President of" & companyName
case"Cooper"
	MsgBox " Cooper is Vice President of" & companyName
case"Parker"
	MsgBox " Parker is Manager of" & companyName
case else 
	MsgBox "hello" &lastname & "you are not part of management team"

end select
