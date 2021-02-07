option explicit 
'on error resume next 

dim johnAge
dim kerryAge

johnAge = InputBox("enter john's age: ")
kerryAge = InputBox("enter kerry's age: ")

if johnAge > kerryAge then
	MsgBox "John is older than Kerry"
elseif johnAge < kerryAge then
	MsgBox "Kerry is older than John"
end if