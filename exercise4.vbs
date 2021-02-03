option explicit 
'on error resume next 

const state_Name = "exercise 4"

dim name
dim age 
dim city

name = InputBox("enter your name")
age = InputBox("enter your age")
city = InputBox("enter your city")

MsgBox name & " is " & age & " years old and lives in " & city,1, state_Name 