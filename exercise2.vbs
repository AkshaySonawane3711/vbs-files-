option explicit
on error resume next

const name = "adam"
const annualSalary = 60000
const taxrates =  0.2
const retirementcontribution =  0.05 
const healthplan = 200

dim grossMonthlySalary 
dim monthlyTax
dim MonthlyretirementContribution
dim monthlyNetSalary

GrossMonthlySalary = annualSalary / 12
monthlyTax = grossMonthlySalary  * taxrates
MonthlyretirementContribution = grossMonthlySalary * retirementcontribution
monthlyNetSalary = grossSalary - monthlyTax - MonthlyretirementContribution - healthplan

MsgBox name & " "& "receives" & monthlyNetSalary & "as monthly salary after all deductions!!!",49," ABC_industry!!!!"