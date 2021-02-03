option explicit 
on error resume next

Dim arryName(9)

 arryName(0) = "Stanton Lugo "  
 arryName(1) = "Corrinne Tunison" 
 arryName(2) = " Wilber Pyne"
 arryName(3) = "Ava Enos"
 arryName(4) = "Marina Forst"
 arryName(5) = "Yetta Wile "
 arryName(6) = "Kum Gorrell " 
 arryName(7) = "Omar Gonser "
 arryName(8) = "Blanch Kluth " 
 arryName(9) = "Joanne Pfaff"

MsgBox arryName(0) & "met " & arryName(5) & "in Washington DC.", 0, "message : 1"

MsgBox arryName(4) & "celebrated her birthday with " & arryName(1), 0, "message : 2"