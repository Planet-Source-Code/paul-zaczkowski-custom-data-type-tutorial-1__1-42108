Attribute VB_Name = "modTypes"
' Copyright© 2003-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit


' It is always better to create your types in a module, rather a form
' becuase that way, all modules will be able to acces it.

' Declare the custom data type, PIZZA
Public Type PIZZA
     
     Topping        As String
     Crust          As Byte
     ExtraCheese    As Boolean
     Size           As Byte
     ' Notice the different date types

End Type


