VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Custom Data Types Tutorial"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrderPizza 
      Caption         =   "&Order Pizza"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame frmCrust 
      Caption         =   "Crust"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton optStuffed 
         Caption         =   "Stuffed"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optReg 
         Caption         =   "Regular"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optThin 
         Caption         =   "Thin"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame frmTopping 
      Caption         =   "Topping"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
      Begin VB.ComboBox cmbTopping 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0019
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "Pepperoni"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frmSize 
      Caption         =   "Pizza Size"
      Height          =   975
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      Begin VB.OptionButton optLarge 
         Caption         =   "Large"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMed 
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optSmall 
         Caption         =   "Small"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox chkExCheese 
      Caption         =   "Extra Cheese"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright© 2003-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit

Dim sCrust As String ' Used to store the actual Pizza Crust string
Dim sSize  As String ' Used to store the actual Pizza Size string

' Declare tPIZZA as our CDT, PIZZA
Dim tPIZZA As PIZZA


Private Sub Form_Load()

     ' Set the default values
     tPIZZA.Crust = 1  ' Regular crust
     tPIZZA.Size = 2   ' Medium sized pizza
     
End Sub


Private Sub optSmall_Click()

     ' Set the value of pizza size to small
     tPIZZA.Size = 1

End Sub

Private Sub optMed_Click()

     ' Set the value of pizza size to medium
     tPIZZA.Size = 2
     
End Sub

Private Sub optLarge_Click()

     ' Set the value of pizza size to large
     tPIZZA.Size = 3
     
End Sub


Private Sub optReg_Click()
     
     ' Set the pizza crust to regular
     tPIZZA.Crust = 1
     
End Sub

Private Sub optStuffed_Click()

     ' Set the pizza crust to stuffed
     tPIZZA.Crust = 2
     
End Sub

Private Sub optThin_Click()

     ' Set the pizza crust to thin
     tPIZZA.Crust = 3

End Sub


Private Sub cmdOrderPizza_Click()

     ' Set all the values into our variable tPIZZA
     tPIZZA.ExtraCheese = chkExCheese.Value
     tPIZZA.Topping = cmbTopping.Text
     
     ' Get the string value of our pizza crust
     Select Case (tPIZZA.Crust)
          
          Case 1:
               sCrust = "Regular"
          Case 2:
               sCrust = "Stuffed"
          Case 3:
               sCrust = "Thin"
               
     End Select
     
     ' Get the string value of our pizza size
     Select Case (tPIZZA.Size)
          
          Case 1:
               sSize = "Small"
          Case 2:
               sSize = "Medium"
          Case 3:
               sSize = "Large"
               
     End Select
               
     
     ' Display the data in a Message Box
     MsgBox "Topping:  " & tPIZZA.Topping & vbNewLine _
          & "Crust:  " & sCrust & vbNewLine _
          & "Size:  " & sSize & vbNewLine _
          & "Extra Cheese:  " & tPIZZA.ExtraCheese & vbNewLine _
          , 0, "Your pizza!"
     
End Sub


Private Sub cmdExit_Click()

     ' End the program
     Unload Me

End Sub
