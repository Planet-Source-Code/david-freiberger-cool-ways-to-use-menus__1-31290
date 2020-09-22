VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CmdButtonMenu 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CmdButtonMenu 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Floating Menu     Right-Click on me"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Menu Sample"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu regmnu 
      Caption         =   "Regular Menu"
      Begin VB.Menu mhello32 
         Caption         =   "Hello!"
      End
      Begin VB.Menu mnu10 
         Caption         =   "Menu"
      End
      Begin VB.Menu mnu20 
         Caption         =   "Menu2"
      End
   End
   Begin VB.Menu mtest 
      Caption         =   "test"
      Visible         =   0   'False
      Begin VB.Menu mhello 
         Caption         =   "Hello!        "
      End
      Begin VB.Menu mnu1 
         Caption         =   "Menu"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Menu2"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Very simple code, but sometimes iv been looking for this same thing and couldnt find it!
Private Sub Command1_Click()
'this takes a popup menu from your menu list and displays it where you specify
'in this case we want it to be below a button, so we add the height of the button - 10 (this button is 375 pixels in height)
'now when you click that button it will appear right below the button, if you want the menu as wide as the button then make sure that the_
'caption on one of your sub-menu items in that list has enough spaces to make it the width of the button
Form1.PopupMenu mtest, , Command1.Left, Command1.Top + 365
End Sub
Private Sub Command2_Click()
'basically the same as the above, only we subtract about a little less than half the width of the button
'to make it popup on top
Form1.PopupMenu mtest, , Command1.Left, Command1.Top - 140
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is a floating menu, it checks to see if you clicked your right mouse button and then displays a menu where your mouse was
If Button = 2 Then Form1.PopupMenu mtest
End Sub
