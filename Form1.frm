VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Your username or password is incorrect"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
