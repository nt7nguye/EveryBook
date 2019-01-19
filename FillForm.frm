VERSION 5.00
Begin VB.Form Fillform 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add book"
      Height          =   615
      Left            =   1440
      TabIndex        =   18
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Condition"
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   17
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Edition"
      Height          =   615
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "ISBN"
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Location"
      Height          =   615
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Teacher"
      Height          =   615
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Number of copies"
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Publication Year"
      Height          =   615
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Publisher/Author"
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Title"
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   8
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   7
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   6
      Left            =   2400
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   5
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Fillform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Fillform
list.Enabled = True
End Sub

Private Sub Form_Load()
Fillform.Visible = True
End Sub
