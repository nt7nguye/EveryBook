VERSION 5.00
Begin VB.Form EveryBook 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EveryBook"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "English/ESL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   0
      Left            =   720
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   7200
      Picture         =   "EveryBook.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   13
      Top             =   720
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   720
      Picture         =   "EveryBook.frx":169D
      ScaleHeight     =   1800
      ScaleWidth      =   1800
      TabIndex        =   12
      Top             =   600
      Width           =   1800
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ADMINISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   11
      Left            =   6840
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Canadian and World Studies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   9
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Business, Technology and Computers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   10
      Left            =   4800
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Health and Physical Education"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   8
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Arts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   7
      Left            =   6840
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Guidance and Career Education"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   6
      Left            =   4800
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Social Sciences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   5
      Left            =   2760
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sciences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   4
      Left            =   720
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mathematics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   3
      Left            =   6840
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spanish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   2
      Left            =   4800
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.CommandButton subBut 
      BackColor       =   &H00C0FFFF&
      Caption         =   "French"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   1
      Left            =   2760
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose your department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EVERYBOOK SYSTEM "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Width           =   6135
   End
End
Attribute VB_Name = "EveryBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
EveryBook.Visible = True
End Sub


Private Sub subBut_Click(Index As Integer)
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\whatfile.txt" For Output As #1
Write #1, Index
Close #1
Load List
Unload EveryBook
End Sub
