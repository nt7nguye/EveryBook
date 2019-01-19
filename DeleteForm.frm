VERSION 5.00
Begin VB.Form DeleteForm 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No"
      Height          =   615
      Index           =   1
      Left            =   2355
      TabIndex        =   0
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Are you sure you want to delete this book?"
      Height          =   495
      Left            =   1350
      TabIndex        =   2
      Top             =   600
      Width           =   1995
   End
End
Attribute VB_Name = "DeleteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Open "G:\ICS4\Group2\delete.txt" For Output As #2
Write #2, Index
Close #2
List.Enabled = True
Unload delForm
End Sub
Private Sub Form_Load()
List.Enabled = False
End Sub

