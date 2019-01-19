VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form List 
   Caption         =   "EveryBook"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18390
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   18390
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "List.frx":0000
      Height          =   8295
      Left            =   360
      TabIndex        =   29
      Top             =   1440
      Width           =   22200
      _ExtentX        =   39158
      _ExtentY        =   14631
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4105
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   19080
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   10320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   14760
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   11160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   14760
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   10320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   11040
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   11160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   11040
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   10320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   8400
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   11160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   8400
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   11160
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   17280
      TabIndex        =   15
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Condition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   12960
      TabIndex        =   14
      Top             =   11160
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ISBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   12960
      TabIndex        =   13
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   9720
      TabIndex        =   12
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Teacher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   9720
      TabIndex        =   11
      Top             =   10320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   6240
      TabIndex        =   10
      Top             =   11160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Publication(Year)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Publisher/Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   11160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Delete Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15360
      Top             =   480
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6960
      TabIndex        =   3
      Text            =   "Title"
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton all 
      BackColor       =   &H00C0FFC0&
      Caption         =   "All books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton gosearch 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   495
      Left            =   15960
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Tan Nguyen\Desktop\EveryBook\everything.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Tan Nguyen\Desktop\EveryBook\everything.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Sheet1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox srch 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   0
      Text            =   "Search..."
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      TabIndex        =   4
      Top             =   120
      Width           =   8535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ADD BOOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   9840
      Width           =   22215
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   18840
         TabIndex        =   28
         Text            =   "English"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   17040
         TabIndex        =   27
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton add 
         Caption         =   "Add book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   20760
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2400
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim confirm As Integer
Dim userversion As Integer
Dim systemversion As Integer
Dim subject As String
Dim mode As Integer
'mode = 0 for users seeing 1 subject only
'mode = 1 for administration

Private Sub all_Click()
If mode = 0 Then
adogrid.RecordSource = "Select * from Sheet1 where Subject ='" + subject + "'"
Else
adogrid.RecordSource = "Select * from Sheet1"
End If
adogrid.Refresh
adogrid.Caption = adogrid.RecordSource
Call formattable
End Sub

Private Sub add_click()
adogrid.Refresh
adogrid.Recordset.AddNew

With adogrid.Recordset
.Fields("Title") = Text1(0).Text
.Fields("Publisher/Author") = Text1(1).Text
.Fields("Publication(Year)") = Text1(2).Text
.Fields("Copies") = Text1(3).Text
.Fields("Teacher") = Text1(4).Text
.Fields("Location") = Text1(5).Text
.Fields("ISBN") = Text1(6).Text
.Fields("Condition") = Text1(7).Text
.Fields("Edition") = Text1(8).Text
End With

If mode = 0 Then
adogrid.Recordset.Fields("Subject") = subject
Else
adogrid.Recordset.Fields("Subject") = Combo2.Text
End If

For X = 0 To 8
Text1(X).Text = ""
Next

Add.Enabled = False

Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Input As #4
Input #4, systemversion
Close #4
If systemversion > 100 Then
userversion = 1
Else
userversion = userversion + 1
End If
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Output As #5
Write #5, userversion
Close #5

Call formattable
adogrid.Recordset.Update
End Sub

Private Sub Command1_Click()
Unload List
Load EveryBook
End Sub

Private Sub Command2_Click()
confirm = MsgBox("Do you want to delete the Record ?", vbYesNo + vbExclamation, "Warning Message")
If confirm = vbYes Then
adogrid.Recordset.Delete
MsgBox "Record Deleted Successfully", vbInformation, "Delete Record Confirmation"
'change version for other users to update their records
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Input As #2
Input #2, systemversion
Close #2
If systemversion > 100 Then
userversion = 1
Else
userversion = userversion + 1
End If
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Output As #3
Write #3, userversion
Close #3

Else
MsgBox "Record Not Deleted", vbInformation, "Record Not Deleted"
End If

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex _
As Integer)
Dim sortField As String
Dim sortString As String

sortField = DataGrid1.Columns(ColIndex).Caption
If InStr(adogrid.Recordset.Sort, "Asc") Then
    sortString = sortField & " Desc"
Else
    sortString = sortField & " Asc"
End If
adogrid.Recordset.Sort = sortString
End Sub

Private Sub Form_Load()
'1. change final dimensions
'2. change alignment
List.Visible = True
srch.Locked = True
Combo1.AddItem "Title"
Combo1.AddItem "Publisher/Author"
Combo1.AddItem "Publication(Year)"
Combo1.AddItem "Teacher"
Combo1.AddItem "Location"
Combo1.AddItem "ISBN"

For X = 0 To 8
Text1(X).Text = ""
Next X
Add.Enabled = False
gosearch.Enabled = False

Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Input As #20
Input #20, systemversion
Close #20
userversion = systemversion

Dim Index As Integer
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\whatfile.txt" For Input As #10
Input #10, Index
Close #10

Command3(9).Visible = False
Combo2.Visible = False

mode = 0
Select Case Index
Case 0
subject = "English"
Case 1
subject = "French"
Case 2
subject = "Spanish"
Case 3
subject = "Math"
Case 4
subject = "Sciences"
Case 5
subject = "Social Sciences"
Case 6
subject = "Guidance"
Case 7
subject = "Arts"
Case 8
subject = "Health"
Case 9
subject = "History"
Case 10
subject = "Business"
Case 11
mode = 1
Combo2.AddItem "English"
Combo2.AddItem "French"
Combo2.AddItem "Spanish"
Combo2.AddItem "Math"
Combo2.AddItem "Sciences"
Combo2.AddItem "Social Sciences"
Combo2.AddItem "Guidance"
Combo2.AddItem "Arts"
Combo2.AddItem "Health"
Combo2.AddItem "History"
Combo2.AddItem "Business"
Combo2.Visible = True
Command3(9).Visible = True
End Select

Dim connection As String

connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Tan Nguyen\Desktop\Everybook\Everything.mdb;Persist Security Info=False"
adogrid.ConnectionString = connection
If mode = 0 Then
adogrid.RecordSource = "Select * from Sheet1 where Subject ='" + subject + "'"
End If
adogrid.Refresh
Call formattable
End Sub
 
Private Sub formattable()
If mode = 1 Then
DataGrid1.Columns("Subject").Visible = True
With DataGrid1
    .Columns("Title").Width = 5100
    .Columns("Publisher/Author").Width = 5100
    .Columns("Publication(Year)").Width = 1900
    .Columns("Teacher").Width = 1600
    .Columns("Copies").Width = 900
    .Columns("Location").Width = 1200
    .Columns("ISBN").Width = 2200
    .Columns("Condition").Width = 1300
    .Columns("Edition").Width = 1000
    .Columns("Subject").Width = 1300
End With
Else
DataGrid1.Columns("Subject").Visible = False
With DataGrid1
    .Columns("Title").Width = 5750
    .Columns("Publisher/Author").Width = 5750
    .Columns("Publication(Year)").Width = 1900
    .Columns("Teacher").Width = 1600
    .Columns("Copies").Width = 900
    .Columns("Location").Width = 1200
    .Columns("ISBN").Width = 2200
    .Columns("Condition").Width = 1300
    .Columns("Edition").Width = 1000
End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
adogrid.ConnectionString = ""
End Sub

Private Sub gosearch_Click()
If Combo1.Text = "Title" Then
adogrid.RecordSource = "Select * from Sheet1 where Title ='" + srch.Text + "' and Subject ='" + subject + "'"
'or Publisher/Author ='" + srch.Text + "' or Publication(Year) ='" + srch.Text + "' or Teacher ='" + srch.Text + "' or Location ='" + srch.Text + "' or ISBN ='" + srch.Text + "'"
ElseIf Combo1.Text = "Publisher/Author" Then
adogrid.RecordSource = "Select * from Sheet1 where Publisher/Author ='" + srch.Text + "' and Subject ='" + subject + "'"
ElseIf Combo1.Text = "Publication(Year)" Then
adogrid.RecordSource = "Select * from Sheet1 where Publication(Year) ='" + srch.Text + "' and Subject ='" + subject + "'"
ElseIf Combo1.Text = "Teacher" Then
adogrid.RecordSource = "Select * from Sheet1 where Teacher ='" + srch.Text + "' and Subject ='" + subject + "'"
ElseIf Combo1.Text = "Location" Then
adogrid.RecordSource = "Select * from Sheet1 where Location ='" + srch.Text + "' and Subject ='" + subject + "'"
ElseIf Combo1.Text = "ISBN" Then
adogrid.RecordSource = "Select * from Sheet1 where ISBN ='" + srch.Text + "' and Subject ='" + subject + "'"
End If
adogrid.Refresh
If adogrid.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Message"
Else
adogrid.Caption = adogrid.RecordSource
End If
gosearch.Enabled = False
srch.Locked = True
srch.Text = "Search..."
gosearch.Enabled = False

Call formattable

End Sub

Private Sub srch_Change()
gosearch.Enabled = True
End Sub

Private Sub srch_Click()
srch.Locked = False
If srch.Text = "Search..." Then
srch.Text = ""
End If
End Sub

Private Sub srch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call gosearch_Click
End Sub

Private Sub Text1_Change(Index As Integer)
Add.Enabled = True
End Sub

Private Sub Timer1_Timer()
'records only update when there is a new version -> efficient
'a new version is when users make change to the records: add or remove book
Open "C:\Users\Tan Nguyen\Desktop\EveryBook\version.txt" For Input As #1
Input #1, systemversion
Close #1
If userversion <> systemversion Then
adogrid.Refresh
userversion = systemversion

Call formattable

End If
End Sub
