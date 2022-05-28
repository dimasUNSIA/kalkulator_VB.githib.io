VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton Command18 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4320
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3360
      TabIndex        =   15
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   11
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op, fn As Integer

Private Sub Command1_Click()
Text1.Text = Text1.Text & 1
End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text & "."
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text & 0
End Sub

Private Sub Command12_Click()
If op = 1 Then
Text1.Text = Val(fn) + Val(Text1.Text)
ElseIf op = 2 Then
Text1.Text = Val(fn) - Val(Text1.Text)
ElseIf op = 3 Then
Text1.Text = Val(fn) * Val(Text1.Text)
ElseIf op = 4 Then
Text1.Text = Val(fn) / Val(Text1.Text)
End If
End Sub

Private Sub Command13_Click()
op = 3
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command14_Click()
op = 2
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command15_Click()
op = 4
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command16_Click()
op = 1
fn = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command17_Click()
Unload Me
End Sub

Private Sub Command18_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & 2
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & 3
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & 4
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text & 5
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & 6
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text & 7
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & 8
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & 9
End Sub
