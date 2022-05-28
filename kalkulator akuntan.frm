VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "kalkulator "
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "exit"
      Height          =   615
      Left            =   5520
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clear"
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "hasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "nilai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "nilai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3 = Val(Text1) + Val(Text2)
End Sub

Private Sub Command2_Click()
Text3 = Val(Text1) - Val(Text2)
End Sub

Private Sub Command3_Click()
Text3 = Val(Text1) * Val(Text2)
End Sub

Private Sub Command4_Click()
Text3 = Val(Text1) / Val(Text2)
End Sub



Private Sub Command6_Click()
End
End Sub


