VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Select The Game of WAR!"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   Icon            =   "frm4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   6000
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to play 2 Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   3015
      Left            =   0
      Picture         =   "frm4.frx":0442
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to play against Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "frm4.frx":7E28C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Load Form3
Form3.Show
Unload Me
End Sub

Private Sub Image2_Click()
Load Form1
Form1.Show
Unload Me
End Sub
