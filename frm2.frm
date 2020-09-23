VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   Caption         =   "WAR!!!"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   Icon            =   "frm2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4650
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2 player: Up (8) down (2) left (4) right (6) (All on Key Pad)"
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1 player: Up (w) Down (x) Left (a) right (d)"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO PLAY!"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Line Line6 
      X1              =   2400
      X2              =   1920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      X1              =   5160
      X2              =   5520
      Y1              =   4250
      Y2              =   4250
   End
   Begin VB.Image Image1 
      Height          =   2475
      Left            =   1680
      Picture         =   "frm2.frx":0442
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4185
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   550
      Top             =   3150
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   1320
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm2.frx":7E28C
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2400
      Y1              =   1320
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   2400
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2280
      Y1              =   960
      Y2              =   1080
   End
   Begin VB.Shape ammo1 
      BackColor       =   &H00004040&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2760
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   2040
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm2.frx":7E404
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Load Form4
Form4.Show
Unload Me
End Sub
