VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAR!!"
   ClientHeight    =   315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   5000
      Left            =   3480
      Top             =   3600
   End
   Begin VB.Timer p2hit 
      Interval        =   1
      Left            =   5280
      Top             =   0
   End
   Begin VB.Timer p1hit 
      Interval        =   1
      Left            =   1680
      Top             =   120
   End
   Begin VB.Timer bul2 
      Interval        =   1
      Left            =   6480
      Top             =   2880
   End
   Begin VB.Timer p2bound 
      Interval        =   1
      Left            =   6840
      Top             =   2160
   End
   Begin VB.Timer p1bound 
      Left            =   120
      Top             =   2280
   End
   Begin VB.Timer bul1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   1440
   End
   Begin VB.Timer amo2 
      Interval        =   1
      Left            =   5640
      Top             =   3600
   End
   Begin VB.Timer amo1 
      Interval        =   1
      Left            =   1320
      Top             =   960
   End
   Begin VB.Timer Tbound 
      Interval        =   1
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer p2trench 
      Interval        =   1
      Left            =   4080
      Top             =   2160
   End
   Begin VB.Timer p1trench 
      Interval        =   1
      Left            =   3000
      Top             =   2160
   End
   Begin VB.Timer Bbound 
      Interval        =   1
      Left            =   3600
      Top             =   4080
   End
   Begin VB.Line Line4 
      X1              =   5760
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   240
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ammo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ammo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Commencing In 5 Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   0
      Width           =   6975
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   7080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape bullet2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape bullet1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   840
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape ammo2 
      BackColor       =   &H00004040&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   6120
      Top             =   3480
      Width           =   135
   End
   Begin VB.Shape ammo1 
      BackColor       =   &H00004040&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1080
      Top             =   360
      Width           =   135
   End
   Begin VB.Image p2 
      Height          =   480
      Left            =   6360
      Picture         =   "frm1.frx":0442
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image p1 
      Height          =   480
      Left            =   360
      Picture         =   "frm1.frx":0529
      Top             =   600
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   3120
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   240
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function collide(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Radius As Single) As Boolean

 If X1 > X2 - Radius And X1 < X2 + Radius And _
    Y1 > Y2 - Radius And Y1 < Y2 + Radius Then
  collide = True
 Else
  collide = False
 End If

End Function

Public Sub p1move(move As Boolean, dir As String)
If move = True Then
If dir$ = "left" Then
p1.Left = p1.Left - 120
ElseIf dir$ = "right" Then
p1.Left = p1.Left + 120
ElseIf dir$ = "up" Then
p1.Top = p1.Top - 120
ElseIf dir$ = "down" Then
p1.Top = p1.Top + 120
End If
ElseIf move = False Then
Exit Sub
End If
End Sub
Public Sub p2move(move As Boolean, dir As String)
If move = True Then
If dir$ = "left" Then
p2.Left = p2.Left - 120
ElseIf dir$ = "right" Then
p2.Left = p2.Left + 120
ElseIf dir$ = "up" Then
p2.Top = p2.Top - 120
ElseIf dir$ = "down" Then
p2.Top = p2.Top + 120
End If
ElseIf move = False Then
Exit Sub
End If
End Sub
Private Sub amo1_Timer()
If collide(bullet2.Left, bullet2.Top, ammo1.Left, ammo1.Top, ammo1.Height + 100) Then
bullet2.Visible = False
bullet2.Left = 8000
bul2.Enabled = False
Else
End If
If Label2.caption < 0 Then
Label2.caption = 0
End If
If p1.Top = 360 And p1.Left <= 480 Then
Label2.caption = 10
Else
End If
End Sub

Private Sub amo2_Timer()
If p2.Left > 5520 And p2.Top >= 3000 Then
p2.Left = 6360
Else
End If
If collide(bullet1.Left, bullet1.Top, ammo2.Left, ammo2.Top, ammo2.Height - 300) Then
bullet1.Visible = False
bullet1.Left = 8000
bul1.Enabled = False
Else
End If
If Label4.caption < 0 Then
Label4.caption = 0
End If
If p2.Top = 3720 And p2.Left >= 6240 Then
Label4.caption = 10
Else
End If
End Sub

Private Sub Bbound_Timer()
If p2.Top = 3840 Then
p2.Top = 3720
Else
End If
If p1.Top = 3840 Then
p1.Top = 3720
Else
End If
End Sub

Private Sub bul1_Timer()
bullet1.Left = bullet1.Left + 120
End Sub

Private Sub bul2_Timer()
bullet2.Left = bullet2.Left - 120
End Sub

Private Sub col1_Timer()

End Sub

Private Sub Command1_Click()
If Text1.text = "5a1vc" Then
Command1.caption = "Player ready"
Command1.Enabled = False
Else
MsgBox ("Wrong Code Player 1")
End If
If Command1.caption = "Player ready" And Command2.caption = "Player ready" Then
label.caption = "Game Commencing in 5 seconds"
Timer.Enabled = True
Else
End If
End Sub

Private Sub Command2_Click()
If Text2.text = "1fc3s" Then
Command2.caption = "Player ready"
Command2.Enabled = False
Else
MsgBox ("Wrong Code Player 2")
End If
If Command1.caption = "Player ready" And Command2.caption = "Player ready" Then
label.caption = "Game Commencing in 5 seconds"
Timer.Enabled = True
Else
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 100 Then
p1move True, "right"
End If
If KeyAscii = 120 Then
p1move True, "down"
End If
If KeyAscii = 97 Then
p1move True, "left"
End If
If KeyAscii = 119 Then
p1move True, "up"
End If
If KeyAscii = 32 Then
If Label2.caption <= 0 Then
Exit Sub
End If
If collide(p1.Left, p1.Top, ammo1.Left, ammo1.Top, ammo1.Height) Then
Exit Sub
Else
End If
Label2.caption = Label2.caption - 1
bullet1.Visible = True
bullet1.Top = p1.Top
bullet1.Left = p1.Left + 600
bul1.Enabled = True
End If
If KeyAscii = 56 Then
p2move True, "up"
End If
If KeyAscii = 50 Then
p2move True, "down"
End If
If KeyAscii = 52 Then
p2move True, "left"
End If
If KeyAscii = 54 Then
p2move True, "right"
End If
If KeyAscii = 48 Then

If Label4.caption <= 0 Then
Exit Sub
End If
If collide(p2.Left, p2.Top, ammo2.Left, ammo2.Top, ammo2.Height - 300) Then
Exit Sub
Else
End If
Label4.caption = Label4.caption - 1
bullet2.Visible = True
bullet2.Top = p2.Top + 350
bullet2.Left = p2.Left
bul2.Enabled = True

End If

End Sub

Private Sub p1hit_Timer()
If collide(bullet2.Left, bullet2.Top, p1.Left, p1.Top, p1.Height + 100) Then
bullet2.Left = 8000
bul2.Enabled = False
bullet2.Visible = False
Label6.caption = Label6.caption - 10
Else
End If
If Label6.caption = "0" Then
MsgBox ("Player 1 Is Dead")
MsgBox ("Returning to main screen")
Load Form2
Form2.Show
Unload Me
Else
End If
End Sub

Private Sub p1trench_Timer()
If p1.Left = 2400 Then
p1.Left = 2280
Else
End If
End Sub

Private Sub p2bound_Timer()
If p2.Left = 6600 Then
p2.Left = 6480
Else
End If
End Sub

Private Sub p2hit_Timer()
If collide(bullet1.Left, bullet1.Top, p2.Left, p2.Top, p2.Height - 300) Then
bullet1.Left = 8000
bul1.Enabled = False
bullet1.Visible = False
Label8.caption = Label8.caption - 10
Else
End If
If Label8.caption = "0" Then
MsgBox ("Player 2 Is Dead")
MsgBox ("Returning to main screen")
Load Form2
Form2.Show
Unload Me
Else
End If

End Sub

Private Sub p2trench_Timer()
If p2.Left = 4320 Then
p2.Left = 4440
Else
End If
End Sub

Private Sub Tbound_Timer()
If p2.Top = 240 Then
p2.Top = 360
Else
End If
If p1.Top = 240 Then
p1.Top = 360
Else
End If
End Sub

Private Sub Timer_Timer()
label.Visible = False
Timer.Enabled = False
Form1.Height = 5070
End Sub
