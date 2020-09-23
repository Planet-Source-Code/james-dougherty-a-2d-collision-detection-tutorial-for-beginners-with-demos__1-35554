VERSION 5.00
Begin VB.Form frmBF 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Brute Force"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1830
   Icon            =   "frmBF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic2DColl 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   960
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox pic2DColl 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Collision!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "frmBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Check_Collision_BF(Sprite1 As PictureBox, Sprite2 As PictureBox) As Boolean
 
 'Simply check the bounderies of the 2 Picture Box's
 If Sprite1.Left > Sprite2.Left + Sprite2.Width Or _
    Sprite1.Left + Sprite1.Width < Sprite2.Left Or _
    Sprite1.Top > Sprite2.Top + Sprite2.Height Or _
    Sprite1.Top + Sprite1.Height < Sprite2.Top Then
  Check_Collision_BF = False
 Else
  Check_Collision_BF = True
 End If
 
End Function

Private Sub Form_Load()
 
 'Might as well check...
 If Check_Collision_BF(pic2DColl(0), pic2DColl(1)) = True Then
  Label1.Caption = "Collision!"
 Else
  Label1.Caption = "Not Colliding"
 End If
 
End Sub

Private Sub pic2DColl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
 'Move the Picture Box
 Select Case KeyCode
  Case vbKeyLeft
   pic2DColl(0).Left = pic2DColl(0).Left - 1
  Case vbKeyRight
   pic2DColl(0).Left = pic2DColl(0).Left + 1
 End Select
 
 'Check for a collision
 If Check_Collision_BF(pic2DColl(0), pic2DColl(1)) = True Then
  Label1.Caption = "Collision!"
 Else
  Label1.Caption = "Not Colliding"
 End If
 
End Sub
