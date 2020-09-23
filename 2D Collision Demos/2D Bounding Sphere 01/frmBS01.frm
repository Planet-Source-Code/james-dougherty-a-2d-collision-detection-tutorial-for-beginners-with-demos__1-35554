VERSION 5.00
Begin VB.Form frmBS01 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "2D Bounding Sphere Method 1"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7500
   Icon            =   "frmBS01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape shpColl 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Height          =   825
      Index           =   1
      Left            =   6000
      Shape           =   1  'Square
      Top             =   1320
      Width           =   825
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
      Left            =   2978
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
   Begin VB.Shape shpColl 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   825
      Index           =   0
      Left            =   600
      Top             =   1320
      Width           =   825
   End
End
Attribute VB_Name = "frmBS01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Heres a 3D Example

'Public Function Check_3D_Collision(Collider_Position As D3DVECTOR, Object_Position As D3DVECTOR, Object_Radius As Single) As Boolean
'  If Collider_Position.X > Object_Position.X - Object_Radius _
' And Collider_Position.X < Object_Position.X + Object_Radius _
' And Collider_Position.Y > Object_Position.Y - Object_Radius _
' And Collider_Position.Y < Object_Position.Y + Object_Radius _
' And Collider_Position.Z > Object_Position.Z - Object_Radius _
' And Collider_Position.Z < Object_Position.Z + Object_Radius Then
'  Check_3D_Collision = True
' Else
'  Check_3D_Collision = False
' End If
'End Function

Public Function SphereCollision_2D1(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Radius As Single) As Boolean

 'Check for the collision
 
 'Radius would be the Width or Height of your sprite(or whatever you're using)
 If X1 > X2 - Radius And X1 < X2 + Radius And _
    Y1 > Y2 - Radius And Y1 < Y2 + Radius Then
  SphereCollision_2D1 = True
 Else
  SphereCollision_2D1 = False
 End If

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 'Move the Shape
 Select Case KeyCode
  Case vbKeyUp
   shpColl(0).Top = shpColl(0).Top - 3
  Case vbKeyDown
   shpColl(0).Top = shpColl(0).Top + 3
  Case vbKeyLeft
   shpColl(0).Left = shpColl(0).Left - 3
  Case vbKeyRight
   shpColl(0).Left = shpColl(0).Left + 3
 End Select
 
 'Check for a collision
 If SphereCollision_2D1(shpColl(0).Left, shpColl(0).Top, _
                        shpColl(1).Left, shpColl(1).Top, 55) = True Then
  Label1.Caption = "Collision!"
 Else
  Label1.Caption = "Not Colliding"
 End If
 
End Sub

Private Sub Form_Load()

 'Initial Check
 If SphereCollision_2D1(shpColl(0).Left, shpColl(0).Top, _
                        shpColl(1).Left, shpColl(1).Top, 55) = True Then
  Label1.Caption = "Collision!"
 Else
  Label1.Caption = "Not Colliding"
 End If
 
End Sub
