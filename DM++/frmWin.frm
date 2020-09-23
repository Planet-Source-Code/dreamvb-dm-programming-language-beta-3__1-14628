VERSION 5.00
Begin VB.Form frmWin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "................"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pcode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   4470
      Left            =   0
      ScaleHeight     =   4470
      ScaleWidth      =   7170
      TabIndex        =   0
      Top             =   0
      Width           =   7170
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6315
         Top             =   3990
      End
      Begin VB.Label lblBlink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    Pcode.Left = 0
    Pcode.Top = 0
    
End Sub

Private Sub Timer1_Timer()
Static i As Integer
    i = i + 1
    If Blink_Text_Enabled Then
    lblBlink.Visible = True
    T = Int(Rnd(1 * i) * 2)
        If T = 0 Then
                lblBlink.ForeColor = First_Colour
            ElseIf T = 1 Then
                lblBlink.ForeColor = Last_Colour
            End If
            i = 0
            lblBlink.Caption = Blink_Text
        Else
            lblBlink.Visible = False
        End If
    lblBlink.Left = BlinkTextLeft
    lblBlink.Top = BlinkTextTop
    
        
End Sub
