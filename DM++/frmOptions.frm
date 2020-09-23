VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Options......"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   2460
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   6585
      Begin VB.PictureBox BCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4470
         ScaleHeight     =   225
         ScaleWidth      =   480
         TabIndex        =   18
         Top             =   1755
         Width           =   540
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   4425
         TabIndex        =   17
         Top             =   1710
         Width           =   1020
      End
      Begin VB.PictureBox FCOL 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   3315
         ScaleHeight     =   225
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1755
         Width           =   540
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   3270
         TabIndex        =   14
         Top             =   1710
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3330
         ScaleHeight     =   630
         ScaleWidth      =   2115
         TabIndex        =   11
         Top             =   660
         Width           =   2175
         Begin VB.Label lblStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ABCDEFabcdef"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   60
            TabIndex        =   12
            Top             =   195
            Width           =   1440
         End
      End
      Begin VB.ListBox lstSize 
         Height          =   1620
         Left            =   2325
         TabIndex        =   9
         Top             =   660
         Width           =   795
      End
      Begin VB.ListBox lstFont 
         Height          =   1620
         Left            =   105
         TabIndex        =   8
         Top             =   660
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Back-Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4395
         TabIndex        =   16
         Top             =   1410
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2460
         TabIndex        =   13
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sample"
         Height          =   195
         Left            =   4110
         TabIndex        =   10
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fore-Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3300
         TabIndex        =   7
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   3840
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   6570
      Begin VB.CheckBox Indent 
         Caption         =   "Auto Indent Text"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   585
         Width           =   3510
      End
      Begin VB.CheckBox ErrCheck 
         Caption         =   "Halt on All Errors"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   3300
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCan_Click()
    Unload frmOptions
    Form1.Show
    
End Sub

Private Sub Command1_Click()
    Unload frmOptions
    WriteINIChanges
    '
    Form1.Text1.fontname = Config.Font_Name
    Form1.Text1.Font = Config.Font_Size
    Form1.Text1.ForeColor = Config.Fore_Colour
    Form1.Text1.BackColor = Config.Back_Colour
    Form1.Show
    
End Sub

Private Sub Command2_Click()
On Error Resume Next
    FCOL.BackColor = ShowColor(hwnd)
    lblStyle.ForeColor = FCOL.BackColor
    Config.Fore_Colour = FCOL.BackColor
    If Err Then Err.Clear
End Sub

Private Sub Command3_Click()
On Error Resume Next
    BCol.BackColor = ShowColor(hwnd)
    Picture1.BackColor = BCol.BackColor
    Config.Back_Colour = BCol.BackColor
    If Err Then Err.Clear
    
End Sub

Private Sub ErrCheck_Click()
    If ErrCheck Then
        SkipErr = 1
    Else
        SkipErr = 0
    End If
    
End Sub
Private Sub Form_Load()
Dim IFont As Integer
    For IFont = 1 To Screen.FontCount - 1
        lstFont.AddItem Screen.Fonts(IFont)
    Next
    
    lstSize.AddItem "8"
    lstSize.AddItem "9"
    lstSize.AddItem "10"
    lstSize.AddItem "12"
    lstSize.AddItem "14"
    lstSize.AddItem "16"
    lstSize.AddItem "18"
    lstSize.AddItem "24"
    IFont = 0
    If Config.Check_Error = 1 Then
        ErrCheck.Value = 1
    Else
        ErrCheck.Value = 0
    End If
    
    If Config.Indent_Text = 1 Then
        Indent.Value = 1
    Else
        Indent.Value = 0
    End If
    
    FCOL.BackColor = Config.Fore_Colour
    BCol.BackColor = Config.Back_Colour
    
    
    
End Sub

Private Sub Indent_Click()
    If Indent Then
       Auto_Indent = 1
    Else
        Auto_Indent = 0
    End If
    
End Sub

Private Sub lstFont_Click()
    lblStyle.Font = lstFont.Text
    Config.Font_Name = lstFont.Text
    
End Sub

Private Sub lstSize_Click()
    lblStyle.FontSize = Val(lstSize.Text)
    lblStyle.Top = Picture1.Height / 2 - 150
    Config.Font_Size = lstSize.Text
    
End Sub
