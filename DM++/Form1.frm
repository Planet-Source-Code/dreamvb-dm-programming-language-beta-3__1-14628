VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM++ Compiler Beta 3"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Bevel Bevel1 
      Height          =   4815
      Left            =   990
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   465
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   8493
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   5340
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2804
            MinWidth        =   2804
            Picture         =   "Form1.frx":0442
            Text            =   "Press F1 for help"
            TextSave        =   "Press F1 for help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Ln = 0"
            TextSave        =   "Ln = 0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Col = 0"
            TextSave        =   "Col = 0"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4830
      Left            =   0
      ScaleHeight     =   4830
      ScaleWidth      =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   960
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "Form1.frx":1F78
         Top             =   90
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6420
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3508
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":394A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4610
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":52D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5718
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   688
      ButtonWidth     =   661
      ButtonHeight    =   635
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find Text"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   1005
      MouseIcon       =   "Form1.frx":5F9C
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":63DE
      Top             =   480
      Width           =   7755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      Index           =   1
      X1              =   -315
      X2              =   1515
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   -315
      X2              =   1515
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Project"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Project"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindTxt 
         Caption         =   "&Find Text"
      End
   End
   Begin VB.Menu mnuComp1 
      Caption         =   "&Compile"
      Begin VB.Menu mnuComp 
         Caption         =   "C&ompile"
      End
      Begin VB.Menu mnuStopComp 
         Caption         =   "&Stop Compile"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "P&ause"
      End
   End
   Begin VB.Menu mnuPro 
      Caption         =   "&Project"
      Begin VB.Menu mnuOp 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGuid 
         Caption         =   "&See Help Guid"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DM++"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyWords(1 To 24) As String
Dim TCount As Integer
Dim RemmberStuff As String

Sub SaveProject()
Dim Data As String, FExt, Filename As String
Dim Ans

    Filename = SaveFile
    FExt = Right(UCase(Filename), 2)
   
    If Len(Filename) = 0 Then
        Exit Sub
    Else
        If FExt = "DM" Then
    Else
        Filename = Filename & ".dm"
    End If
        Data = Text1.Text
        If Not Dir(Filename) = "" Then
            Ans = _
                MsgBox("Do you want to replace this file", _
                vbYesNo)
                If Ans = vbNo Then
                    Exit Sub
                Else
                    Kill Filename
            End If
        End If
        
        Open Filename For Binary As #1
        Put #1, , Data
        Close #1
    End If
    Filename = ""
    FExt = ""
    Data = ""
    
End Sub
Sub OpenProject()
Dim Data As String, Filename, FExt As String
Dim Filenum As Long

    Filenum = FreeFile
    Filename = OpenFile
    If Len(Filename) = 0 Then
        Exit Sub
    Else
        FExt = UCase(Right(Filename, 2))
        If FExt = "DM" Then
            Open Filename For Binary As #Filenum
            Data = Space(LOF(Filenum))
            Get #Filenum, , Data
            Close #Filenum
            Text1.Text = Data
            Text1.Text = Data
        Else
            MsgBox "This is not a viald DM++ project.", vbCritical, "Error...."
            Exit Sub
        End If
    End If
    Filename = ""
    FExt = ""
    Data = ""
        
End Sub
Function RemoveChar(StrString As String, SChar As String) As String
Dim Xpos As Integer
    Xpos = InStr(StrString, SChar)
    If Xpos Then
        RemoveChar = Left(StrString, Xpos - 1)
    Else
        RemoveChar = StrString
    End If
    Xpos = 0
    
End Function

Function TCompile(lzStr As String)
Dim LineNum  As Integer, i As Integer
Dim StrBuff As String, TCode As String
Dim X, Y, Z As Integer

    For i = 1 To Len(lzStr)
        ch = Asc(Mid(lzStr, i, 1))
        If ch <> 13 Then
            StrBuff = StrBuff & Mid(lzStr, i, 1)
        Else
            LineNum = LineNum + 1
            StrBuff = Trim(StrBuff)
            StrBuff = Replace(StrBuff, Chr(9), "")
            '//-------------------------------------------------------------------
            
            If InStr(StrBuff, KeyWords(1)) = 0 And LineNum = 1 Then
                GetLastError 1, LineNum
                Exit Function
            Else
            End If
            
            
            If InStr(StrBuff, KeyWords(2)) Then
                TCount = TCount + 1
                If TCount > 0 Then
                    If FindPart(StrBuff, ";") = 0 Then
                        GetLastError 2, LineNum
                        Exit Function
                    Else
                    TCode = StrBuff
                    PutToScreen GetText(TCode), frmWin.Pcode
                    End If
                End If
            End If
            TCount = 0
            TCode = ""
            '//-------------------------------------------------------------------
            If InStr(StrBuff, KeyWords(3)) Then
                TCount = TCount + 1
                If TCount > 0 Then
                    If FindPart(StrBuff, ";") = 0 Then
                        GetLastError 1, LineNum
                        Exit Function
                    Else
                    TCode = StrBuff
                    TCode = GetText(TCode)
                    If Len(TCode) = 0 Then
                        GetLastError 3, LineNum
                        Exit Function
                    Else
                        If IsDigit(TCode) = False Then
                            GetLastError 9, LineNum
                            Exit Function
                        Else
                        ScreenModes Val(TCode), frmWin.Pcode
                    End If
                    End If
                End If
            End If
        End If
                '//-------------------------------------------------------------------

        TCount = 0
        TCode = ""
      
      If InStr(StrBuff, KeyWords(4)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                TCol = FindPoint(TCode, "=")
                If TCol > 0 Then
                    TCode = Mid(TCode, TCol + 1, Len(TCode))
                    TCode = Left(TCode, Len(TCode) - 1)
                    SetTextColour TCode, frmWin.Pcode
                Else
                    GetLastError 4, LineNum
                    Exit Function
                End If
            End If
        End If
    End If
            '//-------------------------------------------------------------------

    TCode = ""
    TCount = 0
    If InStr(StrBuff, KeyWords(5)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                If FindPoint(TCode, "(") = 0 Then
                    GetLastError 5, LineNum
                ElseIf FindPoint(TCode, ")") = 0 Then
                    GetLastError 6, LineNum
                    Exit Function
                    Else
                        TCode = GetText(TCode)
                        ShowMsg TCode
                End If
            End If
        End If
    End If
    TCount = 0
    TCode = ""
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(6)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            Plot StrBuff, LineNum, frmWin.Pcode
        End If
    End If
            '//-------------------------------------------------------------------
            
    If InStr(StrBuff, KeyWords(7)) Then
        TCount = TCount + 1
        If TCount > 0 Then
           If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                Beep
            End If
        End If
    End If
    TCount = 0
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(8)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                frmWin.Pcode.Cls
            End If
        End If
    End If
    TCount = 0
                '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(9)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            DrawLine StrBuff, LineNum, frmWin.Pcode
        End If
    End If
    TCount = 0
                '//-------------------------------------------------------------------
                
                
      If InStr(StrBuff, KeyWords(11)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                TCol = FindPoint(TCode, "=")
                If TCol > 0 Then
                    TCode = Mid(TCode, TCol + 1, Len(TCode))
                    TCode = Left(TCode, Len(TCode) - 1)
                    SetBKColour TCode, frmWin.Pcode
                Else
                    GetLastError 4, LineNum
                    Exit Function
                End If
            End If
        End If
    End If
    TCode = ""
    TCount = 0
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(12)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
               TCode = StrBuff
                If FindPoint(TCode, "=") = 0 Then
                    GetLastError 11, LineNum
                    Exit Function
                Else
                    SetTextPositionsX Mid(TCode, FindPoint(TCode, "=") + 1, Len(TCode)), LineNum
                End If
            End If
        End If
    End If
    
    '//-------------------------------------------------------------------
    
        If InStr(StrBuff, KeyWords(13)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
               TCode = StrBuff
                If FindPoint(TCode, "=") = 0 Then
                    GetLastError 11, LineNum
                    Exit Function
                Else
                    SetTextPositionsY Mid(TCode, FindPoint(TCode, "=") + 1, Len(TCode)), LineNum
                End If
            End If
        End If
    End If
    TCode = ""
    TCount = 0
    
    If InStr(StrBuff, KeyWords(14)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                DrawEllipse TCode, LineNum, frmWin.Pcode
            End If
        End If
    End If
    TCode = ""
    TCount = 0
    
    '//-------------------------------------------------------------------
    
    If InStr(StrBuff, KeyWords(15)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
            TCode = StrBuff
            If FindPoint(TCode, "=") = 0 Then
                GetLastError 11, LineNum
                Exit Function
            Else
                TCode = GetRightVal(StrBuff, LineNum)
                If IsDigit(TCode) = False Then
                    GetLastError 9, LineNum
                    Exit Function
                Else
                    Delay Val(TCode)
            End If
            End If
        End If
    End If
    End If
    
    '//-------------------------------------------------------------------
   TCode = ""
   TCount = 0
   
   If InStr(StrBuff, KeyWords(16)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ".") = 0 Then
                Exit Function
                    GetLastError 15, LineNum
                    Exit Function
                Else
                    If FindPart(StrBuff, ";") = 0 Then
                        GetLastError 2, LineNum
                        Exit Function
                    Else
                        frmWin.Show
                        frmWin.Refresh
                    End If
                End If
            End If
        End If
        TCount = 0
        TCode = ""
    '//-------------------------------------------------------------------
        If InStr(StrBuff, KeyWords(17)) Then
            TCount = TCount + 1
            If TCount > 0 Then
                If FindPart(StrBuff, ";") = 0 Then
                    GetLastError 2, LineNum
                    Exit Function
                Else
                    frmWin.Hide
                End If
            End If
        End If
        TCount = 0
    '//-------------------------------------------------------------------
    If InStr(StrBuff, KeyWords(18)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPoint(StrBuff, "=") = 0 Then
                GetLastError 11, LineNum
                Exit Function
            Else
                If FindPart(StrBuff, ";") = 0 Then
                    GetLastError 2, LineNum
                    Exit Function
                Else
                    X = FindPoint(StrBuff, "=")
                    Y = FindPoint(StrBuff, ";")
                
                    If Not Mid(StrBuff, X + 1, 1) = Chr(34) Or Not Mid(StrBuff, Y - 1, 1) = Chr(34) Then
                        GetLastError 16, LineNum
                        Exit Function
                    Else
                        frmWin.Caption = GetStringRight(StrBuff)
                    End If
                End If
            End If
        End If
    End If
    '//-------------------------------------------------------------------
    If InStr(StrBuff, KeyWords(19)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(19))
                Y = FindPoint(StrBuff, "(")
                If Y - X = 9 Then
                TCode = StrBuff
                BlinkLable TCode, LineNum
            End If
        End If
    End If
    End If
    X = 0: Y = 0: TCode = ""
    '
    If InStr(StrBuff, KeyWords(20)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(20))
                Y = FindPoint(StrBuff, "=")
                If Y - X = 17 Then
                TCode = StrBuff
                    Blink_Text = GetStringRight(TCode)
                End If
            End If
        End If
    End If
    X = 0: Y = 0: TCode = ""
    '
    If InStr(StrBuff, KeyWords(21)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(21))
                Y = FindPoint(StrBuff, "=")
                If Y - X = 14 Then
                    TCode = StrBuff
                    BlinkTextLeft = GetRightVal(TCode, LineNum)
                End If
            End If
        End If
    End If
    X = 0: Y = 0
    
    '
    If InStr(StrBuff, KeyWords(22)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(22))
                Y = FindPoint(StrBuff, "=")
                If Y - X = 13 Then
                    TCode = StrBuff
                    BlinkTextTop = GetRightVal(TCode, LineNum)
                End If
            End If
        End If
    End If
    X = 0: Y = 0: TCode = ""
    '
    If InStr(StrBuff, KeyWords(23)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(23))
                Y = FindPoint(StrBuff, "=")
                    TCode = StrBuff
                    If Y - X = 18 Then
                    On Error Resume Next
                        frmWin.lblBlink.FontSize = GetRightVal(TCode, LineNum)
                        If Err Then frmWin.lblBlink.FontSize = 12
                    End If
                End If
            End If
        End If
    X = 0: Y = 0: TCode = ""
    '
    If InStr(StrBuff, KeyWords(24)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                X = InStr(StrBuff, KeyWords(24))
                Y = FindPoint(StrBuff, "=")
                    TCode = StrBuff
                    If Y - X = 18 Then
                    On Error Resume Next
                        frmWin.lblBlink.FontName = GetStringRight(TCode)
                        If Err Then frmWin.lblBlink.FontName = "System"
                    End If
                End If
            End If
        End If
    'X = 0: Y = 0: TCode = ""
    
    
    TCount = 0
    TCode = ""
    '

    TCode = ""
    TCount = 0
    
    StrBuff = ""
        i = i + 1
    End If
    Next
    
    i = 0
    TCode = ""
    StrBuff = ""
    TCount = 0
    LineNum = 0

    
    
End Function

Private Sub Form_Load()
Dim lPos As Integer

On Error Resume Next

    Set Var_Names = New Collection
    Set Var_Types = New Collection
    
    For lPos = 0 To 255
        VChar(lPos) = "Char{" & lPos & "}"
    Next
    lPos = 0
    
    KeyWords(1) = "Procedure()"
    KeyWords(2) = "TextOut"
    KeyWords(3) = "Mode"
    KeyWords(4) = "TextColour"
    KeyWords(5) = "ShowMessage"
    KeyWords(6) = "Plot"
    KeyWords(7) = "Beep"
    KeyWords(8) = "Cls"
    KeyWords(9) = "DrawLine"
    KeyWords(10) = "End Sub"
    
    ' New Functions Added as of 16/01/01
    
    KeyWords(11) = "BkColour"
    KeyWords(12) = "CurrentX"
    KeyWords(13) = "CurrentY"
    KeyWords(14) = "DrawEllipse"
    KeyWords(15) = "Delay"
    KeyWords(16) = "Window.Show"
    KeyWords(17) = "Window.Hide"
    KeyWords(18) = "Window.Caption"
    KeyWords(19) = "BlinkText"
    KeyWords(20) = "BlinkText.Caption"
    KeyWords(21) = "BlinkText.Left"
    KeyWords(22) = "BlinkText.Top"
    KeyWords(23) = "BlinkText.FontSize"
    KeyWords(24) = "BlinkText.FontName"
    
    
    
    TCount = 1
    ModPas.CLB_Colour = &H8000000F
    ModPas.CLF_Colour = vbBlack
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(11).Enabled = False
    mnuStopComp.Enabled = False
    mnuPause.Enabled = False
    
    frmWin.Pcode.Top = Text1.Top
    frmWin.Pcode.Left = Text1.Left
    frmWin.Pcode.Width = Text1.Width
    
    ReadWriteINI
    
    Text1.FontName = Config.Font_Name
    Text1.FontSize = Config.Font_Size
    Text1.BackColor = Config.Back_Colour
    Text1.ForeColor = Config.Fore_Colour
    SkipErr = Config.Check_Error
    Auto_Indent = Config.Indent_Text
    
    If Err Then
        Text1.FontName = "Courier New"
        Text1.FontSize = 10
        Text1.BackColor = vbWhite
        Text1.ForeColor = vbBlack
        Auto_Indent = False
        SkipErr = True
    End If
    
    BlinkTextFontSize = 10
    
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = ScaleWidth - 1
    Line1(1).X2 = ScaleWidth - 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form1: End
    
End Sub

Private Sub mnuAbout_Click()
    FrmAbout1.Show
    Form1.Hide
            
End Sub

Private Sub mnuAll_Click()
    EditMenu Text1, "SELALL"
    
End Sub

Private Sub mnuComp_Click()
    Text1.Enabled = False
    frmWin.Pcode.Visible = True
    RemmberStuff = Text1
    Text1 = TCompile(Text1)
    mnuOpen.Enabled = False
    mnuSave.Enabled = False
    mnuStopComp.Enabled = True
    mnuComp.Enabled = False
    mnuPause.Enabled = True
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(11).Enabled = True
    Text1.Text = RemmberStuff
    
    
End Sub



Private Sub mnuCopy_Click()
    EditMenu Text1, "COPY"
    
End Sub

Private Sub mnuCut_Click()
    EditMenu Text1, "CUT"
    
End Sub

Private Sub mnuEx_Click()
    Ans = _
    MsgBox("Do you want to exit this program now", _
    vbYesNo)
    If Ans = vbNo Then
        Exit Sub
    Else
        RemmberStuff = ""
        Unload Form1: End
    End If
            
End Sub

Private Sub mnuFindTxt_Click()
    EditMenu Text1, "FIND"
    
End Sub

Private Sub mnuGuid_Click()
Dim Path As String
    Path = App.Path
    If Right(Path, 1) = "\" Then
        Path = Path
    Else
        Path = Path & "\"
    End If
    Main.RunProgran hwnd, Path & "Help.txt", vsMaxSized
    Path = ""
    
End Sub

Private Sub mnuOp_Click()
    Form1.Hide
    frmOptions.Show
    
End Sub

Private Sub mnuOpen_Click()
    OpenProject
    
End Sub

Private Sub mnuPaste_Click()
    EditMenu Text1, "PASTE"
    
End Sub

Private Sub mnuPause_Click()
    mnuPause.Enabled = False
    mnuComp.Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = False
    Text1.Enabled = False
End Sub

Private Sub mnuSave_Click()
    SaveProject
    
End Sub

Private Sub mnuStopComp_Click()
    Text1.Enabled = True
    frmWin.Pcode.Visible = False
    frmWin.Pcode.Cls
    Text1.Text = RemmberStuff
    mnuOpen.Enabled = True
    mnuSave.Enabled = True
    mnuStopComp.Enabled = False
    mnuComp.Enabled = True
    mnuPause.Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(11).Enabled = False
    
    RemmberStuff = ""
    RestoreOld frmWin.Pcode
    Unload frmWin
    
End Sub

Private Sub Text1_Change()
Dim Cur_Line As Long
Dim Cur_Col As Long

    On Local Error Resume Next
    Cur_Line = SendMessage(Text1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
    Cur_Col = Text1.SelStart
    StatusBar1.Panels(4).Text = "Ln = " & Format(Cur_Line, "##,###")
    StatusBar1.Panels(5).Text = "Col = " & Cur_Col
    
End Sub

Private Sub Text1_Click()
    Text1_Change
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        mnuComp_Click
    ElseIf KeyCode = 112 Then
        mnuGuid_Click
    End If
    
    If Auto_Indent Then
        If KeyCode = 13 Then
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnuOpen_Click
        Case 2
           mnuSave_Click
        Case 4
            mnuCut_Click
        Case 5
            mnuCopy_Click
        Case 6
            mnuPaste_Click
        Case 7
            mnuFindTxt_Click
        Case 9
            mnuComp_Click
        Case 10
            mnuPause_Click
        Case 11
            mnuStopComp_Click
        Case 13
            FrmAbout1.Show
        Case 14
            mnuEx_Click
    End Select
    
End Sub

