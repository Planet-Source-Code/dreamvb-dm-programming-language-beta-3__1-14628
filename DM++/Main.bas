Attribute VB_Name = "Main"
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Public Const EM_LINEFROMCHAR = &HC9


Enum WinShow
    vsHide = 0
    vsNormal = 1
    vsMinSized = 2
    vsMaxSized = 3
End Enum

Type Conf
    Check_Error As String
    Indent_Text As String
    Font_Name As String
    Font_Size As String
    Fore_Colour As String
    Back_Colour As String
End Type



Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Config As Conf
Function WriteINIChanges()
    SaveSetting "DM++", "Config", "ErrorCheck", SkipErr
    SaveSetting "DM++", "Config", "IndentText", Auto_Indent
    SaveSetting "DM++", "Config", "FontName", Config.Font_Name
    SaveSetting "DM++", "Config", "FontSize", Config.Font_Size
    SaveSetting "DM++", "Config", "ForeColour", Config.Fore_Colour
    SaveSetting "DM++", "Config", "BackColour", Config.Back_Colour
    
End Function
Function ReadWriteINI()
    Config.Check_Error = GetSetting("DM++", "Config", "ErrorCheck")
    Config.Indent_Text = GetSetting("DM++", "Config", "IndentText")
    Config.Font_Name = GetSetting("DM++", "Config", "FontName")
    Config.Font_Size = GetSetting("DM++", "Config", "FontSize")
    Config.Fore_Colour = GetSetting("DM++", "Config", "ForeColour")
    Config.Back_Colour = GetSetting("DM++", "Config", "BackColour")
    
    If Config.Check_Error = "" Then
        SaveSetting "DM++", "Config", "ErrorCheck", "1"
    End If
    
    If Config.Indent_Text = "" Then
        SaveSetting "DM++", "Config", "IndentText", "0"
    End If
  
    If Config.Font_Name = "" Then
        SaveSetting "DM++", "Config", "FontName", "Courier New"
    End If

    If Config.Font_Size = "" Then
        SaveSetting "DM++", "Config", "FontSize", "10"
    End If

    If Config.Fore_Colour = "" Then
        SaveSetting "DM++", "Config", "ForeColour", "0"
    End If

    If Config.Back_Colour = "" Then
        SaveSetting "DM++", "Config", "BackColour", "16777215"
    End If
    
End Function
Public Function ShowColor(Handle As Long) As Long
Dim TCol As CHOOSECOLOR
Dim Custcolor(41) As Long
Dim lReturn As Long
    
    TCol.lStructSize = Len(TCol)
    TCol.hwndOwner = Handle
    TCol.hInstance = App.hInstance
    TCol.lpCustColors = StrConv(CustomColors, vbUnicode)
    TCol.flags = 0
    
    If CHOOSECOLOR(TCol) <> 0 Then
        ShowColor = TCol.rgbResult
        CustomColors = StrConv(TCol.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If

End Function
Public Function OpenFile() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM++ Files)" + Chr$(0) + "*.DM"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Open Project"
        ofn.flags = 0
        
        a = GetOpenFileName(ofn)
        If (a) Then
                OpenFile = RemoveNulls(Trim(ofn.lpstrFile))
        End If
        
 End Function
 Public Function SaveFile() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM++ Files)" + Chr$(0) + "*.DM"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Save Project"
        ofn.flags = 0
        
        a = GetSaveFileName(ofn)
        If (a) Then
                SaveFile = RemoveNulls(Trim(ofn.lpstrFile))
        End If
 End Function
Function RemoveNulls(lzString As String) As String
Dim Xpos As Integer
    Xpos = InStr(lzString, vbNullChar)
    If Xpos > 0 Then
        lzString = Left(lzString, Len(lzString) - 1)
        RemoveNulls = lzString
    End If
    
End Function
Function EditMenu(txtBox As TextBox, Cmd As String)
Dim StrFind As String
Dim Xpos As Integer

    Select Case Cmd
        Case "CUT"
            Clipboard.SetText txtBox.SelText
            txtBox.SelText = ""
        Case "COPY"
            Clipboard.SetText txtBox.SelText
        Case "PASTE"
            txtBox.SelText = Clipboard.GetText
        Case "SELALL"
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.Text)
            
        Case "FIND"
            StrFind = InputBox("What do you want to find", "Find Text..", , 5, 5)
            If Len(StrFind) = 0 Then
                Exit Function
            Else
                Xpos = InStr(txtBox.Text, StrFind)
                If Xpos > 0 Then
                    txtBox.SetFocus
                    txtBox.SelStart = Xpos - 1
                    txtBox.SelLength = Len(StrFind)
                Else
                    Beep
                    MsgBox "Serach text " & Chr(34) & StrFind & Chr(34) & " was not found", vbExclamation
                End If
            End If
            Xpos = 0
            Cmd = ""
            StrFind = ""
    End Select
    
End Function
Public Function AddBackSlash(lzPathName As String) As String
Dim mPath As String
    If Right(lzPathName, 1) = "\" Then
        mPath = mPath & lzPathName
    Else
        mPath = mPath & lzPathName & "\"
    End If
    AddBackSlash = mPath

End Function

Public Function FileExists(ByVal Filename As String) As Integer
    If Dir(Filename) = "" Then FileExists = 0 Else FileExists = 1
    
End Function
Public Function RunProgran(mHwnd As Long, ProgramNamePath As String, ShowWindow As WinShow)
    If FileExists(ProgramNamePath) = 0 Then
        MsgBox "Can't find file " & ProgramNamePath, vbInformation
    Else
        ShellExecute mHwnd, vbNullString, ProgramNamePath, vbNullString, vbNullString, ShowWindow
    End If
    
End Function
