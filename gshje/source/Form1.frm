VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "GhostSquadHackers JavaScript Encryptor   -   (c) by Necronomikon"
   ClientHeight    =   4485
   ClientLeft      =   5385
   ClientTop       =   4200
   ClientWidth     =   8730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   5760
      Picture         =   "Form1.frx":058A
      ScaleHeight     =   2865
      ScaleWidth      =   2865
      TabIndex        =   16
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xtras..."
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   9.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear!"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      Left            =   4320
      TabIndex        =   14
      Top             =   2880
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit!"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   11
      Top             =   3600
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Encoding Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   225
      TabIndex        =   2
      Top             =   2160
      Width           =   3930
      Begin VB.OptionButton Option8 
         BackColor       =   &H00000000&
         Caption         =   "Binary Encrypt"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2205
         TabIndex        =   12
         Top             =   810
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Other"
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   135
         TabIndex        =   8
         Top             =   1080
         Width           =   3660
         Begin VB.OptionButton Option9 
            BackColor       =   &H00000000&
            Caption         =   "URL Encode"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   1890
            TabIndex        =   13
            Top             =   630
            Width           =   1635
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00000000&
            Caption         =   "Add trash to code"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   135
            TabIndex        =   10
            Top             =   630
            Width           =   1995
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00000000&
            Caption         =   "Random Octal Quotes"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   315
            Width           =   1995
         End
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "Caesar-Encryption"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   765
         Width           =   1680
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Octal Encoding"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2205
         TabIndex        =   6
         Top             =   495
         Width           =   1500
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Hex Encoding"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2205
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "ASCII codes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   495
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Number calculating"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Encode!"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   4275
      TabIndex        =   1
      Top             =   2160
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   1875
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   120
      Width           =   5355
   End
   Begin VB.Menu bl 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open File"
      End
      Begin VB.Menu lb 
         Caption         =   "-"
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu abt 
         Caption         =   "about"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function junk(Length As Integer)
    Dim i As Integer: Dim p As String: p = ""
    For i = 1 To Length
        Randomize Timer
        If Int(2 * Rnd + 1) = 1 Then
            p = p + Chr(Int((57 - 48 + 1) * Rnd + 48))
        Else
            p = p + Chr(Int((90 - 65 + 1) * Rnd + 65))
        End If
    Next i
    junk = p
End Function



Private Sub abt_Click()
MsgBox "JSE (JavaScript Encoder) encodes/encryptes JavaScript sources. It has 9 different methods and all work fine. Modified version from 'JSE - Genetix -  Javascript Encoder v4'"
End Sub


Public Function URLEncode(sRawURL As String) As String
    On Error GoTo Catch
    Dim iLoop As Integer
    Dim sRtn As String
    Dim sTmp As String
    Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~&"
    If Len(sRawURL) > 0 Then
        For iLoop = 1 To Len(sRawURL)
            sTmp = Mid(sRawURL, iLoop, 1)
            If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                sTmp = Hex(Asc(sTmp))
                If Len(sTmp) = 1 Then
                    sTmp = "%0" & sTmp
                Else
                    sTmp = "%" & sTmp
                End If
            End If
            sRtn = sRtn & sTmp
        Next iLoop
        URLEncode = "unescape('" & sRtn & "')"
    End If
Finally:
    Exit Function
Catch:
    URLEncode = ""
    Resume Finally
End Function

Sub octal()
    Dim b As String
    Dim i As Long
    For i = 1 To Len(Text1.Text)
        b = b & "\" & Oct(Asc(Mid(Text1.Text, i, 1)))
    Next
    Text1.Text = "eval('" & b & "')"
End Sub

Sub ascii()
    Dim i As Long
    Dim arTemp() As Integer
    Dim arOut() As String
    If Text1.Text = "" Then
        Exit Sub
    End If
    ReDim arTemp(Len(Text1.Text) - 1)
    For i = 1 To Len(Text1.Text)
        arTemp(i - 1) = Asc(Mid(Text1.Text, i, 1))
    Next

    ReDim arOut(UBound(arTemp))
    For i = LBound(arTemp) To UBound(arTemp)
        arOut(i) = CStr(arTemp(i)) & ","
    Next
    Text1.Text = "eval(String.fromCharCode(" + Join(arOut, "") & ")"
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 2) + "))"
End Sub

Sub Hex_()
    Dim a As String
    For i = 1 To Len(Text1.Text)
        a = a & "\x" & Hex(Asc(Mid(Text1.Text, i, 1)))
    Next
    Text1.Text = "eval('" & a & "')"
End Sub

Sub numcalc()
    Dim i As Long
    Dim arTemp() As Integer
    Dim arOut() As String
    Dim iRandom As Long
    Dim eRandom As Integer
    If Text1.Text = "" Then
        Exit Sub
    End If
    ReDim arTemp(Len(Text1.Text) - 1)
    For i = 1 To Len(Text1.Text)
        arTemp(i - 1) = Asc(Mid(Text1.Text, i, 1))
    Next

    ReDim arOut(UBound(arTemp))
    For i = LBound(arTemp) To UBound(arTemp)
        Randomize
        iRandom = Int(10 * Rnd) + 9
        eRandom = Int(3 * Rnd) + 1
        If (eRandom = 1) Then arOut(i) = "(" & CStr(arTemp(i) + iRandom) & "-" & CStr(iRandom) & "),"
        If (eRandom = 2) Then arOut(i) = "(" & CStr(arTemp(i) * iRandom) & "/" & CStr(iRandom) & "),"
        If (eRandom = 3) Then arOut(i) = "(" & CStr(arTemp(i) - iRandom) & "+" & CStr(iRandom) & "),"
    Next
    Text1.Text = "eval(String.fromCharCode(" + Join(arOut, "") & ")"
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 2) + "))"
End Sub


Sub Caesar()
    If Text1.Text = "" Then
        MsgBox "Nothing to encrypt", 16, "Error"
        Exit Sub
    End If
    MsgBox "This type of encryption add's 3 ascii code's to the original ascii value. a=d b=e ect" & vbCrLf & "This should only be used after other encryptions or alone!", vbInformation, "Notice"
    Dim tmp, a, b, c As String
    tmp = ""
    For i = 1 To Len(Text1.Text)
        tmp = tmp & Chr(Asc(Mid(Text1.Text, i, 1)) + 3)
    Next
    Text1.Text = "eval(de('" & tmp & "'));"
    Text1.Text = Text1.Text & vbCrLf & "function de(string){" & vbCrLf & "var tmp='';" & _
                 vbCrLf & "eval('\x66\x6F\x72\x28\x69\x3D\x30\x3B\x69\x3C\x73\x74\x72\x69\x6E\x67\x2E\x6C\x65\x6E\x67\x74\x68\x3B\x2B\x2B\x69\x29\x7B\x74\x6D\x70\x2B\x3D\x53\x74\x72\x69\x6E\x67\x2E\x66\x72\x6F\x6D\x43\x68\x61\x72\x43\x6F\x64\x65\x28\x73\x74\x72\x69\x6E\x67\x2E\x63\x68\x61\x72\x43\x6F\x64\x65\x41\x74\x28\x69\x29\x2D\x33\x29\x3B\x7D')" _
                 & vbCrLf & "return(tmp);}"

End Sub

Sub trash()
    Dim strTxt As String
    Dim strParts() As String
    Dim i, x As Integer
    Dim c As String
    If InStr(1, Text1, "eval", 3) Then
        MsgBox "Sorry, Can't use this option with eval encryption", 16, "Error"
        Option1.Value = 0
        Option2.Value = 0
        Option3.Value = 0
        Option4.Value = 0
        Option7.Value = 0
        Exit Sub
    End If
    strTxt = Text1.Text
    strParts = Split(strTxt, vbCrLf)

    For i = 0 To UBound(strParts)
        x = Int(2 * Rnd) + 1
        If x = 1 Then
            c = c & vbCrLf & "//" & junk(Int(50 * Rnd) + 15) & vbCrLf & strParts(i)
        Else
            c = c & vbCrLf & strParts(i)
        End If
        Text1.Text = c
    Next i
End Sub

Sub BinCrypt()
    Dim lInput As Double
    Dim strOutput As String
    Dim strOutput2 As String
    Dim x, i As Integer
    txtText = Text1
    rtn = 0
    strOutput2 = ""
    For i = 1 To Len(Text1.Text)
        lInput = Asc(Mid(Text1.Text, i, 1))
        strOutput = ""
        While lInput > 0
            x = lInput Mod 2
            strOutput = x & strOutput
            lInput = Int(lInput / 2)
        Wend
        While Len(strOutput) < 8
            strOutput = 0 & strOutput
        Wend
        strOutput2 = strOutput2 & strOutput
    Next
    Text1.Text = "eval(de_('" & strOutput2 & "'));" & _
    vbCrLf & "function de_(code){" & vbCrLf _
& "var tmp='';" _
& vbCrLf & "eval('\167\150\151\154\145\50\143\157\144\145\56\154\145\156\147\164\150\40\45\40\70\40\75\75\40\60\40\46\46\40\143\157\144\145\56\154\145\156\147\164\150\40\41\75\40\60\51\173\164\155\160\40\53\75\40\123\164\162\151\156\147\56\146\162\157\155\103\150\141\162\103\157\144\145\50\160\141\162\163\145\111\156\164\50\143\157\144\145\56\163\165\142\163\164\162\50\60\54\70\51\54\40\62\51\51\73\40\143\157\144\145\40\75\40\143\157\144\145\56\162\145\160\154\141\143\145\50\143\157\144\145\56\163\165\142\163\164\162\50\60\54\70\51\54\42\42\51\73\175')" _
& vbCrLf & "return (tmp);}"
End Sub

Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command11_Click(Index As Integer)
If Text1.Text = "" Then
        MsgBox "Nothing to encrypt", 16, "Error"
        Exit Sub
    End If
    If Option1.Value = True Then Call numcalc
    If Option2.Value = True Then Call ascii
    If Option3.Value = True Then Call Hex_
    If Option4.Value = True Then Call octal
    If Option7.Value = True Then Call trash
    If Option8.Value = True Then Call BinCrypt
    If Option9.Value = True Then Text1.Text = URLEncode(Text1.Text)
    If Option5.Value = True Then
        If InStr(5, Text1, "de", 3) Then
            MsgBox "Sorry, Can't use this option more than once!", 16, "Error"
            Exit Sub
        Else
            Call Caesar
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Command3_Click(Index As Integer)
Text1.Text = ""
End Sub

Private Sub ext_Click()
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then Unload Me
End Sub

Sub mnuOpen_Click()
    Dim f As Integer

    On Error GoTo OpenErr1

    CommonDialog1.ShowOpen
    FileName = CommonDialog1.FileName
    On Error GoTo OpenErr2
    f = FreeFile
    Open FileName For Input As f
    Text1.Text = Input$(LOF(f), f)
    Close f
    Form1.Caption = "Text Editor : " & FileName
    Exit Sub


OpenErr2:
    MsgBox "Error opening file", vbOK, "Text editor"
    Close f
    Exit Sub


OpenErr1:
    Exit Sub

End Sub

Private Sub Option6_Click()
    If Option6.Value = True Then Form2.Show
End Sub

