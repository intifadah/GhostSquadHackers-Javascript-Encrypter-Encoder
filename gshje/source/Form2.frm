VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Random Octal Quotes"
   ClientHeight    =   4905
   ClientLeft      =   3330
   ClientTop       =   3735
   ClientWidth     =   9150
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear"
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   5
      Top             =   4440
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   285
      Left            =   7800
      TabIndex        =   3
      Top             =   4440
      Width           =   1275
   End
   Begin VB.CommandButton Command1a 
      Caption         =   "&Encode"
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   4440
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   3840
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   405
      Width           =   8970
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "This Tool will encode string's between """"'s with random octal quotes."
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   4320
      Width           =   5460
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Paste Unencrypted Code Here:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1a_Click(Index As Integer)
 Dim strText As String
    Dim intPos As Integer
    strText = Text1.Text
    intPos = InStr(strText, Chr(34))
    Do While intPos > 0
        strText = Mid(strText, intPos + 1)
        intPos = InStr(strText, Chr(34))
        Text1.Text = Replace(Text1.Text, Left(strText, intPos - 1), rndQuote(Left(strText, intPos - 1)))
        intPos = InStr(intPos + 1, strText, Chr(34))
    Loop
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Function rndQuote(Str)
    Dim b As String
    Dim x As Long
    Dim i As Long
    For i = 1 To Len(Str)
        Randomize
        x = Int(2 * Rnd) + 1
        If x = 1 Then
            b = b & (Mid(Str, i, 1))
        Else
            b = b & "\" & Oct(Asc(Mid(Str, i, 1)))
        End If
    Next
    rndQuote = b
End Function

Private Sub Command3_Click(Index As Integer)
Text1.Text = ""
End Sub
