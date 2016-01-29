VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Unicode Toolkit"
   ClientHeight    =   11490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnDiacriticRemover 
      Caption         =   "Diacritic remover"
      Height          =   600
      Left            =   5550
      TabIndex        =   10
      Top             =   225
      Width           =   1740
   End
   Begin VB.CommandButton btnCode2Char 
      Caption         =   "Code to character"
      Height          =   600
      Left            =   3390
      TabIndex        =   25
      Top             =   225
      Width           =   1740
   End
   Begin VB.CommandButton btnChar2Code 
      Caption         =   "Character to code"
      Height          =   600
      Left            =   1230
      TabIndex        =   27
      Top             =   225
      Width           =   1740
   End
   Begin UnicodeToolkit.FrameW frmDiacriticRemover 
      Height          =   4900
      Left            =   255
      Top             =   6015
      Width           =   7860
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Form1.frx":0442
      Begin VB.CommandButton btnReplace 
         Caption         =   "&Remove Diacritics"
         Height          =   855
         Left            =   2640
         TabIndex        =   28
         Top             =   3720
         Width           =   2535
      End
      Begin UnicodeToolkit.TextBoxW txtReplace 
         Height          =   2985
         Left            =   255
         TabIndex        =   26
         Top             =   480
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   5265
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         WantReturn      =   -1  'True
      End
   End
   Begin UnicodeToolkit.FrameW frmCode2Char 
      Height          =   4900
      Left            =   8415
      Top             =   945
      Width           =   7860
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Form1.frx":049A
      Begin UnicodeToolkit.TextBoxW txtDec2 
         Height          =   330
         Left            =   825
         TabIndex        =   12
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowOnlyNumbers=   -1  'True
         MaxLength       =   7
      End
      Begin UnicodeToolkit.TextBoxW txtHex2 
         Height          =   330
         Left            =   3435
         TabIndex        =   14
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   7
      End
      Begin UnicodeToolkit.TextBoxW txtOct2 
         Height          =   330
         Left            =   6075
         TabIndex        =   16
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowOnlyNumbers=   -1  'True
         MaxLength       =   7
      End
      Begin UnicodeToolkit.TextBoxW txtChr2 
         Height          =   1500
         Left            =   2130
         TabIndex        =   18
         Top             =   1145
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   45.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtName 
         Height          =   330
         Left            =   255
         TabIndex        =   20
         Top             =   2970
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtBlock 
         Height          =   330
         Left            =   255
         TabIndex        =   22
         Top             =   3630
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtPlane 
         Height          =   330
         Left            =   255
         TabIndex        =   24
         Top             =   4290
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.Label lblPlane 
         AutoSize        =   -1  'True
         Caption         =   "Unicode Plane:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   4065
         Width           =   1095
      End
      Begin VB.Label lblBlock 
         AutoSize        =   -1  'True
         Caption         =   "Unicode Block:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   3390
         Width           =   1095
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Top             =   2715
         Width           =   465
      End
      Begin VB.Label lblchr2 
         AutoSize        =   -1  'True
         Caption         =   "Character:"
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblOct 
         AutoSize        =   -1  'True
         Caption         =   "Oct:"
         Height          =   195
         Left            =   5655
         TabIndex        =   15
         Top             =   495
         Width           =   300
      End
      Begin VB.Label lblHex 
         AutoSize        =   -1  'True
         Caption         =   "Hex:"
         Height          =   195
         Left            =   3015
         TabIndex        =   13
         Top             =   495
         Width           =   330
      End
      Begin VB.Label lblDec 
         AutoSize        =   -1  'True
         Caption         =   "Dec:"
         Height          =   195
         Left            =   405
         TabIndex        =   11
         Top             =   495
         Width           =   345
      End
   End
   Begin UnicodeToolkit.FrameW frmChar2Code 
      Height          =   4905
      Left            =   270
      Top             =   975
      Width           =   7860
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Form1.frx":04E0
      Begin VB.CheckBox chkHTML 
         Caption         =   "HTML Syntax"
         Height          =   315
         Left            =   645
         TabIndex        =   9
         Top             =   4395
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin UnicodeToolkit.TextBoxW txtChr 
         Height          =   1560
         Left            =   270
         TabIndex        =   2
         Top             =   585
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   2752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         WantReturn      =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtDec 
         Height          =   1800
         Left            =   270
         TabIndex        =   4
         Top             =   2505
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   3175
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         WantReturn      =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtHex 
         Height          =   1800
         Left            =   2880
         TabIndex        =   6
         Top             =   2505
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   3175
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         WantReturn      =   -1  'True
      End
      Begin UnicodeToolkit.TextBoxW txtOct 
         Height          =   1800
         Left            =   5520
         TabIndex        =   8
         Top             =   2505
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   3175
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         WantReturn      =   -1  'True
      End
      Begin VB.Label lblOctCodes 
         AutoSize        =   -1  'True
         Caption         =   "Octal:"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   2265
         Width           =   420
      End
      Begin VB.Label lblHexCodes 
         AutoSize        =   -1  'True
         Caption         =   "Hexadecimal:"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   2265
         Width           =   960
      End
      Begin VB.Label lblDecCodes 
         AutoSize        =   -1  'True
         Caption         =   "Decimal:"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   2265
         Width           =   615
      End
      Begin VB.Label lblchr 
         AutoSize        =   -1  'True
         Caption         =   "Character(s):"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   10920
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
'Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Private txtDec2_hasfocus As Boolean
Private txtHex2_hasfocus As Boolean
Private txtOct2_hasfocus As Boolean
Private active_btn As Integer

Private Sub pressbtn()
If active_btn = 3 Then
 Call SendMessage(btnChar2Code.hWnd, 243, 0, 0)
 Call SendMessage(btnCode2Char.hWnd, 243, 0, 0)
 Call SendMessage(btnDiacriticRemover.hWnd, 243, 1, 0)
 frmChar2Code.Visible = False
 frmCode2Char.Visible = False
 frmDiacriticRemover.Visible = True
 frmDiacriticRemover.Left = frmChar2Code.Left
 frmDiacriticRemover.Top = frmChar2Code.Top
 frmDiacriticRemover.Width = frmChar2Code.Width
 frmDiacriticRemover.Height = frmChar2Code.Height
 txtReplace.SetFocus
 txtReplace.SelStart = 0
 txtReplace.SelLength = Len(txtReplace.Text)
ElseIf active_btn = 2 Then
 Call SendMessage(btnChar2Code.hWnd, 243, 0, 0)
 Call SendMessage(btnCode2Char.hWnd, 243, 1, 0)
 Call SendMessage(btnDiacriticRemover.hWnd, 243, 0, 0)
 frmChar2Code.Visible = False
 frmDiacriticRemover.Visible = False
 frmCode2Char.Visible = True
 frmCode2Char.Left = frmChar2Code.Left
 frmCode2Char.Top = frmChar2Code.Top
 frmCode2Char.Width = frmChar2Code.Width
 frmCode2Char.Height = frmChar2Code.Height
 txtDec2.SetFocus
 txtDec2.SelStart = 0
 txtDec2.SelLength = Len(txtDec2.Text)
Else
 Call SendMessage(btnChar2Code.hWnd, 243, 1, 0)
 Call SendMessage(btnCode2Char.hWnd, 243, 0, 0)
 Call SendMessage(btnDiacriticRemover.hWnd, 243, 0, 0)
 frmDiacriticRemover.Visible = False
 frmCode2Char.Visible = False
 frmChar2Code.Visible = True
 txtChr.SetFocus
 txtChr.SelStart = 0
 txtChr.SelLength = Len(txtChr.Text)
End If
End Sub

Private Sub btnChar2Code_Click()
 active_btn = 1
 Call pressbtn
End Sub

Private Sub btnChar2Code_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Me.WindowState <> 2 Then
  active_btn = 1
  Call pressbtn
 End If
End Sub

Private Sub btnCode2Char_Click()
 active_btn = 2
 Call pressbtn
End Sub

Private Sub btnCode2Char_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Me.WindowState <> 2 Then
  active_btn = 2
  Call pressbtn
 End If
End Sub

Private Sub btnDiacriticRemover_Click()
 active_btn = 3
 Call pressbtn
End Sub

Private Sub btnDiacriticRemover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Me.WindowState <> 2 Then
  active_btn = 3
  Call pressbtn
 End If
End Sub

Private Sub btnExit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 RemoveMenu GetSystemMenu(Me.hWnd, 0), 2, &H400& 'prevent resizing
 
 Me.Width = 8545
 Me.Height = 7895
 btnExit.Top = 6180
 btnExit.Left = (Me.Width - btnExit.Width - 105) / 2
 
 active_btn = 1
 frmCode2Char.Visible = False
 frmCode2Char.Left = frmChar2Code.Left
 frmCode2Char.Top = frmChar2Code.Top
 frmCode2Char.Width = frmChar2Code.Width
 frmCode2Char.Height = frmChar2Code.Height
 
 frmDiacriticRemover.Visible = False
 frmDiacriticRemover.Left = frmChar2Code.Left
 frmDiacriticRemover.Top = frmChar2Code.Top
 frmDiacriticRemover.Width = frmChar2Code.Width
 frmDiacriticRemover.Height = frmChar2Code.Height
End Sub

Private Sub Form_Activate()
 txtChr.SetFocus
 Call pressbtn
End Sub

Private Sub btnReplace_Click()
 Dim s As String
 s = txtReplace.Text
 Call replace_diacritics(s)
 txtReplace.Text = s
End Sub

Private Sub lblBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SetCursor LoadCursor(0, 32649&)
End Sub

Private Sub lblPlane_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SetCursor LoadCursor(0, 32649&)
End Sub

Private Sub lblBlock_Click()
 ShellExecute 0&, "open", "https://en.wikipedia.org/wiki/Unicode_block", vbNullString, vbNullString, 3&
End Sub

Private Sub lblPlane_Click()
 ShellExecute 0&, "open", "https://en.wikipedia.org/wiki/Plane_(Unicode)", vbNullString, vbNullString, 3&
End Sub

Private Sub txtBlock_KeyPress(KeyChar As Integer)
 'KeyChar = 0
End Sub

Private Sub txtChr_Change()
txtDec.Text = ""
txtHex.Text = ""
txtOct.Text = ""
txtDec2.Text = ""
txtHex2.Text = ""
txtOct2.Text = ""
Dim i As Long
Dim dec As String
Dim hex As String
Dim oct As String
For i = 1 To Len(txtChr.Text)
 If AscW(Mid$(txtChr.Text, i, 1)) <> 13 And AscW(Mid$(txtChr.Text, i, 1)) <> 10 Then
  dec = CStr(char2dec(Mid$(txtChr.Text, i, 1)))
  hex = dec2hex(char2dec(Mid$(txtChr.Text, i, 1)))
  oct = dec2oct(char2dec(Mid$(txtChr.Text, i, 1)))
  If (dec = 0) Then
   dec = CStr(char2dec(Mid$(txtChr.Text, i, 2)))
   hex = dec2hex(char2dec(Mid$(txtChr.Text, i, 2)))
   oct = dec2oct(char2dec(Mid$(txtChr.Text, i, 2)))
   i = i + 1
   If ((dec >= 8960) And (dec <= 11263)) Or (dec > 100000) Then
    txtChr.Font.Name = "Segoe UI Symbol"
   Else
    txtChr.Font.Name = "Segoe UI"
   End If
  End If
  txtDec2.Text = dec
  txtHex2.Text = hex
  txtOct2.Text = oct
  If Len(txtHex2.Text) < 4 Then
  txtHex2.Text = String$(4 - Len(txtHex2.Text), "0") & txtHex2.Text
  End If
  If Len(txtOct2.Text) < 4 Then
  txtOct2.Text = String$(4 - Len(txtOct2.Text), "0") & txtOct2.Text
  End If
  If Len(hex) < 4 And chkHTML.Value = 0 Then
  hex = String$(4 - Len(hex), "0") & hex
  End If
  If Len(oct) < 4 Then
  oct = String$(4 - Len(oct), "0") & oct
  End If
  If chkHTML.Value = 1 Then
  dec = "&#" & dec & ";"
  hex = "&#x" & hex & ";"
  End If
  txtDec.Text = txtDec.Text & dec & vbNewLine
  txtHex.Text = txtHex.Text & hex & vbNewLine
  txtOct.Text = txtOct.Text & oct & vbNewLine
 End If
Next
End Sub

Private Sub chkHTML_Click()
 Call txtChr_Change
End Sub

Private Sub txtChr_KeyUp(KeyCode As Integer, Shift As Integer)
 Call txtChr_Change
End Sub

Private Sub txtChr2_GotFocus()
 txtChr2.SelStart = 0
 txtChr2.SelLength = Len(txtChr2.Text)
End Sub

Private Sub txtChr2_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift = 2) And (KeyCode = 67) Then
 Clipboard.Clear
 Clipboard.SetText txtChr2.Text
End If
End Sub

Private Sub txtChr2_KeyPress(KeyChar As Integer)
 'KeyChar = 0
End Sub

Private Sub txtDec_GotFocus()
 txtDec.SelStart = 0
 txtDec.SelLength = Len(txtDec.Text)
End Sub

Private Sub txtDec2_KeyPress(KeyChar As Integer)
 'If (KeyChar <> 8) And (KeyChar <> 127) And ((KeyChar < 48) Or (KeyChar > 57)) Then
 ' KeyChar = 0
 'End If
End Sub

Private Sub txtHex_GotFocus()
 txtHex.SelStart = 0
 txtHex.SelLength = Len(txtHex.Text)
End Sub

Private Sub txtHex2_KeyPress(KeyChar As Integer)
 If (KeyChar <> 8) And (KeyChar <> 127) And ((KeyChar < 48) Or (KeyChar > 57)) And ((KeyChar < 65) Or (KeyChar > 70)) And ((KeyChar < 97) Or (KeyChar > 102)) Then
  KeyChar = 0
 End If
End Sub

Private Sub txtName_GotFocus()
 txtName.SelStart = 0
 txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtBlock_GotFocus()
 txtBlock.SelStart = 0
 txtBlock.SelLength = Len(txtBlock.Text)
End Sub

Private Sub txtName_KeyPress(KeyChar As Integer)
 'KeyChar = 0
End Sub

Private Sub txtOct2_KeyPress(KeyChar As Integer)
 'If (KeyChar <> 8) And (KeyChar <> 127) And ((KeyChar < 48) Or (KeyChar > 57)) Then
 ' KeyChar = 0
 'End If
End Sub

Private Sub txtPlane_GotFocus()
 txtPlane.SelStart = 0
 txtPlane.SelLength = Len(txtPlane.Text)
End Sub

Private Sub txtPlane_KeyPress(KeyChar As Integer)
 'KeyChar = 0
End Sub

Private Sub txtOct_GotFocus()
 txtOct.SelStart = 0
 txtOct.SelLength = Len(txtOct.Text)
End Sub

Private Sub txtReplace_GotFocus()
 txtReplace.SelStart = 0
 txtReplace.SelLength = Len(txtReplace.Text)
End Sub

'--------------------------------------------------------------------------------------------------
'txtDec2 Functions
'--------------------------------------------------------------------------------------------------
Private Sub txtDec2_GotFocus()
 txtDec2_hasfocus = True
 txtHex2_hasfocus = False
 txtOct2_hasfocus = False
End Sub

Private Sub txtDec2_LostFocus()
 txtDec2_hasfocus = False
 txtHex2_hasfocus = False
 txtOct2_hasfocus = False
End Sub


Private Sub txtDec2_KeyDown(KeyCode As Integer, Shift As Integer)
 If (Shift = 2) And (KeyCode = 86) Then 'CTRL+V
  Dim str As String
  str = Clipboard.GetText
  If is_valid_decimal(str) Then
   txtDec2.Text = str
  End If
 End If
End Sub

Private Sub txtDec2_KeyUp(KeyCode As Integer, Shift As Integer)
 If Len(txtDec2.Text) > 0 Then
  If Val(txtDec2.Text) > 1114111 Then
   txtDec2.Text = "1114111"
  End If
 End If
 Call txtDec2_Change
End Sub

Private Sub txtDec2_Change()
If (txtDec2_hasfocus = True) Then
 If Len(txtDec2.Text) > 0 Then
  If Val(txtDec2.Text) > 1114111 Then
   txtDec2.Text = "1114111"
  End If
  txtHex2.Text = dec2hex(CLng(txtDec2.Text))
  txtOct2.Text = dec2oct(CLng(txtDec2.Text))
  If Len(txtHex2.Text) < 4 Then
  txtHex2.Text = String$(4 - Len(txtHex2.Text), "0") & txtHex2.Text
  End If
  If Len(txtOct2.Text) < 4 Then
  txtOct2.Text = String$(4 - Len(txtOct2.Text), "0") & txtOct2.Text
  End If
 Else
  txtHex2.Text = ""
  txtOct2.Text = ""
 End If
End If
If Len(txtDec2.Text) > 0 Then
 If ((CLng(txtDec2.Text) >= 8960) And (CLng(txtDec2.Text) <= 11263)) Or (CLng(txtDec2.Text) > 100000) Then
  txtChr2.Font.Name = "Segoe UI Symbol"
 Else
  txtChr2.Font.Name = "Segoe UI"
 End If
 If txtDec2.Text = "11" Then
 txtChr2.Text = ""
 Else
 txtChr2.Text = dec2char(CLng(txtDec2.Text))
 End If
 txtName.Text = get_unicode_name(CLng(txtDec2.Text))
 txtBlock.Text = get_unicode_block(CLng(txtDec2.Text))
 txtPlane.Text = get_unicode_plane(CLng(txtDec2.Text))
Else
 txtChr2.Text = ""
 txtName.Text = ""
 txtBlock.Text = ""
 txtPlane.Text = ""
End If
End Sub

'--------------------------------------------------------------------------------------------------
'txtHex2 Functions
'--------------------------------------------------------------------------------------------------
Private Sub txtHex2_GotFocus()
 txtDec2_hasfocus = False
 txtHex2_hasfocus = True
 txtOct2_hasfocus = False
End Sub

Private Sub txtHex2_LostFocus()
 txtDec2_hasfocus = False
 txtHex2_hasfocus = False
 txtOct2_hasfocus = False
End Sub

Private Sub txtHex2_KeyDown(KeyCode As Integer, Shift As Integer)
 If (Shift = 2) And (KeyCode = 86) Then 'CTRL+V
  Dim str As String
  str = Clipboard.GetText
  If is_valid_hex(str) Then
   txtHex2.Text = str
  End If
 End If
End Sub

Private Sub txtHex2_KeyUp(KeyCode As Integer, Shift As Integer)
 If Len(txtHex2.Text) > 0 Then
 If hex2dec(txtHex2.Text) > 1114111 Then
  txtHex2.Text = "10FFFF"
 End If
 End If
 Call txtHex2_Change
End Sub

Private Sub txtHex2_Change()
 If (txtHex2_hasfocus = True) Then
  If Len(txtHex2.Text) > 0 Then
   txtDec2.Text = hex2dec(txtHex2.Text)
   txtOct2.Text = hex2oct(txtHex2.Text)
  Else
   txtDec2.Text = ""
   txtOct2.Text = ""
  End If
 End If
End Sub

'--------------------------------------------------------------------------------------------------
'txtOct2 Functions
'--------------------------------------------------------------------------------------------------
Private Sub txtOct2_GotFocus()
 txtDec2_hasfocus = False
 txtHex2_hasfocus = False
 txtOct2_hasfocus = True
End Sub

Private Sub txtOct2_LostFocus()
 txtDec2_hasfocus = False
 txtHex2_hasfocus = False
 txtOct2_hasfocus = False
End Sub

Private Sub txtOct2_KeyDown(KeyCode As Integer, Shift As Integer)
 If (Shift = 2) And (KeyCode = 86) Then 'CTRL+V
  Dim str As String
  str = Clipboard.GetText
  If is_valid_decimal(str) Then
   txtOct2.Text = str
  End If
 End If
End Sub

Private Sub txtOct2_KeyUp(KeyCode As Integer, Shift As Integer)
 If Len(txtOct2.Text) > 0 Then
  If Val(txtOct2.Text) > 4177777 Then
   txtOct2.Text = "4177777"
  End If
 End If
 Call txtOct2_Change
End Sub

Private Sub txtOct2_Change()
 If (txtOct2_hasfocus = True) Then
  If Len(txtOct2.Text) > 0 Then
   If Val(txtOct2.Text) > 4177777 Then
    txtOct2.Text = "4177777"
   End If
   txtHex2.Text = oct2hex(txtOct2.Text)
   txtDec2.Text = oct2dec(txtOct2.Text)
  Else
   txtHex2.Text = ""
   txtDec2.Text = ""
  End If
 End If
End Sub

'Private Function ByteArrayToString(ByRef Text() As Byte, ByVal Length As Long) As String
' Dim lDataLength As Long
' lDataLength = MultiByteToWideChar(65001, 0, VarPtr(Text(0)), Length, 0, 0)
' ByteArrayToString = String$(lDataLength, 0)
' MultiByteToWideChar 65001, 0, VarPtr(Text(0)), Length, StrPtr(ByteArrayToString), lDataLength
'End Function

'Private Sub ByteArrayToFile(Text() As Byte, FileName As String)
' Dim FileNumber As Integer, BOM As Integer
' FileNumber = FreeFile
' BOM = &HFEFF
' Open FileName For Binary As #FileNumber
'  Put #FileNumber, , BOM
'  Put #FileNumber, , Text
' Close #FileNumber
'End Sub
