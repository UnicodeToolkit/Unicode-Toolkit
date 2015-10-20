VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "Unicode Toolkit"
   ClientHeight    =   12090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7890
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12090
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPlane 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   255
      TabIndex        =   11
      Top             =   8910
      Width           =   7305
   End
   Begin VB.TextBox txtBlock 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   255
      TabIndex        =   10
      Top             =   8220
      Width           =   7305
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   255
      TabIndex        =   9
      Top             =   7530
      Width           =   7305
   End
   Begin VB.TextBox txtOct2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   7
      Top             =   4950
      Width           =   1335
   End
   Begin VB.TextBox txtHex2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3330
      MaxLength       =   6
      TabIndex        =   6
      Top             =   4950
      Width           =   1335
   End
   Begin VB.TextBox txtDec2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   720
      MaxLength       =   7
      TabIndex        =   5
      Top             =   4950
      Width           =   1335
   End
   Begin VB.CheckBox chkHTML 
      Caption         =   "HTML Syntax"
      Height          =   315
      Left            =   630
      TabIndex        =   4
      Top             =   4140
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   855
      Left            =   4320
      TabIndex        =   15
      Top             =   10995
      Width           =   2535
   End
   Begin VB.CommandButton btnReplace 
      Caption         =   "&Remove Diacritics"
      Default         =   -1  'True
      Height          =   855
      Left            =   1200
      TabIndex        =   13
      Top             =   10995
      Width           =   2535
   End
   Begin VB.Label lblReplace 
      Caption         =   "Diacritic remover (á -> a)                                                                                      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   28
      Top             =   9360
      Width           =   7320
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
      TabIndex        =   27
      Top             =   8685
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
      TabIndex        =   26
      Top             =   7980
      Width           =   1095
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   255
      TabIndex        =   25
      Top             =   7275
      Width           =   465
   End
   Begin MSForms.TextBox txtChr2 
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   2160
      TabIndex        =   8
      Top             =   5670
      Width           =   3480
      VariousPropertyBits=   545275931
      Size            =   "6138;2646"
      SpecialEffect   =   3
      FontName        =   "Segoe UI"
      FontHeight      =   915
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblchr2 
      AutoSize        =   -1  'True
      Caption         =   "Character:"
      Height          =   195
      Left            =   255
      TabIndex        =   24
      Top             =   5475
      Width           =   735
   End
   Begin VB.Label lblOct 
      AutoSize        =   -1  'True
      Caption         =   "Oct:"
      Height          =   195
      Left            =   5535
      TabIndex        =   23
      Top             =   5025
      Width           =   300
   End
   Begin VB.Label lblHex 
      AutoSize        =   -1  'True
      Caption         =   "Hex:"
      Height          =   195
      Left            =   2895
      TabIndex        =   22
      Top             =   5025
      Width           =   330
   End
   Begin VB.Label lblDec 
      AutoSize        =   -1  'True
      Caption         =   "Dec:"
      Height          =   195
      Left            =   285
      TabIndex        =   21
      Top             =   5025
      Width           =   345
   End
   Begin VB.Label lblcode2chr 
      Caption         =   "Code to character                                                                                                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   20
      Top             =   4515
      Width           =   7320
   End
   Begin MSForms.TextBox txtOct 
      Height          =   1800
      Left            =   5505
      TabIndex        =   3
      Top             =   2250
      Width           =   2040
      VariousPropertyBits=   -1397733349
      ScrollBars      =   2
      Size            =   "3598;3175"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblOctCodes 
      AutoSize        =   -1  'True
      Caption         =   "Octal:"
      Height          =   195
      Left            =   5505
      TabIndex        =   19
      Top             =   2010
      Width           =   420
   End
   Begin MSForms.TextBox txtHex 
      Height          =   1800
      Left            =   2865
      TabIndex        =   2
      Top             =   2250
      Width           =   2040
      VariousPropertyBits=   -1397733349
      ScrollBars      =   2
      Size            =   "3598;3175"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblHexCodes 
      AutoSize        =   -1  'True
      Caption         =   "Hexadecimal:"
      Height          =   195
      Left            =   2865
      TabIndex        =   18
      Top             =   2010
      Width           =   960
   End
   Begin MSForms.TextBox txtDec 
      Height          =   1800
      Left            =   255
      TabIndex        =   1
      Top             =   2250
      Width           =   2040
      VariousPropertyBits=   -1397733349
      ScrollBars      =   2
      Size            =   "3598;3175"
      SpecialEffect   =   3
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblDecCodes 
      AutoSize        =   -1  'True
      Caption         =   "Decimal:"
      Height          =   195
      Left            =   255
      TabIndex        =   17
      Top             =   2010
      Width           =   615
   End
   Begin MSForms.TextBox txtChr 
      Height          =   1200
      Left            =   255
      TabIndex        =   0
      Top             =   690
      Width           =   7320
      VariousPropertyBits=   -1531951077
      ScrollBars      =   2
      Size            =   "12912;2117"
      SpecialEffect   =   3
      FontName        =   "Segoe UI"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblchr 
      AutoSize        =   -1  'True
      Caption         =   "Character(s):"
      Height          =   195
      Left            =   255
      TabIndex        =   16
      Top             =   435
      Width           =   900
   End
   Begin VB.Label lblchr2code 
      Caption         =   "Character to code                                                                                                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   14
      Top             =   120
      Width           =   7320
   End
   Begin MSForms.TextBox txtReplace 
      Height          =   1065
      Left            =   255
      TabIndex        =   12
      Top             =   9715
      Width           =   7320
      VariousPropertyBits=   -1397733349
      ScrollBars      =   2
      Size            =   "12912;1879"
      SpecialEffect   =   3
      FontName        =   "Segoe UI"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
'Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long

Private txtDec2_hasfocus As Boolean
Private txtHex2_hasfocus As Boolean
Private txtOct2_hasfocus As Boolean

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
RemoveMenu GetSystemMenu(Me.hwnd, 0), 2, &H400& 'prevent resizing
End Sub

Private Sub Form_Activate()
txtChr.SetFocus
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

Private Sub txtChr_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Call txtChr_Change
End Sub

Private Sub chkHTML_Click()
Call txtChr_Change
End Sub

Private Sub txtChr2_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub txtDec_GotFocus()
txtDec.SelStart = 0
txtDec.SelLength = Len(txtDec.Text)
End Sub

Private Sub txtHex_GotFocus()
txtHex.SelStart = 0
txtHex.SelLength = Len(txtHex.Text)
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtBlock_GotFocus()
txtBlock.SelStart = 0
txtBlock.SelLength = Len(txtBlock.Text)
End Sub

Private Sub txtPlane_GotFocus()
txtPlane.SelStart = 0
txtPlane.SelLength = Len(txtPlane.Text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtBlock_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPlane_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtOct_GotFocus()
txtOct.SelStart = 0
txtOct.SelLength = Len(txtOct.Text)
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

Private Sub txtDec2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii <> 127) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
KeyAscii = 0
End If
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

Private Sub txtHex2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii <> 127) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((KeyAscii < 65) Or (KeyAscii > 70)) And ((KeyAscii < 97) Or (KeyAscii > 102)) Then
KeyAscii = 0
End If
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

Private Sub txtOct2_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) And (KeyAscii <> 127) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
KeyAscii = 0
End If
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
