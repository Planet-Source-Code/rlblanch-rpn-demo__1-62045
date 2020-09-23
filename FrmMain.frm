VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "RPN"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TxtAnswr 
      Height          =   2655
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FrmMain.frx":0ECA
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox TxtRPNequation 
      Height          =   2655
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "FrmMain.frx":0ECC
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox TxtEquation 
      Height          =   2655
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "FrmMain.frx":0ECE
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExecute_Click()
  Dim strTmp     As String
  Dim strRPN()   As String
  Dim strEqSrc() As String
  Dim lRPNpntr   As Long
  Dim lX         As Long
  Dim lEqPntr    As Long
  Dim lLast      As Long

  lRPNpntr = 2
  lEqPntr = 1
  TxtAnswr.Text = vbNullString
  TxtRPNequation.Text = vbNullString
  strTmp = TxtEquation.Text
  lX = Len(strTmp)
  lLast = lX + 1
  Do While lX < lLast
    lLast = lX
    strTmp = Replace$(strTmp, vbNewLine, vbCr)
    strTmp = Replace$(strTmp, vbTab, " ")
    strTmp = Replace$(strTmp, "  ", " ")
    strTmp = Replace$(strTmp, vbNewLine, vbCr)
    strTmp = Replace$(strTmp, vbLf, vbNullString)
    lX = Len(strTmp)
  Loop
  strEqSrc = Split(strTmp, vbCr)
  For lEqPntr = LBound(strEqSrc) To UBound(strEqSrc)
    If LenB(Trim$(strEqSrc(lEqPntr))) > 0 Then
      Call Parser(strEqSrc(lEqPntr))
'      strTmp = vbNullString
'      For lX = LBound(gstrParsed) To glParsedSize
'        strTmp = strTmp & gstrParsed(lX) & " "
'      Next lX
'      Debug.Print strTmp
      Call ParsedEqn2RPNorder(strRPN(), lRPNpntr)
      strTmp = vbNullString
      For lX = LBound(strRPN()) To lRPNpntr
        strTmp = strTmp & strRPN(lX) & " "
      Next lX
      TxtRPNequation.Text = TxtRPNequation.Text & strTmp & vbNewLine
      TxtAnswr.Text = TxtAnswr.Text & CalcRPN(strRPN()) & vbNewLine
    End If
  Next lEqPntr
  CmdExit.SetFocus
End Sub

Private Sub CmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  gstrTokens(1) = "-=+*/|();'><^\" '"-=+*/|();'><[]"
  gstrTokens(2) = "AND"
  gstrTokens(3) = "OR"
  gstrTokens(4) = "XOR"
  gstrTokens(5) = "MOD"
  gstrTokens(6) = "<>"
  gstrTokens(7) = "><"
  gstrDoubleQuote = Chr$(34)      '...double quote character..
  'sample equations
  With TxtEquation
    .Text = "y= 1 + 5 - 1"
    .Text = .Text & vbNewLine & "y=25 * 5 + 2" & vbNewLine
    .Text = .Text & "y=((4+2)-(3+2))" & vbNewLine
    .Text = .Text & "y=10* .5" & vbNewLine
    .Text = .Text & "y = 2 ^ 8" & vbNewLine
    .Text = .Text & "y = -1 * 2" & vbNewLine
    .Text = .Text & "z=(((2+3)/(1-0.5))-((-1*2)*(2*3)))"
    .Text = .Text & "y = ( ( 4 * 3 + 6 ) * ( 3 - 4 ) )" & vbNewLine
    .Text = .Text & "y = ( 2 + 4 * 6 - ( 2 - 4 * ( 10 + 20 ) * 2 ) )" & vbNewLine
    .Text = .Text & "y = ( - 1 * 5 ) * ( - 1 )" & vbNewLine
    .Text = .Text & "y = ( - 1 - 5 )" & vbNewLine
    .Text = .Text & "y = - 1 - 5" & vbNewLine
    .Text = .Text & "y = ( - 1 * 5 )" & vbNewLine
    .Text = .Text & "y = 1 + 2 + 3 + 4 + 5" & vbNewLine
    .Text = .Text & "y = 1 - 2 + 3 + 4 - 5" & vbNewLine
    .Text = .Text & "y = 1 * 2 + 3 + 4 * 5" & vbNewLine
    .Text = .Text & "y = 2.5 / .5 * 10" & vbNewLine
    .Text = .Text & "y = 5 mod 2" & vbNewLine
    .Text = .Text & "5 + 7 mod 3" & vbNewLine
    .Text = .Text & "y = ( 5 + 7 + 1 ) mod 3" & vbNewLine
    .Text = .Text & "y = 5 and 8" & vbNewLine
    .Text = .Text & "y = 5 or 8" & vbNewLine
    .Text = .Text & "y = 1 * 2 + 3 + 5 or 8" & vbNewLine
    .Text = .Text & "y = ( 1 * 2 + 3 + 5 ) or 5" & vbNewLine
    .Text = .Text & "Y = &HFFFF AND 2 "
  End With 'TxtEquation
  FrmMain.Caption = FrmMain.Caption & "Execute All Rev. " & (Format$(App.Major, "00") & _
                    "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000"))
End Sub

':)Code Fixer V3.0.9 (8/2/2005 5:15:22 AM) 1 + 91 = 92 Lines Thanks Ulli for inspiration and lots of code.
