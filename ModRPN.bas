Attribute VB_Name = "ModRPN"
Option Explicit

Public glParsedSize              As Long
Public gstrTokens(1 To 7)        As String
Public gstrParsed(0 To 200)      As String
Public gstrDoubleQuote           As String * 1

'**************************************
' Name: CalcRPN
' Description: This function calculates a results from a RPN formula.
' By: Juha Mensola
'
' Inputs: The formula as string, separated by spaces.
'
' Returns: The end result as double
'
'Assumes: For those unfamiliar with RPN,
'     it is a notation somewhat different from
'     the standard way of writing down formulas.
'     For example the calculation (2+3)*(4+5)(= 45)
'     in RPN would be 2 3 + 4 5 + *.
'Also, this function understands multiple operands in the following manner:
'Formula 5*5+1*2*3/4-1 (= 25.5) would be 5 5 * 1 2 3 * * 4 / 1 - +
'The RemoveCell-function is used by the CalcRPN-function and should also be included in your project.
'The function currently understands the following perators: +, -, *, / and \(integer divide).
'But you can add new ones easily. Just add another case-statement and so on.
'
'Side Effects: None known.
'This code is copyrighted and has limited warranties.
'Please see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=24179&lngWId=1
'for details.
'**************************************
'Converting from infix to postfix notation
'The algorithm to convert from normal, infix, notation to postfix notation works as follows,
' converting an expression on the input stream into an equivalent one on the output stream.
' The expression on the input stream is examined symbol by symbol:
'
'
' Identifiers and constants are passed directly to the output stream, but operators and
'   parentheses require special attention.
' An opening parenthesis, '(', is simply placed on the operator stack.
' A closing parenthesis, ')', causes all of the operators on the stack down
'   to the matching '(' to be placed on the output stream, the operators being
'   removed from the top, down. The matching '(' is not placed on the output
'   stream, and neither is the ')' -- these are discarded.
' An operator on the input stream causes one of two actions to be taken:
'
'   If the stack is empty, or there is a '(' on the top of it, then the
'     operator is placed on the top of the stack.
'   If the top of the stack is an operator then if the current operator
'     has a higher priority value than the one on the top, again, it is placed on the top of the stack.
'   If, however, the operator has a priority which is lower than or equal
'     to that of the operator on the top of the stack (again assuming all
'     operators are left-associative), then the operator at the top of the stack is moved to the ouput stream.
'   The whole of the above process is repeated until the operator can be
'   placed on the top of the stack.
'
'This is only a simple introduction to Reverse Polish notation --
' more can be found in most books on compiling. We will not be using
' it any further as we will use the operator precedence method to parse
' expressions and will focus on tree-walking techniques when discussing code generation.
'
'article with animated compile
'http://www.spsu.edu/cs/faculty/bbrown/web_lectures/postfix/
'Converting Infix to Postfix
'
'
'We know that the infix expression (A+B)/(C-D) is equivalent to the postfix expression AB+CD-/.
'Let's convert the former to the latter.
'
'We have to know the rules of operator precedence in order to convert infix to postfix.
'The operations + and - have the same precedence. Multiplication and division, which we will
'represent as * and / also have equal precedence, but both have higher precedence than + and -.
'These are the same rules you learned in high school.
'
'We place a "terminating symbol"  after the infix expression to serve as a marker that we
'have reached the end of the expression. We also push this symbol onto the stack.
'
'After that, the expression is processed according to the following rules:
'
'Variables (in this case letters) are copied to the output
'
'Left parentheses are always pushed onto the stack
'
'When a right parenthesis is encountered, the symbol at the top of the stack is
'popped off the stack and copied to the output. Repeat until the symbol at the top
'of the stack is a left parenthesis. When that occurs, both parentheses are discarded.
'
'Otherwise, if the symbol being scanned has a higher precedence than the symbol at
'the top of the stack, the symbol being scanned is pushed onto the stack and thee
'scan pointer is advanced.
'
'If the precedence of the symbol being scanned is lower than or equal to the
'precedence of the symbol at the top of the stack, one element of the stack is
'popped to the output; the scan pointer is not advanced. Instead, the symbol being
'scanned will be complared with the new top element on the stack.
'
'When the terminating symbol is reached on the input scan, the stack is popped to
'the output until the terminating symbol is also reached on the stack. Then the
'algorithm terminates.
'
'If the top of the stack is a left parenthesis and the terminating symbol is scanned,
'or a right parenthesis is scanned when the terminating symbol is at the top of the stack,
'the parentheses of the original expression were unbalanced and an unrecoverable error has occurred.
Public Function CalcRPN(pstrStatements() As String) As Double
  Dim dOperand1         As Double
  Dim dOperand2         As Double
  Dim dResult           As Double
  Dim bMultipleOperands As Boolean
  Dim strOperator       As String
  Dim strTmp            As String
  Dim lCnt              As Long
  Dim lMultipleStartPos As Long
  Dim lX                As Long

  On Error GoTo ErrHandler
  lCnt = LBound(pstrStatements)
  Do Until False
'    strTmp = vbNullString
'    For lX = LBound(pstrStatements) To UBound(pstrStatements)
'      strTmp = strTmp & pstrStatements(lX) & " "
'    Next lX
'    Debug.Print strTmp
    If Not IsNumeric(pstrStatements(lCnt)) Then
      lCnt = lCnt - 2
    ElseIf Not IsNumeric(pstrStatements(lCnt + 1)) Then 'NOT NOT...
      lCnt = lCnt - 1
    ElseIf IsNumeric(pstrStatements(lCnt + 2)) Then 'NOT NOT...
      bMultipleOperands = True
      lMultipleStartPos = lCnt
      Do Until False
        If Not IsNumeric(pstrStatements(lCnt + 2)) Then Exit Do
        lCnt = lCnt + 1
      Loop
    End If
    dOperand1 = pstrStatements(lCnt)
    dOperand2 = pstrStatements(lCnt + 1)
    strOperator = pstrStatements(lCnt + 2)
    Select Case strOperator
      Case "+":   dResult = dOperand1 + dOperand2
      Case "-":   dResult = dOperand1 - dOperand2
      Case "*":   dResult = dOperand1 * dOperand2
      Case "/":   dResult = dOperand1 / dOperand2
      Case "^":   dResult = dOperand1 ^ dOperand2
      Case "\":   dResult = dOperand1 \ dOperand2
      Case "MOD": dResult = CLng(dOperand1) Mod CLng(dOperand2)
      Case "AND": dResult = CLng(dOperand1) And CLng(dOperand2)
      Case "OR":  dResult = CLng(dOperand1) Or CLng(dOperand2)
      Case "XOR": dResult = CLng(dOperand1) Xor CLng(dOperand2)
    End Select
    If bMultipleOperands Then
      pstrStatements(lCnt) = dResult
      Call RemoveCell(pstrStatements, lCnt + 1, 2)
      lCnt = lMultipleStartPos
      bMultipleOperands = False
    Else 'BMULTIPLEOPERANDS = FALSE/0
      pstrStatements(lCnt) = dResult
      Call RemoveCell(pstrStatements, lCnt + 1, 2)
      lCnt = lCnt + 1
      ' If UBound(pstrStatements) < 2 Then Exit Do
    End If
    If UBound(pstrStatements) < 2 Then Exit Do
  Loop
  CalcRPN = pstrStatements(LBound(pstrStatements))
Exit Function

ErrHandler:
  MsgBox "Error!"
  '        Resume Next
End Function

Private Function flOperPrecedence(ByVal pstrTst As String) As Long
  Select Case pstrTst
    Case "^":      flOperPrecedence = 8
    Case "AND":    flOperPrecedence = 7
    Case "OR":     flOperPrecedence = 6
    Case "XOR":    flOperPrecedence = 5
    Case "*", "/": flOperPrecedence = 4
    Case "\":      flOperPrecedence = 3
    Case "MOD":    flOperPrecedence = 2
    Case "+", "-": flOperPrecedence = 1
    Case "(":      flOperPrecedence = -1
    Case Else:     MsgBox "Unrecognized Operator: " & pstrTst
  End Select
End Function

Public Sub ParsedEqn2RPNorder(pstrRPN() As String, plRPNpntr As Long)
  Dim lX                 As Long
  Dim lY                 As Long
  Dim lParseSize         As Long
  Dim lStrt              As Long
  Dim lStackPntr         As Long
  Dim lTstPrec           As Long
  Dim lStckPrec          As Long
  Dim strTst             As String
  Dim strTmp             As String
  Dim strTmps            As String
  Dim strStack(1 To 100) As String
  Dim bRepeat            As Boolean
  Dim bUnaryNeg          As Boolean

  'test that there are matching parens
'  On Error GoTo HaveanError
  lParseSize = glParsedSize
  lStrt = LBound(gstrParsed()) + 2
  lStackPntr = 0
  plRPNpntr = 0
  bUnaryNeg = False
  'ReDim pstrRPN(1) As String
  For lX = lStrt To lParseSize
    strTst = gstrParsed(lX)
'    strTmp = ":OutStream: "
'    For lY = 1 To plRPNpntr
'      strTmp = strTmp & " " & Trim$(pstrRPN(lY))
'    Next lY
'    strTmps = ":Stack: "
'    For lY = 1 To lStackPntr
'      strTmps = strTmps & " " & Trim$(strStack(lY))
'    Next lY
'    Debug.Print ":TestOp: " & strTst
'    Debug.Print strTmp
'    Debug.Print strTmps
'    Debug.Print
    If strTst = "(" Then
      'push
      lStackPntr = lStackPntr + 1
      strStack(lStackPntr) = strTst
    ElseIf strTst = ")" Then 'NOT STRTST...
      'pop
      If strStack(lStackPntr) <> "(" Then
        plRPNpntr = plRPNpntr + 1
        ReDim Preserve pstrRPN(1 To plRPNpntr) As String
        pstrRPN(plRPNpntr) = strStack(lStackPntr)
      End If
      lStackPntr = lStackPntr - 1
      Do While strStack(lStackPntr) <> "("
        'pop
        plRPNpntr = plRPNpntr + 1
        ReDim Preserve pstrRPN(1 To plRPNpntr) As String
        pstrRPN(plRPNpntr) = strStack(lStackPntr)
        lStackPntr = lStackPntr - 1
        If lStackPntr = 0 Then Exit Do
      Loop
      lStackPntr = lStackPntr - 1
    Else 'NOT STRTST...
      If IsNumeric(strTst) Then
        plRPNpntr = plRPNpntr + 1
        ReDim Preserve pstrRPN(1 To plRPNpntr) As String
        If bUnaryNeg Then
          pstrRPN(plRPNpntr) = Trim$(Str$(-Val(strTst)))
          bUnaryNeg = False
        Else 'BUNARYNEG = FALSE/0
          pstrRPN(plRPNpntr) = strTst
        End If
      Else 'ISNUMERIC(STRTST) = FALSE/0
        bRepeat = True
        Do While bRepeat
          lTstPrec = flOperPrecedence(strTst)
          If lStackPntr = 0 Then
            If lX = lStrt Then
              bUnaryNeg = True
              Exit Do
            Else 'NOT LX...
              lStckPrec = 0
            End If
          Else 'NOT LSTACKPNTR...
            lStckPrec = flOperPrecedence(strStack(lStackPntr))
          End If
          'if lStckPrec negative then unary -
          If (lStckPrec < 0) Then
            If strTst = "-" Then
              If lX = lStrt Then
                bUnaryNeg = True
                Exit Do
              ElseIf gstrParsed(lX - 1) = "(" Then 'NOT LX...
                bUnaryNeg = True
                Exit Do
              End If
            ElseIf gstrParsed(lX - 1) = "(" Then 'NOT STRTST...
              MsgBox "Error"
            End If
          End If
          If lTstPrec <= lStckPrec Then
            'pop
            plRPNpntr = plRPNpntr + 1
            ReDim Preserve pstrRPN(1 To plRPNpntr) As String
            pstrRPN(plRPNpntr) = strStack(lStackPntr)
            lStackPntr = lStackPntr - 1
            bRepeat = True
          Else 'NOT LTSTPREC...
            'push
            lStackPntr = lStackPntr + 1
            strStack(lStackPntr) = strTst
            bRepeat = False
          End If
        Loop
      End If
    End If
  Next lX
  If lStackPntr > 0 Then
    'empty stack
    Do While lStackPntr > 0
      'pop
      plRPNpntr = plRPNpntr + 1
      ReDim Preserve pstrRPN(1 To plRPNpntr) As String
      pstrRPN(plRPNpntr) = strStack(lStackPntr)
      lStackPntr = lStackPntr - 1
    Loop
  End If
  Exit Sub

HaveanError:
  MsgBox "Error in Parseing equation to RPN."
  Resume Next
End Sub

Public Sub Parser(ByVal pstrString As String)
  Dim lStrPtr       As Long
  Dim lY            As Long
  Dim lLen          As Long
  Dim lStrt         As Long
  Dim bIgnoreString As Boolean
  Dim bIgnoreSpace  As Boolean
  Dim bInQuotes     As Boolean
  Dim bSkip         As Boolean
  Dim strVar        As String
  Dim strWork       As String
  Dim strTmp        As String
  Dim strTmp2       As String
  Dim strTmp3       As String
  glParsedSize = 0
  strWork = Trim$(pstrString)
  If InStr(strWork, "[") > 0 Then
    lLen = Len(strWork)
    lY = lLen + 1
    Do While lLen < lY
      lY = lLen
      strWork = Replace$(strWork, " [", "[")
      strWork = Replace$(strWork, "[ ", "[")
      strWork = Replace$(strWork, " ]", "]")
      lLen = Len(strWork)
    Loop
  End If
  'comments are defined by REM and '
  'a string is enclosed with "
  'exspected tokens =-+*/; <> >< and or xor mod
  'multi character tokens must end in a blank to identify their end
  'and not the beginning of a variable name
  glParsedSize = -1
  lStrPtr = InStr(strWork, "REM")
  If lStrPtr > 0 Then
    glParsedSize = glParsedSize + 1
    gstrParsed(glParsedSize) = "REM"
    glParsedSize = glParsedSize + 1
    gstrParsed(glParsedSize) = Mid$(strWork, lStrPtr + 3)
    Exit Sub
  End If
  If Left$(strWork, 1) = "'" Then
    glParsedSize = glParsedSize + 1
    gstrParsed(glParsedSize) = "'"
    glParsedSize = glParsedSize + 1
    gstrParsed(glParsedSize) = Mid$(strWork, lStrPtr + 3)
    Exit Sub
  End If
  'remove comments
  bIgnoreString = False
  lStrPtr = 0
  For lY = 1 To Len(strWork) 'find beginning of ' comments
    strTmp = Mid$(strWork, lY, 1)
    If strTmp = """" Then
      bIgnoreString = Not bIgnoreString
    ElseIf strTmp = "'" Then 'NOT STRTMP...
      If Not bIgnoreString Then lStrPtr = lY
      Exit For
    End If
  Next lY
  If lStrPtr > 0 Then
    'is  a ' comment
    'strComments = Mid$(strWork, lStrPtr + 1)
    '<:-) :WARNING: assigned only variable commented out
    strWork = Trim$(Left$(strWork, lStrPtr - 1))
  End If
  'ucase everything but what is between ""
  bInQuotes = False
  For lY = 1 To Len(strWork)
    strVar = Mid$(strWork, lY, 1)
    If Not bInQuotes Then
      If strVar <> """" Then
        Mid$(strWork, lY, 1) = UCase$(strVar)
      Else 'NOT STRVAR...
        bInQuotes = True
      End If
    ElseIf strVar = """" Then 'NOT NOT...
      bInQuotes = (Mid$(strWork, lY + 1, 1) = """")
    End If
  Next lY
  'start the parseing
  strVar = vbNullString
  strTmp = vbNullString
  bIgnoreSpace = False
  bIgnoreString = False
  For lStrPtr = 1 To Len(strWork)
    strTmp3 = Mid$(strWork, lStrPtr, 1)
    If (strTmp3 <> " ") Then
      If bIgnoreSpace Then
        If Len(Trim$(strVar)) > 0 Then
          bIgnoreSpace = Not bIgnoreSpace
          glParsedSize = glParsedSize + 1
          gstrParsed(glParsedSize) = Trim$(strVar)
        End If
        strVar = vbNullString
        bIgnoreSpace = False
      End If
    End If
    If strTmp3 = gstrDoubleQuote Then
      If Not bIgnoreString Then
        If Len(Trim$(strVar)) > 0 Then
          glParsedSize = glParsedSize + 1
          gstrParsed(glParsedSize) = Trim$(strVar)
          strVar = vbNullString
        End If
        strVar = strTmp3
        bIgnoreString = True
      ElseIf Mid$(strWork, lStrPtr + 1, 1) <> gstrDoubleQuote Then 'NOT NOT...
        strVar = strVar & strTmp3
        If Len(Trim$(strVar)) > 0 Then
          glParsedSize = glParsedSize + 1
          gstrParsed(glParsedSize) = Trim$(strVar)
          strVar = vbNullString
        End If
        bIgnoreString = False
      Else 'get the double double quote'NOT MID$(STRWORK,...
        strVar = strVar & strTmp3 & Mid$(strWork, lStrPtr + 1, 1)
        lStrPtr = lStrPtr + 1 '<:-) :WARNING: Modifies active For-Variable
      End If
    ElseIf bIgnoreString Then 'NOT STRTMP3...
      strVar = strVar & strTmp3
    ElseIf strTmp3 = " " Then 'BIGNORESTRING = FALSE/0
      If Len(Trim$(strVar)) > 0 Then
        glParsedSize = glParsedSize + 1
        gstrParsed(glParsedSize) = Trim$(strVar)
        strVar = vbNullString
      End If
      bIgnoreSpace = Not bIgnoreSpace
    ElseIf strTmp3 = "," Then ' Then make a token
      If Len(strVar) > 0 Then
        glParsedSize = glParsedSize + 1
        gstrParsed(glParsedSize) = Trim$(strVar)
        strVar = vbNullString
      End If
    Else 'NOT STRTMP3...
      bSkip = False
      For lY = 2 To UBound(gstrTokens())
        'multicharacter tokens will have a space before and after
        'if they are in the middle of a string
        lStrt = lStrPtr - 1
        If lStrPtr = 1 Then
          strTmp2 = gstrTokens(lY) & " "
          lLen = Len(gstrTokens(lY)) + 1
          lStrt = lStrPtr
        ElseIf Len(strWork) - lStrPtr = Len(gstrTokens(lY)) Then 'NOT LSTRPTR...
          strTmp2 = " " & gstrTokens(lY)
          lLen = Len(gstrTokens(lY)) + 1
        Else 'NOT LEN(STRWORK)...
          strTmp2 = " " & gstrTokens(lY) & " "
          lLen = Len(gstrTokens(lY)) + 2
        End If
        If Mid$(strWork, lStrt, lLen) = strTmp2 Then
          glParsedSize = glParsedSize + 1
          gstrParsed(glParsedSize) = Trim$(gstrTokens(lY))
          lStrPtr = lStrPtr + Len(gstrTokens(lY)) - 1 '1 is for the "if" adder'<:-) :WARNING: Modifies active For-Variable
          bSkip = True
        End If
      Next lY
      If Not bSkip Then
        If (InStr(gstrTokens(1), strTmp3) > 0) Then 'look for single char tokens
          If LenB(strVar) > 0 Then
            glParsedSize = glParsedSize + 1
            gstrParsed(glParsedSize) = Trim$(strVar)
            strVar = vbNullString
          End If
          glParsedSize = glParsedSize + 1
          gstrParsed(glParsedSize) = Trim$(strTmp3)
          strTmp2 = vbNullString
          lStrPtr = lStrPtr + Len(strTmp3) - 1 'skip past multichar token in strwork'<:-) :WARNING: Modifies active For-Variable
          strTmp3 = vbNullString
        Else 'NOT (INSTR(GSTRTOKENS(1),...
          strVar = strVar & strTmp3
        End If
      End If
    End If
  Next lStrPtr
  If Len(strVar) > 0 Then
    glParsedSize = glParsedSize + 1
    gstrParsed(glParsedSize) = Trim$(strVar)
  End If
  For lY = glParsedSize + 1 To glParsedSize + 6
    gstrParsed(lY) = vbNullString
  Next lY
End Sub

Private Sub RemoveCell(pstrArray() As String, _
                            plIndex As Integer, _
                            Optional plBlockLen As Long = 1)
    Dim lX     As Long
    Dim lY     As Long
    Dim lBnd   As Long
    
    lBnd = (plIndex + plBlockLen - 1)
    'remove array items from plIndex to plBlockLen
    lY = LBound(pstrArray)
    For lX = LBound(pstrArray) To UBound(pstrArray)
        If lX < plIndex Or lBnd < lX Then
            pstrArray(lY) = pstrArray(lX)
            lY = lY + 1
        End If
    Next
    lX = LBound(pstrArray)
    lY = UBound(pstrArray) - plBlockLen
    ReDim Preserve pstrArray(lX To lY)
End Sub

':)Code Fixer V3.0.9 (8/2/2005 5:15:24 AM) 5 + 638 = 643 Lines Thanks Ulli for inspiration and lots of code.

