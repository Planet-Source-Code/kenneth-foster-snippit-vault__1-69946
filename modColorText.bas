Attribute VB_Name = "modColorize"
Option Explicit
   
   Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
   
   Const BlueKeyWords = "#Const*#Else*#ElseIf*#End If*#If*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"
   Const lBlueKeyWords = "#const*#else*#elseif*#end if*#if*alias*and*as*base*binary*boolean*byte*byval*call*case*cbool*cbyte*ccur*cdate*cdbl*cdec*cint*clng*close*compare*const*csng*cstr*currency*cvar*cverr*decimal*declare*defbool*defbyte*defcur*defdate*defdbl*defdec*defint*deflng*defobj*defsng*defstr*defvar*dim*do*double*each*else*elseif*end*enum*eqv*erase*error*exit*explicit*false*for*function*get*global*gosub*goto*if*imp*in*input*input*integer*is*lbound*let*lib*like*line*lock*long*loop*lset*name*new*next*not*object*on*open*option*or*output*print*private*property*public*put*random*read*redim*resume*return*rset*seek*select*set*single*spc*static*string*stop*sub*tab*then*then*true*type*ubound*unlock*variant*wend*while*with*xor*nothing*to*"
   
   Public bInQuotes As Boolean

Public Sub Colorize(rtb As RichTextBox, sText As String)
   If sText = "" Then Exit Sub
   DoEvents
   Screen.MousePointer = vbHourglass
   
   Dim arCode() As String
   Dim arSegment() As String
   Dim iLineCount As Integer
   Dim iSegment As Integer
   Dim bPartialComment As Boolean
   
   arCode = Split(sText, vbCrLf)
   
   With rtb
      
      LockWindowUpdate .hWnd
      
      '-- Loop through Each line of code
      For iLineCount = LBound(arCode) To UBound(arCode)
         DoEvents
         '-- check if line is blank
         If Len(Trim(arCode(iLineCount))) > 0 Then
            '-- Check if line is a comment
            If Left$(Trim(arCode(iLineCount)), 1) = "Rem " Or Left$(Trim(arCode(iLineCount)), 1) = "'" Then
               .SelColor = &HFF00FF     'QBColor(2) '-- Green
               .SelText = arCode(iLineCount) & vbCrLf
            Else
               '-- Split up each word in the line of code
               arSegment = Split(arCode(iLineCount), " ")
               '-- Examine each word for colorzing
               For iSegment = LBound(arSegment) To UBound(arSegment)
                  
                  If Left$(arSegment(iSegment), 1) = "'" Then
                     '-- Comment found in middle of test
                     If Not bInQuotes Or bPartialComment Then
                        .SelColor = &HFF00FF      'QBColor(2) '-- Green
                        .SelText = arSegment(iSegment) & " "
                        bPartialComment = True
                     Else
                        .SelText = arSegment(iSegment) & " "
                     End If
                     
                     '-- Space found Move on
                  ElseIf Left$(arSegment(iSegment), 1) = "" Then
                     .SelText = arSegment(iSegment) & " "
                     
                  Else
                     If bPartialComment Then
                        '-- Word is in a comment
                        .SelColor = &HFF00FF     'QBColor(2) '-- Green
                        .SelText = arSegment(iSegment) & " "
                     Else
                        If InStr(1, lBlueKeyWords, LCase(arSegment(iSegment))) And Not Len(arSegment(iSegment)) = 1 Then
                           If Not bInQuotes Then
                              '-- Word is a Keyword
                              .SelColor = vbBlue  'QBColor(1) '-- Dark Blue
                              '-- Fix Uppercase / Lowercase
                              .SelText = Mid$(BlueKeyWords, InStr(1, lBlueKeyWords, LCase(arSegment(iSegment))), Len(arSegment(iSegment))) & " "
                           Else
                              '-- Word is a Keyword but inside quotes
                              '-- Make it black
                              .SelText = arSegment(iSegment) & " "
                           End If
                        Else
                           '-- Word is not a Keyword or inside quotes or commented
                           If In_Quote(arSegment(iSegment)) Then Debug.Print "in quote"
                           .SelColor = QBColor(0) '-- Black
                           .SelText = arSegment(iSegment) & " "
                        End If
                     End If
                  End If
               Next iSegment
               If Not iLineCount = UBound(arCode) Then .SelText = vbCrLf
            End If
         Else
            'Line is Blank
            .SelText = vbCrLf
         End If
         bPartialComment = False
         bInQuotes = False
      Next iLineCount
      .SelColor = QBColor(0)
   End With
   LockWindowUpdate 0&
   Screen.MousePointer = vbDefault
End Sub

Private Function In_Quote(sSegment As String) As Boolean
   'Check for Quote State
   Dim pos As Integer
   Dim start As Integer
   start = 1
   pos = 1
   Do Until pos = 0
      pos = InStr(start, sSegment, Chr(34))
      If pos > 0 Then bInQuotes = Not bInQuotes
      start = pos + 1
      Loop
      In_Quote = bInQuotes
   End Function
