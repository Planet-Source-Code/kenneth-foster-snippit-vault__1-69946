Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'used to align control captions
Private Const BS_LEFT As Long = &H100
Private Const BS_RIGHT As Long = &H200
Private Const BS_CENTER As Long = &H300
Private Const BS_TOP As Long = &H400
Private Const BS_BOTTOM As Long = &H800
Private Const BS_VCENTER As Long = &HC00

Private Const BS_ALLSTYLES = BS_LEFT Or BS_RIGHT Or BS_CENTER Or BS_TOP Or BS_BOTTOM Or BS_VCENTER
Private Const GWL_STYLE& = (-16)

Public Enum bsHorizontalAlignments
    bsLeft = BS_LEFT
    bsright = BS_RIGHT
    bsCenter = BS_CENTER
End Enum

Public Enum bsVerticalAlignments
    bsTop = BS_TOP
    bsBottom = BS_BOTTOM
    bsVcenter = BS_VCENTER
End Enum


Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
   lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
   
   
Dim textfound As Integer    'used in subs as success indicator
Public foundit As Boolean   'lets the form know if search was successful
Public dontsave As Boolean  'stops you screwing files while you experiment
                          
Public Sub Findit(RTF As RichTextBox, FindMe As String)
    RTF.Find (FindMe)
    textfound = RTF.Find(FindMe)
    If textfound <> -1 Then
        foundit = True 'lets the form know we found it
        RTF.SetFocus 'OK we've found it so select it to show the user
        Exit Sub 'lets' get out of here
    Else
        foundit = False 'lets the form know we cant find it
        Exit Sub
    End If
End Sub

Public Sub FinditNext(RTF As RichTextBox, FindMe As String)
'the text has already been found once. We want the next occurrence
  textfound = RTF.Find(FindMe, RTF.SelStart + Len(FindMe))
    If textfound <> -1 Then 'Found it !
        foundit = True 'let the form know we succeeded
        RTF.SetFocus 'show it to the user
        Exit Sub 'lets bail out now
    Else
        foundit = False 'let the form know we cant find it
        Exit Sub 'lets bail out now
    End If
End Sub

Public Sub ReplaceitAll(RTF As RichTextBox, currentfile As String, FindMe As String, Replaceme As String)
    
    RTF.SetFocus
    RTF.SelStart = 0
    RTF.Find (FindMe) 'find it
    Do Until RTF.SelText = "" 'keep going till we cant find it
        RTF.SelText = Replaceme 'when we find it replace it
        RTF.Find (FindMe), RTF.SelStart + Len(FindMe) 'look for the next one
        If RTF.SelText = "" Then 'get out of the loop when done
            RTF.SelStart = 0
            RTF.SelLength = 0
            Exit Do
        End If
    Loop
    If dontsave = False Then RTF.SaveFile currentfile
    'set to true in normal operation or remove dontsave value
    Exit Sub

End Sub

Public Sub Replaceit(RTF As RichTextBox, currentfile As String, FindMe As String, Replaceme As String)
If RTF.SelText = FindMe Then RTF.SelText = Replaceme
If dontsave = False Then RTF.SaveFile currentfile
'put this in for thoroughness - but didn't bother calling it -
'did the same thing on the form itself
End Sub

Public Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)
   Dim LeftOffset As Long
   Dim LeftMargin As Long
   Dim RightMargin As Long
   Dim TopOffset As Long
   Dim TopMargin As Long
   Dim BottomMargin As Long
   Dim PrinterhDC As Long
   Dim r As Long

   ' Start a print job to initialize printer object
   Printer.Print Space(1)
   Printer.ScaleMode = vbTwips
   
   ' Get the left offset to the printable area on the page in twips
   LeftOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
   LeftOffset = Printer.ScaleX(LeftOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   LeftMargin = LeftMarginWidth - LeftOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   
   ' Calculate the line width
   PrintableWidth = RightMargin - LeftMargin
   
   ' Get the top offset to the printable area on the page in twips
   TopOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
   TopOffset = Printer.ScaleX(TopOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   TopMargin = TopMarginWidth - TopOffset
   BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffset
   
   ' Calculate the line width
   PrintableHeight = BottomMargin - TopMargin
    
   
   ' Create an hDC on the Printer pointed to by the Printer object
   ' This DC needs to remain for the RTF to keep up the WYSIWYG display
   PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

   ' Tell the RTF to base it's display off of the printer
   '    at the desired line width
   r = SendMessage(RTF.hWnd, EM_SETTARGETDEVICE, PrinterhDC, _
      ByVal PrintableWidth)

   ' Abort the temporary print job used to get printer info
   Printer.KillDoc
End Sub

' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins
' RTF - A RichTextBox control to print
' LeftMarginWidth - Width of desired left margin in twips
' TopMarginHeight - Height of desired top margin in twips
' RightMarginWidth - Width of desired right margin in twips
' BottomMarginHeight - Height of desired bottom margin in twips
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
'   parameter you should specify the same LeftMarginWidth and
'   RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, _
   TopMarginHeight, RightMarginWidth, BottomMarginHeight)
   Dim LeftOffset As Long, TopOffset As Long
   Dim LeftMargin As Long, TopMargin As Long
   Dim RightMargin As Long, BottomMargin As Long
   Dim fr As FormatRange
   Dim rcDrawTo As RECT
   Dim rcPage As RECT
   Dim TextLength As Long
   Dim NextCharPosition As Long
   Dim r As Long

   ' Start a print job to get a valid Printer.hDC
   Printer.Print Space(1)
   Printer.ScaleMode = vbTwips

   ' Get the offsett to the printable area on the page in twips
   LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETX), vbPixels, vbTwips)
   TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
      PHYSICALOFFSETY), vbPixels, vbTwips)

   ' Calculate the Left, Top, Right, and Bottom margins
   LeftMargin = LeftMarginWidth - LeftOffset
   TopMargin = TopMarginHeight - TopOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

   ' Set printable area rect
   rcPage.Left = 0
   rcPage.Top = 0
   rcPage.Right = Printer.ScaleWidth
   rcPage.Bottom = Printer.ScaleHeight

   ' Set rect in which to print (relative to printable area)
   rcDrawTo.Left = LeftMargin
   rcDrawTo.Top = TopMargin
   rcDrawTo.Right = RightMargin
   rcDrawTo.Bottom = BottomMargin

   ' Set up the print instructions
   fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
   fr.hdcTarget = Printer.hdc  ' Point at printer hDC
   fr.rc = rcDrawTo            ' Indicate the area on page to draw to
   fr.rcPage = rcPage          ' Indicate entire size of page
   fr.chrg.cpMin = 0           ' Indicate start of text through
   fr.chrg.cpMax = -1          ' end of the text

   ' Get length of text in RTF
   TextLength = Len(RTF.Text)

   ' Loop printing each page until done
   Do
      ' Print the page by sending EM_FORMATRANGE message
      NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
      If NextCharPosition >= TextLength Then Exit Do  'If done then exit
      fr.chrg.cpMin = NextCharPosition ' Starting position for next page
      Printer.NewPage                  ' Move on to next page
      Printer.Print Space(1) ' Re-initialize hDC
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc
   Loop

   ' Commit the print job
   Printer.EndDoc

   ' Allow the RTF to free up memory
   r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub

Public Sub AlignButtonText(cmd As Control, _
Optional ByVal HStyle As bsHorizontalAlignments = _
bsCenter, Optional ByVal VStyle As _
bsVerticalAlignments = bsVcenter)

    Dim oldStyle As Long
    ' retrieve the current style of the control
    oldStyle = GetWindowLong(cmd.hWnd, GWL_STYLE)
    ' change the style
    oldStyle = oldStyle And (Not BS_ALLSTYLES)
    ' set the style of the control to the new style
    Call SetWindowLong(cmd.hWnd, GWL_STYLE, _
    oldStyle Or HStyle Or VStyle)
    cmd.Refresh
End Sub
