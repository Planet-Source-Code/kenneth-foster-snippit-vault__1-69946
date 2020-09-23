VERSION 5.00
Begin VB.UserControl LEDDisplaySTO 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ScaleHeight     =   2565
   ScaleWidth      =   2460
   Begin VB.PictureBox SegDigits 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   2010
      Picture         =   "SolidTinyOnly.ctx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   60
      Width           =   150
   End
   Begin VB.PictureBox SegmentDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   0
      Top             =   30
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "LEDDisplaySTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const SRCCOPY   As Long = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const m_def_Value = 0
Const m_def_DigitCount = 2
Const m_def_LeadingZeros = False
Const m_def_BorderColor = vbWhite


Dim Redo As ParentControls
Dim m_Value As Integer
Dim m_DigitCount As Integer
Dim m_LeadingZeros As Boolean
Dim m_BorderColor As OLE_COLOR

Dim i As Integer
Event Change()
Event Click()
Event DblClick()

Private Sub SegmentDisplay_Click()
   RaiseEvent Click
End Sub

Private Sub SegmentDisplay_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
   m_Value = m_def_Value
   m_DigitCount = m_def_DigitCount
   m_LeadingZeros = m_def_LeadingZeros
   m_BorderColor = m_def_BorderColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Value = PropBag.ReadProperty("Value", m_def_Value)
   m_DigitCount = PropBag.ReadProperty("DigitCount", m_def_DigitCount)
   m_LeadingZeros = PropBag.ReadProperty("LeadingZeros", m_def_LeadingZeros)
   m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
   Value = m_Value
   DigitCount = m_DigitCount
   LeadingZeros = m_LeadingZeros
   BorderColor = m_BorderColor
   UserControl_Resize
End Sub

Private Sub UserControl_Resize()
   SegmentDisplay.Picture = SegmentDisplay.Picture
   SegmentDisplay.Width = DigitCount * 150
   SegmentDisplay.Height = 200
   Shape1.Width = SegmentDisplay.Width + 40
   UserControl.Width = SegmentDisplay.Width + 40
   UserControl.Height = SegmentDisplay.Height + 40
   Shape1.Height = SegmentDisplay.Height + 40
   DrawLED
   SegmentDisplay.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
   Call .WriteProperty("Value", m_Value, m_def_Value)
   Call .WriteProperty("DigitCount", m_DigitCount, m_def_DigitCount)
   Call .WriteProperty("LeadingZeros", m_LeadingZeros, m_def_LeadingZeros)
   Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
   End With
End Sub

Private Sub DrawLED()
Dim strValue As String

   If Len(Str$(Value)) - 1 > DigitCount Then
      Value = 0                              'set display to zero
      'DigitCount = Len(Str$(Value)) - 1     'set number of digits showing to match value
   End If
   
On Error Resume Next
   RaiseEvent Change
   If m_LeadingZeros = False Then
      strValue = strValue + Trim(String(DigitCount - (Len(Str$(Value)) - 1), "-")) + Trim(Value)
   Else
      strValue = strValue + Trim(String(DigitCount - (Len(Str$(Value)) - 1), "0")) + Trim(Value)
   End If
  
   For i = Len(strValue) To 1 Step -1
        BitBlt SegmentDisplay.hDC, DigitCount * 10 - i * 10, 0, 8, 15, SegDigits.hDC, 0, Mid$(strValue, Len(strValue) - i + 1, 1) * 15, SRCCOPY
   Next i
End Sub
Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
   On Error Resume Next
   m_BorderColor = NewBorderColor
   PropertyChanged "BorderColor"
   Shape1.BorderColor = m_BorderColor
   UserControl_Resize
End Property

Public Property Get DigitCount() As Integer
   DigitCount = m_DigitCount
End Property

Public Property Let DigitCount(ByVal NewDigitCount As Integer)
   On Error Resume Next
   m_DigitCount = NewDigitCount
   PropertyChanged "DigitCount"
   UserControl_Resize
End Property

Public Property Get LeadingZeros() As Boolean
   LeadingZeros = m_LeadingZeros
End Property

Public Property Let LeadingZeros(NewLeadingZeros As Boolean)
   m_LeadingZeros = NewLeadingZeros
   PropertyChanged "LeadingZeros"
   UserControl_Resize
End Property

Public Property Get Value() As Integer
   Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Integer)
   On Error Resume Next
   m_Value = NewValue
   PropertyChanged "Value"
   UserControl_Resize
End Property

