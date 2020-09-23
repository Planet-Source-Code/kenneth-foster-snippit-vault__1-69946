VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                           Snippit Vault ver3.0"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmNotes 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Notes"
      Height          =   3630
      Left            =   4320
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   6510
      Begin Project1.ccXPButton cmdClearNote 
         Height          =   480
         Left            =   5625
         TabIndex        =   26
         Top             =   1785
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   847
         Caption         =   "Delete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdSaveNote 
         Height          =   705
         Left            =   5625
         TabIndex        =   24
         Top             =   2385
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1244
         Caption         =   "Save"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdCloseNote 
         Height          =   315
         Left            =   5625
         TabIndex        =   23
         Top             =   240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         Caption         =   "Close"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9ECED&
         Height          =   2850
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   240
         Width           =   5400
      End
      Begin VB.Shape Shape4 
         Height          =   3630
         Left            =   0
         Top             =   0
         Width           =   6510
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To use: Type in notes and press Save."
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   195
         TabIndex        =   27
         Top             =   3135
         Width           =   2895
      End
   End
   Begin VB.ListBox lstTitles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E9ECED&
      Height          =   6660
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   3555
   End
   Begin RichTextLib.RichTextBox rtbCode 
      Height          =   6660
      Left            =   3630
      TabIndex        =   2
      Top             =   30
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   11748
      _Version        =   393217
      BackColor       =   15330541
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":030A
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   15
      ScaleHeight     =   1470
      ScaleWidth      =   11925
      TabIndex        =   0
      Top             =   6780
      Width           =   11955
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10365
         TabIndex        =   29
         Top             =   105
         Width           =   1425
         Begin Project1.ccXPButton cmdPrintSeld 
            Height          =   315
            Left            =   150
            TabIndex        =   31
            Top             =   585
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "Selected"
            ForeColor       =   49152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
         Begin Project1.ccXPButton cmdPrint 
            Height          =   315
            Left            =   150
            TabIndex        =   30
            Top             =   210
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "All"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
      End
      Begin Project1.LEDDisplaySTO Ctr1 
         Height          =   240
         Left            =   870
         TabIndex        =   28
         Top             =   690
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   423
         DigitCount      =   4
         LeadingZeros    =   -1  'True
         BorderColor     =   0
      End
      Begin Project1.ccXPButton cmdShowNotes 
         Height          =   390
         Left            =   105
         TabIndex        =   25
         Top             =   1050
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   688
         Caption         =   "Show Notes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Filename"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   1635
         TabIndex        =   17
         Top             =   105
         Width           =   3615
         Begin Project1.ccXPButton cmdDelete 
            Height          =   360
            Left            =   2430
            TabIndex        =   20
            Top             =   570
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   635
            Caption         =   "Delete"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
         Begin Project1.ccXPButton cmdSave 
            Height          =   360
            Left            =   90
            TabIndex        =   19
            Top             =   570
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   635
            Caption         =   "Save"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
         Begin VB.TextBox txtSave 
            Appearance      =   0  'Flat
            BackColor       =   &H00E9ECED&
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Top             =   240
            Width           =   3420
         End
      End
      Begin Project1.ccXPButton cmdExit 
         Height          =   315
         Left            =   10530
         TabIndex        =   12
         Top             =   1125
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "Exit"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Project1.ccXPButton cmdClear 
         Height          =   450
         Left            =   105
         TabIndex        =   11
         Top             =   75
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   794
         Caption         =   "Clear"
         ForeColor       =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1260
         Left            =   5370
         TabIndex        =   5
         Top             =   105
         Width           =   3075
         Begin Project1.ccXPButton cmdFind 
            Height          =   405
            Left            =   2355
            TabIndex        =   16
            Top             =   165
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   714
            Caption         =   "Find"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
         Begin Project1.ccXPButton cmdFindNext 
            Height          =   405
            Left            =   2355
            TabIndex        =   15
            Top             =   165
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   714
            Caption         =   "Next"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            BackColor       =   &H00E9ECED&
            Height          =   285
            Left            =   795
            TabIndex        =   8
            Top             =   675
            Width           =   2130
         End
         Begin VB.TextBox FindMe 
            Appearance      =   0  'Flat
            BackColor       =   &H00E9ECED&
            Height          =   285
            Left            =   735
            TabIndex        =   6
            Top             =   225
            Width           =   1590
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000009&
            X1              =   15
            X2              =   3060
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   -165
            X2              =   3045
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Type in letter or word and press Enter"
            Height          =   210
            Left            =   285
            TabIndex        =   10
            Top             =   975
            Width           =   2715
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Code List:"
            Height          =   210
            Left            =   60
            TabIndex        =   9
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Code in Window:"
            Height          =   375
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Send to Clipboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1260
         Left            =   8565
         TabIndex        =   4
         Top             =   105
         Width           =   1665
         Begin Project1.ccXPButton cmdSelected 
            Height          =   390
            Left            =   180
            TabIndex        =   14
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   688
            Caption         =   "Selected"
            ForeColor       =   49152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
         Begin Project1.ccXPButton cmdSelAll 
            Height          =   405
            Left            =   165
            TabIndex        =   13
            Top             =   225
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   714
            Caption         =   "All"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FocusRect       =   0   'False
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "See Notes"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1605
         TabIndex        =   32
         Top             =   1185
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   1020
         Left            =   10350
         Top             =   90
         Width           =   1485
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   1050
         Left            =   1605
         Top             =   90
         Width           =   3675
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1305
         Left            =   8550
         Top             =   90
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1305
         Left            =   5355
         Top             =   90
         Width           =   3135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9ECED&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Files"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   1
         Top             =   690
         Width           =   780
      End
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   1140
      Picture         =   "Form1.frx":038C
      Top             =   8610
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
   '*  Snippit Vault ver.3.0.4
   '*      By Ken Foster
   '*            2006
   '*  Freeware---Use any way you want
   '*  Bits and pieces of this code are from PSC
   '*  Thanks to the authors for all their efforts
   '**********************************************
   Option Explicit
   
   Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
   Const LB_FINDSTRING = &H18F
   Dim FileCount As Integer

Private Sub Form_Load()
   LoadList
    'puts frame captions in top center of controls
    Call AlignButtonText(Frame1, bsCenter, bsCenter)
    Call AlignButtonText(Frame2, bsCenter, bsCenter)
    Call AlignButtonText(Frame3, bsCenter, bsVcenter)
     Call AlignButtonText(Frame4, bsCenter, bsVcenter)
End Sub

Private Sub Form_Resize()
   If Form1.WindowState = 1 Then Exit Sub           'prevents error on minimizing form
   rtbCode.Height = lstTitles.Height + 20
   rtbCode.Width = Form1.Width - lstTitles.Width - 230
   lstTitles.Height = rtbCode.Height + 30
   'position and size controls window
   Picture1.Top = lstTitles.Height + 100
   Picture1.Width = Form1.Width - 170
   TileBackground Picture1, Image1
   'keeps shape rect in position with frames
   Shape1.Top = Frame2.Top - 10
   Shape1.Left = Frame2.Left - 23
   Shape2.Top = Frame1.Top - 10
   Shape2.Left = Frame1.Left - 23
   Shape3.Top = Frame3.Top - 10
   Shape3.Left = Frame3.Left - 23
   Shape4.Top = 0
   Shape4.Left = 0
   Shape5.Top = Frame4.Top - 10
   Shape5.Left = Frame4.Left - 23
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub LoadList()
   Dim strPath As String
   Dim strMapName As String
   Dim fStg As String
   Dim fSLen As Integer
   Dim FC As String
   
   lstTitles.Clear
   strPath = Dir(App.Path & "\SnippitFolder" & "\*.rtf")
   FileCount = 0
   If Not strPath = "" Then                                'yes, there are files here so
   Do                                                             'go get them
      strMapName = strPath
      fSLen = Len(strMapName) - 4                    'filename length minus extension
      fStg = Mid$(strMapName, 1, fSLen)            'filename without extension
      lstTitles.AddItem fStg                                'put filename into listbox
      FC = Right$(fStg, 1)
      If FC <> "-" Then FileCount = FileCount + 1  'if file name ends with "-" then do not count as a file
      strPath = Dir$
   Loop Until strPath = ""
Else
   MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
End If
'Label2.Caption = "Total Files   " & FileCount       'update file count window
Ctr1.Value = FileCount
End Sub

Private Sub lstTitles_Click()
   txtSave.Text = lstTitles.Text
   Call RTB_Load
   txtNote.Text = ""
   If frmNotes.Visible = True Then frmNotes.Visible = False
End Sub

Public Sub List_Remove(list As ListBox)
   
   On Error Resume Next
   If list.ListCount < 0 Then Exit Sub
   list.RemoveItem list.ListIndex
End Sub

Private Sub RTB_Save()
   Dim fFile As Integer
   
   fFile = FreeFile
   Open App.Path & "\SnippitFolder\" & txtSave & ".rtf" For Output As fFile     'For Output As 1 ---To overwrite file
   Print #fFile, rtbCode.Text                                                                         ' String location you want To save
   Close fFile
End Sub

Private Sub RTB_Load()
   
   Dim FileLength As Integer
   Dim var1 As String
   Dim fFile As Integer
   
   fFile = FreeFile
   If lstTitles.ListIndex = -1 Then Exit Sub                      'No item selected
   
   rtbCode.Text = ""
   Open App.Path & "\SnippitFolder\" & lstTitles & ".rtf" For Input As #fFile
   FileLength = LOF(fFile)
   var1 = Input(FileLength, #fFile)
   Colorize rtbCode, var1                                               'Color the code
   rtbCode.SelStart = 0                                                 'Puts Beginning of code at top
   Close #fFile
   txtSearch.Text = ""                                                   'clear search text window
   'if notes are available then show
   If Dir(App.Path & "\SnippitFolder\" & lstTitles & ".txt") <> "" Then
      Label1.Visible = True
   Else
      Label1.Visible = False
   End If
End Sub

Private Sub rtbCode_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If rtbCode.SelLength > 0 Then FindMe.Text = rtbCode.SelText
   'makes the find dialog pick up selected text - like vb
End Sub

Private Sub File_Delete(TList As ListBox)
   If TList = "" Then Exit Sub
   Kill App.Path & "\SnippitFolder\" & TList & ".rtf"
   rtbCode.Text = ""
End Sub

Private Sub cmdClear_Click()
   On Error Resume Next
   
   rtbCode.Text = ""
   cmdClear.Enabled = True
   txtSearch.Text = ""                                            ' clear search text window
   txtSave.Text = ""
   lstTitles.ListIndex = -1                                        'clear selection bar
End Sub

Private Sub cmdCloseNote_Click()
   frmNotes.Visible = False
   cmdShowNotes.Caption = "Show Notes"
End Sub

Private Sub cmdDelete_Click()
   On Error Resume Next
   Dim location As String
   Dim fileexists As String
   Dim iResponse As Integer
   Dim lgstg As String
   
   If txtSave.Text = "" And lstTitles.ListIndex = -1 Then
      txtSave.Text = "Make Selection"
      MsgBox " Please make a selection.", vbInformation, "Make a Selection"
      txtSave.Text = ""
      Exit Sub
   End If
   'file exists or not
   location = App.Path & "\SnippitFolder\" & txtSave.Text & ".txt"
   fileexists = Dir$(location) <> ""
   
   iResponse = MsgBox("Are you sure ?", vbYesNo, "Delete this file.")
   If iResponse = 7 Then Exit Sub                                        'no was selected
   Call File_Delete(lstTitles)
   Call List_Remove(lstTitles)
   lgstg = Right$(txtSave.Text, 1)                                          'if filename ends with "-" then its not a file so don't count it
   If lgstg <> "-" Then
      FileCount = FileCount - 1
      If fileexists = False Then GoTo skiptohere                      'if file does'nt exist then there is nothing to delete
      Kill App.Path & "\SnippitFolder\" & txtSave.Text & ".txt"  'remove note file associated with rtb file
skiptohere:
     ' Label2.Caption = "Total Files   " & FileCount                   'update file count window
      Ctr1.Value = FileCount
   End If
   txtSave.Text = ""                                                             'clear text in window
   frmNotes.Visible = False                                                  'if note window is still open then close it
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
   If FindMe.Text = "" Then
      MsgBox "Enter a word to search for !", vbExclamation, "Nothing to search for !"
      Exit Sub
   End If
   MousePointer = 11
   rtbCode.HideSelection = True
   'tells the user your busy
   Call Findit(rtbCode, FindMe.Text)
   'starts the find task in the module
   If foundit = False Then
      MsgBox "Can't find it !"
      cmdFind.Visible = True
      FindMe.Text = ""
      MousePointer = 0
      Exit Sub
      'inform the user
   Else
      cmdFindNext.Enabled = True
      'we found it - the user might want to find the
      'next occurrence so enable the 'Next' button
   End If
   MousePointer = 0
   'lets the user know we're finished
   cmdFind.Visible = False
End Sub

Private Sub cmdFindNext_Click()
   MousePointer = 11
   'tells the user your busy
   Call FinditNext(rtbCode, FindMe.Text)
   'starts the find next task in the module
   MousePointer = 0
   'lets the user know we're finished
   If foundit = False Then
      MsgBox "Search Complete !"
      cmdFind.Visible = True
      cmdFindNext.Enabled = False
      FindMe.Text = ""
      rtbCode.SelStart = 0
   End If
End Sub

Private Sub cmdPrint_Click()
   If rtbCode.Text = "" Then Exit Sub
   ' Print the contents of the RichTextBox with a one inch margin
   PrintRTF rtbCode, 1440, 1440, 1440, 1440
End Sub

Private Sub cmdPrintSeld_Click()
   Dim SelCode As String

   Printer.Print
   Printer.Print
   SelCode = rtbCode.SelText
   Printer.Print SelCode
   Printer.EndDoc
End Sub

Private Sub cmdSave_Click()
   Dim iResponse As Integer
   Dim location As String
   Dim fileexists As String
   
   If txtSave.Text = "" Then
      MsgBox "Please enter a Code Name.", vbInformation, "Need a Name"
      Exit Sub
   End If
   'check if file exists or not
   location = App.Path & "\SnippitFolder\" & txtSave.Text & ".rtf"
   fileexists = Dir$(location) <> ""
   
   If fileexists = False Then
      lstTitles.AddItem txtSave.Text
      RTB_Save                                                      'save Snippit code
      LoadList
     ' Label2.Caption = "Total Files   " & FileCount     'update file count window
      Ctr1.Value = FileCount
      txtSave.Text = ""
   Else
      iResponse = MsgBox("File already exists!" & Chr$(10) & "Do you want to overwrite file ?", vbYesNo, "File Exists!")
      If iResponse = 7 Then                                    'no was selected
      txtSave.Text = ""
      lstTitles.ListIndex = -1                                 'clear selection bar
      Exit Sub
   Else
      RTB_Save                                                  'overwrite existing file
     ' Label2.Caption = "Total Files   " & FileCount 'update file count window
      Ctr1.Value = FileCount
      txtSave.Text = ""
      lstTitles.ListIndex = -1                                 'clear selection bar
   End If
End If
End Sub

Private Sub cmdSaveNote_Click()
   If txtNote.Text = "" Then Exit Sub
   WriteFile App.Path & "\SnippitFolder\" & txtSave.Text & ".txt", txtNote.Text  'save note
   MsgBox "Note Saved", vbOKOnly, "Note Saved"
   txtNote.Text = ""
   frmNotes.Visible = False
   cmdShowNotes.Caption = "Show Notes"
End Sub

Private Sub cmdSelAll_Click()
   If rtbCode.Text = "" Then Exit Sub             'lf button is pressed and no code in window
   rtbCode.HideSelection = False
   rtbCode.SelStart = 0
   rtbCode.SelLength = Len(rtbCode.Text)
   Clipboard.Clear                                        'makes sure clipboard is empty
   Clipboard.SetText rtbCode.Text
End Sub

Private Sub cmdSelected_Click()
   If rtbCode.Text = "" Then Exit Sub                'if button is pressed and no code in window
   If rtbCode.SelLength = 0 Then
      MsgBox "Nothing selected"
      Exit Sub
   End If
   Clipboard.Clear
   Clipboard.SetText rtbCode.SelText
End Sub

Private Sub cmdShowNotes_Click()
   On Error Resume Next
   Dim lgstg As String
   
   lgstg = Right$(txtSave.Text, 1)                              'if filename ends with "-" then its not a file so don't show notes window
   If lgstg = "-" Then Exit Sub
   If txtSave.Text = "" Then Exit Sub                         ' if empty the leave sub
   frmNotes.Visible = Not frmNotes.Visible
   If cmdShowNotes.Caption = "Show Notes" Then     ' toggle caption
      cmdShowNotes.Caption = "Hide Notes"
   Else
      cmdShowNotes.Caption = "Show Notes"
   End If
   frmNotes.Caption = "Notes for " & txtSave.Text
   txtNote.Text = ReadFile(App.Path & "\SnippitFolder\" & txtSave.Text & ".txt")  ' load note into textbox
End Sub

Private Sub cmdClearNote_Click()
    If txtNote.Text = "" Then Exit Sub
    txtNote.Text = ""
    DeleteFile App.Path & "\SnippitFolder\" & txtSave.Text & ".txt"
    MsgBox "Note Deleted", vbOKOnly, "Note Deleted"
    txtNote.Text = ""
    frmNotes.Visible = False
    cmdShowNotes.Caption = "Show Notes"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      lstTitles.ListIndex = SendMessage(lstTitles.hWnd, LB_FINDSTRING, -1, ByVal CStr(txtSearch.Text))
   End If
End Sub

Private Sub TileBackground(Container As Object, PictureSrc As Object, Optional StartY As Long = 0)
   On Error Resume Next
   Dim lngXArea As Single, lngYArea As Single
   Container.AutoRedraw = True
   For lngXArea = 0 To Container.ScaleWidth Step PictureSrc.Width
      For lngYArea = StartY To Container.ScaleHeight _
         Step PictureSrc.Height
         Container.PaintPicture PictureSrc.Picture, lngXArea, lngYArea, PictureSrc.Width, PictureSrc.Height
      Next lngYArea
   Next lngXArea
End Sub

Private Function ReadFile(strPath As String) As Variant
   On Error Resume Next
   Dim iFileNumber As Integer
   Dim blnOpen As Boolean
   iFileNumber = FreeFile
   Open strPath For Input As #iFileNumber
   ReadFile = Input(LOF(iFileNumber), iFileNumber)
   Close #iFileNumber
End Function

Private Function WriteFile(strPath As String, strValue As String) As Boolean
   On Error GoTo eHandler
   Dim iFileNumber As Integer
   Dim blnOpen As Boolean
   iFileNumber = FreeFile
   Open strPath For Output As #iFileNumber
   blnOpen = True
   Print #iFileNumber, strValue
eHandler:
   If blnOpen Then Close #iFileNumber
   If Err Then
      MsgBox Err.Description, vbOKOnly + vbExclamation, Err.Number & " - " & Err.Source
   Else
      WriteFile = True
   End If
End Function
