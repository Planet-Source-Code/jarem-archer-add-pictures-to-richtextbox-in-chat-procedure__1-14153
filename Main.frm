VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Pictures to RichTextBox - Jarem Archer"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   285
      Left            =   5460
      TabIndex        =   2
      Top             =   4500
      Width           =   1470
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   4485
      Width           =   5325
   End
   Begin RichTextLib.RichTextBox rchText1 
      Height          =   3750
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   6615
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Main.frx":00C9
      Height          =   615
      Left            =   75
      TabIndex        =   3
      Top             =   30
      Width           =   6840
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'First, just add plain text:
rchText1.SelText = rchText1.SelText & "Guest: " & Text1 & vbCrLf

'Then change what needs to be changed to pictures:
RefreshPics
DoEvents
rchText1.SelStart = Len(rchText1.Text) 'Put the start at the end, thats where you want to add the next line
Text1 = ""
Call Text1.SetFocus
End Sub

Private Sub Form_Load()
rchText1.OLEObjects.Clear 'Clear the ole objects to prevent errors
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DoEvents
rchText1.OLEObjects.Clear 'You must add this or the
                            'program will crash. This doesnt
                            'happen in Windows 2k
DoEvents
End Sub

Sub RefreshPics() 'This scans the text for :) and (l)'s to change
Dim lFoundPos As Long           'Position of first character
                                        'of match
Dim lFindLength As Long         'Length of string to find

Dim MakeSure As Boolean 'I have this to do the procedure twice, just to "make sure"


GoTo Skip:
Start:
MakeSure = True
Skip:

lFoundPos = rchText1.Find(":)", 0, , rtfNoHighlight)
        While lFoundPos > 0
          rchText1.SelStart = lFoundPos
          'The SelLength property is set to 0 as
          'soon as you change SelStart
          rchText1.SelLength = 2
          rchText1.SelText = ""
          rchText1.OLEObjects.Add , , App.Path & "\smile.bmp" 'Add the picture after it has deleted the string
          DoEvents
          'Attempt to find the next match
          lFoundPos = rchText1.Find(sFindString, lFoundPos + 2, , rtfNoHighlight)
Wend
If MakeSure = False Then GoTo Start

' I guess by changing or adding a few lines, you
'   can make it add more pictures with different strings.

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub
