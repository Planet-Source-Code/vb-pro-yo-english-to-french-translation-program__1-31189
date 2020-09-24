VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   120
   End
   Begin VB.TextBox temp 
      Height          =   1575
      Left            =   7800
      TabIndex        =   4
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8520
      Top             =   120
   End
   Begin RichTextLib.RichTextBox english 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3625
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox french 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5318
      _Version        =   393217
      TextRTF         =   $"Form1.frx":00D5
   End
   Begin VB.Line Line1 
      X1              =   9960
      X2              =   0
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label2 
      Caption         =   "French"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu filepulldown 
      Caption         =   "File"
      Begin VB.Menu newpulldown 
         Caption         =   "New"
      End
      Begin VB.Menu Savepulldown 
         Caption         =   "Save"
      End
      Begin VB.Menu loadpulldown 
         Caption         =   "Load"
      End
      Begin VB.Menu quitpulldown 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu optionspulldown 
      Caption         =   "Options"
      Begin VB.Menu makeworddocpulldown 
         Caption         =   "Make Word Doc"
      End
      Begin VB.Menu makenotepadfilepulldown 
         Caption         =   "Make Notepad File"
      End
      Begin VB.Menu translatetoenglishpulldown 
         Caption         =   "Translate To English"
      End
      Begin VB.Menu translatetofrenchpulldown 
         Caption         =   "Translate To French"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 Mark Belew
'this code was created by me.  It's not that complex
'but it does the job pretty good
'I tried to make it as easy to understand as
'possible because I hate downloading
'code that I don't understand.
'if change it I'd like to see what
'you did different my e-mail is:
'hackmastermark@hotmail.com

'all english to french words are stored in "words.txt"
'all french to english words are stored in "words2.txt"
Dim frenchwords(1 To 1000)      'all of the french words
Dim englishwords(1 To 1000)     'all of the english words
Dim lasttime As String          'holds what the text box last showed to prevent flicker
Dim howmanychr As Long          'no purpose really
Dim anystring As String         'current string used when translating
'Copyright 2002 Mark Belew
Private Sub Form_Load()
Form1.Show                      'shows the form
Dim whatnum As Integer          'holds how many words there are
whatnum = 1
Dim englishwo, frenchwo         'allows program to read file into two varibles
Open "words.txt" For Input As #1   ' Open file for input.
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, englishwo, frenchwo   ' Read data into two variables.
englishwords(whatnum) = englishwo   'puts current word into english array
frenchwords(whatnum) = frenchwo     'puts current word into french array
whatnum = whatnum + 1               'keeps count of how many words have been entered
Loop                                'reads next input
Close #1   ' Close file.
Timer1.Enabled = False                 'disables french to english
Timer2.Enabled = True                  'enables english to french
End Sub

Private Sub loadpulldown_Click()
'the following will someday allow you to use this
'program as a word processor including print this is
'the load sub I am going to in a future release use
'common dialog
nameoffile = InputBox("What is the name of the file?")
nameoffile = nameoffile + ".mab"
Open nameoffile For Input As #1   ' Open file for input.
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, a, b ' Read data into two variables.
english.Text = a
french.Text = b
Loop
Close #1   ' Close file.
End Sub

Private Sub makeworddocpulldown_Click()
'this saves text into word doc
frenchorenglish = InputBox("Which Language Do you wish to make into a Word Doc?")
If frenchorenglish = "french" Then
frenchorenglish = french.Text
End If
If frenchorenglish = "english" Then
frenchorenglish = english.Text
End If
nameoffile = InputBox("What do you want to call the file?")
nameoffile = nameoffile + ".doc"
Open nameoffile For Output As #1
Write #1, frenchorenglish
Close #1
End Sub

Private Sub newpulldown_Click()
lasttime = ""
howmanychr = "0"
anystring = ""
english.Text = ""
french.Text = ""
End Sub

Private Sub Savepulldown_Click()
'saves the file that can be opened later
FileName = InputBox("What do you want the file to be called?--do not include .mab--")
FileName = FileName + ".mab"
Open FileName For Output As #1
Write #1, english.Text
Write #1, french.Text
Close #1
End Sub

Private Sub Timer1_Timer()
'french to english translator VERY BAD
'right now I am just trying to do english
'into french
If french.Text = "" Then Exit Sub
If lasttime = french.Text Then Exit Sub
english.Text = ""
whereareweat = 1
lasttime = french.Text
anystring = french.Text   ' Define string.
For i = 1 To 13
anystring = Replace(anystring, frenchwords(i), englishwords(i) & " ")
Next i
For i = 1 To 1000
anystring = Replace(anystring, frenchwords(i), englishwords(i) & " ")
Next i
french.Text = anystring
End Sub

Private Sub Timer2_Timer()
'English to french translator VERY GOOD the first loop
'checks common phrases since they are sometimes said
'differently than just word for word
'the second check for verb conugations
If english.Text = "" Then Exit Sub
If lasttime = english.Text Then Exit Sub
french.Text = ""
whereareweat = 1
lasttime = english.Text
anystring = english.Text   ' Define string.
For i = 1 To 18
anystring = Replace(anystring, frenchwords(i), englishwords(i) & " ")
Next i
For i = 19 To 66
anystring = Replace(anystring, frenchwords(i), englishwords(i) & " ")
Next i
For i = 1 To 1000
anystring = Replace(anystring, englishwords(i), frenchwords(i) & " ")
'anystring = anystring & "  "
Next i
french.Text = anystring
End Sub

Private Sub translatetoenglishpulldown_Click()
Timer1.Enabled = True
Timer2.Enabled = False
whatnum = 1
Dim englishwo, frenchwo
Open "words2.txt" For Input As #1   ' Open file for input.
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, englishwo, frenchwo   ' Read data into two variables.
englishwords(whatnum) = englishwo
frenchwords(whatnum) = frenchwo
whatnum = whatnum + 1
Loop
Close #1   ' Close file.
End Sub

Private Sub translatetofrenchpulldown_Click()
Timer1.Enabled = False
Timer2.Enabled = True
whatnum = 1
Dim englishwo, frenchwo
Open "words.txt" For Input As #1   ' Open file for input.
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, englishwo, frenchwo   ' Read data into two variables.
englishwords(whatnum) = englishwo
frenchwords(whatnum) = frenchwo
whatnum = whatnum + 1
Loop
Close #1   ' Close file.
End Sub
