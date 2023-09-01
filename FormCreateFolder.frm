VERSION 5.00
Begin VB.Form FormCreateFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bug Information - Create"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3015
   Icon            =   "FormCreateFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All (Start Again)"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton cmdGenerateTXTFiles 
      Caption         =   "Empty Text Files Create"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdImageCreate 
      Caption         =   "Image Folder Create"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtBugName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdBugCreate 
      Caption         =   "Bug Folder Create"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "FormCreateFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''''''''''
''' Author: Laim McKenzie
''' Date: 20 June 2013
'''
''' Course: Higher Computing
''''''''''''''''''''''''''''
Private Sub cmdBugCreate_Click()
If txtBugName.Text = "" Then
MsgBox ("Please enter a bug name.")
Else
MkDir (App.Path & "\data\animals\" & txtBugName.Text)
cmdImageCreate.Enabled = True
cmdBugCreate.Enabled = False
txtBugName.Enabled = False
End If

End Sub

Private Sub cmdClear_Click()
txtBugName.Enabled = True
txtBugName.Text = ""
cmdBugCreate.Enabled = True
cmdImageCreate.Enabled = False
cmdGenerateTXTFiles.Enabled = False
cmdClear.Enabled = False
End Sub

Private Sub cmdGenerateTXTFiles_Click()
Call CreateAfile
cmdGenerateTXTFiles.Enabled = False
cmdClear.Enabled = True
End Sub

Private Sub cmdImageCreate_Click()
MkDir (App.Path & "\data\animals\" & txtBugName.Text & "\image")
cmdImageCreate.Enabled = False
cmdGenerateTXTFiles.Enabled = True
End Sub

Sub ImageCreate()
' Declare variables.
   Dim CX, CY, Limit, Radius   As Integer
   ScaleMode = vbPixels   ' Set scale to pixels.
   AutoRedraw = True ' Turn on AutoRedraw.
   CX = ScaleWidth / 1   ' Set X position.
   CY = ScaleHeight / 1   ' Set Y position.
   Limit = CX   ' Limit size of circles.
   For Radius = 0 To Limit   ' Set radius.
      'Circle (CX, CY), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
      DoEvents   ' Yield for other processing.
   Next Radius
   SavePicture Image, "data\animals\" & txtBugName.Text & "\image\" & txtBugName.Text & ".JPG"   ' Save picture to file.
End Sub

Private Sub Form_Load()
cmdImageCreate.Enabled = False
cmdGenerateTXTFiles.Enabled = False
cmdClear.Enabled = False
End Sub

Sub CreateAfile()
Call ImageCreate
''FILE 1 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\antenna.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 1 END''
''FILE 2 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\class.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 2 END''
''FILE 3 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\color.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 3 END''
''FILE 4 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\common name.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 4 END''
''FILE 5 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\legs.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 5 END''
''FILE 6 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\name.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 6 END''
''FILE 7 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\shape.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 7 END''
''FILE 8 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\size.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 8 END''
''FILE 9 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\wings.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 9 END''
''FILE 10 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtBugName.Text & "\location.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 10 END''
End Sub
