VERSION 5.00
Begin VB.Form ProViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox WebFeed 
      CausesValidation=   0   'False
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   5715
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6240
      Width           =   5775
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtDefault 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picMini 
      Height          =   1815
      Left            =   3600
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtMiniLocation 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtMiniName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "LOCATION: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label lblLocationInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "1.0.0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   7200
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "VERSION: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   3375
   End
   Begin VB.Label lblColorInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "COLOR: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label lblWingsInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label lblAntennaInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblClassInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label lblCommonNameInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label lblMiniShapeInfo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblMiniSizeAmount 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblWings 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "WINGS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label lblAntenna 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "ANTENNA:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "CLASS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label lblCommonName 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "COMMON NAME:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblMiniShape 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "SHAPE:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblMiniSize 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "SIZE:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblMiniLegsAmount 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblMiniLegs 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "LEGS:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "DEFAULT IMAGE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMiniTitle 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "MINI BEAST NAME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblMiniLocation 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "MINI BEAST NAME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblMiniName 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "MINI BEAST NAME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "ProViewer"
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
Dim BugName As String
Dim BugLegs As String
Dim BugSize As String
Dim BugShape As String
Dim BugCommonName As String
Dim BugClass As String
Dim BugAntenna As String
Dim BugWings As String
Dim BugColor As String
Dim BugLocation As String
Dim FormIcon As String
Dim objFSO, strFolder

Private Sub cmdClear_Click()
Call GeneralInformation
End Sub

Private Sub cmdSearch_Click()
Call OpenFiles
''Call Functions^^''
''General Code VV''
txtDefault.Text = txtMiniName.Text
lblMiniTitle = txtMiniName.Text
If txtMiniName.Text = "" Then
'' One does not simpy, code a then function.
Else
picMini.Picture = LoadPicture(App.Path & "\data\animals\" & txtMiniName.Text & "\image\" & txtMiniName.Text & ".JPG")
lblMiniTitle.Caption = BugName ''Bug Name
lblMiniLegsAmount.Caption = BugLegs ''Bug, Amount of Legs
lblMiniSizeAmount.Caption = BugSize ''Bug Size
lblMiniShapeInfo.Caption = BugShape ''Bug Shape, fucking bee shaped.
lblCommonNameInfo.Caption = BugCommonName ''Bug Common Name
lblClassInfo.Caption = BugClass ''Bug Class Name
lblAntennaInfo.Caption = BugAntenna ''Bug Antenna?
lblWingsInfo.Caption = BugWings ''Bug Wings
lblColorInfo.Caption = BugColor ''Bug Color
lblLocationInfo.Caption = BugLocation ''Bug Location
cmdSearch.Enabled = False
txtMiniName.Enabled = False
cmdClear.Enabled = True
End If
End Sub

Sub OpenFiles()
''FORM ICON START''
FormIcon = App.Path & "\data\DO_NOT_MODIFY\bug.ico"
''FORM ICON END''
''If Folder/File doesn't exsist, create it you dumb dumb. START ''
Call IfDoesntExsistCreateIt
''If Folder/File doesn't exsist, create it you dumb dumb. END ''
''BUG NAME START''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\name.txt") For Input As #1
Line Input #1, BugName
''BUG NAME FINISH''
'''
''BUG LEG COUNT START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\legs.txt") For Input As #2
Line Input #2, BugLegs
''BUG LEG COUNT END ''
'''
''BUG SIZE START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\size.txt") For Input As #3
Line Input #3, BugSize
''BUG SIZE END ''
'''
''BUG SHAPE START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\shape.txt") For Input As #4
Line Input #4, BugShape
''BUG SHAPE END ''
'''
''BUG COMMON NAME START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\common name.txt") For Input As #5
Line Input #5, BugCommonName
''BUG COMMON NAME END ''
'''
''BUG CLASS START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\class.txt") For Input As #6
Line Input #6, BugClass
''BUG CLASS END ''
'''
''BUG ANTENNA? START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\antenna.txt") For Input As #7
Line Input #7, BugAntenna
''BUG ANTENNA? END ''
'''
''BUG WINGS START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\wings.txt") For Input As #8
Line Input #8, BugWings
''BUG WINGS END ''
'''
''BUG COLOR START ''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\color.txt") For Input As #9
Line Input #9, BugColor
''BUG COLOR END ''
'''
''BUG LOCATION START''
Open (App.Path & "\data\animals\" & txtMiniName.Text & "\location.txt") For Input As #10
Line Input #10, BugLocation
''BUG LOCATION END''
End Sub
Private Sub Form_Load()
Call GeneralInformation

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Lazy Stuff for me:3'
Label1.Caption = picMini.Width & "    " & picMini.Height ''Tells me the WIDTH and HEIGHT of the picture box

End Sub
Sub GeneralInformation()
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7
Close #8
Close #9
Close #10
''''''''SETS FORM SETTINGS'''''''''''''''
'Height = 2310
ProViewer.Caption = "MiniBeast - LaimWMcKenzie - Higher Computing"

''''''''SETS NAME FOR 'OBJECTS'''''''''''
txtDefault.Text = "default" 'Ignore this, as I advise not changing it; will break the entire program. :(
''
txtMiniName.Text = ""
lblMiniName.Caption = "MINI BEAST NAME"
''
txtMiniLocation.Text = ""
lblMiniLocation.Caption = "MINIBEAST LOCATION"
''
cmdSearch.Caption = "Search For MiniBeast"
cmdClear.Caption = "Clear All"

''SETS NAME FOR BUG INFORMATION(s)'''''''
picMini.Picture = LoadPicture(App.Path & "\data\DO_NOT_MODIFY\" & txtDefault.Text & ".JPG")
lblMiniTitle.Caption = "" '' Bug Name, this will change during input.
'''
''LEGS START''
lblMiniLegs.Caption = "LEGS:" '' Bug Legs Label, just ignore this.
lblMiniLegsAmount.Caption = "" ''Bug Legs Count, this will change during input.
''LEGS FINISH''
'''
''SIZE START''
lblMiniSize.Caption = "SIZE:" '' Bug Size Label, just ignore this.
lblMiniSizeAmount.Caption = "" ''Bug Size Count, this will change during input.
''SIZE FINISH''
'''
''SHAPE START''
lblMiniShape.Caption = "SHAPE:" '' Bug Shape Label, just ignore this.
lblMiniShapeInfo.Caption = "" ''Bug Shape, this will change during input.
''SHAPE FINISH''
'''
''COMMON NAME START''
lblCommonName.Caption = "COMMON NAME:" '' Common Name Label, just ignore this.
lblCommonNameInfo.Caption = "" ''Bug Shape, this will change during input.
''COMMON NAME FINISH''
'''
''CLASS START''
lblClass.Caption = "CLASS:" '' Class Label, just ignore this.
lblClassInfo.Caption = "" ''Bug Shape, this will change during input.
''CLASS FINISH''
'''
''CLASS START''
lblAntenna.Caption = "ANTENNA:" '' Antenna Label, just ignore this.
lblAntennaInfo.Caption = "" ''Bug Antenna?, this will change during input.
''CLASS FINISH''
'''
''WINGS START''
lblWings.Caption = "WINGS:" '' Wings Label, just ignore this.
lblWingsInfo.Caption = "" ''Bug Wings, this will change during input.
''WINGS FINISH''
'''
''CLASS START''
lblColor.Caption = "COLOR:" '' Color Label, just ignore this.
lblColorInfo.Caption = "" ''Bug Color, this will change during input.
''CLASS FINISH''
'''
''CLASS START''
lblLocation.Caption = "LOCATION:" '' Color Label, just ignore this.
lblLocationInfo.Caption = "" ''Bug Color, this will change during input.
''CLASS FINISH''

'FormIcon = App.Path & "\data\animals\DO_NOT_MODIFY\bug.ico"
'Me.Icon = LoadPicture(FormIcon)

''BUTTON SETTINGS AND START UP START''
cmdSearch.Enabled = True
txtMiniName.Enabled = True
cmdClear.Enabled = False
''BUTTON SETTINGS AND START UP END''

''WebFeed.Navigate2 "http://yunodev.com/software/minibeasts/updates/v1/feed.html"
End Sub

Sub IfDoesntExsistCreateIt()
''Basically just the subroutine name, lawl''
strFolder = (App.Path & "\data\animals\" & txtMiniName.Text)
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(strFolder) Then
   objFSO.CreateFolder (strFolder)
End If
''^^Creates The Folder^^''
strFolderImage = (App.Path & "\data\animals\" & txtMiniName.Text & "\image")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists(strFolderImage) Then
   objFSO.CreateFolder (strFolderImage)
End If
''^^CREATES IMAGE FOLDER^^''
''FILE 1 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\antenna.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 1 END''
''FILE 2 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\class.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 2 END''
''FILE 3 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\color.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 3 END''
''FILE 4 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\common name.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 4 END''
''FILE 5 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\legs.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 5 END''
''FILE 6 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\name.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 6 END''
''FILE 7 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\shape.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 7 END''
''FILE 8 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\size.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 8 END''
''FILE 9 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\wings.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 9 END''
''FILE 10 START''
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path & "\data\animals\" & txtMiniName.Text & "\location.txt", True)
    a.WriteLine ("N/A")
    a.Close
''FILE 10 END''
''^^ CREATE THE FILES ^^''
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
   SavePicture Image, "data\animals\" & txtMiniName.Text & "\image\" & txtMiniName.Text & ".JPG"   ' Save picture to file.

''^^CREATES IMAGE^^''
End Sub
