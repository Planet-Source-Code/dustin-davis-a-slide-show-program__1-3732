VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slide Show v1.0"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Add All"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Auto Start Slide Show on Load"
      Height          =   195
      Left            =   3960
      TabIndex        =   14
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "5"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Timer"
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show File Name During Slide Show"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Auto Remove"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start the Slide Show!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   9
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "->"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   6885
      ItemData        =   "Form1.frx":0000
      Left            =   3960
      List            =   "Form1.frx":0002
      TabIndex        =   1
      ToolTipText     =   "List of images to play"
      Top             =   0
      Width           =   4335
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Hidden          =   -1  'True
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      System          =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   7215
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Number of Seconds to show picture"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7215
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'Slide Show v1.0
'Author: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'Slide Show has alot of things it can do, Add as many images as u want,
'add Bitmaps, Jpeg Files and GIF files, mix em up even
'You can add ALL of the image file listed with one click of a button
'One click remove item from list
'empty list with one click
'Auto start show on load
'Pause the show
'timer settings and so on.
'
'Please do not steal this code! It was not coded by you. If you use it,
'please give me proper credit for it!
'ENJOY!!
'This is the form for Settings!
'*********************************************************************************

Private Sub Check1_Click()
'this will turn On/Off the label on form2 during the show
If Check1.Value = Checked Then
    SaveSetting "SlideShow", "settings", "showfile", "1"
    Exit Sub
ElseIf Check1.Value = Unchecked Then
    SaveSetting "SlideShow", "settings", "showfile", "0"
    Exit Sub
End If
End Sub

Private Sub Check3_Click()
'this will turn On/Off Auto Start. So if auto start is On or checked
'then when the program loads, it will automaticly start the show
If Check3.Value = Checked Then
    SaveSetting "SlideShow", "settings", "autostart", "1"
    Exit Sub
ElseIf Check3.Value = Unchecked Then
    SaveSetting "SlideShow", "settings", "autostart", "0"
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
'this will add pictures to the show list
Dim FilePath As String
Dim pic As String
Dim temp As String
Dim length As Integer
Dim where As Integer
'This will get the name and path of the picture
FilePath = File1.Path + "\" + File1.FileName
'this will check to see if the file is a picture file or not
pic = Dir1.Path & "\" & File1.FileName
temp = File1.FileName
length = Len(temp)
where = InStr(temp, ".")
temp = Right$(temp, length - where)

If temp = "bmp" Then 'if it is a bitmap then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
 
ElseIf temp = "jpg" Then 'if it is a jpg file then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
    
ElseIf temp = "gif" Then 'if it is a gif file then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
    
ElseIf temp = "BMP" Then 'if it is a bitmap then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
 
ElseIf temp = "JPG" Then 'if it is a jpg file then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
    
ElseIf temp = "GIF" Then 'if it is a jpg file then add to list
List1.AddItem FilePath
SaveSetting "SlideShow", "pictures", "HowMany", List1.ListCount
SaveSetting "SlideShow", "pictures", List1.ListCount, FilePath
    
Else 'if it isnt a bitmap, jpg or gif file then exit
    'MsgBox "This is not a picture file!!" & vbCrLf & "Must be a .JPG, .BMP or .GIF Format!", vbExclamation
    Exit Sub
End If

End Sub

Private Sub Command2_Click()
'this will remove the selected picture from the list!
SaveSetting "Slideshow", "pictures", List1.ListIndex, ""
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
'this will set the timer
If Text1.Text > 65 Then 'if over 65 seconds, it wont work!
    MsgBox "Please Enter a number between 1-65"
    Exit Sub
End If
SaveSetting "SlideShow", "settings", "delay", Text1.Text 'set timer
End Sub

Private Sub Command4_Click()
'this will clear ALL of the picture settings and locations
Dim TotalImages As Long
Dim CurrentImg As Long
TotalImages = GetSetting("SlideShow", "pictures", "HowMany", "0")
CurrentImg = 0
List1.Clear
Do Until CurrentImg = TotalImages
    CurrentImg = CurrentImg + 1
    SaveSetting "SlideShow", "pictures", CurrentImg, ""
Loop
SaveSetting "SlideShow", "pictures", "HowMany", "0"

End Sub

Private Sub Command5_Click()
'this opens the show
Form2.Visible = True
End Sub

Private Sub Command6_Click()
'this adds all of the picture files to the list, will not add anything else
Dim TotalFiles As Long
Dim CurrentFile As Long
CurrentFile = 0
TotalFiles = File1.ListCount

Do Until CurrentFile >= TotalFiles
    File1.ListIndex = CurrentFile
    CurrentFile = CurrentFile + 1
    Call Command1_Click 'call this to scan the files, if good, then add to list
Loop
End Sub

Private Sub Dir1_Change()
'this will change the path of the files
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
'this will change the path of the folders
On Error GoTo errDrive
Dir1.Path = Drive1.Drive
errDrive:
    If Err.Number = 68 Then 'this means that the drive isnt accessable!
        MsgBox "Drive Not Ready!", vbOKOnly
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Dim TotalImages As Long
Dim CurrentImg As Long
Dim PicturePath As String
'start listing what is in the current list
CurrentImg = 0
TotalImages = GetSetting("SlideShow", "pictures", "HowMany", "0")

If TotalImages = 0 Then
    Exit Sub
Else
    Do Until CurrentImg >= TotalImages
        CurrentImg = CurrentImg + 1
        PicturePath = GetSetting("SlideShow", "pictures", CurrentImg, "0")
        If PicturePath = "0" Then
            List1.AddItem "Empty"
        Else
            List1.AddItem PicturePath
        End If
    Loop
End If
'end listing of current list
'show timer setting
Text1.Text = GetSetting("SlideShow", "settings", "delay", "5")
'check if autostart is enabled
If GetSetting("SlideShow", "settings", "autostart", "0") = 1 Then
    Check3.Value = Checked
End If
'check if showfile is enabled
If GetSetting("SlideShow", "settings", "showfile", "1") = 1 Then
    Check1.Value = Checked
End If
'if auto start is enabled then start the show!
If Check3.Value = Checked Then
    Form2.Visible = True
    Form1.Visible = False
End If

End Sub

Private Sub List1_Click()
'this is for the auto remove so u dont have to push the button!
'just click and erase!! :)
If Check2.Value = Checked Then
    SaveSetting "Slideshow", "pictures", List1.ListIndex, ""
    List1.RemoveItem List1.ListIndex
ElseIf Check2.Value = Unchecked Then
    DoEvents
End If
End Sub

