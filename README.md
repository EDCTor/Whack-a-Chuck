#Whack-a-Chuck

##Description

Whack-a-Chuck is a very simple **whack-a-mole** clone.  The goal is to click or tap the **moles** as they pop out of the holes.

##History

In 2002 the game/idea was originally created in Visual Basic as an in office joke, on a slow afternoon between Christmas and New Years.  Two years later, on a slow afternoon between Christmas and New Years, the game was expanded for other employees.

##Visual Basic 5

```
VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wack-EDC"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   4185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   -120
      Top             =   2520
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000014&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblScoreheader 
      BackColor       =   &H80000014&
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgC3 
      Height          =   825
      Left            =   2040
      Picture         =   "frmMain.frx":030A
      Top             =   1800
      Width           =   930
   End
   Begin VB.Image imgC2 
      Height          =   825
      Left            =   1080
      Picture         =   "frmMain.frx":2BB2
      Top             =   1800
      Width           =   930
   End
   Begin VB.Image imgC1 
      Height          =   825
      Left            =   120
      Picture         =   "frmMain.frx":545A
      Top             =   1800
      Width           =   930
   End
   Begin VB.Image imgB3 
      Height          =   825
      Left            =   2040
      Picture         =   "frmMain.frx":7D02
      Top             =   960
      Width           =   930
   End
   Begin VB.Image imgB2 
      Height          =   825
      Left            =   1080
      Picture         =   "frmMain.frx":A5AA
      Top             =   960
      Width           =   930
   End
   Begin VB.Image imgB1 
      Height          =   825
      Left            =   120
      Picture         =   "frmMain.frx":CE52
      Top             =   960
      Width           =   930
   End
   Begin VB.Image imgA3 
      Height          =   825
      Left            =   2040
      Picture         =   "frmMain.frx":F6FA
      Top             =   120
      Width           =   930
   End
   Begin VB.Image imgA2 
      Height          =   825
      Left            =   1080
      Picture         =   "frmMain.frx":11FA2
      Top             =   120
      Width           =   930
   End
   Begin VB.Image imgA1 
      Height          =   825
      Left            =   120
      Picture         =   "frmMain.frx":1484A
      Top             =   120
      Width           =   930
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStart 
         Caption         =   "Start Game"
      End
      Begin VB.Menu mnuStopGame 
         Caption         =   "Stop Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSpeed 
      Caption         =   "Speed"
      Begin VB.Menu mnuSpeed5 
         Caption         =   "Level 5 - Monday"
      End
      Begin VB.Menu mnuSpeed4 
         Caption         =   "Level 4 - Tuesday"
      End
      Begin VB.Menu mnuSpeed3 
         Caption         =   "Level 3 - Wednesday"
      End
      Begin VB.Menu mnuSpeed2 
         Caption         =   "Level 2 - Thursday"
      End
      Begin VB.Menu mnuSpeed1 
         Caption         =   "Level 1 - Friday"
      End
   End
   Begin VB.Menu mnuEmployee 
      Caption         =   "Employee"
      Begin VB.Menu mnuEmployeeChuck 
         Caption         =   "Chuck"
      End
      Begin VB.Menu mnuEmployeeChris 
         Caption         =   "Chris"
      End
      Begin VB.Menu mnuEmployeeSam 
         Caption         =   "Sam"
      End
      Begin VB.Menu mnuEmployeeTor 
         Caption         =   "Tor"
      End
   End
   Begin VB.Menu mnuSpace1 
      Caption         =   "                       "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public successWacks As Double
Public totalWacks As Double
Public current_X As Integer
Public current_Y As Integer
Public speedInterval As Integer

Const empChuck = 0
Const empChris = 1
Const empSam = 2
Const empTor = 3

Public intSelectedEmployee As Integer

Private Sub Form_Load()
    mnuSpeed1.Checked = True
    
    ' default to Tor
    intSelectedEmployee = 3
    mnuEmployeeTor.Checked = True
    
    
    ' check for all required files.  Exit if missing files!
    If Not FileExists(App.Path & "\chuckN.bmp") Or _
       Not FileExists(App.Path & "\chuckW.bmp") Or _
       Not FileExists(App.Path & "\chrisN.bmp") Or _
       Not FileExists(App.Path & "\chrisW.bmp") Or _
       Not FileExists(App.Path & "\samN.bmp") Or _
       Not FileExists(App.Path & "\samW.bmp") Or _
       Not FileExists(App.Path & "\torN.bmp") Or _
       Not FileExists(App.Path & "\torW.bmp") Then
       MsgBox "Wack-EDC is missing critical image files.", vbCritical, "Cannot Initialize Program"
    End If
    
End Sub

Private Sub imgA1_Click()
    If current_X = 1 And current_Y = 1 Then
        successWacks = successWacks + 1
        imgA1.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgA2_Click()
    If current_X = 1 And current_Y = 2 Then
        successWacks = successWacks + 1
        imgA2.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgA3_Click()
    If current_X = 1 And current_Y = 3 Then
        successWacks = successWacks + 1
        imgA3.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgB1_Click()
    If current_X = 2 And current_Y = 1 Then
        successWacks = successWacks + 1
        imgB1.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgB2_Click()
    If current_X = 2 And current_Y = 2 Then
        successWacks = successWacks + 1
        imgB2.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgB3_Click()
    If current_X = 2 And current_Y = 3 Then
        successWacks = successWacks + 1
        imgB3.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgC1_Click()
    If current_X = 3 And current_Y = 1 Then
        successWacks = successWacks + 1
        imgC1.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgC2_Click()
    If current_X = 3 And current_Y = 2 Then
        successWacks = successWacks + 1
        imgC2.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub

Private Sub imgC3_Click()
    If current_X = 3 And current_Y = 3 Then
        successWacks = successWacks + 1
        imgC3.Picture = LoadPicture(getWackedPictureName())
    End If
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuEmployeeChris_Click()
'
    intSelectedEmployee = empChris
    mnuEmployeeChris.Checked = True
    mnuEmployeeChuck.Checked = False
    mnuEmployeeSam.Checked = False
    mnuEmployeeTor.Checked = False
    
End Sub

Private Sub mnuEmployeeChuck_Click()
'
    intSelectedEmployee = empChuck
    mnuEmployeeChris.Checked = False
    mnuEmployeeChuck.Checked = True
    mnuEmployeeSam.Checked = False
    mnuEmployeeTor.Checked = False
    
End Sub

Private Sub mnuEmployeeSam_Click()
'
    intSelectedEmployee = empSam
    mnuEmployeeChris.Checked = False
    mnuEmployeeChuck.Checked = False
    mnuEmployeeSam.Checked = True
    mnuEmployeeTor.Checked = False
    
End Sub

Private Sub mnuEmployeeTor_Click()
'
    intSelectedEmployee = empTor
    mnuEmployeeChris.Checked = False
    mnuEmployeeChuck.Checked = False
    mnuEmployeeSam.Checked = False
    mnuEmployeeTor.Checked = True
    
End Sub

Private Sub mnuExit_Click()
    Timer1.Enabled = False
    MsgBox "High Score: " & successWacks, vbInformation, "High Score"
    Unload Me
End Sub

Private Sub mnuSpeed1_Click()
    mnuSpeed1.Checked = True
    mnuSpeed2.Checked = False
    mnuSpeed3.Checked = False
    mnuSpeed4.Checked = False
    mnuSpeed5.Checked = False
End Sub

Private Sub mnuSpeed2_Click()
    mnuSpeed1.Checked = False
    mnuSpeed2.Checked = True
    mnuSpeed3.Checked = False
    mnuSpeed4.Checked = False
    mnuSpeed5.Checked = False
End Sub

Private Sub mnuSpeed3_Click()
    mnuSpeed1.Checked = False
    mnuSpeed2.Checked = False
    mnuSpeed3.Checked = True
    mnuSpeed4.Checked = False
    mnuSpeed5.Checked = False
End Sub

Private Sub mnuSpeed4_Click()
    mnuSpeed1.Checked = False
    mnuSpeed2.Checked = False
    mnuSpeed3.Checked = False
    mnuSpeed4.Checked = True
    mnuSpeed5.Checked = False
End Sub

Private Sub mnuSpeed5_Click()
    mnuSpeed1.Checked = False
    mnuSpeed2.Checked = False
    mnuSpeed3.Checked = False
    mnuSpeed4.Checked = False
    mnuSpeed5.Checked = True
End Sub

Private Sub mnuStart_Click()
    
    If mnuSpeed1.Checked Then
        speedInterval = 750
    ElseIf mnuSpeed2.Checked Then
        speedInterval = 650
    ElseIf mnuSpeed3.Checked Then
        speedInterval = 540
    ElseIf mnuSpeed4.Checked Then
        speedInterval = 440
    ElseIf mnuSpeed5.Checked Then
        speedInterval = 320
    Else
        speedInterval = 0
    End If
    
    Timer1.Enabled = True
    Timer1.Interval = speedInterval

End Sub

Private Sub mnuStopGame_Click()
    Timer1.Enabled = False
    
    ' pause here for a second
    
    If current_X = 1 And current_Y = 1 Then
        imgA1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 1 And current_Y = 2 Then
        imgA2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 1 And current_Y = 3 Then
        imgA3.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 1 Then
        imgB1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 2 Then
        imgB2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 3 Then
        imgB3.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 1 Then
        imgC1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 2 Then
        imgC2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 3 Then
        imgC3.Picture = LoadPicture(App.Path & "\hole.bmp")
    End If
    
End Sub

Private Sub Timer1_Timer()
    
    If mnuSpeed1.Checked Then
        speedInterval = 750
    ElseIf mnuSpeed2.Checked Then
        speedInterval = 650
    ElseIf mnuSpeed3.Checked Then
        speedInterval = 540
    ElseIf mnuSpeed4.Checked Then
        speedInterval = 440
    ElseIf mnuSpeed5.Checked Then
        speedInterval = 320
    Else
        speedInterval = 0
    End If
    
    Timer1.Interval = speedInterval
     
    If current_X = 1 And current_Y = 1 Then
        imgA1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 1 And current_Y = 2 Then
        imgA2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 1 And current_Y = 3 Then
        imgA3.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 1 Then
        imgB1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 2 Then
        imgB2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 2 And current_Y = 3 Then
        imgB3.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 1 Then
        imgC1.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 2 Then
        imgC2.Picture = LoadPicture(App.Path & "\hole.bmp")
    ElseIf current_X = 3 And current_Y = 3 Then
        imgC3.Picture = LoadPicture(App.Path & "\hole.bmp")
    End If
    
    totalWacks = totalWacks + 1
    
    Select Case Round((Rnd * 100), 0) Mod 9
        Case 1
            imgA1.Picture = LoadPicture(getNormalPictureName())
            current_X = 1
            current_Y = 1
        Case 2
            imgA2.Picture = LoadPicture(getNormalPictureName())
            current_X = 1
            current_Y = 2
        Case 3
            imgA3.Picture = LoadPicture(getNormalPictureName())
            current_X = 1
            current_Y = 3
        Case 4
            imgB1.Picture = LoadPicture(getNormalPictureName())
            current_X = 2
            current_Y = 1
        Case 5
            imgB2.Picture = LoadPicture(getNormalPictureName())
            current_X = 2
            current_Y = 2
        Case 6
            imgB3.Picture = LoadPicture(getNormalPictureName())
            current_X = 2
            current_Y = 3
        Case 7
            imgC1.Picture = LoadPicture(getNormalPictureName())
            current_X = 3
            current_Y = 1
        Case 8
            imgC2.Picture = LoadPicture(getNormalPictureName())
            current_X = 3
            current_Y = 2
        Case 0
            imgC3.Picture = LoadPicture(getNormalPictureName())
            current_X = 3
            current_Y = 3
        Case Else
        
    End Select
    
    lblScore.Caption = successWacks
    frmMain.Refresh
    
End Sub

Private Function getWackedPictureName() As String
    Dim strName As String
    
    If intSelectedEmployee = empChuck Then
        strName = App.Path & "\chuckW.bmp"
    ElseIf intSelectedEmployee = empChris Then
        strName = App.Path & "\chrisW.bmp"
    ElseIf intSelectedEmployee = empSam Then
        strName = App.Path & "\samW.bmp"
    ElseIf intSelectedEmployee = empTor Then
        strName = App.Path & "\torW.bmp"
    Else
        'error, default to tor
        strName = App.Path & "\torW.bmp"
    End If
    
    getWackedPictureName = strName
    
End Function

Private Function getNormalPictureName() As String
    Dim strName As String
    
    If intSelectedEmployee = empChuck Then
        strName = App.Path & "\chuckN.bmp"
    ElseIf intSelectedEmployee = empChris Then
        strName = App.Path & "\chrisN.bmp"
    ElseIf intSelectedEmployee = empSam Then
        strName = App.Path & "\samN.bmp"
    ElseIf intSelectedEmployee = empTor Then
        strName = App.Path & "\torN.bmp"
    Else
        'error, default to tor
        strName = App.Path & "\torN.bmp"
    End If
    
    getNormalPictureName = strName
End Function

' Author: Torrance Jones
' Description: Check if a file exists
' NOTE: not tested against directories
'
Public Function FileExists(filename As String) As Boolean

On Error GoTo errorHandle
    
    If Len(filename) > 0 Then
        If Dir(filename, vbNormal) = "" Then
            FileExists = False
        Else
            FileExists = True
        End If
    Else
        FileExists = False
    End If
    
    Exit Function
errorHandle:
    FileExists = False
End Function
```
