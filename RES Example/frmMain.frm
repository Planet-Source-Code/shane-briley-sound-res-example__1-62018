VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RES Example - Shane"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Franklin Gothic Medium"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdplay 
      Caption         =   "&Play"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Caution"
      Height          =   225
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Research Complete"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbsite 
      Alignment       =   2  'Center
      Caption         =   "http://smash-masters.org"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RES Sound Example By Shane
'http://smash-masters.org
'ceo@smash-masters.org
'Copyright 2005 - 2006
'First off, lets start with making the resource file !
'Go to Add-Ins at the top and select Add-In Manager.
'Find VB 6 Resource Editor.
'Set to Startup/Loaded. If done correct, a new icon should appear at the top !
'Now click that icon and my already made res file should be there.
'If not click the icon that looks like for 4 squares, and select a WAV file from your pc to add.
'Then save the file and you have now created your res file !
'Now view Private Sub cmdplay_click(), for how to call the sounds.
'Well i hope this example helps you all out !
'Note:
'Using a sound res file, is alot better then OLE.
'It causes less lag in the program and makes it more user friendly, also it keeps lamers from stealing sounds !

Private Sub cmdplay_Click()
On Error Resume Next
If opt1.Value = True Then Call PlaySoundResource(1) 'this plays wav 1 in the resource file, only if opt1 = True.
If opt2.Value = True Then Call PlaySoundResource(2) 'this plays wav 2 in the resource file, only if opt2 = True.
End Sub

Private Sub Form_Unload(Cancel As Integer)
' all this is just to make sure the form unloads all the way on exit !
On Error Resume Next
Unload Me
End
End Sub
