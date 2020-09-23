VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Buttons"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   Picture         =   "FormMain.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CommonDialogButtonCol 
      Left            =   285
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Projekt1.Button ButtonConnect 
      Height          =   630
      Left            =   1350
      TabIndex        =   0
      Top             =   885
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1111
      ForeColor       =   16711680
      TX              =   "www.clk-calculator.de"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnColor        =   14737632
   End
   Begin Projekt1.Button ButtonChangeCol 
      Height          =   630
      Left            =   1350
      TabIndex        =   1
      Top             =   1837
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1111
      ForeColor       =   16777215
      TX              =   "Change Color"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnColor        =   65280
      BtnHlFrameColor =   16711935
      BtnHlColor      =   49152
   End
   Begin Projekt1.Button ButtonExit 
      Height          =   630
      Left            =   1350
      TabIndex        =   2
      Top             =   2790
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1111
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BtnColor        =   8438015
      BtnHlFrameColor =   16777215
      BtnHlColor      =   33023
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AktButtonColor As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Enum OpenUrlShowConstants
  swHide = 0
  swNormal = 1
  swShowMaximized = 3
  swShowMinimized = 2
  swShowMinNoAcive = 7
End Enum

Public Sub OpenURL(URL As String, _
    Optional ByVal ShowMode As OpenUrlShowConstants = swNormal)
    ShellExecute GetDesktopWindow(), "Open", URL, "", "", ShowMode
End Sub

Private Sub ButtonConnect_Click()
    'open your browser and connect to the www...
    OpenURL ("www.clk-calculator.de")
End Sub


Private Sub Form_Initialize()
    'this is here for a correct working of the IDE
    'good job and thanks to Niel Corns
    'manual start the buttons once in the Initialize Sub
    'and they will run nice and smooth
    ButtonConnect.BtnRun = True
    ButtonChangeCol.BtnRun = True
    ButtonExit.BtnRun = True
End Sub



Private Sub ButtonChangeCol_Click()
    On Error Resume Next
    
    Err.Clear
    CommonDialogButtonCol.CancelError = True
    CommonDialogButtonCol.ShowColor
    If Not (Err.Number > 0) Then
        ButtonChangeCol.BtnColor = CommonDialogButtonCol.Color
        
    End If
    
End Sub

Private Sub ButtonExit_Click()
    End
End Sub


