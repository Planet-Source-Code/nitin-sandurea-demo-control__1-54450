VERSION 5.00
Object = "{7B99F42D-0523-11D7-9114-00606770BE77}#6.0#0"; "DemoCreation.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin DemoCreation.DemoCreator DemoCreator1 
      Height          =   105
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   185
      NoOfDays        =   2
      NofTimes        =   2
      ProductName     =   "DEMO123456"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DemoCreator1_DemoCompleted(msg As String)
MsgBox msg, vbExclamation
End Sub
