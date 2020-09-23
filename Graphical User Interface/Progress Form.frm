VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Progress"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Progress Form.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Pfeil und Sanduhr
   ScaleHeight     =   855
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin ComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblInformation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while completing the current task..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Sub cmdCancel_Click()

1         On Error GoTo ErrorHandler

          'Cancel activity and close form
2         frmMain.objACDCore.Cancel
3         frmMain.lngCancel = 1
4         Unload Me

5         On Error GoTo 0

6     Exit Sub

'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Progress Form.frm" & vbNewLine & "Form: frmProgress" & vbNewLine & "Procedure: cmdCancel_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End

End Sub
