VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Design: Copyright 2006, iconaholic.com"
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
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   3165
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.maresweb.de/"
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
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "acd@maresweb.de"
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
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2006, MaresWEB. All rights reserved."
      Height          =   195
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   555
      Width           =   3345
   End
   Begin VB.Label lblProgramVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amazon Cover Downloader v3.0.0.6"
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   285
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "About Form.frx":0000
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal lngHandle As Long, ByVal strOperation As String, ByVal strFile As String, ByVal strParameters As String, ByVal strDirectory As String, ByVal lngMode As Long) As Long

Private Sub cmdOK_Click()

1         On Error GoTo ErrorHandler
          
2         Unload Me
              
3         On Error GoTo 0

4     Exit Sub

      'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: cmdOK_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub Form_Load()

1         On Error GoTo ErrorHandler
          
2         blnWorking = True
          
3         lblProgramVersion.Caption = App.Title & " v" & App.Major & "." & App.Minor & ".0." & App.Revision
4         lblCopyright(0).Caption = "Copyright " & Year(Now()) & ", " & App.CompanyName & ". All rights reserved."
5         lblCopyright(1).Caption = "Icon Design: Copyright " & Year(Now()) & ", iconaholic.com"
          
6         On Error GoTo 0

7     Exit Sub

      'Handle errors
ErrorHandler:
8         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: Form_Load" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
9         End

End Sub

Private Sub Form_Unload(Cancel As Integer)

1         On Error GoTo ErrorHandler
          
2         blnWorking = False
          
3         On Error GoTo 0

4     Exit Sub

      'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: Form_Unload" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub lblCopyright_Click(Index As Integer)

1         On Error GoTo ErrorHandler
          
2         If Index = 1 Then
3             ShellExecute Me.hWnd, vbNullString, "http://www.iconaholic.com/", vbNullString, "C:\", SW_SHOWNORMAL
4         End If
          
5         On Error GoTo 0

6     Exit Sub

      'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: lblCopyright_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End
          
End Sub

Private Sub lblEMail_Click(Index As Integer)

1         On Error GoTo ErrorHandler
          
2         If Index = 1 Then
3             ShellExecute Me.hWnd, vbNullString, "mailto:acd@maresweb.de", vbNullString, "C:\", SW_SHOWNORMAL
4         End If
          
5         On Error GoTo 0

6     Exit Sub

      'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: lblEMail_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End
          
End Sub

Private Sub lblWeb_Click(Index As Integer)

1         On Error GoTo ErrorHandler
          
2         If Index = 1 Then
3             ShellExecute Me.hWnd, vbNullString, "http://www.maresweb.de/", vbNullString, "C:\", SW_SHOWNORMAL
4         End If
          
5         On Error GoTo 0

6     Exit Sub

      'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: About Form.frm" & vbNewLine & "Form: frmAbout" & vbNewLine & "Procedure: lblWeb_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End
          
End Sub
