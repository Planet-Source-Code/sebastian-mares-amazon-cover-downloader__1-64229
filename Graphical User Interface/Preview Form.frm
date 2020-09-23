VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Preview"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   Icon            =   "Preview Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picContainer 
      ClipControls    =   0   'False
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   4575
      Begin VB.Image imgPreview 
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fraPreview 
      Caption         =   "&Preview"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.VScrollBar vsbVScroll 
         Height          =   4575
         LargeChange     =   500
         Left            =   4800
         SmallChange     =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   270
      End
      Begin VB.HScrollBar hsbHScroll 
         Height          =   255
         LargeChange     =   500
         Left            =   120
         SmallChange     =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4920
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Public Sub CalculateDimensions()

1         On Error GoTo ErrorHandler

          'Calculate scrollbar values
2         If imgPreview.Height < picContainer.Height Then
3             vsbVScroll.Enabled = False
4           Else
5             vsbVScroll.Enabled = True
6             vsbVScroll.Max = imgPreview.Height - picContainer.Height
7         End If

8         If imgPreview.Width < picContainer.Width Then
9             hsbHScroll.Enabled = False
10          Else
11            hsbHScroll.Enabled = True
12            hsbHScroll.Max = imgPreview.Width - picContainer.Width
13        End If

14        On Error GoTo 0

15    Exit Sub

'Handle errors
ErrorHandler:
16        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: CalculateDimensions" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
17        End

End Sub

Private Sub cmdClose_Click()

1         On Error GoTo ErrorHandler

2         Unload Me

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: cmdClose_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub Form_Load()

1         On Error GoTo ErrorHandler

2         blnWorking = True
          
          'Subclass picture box and intercept mouse wheel scrolls
3         lngOldPictureBoxWindowProcess = GetWindowLong(picContainer.hWnd, GWL_WNDPROC)
4         SetWindowLong picContainer.hWnd, GWL_WNDPROC, AddressOf PictureBoxWindowProcess

          'Subclass form and intercept resize events
5         lngOldFormWindowProcess = GetWindowLong(Me.hWnd, GWL_WNDPROC)
6         SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FormWindowProcess

          'Load form size
7         Me.Height = GetSetting("Amazon Cover Downloader", "frmPreview", "Height", 6390)
8         Me.Width = GetSetting("Amazon Cover Downloader", "frmPreview", "Width", 5550)

9         On Error GoTo 0

10    Exit Sub

'Handle errors
ErrorHandler:
11        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: Form_Load" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
12        End

End Sub

Private Sub Form_Resize()

1         On Error GoTo ErrorHandler

2         cmdClose.Left = Me.Width - cmdClose.Width - 255
3         cmdClose.Top = Me.Height - cmdClose.Height - 495
4         fraPreview.Height = Me.Height - 1095
5         fraPreview.Width = Me.Width - 375
6         hsbHScroll.Top = fraPreview.Height - hsbHScroll.Height - 120
7         hsbHScroll.Width = fraPreview.Width - 600
8         picContainer.Height = fraPreview.Height - 720
9         picContainer.Width = fraPreview.Width - 600
10        vsbVScroll.Height = fraPreview.Height - 720
11        vsbVScroll.Left = fraPreview.Width - vsbVScroll.Width - 105
12        CalculateDimensions

13        On Error GoTo 0

14    Exit Sub

'Handle errors
ErrorHandler:
15        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: Form_Resize" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
16        End

End Sub

Private Sub Form_Unload(Cancel As Integer)

1         On Error GoTo ErrorHandler

          'Save form size
2         SaveSetting "Amazon Cover Downloader", "frmPreview", "Height", Me.Height
3         SaveSetting "Amazon Cover Downloader", "frmPreview", "Width", Me.Width

          'Unsubclass picture box
4         SetWindowLong picContainer.hWnd, GWL_WNDPROC, lngOldPictureBoxWindowProcess

          'Unsubclass form
5         SetWindowLong Me.hWnd, GWL_WNDPROC, lngOldFormWindowProcess

6         blnWorking = False

7         On Error GoTo 0

8     Exit Sub

'Handle errors
ErrorHandler:
9         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: Form_Unload" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
10        End

End Sub

Private Sub hsbHScroll_Change()

1         On Error GoTo ErrorHandler

2         imgPreview.Left = -hsbHScroll.Value

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: hsbHScroll_Change" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub hsbHScroll_Scroll()

1         On Error GoTo ErrorHandler

2         imgPreview.Left = -hsbHScroll.Value

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: hsbHScroll_Scroll" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub imgPreview_Click()

1         On Error GoTo ErrorHandler

2         picContainer.SetFocus

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: imgPreview_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub vsbVScroll_Change()

1         On Error GoTo ErrorHandler

2         imgPreview.Top = -vsbVScroll.Value

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: vsbVScroll_Change" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub vsbVScroll_Scroll()

1         On Error GoTo ErrorHandler

2         imgPreview.Top = -vsbVScroll.Value

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Preview Form.frm" & vbNewLine & "Form: frmPreview" & vbNewLine & "Procedure: vsbVScroll_Scroll" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub
