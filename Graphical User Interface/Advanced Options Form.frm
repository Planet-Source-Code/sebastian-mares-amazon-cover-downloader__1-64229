VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmAdvancedOptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Advanced Options"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   ClipControls    =   0   'False
   Icon            =   "Advanced Options Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame fraPageLimit 
      Caption         =   "Page Limit"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin ComCtl2.UpDown udPageLimit 
         Height          =   315
         Left            =   3241
         TabIndex        =   3
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPageLimit"
         BuddyDispid     =   196612
         OrigLeft        =   5160
         OrigTop         =   600
         OrigRight       =   5415
         OrigBottom      =   975
         Max             =   3199
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPageLimit 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Unlimited"
         Top             =   600
         Width           =   3120
      End
      Begin VB.Label lblPageLimit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select the number of &pages to parse:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frmAdvancedOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Sub cmdCancel_Click()

1         On Error GoTo ErrorHandler

          'Omit changes and close form
2         Unload Me

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Advanced Options Form.frm" & vbNewLine & "Form: frmAdvancedOptions" & vbNewLine & "Procedure: cmdCancel_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub cmdOK_Click()

1         On Error GoTo ErrorHandler

          'Apply changes and close form
2         lngPageLimit = udPageLimit.Value
3         Unload Me

4         On Error GoTo 0

5     Exit Sub

'Handle errors
ErrorHandler:
6         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Advanced Options Form.frm" & vbNewLine & "Form: frmAdvancedOptions" & vbNewLine & "Procedure: cmdOK_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
7         End

End Sub

Private Sub Form_Load()

1         On Error GoTo ErrorHandler

2         blnWorking = True
          
          'Evaluate current limit value
3         udPageLimit.Value = lngPageLimit

4         On Error GoTo 0

5     Exit Sub

'Handle errors
ErrorHandler:
6         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Advanced Options Form.frm" & vbNewLine & "Form: frmAdvancedOptions" & vbNewLine & "Procedure: Form_Load" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
7         End

End Sub

Private Sub Form_Unload(Cancel As Integer)

1         On Error GoTo ErrorHandler

2         blnWorking = False
          
3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Advanced Options Form.frm" & vbNewLine & "Form: frmAdvancedOptions" & vbNewLine & "Procedure: Form_Unload" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub udPageLimit_Change()

1         On Error GoTo ErrorHandler

          'Display "Unlimited" instead of 0
2         If udPageLimit.Value = 0 Then
3             txtPageLimit.Text = "Unlimited"
4         End If

5         On Error GoTo 0

6     Exit Sub

'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Advanced Options Form.frm" & vbNewLine & "Form: frmAdvancedOptions" & vbNewLine & "Procedure: udPageLimit_Change" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End

End Sub
