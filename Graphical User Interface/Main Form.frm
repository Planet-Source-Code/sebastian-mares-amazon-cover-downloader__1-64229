VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Amazon Cover Downloader"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6120
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6120
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'Kein
      Height          =   4185
      Index           =   0
      Left            =   150
      ScaleHeight     =   4185
      ScaleWidth      =   5805
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   5805
      Begin VB.CommandButton cmdAdvancedOptions 
         Caption         =   "Advanced &Options"
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame fraKeywords 
         Caption         =   "Keywords"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5535
         Begin VB.TextBox txtKeywords 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label lblKeywords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please enter &keywords for the query in the text box below:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Width           =   4080
         End
      End
      Begin VB.Frame fraMediaType 
         Caption         =   "Media Type"
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   5535
         Begin VB.ComboBox cboMediaType 
            Height          =   315
            ItemData        =   "Main Form.frx":0000
            Left            =   120
            List            =   "Main Form.frx":000D
            Sorted          =   -1  'True
            TabIndex        =   6
            Text            =   "Music"
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label lblMediaType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please select a media &type for the query from the list below:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   285
            Width           =   4170
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "&Query"
         Default         =   -1  'True
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   5535
         Begin VB.ComboBox cboServer 
            Height          =   315
            ItemData        =   "Main Form.frx":0027
            Left            =   120
            List            =   "Main Form.frx":003D
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   9
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please select an Amazon Web Services &server from the list below:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   285
            Width           =   4665
         End
      End
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'Kein
      Height          =   4185
      Index           =   1
      Left            =   150
      ScaleHeight     =   4185
      ScaleWidth      =   5805
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   5805
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame fraThumbnails 
         Caption         =   "&Thumbnails"
         Height          =   2295
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   5535
         Begin ComctlLib.ListView lvwThumbnails 
            Height          =   1935
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3413
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   5535
         Begin VB.ComboBox cboImageSize 
            Height          =   315
            ItemData        =   "Main Form.frx":008A
            Left            =   1080
            List            =   "Main Form.frx":0097
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   19
            Top             =   240
            Width           =   4335
         End
         Begin VB.CheckBox chkDegrade 
            Caption         =   "&Degrade image size when desired size is not available"
            Height          =   255
            Left            =   1080
            TabIndex        =   20
            Top             =   675
            Value           =   1  'Aktiviert
            Width           =   4335
         End
         Begin VB.Label lblImageSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Image Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4560
         TabIndex        =   21
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3720
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar staStatus 
      Align           =   2  'Unten ausrichten
      Height          =   285
      Left            =   0
      TabIndex        =   23
      Top             =   4800
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Idle"
            TextSave        =   "Idle"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8934
            MinWidth        =   4410
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip tabTabs 
      Height          =   4575
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8070
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Query"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Result"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlSaveAs 
      Left            =   840
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "jpg"
      Filter          =   "JPEG Files (*.jpg)|*.jpg"
      Flags           =   1030
      MaxFileSize     =   32767
   End
   Begin ComctlLib.ImageList imlThumbnails 
      Left            =   120
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Const CS_DROPSHADOW As Long = &H20000
Private Const ERROR_ALREADY_EXISTS As Long = 183
Private Const GCL_STYLE As Long = -26
Private Const GWL_STYLE As Long = -16
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_HITTEST As Long = LVM_FIRST + 18
Private Const MF_BYPOSITION As Long = &H400
Private Const MF_SEPARATOR As Long = &H800
Private Const PBS_MARQUEE As Long = &H8
Private Const PBM_SETMARQUEE As Long = WM_USER + &HA
Private Const QS_ALLINPUT As Long = &HFF
Private Const SC_RESTORE As Long = &HF120&
Private Const WM_SYSCOMMAND As Long = &H112

Private Type Point
    lngX As Long
    lngY As Long
End Type
Private Type ListViewHitTestStructure
    udtPoint As Point
    lngFlags As Long
    lngItem As Long
    lngSubItem As Long
End Type

Private Declare Function BringWindowToTop Lib "user32" (ByVal lngHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal lngObject As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lngAttributes As Long, ByVal lngOwner As Long, ByVal strName As String) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lngEnumerationFunction As Long, ByVal lngLParameter As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal lngHandle As Long, ByVal lngIndex As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal lngMenu As Long) As Long
Private Declare Function GetQueueStatus Lib "user32" (ByVal lngFlags As Long) As Boolean
Private Declare Function GetSystemMenu Lib "user32" (ByVal lngHandle As Long, ByVal lngRestore As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal lngMenu As Long, ByVal lngPosition As Long, ByVal lngFlags As Long, ByVal lngNewItemID As Long, ByVal lngNewItem As Any) As Long
Private Declare Function IsIconic Lib "user32" (ByVal lngHandle As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal strPath As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal lngMutex As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, anyLParameter As Any, ByVal lngFlags As Long, ByVal lngTimeout As Long, lngResult As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal lngHandle As Long, ByVal lngIndex As Long, ByVal lngNewClassLong As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal lngHandle As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal lngHandle As Long, ByVal strProperty As String, ByVal lngData As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal lngCaller As Long, ByVal strURL As String, ByVal strFilename As String, ByVal lngReserved As Long, ByVal lngCallBack As Long) As Long

Private blnPreviousInstance As Boolean
Private blnProgressDataAvailable As Boolean
Public lngCancel As Long
Private lngCurrentListViewItemIndex As Long
Public WithEvents objACDCore As AmazonCoverDownloaderCore.clsMain
Attribute objACDCore.VB_VarHelpID = -1
Public objToolTip As clsToolTip

Private Sub chkDegrade_Click()

1         On Error GoTo ErrorHandler

          'Don't allow semi-checked state
2         If chkDegrade.Value = vbGrayed Then
3             chkDegrade.Value = vbChecked
4         End If

5         On Error GoTo 0

6     Exit Sub

'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: chkDegrade_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End

End Sub

Private Sub cmdAdvancedOptions_Click()

1         On Error GoTo ErrorHandler

2         frmAdvancedOptions.Show vbModal, Me

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: cmdAdvancedOptions_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub cmdExit_Click(Index As Integer)

1         On Error GoTo ErrorHandler

2         Unload Me

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: cmdExit_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub cmdPreview_Click()

        Dim lngHeight As Long
        Dim lngWidth As Long
        Dim picPicture As Picture

1         On Error GoTo ErrorHandler

2         If Not lvwThumbnails.ListItems.Count = 0 Then
3             If Not lvwThumbnails.SelectedItem Is Nothing Then
                  'Download selected image with the desired size
4                 Select Case AvailableImageSize(Choose(cboImageSize.ListIndex + 1, 3, 2, 1), Mid$(lvwThumbnails.SelectedItem.Key, 9), chkDegrade.Value = vbChecked)
                    Case Large
5                     DownloadImage objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strLargeImageURL, picPicture
6                     lngHeight = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngLargeImageHeight
7                     lngWidth = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngLargeImageWidth
8                   Case Medium
9                     DownloadImage objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strMediumImageURL, picPicture
10                    lngHeight = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngMediumImageHeight
11                    lngWidth = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngMediumImageWidth
12                  Case Small
13                    DownloadImage objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strSmallImageURL, picPicture
14                    lngHeight = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngSmallImageHeight
15                    lngWidth = objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).lngSmallImageWidth
16                  Case Else
17                    InformUser "The desired image size is unfortunately not available.", vbInformation, "Information"
18                    Exit Sub
19                End Select

                  'Preview image
20                If Not picPicture Is Nothing Then
21                    frmPreview.imgPreview.Picture = picPicture
22                    frmPreview.CalculateDimensions
23                    frmPreview.fraPreview.Caption = "&Preview (" & lngWidth & "x" & lngHeight & ")"
24                    frmPreview.Show vbModal, Me
25                  Else
26                    InformUser "An unknown error occurred while attempting to load the desired image.", vbCritical, "Error"
27                End If
28            End If
29        End If

30        On Error GoTo 0

31    Exit Sub

'Handle errors
ErrorHandler:
32        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: cmdPreview_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
33        End

End Sub

Private Sub cmdQuery_Click()

1         On Error GoTo ErrorHandler

          'Perform query
2         Query txtKeywords.Text, cboServer.Text, cboMediaType.Text, lngPageLimit

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: cmdQuery_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub cmdSave_Click()

        Dim blnAutoExit As Boolean
        Dim lngDownloadExitCode As Boolean

1         On Error GoTo ErrorHandler

2         If Not lvwThumbnails.ListItems.Count = 0 Then
3             If Not lvwThumbnails.SelectedItem Is Nothing Then
                  'If the Shift key is pressed, assume that the user wants to exit after the image was saved successfully
4                 blnAutoExit = (GetAsyncKeyState(vbKeyShift) And &H8000) = &H8000

                  'Download image
5                 blnWorking = True
6                 Select Case AvailableImageSize(Choose(cboImageSize.ListIndex + 1, 3, 2, 1), Mid$(lvwThumbnails.SelectedItem.Key, 9), chkDegrade.Value = vbChecked)
                    Case Large
                      'Display "Save As" dialog
7                     cdlSaveAs.ShowSave
8                     lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strLargeImageURL, cdlSaveAs.FileName, 0, 0)
9                   Case Medium
                      'Display "Save As" dialog
10                    cdlSaveAs.ShowSave
11                    lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strMediumImageURL, cdlSaveAs.FileName, 0, 0)
12                  Case Small
                      'Display "Save As" dialog
13                    cdlSaveAs.ShowSave
14                    lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(Mid$(lvwThumbnails.SelectedItem.Key, 9)).strSmallImageURL, cdlSaveAs.FileName, 0, 0)
15                  Case Else
16                    InformUser "The desired image size is unfortunately not available.", vbInformation, "Information"
17                    Exit Sub
18                End Select
19                blnWorking = False

                  'Check if the image was downloaded successfully
20                If lngDownloadExitCode = 0 Then
21                    If blnAutoExit Then
22                        Unload Me
23                    End If
24                  Else
25                    InformUser "An unknown error occurred while attempting to save the desired image.", vbCritical, "Error"
26                End If

27            End If
28        End If

29        On Error GoTo 0

30    Exit Sub

'Handle errors
ErrorHandler:
31        If Not Err.Number = 32755 Then
32            MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: cmdSave_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
33            End
34        End If

End Sub

Private Sub DisableControls()

        Dim ctlControl As Control

1         On Error GoTo ErrorHandler

2         For Each ctlControl In frmMain.Controls
3             If (TypeOf ctlControl Is ComboBox) Or _
                  (TypeOf ctlControl Is CheckBox) Or _
                  (TypeOf ctlControl Is CommandButton) Or _
                  (TypeOf ctlControl Is Frame) Or _
                  (TypeOf ctlControl Is Label) Or _
                  (TypeOf ctlControl Is ListView) Or _
                  (TypeOf ctlControl Is TabStrip) Or _
                  (TypeOf ctlControl Is TextBox) Then
4                 If Not ctlControl.Name = "cmdExit" Then
5                     ctlControl.Enabled = False
6                 End If
7             End If
8         Next ctlControl

9         On Error GoTo 0

10    Exit Sub

'Handle errors
ErrorHandler:
11        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: DisableControls" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
12        End

End Sub

Private Sub EnableControls()

        Dim ctlControl As Control

1         On Error GoTo ErrorHandler

2         For Each ctlControl In frmMain.Controls
3             If (TypeOf ctlControl Is ComboBox) Or _
                  (TypeOf ctlControl Is CheckBox) Or _
                  (TypeOf ctlControl Is CommandButton) Or _
                  (TypeOf ctlControl Is Frame) Or _
                  (TypeOf ctlControl Is Label) Or _
                  (TypeOf ctlControl Is ListView) Or _
                  (TypeOf ctlControl Is TabStrip) Or _
                  (TypeOf ctlControl Is TextBox) Then
4                 ctlControl.Enabled = True
5             End If
6         Next ctlControl

7         On Error GoTo 0

8     Exit Sub

'Handle errors
ErrorHandler:
9         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: EnableControls" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
10        End

End Sub

Private Sub Form_Initialize()

1         On Error GoTo ErrorHandler

          'Instantiate Common Controls 6 window classes
2         InitCommonControls

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: Form_Initialize" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1         On Error GoTo ErrorHandler

          'Toggle view with Ctrl+Tab, Ctrl+Shift+Tab, Ctrl+1 and Ctrl+2
2         If KeyCode = vbKeyTab Then
3             If Shift = vbCtrlMask Or Shift = vbShiftMask + vbCtrlMask Then
4                 If tabTabs.SelectedItem.Index = 1 Then
5                     tabTabs.Tabs(2).Selected = True
6                   Else
7                     tabTabs.Tabs(1).Selected = True
8                 End If
9             End If
10          ElseIf KeyCode = vbKey1 Then
11            If Shift = vbCtrlMask Then
12                tabTabs.Tabs(1).Selected = True
13            End If
14          ElseIf KeyCode = vbKey2 Then
15            If Shift = vbCtrlMask Then
16                tabTabs.Tabs(2).Selected = True
17            End If
18        End If

19        On Error GoTo 0

20    Exit Sub

'Handle errors
ErrorHandler:
21        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: Form_KeyDown" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
22        End

End Sub

Private Sub Form_Load()

        Dim abytParameters() As Byte
        Dim lngApplicationMutex As Long
        Dim lngNumberOfMenuItems As Long
        Dim lngSystemMenuHandle As Long
        Dim udtCopyDataStructure As CopyDataStucture

1         On Error GoTo ErrorHandler

          'Parse command line arguments and display possible errors
2         If Not ParseCommandLineArguments Then
3             InformUser "Some of the command line arguments used were not valid and were therefore ignored.", vbExclamation, "Warning"
4         End If

          'Set application mutex if not in IDE
5         If Not DebugMode Then
6             lngApplicationMutex = CreateMutex(0, 1, "MW_ACD_{EDD1F962-EC56-40BA-B2C5-773F25EF26EA}")
7             If Err.LastDllError = ERROR_ALREADY_EXISTS Then
                  'If the application is already running, release the mutex, pass the parameters to the existing instance and close the application
8                 ReleaseMutex lngApplicationMutex
9                 CloseHandle lngApplicationMutex

10                blnPreviousInstance = True

11                EnumWindows AddressOf EnumerateWindowsProcess, ByVal 0

12                If Not lngFirstACDWindowProcess = 0 Then
13                    abytParameters = StrConv(chkDegrade.Value & "|" & cdlSaveAs.FileName & "|" & cboMediaType.Text & "|" & txtKeywords.Text & "|" & lngPageLimit & "|" & cboServer.Text & "|" & cboImageSize.Text, vbFromUnicode)
14                    udtCopyDataStructure.lngDataPointer = VarPtr(abytParameters(0))
15                    udtCopyDataStructure.lngDataSize = UBound(abytParameters) + 1

16                    SendMessageTimeout lngFirstACDWindowProcess, WM_COPYDATA, 0, udtCopyDataStructure, 0, 5000, 0

                      If Not IsIconic(lngFirstACDWindowProcess) Then
                          SendMessageLong lngFirstACDWindowProcess, WM_SYSCOMMAND, SC_RESTORE, 0
                      End If
                      SetForegroundWindow lngFirstACDWindowProcess
                      BringWindowToTop lngFirstACDWindowProcess
17                End If

18                Unload Me
19                Exit Sub
20            End If
21        End If

          'Set custom property to form so we can pick it easier from the list of windows when needed (when communicating between instances)
22        SetProp Me.hWnd, "MW_ACD_{EDD1F962-EC56-40BA-B2C5-773F25EF26EA}", 1

          'Instantiate core and tool tip class
23        Set objACDCore = New clsMain
24        Set objToolTip = New clsToolTip

25        If Not blnSilentMode Then
              'Apply visual effects, such as fancy shadows and icons
26            SetClassLong Me.hWnd, GCL_STYLE, GetClassLong(Me.hWnd, GCL_STYLE) Or CS_DROPSHADOW
27            SetIcon Me.hWnd, "ACD"

              'Add custom item to the system menu and subclass it
28            lngSystemMenuHandle = GetSystemMenu(Me.hWnd, False)
29            If lngSystemMenuHandle Then
30                lngNumberOfMenuItems = GetMenuItemCount(lngSystemMenuHandle)
31                If lngNumberOfMenuItems Then
32                    lngAboutMenuID = CreateMenu
33                    InsertMenu lngSystemMenuHandle, lngNumberOfMenuItems, MF_BYPOSITION Or MF_SEPARATOR, 0, vbNullString
34                    InsertMenu lngSystemMenuHandle, lngNumberOfMenuItems + 1, MF_BYPOSITION, lngAboutMenuID, "&About"
35                End If
36            End If
37            lngOldSystemMenuWindowProcess = GetWindowLong(Me.hWnd, GWL_WNDPROC)
38            SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf SystemMenuWindowProcess

              'Subclass form for WM_COPYDATA so we can receive messages passed from other instances
39            lngACDWindowProcess = GetWindowLong(Me.hWnd, GWL_WNDPROC)
40            SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf ACDWindowProcess

              'Show form
41            Me.Show

              'Check if there are queued message and parse them
42            If GetQueueStatus(QS_ALLINPUT) Then
43                DoEvents
44            End If

              'Automatically perform query if keywords were supplied
45            If Not LenB(txtKeywords.Text) = 0 Then
46                Query txtKeywords.Text, cboServer.Text, cboMediaType.Text, lngPageLimit
47            End If
48          Else
              'Prevent form from appearing, download the first image and exit
49            Me.Hide

50            If Not LenB(cdlSaveAs.FileName) = 0 Then
51                Query txtKeywords.Text, cboServer.Text, cboMediaType.Text, lngPageLimit
52              Else
53                InformUser "You did not enter a filename for the cover. Please rectify this problem and try again.", vbExclamation, "Warning"
54            End If

55            Unload Me
56        End If

57        On Error GoTo 0

58    Exit Sub

'Handle errors
ErrorHandler:
59        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: Form_Load" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
60        End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

1         On Error GoTo ErrorHandler

          'Cancel activity before exitting
2         If blnWorking Then
3             objACDCore.Cancel
4             lngCancel = 2
5         End If

6         On Error GoTo 0

7     Exit Sub

'Handle errors
ErrorHandler:
8         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: Form_QueryUnload" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
9         End

End Sub

Private Sub Form_Unload(Cancel As Integer)

1         On Error GoTo ErrorHandler

2         If Not blnPreviousInstance Then
              'Unload any loaded forms
3             UnloadForms

              'Unsubclass system menu
4             If Not blnSilentMode Then
5                 SetWindowLong Me.hWnd, GWL_WNDPROC, lngOldSystemMenuWindowProcess
6             End If

7             Set objACDCore = Nothing
8         End If

9         On Error GoTo 0

10    Exit Sub

'Handle errors
ErrorHandler:
11        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: Form_Unload" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
12        End

End Sub

Private Sub lvwThumbnails_DblClick()

1         On Error GoTo ErrorHandler

          'Preview selected cover
2         cmdPreview_Click

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: lvwThumbnails_DblClick" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Private Sub lvwThumbnails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        Dim udtListViewHitTestStructure As ListViewHitTestStructure
        Dim lngItemIndex As Long

1         On Error GoTo ErrorHandler

          'Obtain the mouse coordinates in pixels
2         udtListViewHitTestStructure.udtPoint.lngX = X \ Screen.TwipsPerPixelX
3         udtListViewHitTestStructure.udtPoint.lngY = Y \ Screen.TwipsPerPixelY

          'Obtain the item index of the item under the cursor
4         lngItemIndex = SendMessage(lvwThumbnails.hWnd, LVM_HITTEST, 0, udtListViewHitTestStructure) + 1

          'Check if the mouse is over the same item
5         If Not lngCurrentListViewItemIndex = lngItemIndex Then
6             lngCurrentListViewItemIndex = lngItemIndex

7             If lngCurrentListViewItemIndex = 0 Then
                  'The cursor is not over an item - destroy tool tip
8                 objToolTip.Destroy
9               Else
10                objToolTip.Style = Balloon
11                objToolTip.Title = "Information about " & lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key & " (" & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).strASIN & ")"
12                objToolTip.ToolTip = "General Information:" & vbNewLine & _
                                       "     Title: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).strTitle & vbNewLine & _
                                       "     Artist: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).strArtist & vbNewLine & _
                                       "     Product Group: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).strProductGroup & vbNewLine & _
                                       vbNewLine & _
                                       "Image Dimensions:" & vbNewLine & _
                                       "     Large Image: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngLargeImageWidth & "x" & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngLargeImageHeight & vbNewLine & _
                                       "     Medium Image: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngMediumImageWidth & "x" & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngMediumImageHeight & vbNewLine & _
                                       "     Small Image: " & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngSmallImageWidth & "x" & objACDCore.Item(Mid$(lvwThumbnails.ListItems(lngCurrentListViewItemIndex).Key, 9)).lngSmallImageHeight
13                objToolTip.Create lvwThumbnails.hWnd
14            End If
15        End If

16        On Error GoTo 0

17    Exit Sub

'Handle errors
ErrorHandler:
18        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: lvwThumbnails_MouseMove" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
19        End

End Sub

Private Sub objACDCore_Done(ByVal lngStatusCode As Long)

        Dim blnErrorsOccurred As Boolean
        Dim lngCurrentItem As Long
        Dim lngDownloadExitCode As Long
        Dim lngNumberOfItemsFound As Long
        Dim picPicture As Picture
        Dim strSmallImageURL As String

1         On Error GoTo ErrorHandler

2         If lngStatusCode = 200 Then
              'Retrieve total number of results (mapped to large image height of item 0)
3             lngNumberOfItemsFound = objACDCore.Item(0).lngLargeImageHeight

4             If lngNumberOfItemsFound > 0 Then
5                 If Not blnSilentMode Then
6                     For lngCurrentItem = 1 To lngNumberOfItemsFound
7                         If lngCancel = 1 Then
                              'Stop downloading additional thumbnails if user cancelled
8                             Exit For
9                           ElseIf lngCancel = 2 Then
                              'Cancel whole procedure if the program is quitting
10                            Exit Sub
11                        End If

                          'Inform user about current activity
12                        staStatus.Panels(2).Text = "Downloading thumbnail " & lngCurrentItem & " of " & lngNumberOfItemsFound & "..."
13                        frmProgress.pgbProgress.Value = lngCurrentItem * 50 \ lngNumberOfItemsFound + 50

14                        strSmallImageURL = objACDCore.Item(lngCurrentItem).strSmallImageURL
15                        If Not LenB(strSmallImageURL) = 0 Then
                              'Download thumbnail
16                            DownloadImage strSmallImageURL, picPicture
17                            If Not picPicture Is Nothing Then
18                                imlThumbnails.ListImages.Add Picture:=picPicture
19                                If lvwThumbnails.Icons Is Nothing Then
20                                    lvwThumbnails.Icons = imlThumbnails
21                                End If
                                  'Add thumbnail to list
22                                lvwThumbnails.ListItems.Add Key:="Result #" & lngCurrentItem, Icon:=imlThumbnails.ListImages.Count
23                                Set picPicture = Nothing
24                                If Not tabTabs.Tabs(2).Selected Then
                                      'Switch to result page
25                                    tabTabs.Tabs(2).Selected = True
26                                End If
27                              Else
28                                blnErrorsOccurred = True
29                            End If
30                        End If

                          'Check if there are queued message and parse them
31                        If GetQueueStatus(QS_ALLINPUT) Then
32                            DoEvents
33                        End If
34                    Next lngCurrentItem

                      'Inform user about possible errors
35                    If blnErrorsOccurred Then
36                        InformUser "Some thumbnails could not be downloaded successfully and were therefore ignored.", vbExclamation, "Warning"
37                    End If

38                  Else

39                    For lngCurrentItem = 1 To lngNumberOfItemsFound
40                        strSmallImageURL = objACDCore.Item(lngCurrentItem).strSmallImageURL
41                        If Not LenB(strSmallImageURL) = 0 Then
42                            If Not LenB(cdlSaveAs.FileName) = 0 Then
43                                If Not PathFileExists(cdlSaveAs.FileName) = 1 Then

                                      'Download image
44                                    Select Case AvailableImageSize(Choose(cboImageSize.ListIndex + 1, 3, 2, 1), lngCurrentItem, chkDegrade.Value = vbChecked)
                                        Case Large
45                                        lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(lngCurrentItem).strLargeImageURL, cdlSaveAs.FileName, 0, 0)
46                                      Case Medium
47                                        lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(lngCurrentItem).strMediumImageURL, cdlSaveAs.FileName, 0, 0)
48                                      Case Small
49                                        lngDownloadExitCode = URLDownloadToFile(0, objACDCore.Item(lngCurrentItem).strSmallImageURL, cdlSaveAs.FileName, 0, 0)
50                                      Case Else
51                                        InformUser "The desired image size is unfortunately not available.", vbInformation, "Information"
52                                        Exit For
53                                    End Select

                                      'Check if the image was downloaded successfully
54                                    If lngDownloadExitCode = 0 Then
55                                        Exit For
56                                      Else
57                                        InformUser "An unknown error occurred while attempting to load the desired image.", vbCritical, "Error"
58                                    End If

59                                  Else
60                                    InformUser "A file with the same name as the specified filename for the cover already exists. The cover was not saved.", vbExclamation, "Warning"
61                                End If
62                            End If
63                        End If
64                    Next lngCurrentItem
65                End If
66            End If
67        End If

68        If Not blnSilentMode Then
              'Reset status bar
69            staStatus.Panels(1).Text = "Idle"
70            staStatus.Panels(2).Text = vbNullString

71            Me.MousePointer = vbDefault

72            EnableControls

73            Unload frmProgress
74        End If

          'No more activity
75        blnWorking = False

76        On Error GoTo 0

77    Exit Sub

'Handle errors
ErrorHandler:
78        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: objACDCore_Done" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
79        End

End Sub

Private Sub objACDCore_Progress(ByVal lngCurrentPage As Long, ByVal lngTotalNumberOfPages As Long)

1         On Error GoTo ErrorHandler

2         If Not blnSilentMode Then
3             If blnProgressDataAvailable Then
                  'Inform user about current activity
4                 staStatus.Panels(2).Text = "Loading page " & lngCurrentPage & " of " & lngTotalNumberOfPages & "..."
5                 frmProgress.pgbProgress.Value = lngCurrentPage * 50 \ lngTotalNumberOfPages
6             End If
7         End If

8         On Error GoTo 0

9     Exit Sub

'Handle errors
ErrorHandler:
10        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: objACDCore_Progress" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
11        End

End Sub

Private Sub objACDCore_ProgressDataAvailable()

1         On Error GoTo ErrorHandler

2         If Not blnSilentMode Then
              'Precise progress data is now available
3             blnProgressDataAvailable = True

              'Set progress bar style back to normal
4             GetWindowLong frmProgress.pgbProgress.hWnd, GWL_STYLE
5             SetWindowLong frmProgress.pgbProgress.hWnd, GWL_STYLE, GetWindowLong(frmProgress.pgbProgress.hWnd, GWL_STYLE) - PBS_MARQUEE
6             SendMessageLong frmProgress.pgbProgress.hWnd, PBM_SETMARQUEE, False, 60
7         End If

8         On Error GoTo 0

9     Exit Sub

'Handle errors
ErrorHandler:
10        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: objACDCore_ProgressDataAvailable" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
11        End

End Sub

Private Sub objACDCore_WorkStarted()

1         On Error GoTo ErrorHandler

2         If Not blnSilentMode Then
3             DisableControls

4             tabTabs.Tabs(1).Selected = True

5             Me.MousePointer = vbArrowHourglass

              'Reset flags
6             lngCancel = 0
7             blnProgressDataAvailable = False
8             blnWorking = True

              'Reset list view and image list
9             lvwThumbnails.ListItems.Clear
10            Set lvwThumbnails.Icons = Nothing
11            imlThumbnails.ListImages.Clear

              'Inform user about current activity
12            staStatus.Panels(1).Text = "Working..."
13            staStatus.Panels(2).Text = "Querying..."

              'Set progress bar style to marquee
14            GetWindowLong frmProgress.pgbProgress.hWnd, GWL_STYLE
15            SetWindowLong frmProgress.pgbProgress.hWnd, GWL_STYLE, GetWindowLong(frmProgress.pgbProgress.hWnd, GWL_STYLE) + PBS_MARQUEE
16            SendMessageLong frmProgress.pgbProgress.hWnd, PBM_SETMARQUEE, True, 60
17            frmProgress.Show , Me
18        End If

19        On Error GoTo 0

20    Exit Sub

'Handle errors
ErrorHandler:
21        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: objACDCore_WorkStarted" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
22        End

End Sub

Private Sub tabTabs_Click()

1         On Error GoTo ErrorHandler

          'Toggle view
2         If tabTabs.SelectedItem.Index = 1 Then
3             cmdQuery.Default = True
4             picContainer(0).Visible = True
5             picContainer(1).Visible = False
6           Else
7             cmdSave.Default = True
8             picContainer(0).Visible = False
9             picContainer(1).Visible = True
10        End If

11        On Error GoTo 0

12    Exit Sub

'Handle errors
ErrorHandler:
13        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: tabTabs_Click" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
14        End

End Sub

Private Sub UnloadForms()

        Dim frmForm As Form

1         On Error GoTo ErrorHandler

          'Unload any loaded form
2         For Each frmForm In Forms
3             If Not frmForm.Name = "frmMain" Then
4                 Unload frmForm
5             End If
6         Next frmForm

7         On Error GoTo 0

8     Exit Sub

'Handle errors
ErrorHandler:
9         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Form.frm" & vbNewLine & "Form: frmMain" & vbNewLine & "Procedure: UnloadForms" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
10        End

End Sub
