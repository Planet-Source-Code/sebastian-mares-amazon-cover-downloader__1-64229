VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
DefLng A-Z

Private Const CW_USEDEFAULT As Long = &H80000000
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
Private Const TTDT_AUTOPOP As Long = 2
Private Const TTDT_INITIAL As Long = 3
Private Const TTF_IDISHWND As Long = &H1
Private Const TTF_SUBCLASS As Long = &H10
Private Const TTM_ADDTOOL As Long = WM_USER + 4
Private Const TTM_SETDELAYTIME As Long = WM_USER + 3
Private Const TTM_SETTIPBKCOLOR As Long = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR As Long = WM_USER + 20
Private Const TTM_SETTITLE As Long = WM_USER + 32
Private Const TTM_UPDATETIPTEXTA As Long = WM_USER + 12
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_NOPREFIX As Long = &H2

Public Enum IconType
    None = 0
    Information = 1
    Warning = 2
    Error = 3
End Enum
Public Enum ToolTipStyle
    Standard = 1
    Balloon = 2
End Enum
#If False Then
Private Balloon As Long
Private Error As Long
Private Information As Long
Private None As Long
Private Standard As Long
Private Warning As Long
#End If

Private Type Rectangle
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
End Type
Private Type ToolTipStructure
    lngSize As Long
    lngFlags As Long
    lngHandle As Long
    lngID As Long
    udtRect As Rectangle
    lngInstance As Long
    strText As String
    lngParameter As Long
End Type

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal lngExtendedStyle As Long, ByVal strClassName As String, ByVal strCaption As String, ByVal lngStyle As Long, ByVal lngX As Long, ByVal lngY As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal lngParentHandle As Long, ByVal lngMenu As Long, ByVal lngInstance As Long, anyParameter As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal lngHandle As Long) As Long
Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal clrColor As OLE_COLOR, ByVal lngPalette As Long, lngColor As Long) As Long

Private enuIconType As IconType
Private enuToolTipStyle As ToolTipStyle
Private lngToolTipHandle As Long
Private strText As String
Private strTitle As String
Private udtToolTipStructure As ToolTipStructure

Private Sub Class_Initialize()

1         On Error GoTo ErrorHandler

          'Apply default style and icon
2         Style = Standard
3         Icon = Information

4         On Error GoTo 0

5     Exit Sub

'Handle errors
ErrorHandler:
6         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Class_Initialize" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
7         End

End Sub

Private Sub Class_Terminate()

1         On Error GoTo ErrorHandler

          'Destroy existing tool tip
2         Destroy

3         On Error GoTo 0

4     Exit Sub

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Class_Terminate" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Sub

Public Function Create(ByVal lngParentHandle As Long) As Boolean

        Dim lngColor As Long
        Dim lngWindowStyle As Long

1         On Error GoTo ErrorHandler

          'Destroy any existing tool tip
2         If Not lngToolTipHandle = 0 Then
3             DestroyWindow lngToolTipHandle
4         End If

          'Apply desired window style
5         If enuToolTipStyle = Balloon Then
6             lngWindowStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX Or TTS_BALLOON
7           Else
8             lngWindowStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
9         End If

          'Create window
10        lngToolTipHandle = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lngWindowStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
11        With udtToolTipStructure
12            .lngFlags = TTF_SUBCLASS Or TTF_IDISHWND
13            .lngHandle = lngParentHandle
14            .lngID = lngParentHandle
15            .lngInstance = App.hInstance
16            .lngSize = Len(udtToolTipStructure)
17        End With

          'Register tool with tool tip
18        SendMessage lngToolTipHandle, TTM_ADDTOOL, 0, udtToolTipStructure

          'Supply title and icon)
19        If Not LenB(strTitle) = 0 Or Not enuIconType = None Then
20            SendMessage lngToolTipHandle, TTM_SETTITLE, enuIconType, ByVal strTitle
21        End If

          'Apply foreground and background colors
22        TranslateColor vbInfoText, 0, lngColor
23        SendMessageLong lngToolTipHandle, TTM_SETTIPTEXTCOLOR, lngColor, 0
24        TranslateColor vbInfoBackground, 0, lngColor
25        SendMessageLong lngToolTipHandle, TTM_SETTIPBKCOLOR, lngColor, 0

          'Set delay and time-to-live timers
26        SendMessageLong lngToolTipHandle, TTM_SETDELAYTIME, TTDT_AUTOPOP, 10000
27        SendMessageLong lngToolTipHandle, TTM_SETDELAYTIME, TTDT_INITIAL, 500

28        On Error GoTo 0

29    Exit Function

'Handle errors
ErrorHandler:
30        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Create" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
31        End

End Function

Public Sub Destroy()

1         On Error GoTo ErrorHandler

          'Destroy existing tool tip
2         If Not lngToolTipHandle = 0 Then
3             DestroyWindow lngToolTipHandle
4         End If

5         On Error GoTo 0

6     Exit Sub

'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Destroy" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End

End Sub

Public Property Let Icon(ByVal enuValue As IconType)

1         On Error GoTo ErrorHandler

          'Set icon type
2         enuIconType = enuValue
3         If Not lngToolTipHandle = 0 Then
4             If Not LenB(strTitle) = 0 Then
5                 If Not enuIconType = None Then
6                     SendMessage lngToolTipHandle, TTM_SETTITLE, enuIconType, ByVal strTitle
7                 End If
8             End If
9         End If

10        On Error GoTo 0

11    Exit Property

'Handle errors
ErrorHandler:
12        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Icon" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
13        End

End Property

Public Property Get Icon() As IconType

1         On Error GoTo ErrorHandler

          'Get icon type
2         Icon = enuIconType

3         On Error GoTo 0

4     Exit Property

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Icon" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Property

Public Property Let Style(ByVal enuValue As ToolTipStyle)

1         On Error GoTo ErrorHandler

          'Set style
2         enuToolTipStyle = enuValue

3         On Error GoTo 0

4     Exit Property

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Style" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Property

Public Property Get Style() As ToolTipStyle

1         On Error GoTo ErrorHandler

          'Get style
2         Style = enuToolTipStyle

3         On Error GoTo 0

4     Exit Property

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Style" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Property

Public Property Let Title(ByVal strValue As String)

1         On Error GoTo ErrorHandler

          'Set title
2         strTitle = strValue
3         If Not lngToolTipHandle = 0 Then
4             If Not LenB(strTitle) = 0 Then
5                 If Not enuIconType = None Then
6                     SendMessage lngToolTipHandle, TTM_SETTITLE, enuIconType, ByVal strTitle
7                 End If
8             End If
9         End If

10        On Error GoTo 0

11    Exit Property

'Handle errors
ErrorHandler:
12        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Title" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
13        End

End Property

Public Property Get Title() As String

1         On Error GoTo ErrorHandler

          'Get title
2         Title = udtToolTipStructure.strText

3         On Error GoTo 0

4     Exit Property

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: Title" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Property

Public Property Let ToolTip(ByVal strValue As String)

1         On Error GoTo ErrorHandler

          'Set tool tip text
2         strText = strValue
3         udtToolTipStructure.strText = strText
4         If lngToolTipHandle <> 0 Then
5             SendMessage lngToolTipHandle, TTM_UPDATETIPTEXTA, 0, udtToolTipStructure
6         End If

7         On Error GoTo 0

8     Exit Property

'Handle errors
ErrorHandler:
9         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: ToolTip" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
10        End

End Property

Public Property Get ToolTip() As String

1         On Error GoTo ErrorHandler

          'Get tool tip text
2         ToolTip = strText

3         On Error GoTo 0

4     Exit Property

'Handle errors
ErrorHandler:
5         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: ToolTip Functions.cls" & vbNewLine & "Class Module: clsToolTip" & vbNewLine & "Procedure: ToolTip" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
6         End

End Property
