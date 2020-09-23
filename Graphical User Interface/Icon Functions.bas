Attribute VB_Name = "modIcon"
Option Explicit
DefLng A-Z

Private Const GW_OWNER As Long = 4
Private Const ICON_BIG As Long = 1
Private Const ICON_SMALL As Long = 0
Private Const IMAGE_ICON As Long = 1
Private Const LR_SHARED As Long = &H8000
Private Const SM_CXICON As Long = 11
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYICON As Long = 12
Private Const SM_CYSMICON As Long = 50
Private Const WM_SETICON As Long = &H80

Private Declare Function GetSystemMetrics Lib "user32" (ByVal lngIndex As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal lngHandle As Long, ByVal lngRelationship As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal lngInstance As Long, ByVal strImage As String, ByVal lngType As Long, ByVal lngX As Long, ByVal lngY As Long, ByVal lngMode As Long) As Long

Public Sub SetIcon(ByVal lngWindowHandle As Long, ByVal strIcon As String)

        Dim lngHandle As Long
        Dim lngLargeIcon As Long
        Dim lngParentHandle As Long
        Dim lngSmallIcon As Long

1         On Error GoTo ErrorHandler

          'Find parent window
2         lngHandle = lngWindowHandle
3         lngParentHandle = lngHandle
4         Do While Not lngHandle = 0
5             lngHandle = GetWindow(lngHandle, GW_OWNER)
6             If Not lngHandle = 0 Then
7                 lngParentHandle = lngHandle
8             End If
9         Loop

          'Set large icon (usually 32x32)
10        lngLargeIcon = LoadImage(App.hInstance, strIcon, IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_SHARED)
11        SendMessageLong lngParentHandle, WM_SETICON, ICON_BIG, lngLargeIcon
12        SendMessageLong lngWindowHandle, WM_SETICON, ICON_BIG, lngLargeIcon

          'Set small icon (usually 16x16)
13        lngSmallIcon = LoadImage(App.hInstance, strIcon, IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_SHARED)
14        SendMessageLong lngParentHandle, WM_SETICON, ICON_SMALL, lngSmallIcon
15        SendMessageLong lngWindowHandle, WM_SETICON, ICON_SMALL, lngSmallIcon

16        On Error GoTo 0

17    Exit Sub

'Handle errors
ErrorHandler:
18        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Icon Functions.bas" & vbNewLine & "Module: modIcon" & vbNewLine & "Procedure: SetIcon" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
19        End

End Sub
