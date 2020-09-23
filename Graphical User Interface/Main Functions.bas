Attribute VB_Name = "modMain"
Option Explicit
DefLng A-Z

Public Const GWL_WNDPROC As Long = -4
Private Const STD_ERROR_HANDLE As Long = -12
Public Const WM_USER As Long = &H400

Public Enum ImageSizes
    None = 0
    Small = 1
    Medium = 2
    Large = 3
End Enum
#If False Then
Private Large As Long
Private Medium As Long
Private None As Long
Private Small As Long
#End If

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lngPreviousWindowFunction As Long, ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lngClassIDStringPointer As Long, anyClassID As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (anyDestination As Any, anySource As Any, ByVal lngBytesCopied As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal lngKey As Long) As Integer
Private Declare Function GetStdHandle Lib "kernel32" (ByVal lngStdHandle As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal lngHandle As Long, ByVal lngIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByRef anyLParameter As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal lngHandle As Long, ByVal lngIndex As Long, ByVal lngNewIndex As Long) As Long
Private Declare Function OleLoadPicturePath Lib "oleaut32" (ByVal lngPathPointer As Long, ByVal lngIUnknownPointer As Long, ByVal lngReserved As Long, ByVal clrTransparent As OLE_COLOR, ByRef anyInterfaceID As Any, ByRef anyReturn As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal lngFilePointer As Long, ByVal strData As String, ByVal lngNumberOfBytesToWrite As Long, ByRef lngNumberOfBytesWritten As Long, ByRef anyOverlapped As Any) As Long

Public blnSilentMode As Boolean
Public blnWorking As Boolean
Public lngPageLimit As Long

Public Function AvailableImageSize(ByVal enuDesiredImageSize As ImageSizes, ByVal lngImageNumber As Long, Optional ByVal blnAllowDegrading As Boolean = True) As ImageSizes

1         On Error GoTo ErrorHandler

          'Check if the desired image size for the supplied image index is available and fall back to smaller sizes if allowed
2         If enuDesiredImageSize > None Then
3             If ImageSizeAvailable(enuDesiredImageSize, lngImageNumber) Then
4                 AvailableImageSize = enuDesiredImageSize
5               Else
6                 If blnAllowDegrading Then
7                     AvailableImageSize = AvailableImageSize(enuDesiredImageSize - 1, lngImageNumber)
8                 End If
9             End If
10        End If

11        On Error GoTo 0

12    Exit Function

'Handle errors
ErrorHandler:
13        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: AvailableImageSize" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
14        End

End Function

Public Function DebugMode() As Boolean

1         On Error GoTo ErrorHandler

2         If App.EXEName = "Graphical User Interface" Then
3             DebugMode = True
4         End If

5         On Error GoTo 0

6     Exit Function

'Handle errors
ErrorHandler:
7         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: DebugMode" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
8         End

End Function

Public Sub DownloadImage(ByVal strURL As String, ByRef picPicture As Picture)

        Dim IPictureInterface(15) As Byte

1         On Error GoTo ErrorHandler

          'Load a local or remote image into the supplied IPicture object
2         CLSIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IPictureInterface(0)
3         If Not OleLoadPicturePath(StrPtr(strURL), 0, 0, 0, IPictureInterface(0), picPicture) = 0 Then
4             InformUser "Failed to load the thumbnail " & Chr$(34) & strURL & Chr$(34) & "."
5         End If

6         On Error GoTo 0

7     Exit Sub

'Handle errors
ErrorHandler:
8         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: DownloadImage" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
9         End

End Sub

Private Function ImageSizeAvailable(ByVal enuDesiredImageSize As ImageSizes, ByVal lngImageNumber As Long) As Boolean

1         On Error GoTo ErrorHandler

          'Verify if the desired image size for the supplied image index is available
2         If enuDesiredImageSize = Large Then
3             ImageSizeAvailable = (Not frmMain.objACDCore.Item(lngImageNumber).strLargeImageURL = vbNullString)
4           ElseIf enuDesiredImageSize = Medium Then
5             ImageSizeAvailable = (Not frmMain.objACDCore.Item(lngImageNumber).strMediumImageURL = vbNullString)
6           ElseIf enuDesiredImageSize = Small Then
7             ImageSizeAvailable = (Not frmMain.objACDCore.Item(lngImageNumber).strSmallImageURL = vbNullString)
8         End If

9         On Error GoTo 0

10    Exit Function

'Handle errors
ErrorHandler:
11        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: ImageSizeAvailable" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
12        End

End Function

Public Function InformUser(ByVal strText As String, Optional ByVal mbsMessageBoxStyle As VbMsgBoxStyle, Optional ByVal strMessageBoxTitle As String) As VbMsgBoxResult

1         On Error GoTo ErrorHandler

          'Depending on the operation mode, output message to StdErr or message box
2         If Not blnSilentMode Then
3             InformUser = MsgBox(strText, mbsMessageBoxStyle, strMessageBoxTitle)
4           Else
5             WriteToStdErr strText
6             If mbsMessageBoxStyle = vbYesNo Or mbsMessageBoxStyle = vbYesNoCancel Then
7                 InformUser = vbYes
8               Else
9                 InformUser = vbOK
10            End If
11        End If

12        On Error GoTo 0

13    Exit Function

'Handle errors
ErrorHandler:
14        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: InformUser" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
15        End

End Function

Private Sub LoadDefaultSettings()

1         On Error GoTo ErrorHandler

          'Load default settings which could not be set in design mode
2         frmMain.cboServer.Text = "amazon.com"
3         frmMain.cboImageSize.Text = "Large"

4         On Error GoTo 0

5     Exit Sub

'Handle errors
ErrorHandler:
6         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: LoadDefaultSettings" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
7         End

End Sub

Public Function ParseCommandLineArguments() As Boolean

        Dim lngCurrentArgument As Long
        Dim strArgumentAndValue() As String
        Dim strArguments() As String

1         On Error Resume Next

              'Load default settings
2             LoadDefaultSettings

              'Fill array with command line arguments (one argument/value pair per index)
3             strArguments = Split(Command$, "--acd-", , vbTextCompare)

              'Evaluate each command line argument
4             For lngCurrentArgument = LBound(strArguments) To UBound(strArguments)
5                 If Not LenB(strArguments(lngCurrentArgument)) = 0 Then

                      'Split the command line argument and value
6                     strArgumentAndValue = Split(strArguments(lngCurrentArgument), "=", , vbTextCompare)

                      'Evaluate arguments and and apply setting
7                     Select Case strArgumentAndValue(0)
                        Case "degrade"
8                         frmMain.chkDegrade.Value = Trim$(strArgumentAndValue(1))
9                       Case "filename"
10                        frmMain.cdlSaveAs.FileName = Trim$(strArgumentAndValue(1))
11                      Case "index"
12                        frmMain.cboMediaType.Text = Trim$(strArgumentAndValue(1))
13                      Case "keywords"
14                        frmMain.txtKeywords.Text = Trim$(strArgumentAndValue(1))
15                      Case "pagelimit"
16                        lngPageLimit = Trim$(strArgumentAndValue(1))
17                      Case "server"
18                        frmMain.cboServer.Text = Trim$(strArgumentAndValue(1))
19                      Case "silent"
20                        blnSilentMode = Trim$(strArgumentAndValue(1))
21                      Case "size"
22                        frmMain.cboImageSize.Text = Trim$(strArgumentAndValue(1))
23                      Case Else
24                        Err.Raise 1
25                    End Select

26                End If
27            Next lngCurrentArgument

              'Return "True" if no runtime errors occurred
28            If Err.Number = 0 Then
29                ParseCommandLineArguments = True
30            End If

31        On Error GoTo 0

End Function

Public Sub Query(ByVal strQueryString As String, Optional ByVal strLocaleURL As String = "amazon.com", Optional ByVal strMediaType As String = "Music", Optional ByVal lngNumberOfPagesToProcess As Long)

        Dim lngQueryResult As Long

1         On Error GoTo ErrorHandler

          'Do basic checks
2         If LenB(strQueryString) = 0 Then
3             InformUser "You did not enter any keywords for the query. Please rectify this problem and try again.", vbExclamation, "Warning"
4             Exit Sub
5         End If
6         If LenB(strLocaleURL) = 0 Then
7             InformUser "You did not select an AWS server for the query. Please rectify this problem and try again.", vbExclamation, "Warning"
8             Exit Sub
9         End If
10        If LenB(strMediaType) = 0 Then
11            InformUser "You did not enter a media type for the query. Please rectify this problem and try again.", vbExclamation, "Warning"
12            Exit Sub
13        End If

          'Perform query
14        lngQueryResult = frmMain.objACDCore.Query(strQueryString, , strLocaleURL, strMediaType, lngNumberOfPagesToProcess)

          'Check response and show possible errors
15        If Not lngQueryResult = 200 And Not lngQueryResult = 600 Then
16            If blnSilentMode Then
17                InformUser "The query terminated with code " & lngQueryResult & ": " & frmMain.objACDCore.StatusCodeDescription(lngQueryResult), vbCritical, "Error"
18              Else
19                If InformUser("The query terminated with code " & lngQueryResult & ": " & frmMain.objACDCore.StatusCodeDescription(lngQueryResult) & vbNewLine & vbNewLine & "Would you like to obtain additional error information?", vbQuestion + vbYesNo, "Question") = vbYes Then
20                    Select Case lngQueryResult
                        Case 402
21                        InformUser "Amazon Web Services returned error code " & frmMain.objACDCore.AWSError.strCode & ":" & vbNewLine & vbNewLine & frmMain.objACDCore.AWSError.strMessage, vbCritical, "Error"
22                      Case 500 - 502
23                        InformUser frmMain.objACDCore.RuntimeError.strSource & " generated runtime error " & frmMain.objACDCore.RuntimeError.lngNumber & " on line " & frmMain.objACDCore.RuntimeError.lngLine & ":" & vbNewLine & vbNewLine & frmMain.objACDCore.RuntimeError.strDescription, vbCritical, "Error"
24                      Case Else
25                        If Not frmMain.objACDCore.XMLError.errorCode = 0 Then
26                            InformUser "The XML parser failed with error code " & frmMain.objACDCore.XMLError.errorCode & " while attempting to process document " & frmMain.objACDCore.XMLError.url & " (absolute position: " & frmMain.objACDCore.XMLError.filepos & ", line: " & frmMain.objACDCore.XMLError.Line & ", character: " & frmMain.objACDCore.XMLError.linepos & "):" & vbNewLine & vbNewLine & frmMain.objACDCore.XMLError.reason, vbCritical, "Error"
27                          Else
28                            InformUser "No additional error information available.", vbInformation, "Information"
29                        End If
30                    End Select
31                End If
32            End If
33            Exit Sub
34        End If

35        On Error GoTo 0

36    Exit Sub

'Handle errors
ErrorHandler:
37        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: Query" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
38        End

End Sub

Private Sub WriteToStdErr(ByVal strText As String)

1         On Error GoTo ErrorHandler

          'Format message and write it to StdErr
2         strText = "[" & Now & "]" & String$(5, " ") & Command$ & String$(5, " ") & strText & vbNewLine
3         WriteFile GetStdHandle(STD_ERROR_HANDLE), strText, Len(strText), 0, ByVal 0

4         On Error GoTo 0

5     Exit Sub

'Handle errors
ErrorHandler:
6         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Main Functions.bas" & vbNewLine & "Module: modMain" & vbNewLine & "Procedure: WriteToStdErr" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
7         End

End Sub
