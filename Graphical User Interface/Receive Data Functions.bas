Attribute VB_Name = "modReceiveData"
Option Explicit
DefLng A-Z

Public Const WM_COPYDATA As Long = &H4A

Public Type CopyDataStucture
   lngData As Long
   lngDataSize As Long
   lngDataPointer As Long
End Type

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal lngHandle As Long, ByVal strProperty As String) As Long

Public lngACDWindowProcess As Long
Public lngFirstACDWindowProcess As Long

Public Function ACDWindowProcess(ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long

        Dim abytParameters() As Byte
        Dim astrParameters() As String
        Dim udtCopyDataStructure As CopyDataStucture
          
1         On Error GoTo ErrorHandler

          'We received data from another instance
2         If lngMessage = WM_COPYDATA Then
3             CopyMemory udtCopyDataStructure, ByVal lngLParameter, LenB(udtCopyDataStructure)
4             ReDim abytParameters(udtCopyDataStructure.lngDataSize - 1)
5             CopyMemory abytParameters(0), ByVal udtCopyDataStructure.lngDataPointer, udtCopyDataStructure.lngDataSize
6             astrParameters = Split(StrConv(abytParameters, vbUnicode), "|")
              'If we are not busy or in silent mode, change parameters and perform query
7             If Not blnSilentMode And Not blnWorking Then
8                 frmMain.chkDegrade.Value = astrParameters(0)
9                 frmMain.cdlSaveAs.FileName = astrParameters(1)
10                frmMain.cboMediaType.Text = astrParameters(2)
11                frmMain.txtKeywords.Text = astrParameters(3)
12                lngPageLimit = astrParameters(4)
13                frmMain.cboServer.Text = astrParameters(5)
14                frmMain.cboImageSize.Text = astrParameters(6)
15                If Not LenB(frmMain.txtKeywords.Text) = 0 Then
16                    Query frmMain.txtKeywords.Text, frmMain.cboServer.Text, frmMain.cboMediaType.Text, lngPageLimit
17                End If
18            End If
19        End If

20        ACDWindowProcess = CallWindowProc(lngACDWindowProcess, lngHandle, lngMessage, lngWParameter, lngLParameter)

21        On Error GoTo 0

22    Exit Function

'Handle errors
ErrorHandler:
23        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Receive Data Functions.bas" & vbNewLine & "Module: modReceiveData" & vbNewLine & "Procedure: ACDWindowProcess" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
24        End

End Function

Public Function EnumerateWindowsProcess(ByVal lngHandle As Long, ByVal lngLParameter As Long) As Long

1         On Error GoTo ErrorHandler

2         If GetProp(lngHandle, "MW_ACD_{EDD1F962-EC56-40BA-B2C5-773F25EF26EA}") = 1 Then
3             EnumerateWindowsProcess = 0
4             lngFirstACDWindowProcess = lngHandle
5           Else
6             EnumerateWindowsProcess = 1
7         End If

8         On Error GoTo 0

9     Exit Function

'Handle errors
ErrorHandler:
10        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Receive Data Functions.bas" & vbNewLine & "Module: modReceiveData" & vbNewLine & "Procedure: EnumerateWindowsProcess" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
11        End

End Function
