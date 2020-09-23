Attribute VB_Name = "modSystemMenu"
Option Explicit
DefLng A-Z

Public lngAboutMenuID As Long
Public lngOldSystemMenuWindowProcess As Long

Public Function SystemMenuWindowProcess(ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long

1         On Error GoTo ErrorHandler

          'Check if the "About" menu item was clicked
2         If lngWParameter = lngAboutMenuID Then
3             frmAbout.Show vbModal, frmMain
4         End If

5         SystemMenuWindowProcess = CallWindowProc(lngOldSystemMenuWindowProcess, lngHandle, lngMessage, lngWParameter, lngLParameter)

6         On Error GoTo 0

7     Exit Function

'Handle errors
ErrorHandler:
8         MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: System Menu Functions.bas" & vbNewLine & "Module: modSystemMenu" & vbNewLine & "Procedure: SystemMenuWindowProcess" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
9         End

End Function
