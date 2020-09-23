Attribute VB_Name = "modLimitSize"
Option Explicit
DefLng A-Z

Private Const WM_GETMINMAXINFO As Long = &H24

Private Type Point
    lngX As Long
    lngY As Long
End Type
Private Type MinMaxInfo
    udtReserved As Point
    udtMaximumSize As Point
    udtMaximumPosition As Point
    udtMinimumTrackSize As Point
    udtMaximumTrackSize As Point
End Type

Public lngOldFormWindowProcess As Long

Public Function FormWindowProcess(ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long

        Dim udtMinMaxInfo As MinMaxInfo

1         On Error GoTo ErrorHandler

          'Limit the minimum window size of the preview window
2         If lngMessage = WM_GETMINMAXINFO Then
3             CopyMemory udtMinMaxInfo, ByVal lngLParameter, LenB(udtMinMaxInfo)

              'Minimum window size
4             udtMinMaxInfo.udtMinimumTrackSize.lngX = 370
5             udtMinMaxInfo.udtMinimumTrackSize.lngY = 426

6             CopyMemory ByVal lngLParameter, udtMinMaxInfo, LenB(udtMinMaxInfo)

7             FormWindowProcess = 0
8           Else
9             FormWindowProcess = CallWindowProc(lngOldFormWindowProcess, lngHandle, lngMessage, lngWParameter, lngLParameter)
10        End If

11        On Error GoTo 0

12    Exit Function

'Handle errors
ErrorHandler:
13        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Size Limitation Functions.bas" & vbNewLine & "Module: modLimitSize" & vbNewLine & "Procedure: FormWindowProcess" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
14        End

End Function
