Attribute VB_Name = "modMouseWheel"
Option Explicit
DefLng A-Z

Private Const WM_MOUSEWHEEL As Long = &H20A

Public lngOldPictureBoxWindowProcess As Long

Public Function PictureBoxWindowProcess(ByVal lngHandle As Long, ByVal lngMessage As Long, ByVal lngWParameter As Long, ByVal lngLParameter As Long) As Long

1         On Error GoTo ErrorHandler

          'If the preview form is visible and the mouse wheel was moved, scroll the picture box
2         If frmPreview.Visible Then
3             If lngMessage = WM_MOUSEWHEEL Then
4                 If lngWParameter \ &H10000 > 0 Then
5                     If (GetAsyncKeyState(vbKeyShift) And &H8000) = &H8000 Then
6                         If frmPreview.hsbHScroll.Enabled Then
7                             If Not frmPreview.hsbHScroll.Min > frmPreview.hsbHScroll.Value - frmPreview.hsbHScroll.SmallChange Then
8                                 frmPreview.hsbHScroll.Value = frmPreview.hsbHScroll.Value - frmPreview.hsbHScroll.SmallChange
9                               Else
10                                frmPreview.hsbHScroll.Value = frmPreview.hsbHScroll.Min
11                            End If
12                        End If
13                      Else
14                        If frmPreview.vsbVScroll.Enabled Then
15                            If Not frmPreview.vsbVScroll.Min > frmPreview.vsbVScroll.Value - frmPreview.vsbVScroll.SmallChange Then
16                                frmPreview.vsbVScroll.Value = frmPreview.vsbVScroll.Value - frmPreview.vsbVScroll.SmallChange
17                              Else
18                                frmPreview.vsbVScroll.Value = frmPreview.vsbVScroll.Min
19                            End If
20                        End If
21                    End If
22                  Else
23                    If (GetAsyncKeyState(vbKeyShift) And &H8000) = &H8000 Then
24                        If frmPreview.hsbHScroll.Enabled Then
25                            If Not frmPreview.hsbHScroll.Max < frmPreview.hsbHScroll.Value + frmPreview.hsbHScroll.SmallChange Then
26                                frmPreview.hsbHScroll.Value = frmPreview.hsbHScroll.Value + frmPreview.hsbHScroll.SmallChange
27                              Else
28                                frmPreview.hsbHScroll.Value = frmPreview.hsbHScroll.Max
29                            End If
30                        End If
31                      Else
32                        If frmPreview.vsbVScroll.Enabled Then
33                            If Not frmPreview.vsbVScroll.Max < frmPreview.vsbVScroll.Value + frmPreview.vsbVScroll.SmallChange Then
34                                frmPreview.vsbVScroll.Value = frmPreview.vsbVScroll.Value + frmPreview.vsbVScroll.SmallChange
35                              Else
36                                frmPreview.vsbVScroll.Value = frmPreview.vsbVScroll.Max
37                            End If
38                        End If
39                    End If
40                End If
41            End If
42        End If

43        PictureBoxWindowProcess = CallWindowProc(lngOldPictureBoxWindowProcess, lngHandle, lngMessage, lngWParameter, lngLParameter)

44        On Error GoTo 0

45    Exit Function

'Handle errors
ErrorHandler:
46        MsgBox "Fatal runtime error " & Err.Number & ": " & Err.Description & "." & vbNewLine & vbNewLine & "File: Mouse Wheel Functions.bas" & vbNewLine & "Module: modMouseWheel" & vbNewLine & "Procedure: PictureBoxWindowProcess" & vbNewLine & "Line: " & Erl & vbNewLine & "Timestamp: " & Format$(Now(), "YYYY-MM-DD HH:NN:SS") & vbNewLine & vbNewLine & "The application is going to be terminated." & vbNewLine & "Please copy this error report by pressing the CTRL+C keys and send it to the support for further investigation.", vbCritical, "Error"
47        End

End Function
