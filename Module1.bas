Attribute VB_Name = "Module1"
Option Explicit

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long

Function PlayMidi()
   Dim MidiToOpen As String
   Midi "close all"
      MidiToOpen = App.Path & "\simble.mid" '"E:\R-Soft\Software\Sourse\Downloads\LCDAlarmClock\LCD Alarm Clock\Alarm Music\Chimes.mid"
      OpenMidi (MidiToOpen)
      Midi "play med"

End Function

Function Midi(sCommand As String) As String
    Dim Buff As String * 255
    Call mciSendString(sCommand, Buff, 255, Form1.hWnd)
End Function

Function OpenMidi(fName As String) As Boolean
    OpenMidi = CBool(Val(Midi("open """ & fName & """ alias med")))
End Function


