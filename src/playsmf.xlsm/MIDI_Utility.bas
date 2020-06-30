Attribute VB_Name = "MIDI_Utility"
Option Explicit

'SysExはよくわからん...
'http://www.cs.k.tsukuba-tech.ac.jp/labo/koba/software/midi.html
Public Const GSReset As String = "F0 41 10 42 12 40 00 7F 00 41 F7"

'Master Volume   F0 7F 7F 04 01 00 xx F7 (xx: 設定する音量)
Function MasterVolume(ByVal val As Integer) As String
    MasterVolume = "F0 7F 7F 04 01 00 " & Right("0" & Hex(val), 2) & " F7"
End Function

'Use for Rhythm Part F0 41 10 42 12 [40 1x 15 yy] sum F7  (x: 対象とするチャンネル yy: 割り当てるドラムマップ)
Function UseForRhythmPart(ByVal val As Integer) As String
    UseForRhythmPart = "F0 41 10 42 12 40 1x 15 01 yy F7"
    If val = 10 Then
        val = 0
    End If
    UseForRhythmPart = Replace(UseForRhythmPart, "x", Hex(val))
    Dim sum As Integer
    sum = CInt("&H" & Right("0" & Mid(UseForRhythmPart, 16, 2), 2)) + CInt("&H" & Right("0" & Mid(UseForRhythmPart, 19, 2), 2)) + CInt("&H" & Right("0" & Mid(UseForRhythmPart, 22, 2), 2)) + CInt("&H" & Right("0" & Mid(UseForRhythmPart, 25, 2), 2))
    UseForRhythmPart = Replace(UseForRhythmPart, "yy", Right("0" & Hex(sum), 2))
End Function

'メッセージを送信する
Sub midiOutSendMsg(ByRef hmo As LongPtr, ByVal msg As String)
    
    Dim ret As Integer
    msg = Replace(msg, " ", "")
    
    If Left(msg, 1) = "F" Then 'Fから始まる場合のみLongMsgに送る
        Dim i As Long
        Dim lpMidiOutHdr As MIDIHDR
        Dim buffer() As Long
        
        ReDim buffer(0 To Len(msg) / 2 - 1)
        For i = 1 To Len(msg) / 2
            buffer(i - 1) = CLng("&H" & Mid(msg, 2 * i - 1, 2))
        Next
        
        lpMidiOutHdr.lpData = VarPtr(buffer(0))
        lpMidiOutHdr.dwBufferLength = UBound(buffer)
        lpMidiOutHdr.dwFlags = 0
        ret = midiOutPrepareHeader(hmo, lpMidiOutHdr, LenB(lpMidiOutHdr))
        ret = midiOutLongMsg(hmo, lpMidiOutHdr, LenB(lpMidiOutHdr))
        ret = midiOutUnprepareHeader(hmo, lpMidiOutHdr, LenB(lpMidiOutHdr))
        
    Else 'それ以外はShortMsg
        Dim temp As String
        temp = ""

        For i = 1 To Len(msg) / 2
            temp = temp + Mid(msg, (-2 * (i - 1)) + (Len(msg) - 1), 2)
        Next

        ret = midiOutShortMsg(hmo, CLng("&H" & temp))
    End If
    
End Sub


