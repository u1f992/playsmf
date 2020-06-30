Attribute VB_Name = "WindowsAPI"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

'ウィンドウの「x」ボタンを無効化する
Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As Long
Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'midiOutOpen => MIDIデバイスを開く
Declare PtrSafe Function midiOutOpen Lib "winmm" (lphMidiOut As LongPtr, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
'<引数>
'lphMidiOut：MIDIデバイスのハンドル
'uDeviceID：デバイスID = -1(MIDIマッパー)
'dwCallback：コールバックパラメータ = 0
'dwInstance：コールバックに渡されるデータ = 0
'dwFlags：コールバックフラグ = 0
'<戻り値>
'MMRESULT エラー

'UINT midiOutOpen(
'  LPHMIDIOUT lphmo,
'  UINT uDeviceID,
'  DWORD dwCallback,
'  DWORD dwCallbackInstance,
'  DWORD dwFlags
');


'midiOutShortMsg => システムエクスクルーシブおよびストリームメッセージ以外のMIDIメッセージを送信する
Declare PtrSafe Function midiOutShortMsg Lib "winmm" (ByVal hMidiOut As LongPtr, ByVal dwMsg As Long) As Long
'<引数>
'hMidiOut：MIDIデバイスのハンドル
'dwMsg：音階
'第1バイト... MIDIステータスバイト
'第2バイト... MIDIデータ1バイト目
'第3バイト... MIDIデータ2バイト目
'第4バイト... 使用されません
'<戻り値>
'MMRESULT エラー

'MMRESULT midiOutShortMsg(
'  HMIDIOUT hmo,
'  DWORD dwMsg
');

 
'midiOutReset => MIDI 出力デバイスのすべてのチャンネルのノートをオフにする。
Declare PtrSafe Function midiOutReset Lib "winmm" (ByVal hMidiOut As LongPtr) As Long
'<引数>
'hMidiOut：MIDIデバイスのハンドル
'<戻り値>
'MMRESULT エラー

'MRESULT midiOutReset(
'    HMIDIOUT hmo   // MIDI出力デバイスのハンドル
');


'midiOutClose => MIDIデバイスを閉じる
Declare PtrSafe Function midiOutClose Lib "winmm" (ByVal hMidiOut As LongPtr) As Long
'<引数>
'hMidiOut：MIDIデバイスのハンドル
'<戻り値>
'MMRESULT エラー

'MMRESULT midiOutClose(
'  hMidiOut hmo
');


'MIDIHDR構造体   4*6+10+文字列の長さ(LenB)バイト
Type MIDIHDR
    lpData          As LongPtr '実際のMIDIデータ
    dwBufferLength  As Long
    dwBytesRecorded As Long
    dwUser          As Long
    dwFlags         As Long
    lpNext          As Long
    Reserved        As Long
End Type
'lpData : MIDI データを格納したバッファのアドレスが格納されます
'dwBufferLength : データバッファのサイズが格納されます
'dwBytesRecorded : バッファ中の実際のデータサイズが格納されます。 dwBufferLength メンバで指定された値以下でなければなりません
'dwUser : カスタムユーザーデータが格納されます
'dwFlags :バッファに関する情報のフラグが格納されます。必ず 0 に設定する
'lpNext : 使用不可
'Reserved : 使用不可
'dwOffset : コールバック処理時のバッファのオフセットが格納されます
'dwReserved : 使用不可

'midiOutPrepareHeader => MIDIシステム排他バッファを準備する
Declare PtrSafe Function midiOutPrepareHeader Lib "winmm" (ByVal hmo As LongPtr, ByRef lpMidiOutHdr As MIDIHDR, ByVal cbMidiOutHdr As Long) As Long
'<引数>
'hmo : MIDI 出力デバイスのハンドルを指定する
'lpMidiOutHdr : 準備するバッファを識別するMIDIHDR構造体のアドレスを指定する
'cbMidiOutHdr : MIDIHDR 構造体のサイズをバイト単位で指定する
'<戻り値>
'MMSYSERRエラー
'
'lpDataにMIDIデータをセット、dwBufferLengthに構造体サイズをセット、dwFlagsに0をセットしてから使う

'midiOutLongMsg =>指定された MIDI 出力デバイスにシステム排他 MIDI メッセージを送信する
Declare PtrSafe Function midiOutLongMsg Lib "winmm" (ByVal hmo As LongPtr, ByRef lpMidiOutHdr As MIDIHDR, ByVal cbMidiOutHdr As Long) As Long
'<引数>
'hmo : MIDI出力デバイスのハンドルを指定する
'lpMidiOutHdr : MIDIバッファを識別するMIDIHDR構造体のアドレスを指定する。
'cbMidiOutHdr:  MIDIHDR構造体のサイズをバイト単位で指定する｡

Declare PtrSafe Function midiOutUnprepareHeader Lib "winmm" (ByVal hmo As LongPtr, lpMidiOutHdr As MIDIHDR, ByVal cbMidiOutHdr As Long) As Long
'<引数>
'hmo : MIDI 出力デバイスのハンドルを指定する
'lpMidiOutHdr : クリーンアップするバッファを識別するMIDIHDR構造体のアドレスを指定する
'cbMidiOutHdr : MIDIHDR 構造体のサイズをバイト単位で指定する
