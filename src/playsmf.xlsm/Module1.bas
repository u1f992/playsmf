Attribute VB_Name = "Module1"
Option Explicit

'1マイクロ秒間に増えるカウント数を取得する
Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef freq As LongLong) As Long
'システムが起動してからのカウント数を取得する
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef procTime As LongLong) As Long
Public freq As LongLong

Type midiHeaderChunk
    ChunkID As String
    DataLength As Long
    Format As Long
    Tracks As Long
    Division As Long
End Type

Type midiTrackChunk
    ChunkID As String
    DataLength As Long
    Data() As Byte
End Type

Sub Main()
    
    DeleteAllSheets
    
    QueryPerformanceFrequency freq
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filename As String
    filename = Application.GetOpenFilename("Standard MIDI File,*.mid")
    If filename = "False" Or fso.FileExists(filename) = False Then
        Exit Sub
    End If
    
    Dim smf() As Byte
    BinaryReader smf(), filename
    
    Dim offset As Long                           '読み込み位置
    offset = 0
    
    Dim tempo As Long                            'Set Tempo : 4分音符の長さをマイクロ秒で表す
    tempo = 500000
    
    Dim templen As Long
    Dim tempoffset As Long
    Dim trackoffset As Long                      '.Data 読み込み位置
    Dim trackfirst As Boolean                    'トラック開始時にTrue
    
    Dim count As Long                            '行
    count = 1
    
    Dim smfHeader As midiHeaderChunk
    smfHeader.ChunkID = GetValueByString(smf, offset, 4)
    smfHeader.DataLength = GetValueByLong(smf, offset, 4)
    smfHeader.Format = GetValueByLong(smf, offset, 2)
    smfHeader.Tracks = GetValueByLong(smf, offset, 2)
    smfHeader.Division = GetValueByLong(smf, offset, 2)
    
    Dim tempmsg As String
    
    tempmsg = "Header Chunk" & vbNewLine & _
              " - Chunk ID" & vbTab & vbTab & ": " & smfHeader.ChunkID & vbNewLine & _
              " - Data Length" & vbTab & ": " & smfHeader.DataLength & vbNewLine & _
              " - Format" & vbTab & vbTab & ": " & smfHeader.Format & vbNewLine & _
              " - Tracks" & vbTab & vbTab & ": " & smfHeader.Tracks & vbNewLine & _
              " - Division" & vbTab & vbTab & ": " & smfHeader.Division & vbNewLine
    
    Dim smfTrack() As midiTrackChunk
    ReDim smfTrack(0 To smfHeader.Tracks - 1)
    
    Dim t As Double
    t = Timer
    
    For j = 0 To smfHeader.Tracks - 1
        
        'トラック番号のワークシートに記録する。
        'Worksheets.Add(after:=Worksheets(Worksheets.count)).Name = j
        
        With smfTrack(j)
    
            'offsetからLength分、smfTrack.Data()に読み込む。
            .ChunkID = GetValueByString(smf, offset, 4)
            .DataLength = GetValueByLong(smf, offset, 4)
            
            ReDim .Data(0 To .DataLength - 1)
            For i = 0 To UBound(.Data)
                .Data(i) = smf(offset)
                offset = offset + 1
            Next
            
            trackoffset = 0
            'count = 1
            trackfirst = True
            
            Do While True
                
                '1列目 : デルタタイム
                If trackfirst = False Then
                    ActiveSheet.Cells(count, 1).Value = ActiveSheet.Cells(count - 1, 1).Value + GetDeltaTime(.Data, trackoffset)
                Else
                    ActiveSheet.Cells(count, 1).Value = GetDeltaTime(.Data, trackoffset)
                    trackfirst = False
                End If
                
                ActiveSheet.Cells(count, 2).Value = GetValueByLong(.Data, trackoffset, 1)
                
                k = 3

                Select Case ActiveSheet.Cells(count, k - 1).Value
            
                Case &HFF 'FF nn len ~
                    
                    ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                    k = k + 1

                    tempoffset = trackoffset

                    templen = GetDeltaTime(.Data, trackoffset)
                    For l = trackoffset - tempoffset To 1 Step -1
                        ActiveSheet.Cells(count, k).Value = .Data(trackoffset - l)
                        k = k + 1
                    Next

                    For l = 1 To templen
                        ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                        k = k + 1
                    Next l

                Case &HF0, &HF7 'FO len ~ F7

                    tempoffset = trackoffset

                    templen = GetDeltaTime(.Data, trackoffset)
                    For l = trackoffset - tempoffset To 1 Step -1
                        ActiveSheet.Cells(count, k).Value = .Data(trackoffset - l)
                        k = k + 1
                    Next
                    For l = 1 To templen
                        ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                        k = k + 1
                    Next l

                Case Else
                    
                    If ActiveSheet.Cells(count, k - 1).Value > &H7F And ActiveSheet.Cells(count, k - 1).Value > &HBF And ActiveSheet.Cells(count, k - 1).Value < &HE0 Then 'Cn と Dnは2Byte
                        ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                        k = k + 1
                    ElseIf ActiveSheet.Cells(count, k - 1).Value > &H7F Then 'それ以外は3Byte
                        For l = 1 To 2
                            ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                            k = k + 1
                        Next
                    Else                         '&H7Fより小さい場合、ステータスバイトは直前のものと同一
                        ActiveSheet.Cells(count, k).Value = ActiveSheet.Cells(count, k - 1).Value
                        ActiveSheet.Cells(count, k - 1).Value = ActiveSheet.Cells(count - 1, k - 1).Value
                        k = k + 1
                        ActiveSheet.Cells(count, k).Value = GetValueByLong(.Data, trackoffset, 1)
                        k = k + 1
                    End If

                End Select

                If k = 5 Then
                    If ActiveSheet.Cells(count, k - 1).Value = &H0 And ActiveSheet.Cells(count, k - 2).Value = &H2F And ActiveSheet.Cells(count, k - 3).Value = &HFF Or trackoffset = .DataLength Then
                        Exit Do
                    End If
                End If

                count = count + 1
                
                Application.StatusBar = ProgressBar((j / smfHeader.Tracks) + (((j + 1) / smfHeader.Tracks) - (j / smfHeader.Tracks)) * (trackoffset / .DataLength))
                DoEvents
                
            Loop
            
            tempmsg = tempmsg & "Track Chunk(" & j & ")" & vbNewLine & _
                      " - Chunk ID" & vbTab & vbTab & ": " & .ChunkID & vbNewLine & _
                      " - Data Length" & vbTab & ": " & .DataLength & vbNewLine
    
        End With
        
        DoEvents
        
    Next
    
    Application.StatusBar = ProgressBar(1)
    'Debug.Print
    
    Application.StatusBar = "Sorting..."
    ActiveSheet.UsedRange.Sort ActiveSheet.Columns(1)
    Application.StatusBar = "Sorted"
    
    Debug.Print tempmsg
    tempmsg = ""
    
    '_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
    
    Dim hmo As LongPtr
    Debug.Print "midiOutOpen" & vbTab & vbTab & ": " & midiOutOpen(hmo, -1, 0, 0, 0)
    
    Application.StatusBar = "Reade to Play"
    MsgBox "再生の準備が整いました。"
    WaitDoEvents 1000000
    
    For i = 1 To ActiveSheet.UsedRange.Rows.count
        
        ActiveSheet.Rows(i).Select
        
        tempmsg = ""
        j = 2
        
        Do While ActiveSheet.Cells(i, j).Value <> ""
            tempmsg = tempmsg & Right("0" & Hex(ActiveSheet.Cells(i, j).Value), 2)
            j = j + 1
        Loop
        
        If i = 1 Then
            WaitDoEvents tempo / smfHeader.Division * ActiveSheet.Cells(i, 1).Value
        Else
            WaitDoEvents tempo / smfHeader.Division * (ActiveSheet.Cells(i, 1).Value - ActiveSheet.Cells(i - 1, 1).Value)
        End If
        midiOutSendMsg hmo, tempmsg
        
        If ActiveSheet.Cells(i, 2).Value = &HFF And ActiveSheet.Cells(i, 3).Value = &H51 Then
            tempo = (ActiveSheet.Cells(i, 5).Value * 65536) + (ActiveSheet.Cells(i, 6).Value * 256) + ActiveSheet.Cells(i, 7).Value
        End If
        
    Next
    
    WaitDoEvents 2000000
    Debug.Print "midiOutClose" & vbTab & ": " & midiOutClose(hmo)
    
    Application.StatusBar = False
    
    DeleteAllSheets
    
End Sub

'''<summary>
'''バイナリファイルを配列に読み込む。
'''</summary>
'''
'''<param name="arr">結果を格納する配列</param>
'''<param name="path">ファイルパス</param>
'''
'''<returns></returns>
Sub BinaryReader(ByRef arr() As Byte, ByVal path As String)

    Dim objStream As ADODB.Stream
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
    
        .Open
        .Type = adTypeBinary
        .LoadFromFile path
        
        arr = .Read(adReadAll)
        
        .Close
        
    End With
    
End Sub

'''<summary>
'''offsetからlength個、Byte配列を読み込み整数値を返す。
'''</summary>
'''
'''<param name="arr">配列</param>
'''<param name="offset">現在位置</param>
'''<param name="length">読み込む数</param>
'''
'''<returns>整数値</returns>
Function GetValueByLong(ByRef arr() As Byte, ByRef offset As Long, ByVal length As Long) As Long
    
    Dim i As Long
    
    Dim power As Long
    power = (length - 1) * 2
    
    For i = offset To offset + length - 1
        GetValueByLong = GetValueByLong + (arr(i) * (16 ^ power))
        power = power - 2
    Next
    
    offset = offset + length

End Function

'''<summary>
'''offsetからlength個、Byte配列を読み込みASCIIを返す。
'''</summary>
'''
'''<param name="arr">配列</param>
'''<param name="offset">現在位置</param>
'''<param name="length">読み込む数</param>
'''
'''<returns>文字列</returns>
Function GetValueByString(ByRef arr() As Byte, ByRef offset As Long, ByVal length As Long) As String
    
    Dim i As Long
    
    For i = offset To offset + length - 1
        GetValueByString = GetValueByString & Chr(arr(i))
    Next
    
    offset = offset + length
    
End Function

'''<summary>
'''arr()のoffset位置からデルタタイムを読み込む
'''</summary>
'''
'''<param name="arr">配列</param>
'''<param name="offset">現在位置</param>
'''
'''<returns>デルタタイム</returns>
Function GetDeltaTime(ByRef arr() As Byte, ByRef offset As Long) As Long
    
    Dim i As Long
    i = 0
    
    Do While True
        If arr(offset + i) > &H7F Then
            i = i + 1
        Else
            i = i + 1
            Exit Do
        End If
    Loop
    
    offset = offset + i
    
    Dim j As Long
    For j = 0 To i - 1
        If j = 0 Then
            GetDeltaTime = arr(offset - (j + 1))
        Else
            GetDeltaTime = GetDeltaTime + (arr(offset - (j + 1)) - &H80) * (2 ^ (7 * j))
        End If
    Next
    
End Function

'Sub testGetDeltaTime()
'    Dim arr(1) As Byte
'    arr(0) = &H81
'    arr(1) = &H2F
'    Debug.Print GetDeltaTime(arr, 0)
'End Sub

Function DeleteAllSheets()
    Application.DisplayAlerts = False
    Do While Worksheets.count > 1
        Worksheets(Worksheets.count).Delete
    Loop
    Application.DisplayAlerts = True
    Cells.Clear
    Cells(1, 1).Select
End Function

'Debug Tools
'Sub SelectionHex()
'    Dim cell As Range
'    For Each cell In Selection
'        If cell.Value <> "" Then
'            cell.Value = "'" & Right("0" & Hex(cell.Value), 2)
'        End If
'    Next
'End Sub
'
'Sub SelectionDex()
'    Dim cell As Range
'    For Each cell In Selection
'        If cell.Value <> "" Then
'            cell.Value = CLng("&H" & cell.Value)
'        End If
'    Next
'End Sub

Function GetMicroSecond(ByVal freq As LongLong) As Double
    
    Dim procTime As LongLong
    
    '念のため初期化
    GetMicroSecond = 0
    
    'カウント数を「1マイクロ秒間に増えるカウント数」で割ることでマイクロ秒を取得できる
    QueryPerformanceCounter procTime
    GetMicroSecond = procTime / freq
    
End Function

Function HourMinSec(ByVal sec As Long) As String
    Dim s As Long
    Dim m As Long
    Dim h As Long
    
    Dim result As String
    
    s = sec
    m = Int(s / 60)
    h = Int(m / 60)
    
    m = m - (h * 60)
    s = s - (h * (60 ^ 2)) - (m * 60)
    
    If h < 10 Then
        result = Right("0" & CStr(h), 2) & ":" & Right("0" & CStr(m), 2) & ":" & Right("0" & CStr(s), 2)
    Else
        result = CStr(h) & ":" & Right("0" & CStr(m), 2) & ":" & Right("0" & CStr(s), 2)
    End If

    HourMinSec = result
    
End Function

Function ProgressBar(ByVal percent As Double) As String
    
    Dim temp As String
    Dim j As Long
    
    percent = Int(percent * 100)
    temp = "["
    For j = 1 To Int(percent / 4)
        temp = temp & "■"
    Next
    'temp = temp & ">"
    For j = 1 To 25 - Int(percent / 4)
        temp = temp & "□"
    Next
    temp = temp & "]"
    
    ProgressBar = temp & " " & Right("  " & percent, 3) & "%"

End Function

Function WaitDoEvents(microsec As LongLong)
    
    Dim start As Double
    start = GetMicroSecond(freq)
    
    Do While (GetMicroSecond(freq) - start) < (microsec / (1000 ^ 2))
        DoEvents
    Loop
    
End Function
