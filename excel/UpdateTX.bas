Attribute VB_Name = "modTX"
Option Explicit
' 台指期全日盤日線自動更新（按鈕「更新資料」）
' 從期交所下載 TX 近月：開=夜盤開、高低=夜盤+日盤、收=結算價、量=夜盤+日盤
' 從「收盤價」最後一筆的下一天補到今天；公式欄 G:Q 不夠時自動往下複製

Private Const TAIFEX_URL As String = "https://www.taifex.com.tw/cht/3/futDataDown"

Public Sub UpdateTX()
    Dim msg As String
    msg = RunUpdate()
    MsgBox msg, IIf(Left(msg, 4) = "更新失敗", vbExclamation, vbInformation)
End Sub

' 實際更新；回傳結果訊息（不跳視窗，方便測試）
Public Function RunUpdate() As String
    Dim ws As Worksheet, lastRow As Long, lastDate As Date
    Dim d As Date, e As Date, csv As String, added As Long
    Dim dates As Object, dayRows As Object, nightRows As Object
    Dim k As Variant, f As Variant, n As Variant, m As String
    Dim o As Double, h As Double, l As Double, c As Double, v As Double
    Dim r As Long, lines() As String, i As Long, sess As String
    Dim sDay As String, sNight As String

    On Error GoTo Fail
    sDay = ChrW(&H4E00) & ChrW(&H822C)    ' 一般
    sNight = ChrW(&H76E4) & ChrW(&H5F8C)  ' 盤後

    Set ws = ThisWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row     ' E 欄（收盤價）最後一筆
    lastDate = ParseDate(ws.Cells(lastRow, 1).Value)
    If lastDate >= Date Then
        RunUpdate = "已是最新（最後一筆 " & Format(lastDate, "yyyy/mm/dd") & "）"
        Exit Function
    End If

    Application.StatusBar = "正在從期交所下載資料…"
    Application.Cursor = xlWait
    Set dates = CreateObject("Scripting.Dictionary")
    Set dayRows = CreateObject("Scripting.Dictionary")
    Set nightRows = CreateObject("Scripting.Dictionary")

    d = lastDate + 1
    Do While d <= Date
        e = DateSerial(Year(d), Month(d) + 1, 0)
        If e > Date Then e = Date
        csv = Download(d, e)
        lines = Split(Replace(csv, vbCr, ""), vbLf)
        For i = 1 To UBound(lines)
            f = Split(lines(i), ",")
            If UBound(f) >= 17 Then
                m = Trim(f(2))
                sess = Trim(f(17))
                If Len(m) = 6 And IsNumeric(m) And IsNumeric(f(3)) Then
                    k = Trim(f(0))
                    If sess = sDay Then
                        dayRows(k & "|" & m) = f
                        If Not dates.Exists(k) Then
                            dates(k) = m
                        ElseIf m < dates(k) Then
                            dates(k) = m
                        End If
                    ElseIf sess = sNight Then
                        nightRows(k & "|" & m) = f
                    End If
                End If
            End If
        Next i
        d = e + 1
    Loop

    r = lastRow
    For Each k In dates.Keys
        f = dayRows(k & "|" & dates(k))
        If IsNumeric(f(10)) Then                       ' 結算價出來了才算收盤
            c = CDbl(f(10))
            o = CDbl(f(3)): h = CDbl(f(4)): l = CDbl(f(5)): v = Val(f(9))
            If nightRows.Exists(k & "|" & dates(k)) Then
                n = nightRows(k & "|" & dates(k))
                o = CDbl(n(3))
                If CDbl(n(4)) > h Then h = CDbl(n(4))
                If CDbl(n(5)) < l Then l = CDbl(n(5))
                v = v + Val(n(9))
            End If
            r = r + 1
            ws.Cells(r, 1).NumberFormat = "@"
            ws.Cells(r, 1).Value = Replace(Mid(k, 1, 4) & "-" & Mid(k, 6, 2) & Mid(k, 9, 2), "/", "")
            ws.Cells(r, 2).Value = o
            ws.Cells(r, 3).Value = h
            ws.Cells(r, 4).Value = l
            ws.Cells(r, 5).Value = c
            ws.Cells(r, 6).Value = v
            If Not ws.Cells(r, 7).HasFormula Then
                ws.Range(ws.Cells(r - 1, 7), ws.Cells(r - 1, 17)).Copy ws.Cells(r, 7)
            End If
            added = added + 1
        End If
    Next k

    PlaceButton
    Application.StatusBar = False
    Application.Cursor = xlDefault
    If added = 0 Then
        RunUpdate = "期交所還沒有新的完整日線（日盤收盤、結算價公布後才會有）。" & vbCrLf & _
               "目前最後一筆：" & Format(lastDate, "yyyy/mm/dd")
    Else
        RunUpdate = "已補入 " & added & " 天，最後一筆：" & ws.Cells(r, 1).Value
    End If
    Exit Function

Fail:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    RunUpdate = "更新失敗：" & Err.Description
End Function

Private Function ParseDate(x As Variant) As Date
    Dim s As String
    If IsDate(x) And VarType(x) = vbDate Then
        ParseDate = x
    Else
        s = Trim(CStr(x))                                ' 例：2026-1106
        ParseDate = DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 6, 2)), CInt(Mid(s, 8, 2)))
    End If
End Function

Private Function Download(d1 As Date, d2 As Date) As String
    Dim http As Object, stm As Object, body As String
    body = "down_type=1&commodity_id=TX&commodity_id2=" & _
           "&queryStartDate=" & Format(d1, "yyyy\/mm\/dd") & _
           "&queryEndDate=" & Format(d2, "yyyy\/mm\/dd")
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", TAIFEX_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send body
    If http.Status <> 200 Then Err.Raise vbObjectError + 1, , "期交所回應 " & http.Status
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write http.responseBody
    stm.Position = 0: stm.Type = 2: stm.Charset = "big5"
    Download = stm.ReadText
    stm.Close
End Function

' 把「更新資料」按鈕放到最後一筆收盤價往上 6 列的 R 欄
Public Sub PlaceButton()
    Dim ws As Worksheet, lr As Long, cel As Range
    Set ws = ThisWorkbook.Worksheets(1)
    lr = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row - 6
    If lr < 2 Then lr = 2
    Set cel = ws.Cells(lr, 18)
    With ws.Buttons("btnUpdateTX")
        .Top = cel.Top
        .Left = cel.Left
    End With
End Sub
