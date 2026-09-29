Attribute VB_Name = "WeatherMacro"
Option Explicit

' 기상청 초단기실황(getUltraSrtNcst) 데이터를 "날씨" 시트에 입력
' 시트 "날씨": B1=API 키(공공데이터포털), B2=위도, B3=경도

Private Const PI As Double = 3.14159265358979

Sub 기상청날씨가져오기()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("날씨")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "'날씨' 시트가 없습니다.", vbExclamation
        Exit Sub
    End If

    Dim apiKey As String
    apiKey = Trim(CStr(ws.Range("B1").Value))
    If apiKey = "" Then
        MsgBox "B1 셀에 공공데이터포털 API 인증키를 입력하세요.", vbExclamation
        Exit Sub
    End If

    Dim nx As Long, ny As Long
    ToGrid CDbl(ws.Range("B2").Value), CDbl(ws.Range("B3").Value), nx, ny

    Dim t As Date
    t = DateAdd("n", -40, Now)  ' 실황은 매시 정각 생성, 약 10분 뒤 제공

    Dim url As String
    url = "http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtNcst" & _
          "?serviceKey=" & apiKey & "&pageNo=1&numOfRows=10&dataType=XML" & _
          "&base_date=" & Format(t, "yyyymmdd") & "&base_time=" & Format(t, "hh") & "00" & _
          "&nx=" & nx & "&ny=" & ny

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo Fail
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then GoTo Fail

    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.LoadXML http.responseText
    Dim code As Object
    Set code = doc.SelectSingleNode("//resultCode")
    If code Is Nothing Then GoTo Fail
    If code.Text <> "00" Then
        MsgBox "API 오류: " & doc.SelectSingleNode("//resultMsg").Text, vbExclamation
        Exit Sub
    End If

    ws.Range("A5:C30").ClearContents
    ws.Range("A5:C5").Value = Array("코드", "항목", "값")
    Dim items As Object, i As Long, r As Long, cat As String
    Set items = doc.SelectNodes("//item")
    r = 6
    For i = 0 To items.Length - 1
        cat = items(i).SelectSingleNode("category").Text
        ws.Cells(r, 1).Value = cat
        ws.Cells(r, 2).Value = CategoryName(cat)
        ws.Cells(r, 3).Value = items(i).SelectSingleNode("obsrValue").Text
        r = r + 1
    Next i
    ws.Range("A4").Value = "기준: " & Format(t, "yyyy-mm-dd hh") & ":00  격자 (" & nx & "," & ny & ")"
    MsgBox "날씨 데이터를 입력했습니다.", vbInformation
    Exit Sub
Fail:
    MsgBox "날씨 데이터를 가져오지 못했습니다. 인터넷/API 키를 확인하세요.", vbExclamation
End Sub

Private Function CategoryName(c As String) As String
    Select Case c
        Case "T1H": CategoryName = "기온(℃)"
        Case "RN1": CategoryName = "1시간 강수량(mm)"
        Case "REH": CategoryName = "습도(%)"
        Case "PTY": CategoryName = "강수형태(0없음 1비 2비/눈 3눈 5빗방울 6빗방울눈날림 7눈날림)"
        Case "WSD": CategoryName = "풍속(m/s)"
        Case "VEC": CategoryName = "풍향(deg)"
        Case "UUU": CategoryName = "동서바람성분(m/s)"
        Case "VVV": CategoryName = "남북바람성분(m/s)"
        Case Else: CategoryName = c
    End Select
End Function

' 기상청 Lambert Conformal Conic 격자 변환
Private Sub ToGrid(lat As Double, lon As Double, ByRef x As Long, ByRef y As Long)
    Dim d As Double, re As Double, s1 As Double, s2 As Double, ol As Double, oa As Double
    d = PI / 180#
    re = 6371.00877 / 5#
    s1 = 30# * d: s2 = 60# * d: ol = 126# * d: oa = 38# * d

    Dim sn As Double, sf As Double, ro As Double, ra As Double, th As Double
    sn = Log(Cos(s1) / Cos(s2)) / Log(Tan(PI / 4 + s2 / 2) / Tan(PI / 4 + s1 / 2))
    sf = (Tan(PI / 4 + s1 / 2) ^ sn) * Cos(s1) / sn
    ro = re * sf / (Tan(PI / 4 + oa / 2) ^ sn)
    ra = re * sf / (Tan(PI / 4 + lat * d / 2) ^ sn)
    th = lon * d - ol
    If th > PI Then th = th - 2 * PI
    If th < -PI Then th = th + 2 * PI
    th = th * sn

    x = Int(ra * Sin(th) + 43 + 0.5)
    y = Int(ro - ra * Cos(th) + 136 + 0.5)
End Sub
