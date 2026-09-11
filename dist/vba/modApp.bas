Attribute VB_Name = "modApp"
'==========================================================================
' modApp  シートの読み書きとマクロ本体
'   - 児童生徒の個人情報は扱いません。管理は出席番号のみです。
'   - 「所見を作る」ボタン（またはマクロ GenerateAll）で全員分を生成します。
'==========================================================================
Option Explicit

Private Const DATA_START As Long = 2       ' 1行目は見出し

' 所見シートの列
Private Const C_NO As Long = 1
Private Const C_BUMP As Long = 2
Private Const C_REPORT As Long = 3
Private Const C_REPORT_LEN As Long = 4
Private Const C_YOUROKU As Long = 5
Private Const C_YOUROKU_LEN As Long = 6
Private Const C_MORAL_R As Long = 7
Private Const C_MORAL_R_LEN As Long = 8
Private Const C_MORAL_Y As Long = 9
Private Const C_MORAL_Y_LEN As Long = 10
Private Const C_SOUGOU_R As Long = 11
Private Const C_SOUGOU_R_LEN As Long = 12
Private Const C_SOUGOU_Y As Long = 13
Private Const C_SOUGOU_Y_LEN As Long = 14
Private Const C_WARN As Long = 15
Private Const C_LOCK As Long = 16

'--- 小物 -----------------------------------------------------------------

Private Function YesStr() As String
    YesStr = ChrW$(&H306F) & ChrW$(&H3044)                  ' はい
End Function

Private Function ParseBool(ByVal v As Variant) As Boolean
    Dim s As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then ParseBool = CBool(v): Exit Function
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If s = YesStr() Or s = ChrW$(&H25CB) Or s = ChrW$(&H25CE) Or s = "1" Then ParseBool = True: Exit Function
    If LCase$(s) = "true" Or LCase$(s) = "yes" Then ParseBool = True
End Function

' シート全体を配列で読む（空なら Empty）
Private Function ReadSheet(ByVal name As String) As Variant
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim fRow As Range, fCol As Range
    Set ws = ThisWorkbook.Worksheets(name)
    Set fRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set fCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If fRow Is Nothing Or fCol Is Nothing Then Exit Function
    lastRow = fRow.Row
    lastCol = fCol.Column
    If lastRow < DATA_START Then Exit Function
    ReadSheet = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
End Function

Private Function Cell(ByVal arr As Variant, ByVal r As Long, ByVal c As Long) As String
    If IsEmpty(arr) Then Exit Function
    If r < LBound(arr, 1) Or r > UBound(arr, 1) Then Exit Function
    If c < LBound(arr, 2) Or c > UBound(arr, 2) Then Exit Function
    If IsError(arr(r, c)) Then Exit Function
    If IsNull(arr(r, c)) Then Exit Function
    Cell = Trim$(CStr(arr(r, c)))
End Function

'--- 設定の読み込み -------------------------------------------------------
'   設定シート:  A 項目 / B キー / C 値
'                E 順番 / F キー / G 使う          （所見に入れる項目）
'                I 順番 / J キー / K 表示          （行動の記録の並び順）
Public Function ReadSettings() As Object
    Dim arr As Variant, s As Object, r As Long, k As String
    Dim order() As String, bOrder() As String, enabled As Object
    Dim nOrder As Long, nB As Long

    Set s = NewDict()
    Set enabled = NewDict()
    arr = ReadSheet(ShSettings())

    ' 既定値
    s.Add "term", ChrW$(&H0032) & ChrW$(&H5B66) & ChrW$(&H671F)   ' 2学期
    s.Add "count", 35
    s.Add "limitReport", 120
    s.Add "limitYouroku", 60
    s.Add "moralLimitReport", 120
    s.Add "moralLimitYouroku", 60
    s.Add "sougouLimitReport", 120
    s.Add "sougouLimitYouroku", 60
    s.Add "behaviorMax", 2
    s.Add "behaviorPick", "order"
    s.Add "useClosing", False
    s.Add "moralIntro", True
    s.Add "moralClosing", True
    s.Add "sougouClosing", True

    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            k = Cell(arr, r, 2)
            If Len(k) > 0 Then
                Select Case k
                    Case "term"
                        If Len(Cell(arr, r, 3)) > 0 Then s("term") = Cell(arr, r, 3)
                    Case "count", "limitReport", "limitYouroku", "moralLimitReport", _
                         "moralLimitYouroku", "sougouLimitReport", "sougouLimitYouroku", "behaviorMax"
                        If Len(Cell(arr, r, 3)) > 0 Then s(k) = ToNumber(Cell(arr, r, 3))
                    Case "behaviorPick"
                        If Len(Cell(arr, r, 3)) > 0 Then s(k) = LCase$(Cell(arr, r, 3))
                    Case "useClosing", "moralIntro", "moralClosing", "sougouClosing"
                        s(k) = ParseBool(Cell(arr, r, 3))
                End Select
            End If
        Next r

        ' 所見に入れる項目と順番
        nOrder = 0
        ReDim order(0 To 10)
        For r = DATA_START To UBound(arr, 1)
            k = Cell(arr, r, 6)
            If Len(k) > 0 Then
                order(nOrder) = k
                nOrder = nOrder + 1
                enabled(k) = ParseBool(Cell(arr, r, 7))
            End If
        Next r
        If nOrder > 0 Then
            ReDim Preserve order(0 To nOrder - 1)
            s.Add "order", order
        End If

        ' 行動の記録の並び順
        nB = 0
        ReDim bOrder(0 To 20)
        For r = DATA_START To UBound(arr, 1)
            k = Cell(arr, r, 10)
            If Len(k) > 0 Then
                bOrder(nB) = k
                nB = nB + 1
            End If
        Next r
        If nB > 0 Then
            ReDim Preserve bOrder(0 To nB - 1)
            s.Add "behaviorOrder", bOrder
        End If
    End If

    If Not s.Exists("order") Then
        s.Add "order", Array("observation", "event", "committee", "kakari", "behavior")
        enabled("observation") = True: enabled("event") = True: enabled("committee") = True
        enabled("kakari") = True: enabled("behavior") = True
    End If
    If Not s.Exists("behaviorOrder") Then s.Add "behaviorOrder", BehaviorKeys()
    s.Add "enabled", enabled
    If CLng(s("count")) < 1 Then s("count") = 35

    Set ReadSettings = s
End Function

'--- 見取りの文型（表示名 -> キー）---------------------------------------
Private Function LoadObsLabels() As Object
    Dim ws As Worksheet, lastRow As Long, r As Long, d As Object, lab As String, k As String
    Set d = NewDict()
    Set ws = ThisWorkbook.Worksheets(ShTpl())
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "obs" Then
            lab = Trim$(CStr(ws.Cells(r, 1).Value))
            k = Trim$(CStr(ws.Cells(r, 3).Value))
            If Len(lab) > 0 And Len(k) > 0 Then
                If Not d.Exists(lab) Then d.Add lab, k
            End If
        End If
    Next r
    Set LoadObsLabels = d
End Function

'--- 児童データの読み込み -------------------------------------------------

Private Function EnsureStudent(ByVal students As Object, ByVal no As Long) As Object
    Dim st As Object, key As String
    key = CStr(no)
    If students.Exists(key) Then
        Set EnsureStudent = students(key)
        Exit Function
    End If
    Set st = NewDict()
    st.Add "no", no
    st.Add "committee", ""
    st.Add "committeeRole", ""
    st.Add "committeeWork", ""
    st.Add "kakari", ""
    st.Add "kakariWork", ""
    st.Add "events", New Collection
    st.Add "behavior", NewDict()
    st.Add "observations", New Collection
    st.Add "moral", New Collection
    st.Add "sougou", New Collection
    students.Add key, st
    Set EnsureStudent = st
End Function

Private Function ReadStudents(ByVal settings As Object) As Object
    Dim students As Object, arr As Variant, r As Long, no As Long, st As Object
    Dim i As Long, j As Long, keys As Variant, ev As Object, ob As Object
    Dim obsLabels As Object, lab As String, mk As String, tx As String

    Set students = NewDict()
    Set obsLabels = LoadObsLabels()

    ' 委員会・係
    arr = ReadSheet(ShRoster())
    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 Then
                Set st = EnsureStudent(students, no)
                st("committee") = Cell(arr, r, 2)
                st("committeeRole") = Cell(arr, r, 3)
                st("committeeWork") = Cell(arr, r, 4)
                st("kakari") = Cell(arr, r, 5)
                st("kakariWork") = Cell(arr, r, 6)
            End If
        Next r
    End If

    ' 行事（1人3つまで: B..D / E..G / H..J）
    arr = ReadSheet(ShEvent())
    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 Then
                Set st = EnsureStudent(students, no)
                For i = 0 To 2
                    If Len(Cell(arr, r, 2 + i * 3)) > 0 Then
                        Set ev = NewDict()
                        ev.Add "name", Cell(arr, r, 2 + i * 3)
                        ev.Add "role", Cell(arr, r, 3 + i * 3)
                        ev.Add "effort", Cell(arr, r, 4 + i * 3)
                        st("events").Add ev
                    End If
                Next i
            End If
        Next r
    End If

    ' 行動の記録（B..K が10項目・列の順番は固定）
    arr = ReadSheet(ShBehavior())
    If Not IsEmpty(arr) Then
        keys = BehaviorKeys()
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 Then
                Set st = EnsureStudent(students, no)
                For j = 0 To UBound(keys)
                    mk = NormalizeMark(Cell(arr, r, 2 + j))
                    If Len(mk) > 0 Then st("behavior")(CStr(keys(j))) = mk
                Next j
            End If
        Next r
    End If

    ' 見取り（1人3つまで: B/C, D/E, F/G）
    arr = ReadSheet(ShObs())
    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 Then
                Set st = EnsureStudent(students, no)
                For i = 0 To 2
                    tx = Cell(arr, r, 2 + i * 2)
                    If Len(tx) > 0 Then
                        Set ob = NewDict()
                        ob.Add "text", tx
                        lab = Cell(arr, r, 3 + i * 2)
                        If obsLabels.Exists(lab) Then
                            ob.Add "pattern", obsLabels(lab)
                        Else
                            ob.Add "pattern", "sugata"
                        End If
                        st("observations").Add ob
                    End If
                Next i
            End If
        Next r
    End If

    ' 道徳（1人何行でも）
    arr = ReadSheet(ShMoral())
    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 And Len(Cell(arr, r, 2)) > 0 Then
                Set st = EnsureStudent(students, no)
                Set ob = NewDict()
                ob.Add "material", Cell(arr, r, 2)
                ob.Add "value", Cell(arr, r, 3)
                ob.Add "note", Cell(arr, r, 4)
                st("moral").Add ob
            End If
        Next r
    End If

    ' 総合（1人何行でも）
    arr = ReadSheet(ShSougou())
    If Not IsEmpty(arr) Then
        For r = DATA_START To UBound(arr, 1)
            no = ToNumber(Cell(arr, r, 1))
            If no > 0 And Len(Cell(arr, r, 2)) > 0 Then
                Set st = EnsureStudent(students, no)
                Set ob = NewDict()
                ob.Add "unit", Cell(arr, r, 2)
                ob.Add "theme", Cell(arr, r, 3)
                ob.Add "activity", Cell(arr, r, 4)
                ob.Add "result", Cell(arr, r, 5)
                st("sougou").Add ob
            End If
        Next r
    End If

    Set ReadStudents = students
End Function

'--- 警告文の組み立て -----------------------------------------------------

Private Function WarnText(ByVal res As Object, ByVal label As String) As String
    Dim over As String
    If res("over") = True Then
        over = label & ChrW$(&H304C) & CStr(res("length")) & ChrW$(&H5B57) & ChrW$(&HFF08) & _
               ChrW$(&H4E0A) & ChrW$(&H9650) & CStr(res("limit")) & ChrW$(&HFF09)
    End If
    If res("unused").Count > 0 Then
        If Len(over) > 0 Then over = over & " / "
        over = over & label & ChrW$(&H306E) & ChrW$(&H672A) & ChrW$(&H4F7F) & ChrW$(&H7528) & _
               CStr(res("unused").Count) & ChrW$(&H6587)
    End If
    WarnText = over
End Function

'==========================================================================
' マクロ本体：全員分の所見を作る
'==========================================================================
Public Sub GenerateAll()
    Dim tpls As Object, settings As Object, students As Object
    Dim ws As Worksheet, no As Long, cnt As Long, r As Long
    Dim st As Object, bump As Long, warn As String, pi As String
    Dim rRep As Object, rYou As Object, rMoR As Object, rMoY As Object, rSoR As Object, rSoY As Object
    Dim locked As Boolean, made As Long

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set tpls = LoadTemplates()
    Set settings = ReadSettings()
    Set students = ReadStudents(settings)

    Set ws = ThisWorkbook.Worksheets(ShOut())
    cnt = CLng(settings("count"))

    ' 既存の出力を消す（別案・手直しの指定は残す）
    If ws.Cells(ws.Rows.Count, C_NO).End(xlUp).Row >= DATA_START Then
        ws.Range(ws.Cells(DATA_START, C_REPORT), ws.Cells(ws.Rows.Count, C_WARN)).ClearContents
    End If

    For no = 1 To cnt
        r = DATA_START + no - 1
        ws.Cells(r, C_NO).Value = no
        locked = ParseBool(ws.Cells(r, C_LOCK).Value)
        bump = ToNumber(ws.Cells(r, C_BUMP).Value)

        Set st = EnsureStudent(students, no)
        Set rRep = BuildMain(st, settings, tpls, "report", bump)
        Set rYou = BuildMain(st, settings, tpls, "youroku", bump)
        Set rMoR = BuildMoral(st, settings, tpls, "report", bump)
        Set rMoY = BuildMoral(st, settings, tpls, "youroku", bump)
        Set rSoR = BuildSougou(st, settings, tpls, "report", bump)
        Set rSoY = BuildSougou(st, settings, tpls, "youroku", bump)

        If Not locked Then
            ws.Cells(r, C_REPORT).Value = rRep("text")
            ws.Cells(r, C_YOUROKU).Value = rYou("text")
            ws.Cells(r, C_MORAL_R).Value = rMoR("text")
            ws.Cells(r, C_MORAL_Y).Value = rMoY("text")
            ws.Cells(r, C_SOUGOU_R).Value = rSoR("text")
            ws.Cells(r, C_SOUGOU_Y).Value = rSoY("text")
        End If

        ws.Cells(r, C_REPORT_LEN).Value = CountChars(CStr(ws.Cells(r, C_REPORT).Value))
        ws.Cells(r, C_YOUROKU_LEN).Value = CountChars(CStr(ws.Cells(r, C_YOUROKU).Value))
        ws.Cells(r, C_MORAL_R_LEN).Value = CountChars(CStr(ws.Cells(r, C_MORAL_R).Value))
        ws.Cells(r, C_MORAL_Y_LEN).Value = CountChars(CStr(ws.Cells(r, C_MORAL_Y).Value))
        ws.Cells(r, C_SOUGOU_R_LEN).Value = CountChars(CStr(ws.Cells(r, C_SOUGOU_R).Value))
        ws.Cells(r, C_SOUGOU_Y_LEN).Value = CountChars(CStr(ws.Cells(r, C_SOUGOU_Y).Value))

        warn = ""
        If Not locked Then
            warn = WarnText(rRep, ChrW$(&H901A) & ChrW$(&H4FE1) & ChrW$(&H8868))
            If Len(WarnText(rYou, ChrW$(&H8981) & ChrW$(&H9332))) > 0 Then
                If Len(warn) > 0 Then warn = warn & " / "
                warn = warn & WarnText(rYou, ChrW$(&H8981) & ChrW$(&H9332))
            End If
        End If

        ' 氏名らしい表現のチェック（個人情報は入力しません）
        pi = FindPersonalInfo(CStr(ws.Cells(r, C_REPORT).Value) & CStr(ws.Cells(r, C_YOUROKU).Value) & _
                              CStr(ws.Cells(r, C_MORAL_R).Value) & CStr(ws.Cells(r, C_SOUGOU_R).Value))
        If Len(pi) > 0 Then
            If Len(warn) > 0 Then warn = warn & " / "
            warn = warn & ChrW$(&H6C0F) & ChrW$(&H540D) & ChrW$(&H304B) & ChrW$(&H3082) & "?: " & pi
        End If
        ws.Cells(r, C_WARN).Value = warn

        If Len(CStr(ws.Cells(r, C_REPORT).Value)) > 0 Then made = made + 1
    Next no

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox ChrW$(&H6240) & ChrW$(&H898B) & ChrW$(&H3092) & ChrW$(&H4F5C) & ChrW$(&H308A) & _
           ChrW$(&H307E) & ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H3002) & vbCrLf & _
           CStr(made) & "/" & CStr(cnt) & ChrW$(&H4EBA) & vbCrLf & vbCrLf & _
           ChrW$(&H751F) & ChrW$(&H6210) & ChrW$(&H3055) & ChrW$(&H308C) & ChrW$(&H305F) & _
           ChrW$(&H6587) & ChrW$(&H306F) & ChrW$(&H6587) & ChrW$(&H4F8B) & ChrW$(&H306E) & _
           ChrW$(&H7D44) & ChrW$(&H307F) & ChrW$(&H5408) & ChrW$(&H308F) & ChrW$(&H305B) & _
           ChrW$(&H3067) & ChrW$(&H3059) & ChrW$(&H3002) & vbCrLf & _
           ChrW$(&H5FC5) & ChrW$(&H305A) & ChrW$(&H305D) & ChrW$(&H306E) & ChrW$(&H5B50) & _
           ChrW$(&H306E) & ChrW$(&H4E8B) & ChrW$(&H5B9F) & ChrW$(&H3068) & ChrW$(&H7167) & _
           ChrW$(&H3089) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H78BA) & ChrW$(&H8A8D) & _
           ChrW$(&H30FB) & ChrW$(&H4FEE) & ChrW$(&H6B63) & ChrW$(&H3057) & ChrW$(&H3066) & _
           ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002), _
           vbInformation
    Exit Sub

Failed:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox ChrW$(&H4F5C) & ChrW$(&H6210) & ChrW$(&H4E2D) & ChrW$(&H306B) & ChrW$(&H554F) & _
           ChrW$(&H984C) & ChrW$(&H304C) & ChrW$(&H8D77) & ChrW$(&H304D) & ChrW$(&H307E) & _
           ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H3002) & vbCrLf & Err.Description, vbExclamation
End Sub

'==========================================================================
' 入力データの全消去（設定・文例は残します）
'==========================================================================
Public Sub ClearAllData()
    Dim names As Variant, i As Long, ws As Worksheet, ans As VbMsgBoxResult, lastRow As Long

    ans = MsgBox(ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H305F) & ChrW$(&H30C7) & _
                 ChrW$(&H30FC) & ChrW$(&H30BF) & ChrW$(&H3092) & ChrW$(&H3059) & ChrW$(&H3079) & _
                 ChrW$(&H3066) & ChrW$(&H6D88) & ChrW$(&H3057) & ChrW$(&H307E) & ChrW$(&H3059) & _
                 ChrW$(&H3002) & vbCrLf & ChrW$(&H3088) & ChrW$(&H308D) & ChrW$(&H3057) & _
                 ChrW$(&H3044) & ChrW$(&H3067) & ChrW$(&H3059) & ChrW$(&H304B) & ChrW$(&HFF1F), _
                 vbYesNo + vbExclamation)
    If ans <> vbYes Then Exit Sub

    names = Array(ShRoster(), ShEvent(), ShBehavior(), ShObs(), ShMoral(), ShSougou(), ShOut())
    For i = LBound(names) To UBound(names)
        Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= DATA_START Then
            ws.Range(ws.Rows(DATA_START), ws.Rows(lastRow)).ClearContents
        End If
    Next i

    MsgBox ChrW$(&H6D88) & ChrW$(&H53BB) & ChrW$(&H3057) & ChrW$(&H307E) & ChrW$(&H3057) & _
           ChrW$(&H305F) & ChrW$(&H3002), vbInformation
End Sub
