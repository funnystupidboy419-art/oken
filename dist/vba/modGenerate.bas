Attribute VB_Name = "modGenerate"
'==========================================================================
' modGenerate  所見の組み立て
'   - 「どの項目についての文か」を保持したまま文を並べます。
'   - 並び順は「設定」シートに従い、上にあるものほど優先度が高いものとして
'     文字数超過時は下の項目から短縮・削除します。
'   - 直前の文と文末が重ならないよう、候補をずらして選びます。
'==========================================================================
Option Explicit

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

'--- 語句の組み立て補助 ---------------------------------------------------

' 末尾の助詞が重複しないように付ける
Private Function WithParticle(ByVal text As String, ByVal particle As String) As String
    Dim t As String, lastCh As String
    t = Trim$(text)
    If Len(t) = 0 Then Exit Function
    If Len(t) >= Len(particle) Then
        If Right$(t, Len(particle)) = particle Then WithParticle = t: Exit Function
    End If
    If Len(particle) = 1 Then
        lastCh = Right$(t, 1)
        Select Case lastCh
            Case ChrW$(&H3092), ChrW$(&H306B), ChrW$(&H3067), ChrW$(&H3078), ChrW$(&H3068), _
                 ChrW$(&H3084), ChrW$(&H306E), ChrW$(&H306F), ChrW$(&H304C), ChrW$(&H3082), _
                 ChrW$(&H3001), ChrW$(&H3002)
                WithParticle = t
                Exit Function
        End Select
    End If
    WithParticle = t & particle
End Function

Private Function SuffixStr(ByVal text As String, ByVal tail As String) As String
    Dim t As String
    t = Trim$(text)
    If Len(t) = 0 Then Exit Function
    SuffixStr = t & tail
End Function

' 文例の選択（出席番号と「別案」の回数で散らす）
Private Function PickOne(ByVal col As Collection, ByVal seed As Long) As String
    Dim n As Long, i As Long
    If col Is Nothing Then Exit Function
    n = col.Count
    If n = 0 Then Exit Function
    i = ((seed Mod n) + n) Mod n
    PickOne = col(i + 1)
End Function

'--- パーツ生成 -----------------------------------------------------------

Private Function NewPart(ByVal cat As String, ByVal label As String, ByVal detail As String, _
                         ByVal longCol As Collection, ByVal shortCol As Collection, _
                         ByVal vars As Object, ByVal seed As Long, _
                         Optional ByVal raw As Boolean = False) As Object
    Dim p As Object
    Set p = NewDict()
    p.Add "cat", cat
    p.Add "label", label
    p.Add "detail", detail
    p.Add "long", longCol
    p.Add "short", shortCol
    p.Add "vars", vars
    p.Add "seed", seed
    p.Add "raw", raw
    Set NewPart = p
End Function

Private Sub AddCommittee(ByVal parts As Collection, ByVal st As Object, ByVal tpls As Object, ByVal seed As Long)
    Dim nm As String, vars As Object, detail As String
    nm = Trim$(CStr(st("committee")))
    If Len(nm) = 0 Then Exit Sub
    Set vars = NewDict()
    vars.Add ChrW$(&H59D4) & ChrW$(&H54E1) & ChrW$(&H4F1A), nm
    vars.Add ChrW$(&H5F79) & ChrW$(&H5272) & ChrW$(&H53E5), _
             SuffixStr(CStr(st("committeeRole")), ChrW$(&H3068) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H3001))
    If Len(Trim$(CStr(st("committeeWork")))) > 0 Then
        vars.Add ChrW$(&H6D3B) & ChrW$(&H52D5) & ChrW$(&H53E5), WithParticle(CStr(st("committeeWork")), ChrW$(&H306B))
    Else
        vars.Add ChrW$(&H6D3B) & ChrW$(&H52D5) & ChrW$(&H53E5), ChrW$(&H6D3B) & ChrW$(&H52D5) & ChrW$(&H306B)
    End If
    detail = nm
    If Len(Trim$(CStr(st("committeeRole")))) > 0 Then detail = detail & "/" & Trim$(CStr(st("committeeRole")))
    parts.Add NewPart("committee", CategoryLabel("committee"), detail, _
                      TplPair(tpls, "committee", "", "", False), _
                      TplPair(tpls, "committee", "", "", True), vars, seed)
End Sub

Private Sub AddKakari(ByVal parts As Collection, ByVal st As Object, ByVal tpls As Object, ByVal seed As Long)
    Dim nm As String, vars As Object
    nm = Trim$(CStr(st("kakari")))
    If Len(nm) = 0 Then Exit Sub
    Set vars = NewDict()
    vars.Add ChrW$(&H4FC2), nm
    If Len(Trim$(CStr(st("kakariWork")))) > 0 Then
        vars.Add ChrW$(&H4ED5) & ChrW$(&H4E8B) & ChrW$(&H53E5), WithParticle(CStr(st("kakariWork")), ChrW$(&H3092))
    Else
        vars.Add ChrW$(&H4ED5) & ChrW$(&H4E8B) & ChrW$(&H53E5), _
                 ChrW$(&H81EA) & ChrW$(&H5206) & ChrW$(&H306E) & ChrW$(&H4ED5) & ChrW$(&H4E8B) & ChrW$(&H3092)
    End If
    parts.Add NewPart("kakari", CategoryLabel("kakari"), nm, _
                      TplPair(tpls, "kakari", "", "", False), _
                      TplPair(tpls, "kakari", "", "", True), vars, seed)
End Sub

Private Sub AddEvents(ByVal parts As Collection, ByVal st As Object, ByVal tpls As Object, ByVal seed As Long)
    Dim evs As Collection, i As Long, ev As Object, nm As String, vars As Object, detail As String
    Set evs = st("events")
    If evs Is Nothing Then Exit Sub
    For i = 1 To evs.Count
        Set ev = evs(i)
        nm = Trim$(CStr(ev("name")))
        If Len(nm) > 0 Then
            Set vars = NewDict()
            vars.Add ChrW$(&H884C) & ChrW$(&H4E8B), nm
            vars.Add ChrW$(&H5F79) & ChrW$(&H5272) & ChrW$(&H53E5), _
                     SuffixStr(CStr(ev("role")), ChrW$(&H3068) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H3001))
            vars.Add ChrW$(&H9811) & ChrW$(&H5F35) & ChrW$(&H308A) & ChrW$(&H53E5), _
                     SuffixStr(CStr(ev("effort")), ChrW$(&H3001))
            detail = nm
            If Len(Trim$(CStr(ev("role")))) > 0 Then detail = detail & "/" & Trim$(CStr(ev("role")))
            parts.Add NewPart("event", CategoryLabel("event"), detail, _
                              TplPair(tpls, "event", "", "", False), _
                              TplPair(tpls, "event", "", "", True), vars, seed + i - 1)
        End If
    Next i
End Sub

Private Function BehaviorTplFor(ByVal tpls As Object, ByVal key As String, ByVal mk As String, _
                                ByVal wantShort As Boolean) As Collection
    Dim col As Collection
    Set col = TplPair(tpls, "behavior", key, mk, wantShort)
    If col.Count = 0 Then Set col = TplPair(tpls, "behavior", key, MarkB(), wantShort)
    Set BehaviorTplFor = col
End Function

Private Sub AddBehavior(ByVal parts As Collection, ByVal st As Object, ByVal settings As Object, _
                        ByVal tpls As Object, ByVal seed As Long)
    Dim marks As Object, order As Variant, i As Long, k As String, mk As String
    Dim picked As Collection, ordered As Collection, maxN As Long, n As Long, detail As String

    Set marks = st("behavior")
    order = settings("behaviorOrder")
    Set picked = New Collection

    For i = LBound(order) To UBound(order)
        k = CStr(order(i))
        If marks.Exists(k) Then
            mk = CStr(marks(k))
            If mk = MarkA() Or mk = MarkB() Then picked.Add k
        End If
    Next i

    ' 「◎を優先」のときは ◎ を先に、その中では学校の並び順を保つ
    If CStr(settings("behaviorPick")) = "mark" Then
        Set ordered = New Collection
        For i = 1 To picked.Count
            If CStr(marks(picked(i))) = MarkA() Then ordered.Add picked(i)
        Next i
        For i = 1 To picked.Count
            If CStr(marks(picked(i))) = MarkB() Then ordered.Add picked(i)
        Next i
        Set picked = ordered
    End If

    maxN = CLng(settings("behaviorMax"))
    If maxN < 1 Then maxN = 1
    n = picked.Count
    If n > maxN Then n = maxN

    For i = 1 To n
        k = picked(i)
        mk = CStr(marks(k))
        detail = BehaviorLabel(k) & ChrW$(&HFF08) & mk & ChrW$(&HFF09)
        parts.Add NewPart("behavior", CategoryLabel("behavior"), detail, _
                          BehaviorTplFor(tpls, k, mk, False), _
                          BehaviorTplFor(tpls, k, mk, True), NewDict(), seed + i - 1)
    Next i
End Sub

Private Sub AddObservations(ByVal parts As Collection, ByVal st As Object, ByVal tpls As Object, ByVal seed As Long)
    Dim obs As Collection, i As Long, ob As Object, tx As String
    Dim tpl As String, col As Collection, vars As Object, isRaw As Boolean, detail As String
    Set obs = st("observations")
    If obs Is Nothing Then Exit Sub
    For i = 1 To obs.Count
        Set ob = obs(i)
        tx = Trim$(CStr(ob("text")))
        If Len(tx) > 0 Then
            tpl = ObsPattern(tpls, CStr(ob("pattern")))
            isRaw = (CStr(ob("pattern")) = "as-is")
            Set col = New Collection
            col.Add tpl
            Set vars = NewDict()
            vars.Add ChrW$(&H898B) & ChrW$(&H53D6) & ChrW$(&H308A), tx
            detail = ChrW$(&H898B) & ChrW$(&H53D6) & ChrW$(&H308A) & CStr(i)
            parts.Add NewPart("observation", CategoryLabel("observation"), detail, _
                              col, col, vars, seed + i - 1, isRaw)
        End If
    Next i
End Sub

'--- 文字列化 -------------------------------------------------------------
'   同じ文末（「～取り組んでいました。」など）が続くと、まとめて読んだときに
'   コピー＆ペーストしたような印象になる。保護者が読む文章として自然に
'   なるよう、直前の文と文末が重なる候補は避けて選び直す。

Private Function EndingKey(ByVal text As String) As String
    If Len(text) <= 9 Then
        EndingKey = text
    Else
        EndingKey = Right$(text, 9)
    End If
End Function

Private Function RenderCandidate(ByVal part As Object, ByVal tpl As String, ByVal style As String) As String
    Dim t As String, lastCh As String
    t = RenderTpl(tpl, part("vars"), style)
    ' 「そのまま使う」の見取りは、指導要録用に常体へ機械変換する
    If part("raw") = True And style = "plain" Then t = ToPlain(t)
    If Len(t) > 0 Then
        lastCh = Right$(t, 1)
        Select Case lastCh
            Case ChrW$(&H3002), ChrW$(&H300D), ChrW$(&HFF09), ")"
                ' そのまま
            Case Else
                t = t & ChrW$(&H3002)
        End Select
    End If
    RenderCandidate = t
End Function

' 直前の文と文末が同じにならないよう、候補を順にずらしながら選ぶ
Private Function RenderPartVaried(ByVal part As Object, ByVal useShort As Boolean, ByVal style As String, _
                                  ByVal bump As Long, ByVal prevEnding As String) As Object
    Dim col As Collection, n As Long, startIdx As Long, k As Long, idx As Long
    Dim t As String, fallback As String, hasFallback As Boolean, res As Object

    Set res = NewDict()
    res.Add "text", ""
    res.Add "ending", prevEnding

    If useShort Then Set col = part("short") Else Set col = part("long")
    If col Is Nothing Then Set RenderPartVaried = res: Exit Function
    n = col.Count
    If n = 0 Then Set RenderPartVaried = res: Exit Function

    startIdx = (((CLng(part("seed")) + bump) Mod n) + n) Mod n

    For k = 0 To n - 1
        idx = ((startIdx + k) Mod n) + 1
        t = RenderCandidate(part, col(idx), style)
        If Not hasFallback Then
            fallback = t
            hasFallback = True
        End If
        If Len(t) > 0 Then
            If Len(prevEnding) = 0 Or EndingKey(t) <> prevEnding Then
                res("text") = t
                res("ending") = EndingKey(t)
                Set RenderPartVaried = res
                Exit Function
            End If
        End If
    Next k

    res("text") = fallback
    res("ending") = EndingKey(fallback)
    Set RenderPartVaried = res
End Function

Private Function TotalChars(ByVal rendered As Collection, ByVal upTo As Long) As Long
    Dim i As Long, it As Object, s As String, total As Long
    For i = 1 To upTo
        Set it = rendered(i)
        If it("useShort") = True Then s = CStr(it("shortText")) Else s = CStr(it("longText"))
        total = total + CountChars(s)
    Next i
    TotalChars = total
End Function

Private Function Assemble(ByVal parts As Collection, ByVal limit As Long, ByVal style As String, _
                          ByVal bump As Long) As Object
    Dim rendered As Collection, i As Long, p As Object, it As Object
    Dim longR As Object, shortR As Object
    Dim prevLongEnding As String, prevShortEnding As String
    Dim keptCount As Long, res As Object, outParts As Collection, unused As Collection
    Dim txt As String, rec As Object

    Set rendered = New Collection
    For i = 1 To parts.Count
        Set p = parts(i)
        Set longR = RenderPartVaried(p, False, style, bump, prevLongEnding)
        Set shortR = RenderPartVaried(p, True, style, bump, prevShortEnding)
        If Len(CStr(longR("text"))) > 0 Then prevLongEnding = CStr(longR("ending"))
        If Len(CStr(shortR("text"))) > 0 Then prevShortEnding = CStr(shortR("ending"))
        If Len(CStr(longR("text"))) > 0 Or Len(CStr(shortR("text"))) > 0 Then
            Set it = NewDict()
            it.Add "cat", p("cat")
            it.Add "label", p("label")
            it.Add "detail", p("detail")
            it.Add "longText", longR("text")
            it.Add "shortText", shortR("text")
            it.Add "useShort", False
            rendered.Add it
        End If
    Next i

    ' 1) 後ろ（優先度の低い方）から順に短い文例へ置き換える
    i = rendered.Count
    Do While i >= 1
        If limit <= 0 Then Exit Do
        If TotalChars(rendered, rendered.Count) <= limit Then Exit Do
        Set it = rendered(i)
        If Len(CStr(it("shortText"))) > 0 And CStr(it("shortText")) <> CStr(it("longText")) Then
            it("useShort") = True
        End If
        i = i - 1
    Loop

    ' 2) それでも収まらなければ、後ろの文から落とす（最低1文は残す）
    keptCount = rendered.Count
    Do While limit > 0 And keptCount > 1
        If TotalChars(rendered, keptCount) <= limit Then Exit Do
        keptCount = keptCount - 1
    Loop

    Set outParts = New Collection
    Set unused = New Collection
    txt = ""
    For i = 1 To rendered.Count
        Set it = rendered(i)
        Set rec = NewDict()
        rec.Add "cat", it("cat")
        rec.Add "label", it("label")
        rec.Add "detail", it("detail")
        If it("useShort") = True Then
            rec.Add "text", it("shortText")
        Else
            rec.Add "text", it("longText")
        End If
        rec.Add "shortened", it("useShort")
        If i <= keptCount Then
            outParts.Add rec
            txt = txt & CStr(rec("text"))
        Else
            unused.Add rec
        End If
    Next i

    Set res = NewDict()
    res.Add "text", txt
    res.Add "parts", outParts
    res.Add "unused", unused
    res.Add "length", CountChars(txt)
    res.Add "limit", limit
    res.Add "over", (limit > 0 And CountChars(txt) > limit)
    Set Assemble = res
End Function

'==========================================================================
' 通信表／指導要録の所見（行動・活動・見取り）
'   mode: "report"(通信表・敬体) / "youroku"(指導要録・常体)
'==========================================================================
Public Function BuildMain(ByVal st As Object, ByVal settings As Object, ByVal tpls As Object, _
                          ByVal mode As String, ByVal bump As Long) As Object
    Dim style As String, limit As Long, seed As Long, parts As Collection
    Dim order As Variant, i As Long, k As String, enabled As Object, useCat As Boolean
    Dim res As Object, closingCol As Collection, closing As String, rec As Object

    If mode = "youroku" Then style = "plain" Else style = "polite"
    If mode = "youroku" Then limit = CLng(settings("limitYouroku")) Else limit = CLng(settings("limitReport"))
    seed = CLng(st("no"))

    Set parts = New Collection
    order = settings("order")
    Set enabled = settings("enabled")

    For i = LBound(order) To UBound(order)
        k = CStr(order(i))
        If enabled.Exists(k) Then useCat = (enabled(k) = True) Else useCat = True
        If useCat Then
            Select Case k
                Case "committee": AddCommittee parts, st, tpls, seed
                Case "kakari": AddKakari parts, st, tpls, seed
                Case "event": AddEvents parts, st, tpls, seed
                Case "behavior": AddBehavior parts, st, settings, tpls, seed
                Case "observation": AddObservations parts, st, tpls, seed
            End Select
        End If
    Next i

    Set res = Assemble(parts, limit, style, bump)

    ' 通信表のみ、任意で結びの一文を添える
    If mode = "report" And settings("useClosing") = True And Len(CStr(res("text"))) > 0 Then
        Set closingCol = TplList(tpls, "closing", "", "", "long")
        If closingCol.Count > 0 Then
            closing = RenderTpl(PickOne(closingCol, seed + bump), NewDict(), style)
            If Len(closing) > 0 Then
                Set rec = NewDict()
                rec.Add "cat", "closing"
                rec.Add "label", CategoryLabel("closing")
                rec.Add "detail", ""
                rec.Add "text", closing
                rec.Add "shortened", False
                If limit <= 0 Or CountChars(CStr(res("text")) & closing) <= limit Then
                    res("text") = CStr(res("text")) & closing
                    res("parts").Add rec
                Else
                    res("unused").Add rec
                End If
            End If
        End If
    End If

    res("length") = CountChars(CStr(res("text")))
    res("over") = (limit > 0 And CLng(res("length")) > limit)
    Set BuildMain = res
End Function

'==========================================================================
' 道徳の所見（複数教材をまとめて「大くくりのまとまり」として記述）
'==========================================================================
Public Function BuildMoral(ByVal st As Object, ByVal settings As Object, ByVal tpls As Object, _
                           ByVal mode As String, ByVal bump As Long) As Object
    Dim style As String, limit As Long, seed As Long, parts As Collection
    Dim items As Collection, i As Long, m As Object, material As String, note As String
    Dim vars As Object, setKey As String, detail As String, res As Object

    If mode = "youroku" Then style = "plain" Else style = "polite"
    If mode = "youroku" Then limit = CLng(settings("moralLimitYouroku")) Else limit = CLng(settings("moralLimitReport"))
    seed = CLng(st("no"))
    Set parts = New Collection

    If settings("moralIntro") = True Then
        Set vars = NewDict()
        vars.Add ChrW$(&H5B66) & ChrW$(&H671F), CStr(settings("term"))
        parts.Add NewPart("moral-intro", ChrW$(&H5C0E) & ChrW$(&H5165), "", _
                          TplPair(tpls, "moral-intro", "", "", False), _
                          TplPair(tpls, "moral-intro", "", "", True), vars, seed + bump)
    End If

    Set items = st("moral")
    If Not items Is Nothing Then
        For i = 1 To items.Count
            Set m = items(i)
            material = Trim$(CStr(m("material")))
            If Len(material) > 0 Then
                note = Trim$(CStr(m("note")))
                If Len(note) > 0 Then setKey = "moral-material" Else setKey = "moral-material-nonote"
                Set vars = NewDict()
                vars.Add ChrW$(&H6559) & ChrW$(&H6750), material
                If Len(Trim$(CStr(m("value")))) > 0 Then
                    vars.Add ChrW$(&H9805) & ChrW$(&H76EE) & ChrW$(&H53E5), Trim$(CStr(m("value")))
                Else
                    vars.Add ChrW$(&H9805) & ChrW$(&H76EE) & ChrW$(&H53E5), _
                             ChrW$(&H5927) & ChrW$(&H5207) & ChrW$(&H306A) & ChrW$(&H3053) & ChrW$(&H3068)
                End If
                If Len(note) > 0 Then
                    If Right$(note, 1) = ChrW$(&H3002) Then note = Left$(note, Len(note) - 1)
                    vars.Add ChrW$(&H6C17) & ChrW$(&H4ED8) & ChrW$(&H304D) & ChrW$(&H53E5), _
                             ChrW$(&H300C) & note & ChrW$(&H300D)
                Else
                    vars.Add ChrW$(&H6C17) & ChrW$(&H4ED8) & ChrW$(&H304D) & ChrW$(&H53E5), ""
                End If
                detail = material
                If Len(Trim$(CStr(m("value")))) > 0 Then detail = detail & "/" & Trim$(CStr(m("value")))
                parts.Add NewPart("moral", ChrW$(&H9053) & ChrW$(&H5FB3), detail, _
                                  TplPair(tpls, setKey, "", "", False), _
                                  TplPair(tpls, setKey, "", "", True), vars, seed + i - 1 + bump)
            End If
        Next i
    End If

    If settings("moralClosing") = True And parts.Count > 0 Then
        parts.Add NewPart("moral-closing", ChrW$(&H7D50) & ChrW$(&H3073), "", _
                          TplPair(tpls, "moral-closing", "", "", False), _
                          TplPair(tpls, "moral-closing", "", "", True), NewDict(), seed + bump)
    End If

    Set res = Assemble(parts, limit, style, 0)
    Set BuildMoral = res
End Function

'==========================================================================
' 総合的な学習の時間の所見
'==========================================================================
Public Function BuildSougou(ByVal st As Object, ByVal settings As Object, ByVal tpls As Object, _
                            ByVal mode As String, ByVal bump As Long) As Object
    Dim style As String, limit As Long, seed As Long, parts As Collection
    Dim items As Collection, i As Long, g As Object, unit As String
    Dim vars As Object, detail As String, res As Object

    If mode = "youroku" Then style = "plain" Else style = "polite"
    If mode = "youroku" Then limit = CLng(settings("sougouLimitYouroku")) Else limit = CLng(settings("sougouLimitReport"))
    seed = CLng(st("no"))
    Set parts = New Collection

    Set items = st("sougou")
    If Not items Is Nothing Then
        For i = 1 To items.Count
            Set g = items(i)
            unit = Trim$(CStr(g("unit")))
            If Len(unit) > 0 Then
                Set vars = NewDict()
                vars.Add ChrW$(&H5358) & ChrW$(&H5143), unit
                If Len(Trim$(CStr(g("theme")))) > 0 Then
                    vars.Add ChrW$(&H8AB2) & ChrW$(&H984C) & ChrW$(&H53E5), _
                             WithParticle(CStr(g("theme")), ChrW$(&H306B) & ChrW$(&H3064) & ChrW$(&H3044) & ChrW$(&H3066))
                Else
                    vars.Add ChrW$(&H8AB2) & ChrW$(&H984C) & ChrW$(&H53E5), ""
                End If
                vars.Add ChrW$(&H6D3B) & ChrW$(&H52D5) & ChrW$(&H53E5), SuffixStr(CStr(g("activity")), ChrW$(&H3001))
                vars.Add ChrW$(&H6210) & ChrW$(&H679C) & ChrW$(&H53E5), SuffixStr(CStr(g("result")), ChrW$(&H3001))
                detail = unit
                If Len(Trim$(CStr(g("theme")))) > 0 Then detail = detail & "/" & Trim$(CStr(g("theme")))
                parts.Add NewPart("sougou", ChrW$(&H7DCF) & ChrW$(&H5408), detail, _
                                  TplPair(tpls, "sougou", "", "", False), _
                                  TplPair(tpls, "sougou", "", "", True), vars, seed + i - 1 + bump)
            End If
        Next i
    End If

    If settings("sougouClosing") = True And parts.Count > 0 Then
        parts.Add NewPart("sougou-closing", ChrW$(&H7D50) & ChrW$(&H3073), "", _
                          TplPair(tpls, "sougou-closing", "", "", False), _
                          TplPair(tpls, "sougou-closing", "", "", True), NewDict(), seed + bump)
    End If

    Set res = Assemble(parts, limit, style, 0)
    Set BuildSougou = res
End Function
