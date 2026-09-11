Attribute VB_Name = "modUtil"
'==========================================================================
' modUtil  共通ユーティリティ
'   文字数 / テンプレート展開 / 文体変換 / 記号の正規化 / 個人情報チェック
'   ※ 外部ライブラリ・外部通信は一切使用しません。
'==========================================================================
Option Explicit

Private Const MARK_CHAR As Long = &HE000&

'--------------------------------------------------------------------------
' 文字数
'   通信表・指導要録は「句読点を含む文字数」で数えるのが一般的。改行は数えない。
'--------------------------------------------------------------------------
Public Function CountChars(ByVal s As String) As Long
    If Len(s) = 0 Then Exit Function
    s = Replace$(s, vbCrLf, "")
    s = Replace$(s, vbLf, "")
    s = Replace$(s, vbCr, "")
    CountChars = Len(s)
End Function

Public Function NewRegExp(ByVal pattern As String, Optional ByVal globalFlag As Boolean = True) As Object
    Set NewRegExp = CreateObject("VBScript.RegExp")
    NewRegExp.pattern = pattern
    NewRegExp.Global = globalFlag
    NewRegExp.MultiLine = False
    NewRegExp.IgnoreCase = False
End Function

Public Function Trim2(ByVal s As String) As String
    Trim2 = Trim$(Replace$(Replace$(s, vbTab, " "), ChrW$(&H3000), " "))
End Function

'--------------------------------------------------------------------------
' 余分な空白・重複した読点などを整える
'--------------------------------------------------------------------------
Public Function Tidy(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace$(s, " ", "")
    s = Replace$(s, ChrW$(&H3000), "")
    Do While InStr(s, ChrW$(&H3001) & ChrW$(&H3001)) > 0
        s = Replace$(s, ChrW$(&H3001) & ChrW$(&H3001), ChrW$(&H3001))
    Loop
    Do While InStr(s, ChrW$(&H3002) & ChrW$(&H3002)) > 0
        s = Replace$(s, ChrW$(&H3002) & ChrW$(&H3002), ChrW$(&H3002))
    Loop
    s = Replace$(s, ChrW$(&H3001) & ChrW$(&H3002), ChrW$(&H3002))
    Do While Len(s) > 0
        If Left$(s, 1) = ChrW$(&H3001) Or Left$(s, 1) = ChrW$(&H3002) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    Tidy = Trim$(s)
End Function

'--------------------------------------------------------------------------
' テンプレート展開
'   {{キー}}      … 差し込み
'   {敬体|常体}   … 文体による出し分け
'--------------------------------------------------------------------------
Public Function RenderTpl(ByVal tpl As String, ByVal vars As Object, ByVal style As String) As String
    Dim s As String, re As Object, ms As Object, i As Long, k As String, v As String
    If Len(tpl) = 0 Then Exit Function
    s = tpl

    Set re = NewRegExp("\{([^{}|]*)\|([^{}|]*)\}")
    Set ms = re.Execute(s)
    For i = ms.Count - 1 To 0 Step -1
        If style = "plain" Then
            v = ms(i).SubMatches(1)
        Else
            v = ms(i).SubMatches(0)
        End If
        s = Left$(s, ms(i).FirstIndex) & v & Mid$(s, ms(i).FirstIndex + ms(i).Length + 1)
    Next i

    Set re = NewRegExp("\{\{([^{}]+)\}\}")
    Set ms = re.Execute(s)
    For i = ms.Count - 1 To 0 Step -1
        k = Replace$(Replace$(ms(i).SubMatches(0), " ", ""), ChrW$(&H3000), "")
        v = ""
        If Not vars Is Nothing Then
            If vars.Exists(k) Then v = CStr(vars(k))
        End If
        s = Left$(s, ms(i).FirstIndex) & v & Mid$(s, ms(i).FirstIndex + ms(i).Length + 1)
    Next i

    RenderTpl = Tidy(s)
End Function

'--------------------------------------------------------------------------
' 全角数字を含む文字列から出席番号を取り出す（取れなければ 0）
'--------------------------------------------------------------------------
Public Function ToNumber(ByVal v As Variant) As Long
    Dim s As String, out As String, i As Long, c As String, code As Long
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If code >= &HFF10& And code <= &HFF19& Then
            out = out & Chr$(code - &HFEE0&)
        ElseIf c >= "0" And c <= "9" Then
            out = out & c
        End If
    Next i
    If Len(out) = 0 Then Exit Function
    ToNumber = CLng(Left$(out, 9))
End Function

'--------------------------------------------------------------------------
' 評価記号の正規化 (◎ / ○ / 空)
'--------------------------------------------------------------------------
Public Function NormalizeMark(ByVal v As Variant) As String
    Dim s As String, re As Object
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    s = Replace$(Replace$(Replace$(s, " ", ""), ChrW$(&H3000), ""), vbTab, "")
    If Len(s) = 0 Then Exit Function
    Set re = NewRegExp("^(" & ChrW$(&H25CE) & "|" & ChrW$(&HFF20) & "|@|A|a|" & ChrW$(&HFF41) & "|" & ChrW$(&HFF21) & ")")
    If re.Test(s) Then NormalizeMark = ChrW$(&H25CE): Exit Function
    Set re = NewRegExp("^(" & ChrW$(&H25CB) & "|" & ChrW$(&H3007) & "|" & ChrW$(&H25EF) & "|O|o|" & ChrW$(&HFF2F) & "|" & ChrW$(&HFF4F) & "|B|b|1)")
    If re.Test(s) Then NormalizeMark = ChrW$(&H25CB)
End Function

'==========================================================================
' 文体変換（敬体 → 常体）
'   自由記述欄の補助機能。機械的な変換のため、指導要録に使う前に
'   必ず内容を目視で確認すること。
'==========================================================================

Private Function ItoU(ByVal c As String) As String
    Select Case c
        Case ChrW$(&H3044): ItoU = ChrW$(&H3046)
        Case ChrW$(&H304D): ItoU = ChrW$(&H304F)
        Case ChrW$(&H304E): ItoU = ChrW$(&H3050)
        Case ChrW$(&H3057): ItoU = ChrW$(&H3059)
        Case ChrW$(&H3061): ItoU = ChrW$(&H3064)
        Case ChrW$(&H306B): ItoU = ChrW$(&H306C)
        Case ChrW$(&H3072): ItoU = ChrW$(&H3075)
        Case ChrW$(&H3073): ItoU = ChrW$(&H3076)
        Case ChrW$(&H307F): ItoU = ChrW$(&H3080)
        Case ChrW$(&H308A): ItoU = ChrW$(&H308B)
    End Select
End Function

Private Function ItoA(ByVal c As String) As String
    Select Case c
        Case ChrW$(&H3044): ItoA = ChrW$(&H308F)
        Case ChrW$(&H304D): ItoA = ChrW$(&H304B)
        Case ChrW$(&H304E): ItoA = ChrW$(&H304C)
        Case ChrW$(&H3057): ItoA = ChrW$(&H3055)
        Case ChrW$(&H3061): ItoA = ChrW$(&H305F)
        Case ChrW$(&H306B): ItoA = ChrW$(&H306A)
        Case ChrW$(&H3072): ItoA = ChrW$(&H306F)
        Case ChrW$(&H3073): ItoA = ChrW$(&H3070)
        Case ChrW$(&H307F): ItoA = ChrW$(&H307E)
        Case ChrW$(&H308A): ItoA = ChrW$(&H3089)
    End Select
End Function

' 語幹がイ段でも一段活用になる代表語（例外辞書）
Private Function IchidanTails() As Variant
    IchidanTails = Array( _
        ChrW$(&H3067) & ChrW$(&H304D), _
        ChrW$(&H898B), ChrW$(&H89B3), ChrW$(&H7740), _
        ChrW$(&H8D77) & ChrW$(&H304D), _
        ChrW$(&H843D) & ChrW$(&H3061), _
        ChrW$(&H751F) & ChrW$(&H304D), _
        ChrW$(&H904E) & ChrW$(&H304E), _
        ChrW$(&H611F) & ChrW$(&H3058), _
        ChrW$(&H4FE1) & ChrW$(&H3058), _
        ChrW$(&H5FDC) & ChrW$(&H3058), _
        ChrW$(&H9589) & ChrW$(&H3058), _
        ChrW$(&H8DB3) & ChrW$(&H308A), _
        ChrW$(&H501F) & ChrW$(&H308A), _
        ChrW$(&H7528) & ChrW$(&H3044), _
        ChrW$(&H7387) & ChrW$(&H3044), _
        ChrW$(&H5831) & ChrW$(&H3044), _
        ChrW$(&H6D74) & ChrW$(&H3073), _
        ChrW$(&H5EF6) & ChrW$(&H3073), _
        ChrW$(&H4F38) & ChrW$(&H3073))
End Function

Private Function IsIchidan(ByVal stem As String) As Boolean
    Dim re As Object, tails As Variant, i As Long, t As String
    If Len(stem) = 0 Then Exit Function
    Set re = NewRegExp("[" & ChrW$(&H3048) & ChrW$(&H3051) & ChrW$(&H3052) & ChrW$(&H305B) & ChrW$(&H305C) & _
                       ChrW$(&H3066) & ChrW$(&H3067) & ChrW$(&H306D) & ChrW$(&H3078) & ChrW$(&H3079) & _
                       ChrW$(&H307A) & ChrW$(&H3081) & ChrW$(&H308C) & "]$")
    If re.Test(stem) Then IsIchidan = True: Exit Function
    tails = IchidanTails()
    For i = LBound(tails) To UBound(tails)
        t = tails(i)
        If Len(stem) >= Len(t) Then
            If Right$(stem, Len(t)) = t Then IsIchidan = True: Exit Function
        End If
    Next i
End Function

' サ変（〜する）かどうか：「活動し」「チャレンジし」など
Private Function IsSuru(ByVal stem As String) As Boolean
    Dim re As Object
    If Len(stem) = 0 Then IsSuru = True: Exit Function
    If stem = ChrW$(&H3057) Then IsSuru = True: Exit Function
    Set re = NewRegExp("(?:[" & ChrW$(&H4E00) & "-" & ChrW$(&H9FA0) & ChrW$(&H3005) & "]{2,}|[" & _
                       ChrW$(&H30A1) & "-" & ChrW$(&H30F6) & ChrW$(&H30FC) & "]{2,})" & ChrW$(&H3057) & "$")
    IsSuru = re.Test(stem)
End Function

Private Function SuruBase(ByVal stem As String) As String
    If Len(stem) = 0 Or stem = ChrW$(&H3057) Then Exit Function
    SuruBase = Left$(stem, Len(stem) - 1)
End Function

' 「〜ます」の語幹 → 終止形
Private Function PlainDict(ByVal stem As String) As String
    Dim last As String, u As String
    If IsSuru(stem) Then PlainDict = SuruBase(stem) & ChrW$(&H3059) & ChrW$(&H308B): Exit Function
    If stem = ChrW$(&H6765) Or stem = ChrW$(&H304D) Then PlainDict = ChrW$(&H6765) & ChrW$(&H308B): Exit Function
    If IsIchidan(stem) Then PlainDict = stem & ChrW$(&H308B): Exit Function
    last = Right$(stem, 1)
    u = ItoU(last)
    If Len(u) > 0 Then
        PlainDict = Left$(stem, Len(stem) - 1) & u
    Else
        PlainDict = stem & ChrW$(&H308B)
    End If
End Function

' 「〜ました」の語幹 → 過去形（音便あり）
Private Function PlainPast(ByVal stem As String) As String
    Dim last As String, head As String, re As Object
    Dim TA As String, DA As String, TTA As String, NDA As String
    TA = ChrW$(&H305F): DA = ChrW$(&H3060)
    TTA = ChrW$(&H3063) & TA: NDA = ChrW$(&H3093) & DA

    If IsSuru(stem) Then PlainPast = SuruBase(stem) & ChrW$(&H3057) & TA: Exit Function
    If stem = ChrW$(&H6765) Or stem = ChrW$(&H304D) Then PlainPast = ChrW$(&H6765) & TA: Exit Function
    If IsIchidan(stem) Then PlainPast = stem & TA: Exit Function

    ' 行く は促音便
    Set re = NewRegExp("(" & ChrW$(&H884C) & ChrW$(&H304D) & "|" & ChrW$(&H3044) & ChrW$(&H304D) & ")$")
    If re.Test(stem) Then PlainPast = Left$(stem, Len(stem) - 1) & TTA: Exit Function

    last = Right$(stem, 1)
    head = Left$(stem, Len(stem) - 1)
    Select Case last
        Case ChrW$(&H304D): PlainPast = head & ChrW$(&H3044) & TA
        Case ChrW$(&H304E): PlainPast = head & ChrW$(&H3044) & DA
        Case ChrW$(&H3057): PlainPast = head & ChrW$(&H3057) & TA
        Case ChrW$(&H3044), ChrW$(&H3061), ChrW$(&H308A): PlainPast = head & TTA
        Case ChrW$(&H306B), ChrW$(&H3073), ChrW$(&H307F): PlainPast = head & NDA
        Case Else: PlainPast = stem & TA
    End Select
End Function

' 「〜ません」の語幹 → 否定形
Private Function PlainNeg(ByVal stem As String) As String
    Dim last As String, a As String, NAI As String
    NAI = ChrW$(&H306A) & ChrW$(&H3044)
    If IsSuru(stem) Then PlainNeg = SuruBase(stem) & ChrW$(&H3057) & NAI: Exit Function
    If stem = ChrW$(&H6765) Or stem = ChrW$(&H304D) Then PlainNeg = ChrW$(&H6765) & NAI: Exit Function
    If stem = ChrW$(&H3042) & ChrW$(&H308A) Then PlainNeg = NAI: Exit Function
    If IsIchidan(stem) Then PlainNeg = stem & NAI: Exit Function
    last = Right$(stem, 1)
    a = ItoA(last)
    If Len(a) > 0 Then
        PlainNeg = Left$(stem, Len(stem) - 1) & a & NAI
    Else
        PlainNeg = stem & NAI
    End If
End Function

Private Function PlainNegPast(ByVal stem As String) As String
    Dim s As String, NAI As String, NAKATTA As String
    NAI = ChrW$(&H306A) & ChrW$(&H3044)
    NAKATTA = ChrW$(&H306A) & ChrW$(&H304B) & ChrW$(&H3063) & ChrW$(&H305F)
    s = PlainNeg(stem)
    If Right$(s, Len(NAI)) = NAI Then
        PlainNegPast = Left$(s, Len(s) - Len(NAI)) & NAKATTA
    Else
        PlainNegPast = s
    End If
End Function

' 語幹 + 活用の一括置換（後ろから置き換えて位置ずれを防ぐ）
Private Function ReplaceStems(ByVal s As String, ByVal tail As String, ByVal mode As Long) As String
    Dim re As Object, ms As Object, i As Long, st As String, rep As String
    Dim STEM As String
    STEM = "([" & ChrW$(&H3041) & "-" & ChrW$(&H3093) & ChrW$(&H30A1) & "-" & ChrW$(&H30F6) & _
           ChrW$(&H4E00) & "-" & ChrW$(&H9FA0) & ChrW$(&H3005) & ChrW$(&H30FC) & "]+?)"
    Set re = NewRegExp(STEM & tail)
    Set ms = re.Execute(s)
    For i = ms.Count - 1 To 0 Step -1
        st = ms(i).SubMatches(0)
        Select Case mode
            Case 1: rep = PlainNegPast(st)
            Case 2: rep = PlainNeg(st)
            Case 3
                rep = PlainDict(st)
                If Right$(rep, 1) = ChrW$(&H308B) Then
                    rep = Left$(rep, Len(rep) - 1) & ChrW$(&H3088) & ChrW$(&H3046)
                End If
            Case 4: rep = PlainPast(st)
            Case 5
                rep = PlainPast(st)
                If Right$(rep, 1) = ChrW$(&H305F) Then
                    rep = Left$(rep, Len(rep) - 1) & ChrW$(&H3066)
                ElseIf Right$(rep, 1) = ChrW$(&H3060) Then
                    rep = Left$(rep, Len(rep) - 1) & ChrW$(&H3067)
                End If
            Case 6: rep = PlainDict(st) & ChrW$(&H306E) & ChrW$(&H3067)
            Case Else: rep = PlainDict(st)
        End Select
        s = Left$(s, ms(i).FirstIndex) & rep & Mid$(s, ms(i).FirstIndex + ms(i).Length + 1)
    Next i
    ReplaceStems = s
End Function

Public Function ToPlain(ByVal text As String) As String
    Dim s As String, MK As String
    Dim MASU As String, MASHITA As String, MASEN As String
    If Len(text) = 0 Then Exit Function
    s = text
    MK = ChrW$(MARK_CHAR)
    MASU = ChrW$(&H307E) & ChrW$(&H3059)
    MASHITA = MASU & ChrW$(&H3057) & ChrW$(&H305F)
    MASEN = ChrW$(&H307E) & ChrW$(&H305B) & ChrW$(&H3093)

    ' 「ますます」を退避
    s = Replace$(s, MASU & MASU, MK)

    ' 「〜ています」「〜ておりました」など補助動詞の形
    s = RegexReplace(s, "(" & ChrW$(&H3066) & "|" & ChrW$(&H3067) & ")(?:" & ChrW$(&H3044) & "|" & _
        ChrW$(&H304A) & ChrW$(&H308A) & ")" & MASEN & ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F), _
        "$1" & ChrW$(&H3044) & ChrW$(&H306A) & ChrW$(&H304B) & ChrW$(&H3063) & ChrW$(&H305F))
    s = RegexReplace(s, "(" & ChrW$(&H3066) & "|" & ChrW$(&H3067) & ")(?:" & ChrW$(&H3044) & "|" & _
        ChrW$(&H304A) & ChrW$(&H308A) & ")" & MASEN, "$1" & ChrW$(&H3044) & ChrW$(&H306A) & ChrW$(&H3044))
    s = RegexReplace(s, "(" & ChrW$(&H3066) & "|" & ChrW$(&H3067) & ")(?:" & ChrW$(&H3044) & "|" & _
        ChrW$(&H304A) & ChrW$(&H308A) & ")" & MASHITA, "$1" & ChrW$(&H3044) & ChrW$(&H305F))
    s = RegexReplace(s, "(" & ChrW$(&H3066) & "|" & ChrW$(&H3067) & ")(?:" & ChrW$(&H3044) & "|" & _
        ChrW$(&H304A) & ChrW$(&H308A) & ")" & MASU, "$1" & ChrW$(&H3044) & ChrW$(&H308B))

    ' 名詞述語・形容動詞
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H306F) & ChrW$(&H3042) & ChrW$(&H308A) & MASEN & ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F), _
                    ChrW$(&H3067) & ChrW$(&H306F) & ChrW$(&H306A) & ChrW$(&H304B) & ChrW$(&H3063) & ChrW$(&H305F))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H306F) & ChrW$(&H3042) & ChrW$(&H308A) & MASEN, _
                    ChrW$(&H3067) & ChrW$(&H306F) & ChrW$(&H306A) & ChrW$(&H3044))
    s = Replace$(s, ChrW$(&H3042) & ChrW$(&H308A) & MASEN & ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F), _
                    ChrW$(&H306A) & ChrW$(&H304B) & ChrW$(&H3063) & ChrW$(&H305F))
    s = Replace$(s, ChrW$(&H3042) & ChrW$(&H308A) & MASEN, ChrW$(&H306A) & ChrW$(&H3044))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H3087) & ChrW$(&H3046), ChrW$(&H3060) & ChrW$(&H308D) & ChrW$(&H3046))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F), ChrW$(&H3060) & ChrW$(&H3063) & ChrW$(&H305F))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3059) & ChrW$(&H304C), ChrW$(&H3060) & ChrW$(&H304C))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3059) & ChrW$(&H306E) & ChrW$(&H3067), ChrW$(&H306A) & ChrW$(&H306E) & ChrW$(&H3067))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H3066), ChrW$(&H3067))
    s = Replace$(s, ChrW$(&H3067) & ChrW$(&H3059), ChrW$(&H3060))
    s = Replace$(s, ChrW$(&H3054) & ChrW$(&H3056) & ChrW$(&H3044) & MASU, ChrW$(&H3042) & ChrW$(&H308B))

    ' 動詞
    s = ReplaceStems(s, MASEN & ChrW$(&H3067) & ChrW$(&H3057) & ChrW$(&H305F), 1)
    s = ReplaceStems(s, MASEN, 2)
    s = ReplaceStems(s, ChrW$(&H307E) & ChrW$(&H3057) & ChrW$(&H3087) & ChrW$(&H3046), 3)
    s = ReplaceStems(s, MASHITA, 4)
    s = ReplaceStems(s, ChrW$(&H307E) & ChrW$(&H3057) & ChrW$(&H3066), 5)
    s = ReplaceStems(s, MASU & ChrW$(&H306E) & ChrW$(&H3067), 6)
    s = ReplaceStems(s, MASU, 7)

    s = Replace$(s, MK, MASU & MASU)
    ToPlain = s
End Function

Public Function RegexReplace(ByVal s As String, ByVal pattern As String, ByVal repl As String) As String
    Dim re As Object
    Set re = NewRegExp(pattern)
    RegexReplace = re.Replace(s, repl)
End Function

'--------------------------------------------------------------------------
' 個人情報チェック（氏名らしき表現の検出）
'   ※ このアプリは出席番号のみで管理します。氏名は入力しません。
'--------------------------------------------------------------------------
Private Function PersonExclude() As Variant
    PersonExclude = Array( _
        ChrW$(&H304A) & ChrW$(&H6BCD) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H304B) & ChrW$(&H3042) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H6BCD) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H7236) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H3068) & ChrW$(&H3046) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H7236) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H5144) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H5144) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H59C9) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H59C9) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H3070) & ChrW$(&H3042) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H3058) & ChrW$(&H3044) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H3070) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H3058) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H8D64) & ChrW$(&H3061) & ChrW$(&H3083) & ChrW$(&H3093), _
        ChrW$(&H307F) & ChrW$(&H306A) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H7686) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H304A) & ChrW$(&H5BA2) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H5E97) & ChrW$(&H54E1) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H8FB2) & ChrW$(&H5BB6) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H8077) & ChrW$(&H54E1) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H904B) & ChrW$(&H8EE2) & ChrW$(&H624B) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H99C5) & ChrW$(&H54E1) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H5927) & ChrW$(&H5DE5) & ChrW$(&H3055) & ChrW$(&H3093), _
        ChrW$(&H6F01) & ChrW$(&H5E2B) & ChrW$(&H3055) & ChrW$(&H3093))
End Function

' 氏名らしい表現を見つけたら、その語をカンマ区切りで返す（無ければ空文字）
Public Function FindPersonalInfo(ByVal text As String) As String
    Dim re As Object, ms As Object, i As Long, j As Long
    Dim w As String, ex As Variant, skip As Boolean, hits As String
    Dim SAN As String, YA As String
    If Len(text) = 0 Then Exit Function
    SAN = ChrW$(&H3055) & ChrW$(&H3093)
    YA = ChrW$(&H5C4B) & SAN
    ex = PersonExclude()

    Set re = NewRegExp("[" & ChrW$(&H4E00) & "-" & ChrW$(&H9FA0) & ChrW$(&H3005) & ChrW$(&H3041) & "-" & _
        ChrW$(&H3093) & ChrW$(&H30A1) & "-" & ChrW$(&H30F6) & ChrW$(&H30FC) & "]{1,4}(" & SAN & "|" & _
        ChrW$(&H304F) & ChrW$(&H3093) & "|" & ChrW$(&H541B) & "|" & ChrW$(&H3061) & ChrW$(&H3083) & ChrW$(&H3093) & ")")
    Set ms = re.Execute(text)
    For i = 0 To ms.Count - 1
        w = ms(i).Value
        skip = False
        For j = LBound(ex) To UBound(ex)
            If w = ex(j) Then skip = True: Exit For
        Next j
        If Not skip Then
            If Len(w) >= Len(YA) Then
                If Right$(w, Len(YA)) = YA Then skip = True
            End If
        End If
        If Not skip Then
            If InStr(hits, w) = 0 Then
                If Len(hits) > 0 Then hits = hits & "/"
                hits = hits & w
            End If
        End If
    Next i
    FindPersonalInfo = hits
End Function
