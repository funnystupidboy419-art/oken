Attribute VB_Name = "modTemplates"
'==========================================================================
' modTemplates  文例テンプレートの読み込み
'   文例は「文例」シートに置いてあります。
'   シートを直せば、そのまま生成結果に反映されます（VBAの編集は不要）。
'
'   文例シートの列
'     A 区分(説明)  B キー  C 項目  D 記号  E 長短  F 文例
'==========================================================================
Option Explicit

'--- シート名（日本語はコード化けを避けるため ChrW で表記）-----------------
Public Function ShSettings() As String
    ShSettings = ChrW$(&H8A2D) & ChrW$(&H5B9A)                                   ' 設定
End Function
Public Function ShRoster() As String
    ShRoster = ChrW$(&H59D4) & ChrW$(&H54E1) & ChrW$(&H4F1A) & ChrW$(&H30FB) & ChrW$(&H4FC2)  ' 委員会・係
End Function
Public Function ShEvent() As String
    ShEvent = ChrW$(&H884C) & ChrW$(&H4E8B)                                      ' 行事
End Function
Public Function ShBehavior() As String
    ShBehavior = ChrW$(&H884C) & ChrW$(&H52D5) & ChrW$(&H306E) & ChrW$(&H8A18) & ChrW$(&H9332)  ' 行動の記録
End Function
Public Function ShObs() As String
    ShObs = ChrW$(&H898B) & ChrW$(&H53D6) & ChrW$(&H308A)                        ' 見取り
End Function
Public Function ShMoral() As String
    ShMoral = ChrW$(&H9053) & ChrW$(&H5FB3)                                      ' 道徳
End Function
Public Function ShSougou() As String
    ShSougou = ChrW$(&H7DCF) & ChrW$(&H5408)                                     ' 総合
End Function
Public Function ShTpl() As String
    ShTpl = ChrW$(&H6587) & ChrW$(&H4F8B)                                        ' 文例
End Function
Public Function ShOut() As String
    ShOut = ChrW$(&H6240) & ChrW$(&H898B)                                        ' 所見
End Function

Public Function MarkA() As String
    MarkA = ChrW$(&H25CE)                                                        ' ◎
End Function
Public Function MarkB() As String
    MarkB = ChrW$(&H25CB)                                                        ' ○
End Function

'--- 行動の記録 10項目（キーと表示名）-------------------------------------
Public Function BehaviorKeys() As Variant
    BehaviorKeys = Array("seikatsu", "kenko", "jishu", "sekinin", "soui", _
                         "omoiyari", "seimei", "kinro", "kosei", "kokyo")
End Function

Public Function BehaviorLabel(ByVal key As String) As String
    Select Case key
        Case "seikatsu": BehaviorLabel = ChrW$(&H57FA) & ChrW$(&H672C) & ChrW$(&H7684) & ChrW$(&H306A) & ChrW$(&H751F) & ChrW$(&H6D3B) & ChrW$(&H7FD2) & ChrW$(&H6163)
        Case "kenko": BehaviorLabel = ChrW$(&H5065) & ChrW$(&H5EB7) & ChrW$(&H30FB) & ChrW$(&H4F53) & ChrW$(&H529B) & ChrW$(&H306E) & ChrW$(&H5411) & ChrW$(&H4E0A)
        Case "jishu": BehaviorLabel = ChrW$(&H81EA) & ChrW$(&H4E3B) & ChrW$(&H30FB) & ChrW$(&H81EA) & ChrW$(&H5F8B)
        Case "sekinin": BehaviorLabel = ChrW$(&H8CAC) & ChrW$(&H4EFB) & ChrW$(&H611F)
        Case "soui": BehaviorLabel = ChrW$(&H5275) & ChrW$(&H610F) & ChrW$(&H5DE5) & ChrW$(&H592B)
        Case "omoiyari": BehaviorLabel = ChrW$(&H601D) & ChrW$(&H3044) & ChrW$(&H3084) & ChrW$(&H308A) & ChrW$(&H30FB) & ChrW$(&H5354) & ChrW$(&H529B)
        Case "seimei": BehaviorLabel = ChrW$(&H751F) & ChrW$(&H547D) & ChrW$(&H5C0A) & ChrW$(&H91CD) & ChrW$(&H30FB) & ChrW$(&H81EA) & ChrW$(&H7136) & ChrW$(&H611B) & ChrW$(&H8B77)
        Case "kinro": BehaviorLabel = ChrW$(&H52E4) & ChrW$(&H52B4) & ChrW$(&H30FB) & ChrW$(&H5949) & ChrW$(&H4ED5)
        Case "kosei": BehaviorLabel = ChrW$(&H516C) & ChrW$(&H6B63) & ChrW$(&H30FB) & ChrW$(&H516C) & ChrW$(&H5E73)
        Case "kokyo": BehaviorLabel = ChrW$(&H516C) & ChrW$(&H5171) & ChrW$(&H5FC3) & ChrW$(&H30FB) & ChrW$(&H516C) & ChrW$(&H5FB3) & ChrW$(&H5FC3)
        Case Else: BehaviorLabel = key
    End Select
End Function

'--- カテゴリの表示名 -----------------------------------------------------
Public Function CategoryLabel(ByVal key As String) As String
    Select Case key
        Case "committee": CategoryLabel = ChrW$(&H59D4) & ChrW$(&H54E1) & ChrW$(&H4F1A)
        Case "kakari": CategoryLabel = ChrW$(&H4FC2)
        Case "event": CategoryLabel = ChrW$(&H884C) & ChrW$(&H4E8B)
        Case "behavior": CategoryLabel = ChrW$(&H884C) & ChrW$(&H52D5) & ChrW$(&H306E) & ChrW$(&H8A18) & ChrW$(&H9332)
        Case "observation": CategoryLabel = ChrW$(&H65E5) & ChrW$(&H5E38) & ChrW$(&H306E) & ChrW$(&H898B) & ChrW$(&H53D6) & ChrW$(&H308A)
        Case "closing": CategoryLabel = ChrW$(&H7D50) & ChrW$(&H3073)
        Case Else: CategoryLabel = key
    End Select
End Function

'--------------------------------------------------------------------------
' 文例シートを読み込み、Dictionary(キー -> Collection(文例)) にする
'   辞書キー:  キー|項目|記号|長短
'--------------------------------------------------------------------------
Public Function LoadTemplates() As Object
    Dim ws As Worksheet, d As Object, lastRow As Long, r As Long
    Dim k As String, item As String, mk As String, va As String, tx As String
    Dim dictKey As String, col As Collection

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(ShTpl())
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    For r = 2 To lastRow
        k = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(k) > 0 Then
            item = Trim$(CStr(ws.Cells(r, 3).Value))
            mk = Trim$(CStr(ws.Cells(r, 4).Value))
            va = LCase$(Trim$(CStr(ws.Cells(r, 5).Value)))
            tx = CStr(ws.Cells(r, 6).Value)
            If Len(tx) > 0 Then
                dictKey = k & "|" & item & "|" & mk & "|" & va
                If d.Exists(dictKey) Then
                    Set col = d(dictKey)
                Else
                    Set col = New Collection
                    d.Add dictKey, col
                End If
                col.Add tx
            End If
        End If
    Next r

    Set LoadTemplates = d
End Function

'--------------------------------------------------------------------------
' 文例リストの取り出し（無ければ空のコレクション）
'--------------------------------------------------------------------------
Public Function TplList(ByVal tpls As Object, ByVal k As String, ByVal item As String, _
                        ByVal mk As String, ByVal va As String) As Collection
    Dim dictKey As String
    dictKey = k & "|" & item & "|" & mk & "|" & va
    If tpls.Exists(dictKey) Then
        Set TplList = tpls(dictKey)
    Else
        Set TplList = New Collection
    End If
End Function

'--------------------------------------------------------------------------
' 長文が無いときは短文を、短文が無いときは長文を使う
'--------------------------------------------------------------------------
Public Function TplPair(ByVal tpls As Object, ByVal k As String, ByVal item As String, _
                        ByVal mk As String, ByVal wantShort As Boolean) As Collection
    Dim primary As Collection, fallback As Collection
    If wantShort Then
        Set primary = TplList(tpls, k, item, mk, "short")
        Set fallback = TplList(tpls, k, item, mk, "long")
    Else
        Set primary = TplList(tpls, k, item, mk, "long")
        Set fallback = TplList(tpls, k, item, mk, "short")
    End If
    If primary.Count > 0 Then
        Set TplPair = primary
    Else
        Set TplPair = fallback
    End If
End Function

'--------------------------------------------------------------------------
' 見取りの文末パターン（キー -> テンプレート）
'--------------------------------------------------------------------------
Public Function ObsPattern(ByVal tpls As Object, ByVal key As String) As String
    Dim col As Collection
    Set col = TplList(tpls, "obs", key, "", "")
    If col.Count > 0 Then
        ObsPattern = col(1)
    Else
        Set col = TplList(tpls, "obs", "sugata", "", "")
        If col.Count > 0 Then ObsPattern = col(1)
    End If
End Function
