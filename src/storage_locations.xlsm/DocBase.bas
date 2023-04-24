Attribute VB_Name = "DocBase"
Option Explicit
'VBAを好まないためRustの命名規則を適用した
Const MAIN_SHEET As String = "保管検索"
Const TARGET_SHEET As String = "備品管理一覧"
Const PURPOSE_CELL As String = "B2"
Const MSG_CELL As String = "C3"

'検索ユースケース
Private Sub search_usecase()
    '検索文字列
    Dim purpose As Range: Set purpose = Worksheets(MAIN_SHEET).Range(PURPOSE_CELL)
    '検索結果
    Dim results_dto() As Integer: ReDim results_dto(0)

    '結果をレンダリング
    presenter doc_dao(purpose, results_dto())
    '初期化
    Erase results_dto
End Sub

'文書保管DAO
Private Function doc_dao(ByVal purpose As String, ByRef results_dto() As Integer) As Integer()
    '検索対象シート(マスタ)
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim i As Integer
    Dim key As Variant
    '検索ロジック
    For i = 1 To tar.Cells(Rows.count, "A").End(xlUp).Row
        'FIXME: 要件追加のたびにここが長くなる
        If tar.Cells(i, "G") Like "*" & purpose & "*" Or tar.Cells(i, "O") Like "*" & purpose Then
            Dim is_registered As Boolean: is_registered = False
            'FIXME: ここに書くとDAOがドメイン知識持ってるみたいになるためサービス化したい
            For Each key In results_dto
                If key = i Then
                    '登録しない
                    is_registered = True
                End If
            Next
            If Not is_registered Then
                ReDim Preserve results_dto(UBound(results_dto) + 1)
                results_dto(UBound(results_dto)) = i
            End If
        End If
    Next

    doc_dao = results_dto()

End Function

'表示
Private Function presenter(ByRef results_dto() As Integer)
    Dim main As Worksheet: Set main = Worksheets(MAIN_SHEET)
    Dim msg As Range: Set msg = main.Range(MSG_CELL)
    Dim key As Variant
    Dim tar As Worksheet: Set tar = Worksheets(TARGET_SHEET)
    Dim count As Integer: count = UBound(results_dto)
    Dim j As Integer: j = 1

    'シートの保護解除
    main.Unprotect "tyco"

    '初期化
    Rows("6:6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    '一覧化
    For Each key In results_dto
        If key <> 0 Then
            main.Cells(4 + j * 2, "B") = tar.Cells(key, "B")
            main.Hyperlinks.Add Anchor:=main.Cells(4 + j * 2, "E"), Address:="", SubAddress:= _
        "備品管理一覧!A" & key, TextToDisplay:=tar.Cells(key, "F").Value
            j = j + 1
        End If
    Next key

    If count > 0 Then
        msg = count & "件見つかりました．"
    Else
        msg = "見つかりませんでした．"
    End If

    'セル戻り
    Range("A1").Select

    'シートの保護有効化
    main.Protect "tyco"

End Function