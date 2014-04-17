Attribute VB_Name = "Main"
Option Explicit

' created by D*isuke YAMAKAWA (ClockAhead)

' change log :
' 2014/4/15  rev.3
'       + 墨付き括弧内の下線は消すように
'       + 英文明細書の補正書の書式に合わせた処理追加
'               ref: http://ameblo.jp/gidgeerock/entry-11492099859.html
'       + リファクタリング。経過状態の表示処理削除
'
' 2012/6/27  rev.2
'
' 2011/7/29  rev.1
'       + 処理の経過状態を表示するように　（ステータスバー使用）
'           画面が白く飛ぶのを防止 (メッセージポンプを強制的にまわすように)
'       + 前処理（変更履歴をＯＦＦに）と後処理（変更履歴の反映）追加

Public Sub 変更箇所に下線を引く_JP()
    ActiveDocument.TrackRevisions = False
    Call AddUnderlineToRevisionInsert
    WordBasic.AcceptAllChangesInDoc
    Call ClearUnderlineInBracket
End Sub

Public Sub 変更箇所に下線破線を引く_EN()
    ActiveDocument.TrackRevisions = False
    Call AddUnderlineToRevisionInsert
    Call AddStrikeThroughToRevisionDelete
    WordBasic.AcceptAllChangesInDoc
End Sub

'---------------------------------------------------------------------------------------------
Private Sub AddUnderlineToRevisionInsert()
    Dim myRev As Revision: For Each myRev In ActiveDocument.Revisions
        If myRev.Type = wdRevisionInsert Then
            myRev.range.Underline = wdUnderlineSingle
        End If
        DoEvents
    Next
End Sub

Private Sub AddStrikeThroughToRevisionDelete()
    Dim myRev As Revision: For Each myRev In ActiveDocument.Revisions
    With myRev
        If .Type = wdRevisionDelete Then
            .range.Font.StrikeThrough = True
            .Reject
        End If
        DoEvents
    End With
    Next
End Sub

Private Sub ClearUnderlineInBracket()
    Dim myRange: Set myRange = ActiveDocument.range()
    Call SetFindOption(myRange)
    
    With myRange.Find
        .text = "【[!【】]{2,10}】"
        
        Do While .Execute = True
            .Parent.Underline = wdUnderlineNone
            .Parent.Move
        Loop
    End With
End Sub

Private Sub SetFindOption(myRange)
    With myRange.Find
        .ClearFormatting
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .MatchFuzzy = False
    End With
End Sub
