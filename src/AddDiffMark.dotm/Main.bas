Attribute VB_Name = "Main"
Option Explicit

' created by D*isuke YAMAKAWA (ClockAhead)

' change log :
' 2014/4/15  rev.3
'       + �n�t�����ʓ��̉����͏����悤��
'       + �p�����׏��̕␳���̏����ɍ��킹�������ǉ�
'               ref: http://ameblo.jp/gidgeerock/entry-11492099859.html
'       + ���t�@�N�^�����O�B�o�ߏ�Ԃ̕\�������폜
'
' 2012/6/27  rev.2
'
' 2011/7/29  rev.1
'       + �����̌o�ߏ�Ԃ�\������悤�Ɂ@�i�X�e�[�^�X�o�[�g�p�j
'           ��ʂ�������Ԃ̂�h�~ (���b�Z�[�W�|���v�������I�ɂ܂킷�悤��)
'       + �O�����i�ύX�������n�e�e�Ɂj�ƌ㏈���i�ύX�����̔��f�j�ǉ�

Public Sub �ύX�ӏ��ɉ���������_JP()
    ActiveDocument.TrackRevisions = False
    Call AddUnderlineToRevisionInsert
    WordBasic.AcceptAllChangesInDoc
    Call ClearUnderlineInBracket
End Sub

Public Sub �ύX�ӏ��ɉ����j��������_EN()
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
        .text = "�y[!�y�z]{2,10}�z"
        
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
