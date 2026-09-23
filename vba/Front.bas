Attribute VB_Name = "Front"
Option Compare Database
Option Explicit
Option Base 0

Public Function SortInbox()
    Dim db As DAO.Database
    Dim inboxId As Long
    Dim title As String
    Dim ansBuf As String
    Dim actionQuestionAns As String
    Dim desiredResult As String
    Dim todo As String
    Set db = CurrentDb
    inboxId = InputBox("inboxのidを入力")
    title = InputBox("件名を入力")
    Do
        actionQuestionAns = InputBox("行動を起こすべき？(y/n)")
        If (actionQuestionAns = "n") Then
            Do
                ansBuf = InputBox("現時点では無価値で、もう必要のないもの?(y/n)")
                If (ansBuf = "y") Then
                    DeleteInboxCurrent db, inboxId
                    InsertTrash db, title
                ElseIf (ansBuf = "n") Then
                    DeleteInboxCurrent db, inboxId
                    InsertSomeday db, title
                End If
            Loop While (ansBuf <> "y" And ansBuf <> "n")
        ElseIf (actionQuestionAns = "y") Then
            Do
                ansBuf = InputBox("次にとるべき行動は一つ?(y/n)")
                If (ansBuf = "y") Then
                    DeleteInboxCurrent db, inboxId
                    todo = InputBox("次にとるべき行動を入力してください")
                    InsertAction db, todo
                ElseIf (ansBuf = "n") Then
                    DeleteInboxCurrent db, inboxId
                    desiredResult = InputBox("望んでいる結果は?")
                    InsertProject db, title, desiredResult
                    
                End If
            Loop While (ansBuf <> "y" And ansBuf <> "n")
        
        End If
    Loop While (actionQuestionAns <> "y" And actionQuestionAns <> "n")
    db.Close
End Function
Public Function AddInbox()
    Dim idea As String
    idea = InputBox("ideaを入力")
    InsertInbox CurrentDb, idea
End Function

