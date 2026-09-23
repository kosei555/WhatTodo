Attribute VB_Name = "Core"
Option Compare Database
Option Explicit
Option Base 0

Public Function DeleteInboxCurrent(db As DAO.Database, inboxId As Long)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim inboxIdParamName As String
    queryName = "DeleteInboxCurrent(p_id)"
    inboxIdParamName = "p_id"
    Set qdf = db.QueryDefs(queryName)
    qdf.Parameters(inboxIdParamName).Value = inboxId
    qdf.Execute
    qdf.Close
End Function
Public Function InsertInbox(db As DAO.Database, idea As String)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim ideaParamName As String
    queryName = "InsertInboxCurrent(p_idea)"
    ideaParamName = "p_idea"
    Set qdf = db.QueryDefs(queryName)
    qdf.Parameters(ideaParamName).Value = idea
    qdf.Execute
    qdf.Close
End Function
Public Function InsertProject(db As DAO.Database, title As String, desiredResult As String)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim titlePramName As String
    Dim desiredResultParamName As String
    queryName = "InsertProject(p_title,p_desired_result)"
    titlePramName = "p_title"
    desiredResultParamName = "p_desired_result"
    Set qdf = db.QueryDefs(queryName)
    qdf.Parameters(titlePramName).Value = title
    qdf.Parameters(desiredResultParamName).Value = desiredResult
    qdf.Execute
    qdf.Close
End Function
Public Function InsertSomeday(db As DAO.Database, title As String)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim titlePramName As String
    queryName = "InsertSomedayCurrent(p_title)"
    titlePramName = "p_title"
    Set qdf = db.QueryDefs(queryName)
    qdf.Parameters(titlePramName).Value = title
    qdf.Execute
    qdf.Close
End Function
Public Function InsertTrash(db As DAO.Database, title As String)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim titlePramName As String
    queryName = "InsertTrashCurrent(p_title)"
    titlePramName = "p_title"
    Set qdf = db.QueryDefs("InsertTrashCurrent(p_title)")
    qdf.Parameters("p_title").Value = title
    qdf.Execute
    qdf.Close
End Function
Public Function InsertAction(db As DAO.Database, todo As String)
    Dim queryName As String
    Dim qdf As QueryDef
    Dim todoPramName As String
    Dim projectIdParamName As String
    queryName = "InsertAction(p_todo,p_project_id)"
    todoPramName = "p_todo"
    projectIdParamName = "p_project_id"
    Set qdf = db.QueryDefs(queryName)
    qdf.Parameters(todoPramName).Value = todo
    qdf.Parameters(projectIdParamName).Value = Null
    qdf.Execute
    qdf.Close
End Function

