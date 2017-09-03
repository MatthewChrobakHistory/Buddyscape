Attribute VB_Name = "modQuestInfo"
Option Explicit

'This module exists because were I to add this to the real quest rec, RTE 28 (stack overflow) would occur.
'I'd be sending too much information (that's how I understand it)
'So I manually place the values here.

Public QuestInfo(1 To MAX_QUESTS) As QuestInfoRec

Private Type TaskInfoRec
    HelpInfo As String
    IntroInfo As String
End Type

Private Type QuestInfoRec
    Task(0 To MAX_QUEST_TASKS) As TaskInfoRec
    Summary As String
End Type

Public Sub SetQuestInfo()
Dim i As Long
Dim x As Long

For i = 1 To MAX_QUESTS
    QuestInfo(i).Summary = "None: If you see this, report it to a moderator. " & i

    For x = 0 To MAX_QUEST_TASKS
        QuestInfo(i).Task(x).HelpInfo = "None: If you see this, report it to a moderator. " & i & "." & x
        QuestInfo(i).Task(x).IntroInfo = "None: If you see this, report it to a moderator. " & i & "." & x
    Next
Next

' Setting structure; prepare to no-life typing
' Chances are, I won't need to type out all fifty spaces, but at least 3-4 per quest.

Exit Sub

For i = 1 To MAX_QUESTS
    With QuestInfo(i)
        Select Case i
            Case 1 'Quest 1
                .Summary = ""
                .Task(1).HelpInfo = ""
                .Task(1).IntroInfo = ""
        End Select
    End With
Next

End Sub
