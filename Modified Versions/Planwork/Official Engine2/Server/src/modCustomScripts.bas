Attribute VB_Name = "modCustomScripts"
Option Explicit

Public Sub ItemScript(ByVal index As Long, ByVal itemnum As Long)
Dim NewItem As Long
Select Case Item(itemnum).CustomScript
    Case 0
        Exit Sub 'just to avoid possible errors.
    Case 1 'random item
        NewItem = RAND(2, 274)
        Call GiveInvItem(index, NewItem, 1, True)
End Select

End Sub

Public Sub AliveNpcScript(ByVal index As Long, ByVal npcNum As Long, ByVal MapNpcNum As Long)

With MapNpc(GetPlayerMap(index)).Npc(MapNpcNum)
    Select Case Npc(npcNum).AliveCustomScript
        Case 0
            Exit Sub 'just to avoid possible errors.
        Case 1 'example case
            If .ScriptValues.ValueBoolean(1) = False Then
                If .Vital(Vitals.HP) < Npc(npcNum).HP * 0.9 Then
                    Call MapMsg(GetPlayerMap(index), Trim$(Npc(npcNum).Name) & ": You think you can run, mortal?", Cyan)
                    .ScriptValues.ValueBoolean(1) = True
                    Exit Sub
                End If
            End If
            
            If .ScriptValues.ValueBoolean(2) = False Then
                If .Vital(Vitals.HP) < Npc(npcNum).HP * 0.7 Then
                    Call MapMsg(GetPlayerMap(index), Trim$(Npc(npcNum).Name) & ": You cannot run.", Cyan)
                    .ScriptValues.ValueBoolean(2) = True
                    Exit Sub
                End If
            End If
            
            If .ScriptValues.ValueBoolean(3) = False Then
                If .Vital(Vitals.HP) < Npc(npcNum).HP * 0.5 Then
                    .ScriptValues.ValueBoolean(3) = True
                    Exit Sub
                End If
            End If
            
            If .ScriptValues.ValueBoolean(4) = False Then
                If .Vital(Vitals.HP) < Npc(npcNum).HP * 0.2 Then
                    Call MapMsg(GetPlayerMap(index), Trim$(Npc(npcNum).Name) & ": You're pathetic. You know that?", Cyan)
                    .ScriptValues.ValueBoolean(4) = True
                    Exit Sub
                End If
            End If
    End Select
End With

End Sub

Public Sub DeadNpcScript(ByVal index As Long, ByVal npcNum As Long, ByVal MapNpcNum As Long)

With MapNpc(GetPlayerMap(index)).Npc(MapNpcNum).ScriptValues
    Select Case Npc(npcNum).DeadCustomScript
        Case 0
            Exit Sub 'just to avoid possible errors.
        Case 1
            Call InstanceMap(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call SendStartConversation(index, 1, npcNum)
    End Select
End With

End Sub

Public Sub AliveResourceScript(ByVal index As Long, ByVal resource_index As Long, ByVal rX As Long, ByVal rY As Long)

End Sub

Public Sub DeadResourceScript(ByVal index As Long, ByVal resource_index As Long, ByVal rX As Long, ByVal rY As Long)

End Sub

Public Sub QuestScript(ByVal index As Long, ByVal QuestNum As Long)

End Sub

Public Sub ConvScript(ByVal index As Long, ByVal ConIndex As Long, ByVal PageOn As Long)

End Sub
