Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).Num = itemnum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).Npc(MapNpcNum)

    If npcNum > 0 Then
    
        MapNpc(mapnum).Npc(MapNpcNum).Num = npcNum
        MapNpc(mapnum).Npc(MapNpcNum).target = 0
        MapNpc(mapnum).Npc(MapNpcNum).targetType = 0 ' clear
        
        MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
        MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
        
        MapNpc(mapnum).Npc(MapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = MapNpcNum Then
                        MapNpc(mapnum).Npc(MapNpcNum).x = x
                        MapNpc(mapnum).Npc(MapNpcNum).y = y
                        MapNpc(mapnum).Npc(MapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(mapnum).MaxX)
                y = Random(0, Map(mapnum).MaxY)
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).Npc(MapNpcNum).x = x
                    MapNpc(mapnum).Npc(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).Npc(MapNpcNum).x = x
                        MapNpc(mapnum).Npc(MapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Num
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Dir
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).Npc(LoopI).Num > 0 Then
            If MapNpc(mapnum).Npc(LoopI).x = x Then
                If MapNpc(mapnum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).Npc(MapNpcNum).x
    y = MapNpc(mapnum).Npc(MapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(MapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(MapNpcNum).x, MapNpc(mapnum).Npc(MapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(MapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(MapNpcNum).x, MapNpc(mapnum).Npc(MapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(MapNpcNum).x - 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(MapNpcNum).x, MapNpc(mapnum).Npc(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(MapNpcNum).x + 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(MapNpcNum).x, MapNpc(mapnum).Npc(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).Npc(MapNpcNum).y = MapNpc(mapnum).Npc(MapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).Npc(MapNpcNum).y = MapNpc(mapnum).Npc(MapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).Npc(MapNpcNum).x = MapNpc(mapnum).Npc(MapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).Npc(MapNpcNum).x = MapNpc(mapnum).Npc(MapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(MapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(MapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(mapnum).DoorTimer = 0
    ReDim TempTile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            TempTile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
Dim x As Long, y As Long, Resource_Count As Long
Resource_Count = 0

For x = 0 To Map(mapnum).MaxX
         For y = 0 To Map(mapnum).MaxY

                 If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                         Resource_Count = Resource_Count + 1
                         ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                         ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                         ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                         ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = RAND(Resource(Map(mapnum).Tile(x, y).Data1).health_min, Resource(Map(mapnum).Tile(x, y).Data1).health_max)
                 End If

         Next
Next

ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
        Exit Sub
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        If Player(index).Sprite = 1 Or Player(index).Sprite = 3 Then
            Player(index).Sprite = 2
        End If
        If Player(index).Sprite = 4 Or Player(index).Sprite = 6 Then
            Player(index).Sprite = 5
        End If
        If Player(index).Sprite = 7 Or Player(index).Sprite = 9 Then
            Player(index).Sprite = 8
        End If
        If Player(index).Sprite = 10 Or Player(index).Sprite = 12 Then
            Player(index).Sprite = 11
        End If
        If Player(index).Sprite = 13 Or Player(index).Sprite = 15 Then
            Player(index).Sprite = 14
        End If
        If Player(index).Sprite = 16 Or Player(index).Sprite = 18 Then
            Player(index).Sprite = 17
        End If
        Call SendPlayerData(index)
        Exit Sub
    End If
    
    If GetPlayerEquipment(index, weapon) > 0 Then
        If GetPlayerEquipment(index, Shield) = 0 And Item(GetPlayerEquipment(index, weapon)).istwohander = False Then
            If Player(index).Sprite = 2 Or Player(index).Sprite = 3 Then
                Player(index).Sprite = 1
            End If
            If Player(index).Sprite = 5 Or Player(index).Sprite = 6 Then
                Player(index).Sprite = 4
            End If
            If Player(index).Sprite = 8 Or Player(index).Sprite = 9 Then
                Player(index).Sprite = 7
            End If
            If Player(index).Sprite = 11 Or Player(index).Sprite = 12 Then
                Player(index).Sprite = 10
            End If
            If Player(index).Sprite = 14 Or Player(index).Sprite = 15 Then
                Player(index).Sprite = 13
            End If
            If Player(index).Sprite = 17 Or Player(index).Sprite = 18 Then
                Player(index).Sprite = 16
            End If
            Call SendPlayerData(index)
            Exit Sub
        End If
    End If
    
    If GetPlayerEquipment(index, weapon) = 0 And GetPlayerEquipment(index, Shield) = 0 Then
            If Player(index).Sprite = 2 Or Player(index).Sprite = 3 Then
                Player(index).Sprite = 1
            End If
            If Player(index).Sprite = 5 Or Player(index).Sprite = 6 Then
                Player(index).Sprite = 4
            End If
            If Player(index).Sprite = 8 Or Player(index).Sprite = 9 Then
                Player(index).Sprite = 7
            End If
            If Player(index).Sprite = 11 Or Player(index).Sprite = 12 Then
                Player(index).Sprite = 10
            End If
            If Player(index).Sprite = 14 Or Player(index).Sprite = 15 Then
                Player(index).Sprite = 13
            End If
            If Player(index).Sprite = 17 Or Player(index).Sprite = 18 Then
                Player(index).Sprite = 16
            End If
            Call SendPlayerData(index)
            Exit Sub
    End If
    
End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                
                If Player(index).Map > MinMap And Player(index).Map < MaxMap Then
                    Call SimplePlayerWarp(index, 1, 6, 11)
                End If
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                    If Player(index).Map > MinMap And Player(index).Map < MaxMap Then
                        Call SimplePlayerWarp(index, 1, 6, 11)
                    End If
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partyNum).MemberCount
    leftOver = exp Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(RAND(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

' projectiles
Public Sub HandleProjecTile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, i As Long
Dim NN As Long
Dim Offense As Long
Dim Defense As Long
Dim COA As Long
Dim damage As Long


    ' check for subscript out of range
    If index < 1 Or index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetTickCount > TempPlayer(index).ProjecTile(PlayerProjectile).TravelTime Then
        With TempPlayer(index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case DIR_DOWN
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) + .Range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' up
                Case DIR_UP
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) - .Range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' right
                Case DIR_RIGHT
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(index) + .Range) + 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
                ' left
                Case DIR_LEFT
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(index) - .Range) - 1 Then ClearProjectile index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    x = TempPlayer(index).ProjecTile(PlayerProjectile).x
    y = TempPlayer(index).ProjecTile(PlayerProjectile).y
    
    ' check if left map
    If x > Map(Player(index).Map).MaxX Or y > Map(Player(index).Map).MaxY Or x < 0 Or y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if hit player
    For i = 1 To Player_HighIndex
        ' make sure they're actually playing
        If IsPlaying(i) Then
            ' check coordinates
            If x = Player(i).x And y = GetPlayerY(i) Then
                ' make sure it's not the attacker
                If Not x = Player(index).x Or Not y = GetPlayerY(index) Then
                    ' check if player can attack
                    If CanPlayerAttackPlayer(index, i, False, True) = True Then
                        ' attack the player and kill the project tile
                        
                        'based on agility
                        If CanPlayerDodge(i) Then
                            SendActionMsg GetPlayerMap(i), "Dodge!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                            Exit Sub
                        End If
                        
                        'based on strength
                        If CanPlayerParry(i) Then
                            SendActionMsg GetPlayerMap(i), "Parry!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                            Exit Sub
                        End If
                        
                        If GetPlayerEquipment(index, weapon) > 0 Then
                            Select Case Item(GetPlayerEquipment(index, weapon)).CombatType
                                Case 0
                                    Offense = GetMeleeOffense(index)
                                    Defense = GetMeleeDefense(i)
                                Case 1
                                    Offense = GetRangedOffense(index)
                                    Defense = GetRangedDefense(i)
                                Case 2
                                    Offense = GetMagicOffense(index)
                                    Defense = GetMagicDefense(index)
                            End Select
                        Else
                            Offense = 0
                            Defense = GetMeleeDefense(i)
                        End If
                        
                        COA = (Player(index).Stat(Stats.Attack) / 4) + (Offense / 6.5)
                        Defense = (Player(i).Stat(Stats.Defense) / 4) + (Defense / 6.5)
                        
                        COA = RAND(1, COA)
                        Defense = RAND(1, Defense)
                        
                        If COA < Defense Then
                            Call PlayerMsg(index, "Your attack was blocked!", BrightRed)
                            Exit Sub
                        End If
                        
                        damage = GetPlayerDamage(index)
                        damage = RAND(1, damage)
                        
                        If CanPlayerCrit(index) Then
                            damage = damage * 1.5
                            SendActionMsg GetPlayerMap(index), "Critial!", BrightCyan, 1, (GetPlayerX(index) * 32), GetPlayerY(index)
                        End If
                        
                        'EOC
                        If damage > 0 Then
                            PlayerAttackPlayer index, i, damage
                        End If
                        'EOC
                        
                        
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    Else
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If x = MapNpc(GetPlayerMap(index)).Npc(i).x And y = MapNpc(GetPlayerMap(index)).Npc(i).y Then
            NN = MapNpc(GetPlayerMap(index)).Npc(i).Num
            If NN = 0 Then Exit For
            
                ' they're hit, remove it and deal that damage
            If Npc(NN).HelmetReq > 0 Then
                If GetPlayerEquipment(index, Helmet) > 0 Then
                    If Item(GetPlayerEquipment(index, Helmet)).Data3 <> Npc(index).HelmetReq Then
                        Call PlayerMsg(index, "You should use a more appropriate helmet before attacking this creature.", BrightRed)
                        Exit For
                    End If
                Else
                    Call PlayerMsg(index, "You should use a specific helmet before attacking this creature.", BrightRed)
                    Exit For
                End If
            End If
    
            If Npc(NN).ArmorReq > 0 Then
                If GetPlayerEquipment(index, Armor) > 0 Then
                    If Item(GetPlayerEquipment(index, Armor)).Data3 <> Npc(index).ArmorReq Then
                        Call PlayerMsg(index, "You should use more appropriate body armor before attacking this creature.", BrightRed)
                        Exit For
                    End If
                Else
                    Call PlayerMsg(index, "You should use a more appropriate body armor before attacking this creature.", BrightRed)
                    Exit For
                End If
            End If
            
            If Npc(NN).LegsReq > 0 Then
                If GetPlayerEquipment(index, Legs) > 0 Then
                    If Item(GetPlayerEquipment(index, Legs)).Data3 <> Npc(index).LegsReq Then
                        Call PlayerMsg(index, "You should use a more appropriate leg armor before attacking this creature.", BrightRed)
                        Exit For
                    End If
                Else
                    Call PlayerMsg(index, "You should use a more appropriate leg armor before attacking this creature.", BrightRed)
                    Exit For
                End If
            End If

        If Npc(NN).ShieldReq > 0 Then
            If GetPlayerEquipment(index, Shield) > 0 Then
                If Item(GetPlayerEquipment(index, Shield)).Data3 <> Npc(index).ShieldReq Then
                    Call PlayerMsg(index, "You should use a more appropriate shield before attacking this creature.", BrightRed)
                    Exit For
                End If
            Else
                Call PlayerMsg(index, "You should use a more appropriate shield before attacking this creature.", BrightRed)
                Exit For
            End If
        End If

        If Npc(NN).WeaponReq > 0 Then
            If GetPlayerEquipment(index, weapon) > 0 Then
                If Item(GetPlayerEquipment(index, weapon)).Data3 <> Npc(index).WeaponReq Then
                    Call PlayerMsg(index, "You should use a more appropriate weapon before attacking this creature.", BrightRed)
                    Exit For
                End If
            Else
                Call PlayerMsg(index, "You should use a more appropriate weapon before attacking this creature.", BrightRed)
                Exit For
            End If
        End If
        
        If Npc(NN).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or Npc(index).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then Exit Sub
                        
        If GetPlayerEquipment(index, weapon) > 0 Then
            Select Case Item(GetPlayerEquipment(index, weapon)).CombatType
                Case 0
                    Offense = GetMeleeOffense(index)
                    Defense = Npc(NN).MeleeDefense
                Case 1
                    Offense = GetRangedOffense(index)
                    Defense = Npc(NN).RangedDefense
                Case 2
                    Offense = GetMagicOffense(index)
                    Defense = Npc(NN).MagicDefense
            End Select
        Else
            Offense = 0
            Defense = Npc(NN).MeleeDefense
        End If
                        
        COA = (Player(index).Stat(Stats.Attack) / 4) + (Offense / 6.5)
        Defense = (Npc(NN).Stat(Stats.Defense) / 4) + (Defense / 6.5)
                        
        COA = RAND(1, COA)
        Defense = RAND(1, Defense)
        
        damage = 1
        If COA < Defense Then
            damage = 0
        End If
                   
        If damage > 0 Then
            damage = GetPlayerDamage(index)
            damage = RAND(1, damage)
            If CanPlayerCrit(index) Then
                damage = damage * 1.5
                SendActionMsg GetPlayerMap(index), "Critial!", BrightCyan, 1, (GetPlayerX(index) * 32), GetPlayerY(index) * 32
            End If
        End If
                        
        'EOC
        
        PlayerAttackNpc index, i, damage
        If damage > 0 Then
            If GetPlayerEquipment(index, weapon) > 0 Then
                If Item(GetPlayerEquipment(index, weapon)).Animation > 0 Then
                    Call SendAnimation(GetPlayerMap(index), Item(GetPlayerEquipment(index, weapon)).Animation, MapNpc(GetPlayerMap(index)).Npc(i).x, MapNpc(GetPlayerMap(index)).Npc(i).y)
                End If
            End If
        End If
        
        'EOC
            
            ClearProjectile index, PlayerProjectile
            Exit Sub
        End If
    Next
    
    ' hit a block
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        ' hit a block, clear it.
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
End Sub

Public Sub CheckFinishQuest(ByVal index As Long, ByVal QuestNum As Long, ByVal TaskOn As Long)
    ' if it's over the number of tasks we've finished it
    If Player(index).PlayerQuest(QuestNum).TaskOn > Quest(QuestNum).TaskCount Then
        Player(index).PlayerQuest(QuestNum).QuestStatus = 2
        
        'quest rewards
        If Quest(QuestNum).Reward > 0 Then
            If FindOpenInvSlot(index, Quest(QuestNum).Reward) > 0 Then
                Call GiveInvItem(index, Quest(QuestNum).Reward, Quest(QuestNum).RewardAmount)
            Else
                Call SpawnItem(Quest(QuestNum).Reward, Quest(QuestNum).RewardAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMsg(index, "Your reward fell to the floor.", Cyan)
            End If
        End If
        
        If Quest(QuestNum).XPReward > 0 Then
            Call GivePlayerEXP(index, Quest(QuestNum).XPReward * XPRate)
        End If
        
        Call PlayerMsg(index, "Congratulations! You completed '" & Trim$(Quest(QuestNum).Name) & "'!", BrightRed)
    End If
End Sub

Public Sub StartQuest(ByVal index As Long, ByVal QuestIndex As Long)
    ' make sure we've got a quest
    If QuestIndex = 0 Then Exit Sub
    If Player(index).PlayerQuest(QuestIndex).QuestStatus = 0 Then
        Player(index).PlayerQuest(QuestIndex).QuestStatus = 1
        Player(index).PlayerQuest(QuestIndex).TaskOn = 1
        Player(index).PlayerQuest(QuestIndex).DataAmountLeft = Quest(QuestIndex).Task(1).DataAmount
        Call SendPlayerData(index)
        Call PlayerMsg(index, "You have accepted '" & Trim$(Quest(QuestIndex).Name) & "'.", Cyan)
    Else
        Call PlayerMsg(index, "You are already on this quest!", BrightRed)
        Call SendCloseConv(index)
    End If
End Sub

Public Sub AdvanceQuest(ByVal index As Long, ByVal QuestIndex As Long, ByVal TNC) ' TNC = Task needed to complete
Dim CurTaskOn As Long

CurTaskOn = Player(index).PlayerQuest(QuestIndex).TaskOn

    ' make sure they're on the quest
    If Player(index).PlayerQuest(QuestIndex).QuestStatus = 1 Then
        ' make sure that the task is the task we need to complete
        If CurTaskOn = TNC Then
        
            ' if so, go to the next task and check to see if we finished the quest
            Call QuestScript(index, QuestIndex)
            Player(index).PlayerQuest(QuestIndex).TaskOn = CurTaskOn + 1
            CurTaskOn = CurTaskOn + 1
            
            'find out if there's dataamount to be set
            If Quest(QuestIndex).Task(CurTaskOn).TaskType = 4 Or Quest(QuestIndex).Task(CurTaskOn).TaskType = 2 Then
                Player(index).PlayerQuest(QuestIndex).DataAmountLeft = Quest(QuestIndex).Task(CurTaskOn).DataAmount
            End If
            
            Call PlayerMsg(index, "You have completed a task in '" & Trim$(Quest(QuestIndex).Name) & "'!", BrightGreen)
            Call CheckFinishQuest(index, QuestIndex, CurTaskOn)
            Call SendPlayerData(index)
            
        ElseIf CurTaskOn > TNC Then
            Call PlayerMsg(index, "You have already completed this task!", BrightRed)
        End If
    Else
        If Player(index).PlayerQuest(QuestIndex).QuestStatus = 0 Then Call PlayerMsg(index, "You haven't started this quest yet.", BrightRed)
        If Player(index).PlayerQuest(QuestIndex).QuestStatus = 2 Then Call PlayerMsg(index, "You have already finished this quest.", BrightRed)
        Call SendCloseConv(index)
    End If
    
End Sub

Public Sub InstanceMap(ByVal index As Long, ByVal MapIndex As Long, ByVal MapX As Long, ByVal MapY As Long)
Dim x As Long
Dim y As Long
Dim i As Long
Dim d As Long
Dim InstMap As Long

InstMap = 0

Player(index).LNIM = MapIndex

'find an instanced map with nobody on it.
For i = 1 To MAX_MAPS
    If Map(i).InstanceHost = True Then
        If GetTotalMapPlayers(i) = 0 Then
            InstMap = i
        End If
    End If
Next

'if we couldn't find an instanced map, warp the player to the original map
If InstMap = 0 Then
    Call SimplePlayerWarp(index, MapIndex, MapX, MapY)
    Exit Sub
End If

'we found an instanced map, so copy the data
Map(InstMap).MaxX = Map(MapIndex).MaxX
Map(InstMap).MaxY = Map(MapIndex).MaxY
Map(InstMap).BootMap = Map(MapIndex).BootMap
Map(InstMap).DropItemsOnDeath = Map(MapIndex).DropItemsOnDeath
Map(InstMap).Moral = Map(MapIndex).Moral
Map(InstMap).Music = Map(MapIndex).Music
Map(InstMap).Name = Trim$(Map(MapIndex).Name)
For i = 1 To MAX_MAP_NPCS
    Map(InstMap).Npc(i) = Map(MapIndex).Npc(i)
Next
Map(InstMap).Revision = Map(MapIndex).Revision
Map(InstMap).Down = Map(MapIndex).Down
Map(InstMap).Up = Map(MapIndex).Up
Map(InstMap).Left = Map(MapIndex).Left
Map(InstMap).Right = Map(MapIndex).Right

For x = 0 To Map(MapIndex).MaxX
    For y = 0 To Map(MapIndex).MaxY
        For d = 1 To MapLayer.Layer_Count - 1
            Map(InstMap).Tile(x, y).Layer(d).Tileset = Map(MapIndex).Tile(x, y).Layer(d).Tileset
            Map(InstMap).Tile(x, y).Layer(d).y = Map(MapIndex).Tile(x, y).Layer(d).y
            Map(InstMap).Tile(x, y).Layer(d).x = Map(MapIndex).Tile(x, y).Layer(d).x
        Next
        Map(InstMap).Tile(x, y).Data1 = Map(MapIndex).Tile(x, y).Data1
        Map(InstMap).Tile(x, y).Data2 = Map(MapIndex).Tile(x, y).Data2
        Map(InstMap).Tile(x, y).Data3 = Map(MapIndex).Tile(x, y).Data3
        Map(InstMap).Tile(x, y).Data4 = Map(MapIndex).Tile(x, y).Data4
        Map(InstMap).Tile(x, y).Type = Map(MapIndex).Tile(x, y).Type
        Map(InstMap).Tile(x, y).DirBlock = Map(MapIndex).Tile(x, y).DirBlock
    Next
Next

'respawn the map

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(x, InstMap)
    Next
    Call SendMapNpcsToMap(InstMap)
    Call SpawnMapNpcs(InstMap)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, InstMap, MapItem(InstMap, i).x, MapItem(InstMap, i).y)
        Call ClearMapItem(i, InstMap)
    Next
    ' Respawn
    Call SpawnMapItems(InstMap)
    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, InstMap)
        For x = 1 To 20
            MapNpc(GetPlayerMap(index)).Npc(i).ScriptValues.ValueBoolean(x) = False
            MapNpc(GetPlayerMap(index)).Npc(i).ScriptValues.ValueLong(x) = 0
        Next
    Next
    
        ' Save the map
    Call SaveMap(InstMap)
    Call MapCache_Create(InstMap)
    Call ClearTempTile(InstMap)
    Call CacheResources(InstMap)

Call SimplePlayerWarp(index, InstMap, MapX, MapY)

End Sub

Public Sub MakeDungeon()
Dim InstMap As Long
Dim MapIndex As Long
Dim i As Long
Dim y As Long
Dim x As Long
Dim d As Long
MapIndex = 100

For InstMap = 101 To 200
    'we found an instanced map, so copy the data
    Map(InstMap).MaxX = Map(MapIndex).MaxX
    Map(InstMap).MaxY = Map(MapIndex).MaxY
    Map(InstMap).BootMap = Map(MapIndex).BootMap
    Map(InstMap).DropItemsOnDeath = Map(MapIndex).DropItemsOnDeath
    Map(InstMap).Moral = Map(MapIndex).Moral
    Map(InstMap).Music = Map(MapIndex).Music
    Map(InstMap).Name = Trim$(Map(MapIndex).Name)
    For i = 1 To MAX_MAP_NPCS
        Map(InstMap).Npc(i) = Map(MapIndex).Npc(i)
    Next
    Map(InstMap).Revision = Map(MapIndex).Revision
    Map(InstMap).Down = Map(MapIndex).Down
    Map(InstMap).Up = Map(MapIndex).Up
    Map(InstMap).Left = Map(MapIndex).Left
    Map(InstMap).Right = Map(MapIndex).Right
    
    For x = 0 To Map(MapIndex).MaxX
        For y = 0 To Map(MapIndex).MaxY
            For d = 1 To MapLayer.Layer_Count - 1
                Map(InstMap).Tile(x, y).Layer(d).Tileset = Map(MapIndex).Tile(x, y).Layer(d).Tileset
                Map(InstMap).Tile(x, y).Layer(d).y = Map(MapIndex).Tile(x, y).Layer(d).y
                Map(InstMap).Tile(x, y).Layer(d).x = Map(MapIndex).Tile(x, y).Layer(d).x
            Next
            Map(InstMap).Tile(x, y).Data1 = Map(MapIndex).Tile(x, y).Data1
            Map(InstMap).Tile(x, y).Data2 = Map(MapIndex).Tile(x, y).Data2
            Map(InstMap).Tile(x, y).Data3 = Map(MapIndex).Tile(x, y).Data3
            Map(InstMap).Tile(x, y).Data4 = Map(MapIndex).Tile(x, y).Data4
            Map(InstMap).Tile(x, y).Type = Map(MapIndex).Tile(x, y).Type
            Map(InstMap).Tile(x, y).DirBlock = Map(MapIndex).Tile(x, y).DirBlock
        Next
    Next
    
    'respawn the map
    
        For i = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, InstMap)
        Next
        Call SendMapNpcsToMap(InstMap)
        Call SpawnMapNpcs(InstMap)
    
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, InstMap, MapItem(InstMap, i).x, MapItem(InstMap, i).y)
            Call ClearMapItem(i, InstMap)
        Next
        ' Respawn
        Call SpawnMapItems(InstMap)
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, InstMap)
            For x = 1 To 20
                MapNpc(InstMap).Npc(i).ScriptValues.ValueBoolean(x) = False
                MapNpc(InstMap).Npc(i).ScriptValues.ValueLong(x) = 0
            Next
        Next
        
            ' Save the map
        Call SaveMap(InstMap)
        Call MapCache_Create(InstMap)
        Call ClearTempTile(InstMap)
        Call CacheResources(InstMap)
Next
End Sub
