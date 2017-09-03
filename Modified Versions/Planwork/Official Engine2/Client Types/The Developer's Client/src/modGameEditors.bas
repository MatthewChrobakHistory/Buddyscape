Attribute VB_Name = "modGameEditors"
Option Explicit
' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.Max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' render the tiles
    Call EditorMap_BltTileset
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim x As Long
Dim Y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.Text = Trim$(Map.Name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(Map.Up)
        .txtDown.Text = CStr(Map.Down)
        .txtLeft.Text = CStr(Map.Left)
        .txtRight.Text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.Text = CStr(Map.BootMap)
        .txtBootX.Text = CStr(Map.BootX)
        .txtBootY.Text = CStr(Map.BootY)

        ' show the map npcs
        .lstNpcs.Clear
        For x = 1 To MAX_MAP_NPCS
            If Map.Npc(x) > 0 Then
            .lstNpcs.AddItem x & ": " & Trim$(Npc(Map.Npc(x)).Name)
            Else
                .lstNpcs.AddItem x & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For x = 1 To MAX_NPCS
            .cmbNpc.AddItem x & ": " & Trim$(Npc(x).Name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
        If Map.DropItemsOnDeath = True Then
            .chkDIOD.Value = 1
        Else
            .chkDIOD.Value = 0
        End If
        If Map.InstanceHost = True Then
            .chkInstanceHost.Value = 1
        Else
            .chkInstanceHost.Value = 0
        End If
        If Map.IsInstanced = True Then
            .chkIsInstanced.Value = 1
        Else
            .chkIsInstanced.Value = 0
        End If
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal x As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not multitile Then ' single
        With Map.Tile(x, Y)
            ' set layer
            .Layer(CurLayer).x = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For x = CurX To CurX + EditorTileWidth - 1
                If x >= 0 And x <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(x, Y)
                            .Layer(CurLayer).x = EditorTileX + x2
                            .Layer(CurLayer).Y = EditorTileY + y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal x As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = 0
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' door
                If frmEditor_Map.optDoor.Value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = DoorEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                'sprite
                If frmEditor_Map.optSprite.Value Then
                    .Type = TILE_TYPE_SPRITE
                    .Data1 = TileSprite
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                If frmEditor_Map.optRec.Value Then
                    .Type = TILE_TYPE_REC
                    .Data1 = RecType
                    .Data2 = RecIndex
                    .Data3 = RecAditional
                    .Data4 = RecCondi
                End If
                
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            x = x - ((x \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If x >= DirArrowX(i) And x <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).x = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
            End With
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, x As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = x \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmEditor_Map.shpSelected.top = EditorTileY * PIC_Y
        frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
        
        frmEditor_Map.shpSelected.width = PIC_X
        frmEditor_Map.shpSelected.height = PIC_Y
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, x As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        x = (x \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If x < 0 Then x = 0
        If x > frmEditor_Map.picBackSelect.width / PIC_X Then x = frmEditor_Map.picBackSelect.width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.height / PIC_Y Then Y = frmEditor_Map.picBackSelect.height / PIC_Y
        ' find out what to set the width + height of map editor to
        If x > EditorTileX Then ' drag right
            EditorTileWidth = x - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        frmEditor_Map.shpSelected.width = EditorTileWidth * PIC_X
        frmEditor_Map.shpSelected.height = EditorTileHeight * PIC_Y
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.width < frmEditor_Map.picBack.width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.Value * PIC_X) * -1
    End If
    
    ' vertical scrolling
    If frmEditor_Map.picBackSelect.height < frmEditor_Map.picBack.height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
        frmEditor_Map.picBackSelect.top = (frmEditor_Map.scrlPictureY.Value * PIC_Y) * -1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim x As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(x, Y).Layer(CurLayer).x = 0
                Map.Tile(x, Y).Layer(CurLayer).Y = 0
                Map.Tile(x, Y).Layer(CurLayer).Tileset = 0
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim x As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(x, Y).Layer(CurLayer).x = EditorTileX
                Map.Tile(x, Y).Layer(CurLayer).Y = EditorTileY
                Map.Tile(x, Y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim x As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then

        For x = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(x, Y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.Name)
        If .Pic > frmEditor_Item.scrlPic.Max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.Value = .Speed
            frmEditor_Item.cmbStyle.ListIndex = .CombatType
            
            If Item(EditorIndex).istwohander Then
                frmEditor_Item.ChkTwoh.Value = 1
            Else
                frmEditor_Item.ChkTwoh.Value = 0
            End If
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlMeleeOffense.Value = .MeleeOffense
            frmEditor_Item.scrlMeleeDefense.Value = .MeleeDefense
            frmEditor_Item.scrlRangedOffense.Value = .RangedOffense
            frmEditor_Item.scrlRangedDefense.Value = .RangedDefense
            frmEditor_Item.scrlMagicOffense.Value = .MagicOffense
            frmEditor_Item.scrlMagicDefense.Value = .MagicDefense
            frmEditor_Item.scrlPaperdoll1.Value = .Paperdoll1
            frmEditor_Item.scrlPaperdoll2.Value = .Paperdoll2
            frmEditor_Item.scrlPaperdoll3.Value = .Paperdoll3
            frmEditor_Item.scrlPaperdoll4.Value = .Paperdoll4
            frmEditor_Item.scrlPaperdoll5.Value = .Paperdoll5
            frmEditor_Item.scrlPaperdoll6.Value = .Paperdoll6
            frmEditor_Item.scrlWoodcutting = .WcXP
            frmEditor_Item.scrlFishing = .FXP
            frmEditor_Item.scrlMining = .MXP
            frmEditor_Item.scrlWeight.Value = .Weight
            
            
            ' projectiles
            If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
                With Item(EditorIndex).ProjecTile
                    frmEditor_Item.scrlProjectileDamage.Value = .Damage
                    frmEditor_Item.scrlProjectilePic.Value = .Pic
                    frmEditor_Item.scrlProjectileRange.Value = .Range
                    frmEditor_Item.scrlProjectileSpeed.Value = .Speed
                End With
            End If
        Else
            frmEditor_Item.fraEquipment.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
            frmEditor_Item.scrlItem.Value = .ConsumeItem
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.Value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
        frmEditor_Item.scrlSmithing.Value = .SmXP
        frmEditor_Item.scrlCooking.Value = .CoXP
        frmEditor_Item.scrlFletching.Value = .FlXP
        frmEditor_Item.scrlCrafting.Value = .CrXP
        frmEditor_Item.scrlSmithingRew.Value = .SmRew
        frmEditor_Item.scrlCookingRew.Value = .CoRew
        frmEditor_Item.scrlFletchingRew.Value = .FlRew
        frmEditor_Item.scrlCraftingRew.Value = .CrRew
        frmEditor_Item.scrlSmithingXP.Value = .EqSmXP
        frmEditor_Item.scrlCookingXP.Value = .EqCoXP
        frmEditor_Item.scrlFletchingXP.Value = .EqFlXP
        frmEditor_Item.scrlCraftingXP.Value = .EqCrXP
        frmEditor_Item.scrlPotionBrewing.Value = .EqPBXP
        frmEditor_Item.scrlPotionBrewingRew.Value = .PBRew
        frmEditor_Item.scrlPotionBrewingXP.Value = .PBXP
        
        frmEditor_Item.scrlQuestIndex.Value = .QuestIndex
        frmEditor_Item.scrlQuestType.Value = .QuestType
        If .QuestTask = 0 Then .QuestTask = 1
        frmEditor_Item.scrlQuestTask.Value = .QuestTask
        
        frmEditor_Item.scrlRQuestIndex.Value = .RQuestIndex
        frmEditor_Item.scrlRQuestType.Value = .RQuestType
        frmEditor_Item.scrlRQuestTask.Value = .RQuestTask
        frmEditor_Item.scrlCustomScript.Value = .CustomScript
    
    With frmEditor_Item
        .scrlQuestIndex.Visible = False
        .lblQuestIndex.Visible = False
        .scrlQuestTask.Visible = False
        .lblQuestTask.Visible = False
        
        Select Case .scrlQuestType.Value
            Case 0
                .lblQuestType.Caption = "Quest Type: None"
            Case 1
                .lblQuestType.Caption = "Quest Type: Start Quest"
                .scrlQuestIndex.Visible = True
                .lblQuestIndex.Visible = True
            Case 2
                .lblQuestType.Caption = "Quest Type: Advance Quest"
                .scrlQuestIndex.Visible = True
                .lblQuestIndex.Visible = True
                .scrlQuestTask.Visible = True
                .lblQuestTask.Visible = True
        End Select
    End With
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Call EditorItem_BltItem
    Call EditorItem_BltPaperdoll
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.Name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Call EditorAnim_BltAnim
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.Text = Trim$(Npc(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.Max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.Value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtChance.Text = CStr(Npc(EditorIndex).DropChance(1))
        .scrlNum.Value = Npc(EditorIndex).DropItem(1)
        .scrlValue.Value = Npc(EditorIndex).DropItemValue(1)
        .txtHP.Text = Npc(EditorIndex).HP
        .txtEXP.Text = Npc(EditorIndex).EXP
        .txtLevel.Text = Npc(EditorIndex).Level
        .txtDamage.Text = Npc(EditorIndex).Damage
        .scrlSpellNum.Max = MAX_NPC_SPELLS
        .scrlSpellNum.Value = 1
        .txtSpeed.Text = Npc(EditorIndex).Speed
        
        'Tools required
        .scrlHelmet.Value = Npc(EditorIndex).HelmetReq
        .scrlArmor.Value = Npc(EditorIndex).ArmorReq
        .scrlLegs.Value = Npc(EditorIndex).LegsReq
        .scrlShield.Value = Npc(EditorIndex).ShieldReq
        .scrlWeapon.Value = Npc(EditorIndex).WeaponReq
        
        .scrlConv.Value = Npc(EditorIndex).Conv
        .scrlFace.Value = Npc(EditorIndex).Face
        
        .scrlRQuestIndex = Npc(EditorIndex).RQuestIndex
        .scrlRQuestTask = Npc(EditorIndex).RQuestTask
        .scrlRQuestType = Npc(EditorIndex).RQuestType
        
        .scrlMeleeOffense.Value = Npc(EditorIndex).MeleeOffense
        .scrlMeleeDefense.Value = Npc(EditorIndex).MeleeDefense
        .scrlRangedOffense.Value = Npc(EditorIndex).RangedOffense
        .scrlRangedDefense.Value = Npc(EditorIndex).RangedDefense
        .scrlMagicOffense.Value = Npc(EditorIndex).MagicOffense
        .scrlMagicDefense.Value = Npc(EditorIndex).MagicDefense
        .cmbStyle.ListIndex = Npc(EditorIndex).AttackType
        
        .scrlOnAttack.Value = Npc(EditorIndex).AliveCustomScript
        .scrlOnDeath.Value = Npc(EditorIndex).DeadCustomScript
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Npc(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = Npc(EditorIndex).Stat(i)
        Next
    End With
    
    Call EditorNpc_BltSprite
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.Max = NumResources
        .scrlNormalPic.Max = NumResources
        .scrlAnimation.Max = MAX_ANIMATIONS
        
        .txtName.Text = Trim$(Resource(EditorIndex).Name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealthMax.Value = Resource(EditorIndex).health_max
        .scrlHealth.Value = Resource(EditorIndex).health_min
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .txtChance.Text = CStr(Resource(EditorIndex).DropChance(1))
        .scrlNum.Value = Resource(EditorIndex).DropItem(1)
        .scrlValue.Value = Resource(EditorIndex).DropItemValue(1)
        .scrlRewardXP.Value = Resource(EditorIndex).RewardXP
        .scrlWoodcutting.Value = Resource(EditorIndex).WcReq
        .scrlFishing.Value = Resource(EditorIndex).FReq
        .scrlMining.Value = Resource(EditorIndex).MReq
        
        .scrlQuestIndex.Value = Resource(EditorIndex).QuestIndex
        .scrlQuestTask.Value = Resource(EditorIndex).QuestTask
        .scrlQuestType.Value = Resource(EditorIndex).QuestType
        If Resource(EditorIndex).FXP Then
            frmEditor_Resource.chkFXP.Value = 1
        Else
            frmEditor_Resource.chkFXP.Value = 0
        End If
        
        If Resource(EditorIndex).WcXP Then
            frmEditor_Resource.chkWcXP.Value = 1
        Else
            frmEditor_Resource.chkWcXP.Value = 0
        End If
        
        If Resource(EditorIndex).MXP Then
            frmEditor_Resource.chkMXP.Value = 1
        Else
            frmEditor_Resource.chkMXP.Value = 0
        End If
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
        
    Call EditorResource_BltSprite
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.scrlShopType.Value = Shop(EditorIndex).ShopType
    
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"
    frmEditor_Shop.cmbCostItem2.Clear
    frmEditor_Shop.cmbCostItem2.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
        
        frmEditor_Shop.cmbCostItem2.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem2.ListIndex = 0
    frmEditor_Shop.txtNumCosts.Text = 1
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).Name) & " and " & .CostValue2 & "x " & Trim$(Item(.CostItem2).Name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next
        .cmbClass.ListIndex = 0
        
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).x
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).Dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        .ScrlTrans.Value = Spell(EditorIndex).trans
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    EditorSpell_BltIcon
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'/////////
'//DOORS//
'/////////

Public Sub DoorEditorInit()
    If frmEditor_Doors.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Doors.lstIndex.ListIndex + 1
    
    With frmEditor_Doors
    
        .txtName.Text = Doors(EditorIndex).Name
        If Doors(EditorIndex).DoorType = 0 Then
           .optDoor(0).Value = True
        Else
            .optDoor(1).Value = True
        End If
        .scrlKey.Value = Doors(EditorIndex).key
        .scrlSwitch.Value = Doors(EditorIndex).Switch
        .scrlMap.Value = Doors(EditorIndex).WarpMap
        .scrlX.Value = Doors(EditorIndex).WarpX
        .scrlY.Value = Doors(EditorIndex).WarpY
        If Doors(EditorIndex).UnlockType = 0 Then
            .OptUnlock(0).Value = True
        ElseIf Doors(EditorIndex).UnlockType = 1 Then
            .OptUnlock(1).Value = True
        Else
            .OptUnlock(2).Value = True
        End If
        
    End With
    Door_Changed(EditorIndex) = True
End Sub

Public Sub DoorEditorOk()
    Dim i As Long

    For i = 1 To MAX_DOORS
        If Door_Changed(i) Then
            Call SendSavedoor(i)
        End If
    Next
    
    Unload frmEditor_Doors
    Editor = 0
    ClearChanged_Doors
End Sub

Public Sub DoorEditorCancel()
    Editor = 0
    Unload frmEditor_Doors
    ClearChanged_Doors
    ClearDoors
    SendRequestDoors
End Sub

Public Sub ClearChanged_Doors()
    ZeroMemory Door_Changed(1), MAX_DOORS * 2 ' 2 = boolean length
End Sub


Public Sub InitEventData()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Change the data labels to match the event
    With frmEditor_Conv
        Select Case .cmbEvent.ListIndex
            Case 0 ' None
                .lblData1.Visible = False
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = False
                .scrlData2.Visible = False
                .scrlData3.Visible = False
            Case 1 ' Open Bank
                .lblData1.Visible = False
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = False
                .scrlData2.Visible = False
                .scrlData3.Visible = False
            Case 2 ' Open Shop
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the shop
                If .scrlData1.Value > MAX_SHOPS Then .scrlData1.Value = MAX_SHOPS
                .lblData1.Caption = "Shop Index: " & .scrlData1.Value
                .scrlData1.Max = MAX_SHOPS
                .scrlData1.Min = 1
            Case 3 ' Give Item
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the item
                If .scrlData1.Value > MAX_ITEMS Then .scrlData1.Value = MAX_ITEMS
                .lblData1.Caption = "Item Index: " & .scrlData1.Value
                .scrlData1.Max = MAX_ITEMS
                .scrlData1.Min = 1
                
                ' Allow them to set the amount of the item
                If .scrlData2.Value > 32767 Then .scrlData2.Value = 32767
                .lblData2.Caption = "Value: " & .scrlData2.Value
                .scrlData2.Max = 32767
                .scrlData2.Min = 1
            Case 4 ' Take Item
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the item
                If .scrlData1.Value > MAX_ITEMS Then .scrlData1.Value = MAX_ITEMS
                .lblData1.Caption = "Item Index: " & .scrlData1.Value
                .scrlData1.Max = MAX_ITEMS
                .scrlData1.Min = 1
                
                ' Allow them to set the amount of the item
                If .scrlData2.Value > 32767 Then .scrlData2.Value = 32767
                .lblData2.Caption = "Value: " & .scrlData2.Value
                .scrlData2.Max = 32767
                .scrlData2.Min = 1
            Case 5 'warp
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = True
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = True
                If .scrlData1.Value > MAX_MAPS Then .scrlData1.Value = MAX_MAPS
                If .scrlData2.Value > 250 Then .scrlData2.Value = 250
                If .scrlData3.Value > 250 Then .scrlData3.Value = 250
                
                .lblData1.Caption = "Map Index: " & .scrlData1.Value
                .lblData2.Caption = "X index: " & .scrlData2.Value
                .lblData3.Caption = "Y index: " & .scrlData3.Value
                .scrlData2.Max = 250
                .scrlData3.Max = 250
                .scrlData1.Max = MAX_MAPS
            Case 6 'start quest
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                If .scrlData1.Value > MAX_QUESTS Then .scrlData1.Value = MAX_QUESTS
                
                .lblData1.Caption = "Quest Index: " & .scrlData1.Value
                .scrlData1.Max = MAX_QUESTS
            Case 7 'advance quest
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = False
                If .scrlData1.Value > MAX_QUESTS Then .scrlData1.Value = MAX_QUESTS
                If .scrlData1.Value > MAX_QUEST_TASKS Then .scrlData1.Value = MAX_QUEST_TASKS
                
                .lblData1.Caption = "Quest Index: " & .scrlData1.Value
                .lblData2.Caption = "Finish Quest Task: " & .scrlData2.Value
                .scrlData1.Max = MAX_QUESTS
                .scrlData2.Max = MAX_QUEST_TASKS
        End Select
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitConvData", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'***********************
'* Conversation Editor *
'***********************
Public Sub ConvEditorInit(Optional ByVal ChatChange As Boolean = False)
Dim i As Long, j As Long
Dim CurChat As Byte
Dim SoundSet As Boolean

    With frmEditor_Conv
        ' Not loaded yet
        ConvEditorLoaded = False
        
        ' Set the index
        EditorIndex = frmEditor_Conv.lstIndex.ListIndex + 1
        
        ' Set the name
        .txtName = Trim$(Conv(EditorIndex).Name)
        
        ' Reset the bars
        If Conv(EditorIndex).ChatCount <= 0 Then Conv(EditorIndex).ChatCount = 1
        .scrlChatCount.Value = Conv(EditorIndex).ChatCount
        If ChatChange = False Then .scrlCurChat.Value = 1
        
        ' Lazyness
        CurChat = .scrlCurChat.Value

        ' Maximum amount of chats
        .scrlChatCount.Max = MAX_CONV_CHATS
        .fraConv.Caption = "Conversation: (" & .scrlCurChat.Value & "/" & .scrlChatCount.Value & ")"
        
        ' Populate the reply comboboxes
        For i = 1 To 4
            .cmbToConv(i).Clear
        Next
        
        ' populate the cache if we need to
        If Not hasPopulated Then
            PopulateLists
        End If
        
        ' Add the array to the combo
        .cmbSound.Clear
        .cmbSound.AddItem "None."
        For i = 1 To UBound(soundCache)
            .cmbSound.AddItem soundCache(i)
        Next
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Conv(EditorIndex).Chat(CurChat).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        .scrlCustomScript.Value = Conv(EditorIndex).Chat(CurChat).CustomScript
        
        .txtInterval.Text = Conv(EditorIndex).Chat(CurChat).SoundLength
        
        ' Add all the possible conv-to's
        For i = 1 To 4
            .cmbToConv(i).AddItem "None", 0
            For j = 1 To Conv(EditorIndex).ChatCount
                .cmbToConv(i).AddItem CStr(j), j
            Next
            
            If Conv(EditorIndex).Chat(CurChat).ReplyConvTo(i) > .scrlChatCount.Value Then
                .cmbToConv(i).ListIndex = 1
            Else
                .cmbToConv(i).ListIndex = Conv(EditorIndex).Chat(CurChat).ReplyConvTo(i)
            End If
        Next
        
        ' Set the replies
        For i = 1 To 4
            .txtReply(i).Text = Trim$(Conv(EditorIndex).Chat(CurChat).ReplyText(i))
        Next
        
        'set the quest requirements
        For i = 1 To 4
            If Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Index = 0 Then Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Index = 1
            If Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Condition = 0 Then Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Condition = 1
            
            .scrlQuestIndex(i).Value = Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Index
            .scrlTaskIndex(i).Value = Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Task
            .scrlCondi(i).Value = Conv(EditorIndex).Chat(CurChat).QuestRequirement(i).Condition
        Next
        
        ' Populate the event combobox
        .cmbEvent.Clear
        .cmbEvent.AddItem "None", 0
        .cmbEvent.AddItem "Open Bank", 1
        .cmbEvent.AddItem "Open Shop", 2
        .cmbEvent.AddItem "Give Item", 3
        .cmbEvent.AddItem "Take Item", 4
        .cmbEvent.AddItem "Warp", 5
        .cmbEvent.AddItem "Start Quest", 6
        .cmbEvent.AddItem "Advance Quest", 7
        
        ' Set it to the event
        .cmbEvent.ListIndex = Conv(EditorIndex).Chat(CurChat).Event
        
        ' Set the event data
        .scrlData1 = Conv(EditorIndex).Chat(CurChat).Data1
        .scrlData2 = Conv(EditorIndex).Chat(CurChat).Data2
        .scrlData3 = Conv(EditorIndex).Chat(CurChat).Data3
        InitEventData
        
        ' Set the text
        .txtConvText = Trim$(Conv(EditorIndex).Chat(CurChat).Text)
        
        ' Loaded
        ConvEditorLoaded = True
    End With
    
    ' Allows us to clear it if we make a mistake
    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvEditorOK()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If
    Next
    
    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ConvEditorOK", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ConvEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
    
    ClearConvs
    SendRequestConvs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ConvEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Conv()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Conv_Changed(1), MAX_CONVS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Conv", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub QuestEditorInit()
Dim i As Long
Dim CurChat As Long

frmEditor_Quest.scrlTask.Value = 1
CurChat = frmEditor_Quest.scrlTask.Value

    With frmEditor_Quest
        ' Not loaded yet
        QuestEditorLoaded = False
        
        ' Set the index
        EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
        .cmbTaskType.ListIndex = Quest(EditorIndex).Task(CurChat).TaskType
        
        'set values
        .txtName.Text = Trim$(Quest(EditorIndex).Name)
        If Quest(EditorIndex).TaskCount = 0 Then Quest(EditorIndex).TaskCount = 1
        .scrlTaskCount = Quest(EditorIndex).TaskCount
        
        .scrlReward.Max = MAX_ITEMS
        .scrlReward.Value = Quest(EditorIndex).Reward
        If Quest(EditorIndex).Reward > 0 Then
            If Item(Quest(EditorIndex).Reward).Type = ITEM_TYPE_CURRENCY Then
                .scrlAmount.Value = Quest(EditorIndex).RewardAmount
                .scrlAmount.Visible = True
                .lblAmount.Visible = True
            Else
                .scrlAmount.Visible = False
                .lblAmount.Visible = False
            End If
        Else
            .scrlAmount.Visible = False
        End If
        
        .scrlXP.Value = Quest(EditorIndex).XPReward
        If Quest(EditorIndex).XPReward > 0 Then
            .lblXP.Caption = "XP Reward: " & Quest(EditorIndex).XPReward
        Else
            .lblXP.Caption = "XP Reward: None"
        End If
        
        'curchats
        .scrlDataValue.Value = Quest(EditorIndex).Task(CurChat).DataAmount
        .scrlDataIndex.Value = Quest(EditorIndex).Task(CurChat).DataIndex
        
        ' Loaded
        QuestEditorLoaded = True
    End With
    
    ' Allows us to clear it if we make a mistake
    Quest_Changed(EditorIndex) = True
End Sub

Public Sub QuestEditorOK()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "QuestEditorOK", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub QuestEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
    ClearQuests
    SendRequestQuests
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "QuestEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Quest()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Quest", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


