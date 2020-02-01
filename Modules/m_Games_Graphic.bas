Attribute VB_Name = "m_Games_Graphic"
Option Explicit


Public Sub Game_Initialize(this_hdc As Long, this_FilePath As m_FilePath)     '初始化窗体
    Call Game_MapDataLoad(this_FilePath)
    Call Game_LoadPictures(this_hdc, this_FilePath)
    
    EventBreak = False
 
    

    
    
End Sub

Public Sub Game_LoadPictures(this_hdc As Long, this_FilePath As m_FilePath)
Dim i As Integer

    With this_Graphic
        
        .Buffer.BackBuffer = CreateCompatibleDC(this_hdc)
        .Buffer.BackBufferBmp = CreateCompatibleBitmap(this_hdc, .Screen.Width, .Screen.Height)
        .Buffer.OldBackBufferDC = SelectObject(.Buffer.BackBuffer, .Buffer.BackBufferBmp)
        
        For i = 0 To 11
        
            .Buffer.TileSetBmp(i) = CreateCompatibleDC(this_hdc)
            .Buffer.OldTilesetBmpDC(i) = SelectObject(.Buffer.TileSetBmp(i), LoadPicture(this_FilePath.Graphics & .GraphicFiles(i).FilesName))
        Next i
                
        
        
    End With
End Sub

Public Sub GameDraw(this_hdc As Long, this_Graphic As m_Graphic, this_Switch As m_Switch)
    Call Draw_Background(this_Graphic, this_Switch)
    Call Draw_Tiles0(this_Graphic, this_Switch)
    Call Draw_Tiles1(this_Graphic, this_Switch)

    BitBlt this_hdc, this_Graphic.Posi.x, this_Graphic.Posi.y, this_Graphic.Screen.Width, this_Graphic.Screen.Height, this_Graphic.Buffer.BackBuffer, 0, 0, vbSrcCopy    '将地图绘制在屏幕上

End Sub
Public Sub Draw_Background(this_Graphic As m_Graphic, m_Switch As m_Switch)    'Draw background picture
    With this_Graphic
        BitBlt .Buffer.BackBuffer, 0, 0, .Screen.Width, .Screen.Height, .Buffer.TileSetBmp(0), 0, 0, vbSrcCopy
    End With
End Sub

Public Sub Draw_Tiles0(this_Graphic As m_Graphic, m_Switch As m_Switch)   '在地图上重绘第一层图块 基础图块层
    Dim i, j As Integer

    With this_Graphic

            If .Map.Tiles(.Map.TileID).GraphicPosition.Width = 0 Then .Map.Tiles(.Map.TileID).GraphicPosition.Width = 1
            If .Map.Tiles(.Map.TileID).GraphicPosition.Height = 0 Then .Map.Tiles(.Map.TileID).GraphicPosition.Height = 1
            For i = 0 To (.Map.TilesInfo.Width / .Map.Tiles(.Map.TileID).GraphicPosition.Width) - 1
                For j = 0 To (.Map.TilesInfo.Height / .Map.Tiles(.Map.TileID).GraphicPosition.Height) - 1
                   BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.Tiles(.Map.TileID).GraphicPosition.Width, .Map.TilesInfo.y + j * .Map.Tiles(.Map.TileID).GraphicPosition.Height, .Map.Tiles(.Map.TileID).GraphicPosition.Width, .Map.Tiles(.Map.TileID).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Tiles(.Map.TileID).GraphicID), .Map.Tiles(.Map.TileID).GraphicPosition.x, .Map.Tiles(.Map.TileID).GraphicPosition.y, vbSrcCopy
                Next j
            Next i
    End With
End Sub
    
Public Sub Draw_Tiles1(this_Graphic As m_Graphic, m_Switch As m_Switch)   '在地图上重绘第一层图块 基础图块层
    Dim i, j, II, jj, a As Integer
    Dim f As Byte
    Dim s As Integer
    s = -14
    With this_Graphic
        .Player(0).Pic.x = .Player(0).Pic_Temp(0).x
        .Player(0).Pic.y = .Player(0).Pic_Temp(0).y
                                
            
            If m_Switch.Pathway = True Then
                For i = 0 To 15
                    For j = 0 To 13
                        If .Map.BlockTiles(i, j) = 1 Then
                            '调整地砖方向的
                            f = 6
                            jj = j
                            II = i
                            If jj = 0 Then jj = 1
                            If II = 0 Then II = 1
                            
                            Select Case .Map.BlockTiles(II, jj - 1)
                                Case 1
                                    Select Case .Map.BlockTiles(II - 1, jj)
                                        Case 1
                                            Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 4
                                                        Case 0
                                                            f = 7
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            If i = 0 Then
                                                                f = 3
                                                            Else
                                                                f = 5
                                                            End If
                                                        Case 0
                                                            f = 8
                                                    End Select
                                            End Select
                                        Case 0
                                              Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                            End Select
                                        End Select
                                Case 0
                                     Select Case .Map.BlockTiles(II - 1, jj)
                                        Case 1
                                            Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 1
                                                        Case 0
                                                            f = 7
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 2
                                                        Case 0
                                                            f = 8
                                                    End Select
                                            End Select
                                        Case 0
                                              Select Case .Map.BlockTiles(II + 1, jj)
                                                Case 1
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 0
                                                        Case 0
                                                            f = 6
                                                    End Select
                                                Case 0
                                                    Select Case .Map.BlockTiles(II, jj + 1)
                                                        Case 1
                                                            f = 3
                                                        Case 0
                                                            f = 6
                                                    End Select
                                            End Select
                                            
                                        End Select
                            End Select
                            
                            If j = 13 Then
                                If i = 0 Then
                                    f = 6
                                Else
                                    If .Map.BlockTiles(i, 12) = 1 Then
                                        If .Map.BlockTiles(i - 1, 12) = 0 Then
                                            If .Map.BlockTiles(i + 1, 12) = 1 Then
                                                f = 6
                                            Else
                                                f = 5
                                            End If
                                        Else
                                            If .Map.BlockTiles(i + 1, 12) = 1 Then
                                                f = 7
                                            Else
                                                f = 8
                                            End If
                                        End If
                                    Else
                                        f = 7
                                    End If
                                End If
                            End If
                            
                            If m_Switch.PathwayWithBlock = True Then
                                .Map.Block(i, j) = 1
                            Else
                                .Map.Block(i, j) = 0
                            End If
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.TilesInfo.y - 10 + j * .Map.BlockTilesInfo(f).GraphicPosition.Height, .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.BlockTilesInfo(f).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.BlockTilesInfo(f).GraphicID), .Map.BlockTilesInfo(f).GraphicPosition.x, .Map.BlockTilesInfo(f).GraphicPosition.y, vbSrcCopy
                            '自动填充地砖的下方
                            If f > 5 And f < 9 Then
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + i * .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.TilesInfo.y - 10 + (j + 1) * .Map.BlockTilesInfo(f).GraphicPosition.Height, .Map.BlockTilesInfo(f).GraphicPosition.Width, .Map.BlockTilesInfo(f + 3).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.BlockTilesInfo(f).GraphicID), .Map.BlockTilesInfo(f + 3).GraphicPosition.x, .Map.BlockTilesInfo(f + 3).GraphicPosition.y, vbSrcCopy
                            End If

                            
                            If .Player(CurrentPlayerID).Info.C_Position.x = i And .Player(CurrentPlayerID).Info.C_Position.y = j Then
                                .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                            End If
                            
                        End If


                            
                    Next j
                Next i
            End If
            
            If m_Switch.Letters = True Then
                For i = 0 To 3
                    For j = 0 To 25
                        For a = 0 To 31
                            If .Map.Letters(i, j).MapPosition(a).x = -1 And .Map.Letters(i, j).MapPosition(a).y = -1 Then
                            
                            Else
                                BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Letters(i, j).MapPosition(a).x * .Map.Tile_Object.Width, .Map.TilesInfo.y - 10 + .Map.Letters(i, j).MapPosition(a).y * .Map.Tile_Object.Height, .Map.Letters(i, j).GraphicPosition.Width, .Map.Letters(i, j).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Letters(i, j).GraphicID), .Map.Letters(i, j).GraphicPosition.x, .Map.Letters(i, j).GraphicPosition.y, vbSrcCopy
                            
                            
                                If .Player(CurrentPlayerID).Info.C_Position.x = .Map.Letters(i, j).MapPosition(a).x And .Player(CurrentPlayerID).Info.C_Position.y = .Map.Letters(i, j).MapPosition(a).y Then
                                    .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                    .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                                End If
                            End If
                        Next a
                    Next j
                Next i
                       
                For i = 0 To 51
                    For a = 0 To 31
                        If .Map.Letter(i).MapPosition(a).x = -1 And .Map.Letter(i).MapPosition(a).y = -1 Then
                        
                        Else
                            BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x + .Map.Letter(i).MapPosition(a).x * .Map.Tile_Object.Width, .Map.TilesInfo.y - 10 + .Map.Letter(i).MapPosition(a).y * .Map.Tile_Object.Height, .Map.Letter(i).GraphicPosition.Width, .Map.Letter(i).GraphicPosition.Height, .Buffer.TileSetBmp(.Map.Letter(i).GraphicID), .Map.Letter(i).GraphicPosition.x, .Map.Letter(i).GraphicPosition.y, vbSrcCopy
                        
                        
                            If .Player(CurrentPlayerID).Info.C_Position.x = .Map.Letter(i).MapPosition(a).x And .Player(CurrentPlayerID).Info.C_Position.y = .Map.Letter(i).MapPosition(a).y Then
                                .Player(CurrentPlayerID).Pic.x = .Player(CurrentPlayerID).Pic_Temp(1).x
                                .Player(CurrentPlayerID).Pic.y = .Player(CurrentPlayerID).Pic_Temp(1).y
                            End If
                        End If
                    Next a
                Next i
                
            End If
            
            
            
           If m_Switch.Steps = True Then
               
                For i = 1 To 99
                    If Steps(i).Visible = True Then
                        If Steps(i - 1).Position.y <> -1 And Steps(i - 1).Position.x <> -1 Then
                        
                            Select Case Steps(i).Direction
                                Case 0 '向下
                                    Steps(i).Position.y = Steps(i - 1).Position.y + 1
                                    Steps(i).Position.x = Steps(i - 1).Position.x
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                    If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                        s = -24
                                    Else
                                        s = -14
                                    End If
                                    End If
                                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                        BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                                Case 1 '向左
                                    Steps(i).Position.y = Steps(i - 1).Position.y
                                    Steps(i).Position.x = Steps(i - 1).Position.x - 1
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                               
                                Case 2 '向右
                            
                                    Steps(i).Position.y = Steps(i - 1).Position.y
                                    Steps(i).Position.x = Steps(i - 1).Position.x + 1
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                                
                                Case 3 ' 向上
                                    Steps(i).Position.y = Steps(i - 1).Position.y - 1
                                    Steps(i).Position.x = Steps(i - 1).Position.x
                                    If Steps(i).Position.x <= 16 And Steps(i).Position.x >= 0 And Steps(i).Position.y <= 14 And Steps(i).Position.y >= 0 Then
                                        If .Map.BlockTiles(Steps(i).Position.x, Steps(i).Position.y) = 1 Then
                                            s = -24
                                        Else
                                            s = -14
                                        End If
                                    End If
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x + 512, Steps(i).Arrow(Steps(i).Direction).y, vbSrcAnd
                                    BitBlt .Buffer.BackBuffer, .Map.TilesInfo.x - 6 + Steps(i).Position.x * .Map.Tile_Object.Width, .Map.TilesInfo.y + s + Steps(i).Position.y * .Map.Tile_Object.Height, Steps(i).Arrow(Steps(i).Direction).Width, Steps(i).Arrow(Steps(i).Direction).Height, .Buffer.TileSetBmp(Steps(i).GraphicID), Steps(i).Arrow(Steps(i).Direction).x, Steps(i).Arrow(Steps(i).Direction).y, vbSrcPaint
                            
                            End Select
                    
                        End If
                    End If
                Next i
           
           End If
    End With
    
    
    
End Sub

