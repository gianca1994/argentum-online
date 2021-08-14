Attribute VB_Name = "mod_Castillo"
Option Explicit

Public CastilloGlask As String
Public MapaCastilloGlask As Integer
Public RegaloCastillo As Integer
Public CantidadRegalo As Integer
Public CastilloActivo As Boolean

Public Sub CargaVariablesCastillo()
    CastilloGlask = GetVar(DatPath & "CastilloGlask.ini", "CASTILLO", "Castillo")
    MapaCastilloGlask = GetVar(DatPath & "CastilloGlask.ini", "MAPACASTILLO", "MapCastillo")
    RegaloCastillo = GetVar(DatPath & "CastilloGlask.ini", "PREMIO", "ItemPremio")
    CantidadRegalo = GetVar(DatPath & "CastilloGlask.ini", "PREMIO", "CantidadPremio")
End Sub

Public Sub RecargarDatosCastillo()
    Call WriteVar(DatPath & "CastilloGlask.ini", "CASTILLO", "Castillo", "Vacante")
    CastilloGlask = GetVar(DatPath & "CastilloGlask.ini", "CASTILLO", "Castillo")
End Sub

Public Sub ClickeamosAlRey(ByVal TempCharIndex As Integer, ByVal UserIndex As Integer)

    With Npclist(TempCharIndex)
        If .Pos.map = MapaCastilloGlask And .NPCtype = ReyCastillo Then
            If CastilloGlask = "Vacante" Then
                Call WriteChatOverHead(UserIndex, "¡Maldito bastardo, alejate de mi castillo!", Str(.Char.CharIndex), vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Estoy al servicio del clan: " & CastilloGlask & "", Str(.Char.CharIndex), vbWhite)
            End If
        End If
    End With
End Sub

Public Sub MensajeAlAtacarREY(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    With Npclist(NpcIndex)
        If .Pos.map = MapaCastilloGlask And .NPCtype = ReyCastillo And .Stats.MinHp > 750 And .Stats.MinHp <> 1500 Then
            If RandomNumber(1, 100) <= 60 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Rey del castillo esta siendo atacado por el clan " & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            End If
        ElseIf .Pos.map = MapaCastilloGlask And .NPCtype = ReyCastillo And .Stats.MinHp > 0 And .Stats.MinHp < 500 Then
            If RandomNumber(1, 100) <= 25 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Rey del castillo esta  a punto de caer en las manos del clan " & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
    End With
End Sub

Public Sub MuereRey(ByVal UserIndex As Integer, NpcIndex As Integer)
      Dim castillo As Integer
      Dim Fundador As String
      castillo = 0

      With UserList(UserIndex)
            If .Pos.map = MapaCastilloGlask Then castillo = 1
            If castillo = 0 Then Exit Sub
            
            If castillo = 1 Then
                CastilloGlask = modGuilds.GuildName(.GuildIndex)
                Call WriteVar(DatPath & "CastilloGlask.ini", "CASTILLO", "Castillo", modGuilds.GuildName(.GuildIndex))
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El clan " & (modGuilds.GuildName(.GuildIndex)) & " ha conquistado el Castillo y como recompensa obtienen 50 Puntos de Clan!!", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToAll, UserIndex, PrepareMessagePlayWave(44, .Pos.X, .Pos.Y))
            End If
               
            ' Le otorgamos puntos de clan, al clan que mate al rey
            Call guilds(.GuildIndex).ManejoPuntosClan(50)
            .Stats.AztecPiece = .Stats.AztecPiece + 100
            Call WriteUpdateGold(UserIndex)
            Call QuitarNPC(NpcIndex)
            Call WriteConsoleMsg(UserIndex, "¡¡Has matado al rey!!", FontTypeNames.FONTTYPE_GUILD)
            CastilloActivo = True
            
            If guilds(UserList(UserIndex).GuildIndex).GuildName = CastilloGlask Then
                Call WriteClanConquistoColor(UserIndex, True)
            End If
      End With
End Sub

Public Sub ReviveRey(ByVal NpcIndex As Integer)

Dim reNpcPos As WorldPos, reNpcIndex As Integer, ReyNpcN As Integer

    reNpcPos.map = MapaCastilloGlask
    reNpcPos.X = 50
    reNpcPos.Y = 50
    reNpcIndex = NpcIndex
    ReyNpcN = 667
 
    Call SpawnNpc(ReyNpcN, reNpcPos, True, True)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡El Rey del castillo ha vuelto al poder! Ve, acaba con el y sus secuaces y reclama el trono!", FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub CastilloPasarTiempo(ByVal PasarTiempo As Long)

Dim NpcIndex As Integer
Dim LoopC As Long

    Select Case PasarTiempo
        Case 1
            Call DarPremioCastillo
        Case 2
            Call DarPremioCastillo
        Case 3
            Call DarPremioCastillo
            
                For LoopC = 1 To LastUser
                    If guilds(UserList(LoopC).GuildIndex).GuildName = CastilloGlask Then
                        Call WriteClanConquistoColor(LoopC, False)
                        Call WriteUpdateUserStats(LoopC)
                    End If
                Next LoopC
            
            Call RecargarDatosCastillo
            Call ReviveRey(NpcIndex)
            CastilloActivo = False
            Exit Sub
    End Select
End Sub

Public Sub DarPremioCastillo()

On Error GoTo handler

Dim PremioCastillo As Obj
Dim LoopC As Integer

PremioCastillo.objIndex = RegaloCastillo
PremioCastillo.Amount = CantidadRegalo
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).GuildIndex <> 0 Then
            If guilds(UserList(LoopC).GuildIndex).GuildName = CastilloGlask Then
                If UserList(LoopC).flags.UserLogged Then
                    Call MeterItemEnInventario(LoopC, PremioCastillo)
                    Call WriteConsoleMsg(LoopC, "Servidor> Han recibido " & PremioCastillo.Amount & " " & ObjData(PremioCastillo.objIndex).Name & ", por mantener el Castillo en su poder!!", FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        End If
    Next LoopC
    Exit Sub
    
handler:
    Call LogError("Error en DarPremioCastillos.")
End Sub

