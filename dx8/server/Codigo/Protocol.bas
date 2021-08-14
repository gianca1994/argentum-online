Attribute VB_Name = "Protocol"
Private Enum ServerPacketID


End Enum


Private Enum ClientPacketID


End Enum


Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean

        '@@@@@@@@@@ Mascotas GlaskRigAO @@@@@@@@@@
       
        Case ClientPacketID.Desinvocar
            Call HandleDesinvocar(UserIndex)
            
        Case ClientPacketID.ActualizarElementosC
            Call HandleActualizarElementosC(UserIndex)
        
        Case ClientPacketID.InfoPetC
            Call HandleInfoPet(UserIndex)

        Case ClientPacketID.UsarElementoDaño
            Call HandlePetElementoDGM(UserIndex)
            
        Case ClientPacketID.UsarElementoDistancia
            Call HandlePetElementoDISTANCIA(UserIndex)

        Case ClientPacketID.Premium
            Call HandlePremium(UserIndex)
        
        '@@@@@@@@@ FIN Mascotas GlaskRigAO @@@@@@@@@

        '@@@@@@@@@@@ CANJES GlaskRigAO @@@@@@@@@@@
       
        Case ClientPacketID.ExchangesTICKETS
            Call HandleExchangesTICKETS(UserIndex)
            
        Case ClientPacketID.ExchangesKANDAHAR
            Call HandleExchangesKANDAHAR(UserIndex)
            
        Case ClientPacketID.ExchangesJUTLANDS
            Call HandleExchangesJUTLANDS(UserIndex)
            
        Case ClientPacketID.ExchangesANDRAMELECH
            Call HandleExchangesANDRAMELECH(UserIndex)
        
        '@@@@@@@@@@ FIN CANJES GlaskRigAO @@@@@@@@@@

End Function


Public Sub WriteInfoPet(ByVal UserIndex As Integer)

On Error GoTo Errhandler
   
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.InfoPetS)
       
        Call .WriteASCIIString(UserList(UserIndex).Name)
        Call .WriteByte(UserList(UserIndex).Pet.PetLevel)
        Call .WriteLong(UserList(UserIndex).Pet.PetExp)
        Call .WriteLong(UserList(UserIndex).Pet.PetELU)
        Call .WriteInteger(UserList(UserIndex).Pet.PetDañoMaximo)
        Call .WriteInteger(UserList(UserIndex).Pet.PetDañoMinimo)
        Call .WriteInteger(UserList(UserIndex).Pet.PetMaxHP)
        Call .WriteInteger(UserList(UserIndex).Pet.PetMinHP)
        Call .WriteByte(UserList(UserIndex).Pet.PetElemDISTANCIA)
        Call .WriteByte(UserList(UserIndex).Pet.PetElemDAÑO)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleInfoPet(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
      
        Call WriteInfoPet(UserIndex)
    End With
End Sub

Private Sub HandleActualizarElementosC(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        Call WriteActualizarElementos(UserIndex)
    End With
End Sub

Private Sub WriteActualizarElementos(ByVal UserIndex As Integer)
    
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ActualizarElementosS)
        Call .WriteByte(UserList(UserIndex).Pet.PetElemDAÑO)
        Call .WriteByte(UserList(UserIndex).Pet.PetElemDISTANCIA)
        Call .WriteInteger(UserList(UserIndex).Pet.PetDañoMaximo)
        Call .WriteInteger(UserList(UserIndex).Pet.PetDañoMinimo)
    End With
Exit Sub
Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub HandlePetElementoDGM(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If .Stats.ELV < 30 Then
            Call WriteConsoleMsg(UserIndex, "Los elementos del PET pueden ser usados, a partir del nivel 30", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If TieneObjetos(411, 1, UserIndex) = False Then
            Call WriteConsoleMsg(UserIndex, "No tienes ningun elemento de daño", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        If .Pet.PetElemDAÑO >= 3 Then
            .Pet.PetElemDAÑO = .Pet.PetElemDAÑO
            Call WriteConsoleMsg(UserIndex, "Solo puedes utilizar hasta 3 elementos de daño", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        .Pet.PetElemDAÑO = .Pet.PetElemDAÑO + 1
        
        Call WriteConsoleMsg(UserIndex, "¡¡Has aumentado el daño de tu mascota!!", FontTypeNames.FONTTYPE_GUILD)
        Call QuitarObjetos(411, 1, UserIndex)
        
        WriteUpdateUserStats (UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    End With
End Sub

Public Sub HandlePetElementoDISTANCIA(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If .Stats.ELV < 30 Then
            Call WriteConsoleMsg(UserIndex, "Los elementos del PET pueden ser usados, a partir del nivel 30", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If TieneObjetos(412, 1, UserIndex) = False Then
            Call WriteConsoleMsg(UserIndex, "No tienes ningun elemento de distancia", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        If .Pet.PetElemDISTANCIA >= 4 Then
            .Pet.PetElemDISTANCIA = .Pet.PetElemDISTANCIA
            Call WriteConsoleMsg(UserIndex, "Solo puedes utilizar hasta 3 elementos de distancia", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
        End If
        
        .Pet.PetElemDISTANCIA = .Pet.PetElemDISTANCIA + 1
        
        Call WriteConsoleMsg(UserIndex, "¡¡Has aumentado el distancia de tu mascota!!", FontTypeNames.FONTTYPE_GUILD)
        Call QuitarObjetos(412, 1, UserIndex)
        
        WriteUpdateUserStats (UserIndex)
        Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    End With
End Sub

Private Sub HandleDesinvocar(ByVal UserIndex As Integer)
 
Dim i As Integer
Dim NroPets As Integer
Dim InvocadosMatados As Integer
 
Call UserList(UserIndex).incomingData.ReadByte
 
    NroPets = UserList(UserIndex).NroMascotas
    InvocadosMatados = 0
 
    With UserList(UserIndex)
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call QuitarNPC(.MascotasIndex(i))
                    .MascotasIndex(i) = 0
                    InvocadosMatados = InvocadosMatados + 1
                    NroPets = NroPets - 1
                End If
            End If
        Next i
       
        If InvocadosMatados > 0 Then _
            Call WriteConsoleMsg(UserIndex, "Has desinvocado a tu PET.", FontTypeNames.FONTTYPE_INFO)
            .NroMascotas = NroPets
    End With
End Sub

Public Sub HandleExchangesTICKETS(ByVal UserIndex As Integer)
Dim Num1 As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
            Num1 = .incomingData.ReadByte
        Call Canjes1(UserIndex, Num1)
    End With
End Sub

Public Sub HandleExchangesKANDAHAR(ByVal UserIndex As Integer)
Dim Num2 As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
            Num2 = .incomingData.ReadByte
        Call Canjes2(UserIndex, Num2)
    End With
End Sub

Public Sub HandleExchangesJUTLANDS(ByVal UserIndex As Integer)
Dim Num3 As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
            Num3 = .incomingData.ReadByte
        Call Canjes3(UserIndex, Num3)
    End With
End Sub

Public Sub HandleExchangesANDRAMELECH(ByVal UserIndex As Integer)
Dim Num4 As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
            Num4 = .incomingData.ReadByte
        Call Canjes4(UserIndex, Num4)
    End With
End Sub

Public Sub WriteHappyHourActivo(ByVal UserIndex As Integer, ByVal HappyActivo As Boolean)
    
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.HappyHourActivo)
        Call .WriteBoolean(HappyActivo)
    End With
Exit Sub
Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub HandlePremium(ByVal UserIndex As Integer)
 
    Dim MiObj As Obj
    Dim Manager As clsIniManager
    Dim Pase7dias As Integer, Pase14dias As Integer, Pase30dias As Integer
    
    Pase7dias = 1075
    Pase14dias = 1076
    Pase30dias = 1077
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
 
        If .Vip.Premium = 1 Then
            Call WriteConsoleMsg(UserIndex, "Vip> Ya eres un usuario VIP, debes esperar a que se te termine para volver a activarlo!", FontTypeNames.FONTTYPE_PREMIUM)
            Exit Sub
        End If
                            
         If TieneObjetos(Pase7dias, 1, UserIndex) = False And TieneObjetos(Pase14dias, 1, UserIndex) = False _
                                                                                                    And TieneObjetos(Pase30dias, 1, UserIndex) = False Then
            Call WriteConsoleMsg(UserIndex, "Vip> No tienes ningun pase VIP, para poder activar tu VIP", FontTypeNames.FONTTYPE_PREMIUM)
            Exit Sub
         End If
         
       If TieneObjetos(Pase7dias, 1, UserIndex) = True Then
            Call WriteConsoleMsg(UserIndex, "Vip> Te has convertido en usuario VIP!, durante 7 dias tendras grandes beneficios!", FontTypeNames.FONTTYPE_PREMIUM)
            Call QuitarObjetos(Pase7dias, 1, UserIndex)
            Call ActivateUserVIP(UserIndex, 7)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub
            
        ElseIf TieneObjetos(Pase14dias, 1, UserIndex) = True Then
            Call WriteConsoleMsg(UserIndex, "Vip> Te has convertido en usuario VIP!, durante 14 dias tendras grandes beneficios!", FontTypeNames.FONTTYPE_PREMIUM)
            Call QuitarObjetos(Pase14dias, 1, UserIndex)
            Call ActivateUserVIP(UserIndex, 14)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub

        ElseIf TieneObjetos(Pase30dias, 1, UserIndex) = True Then
            Call WriteConsoleMsg(UserIndex, "Vip> Te has convertido en usuario VIP!, durante 30 dias tendras grandes beneficios!", FontTypeNames.FONTTYPE_PREMIUM)
            Call QuitarObjetos(Pase30dias, 1, UserIndex)
            Call ActivateUserVIP(UserIndex, 30)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            
            Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
            Exit Sub
         End If
    End With
End Sub

