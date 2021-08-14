Attribute VB_Name = "Protocol"
Private Enum ServerPacketID
    ActualizarElementosS
    InfoPetS
    HappyHourActivo
End Enum

Private Enum ClientPacketID
    Premium
    InfoPetC
    Desinvocar
    UsarElementoDaño
    UsarElementoDistancia
    ActualizarElementosC
    ExchangesTICKETS
    ExchangesKANDAHAR
    ExchangesJUTLANDS
    ExchangesANDRAMELECH
End Enum


Public Sub HandleIncomingData()

        
        Case ServerPacketID.InfoPetS
            Call HandleInfoPetS
        
        Case ServerPacketID.ActualizarElementosS
            Call HandleActualizarElementosS
            
        Case ServerPacketID.HappyHourActivo
            Call HandleHappyHourActivo
   
End Sub


Public Sub WriteInfoPet()
    Call outgoingData.WriteByte(ClientPacketID.InfoPetC)
End Sub

Private Sub HandleInfoPetS()

    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo ErrHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
   
    Call Buffer.ReadByte
   
    Dim NombreUser As String, PetLevel As Byte, PetExp As Long, PetElu As Long _
    , PetMinDMG As Integer, PetMaxDMG As Integer, PetMinHP As Integer, PetMaxHP As Integer, ElemDAÑO As Byte _
    , ElemDISTANCIA As Byte
   
    NombreUser = Buffer.ReadASCIIString()
    PetLevel = Buffer.ReadByte()
    PetExp = Buffer.ReadLong()
    PetElu = Buffer.ReadLong()
    PetMaxDMG = Buffer.ReadInteger()
    PetMinDMG = Buffer.ReadInteger()
    PetMaxHP = Buffer.ReadInteger()
    PetMinHP = Buffer.ReadInteger()
    ElemDISTANCIA = Buffer.ReadByte()
    ElemDAÑO = Buffer.ReadByte()
    
    With frmInfoPet
        .Label1.Caption = "Dueño: " & NombreUser
        .Label2.Caption = "Nivel: " & PetLevel
        If PetExp < 1 Then
            .Label3.Caption = "Exp: " & PetExp & " / " & Format$(PetElu, "##,##")
        Else
            .Label3.Caption = "Exp: " & Format$(PetExp, "##,##") & " / " & Format$(PetElu, "##,##")
        End If
        .Label4.Caption = "Daño: " & PetMinDMG + (ElemDAÑO * 25) & " / " & PetMaxDMG + (ElemDAÑO * 25)
        .Label5.Caption = "Vida: " & PetMaxHP & " / " & PetMinHP
        .Label6.Caption = "Puntos: " & ElemDISTANCIA
        .Label7.Caption = "Puntos: " & ElemDAÑO
    End With

    Call incomingData.CopyBuffer(Buffer)
   
    frmInfoPet.Show vbModeless, frmMain
   
ErrHandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0

    Set Buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub WriteActualizarElementos()
    Call outgoingData.WriteByte(ClientPacketID.ActualizarElementosC)
End Sub
Public Sub WriteUsarElementoDaño()
    Call outgoingData.WriteByte(ClientPacketID.UsarElementoDaño)
End Sub
Public Sub WriteUsarElementoDistancia()
    Call outgoingData.WriteByte(ClientPacketID.UsarElementoDistancia)
End Sub

Public Sub HandleActualizarElementosS()

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    
    PetElementoDMG = incomingData.ReadByte()
    PetElementoDIST = incomingData.ReadByte()
    PetMaximoDMG = incomingData.ReadInteger()
    PetMinimoDMG = incomingData.ReadInteger()
    
    With frmInfoPet
        .Label4.Caption = "Daño: " & PetMinimoDMG + (PetElementoDMG * 25) & " / " & PetMaximoDMG + (PetElementoDMG * 25)
        .Label6.Caption = "Puntos: " & PetElementoDIST
        .Label7.Caption = "Puntos: " & PetElementoDMG
    End With
End Sub

Public Sub WriteDesinvocar()
    Call outgoingData.WriteByte(ClientPacketID.Desinvocar)
End Sub

Public Sub WriteExchangesTICKETS(ByVal Num1 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.ExchangesTICKETS)
        Call .WriteByte(Num1)
    End With
End Sub

Public Sub WriteExchangesKANDAHAR(ByVal Num2 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.ExchangesKANDAHAR)
        Call .WriteByte(Num2)
    End With
End Sub

Public Sub WriteExchangesJUTLANDS(ByVal Num3 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.ExchangesJUTLANDS)
        Call .WriteByte(Num3)
    End With
End Sub

Public Sub WriteExchangesANDRAMELECH(ByVal Num4 As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.ExchangesANDRAMELECH)
        Call .WriteByte(Num4)
    End With
End Sub

Public Sub HandleHappyHourActivo()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Call incomingData.ReadByte
    HappyActivo = incomingData.ReadBoolean()
End Sub

Public Sub WriteVip()
    Call outgoingData.WriteByte(ClientPacketID.Premium)
End Sub
