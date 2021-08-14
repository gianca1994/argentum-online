Attribute VB_Name = "mod_Mascotas"
Option Explicit

Public Const MascotaGlask As Integer = 668

Public Type PetStats
    PetELU As Long
    PetExp As Long
    PetLevel As Long
    PetElemDAÑO As Byte
    PetElemDISTANCIA As Byte
    PetDañoMaximo As Integer
    PetDañoMinimo As Integer
    PetMinHP As Integer
    PetMaxHP As Integer
End Type

Private Const RESET_EXP As Byte = 0

Public Sub LoadVariablesInicialesPET()
    PETLVLMAX = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "NivelMaximo"))
    PETELUINICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "EluInicial"))
    PETMULTELU = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "MultiplicadorELU"))
    PETDMGMAXINICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "DañoMaxInicial"))
    PETDMGMININICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "DañoMinInicial"))
    PETMAXHPINICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "VidaMaxInicial"))
    PETMINHPINICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "VidaMinInicial"))
    PETPEGAALHITEAR = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "PegaAlHitear"))
    PETPEGAALSKILLEAR = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "PegaAlSkillear"))
    PETELEMDISTANCIAINCIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "ElemDistanciaInicial"))
    PETELEMDAÑOINICIAL = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "PET", "ElemDañoInicial"))
End Sub

Public Sub LoadUserPet(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    With UserList(UserIndex).Pet
        .PetLevel = CByte(UserFile.GetValue("PET", "PetLVL"))
        .PetExp = CLng(UserFile.GetValue("PET", "PetEXP"))
        .PetELU = CLng(UserFile.GetValue("PET", "PetELU"))
        .PetElemDAÑO = CByte(UserFile.GetValue("PET", "ElementoDMG"))
        .PetElemDISTANCIA = CByte(UserFile.GetValue("PET", "ElementoDISTANCIA"))
        .PetDañoMaximo = CInt(UserFile.GetValue("PET", "PetDañoMaximo"))
        .PetDañoMinimo = CInt(UserFile.GetValue("PET", "PetDañoMinimo"))
        .PetMinHP = CInt(UserFile.GetValue("PET", "PetMinHP"))
        .PetMaxHP = CInt(UserFile.GetValue("PET", "PetMaxHP"))
        PetDISTANCIA = .PetElemDISTANCIA
    End With
End Sub

Public Sub SavePet(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    With UserList(UserIndex).Pet
        Call UserFile.ChangeValue("PET", "PetLVL", CStr(.PetLevel))
        Call UserFile.ChangeValue("PET", "PetEXP", CStr(.PetExp))
        Call UserFile.ChangeValue("PET", "PetELU", CStr(.PetELU))
        Call UserFile.ChangeValue("PET", "ElementoDMG", CStr(.PetElemDAÑO))
        Call UserFile.ChangeValue("PET", "ElementoDISTANCIA", CStr(.PetElemDISTANCIA))
        Call UserFile.ChangeValue("PET", "PetDañoMaximo", CStr(.PetDañoMaximo))
        Call UserFile.ChangeValue("PET", "PetDañoMinimo", CStr(.PetDañoMinimo))
        Call UserFile.ChangeValue("PET", "PetMinHP", CStr(.PetMinHP))
        Call UserFile.ChangeValue("PET", "PetMaxHP", CStr(.PetMaxHP))
    End With
End Sub

Public Sub PetUpdate(ByVal MasterIndex As Long, ByVal NcpIndex As Long, Optional ByVal UpdateName As Boolean = False)
    With Npclist(NcpIndex)
        .Stats.MaxHIT = UserList(MasterIndex).Pet.PetDañoMaximo
        .Stats.MinHIT = UserList(MasterIndex).Pet.PetDañoMinimo
        
        .Stats.MaxHp = UserList(MasterIndex).Pet.PetMaxHP
        .Stats.MinHp = UserList(MasterIndex).Pet.PetMinHP
        
        '.Stats.def = UserList(MasterIndex).Pet.def
        '.PoderAtaque = UserList(MasterIndex).Pet.AP
        '.PoderEvasion = UserList(MasterIndex).Pet.Evasion
    End With
End Sub
 
Public Sub CheckPetLEVEL(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
    
        If .MascotasIndex(1) < 1 Then Exit Sub
        If .Pet.PetLevel >= PETLVLMAX Then Exit Sub
        
        If .Pet.PetExp >= .Pet.PetELU Then
        
            .Pet.PetLevel = .Pet.PetLevel + 1
            .Pet.PetExp = RESET_EXP
            .Pet.PetELU = .Pet.PetELU * PETMULTELU
            .Pet.PetMinHP = .Pet.PetMaxHP
            
            .Pet.PetDañoMinimo = .Pet.PetDañoMinimo + PETDMGMININICIAL
            .Pet.PetDañoMaximo = .Pet.PetDañoMaximo + PETDMGMAXINICIAL
            
            .Pet.PetMinHP = .Pet.PetMinHP * 1.05
            .Pet.PetMaxHP = .Pet.PetMaxHP * 1.05
            
            Call PetUpdate(UserIndex, .MascotasIndex(1))
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
            Call WriteConsoleMsg(UserIndex, "¡Tu mascota ha subido de nivel!. Nivel actual: " & .Pet.PetLevel & ".", FontTypeNames.FONTTYPE_PET)
        End If
    End With
End Sub

Public Sub CheckPETexp(ByVal UserIndex As Integer, ByVal Puntos As Long)

    With UserList(UserIndex)
        If .MascotasIndex(1) < 1 Then Exit Sub
        If .Pet.PetLevel >= PETLVLMAX Then Exit Sub
        
        .Pet.PetExp = .Pet.PetExp + CLng(Puntos)
        Call WriteConsoleMsg(UserIndex, "Tu mascota gana " & CLng(Puntos) & " puntos de experiencia.", FontTypeNames.FONTTYPE_PET)
    End With
End Sub

Public Sub PetUpdateStats(ByVal NpcIndex As Long, ByVal DañoAlPet As Long)
    With Npclist(NpcIndex)
    
        If .Numero <> MascotaGlask Then Exit Sub
        If .MaestroUser = 0 Then Exit Sub
    
        UserList(.MaestroUser).Pet.PetMinHP = UserList(.MaestroUser).Pet.PetMinHP - DañoAlPet
        Call MuerePET(.MaestroUser, NpcIndex)
    End With
End Sub

Public Sub MuerePET(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
 
    With UserList(UserIndex)

        If .MascotasIndex(1) <> NpcIndex Then Exit Sub
        If Npclist(NpcIndex).Contadores.TiempoExistencia = 0 Then Exit Sub
        
        If .Pet.PetMinHP <= 0 Then
            .Pet.PetMinHP = 0
            Call QuitarNPC(NpcIndex)
            Call WriteConsoleMsg(UserIndex, "¡Tu mascota ha muerto! Deberas revivirla con el sacerdote!", FontTypeNames.FONTTYPE_PET)
        End If
    End With
End Sub
