Attribute VB_Name = "mod_HappyHour"
Option Explicit

Public Sub LoadVariablesHappyHour()
    CHAPPYHOUR = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYHOUR"))
    CHAPPYAVISO = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYAVISO"))
    CHAPPYACTIVO = Val(GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYACTIVO"))
    cHAPPYHORAPREINICIO = GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYHORAPREINICIO")
    cHAPPYHORAINICIO = GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYHORAINICIO")
    cHAPPYHORAFIN = GetVar(IniPath & "ConfigGlaskRigAO.ini", "BONUSHAPPY", "cHAPPYHORAFIN")
End Sub

Public Sub MensajeHappyHour(ByVal UserIndex As Integer)

    'Dim thisDate As Date
    
   ' If thisDate = Friday Then
        If CHAPPYAVISO = 1 Then
            Select Case Time
                Case cHAPPYHORAPREINICIO
                    Call WriteConsoleMsg(UserIndex, "Servidor> El HappyHour se activara en 5 minutos!!", FontTypeNames.FONTTYPE_ORO)
                Case cHAPPYHORAINICIO
                    Call WriteConsoleMsg(UserIndex, "Servidor> Damos Comienzo al HappyHour, durante este evento la experiencia se multiplicara x" & CHAPPYHOUR & ".", FontTypeNames.FONTTYPE_ORO)
                    Call WriteHappyHourActivo(UserIndex, True)
                Case cHAPPYHORAFIN
                    Call WriteConsoleMsg(UserIndex, "Servidor> Ha terminado el HappyHour!!", FontTypeNames.FONTTYPE_ORO)
                    Call WriteHappyHourActivo(UserIndex, False)
            End Select
        End If
    'End If
End Sub

Public Sub UserLoginInHappyHour(ByVal UserIndex As Integer)
    If Time > cHAPPYHORAINICIO And Time < cHAPPYHORAFIN Then
        Call WriteHappyHourActivo(UserIndex, True)
    End If
End Sub
