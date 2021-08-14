Attribute VB_Name = "mod_Vip"
Option Explicit

Public DateFinVip As String
Public DateActual As Date
Private HoraInicioVip As String

Public Type VipStats
    Premium As Byte
    HoraInicioVip As String
    FechaFinVip As String
End Type

Public Sub CheckUserIsVIP(ByVal UserIndex As Integer)
    With UserList(UserIndex).Vip
        If Date > .FechaFinVip And Time > .HoraInicioVip Then
            'Call WriteConsoleMsg(UserIndex, "Vip> Tu tiempo como usuario VIP, ha terminado!", FontTypeNames.FONTTYPE_PREMIUM)
            .Premium = 0
            .FechaFinVip = "No es VIP"
            .HoraInicioVip = "No es VIP"
            Exit Sub
        End If
    End With
End Sub

Public Sub ActivateUserVIP(ByVal UserIndex As Integer, ByVal TiempoVip As Byte)
    
    With UserList(UserIndex)
        
        HoraInicioVip = Time
        DateFinVip = DateAdd("d", TiempoVip, Date)
        
        .Vip.Premium = 1
        .Vip.HoraInicioVip = HoraInicioVip
        .Vip.FechaFinVip = DateFinVip
        Exit Sub
    End With
End Sub

Public Sub LoadUserVIP(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    With UserList(UserIndex).Vip
        .Premium = CByte(UserFile.GetValue("VIP", "VipActivo"))
        .HoraInicioVip = CStr(UserFile.GetValue("VIP", "HoraInicioVip"))
        .FechaFinVip = CStr(UserFile.GetValue("VIP", "FechaFinVip"))
    End With
End Sub

Public Sub SaveUserVip(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)

    With UserList(UserIndex).Vip
        Call UserFile.ChangeValue("VIP", "VipActivo", CByte(.Premium))
        Call UserFile.ChangeValue("VIP", "HoraInicioVip", CStr(.HoraInicioVip))
        Call UserFile.ChangeValue("VIP", "FechaFinVip", CStr(.FechaFinVip))
    End With
End Sub


