Attribute VB_Name = "mod_TpPolo"
Option Explicit

Private Const MapaInicio As Integer = 1
Private Const XInicio As Integer = 60
Private Const YInicio As Integer = 45

Private Const MapaDestino As Integer = 283
Private Const XDestino As Integer = 45
Private Const YDestino As Integer = 49

Public Sub CrearTPaPOLO()

    Dim TELEPORT As Obj
    TELEPORT.Amount = 1
    TELEPORT.objIndex = TELEP_OBJ_INDEX

    With MapData(MapaInicio, XInicio, YInicio)
        .TileExit.map = MapaDestino
        .TileExit.X = XDestino
        .TileExit.Y = YDestino
    End With
    Call MakeObj(TELEPORT, MapaInicio, XInicio, YInicio)
End Sub

Public Sub BorrarTPaPOLO()
    With MapData(MapaInicio, XInicio, YInicio)
        Call EraseObj(.ObjInfo.Amount, MapaInicio, XInicio, YInicio)
        .TileExit.map = 0
        .TileExit.X = 0
        .TileExit.Y = 0
    End With
End Sub

