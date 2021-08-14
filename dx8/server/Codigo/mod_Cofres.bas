Attribute VB_Name = "mod_Cofres"
Option Explicit

Private Type tDrops
    objIndex As Integer
    Amount As Long
    Probability As Byte
End Type

Public Const MAX_ITEM_DROPS As Byte = 5
Private Const LlaveDelCofre As Integer = 1070

Public Type e_Reward
    Drop(1 To MAX_ITEM_DROPS) As tDrops
End Type

Sub LoadItemsDeCofres(ByVal objIndex As Integer, ByRef Leer As clsIniManager)

' Agregar abajo de:
'   Case eOBJType.otForos
'       Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
' Esto:
'   Case eOBJType.otCofresMagicos
'       Call mod_Cofres.LoadItemsDeCofres(Object, Leer)

    Dim LoopC  As Long
    Dim tmpStr As String
    Dim AscII  As Integer
    
    AscII = Asc("-")
    
    With ObjData(objIndex)
        For LoopC = 1 To MAX_ITEM_DROPS
            tmpStr = Leer.GetValue("OBJ" & objIndex, "Drop" & LoopC)
            .Cofres.Drop(LoopC).objIndex = Val(ReadField(1, tmpStr, AscII))
            .Cofres.Drop(LoopC).Amount = Val(ReadField(2, tmpStr, AscII))
            .Cofres.Drop(LoopC).Probability = Val(ReadField(3, tmpStr, AscII))
        Next LoopC
    End With
End Sub

Public Sub ItemDropCofre(ByVal UserIndex As Integer, ByVal objIndex As Integer, ByVal Slot As Byte)

' Agregar abajo de:
'   If .InvSuma = 0 Then
'       If Not .flags.UltimoMensaje = 100 Then
'           .flags.UltimoMensaje = 100
'           Call WriteConsoleMsg(UserIndex, "Recuerda tener espacio suficiente en tu inventario! De lo contrario, el item caera al suelo.", FontTypeNames.FONTTYPE_INFO)
'       End If
'       .InvSuma = .InvSuma + 1
'   Exit Sub
'End If
' Esto:
'   If .InvSuma = 1 Then
'       Call ItemDropCofre(UserIndex, objIndex, Slot)
'       .InvSuma = 0
'       Exit Sub
'   End If

    Dim i As Long
    Dim MiObj As Obj
    
    If ObjData(objIndex).OBJType = 40 Then
        If TieneObjetos(LlaveDelCofre, 1, UserIndex) Then
            Call QuitarObjetos(LlaveDelCofre, 1, UserIndex)
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
            If objIndex > 0 Then
                For i = 1 To MAX_ITEM_DROPS
                    If RandomNumber(1, 100) <= ObjData(objIndex).Cofres.Drop(i).Probability Then
                        MiObj.objIndex = ObjData(objIndex).Cofres.Drop(i).objIndex
                        MiObj.Amount = ObjData(objIndex).Cofres.Drop(i).Amount
                        If Not MeterItemEnInventario(UserIndex, MiObj) Then
                            Call WriteConsoleMsg(UserIndex, "¡¡Por falta de espacio en tu inventario, el item cayo al piso!!", FontTypeNames.FONTTYPE_INFO)
                            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
                        End If
                    End If
                    Call WriteConsoleMsg(UserIndex, "Servidor> ¡Haz abierto un cofre " & ObjData(objIndex).Name & "!", FontTypeNames.FONTTYPE_GUILD)
                    Exit Sub
                Next i
            End If
         Else
            Call WriteConsoleMsg(UserIndex, "Servidor> ¡¡No tienes llave para abrir este cofre!!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Servidor> ¡No es un cofre encantado, cerrado con llave!", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
        Exit Sub
End Sub

Public Sub ChanceDeObtenerCofre(ByVal UserIndex As Integer, ByVal Num As Byte)

' Reemplazar este if:
'   If .clase = eClass.Worker Then
'       CantidadItems = MaxItemsExtraibles(.Stats.ELV)
'
' En estos 2 subs:
'   Public Sub DoPescar(ByVal UserIndex As Integer)
' Y en:
'   Public Sub DoPescar(ByVal UserIndex As Integer)
'
' Por este, en ambos subs:
'   If .clase = eClass.Worker Then
'       CantidadItems = MaxItemsExtraibles(.Stats.ELV)
'
'       ChancesCofre = RandomNumber(1, 1000)
'
'       If ChancesCofre >= 700 And ChancesCofre > 400 Then CDarCofre = 1
'       If ChancesCofre <= 400 And ChancesCofre > 200 Then CDarCofre = 2
'       If ChancesCofre <= 200 And ChancesCofre > 0 Then CDarCofre = 0
'
'       Call ChanceDeObtenerCofre(UserIndex, CDarCofre)
'
'       MiObj.Amount = RandomNumber(1, CantidadItems)
'   Else
'       MiObj.Amount = 1
'   End If


    Dim CofreComun As Integer
    Dim CofreMejorado As Integer
    Dim CofreMaximo As Integer
    Dim MiObj As Obj
    
    CofreComun = 1067
    CofreMejorado = 1068
    CofreMaximo = 1069
    
    With UserList(UserIndex)
        If .clase = eClass.Worker Then
    
            Select Case Num
                Case 0 'Cofre comun
                    MiObj.objIndex = CofreComun
                    MiObj.Amount = 1
                    
                Case 1 'Cofre Mejorado
                    MiObj.objIndex = CofreMejorado
                    MiObj.Amount = 1
                    
                Case 2 'Cofre Maximo
                    MiObj.objIndex = CofreMaximo
                    MiObj.Amount = 1
            End Select
                            
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If
                Call WriteConsoleMsg(UserIndex, "Servidor> ¡Obtuviste un cofre " & ObjData(MiObj.objIndex).Name & "!", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub


