Attribute VB_Name = "mod_AutoRegalos"
Option Explicit

Public ActivatedAutoRegalos As Boolean

Public Sub EventoAutoRegalos(ByVal Time As Long)

' Agregar en:
' Private Sub AutoSave_Timer()
' Esto:
' If ActivatedAutoRegalos Then
'     Call EventoAutoRegalos(AutoRegalosTick)
' End If

Dim UserIndexAlAzar As Integer
Dim Tiempo1 As Byte, Tiempo2 As Byte, Tiempo3 As Byte

    UserIndexAlAzar = RandomNumber(1, LastUser)

    Tiempo1 = RandomNumber(1, 180)
    Tiempo2 = RandomNumber(240, 360)
    Tiempo3 = RandomNumber(420, 540)
    
    Select Case Time
        Case Tiempo1
            Call DarRegaloAutomatico(UserIndexAlAzar)
        Case Tiempo2
            Call DarRegaloAutomatico(UserIndexAlAzar)
        Case Tiempo3
            Call DarRegaloAutomatico(UserIndexAlAzar)
        Case Else
            ActivatedAutoRegalos = False
            Exit Sub
    End Select
End Sub

Public Sub DarRegaloAutomatico(ByVal UserIndexAlAzar As Integer)

    Dim ItemRegalo As Obj
    Dim Regalo As Byte
    Dim Probabilidad As Integer
    Dim N As Integer
    
    N = FreeFile
    Open App.Path & "\logs\AutoRegalos.log" For Append Shared As #N

    Regalo = RandomNumber(1, 2)
    Probabilidad = RandomNumber(1, 1000)
    
    If UserList(UserIndexAlAzar).flags.UserLogged Then
        With ItemRegalo
            If Probabilidad >= 500 Then
                Select Case Regalo
                    Case 1
                        .objIndex = 541
                        .Amount = 1
                    Case Else
                        .objIndex = 541
                        .Amount = 2
                End Select
                
            ElseIf Probabilidad < 500 And Probabilidad >= 200 Then
                Select Case Regalo
                    Case 1
                        .objIndex = 541
                        .Amount = 4
                    Case Else
                        .objIndex = 541
                        .Amount = 5
                End Select
                
            ElseIf Probabilidad < 200 Then
                Select Case Regalo
                    Case 1
                        .objIndex = 541
                        .Amount = 7
                    Case Else
                        .objIndex = 541
                        .Amount = 10
                End Select
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El regalo al azar de: " & .Amount & " " & ObjData(.objIndex).Name & ", se le otorgo a " & UserList(UserIndexAlAzar).Name & ", tu puedes ser el proximo!!", FontTypeNames.FONTTYPE_GUILD))
            Print #N, "El usuario: " & UserList(UserIndexAlAzar).Name & " Recibio: " & ItemRegalo.Amount & " / " & ObjData(.objIndex).Name
        End With
            Call MeterItemEnInventario(UserIndexAlAzar, ItemRegalo)
        
     ElseIf Not UserList(UserIndexAlAzar).flags.UserLogged Then
        Print #N, "No UserOnline"
    End If
    
    Close #N
End Sub
