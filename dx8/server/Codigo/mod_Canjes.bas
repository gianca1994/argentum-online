Attribute VB_Name = "mod_Canjes"
Option Explicit
 
Type tCanjes
    Objeto As Obj 'Objeto a canjear
    Copa As Obj 'Copas necesarias para canjear
End Type

Private Const TicketIndex As Integer = 1071
Private CantidadDeTickets As Integer
 
Private Canje1 As tCanjes
Private Canje2 As tCanjes
Private Canje3 As tCanjes
Private Canje4 As tCanjes
 
Public Sub Canjes1(ByVal UserIndex As Integer, ByVal Num1 As Byte)
    
    Dim PiezasAztecas As Integer
    
    With UserList(UserIndex)
        Select Case Num1
        
            Case 0 'Pack 10 Pieza Azteca
                PiezasAztecas = 10
                CantidadDeTickets = 5
                
            Case 1 'Pack 100 Pieza Azteca
                PiezasAztecas = 100
                CantidadDeTickets = 45
                
            Case 2 'Pack 1000 Pieza Azteca
                PiezasAztecas = 1000
                CantidadDeTickets = 400
                
            Case 3 'Pack 1000 Pieza Azteca
                PiezasAztecas = 10000
                CantidadDeTickets = 1000
                
            Case 4 'PERGAMINO DE EXPERIENCIA
                Canje1.Objeto.objIndex = 1067 'PERGAMINO DE EXPERIENCIA
                Canje1.Objeto.Amount = 1
                CantidadDeTickets = 50
                
            Case 5 'ANILLO DE EXPERIENCIA
                Canje1.Objeto.objIndex = 1068 'ANILLO DE EXPERIENCIA
                Canje1.Objeto.Amount = 1
                CantidadDeTickets = 25
                
            Case 6 'PASE VIP
                Canje1.Objeto.objIndex = 1069 'PASE VIP
                Canje1.Objeto.Amount = 1
                CantidadDeTickets = 100
        End Select
        
        If TieneObjetos(TicketIndex, CantidadDeTickets, UserIndex) Then
            Call QuitarObjetos(TicketIndex, CantidadDeTickets, UserIndex)
            
            If PiezasAztecas > 0 Then
                .Stats.AztecPiece = .Stats.AztecPiece + PiezasAztecas
                Call WriteUpdateGold(UserIndex)
            Else
                If Not MeterItemEnInventario(UserIndex, Canje1.Objeto) Then
                    Call TirarItemAlPiso(.Pos, Canje1.Objeto)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Servidor> 　Canje exitoso, Muchas gracias por colaborar con el servidor!!", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> 　No tienes suficientes TICKETS para Canjear!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
End Sub

Public Sub Canjes2(ByVal UserIndex As Integer, ByVal Num2 As Byte)
    With UserList(UserIndex)
        Select Case Num2
            Case 0
                Canje2.Objeto.objIndex = 1049 ' Pieza Azteca
                Canje2.Objeto.Amount = 1
                CantidadDeTickets = 50
        End Select
        
        'Comprobamos que tenga las copas antes de darle el objeto
        If TieneObjetos(TicketIndex, CantidadDeTickets, UserIndex) Then
            'Quitamos las copas
            Call QuitarObjetos(TicketIndex, CantidadDeTickets, UserIndex)
            
            'Nos da el objeto
            If Not MeterItemEnInventario(UserIndex, Canje2.Objeto) Then
                Call TirarItemAlPiso(.Pos, Canje2.Objeto)
            End If
            
            Call WriteConsoleMsg(UserIndex, "Servidor> 　Canje exitoso, Muchas gracias por colaborar con el servidor!!", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> 　No tienes suficientes TICKETS para Canjear!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
End Sub

Public Sub Canjes3(ByVal UserIndex As Integer, ByVal Num3 As Byte)
    With UserList(UserIndex)
        
        Select Case Num3
            Case 0
                Canje3.Objeto.objIndex = 1050 ' Pieza Azteca
                Canje3.Objeto.Amount = 1
                CantidadDeTickets = 50
        End Select
        
        'Comprobamos que tenga las copas antes de darle el objeto
        If TieneObjetos(TicketIndex, CantidadDeTickets, UserIndex) Then
            'Quitamos las copas
            Call QuitarObjetos(TicketIndex, CantidadDeTickets, UserIndex)
            
            'Nos da el objeto
            If Not MeterItemEnInventario(UserIndex, Canje3.Objeto) Then
                Call TirarItemAlPiso(.Pos, Canje3.Objeto)
            End If
            
            Call WriteConsoleMsg(UserIndex, "Servidor> 　Canje exitoso, Muchas gracias por colaborar con el servidor!!", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> 　No tienes suficientes TICKETS para Canjear!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
End Sub

Public Sub Canjes4(ByVal UserIndex As Integer, ByVal Num4 As Byte)
    With UserList(UserIndex)
        
        Select Case Num4
            Case 0
                Canje4.Objeto.objIndex = 1038 ' Pieza Azteca
                Canje4.Objeto.Amount = 1
                CantidadDeTickets = 50
        End Select
        
        'Comprobamos que tenga las copas antes de darle el objeto
        If TieneObjetos(TicketIndex, CantidadDeTickets, UserIndex) Then
            'Quitamos las copas
            Call QuitarObjetos(TicketIndex, CantidadDeTickets, UserIndex)
            
            'Nos da el objeto
            If Not MeterItemEnInventario(UserIndex, Canje4.Objeto) Then
                Call TirarItemAlPiso(.Pos, Canje4.Objeto)
            End If
            
            Call WriteConsoleMsg(UserIndex, "Servidor> 　Canje exitoso, Muchas gracias por colaborar con el servidor!!", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> 　No tienes suficientes TICKETS para Canjear!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
End Sub


