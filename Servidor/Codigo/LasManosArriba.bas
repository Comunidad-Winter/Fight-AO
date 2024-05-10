Attribute VB_Name = "LasManosArriba"
   Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = tIndex
    UserList(tIndex).flags.EstaDueleando = True
    Call WarpUserChar(tIndex, 14, 27, 46)
    UserList(tIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 14, 40, 55)
    Call SendData(ToAll, 0, 0, "||Retos: " & UserList(tIndex).name & " y " & UserList(UserIndex).name & " van a jugar un reto." & "~0~200~0~0~0")
    End Sub
    Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal tIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    UserList(tIndex).flags.EsperandoDuelo = False
    UserList(tIndex).flags.Oponente = 0
    UserList(tIndex).flags.EstaDueleando = False
    Call WarpUserChar(UserIndex, 1, 53, 45)
    Call WarpUserChar(tIndex, 1, 53, 44)
    End Sub
    Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos: " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto." & "~0~200~0~0~1")
    UserList(Ganador).Stats.RetosGanados = UserList(Ganador).Stats.RetosGanados + 1
    UserList(Perdedor).Stats.RetosPerdidos = UserList(Perdedor).Stats.RetosPerdidos + 1
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & "~0~200~0~0~1")
    Call ResetDuelo(Ganador, Perdedor)
    End Sub


