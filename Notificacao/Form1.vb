Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        'userLOGADO()
        'pegaIdUser()
        'procuraTarefas()
        'popUp()
        'conexaoAccess.Close()
    End Sub
    
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        userLOGADO()
        'pegaIdUser()
        procuraTarefas()
        If SW = True Then
            consultaIntervalo()
            If Intervalo = 0 Then Exit Sub
            Timer1.Interval = Intervalo
            popUp()
        End If

        conexaoAccess.Close()

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
    End Sub
End Class
