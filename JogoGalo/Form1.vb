Public Class Form1
    Dim Ajogar, Galo(9) As Integer
    Dim Quadros() As PictureBox
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call InicializaGalo()
    End Sub
    Private Sub InicializaGalo()
        For i = 0 To 8
            Galo(i) = 0
            Call Imagem(i, 0)
        Next
    End Sub
    Private Sub Clicar(sender As System.Object, e As System.EventArgs) Handles Q1.Click, Q2.Click, Q3.Click, Q4.Click, Q5.Click, Q6.Click, Q7.Click, Q8.Click, Q9.Click
        Dim quadro As Integer
        Select Case sender.name
            Case "Q1" : quadro = 0
            Case "Q2" : quadro = 1
            Case "Q3" : quadro = 2
            Case "Q4" : quadro = 3
            Case "Q5" : quadro = 4
            Case "Q6" : quadro = 5
            Case "Q7" : quadro = 6
            Case "Q8" : quadro = 7
            Case "Q9" : quadro = 8
        End Select
        If (Galo(quadro) = 0) Then
            Call Imagem(quadro, 1)
            Call Verseganhou()
            Call JogaPC()
            Call Verseganhou()
        End If
    End Sub
    Private Sub Imagem(quadro, img)
        Dim fig As New PictureBox
        If img = 1 Then : fig.Image = My.Resources.Cruz
        ElseIf img = 2 Then : fig.Image = My.Resources.Bola
        Else : fig.Image = My.Resources.Vazio
        End If
        Quadros(quadro).Image = fig.Image
        Galo(quadro) = img
    End Sub
    Private Function M(a, b, c) As Integer
        Dim r = Galo(a) * Galo(b) * Galo(c)
        If r = 1 Or r = 8 Then : Return r
        Else : Return (0)
        End If
    End Function
    Private Sub Verseganhou()
        Dim g = M(0, 1, 2) + M(3, 4, 5) + M(6, 7, 8) + M(0, 3, 6) + M(1, 4, 7) + M(2, 5, 8) + M(0, 4, 8) + M(2, 4, 6)
        If g > 7 Then
            MsgBox("Perdestes!!! Jogar Novamente? ")
            Call InicializaGalo()
        ElseIf g > 0 Then
            MsgBox("Ganhastes!!! Jogar Novamente? ")
            Call InicializaGalo()
        Else
            Dim empate = True
            For i = 0 To 8
                If Galo(i) = 0 Then empate = False
            Next
            If empate Then
                MsgBox("empatastes!!! Jogar Novamente? ")
                Call InicializaGalo()
            End If
        End If
    End Sub
    Private Function A(q1, q2, q3, v) As Integer
        If Galo(q1) = 0 And Galo(q2) = 2 And Galo(q3) = 2 Then : Return q1
        ElseIf Galo(q1) = 2 And Galo(q2) = 0 And Galo(q3) = 2 Then : Return q2
        ElseIf Galo(q1) = 2 And Galo(q2) = 2 And Galo(q3) = 0 Then : Return q3
        Else : Return v
        End If
    End Function
    Private Function D(q1, q2, q3, v) As Integer
        If Galo(q1) = 0 And Galo(q2) = 1 And Galo(q3) = 1 Then : Return q1
        ElseIf Galo(q1) = 1 And Galo(q2) = 0 And Galo(q3) = 1 Then : Return q2
        ElseIf Galo(q1) = 1 And Galo(q2) = 1 And Galo(q3) = 0 Then : Return q3
        Else : Return v
        End If
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Quadros = {Q1, Q2, Q3, Q4, Q5, Q6, Q7, Q8, Q9}
    End Sub
    Private Sub JogaPC()
        Dim i As Integer
        Do : i = Int(Rnd() * 9)
        Loop While Galo(i) > 0
        i = D(0, 1, 2, i)
        i = D(3, 4, 5, i)
        i = D(6, 7, 8, i)
        i = D(0, 3, 6, i)
        i = D(1, 4, 7, i)
        i = D(2, 5, 8, i)
        i = D(0, 4, 8, i)
        i = D(2, 4, 6, i)
        i = A(0, 1, 2, i)
        i = A(3, 4, 5, i)
        i = A(6, 7, 8, i)
        i = A(0, 3, 6, i)
        i = A(1, 4, 7, i)
        i = A(2, 5, 8, i)
        i = A(0, 4, 8, i)
        i = A(2, 4, 6, i)
        Call Imagem(i, 2)
    End Sub
End Class