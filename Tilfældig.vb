Public Class Form1

    Shared random As New Random()
    Dim Kort(52) As Integer
    Dim Bunke(24) As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For UdKort As Integer = 1 To 28
            MsgBox(random.Next(1, 52))
        Next

    End Sub
End Class
