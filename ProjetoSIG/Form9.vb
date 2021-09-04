Public Class Form9

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Form0.Close()
        Form1.Close()
        Form2.Close()
        Form3.Close()
        Form4.Close()
        Form5.Close()
        Me.Close()





    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Visible = False
        Form1.Show()


    End Sub
End Class