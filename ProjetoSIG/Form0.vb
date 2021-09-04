Imports System.Data.OleDb

Public Class Form0
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim conta As Integer
    Dim login, senha As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        login = TextBox1.Text
        senha = TextBox2.Text

        If login = "" Or senha = "" Then
            MessageBox.Show("Informação Incompleta.Preencha os campos com o nome do usuário e senha.", "Informação incompleta")
            TextBox1.Focus()
        End If
        Dim sql As String = "Select Usuario, Senha from Usuarios where Usuario = '" & login & "'"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Dim flag As Boolean
        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    If (dr.Item("Senha") = senha) And (dr.Item("Usuario") = login) Then
                        flag = True


                    End If


                End While


            End If
            If flag = True Then
                Visible = False

                Form1.ShowDialog()



            Else
                MessageBox.Show("Acesso Negado!")
                conta = conta + 1 '
            End If
            If conta = 3 Then
                MessageBox.Show("Limite de erros alcançados.")
                Me.Close()
                End
            End If
            DBCon.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro Genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        login = TextBox1.Text
        senha = TextBox2.Text

        If login = "" Or senha = "" Then
            MessageBox.Show("Informação Incompleta.Preencha os campos com o nome do usuário e senha.", "Informação incompleta")
            TextBox1.Focus()
        End If
        Dim sql As String = "Select funcionario, senha2 from Usuarios where funcionario = '" & login & "'"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Dim flag As Boolean
        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    If (dr.Item("senha2") = senha) And (dr.Item("funcionario") = login) Then
                        flag = True


                    End If


                End While


            End If
            If flag = True Then
                Visible = False

                Form27.ShowDialog()



            Else
                MessageBox.Show("Acesso Negado!")
                conta = conta + 1 '
            End If
            If conta = 3 Then
                MessageBox.Show("Limite de erros alcançados.")
                Me.Close()
                End
            End If
            DBCon.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro Genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
