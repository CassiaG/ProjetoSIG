Imports System.Data.OleDb
Public Class Form10
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Codigo, Nome, RG, CPF, Telefone, Celular As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Codigo = TextBox1.Text

        Dim sql As String = "Select * from clientes where Codigo ='" + TextBox1.Text + "'"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Dim flag As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    If (dr.Item("Codigo") = Codigo) Then
                        TextBox2.Text = dr.Item("Nome")
                        TextBox3.Text = dr.Item("Cidade")
                        TextBox4.Text = dr.Item("Estado")
                        TextBox5.Text = dr.Item("Telefone")
                        TextBox6.Text = dr.Item("Celular")
                        TextBox7.Text = dr.Item("Email")
                        DateTimePicker1.Text = dr.Item("Agendamento")


                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        sql = "update clientes set Codigo=@Codigo, Nome=@Nome, Cidade=@Cidade, Estado=@Estado, Telefone=@Telefone, Celular=@Celular, Email=@Email, Agendamento=@Agendamento where Codigo ='" + TextBox1.Text + "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        cm.Parameters.AddWithValue("@Codigo", TextBox1.Text)
        cm.Parameters.AddWithValue("@Nome", TextBox2.Text)
        cm.Parameters.AddWithValue("@Cidade", TextBox3.Text)
        cm.Parameters.AddWithValue("@Estado", TextBox4.Text)
        cm.Parameters.AddWithValue("@Telefone", TextBox5.Text)
        cm.Parameters.AddWithValue("@Celular", TextBox6.Text)
        cm.Parameters.AddWithValue("@Email", TextBox7.Text)
        cm.Parameters.AddWithValue("@Agendamento", DateTimePicker1.Text)

        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            MessageBox.Show("Alterado com Sucesso!")
            cm.ExecuteNonQuery()
            DBCon.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "erro")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Digite um código para consulta", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox1.Focus()
            Return
        End If
        Codigo = TextBox1.Text

        Dim flag As Boolean
        Dim sql As String = "Delete from clientes where Codigo ='" & Codigo & "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Try
            If MessageBox.Show("Você deseja excluir os dados do sistema?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                DBCon.Open()
                cm.ExecuteNonQuery()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()

                DBCon.Close()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form24.ShowDialog()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) = False Then
            TextBox5.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox6.Text) = False Then
            TextBox6.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Form10_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub
End Class