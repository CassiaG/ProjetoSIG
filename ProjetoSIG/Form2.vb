Imports System.Data.OleDb

Public Class Form2
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Codigo, Nome, RG, CPF, Telefone, Celular As String

  

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim sql As String
        sql = "Insert into clientes(Codigo, Nome, Cidade, Estado, Telefone, Celular, Email, Agendamento) Values (@Codigo, @Nome, @Cidade, @Estado, @Telefone, @Celular, @Email, @Agendamento)"
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
           

        End If
        Try
            DBCon.Open()
            cm.ExecuteNonQuery()
            MessageBox.Show("Obrigado!")

            DBCon.Close()
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "erro")
        End Try
        DBCon.Close()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
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

        End If
        'MessageBox.Show("Digite um valor númerico", "Aviso")
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox6.Text) = False Then
            TextBox6.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub
End Class