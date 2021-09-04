Imports System.Data.OleDb
Public Class Form4
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(ConString)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim sql As String
        sql = "Insert into fornecedores(CNPJ, Nome_Empresa, Data_Cad, Data_Alt, Telefone, Celular, Endereco, Numero, Bairro, Cidade, Cep, Estado, E_mail, Razao_social, Atendimento) Values (@CNPJ, @Nome_Empresa, @Data_Cad, @Data_Alt, @Telefone, @Celular, @Endereço, @Numero, @Bairro, @Cidade, @CEP, @Estado, @E_mail, @Razao_social, @Atendimento)"
        Dim cm As New OleDbCommand(sql, DBCon)
        cm.Parameters.AddWithValue("@CNPJ", TextBox1.Text)
        cm.Parameters.AddWithValue("@Nome_Empresa", TextBox2.Text)
        cm.Parameters.AddWithValue("@Data_Cad", TextBox3.Text)
        cm.Parameters.AddWithValue("@Data_Alt", TextBox4.Text)
        cm.Parameters.AddWithValue("@Telefone", TextBox5.Text)
        cm.Parameters.AddWithValue("@Celular", TextBox6.Text)
        cm.Parameters.AddWithValue("@Endereco", TextBox11.Text)
        cm.Parameters.AddWithValue("@Numero", TextBox12.Text)
        cm.Parameters.AddWithValue("@Bairro", TextBox13.Text)
        cm.Parameters.AddWithValue("@Cidade", TextBox14.Text)
        cm.Parameters.AddWithValue("@Cep", TextBox15.Text)
        cm.Parameters.AddWithValue("@Estado", TextBox16.Text)
        cm.Parameters.AddWithValue("@E_mail", TextBox17.Text)
        cm.Parameters.AddWithValue("@Razao_social", TextBox18.Text)
        cm.Parameters.AddWithValue("@Atendimento", TextBox19.Text)
        If DBCon.State = ConnectionState.Open Then
            MessageBox.Show("Obrigado!")
            DBCon.Close()
            Me.Close()

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


    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Form5.ShowDialog()

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        'If IsNumeric(TextBox3.Text) = False Then
        '    TextBox3.Text = ""
        '    'MessageBox.Show("Digite um valor númerico", "Aviso")

        'End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        'If IsNumeric(TextBox4.Text) = False Then
        '    TextBox4.Text = ""
        '    'MessageBox.Show("Digite um valor númerico", "Aviso")

        'End If
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

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        If IsNumeric(TextBox12.Text) = False Then
            TextBox12.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox15.TextChanged
        If IsNumeric(TextBox15.Text) = False Then
            TextBox15.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub Form4_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
    End Sub
End Class