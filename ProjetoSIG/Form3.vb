Imports System.Data.OleDb
Public Class Form3
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(ConString)
    Dim Nome, Cargo, RG, CPF, Telefone, Celular, Data_nasc, Sexo, Estado_civil, Nome_Pais, Endereco, Numero, Bairo, Cidade, CEP, Estado, Email, Grau_esc, Horario_de_trabalho, Num_cont, Digito, Agencia As String

    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Form3_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox7.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        NumericUpDown1.Text = ""
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        sql = "Insert into funcionarios(Nome, Cargo, RG, CPF,Telefone,Celular,Data_nasc,Sexo,Estado_civil,Nome_Pais,Endereco,Numero,Bairro,Cidade,CEP,Estado,Email,Grau_esc,Horario_de_trabalho,salario, Vale_Transp, Num_cont,Digito,Agencia,Preco_Transp, preco_passag,nomepais2) Values (@Nome, @Cargo, @RG, @CPF, @Telefone, @Celular, @Data_nasc, @Sexo, @Estado_civil, @Nome_Pais, @Endereco, @Numero, @Bairro, @Cidade, @CEP, @Estado, @Email, @Grau_esc, @Horario_de_trabalho, @salario, @Vale_Transp, @Num_cont, @Digito, @Agencia,@Preco_Transp,@preco_passag,@nomepais2)"
        Dim cm As New OleDbCommand(sql, DBCon)
        cm.Parameters.AddWithValue("@Nome", TextBox1.Text)
        cm.Parameters.AddWithValue("@Cargo", TextBox2.Text)
        cm.Parameters.AddWithValue("@RG", TextBox3.Text)
        cm.Parameters.AddWithValue("@CPF", TextBox4.Text)
        cm.Parameters.AddWithValue("@Telefone", TextBox5.Text)
        cm.Parameters.AddWithValue("@Celular", TextBox6.Text)
        cm.Parameters.AddWithValue("@Data_nasc", DateTimePicker1.Text)
        cm.Parameters.AddWithValue("@Sexo", TextBox8.Text)
        cm.Parameters.AddWithValue("@Estado_civil", TextBox9.Text)
        cm.Parameters.AddWithValue("@Nome_Pais", TextBox10.Text)
        cm.Parameters.AddWithValue("@Endereco", TextBox11.Text)
        cm.Parameters.AddWithValue("@Numero", TextBox12.Text)
        cm.Parameters.AddWithValue("@Bairro", TextBox13.Text)
        cm.Parameters.AddWithValue("@Cidade", TextBox14.Text)
        cm.Parameters.AddWithValue("@CEP", TextBox15.Text)
        cm.Parameters.AddWithValue("@Estado", TextBox16.Text)
        cm.Parameters.AddWithValue("@Email", TextBox17.Text)
        cm.Parameters.AddWithValue("@Grau_esc", TextBox18.Text)
        cm.Parameters.AddWithValue("@Horario_de_trabalho", TextBox19.Text)
        cm.Parameters.AddWithValue("@salario", TextBox23.Text)
        cm.Parameters.AddWithValue("@Vale_Transp", NumericUpDown1.Text)
        cm.Parameters.AddWithValue("@Num_cont", TextBox20.Text)
        cm.Parameters.AddWithValue("@Digito", TextBox21.Text)
        cm.Parameters.AddWithValue("@Agencia", TextBox22.Text)
        cm.Parameters.AddWithValue("@Preco_Transp", TextBox25.Text)
        cm.Parameters.AddWithValue("@preco_passag", TextBox24.Text)
        cm.Parameters.AddWithValue("@nomepais2", TextBox7.Text)
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

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            TextBox25.Text = NumericUpDown1.Value * TextBox24.Text
        Catch

        End Try
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) = False Then
            TextBox3.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) = False Then
            TextBox4.Text = ""
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

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox7.Text) = False Then
            TextBox7.Text = ""
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

    Private Sub TextBox23_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox23.TextChanged
        If IsNumeric(TextBox23.Text) = False Then
            TextBox23.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox24_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox24.TextChanged
        If IsNumeric(TextBox24.Text) = False Then
            TextBox24.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox20.TextChanged
        If IsNumeric(TextBox20.Text) = False Then
            TextBox20.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox21_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox21.TextChanged
        If IsNumeric(TextBox21.Text) = False Then
            TextBox21.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged
     
    End Sub

    Private Sub TextBox22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"

    End Sub

    Private Sub Label9_Click(sender As System.Object, e As System.EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label17_Click(sender As System.Object, e As System.EventArgs) Handles Label17.Click

    End Sub

    Private Sub TextBox10_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged

    End Sub
End Class