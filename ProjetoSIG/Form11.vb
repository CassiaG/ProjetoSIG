Imports System.Data.OleDb
Public Class Form11
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(ConString)
    Dim Nome, Cargo, RG, CPF, Telefone, Celular, Data_nasc, Sexo, Estado_civil, Nome_Pais, Endereco, Numero, Bairo, Cidade, CEP, Estado, Email, Grau_esc, Horario_de_trabalho, Num_cont, Digito, Agencia As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Nome = TextBox1.Text

        Dim sql As String = "Select * from funcionarios where Nome ='" + TextBox1.Text + "'"


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
                    If (dr.Item("Nome") = Nome) Then
                        TextBox2.Text = dr.Item("Cargo")
                        TextBox3.Text = dr.Item("RG")
                        TextBox4.Text = dr.Item("CPF")
                        TextBox5.Text = dr.Item("Telefone")
                        TextBox6.Text = dr.Item("Celular")
                        DateTimePicker1.Text = dr.Item("Data_nasc")
                        TextBox8.Text = dr.Item("Sexo")
                        TextBox9.Text = dr.Item("Estado_civil")
                        TextBox10.Text = dr.Item("Nome_Pais")
                        TextBox11.Text = dr.Item("Endereco")
                        TextBox12.Text = dr.Item("Numero")
                        TextBox13.Text = dr.Item("Bairro")
                        TextBox14.Text = dr.Item("Cidade")
                        TextBox15.Text = dr.Item("CEP")
                        TextBox16.Text = dr.Item("Estado")
                        TextBox17.Text = dr.Item("Email")
                        TextBox18.Text = dr.Item("Grau_esc")
                        TextBox19.Text = dr.Item("Horario_de_trabalho")
                        TextBox23.Text = dr.Item("salario")
                        NumericUpDown1.Text = dr.Item("Vale_Transp")
                        TextBox20.Text = dr.Item("Num_cont")
                        TextBox21.Text = dr.Item("Digito")
                        TextBox22.Text = dr.Item("Agencia")
                        TextBox25.Text = dr.Item("Preco_Transp")
                        TextBox24.Text = dr.Item("preco_passag")
                        TextBox7.Text = dr.Item("nomepais2")
                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Funcionário não cadastrado no sistema", "Atenção")

            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Digite um código para consulta", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox1.Focus()
            Return
        End If
        Nome = TextBox1.Text

        Dim flag As Boolean
        Dim sql As String = "Delete from funcionarios where Nome ='" & Nome & "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Try
            If MessageBox.Show("Você deseja excluir os dados do sistema?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                DBCon.Open()
                cm.ExecuteNonQuery()
                DateTimePicker1.Text = ""
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
                DBCon.Close()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        sql = "update funcionarios set Nome=@Nome, Cargo=@Cargo, RG=@RG, CPF=@CPF, Telefone=@Telefone, Celular=@Celular, Data_nasc=@Data_nasc, Sexo=@Sexo, Estado_civil=@Estado_civil, Nome_Pais=@Nome_Pais, Endereco=@Endereco, Numero=@Numero, Bairro=@Bairro, Cidade=@Cidade, CEP=@CEP, Estado=@Estado, Email=@Email, Grau_esc=@Grau_esc, Horario_de_trabalho=@Horario_de_trabalho, salario=@salario, Vale_Transp=@Vale_Transp, Num_cont=@Num_cont, Digito=@Digito, Agencia=@Agencia, Preco_Transp=@Preco_Transp, preco_passag=@preco_passag, nomepais2=@nomepais2 where Nome ='" + TextBox1.Text + "'"
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
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
        TextBox25.Text = ""
        TextBox24.Text = ""
        TextBox7.Text = ""
        DateTimePicker1.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form23.ShowDialog()
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

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

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged
        If IsNumeric(TextBox25.Text) = False Then
            TextBox25.Text = ""
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

    Private Sub Form11_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Text = ""
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

    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
    End Sub
End Class