Imports System.Data.OleDb
Public Class Form12
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(ConString)
    Dim Nome_Empresa As String
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Digite um código para consulta", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox2.Focus()
            Return
        End If
        Nome_Empresa = TextBox2.Text

        Dim flag As Boolean
        Dim sql As String = "Delete from fornecedores where Nome_Empresa ='" & Nome_Empresa & "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Try
            If MessageBox.Show("Você deseja excluir os dados do sistema?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                DBCon.Open()
                cm.ExecuteNonQuery()
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
                DBCon.Close()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Nome_Empresa = TextBox2.Text

        Dim sql As String = "Select * from fornecedores where Nome_Empresa ='" + TextBox2.Text + "'"
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
                    If (dr.Item("Nome_Empresa") = Nome_Empresa) Then
                        TextBox1.Text = dr.Item("CNPJ")
                        TextBox3.Text = dr.Item("Data_Cad")
                        TextBox4.Text = dr.Item("Data_Alt")
                        TextBox5.Text = dr.Item("Telefone")
                        TextBox6.Text = dr.Item("Celular")
                        TextBox11.Text = dr.Item("Endereco")
                        TextBox12.Text = dr.Item("Numero")
                        TextBox13.Text = dr.Item("Bairro")
                        TextBox14.Text = dr.Item("Cidade")
                        TextBox15.Text = dr.Item("Cep")
                        TextBox16.Text = dr.Item("Estado")
                        TextBox17.Text = dr.Item("E_mail")
                        TextBox18.Text = dr.Item("Razao_social")
                        TextBox19.Text = dr.Item("Atendimento")



                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Empresa não cadastrada no sistema", "Atenção")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        sql = "update fornecedores set CNPJ=@CNPJ, Nome_Empresa=@Nome_Empresa, Data_Cad=@Data_Cad, Data_Alt=@Data_Alt, Telefone=@Telefone, Celular=@Celular, Endereco=@Endereco, Numero=@Numero, Bairro=@Bairro, Cidade= @Cidade, CEP=@CEP, Estado=@Estado, E_mail=@E_mail, Razao_social=@Razao_social, Atendimento=@Atendimento where Nome_Empresa ='" + TextBox2.Text + "'"
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form22.ShowDialog()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        'If IsNumeric(TextBox1.Text) = False Then
        '    TextBox1.Text = ""
        '    MessageBox.Show("Digite um valor númerico", "Aviso")

        'End If
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

    Private Sub Form12_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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