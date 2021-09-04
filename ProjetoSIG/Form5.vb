Imports System.Data.OleDb
Public Class Form5
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(ConString)
    'Dim Cod_produto, nomeE, nomeP, preco, quantidade, desc As String

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim login, senha As String
        login = TextBox8.Text
        senha = TextBox7.Text

        If login = "" Or senha = "" Then
            MessageBox.Show("Informação Incompleta. Busque os campos com o nome do telefone e celular da Empresa.", "Informação incompleta")
            Exit Sub
        End If

        Dim sql As String
        sql = "INSERT INTO produtos values (@cod_produto,@nomeE,@nomeP,@preco,@quantidade,@desc,@Quantidade_final,@valor_compra, @preco_total)"
        Dim cm As New OleDbCommand(Sql, DBCon)

        cm.Parameters.AddWithValue("@cod_produto", TextBox1.Text)
        cm.Parameters.AddWithValue("@nomeE", TextBox6.Text)
        cm.Parameters.AddWithValue("@nomeP", TextBox2.Text)
        cm.Parameters.AddWithValue("@preco", TextBox4.Text)
        cm.Parameters.AddWithValue("@quantidade", TextBox3.Text)
        cm.Parameters.AddWithValue("@desc", TextBox5.Text)
        cm.Parameters.AddWithValue("@Quantidade_final", TextBox9.Text)
        cm.Parameters.AddWithValue("@valor_compra", TextBox10.Text)
        cm.Parameters.AddWithValue("@preco_total", TextBox11.Text)
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox6.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Dim Cod_prod As String
        Cod_prod = TextBox6.Text

        Dim sql As String = "Select * from fornecedores where Nome_Empresa ='" + TextBox6.Text + "'"
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
                    If (dr.Item("Nome_Empresa") = Cod_prod) Then
                        TextBox6.Text = dr.Item("Nome_Empresa")
                        TextBox8.Text = dr.Item("Telefone")
                        TextBox7.Text = dr.Item("Celular")

                    End If
                    flag = True

                End While
            End If

            If flag = False Then
                MessageBox.Show("Empresa não cadastrado no sistema", "Atenção")

            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        If IsNumeric(TextBox11.Text) = False Then
            TextBox11.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        If IsNumeric(TextBox8.Text) = False Then
            TextBox8.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        If IsNumeric(TextBox7.Text) = False Then
            TextBox7.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) = False Then
            TextBox4.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) = False Then
            TextBox3.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) = False Then
            TextBox9.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged
        If IsNumeric(TextBox10.Text) = False Then
            TextBox10.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub Form5_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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
    End Sub
End Class