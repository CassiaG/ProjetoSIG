Imports System.Data.OleDb
Public Class Form14
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Cod_prod, Nome, RG, CPF, Telefone, Celular As String


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Cod_prod = TextBox1.Text

        Dim sql As String = "Select * from produtos where cod_produto ='" + TextBox1.Text + "'"
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
                    If (dr.Item("cod_produto") = Cod_prod) Then
                        TextBox2.Text = dr.Item("nomeE")
                        TextBox6.Text = dr.Item("nomeP")
                        TextBox3.Text = dr.Item("preco")
                        TextBox4.Text = dr.Item("quantidade")
                        TextBox5.Text = dr.Item("desc")

                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sql As String = "update produtos set nomeE=@nomeE, nomeP=@nomeP, preco=@preco, quantidade=@quantidade, desc=@desc where cod_produto ='" & TextBox1.Text & "'"
        'sqli = ("UPDATE produtos SET @nomeE = '" & TextBox2.Text & "', @nomeP= '" & TextBox6.Text & "', @preco= '" & TextBox3.Text & "', @quantidade= '" & TextBox4.Text & "', @desc= '" & TextBox6.Text & "' where @cod_produto='" & TextBox1.Text & "'")

        Dim cm As New OleDbCommand(sql, DBCon)

        cm.Parameters.AddWithValue("@nomeE", TextBox2.Text)
        cm.Parameters.AddWithValue("@nomeP", TextBox6.Text)
        cm.Parameters.AddWithValue("@preco", TextBox3.Text)
        cm.Parameters.AddWithValue("@quantidade", TextBox4.Text)
        cm.Parameters.AddWithValue("@desc", TextBox6.Text)

        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            cm.ExecuteNonQuery()
            DBCon.Close()
            MsgBox("Alteroou")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "erro")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Digite um código para consulta", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox1.Focus()
            Return
        End If
        Cod_prod = TextBox1.Text

        Dim flag As Boolean
        Dim sql As String = "Delete from produtos where cod_produto ='" & Cod_prod & "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Try
            If MessageBox.Show("Você deseja excluir os dados do sistema?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                DBCon.Open()
                cm.ExecuteNonQuery()
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
     

    End Sub


    Private Sub Button5_Click_1(sender As System.Object, e As System.EventArgs) Handles Button5.Click

        Form20.ShowDialog()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
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

    Private Sub Form14_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
 
    End Sub
End Class