Imports System.Data.OleDb
Public Class Form6
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Cod_prod, Nome, RG, CPF, Telefone, Celular As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      
        Dim login, login2 As String
        login = TextBox2.Text

        login2 = TextBox3.Text
        

        If login = "" Or login2 = "" Then
            MessageBox.Show("Informação Incompleta. Busque os campos que estão faltando.", "Informação incompleta")
            Exit Sub
        End If
     
        If TextBox7.Text >= Label7.Text Then 
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()
        End If

        Dim sqlt As String
        sqlt = "Insert into VendaProd(Nome_P, Nome_E, Preco, Quantidade_S, Valor_P,cod_prod) Values (@Nome_P, @Nome_E, @Preco, @Quantidade_S, @Valor_P,@cod_prod)"
        Dim cmt As New OleDbCommand(sqlt, DBCon)

        cmt.Parameters.AddWithValue("@Nome_P", TextBox2.Text)
        cmt.Parameters.AddWithValue("@Nome_E", TextBox6.Text)
        cmt.Parameters.AddWithValue("@Preco", TextBox3.Text)
        cmt.Parameters.AddWithValue("@Quantidade_S", NumericUpDown1.Value)
        cmt.Parameters.AddWithValue("@Valor_P", TextBox7.Text)

        cmt.Parameters.AddWithValue("@cod_prod", TextBox1.Text)


           

        Try
            DBCon.Open()
            cmt.ExecuteNonQuery()
            'MessageBox.Show("Obrigado!")

            DBCon.Close()
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "erro")

        End Try
        Else
            MessageBox.Show("Valor insuficiente para a compra", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        Dim S, D As Double
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
                        TextBox4.Text = dr.Item("quantidade")
                        TextBox4.Text = TextBox4.Text - NumericUpDown1.Value
                        TextBox8.Text = dr.Item("preco_total")
                        S = TextBox8.Text
                        D = Label7.Text
                        S = S + D
                        'TextBox8.Text = Int(TextBox8.Text) + Int(Label7.Text)


                    End If

                End While
            End If
            'If flagl = False Then
            '    'MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

            'End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Dim sqli As String
        sqli = "update produtos set quantidade=@quantidade, preco_total=@preco_total where cod_produto ='" + TextBox1.Text + "'"
        Dim cmi As New OleDbCommand(sqli, DBCon)

        cmi.Parameters.AddWithValue("@quantidade", TextBox4.Text)
        cmi.Parameters.AddWithValue("@preco_total", S)
        'TextBox8.Text
        If DBCon.State = ConnectionState.Open Then
            MessageBox.Show("Obrigado!")

            DBCon.Close()
            Me.Close()


        End If
        Try
            DBCon.Open()
            cm.ExecuteNonQuery()
            cmi.ExecuteNonQuery()
            'MessageBox.Show("Obrigado!")

            DBCon.Close()
            Me.Close()

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "erro")
        End Try
    End Sub


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
                        TextBox6.Text = dr.Item("nomeE")
                        TextBox2.Text = dr.Item("nomeP")
                        TextBox3.Text = dr.Item("preco")
                        TextBox4.Text = dr.Item("quantidade")
                        NumericUpDown1.Maximum = dr.Item("quantidade")
                        TextBox9.Text = dr.Item("desc")
                        TextBox5.Text = dr.Item("Quantidade_final")
                        TextBox8.Text = dr.Item("preco_total")
                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs)

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        Label7.Text = "0"
        Label10.Text = "0"

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) = False Then
            TextBox3.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
     


    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        Label10.Text = 0


    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) = False Then
            TextBox4.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If

    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim A As String
        Dim sql As String = "Select * from produtos where cod_produto"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        'Dim flag As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read

                    If (dr.Item("cod_produto") = Cod_prod) Then
                        'TextBox10.Text = 0
                        A = NumericUpDown1.Value
                        TextBox3.Text = dr.Item("preco")
                        Label7.Text = dr.Item("preco")
                        Label7.Text = Label7.Text * A

                        Try
                            TextBox7.Text = 0
                            Label10.Text = TextBox7.Text - Label7.Text
                        Catch ex As Exception
                            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                        'flag = True
                    End If


                End While
            End If
            'If flag = False Then
            'MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

            'End If
        Catch ex As Exception
            '    'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()

        'Dim C As String
        'Dim sqll As String = "Select * from produtos where cod_produto"
        'Dim cml As New OleDb.OleDbCommand(sqll, DBCon)
        'Dim drl As OleDb.OleDbDataReader
        'Dim flagl As Boolean
        'If DBCon.State = ConnectionState.Open Then
        '    DBCon.Close()

        'End If
        'Try
        '    DBCon.Open()
        '    drl = cml.ExecuteReader
        '    If drl.HasRows Then
        '        While drl.Read

        '            If (drl.Item("cod_produto") = Cod_prod) Then
        '                C = TextBox10.Text
        '                TextBox4.Text = drl.Item("quantidade")
        '                Label12.Text = drl.Item("quantidade")
        '                Label12.Text = Label12.Text - C

        '            End If
        '        End While
        '    End If
        '    'If flagl = False Then
        '    '    'MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

        '    'End If
        'Catch ex As Exception
        '    'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'DBCon.Close()

        'Dim sq As String = "Select * from VendaProd where cod_prod"
        'Dim cmm As New OleDb.OleDbCommand(sq, DBCon)
        'Dim drr As OleDb.OleDbDataReader
        'Dim flag As Boolean
        'Dim B As String
        'If DBCon.State = ConnectionState.Open Then
        '    DBCon.Close()

        'End If
        'Try
        '    DBCon.Open()
        '    drr = cmm.ExecuteReader
        '    If drr.HasRows Then
        '        While drr.Read

        '            If (drr.Item("cod_prod") = Cod_prod) Then

        '                'TextBox4.Text = drl.Item("quantidade")

        '                Label13.Text = drr.Item("quantidade_f")
        '                Label13.Text = Label13.Text - TextBox10.Text

        '            End If


        '        End While
        '    End If
        '    'If flagl = False Then
        '    '    'MessageBox.Show("Produto não cadastrado no sistema", "Atenção")

        '    'End If
        'Catch ex As Exception
        '    'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'DBCon.Close()

    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        If IsNumeric(TextBox7.Text) = False Then
            TextBox7.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            Dim a As Double = TextBox7.Text
            Dim b As Double = Label7.Text


            Label10.Text = a - b

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox5_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) = False Then
            TextBox5.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub Label8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            Label7.Text = NumericUpDown1.Value * TextBox3.Text
        Catch

        End Try
    End Sub

    Private Sub GroupBox2_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Form6_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        Label10.Text = "0"
        Label7.Text = "0"
        NumericUpDown1.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        If IsNumeric(TextBox8.Text) = False Then
            TextBox8.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub
End Class