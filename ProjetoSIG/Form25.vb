Imports System.Data.OleDb
Public Class Form25
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Codigo, Nome, RG, CPF, Telefone, Celular As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        sql = "Insert into gastolucro(l_gravacao, l_ensaio, l_produtos, d_luz, d_agua, d_tel, d_produto, d_salario, d_transp, d_outras, total_ganhos, total_despesas, total_lucro, dia, aluguel) Values (@l_gravacao, @l_ensaio, @l_produtos, @d_luz, @d_agua, @d_tel, @d_produto, @d_salario, @d_transp, @d_outras, @total_ganhos, @total_despesas, @total_lucro, @dia, @aluguel)"
        Dim cm As New OleDbCommand(sql, DBCon)
        cm.Parameters.AddWithValue("@l_gravacao", TextBox1.Text)
        cm.Parameters.AddWithValue("@l_ensaio", TextBox13.Text)
        cm.Parameters.AddWithValue("@l_produtos", TextBox2.Text)
        cm.Parameters.AddWithValue("@d_luz", TextBox5.Text)
        cm.Parameters.AddWithValue("@d_agua", TextBox6.Text)
        cm.Parameters.AddWithValue("@d_tel", TextBox12.Text)
        cm.Parameters.AddWithValue("@d_produto", TextBox7.Text)
        cm.Parameters.AddWithValue("@d_salario", TextBox8.Text)
        cm.Parameters.AddWithValue("@d_transp", TextBox10.Text)
        cm.Parameters.AddWithValue("@aluguel", TextBox3.Text)
        cm.Parameters.AddWithValue("@d_outras", TextBox11.Text)
        cm.Parameters.AddWithValue("@total_ganhos", TextBox4.Text)
        cm.Parameters.AddWithValue("@total_despesas", TextBox9.Text)
        cm.Parameters.AddWithValue("@total_lucros", Label16.Text)
        cm.Parameters.AddWithValue("@dia", DateTimePicker1.Text)


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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim A, B As Double
        Dim sql As String = "Select valor_f from gravacao"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda

        Dim cm As New OleDb.OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Dim flag As Boolean
        Dim Somatorio As Double
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

            End If
        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    Somatorio += dr.Item("valor_f")
                    TextBox1.Text = Somatorio



                    flag = True



                    End While
                End If
            'If flag = False Then
            '    

            '    End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        DBCon.Close()

        Dim sqla As String = "Select valor_i from Ensaio"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda
        Dim somat As Double
        Dim cma As New OleDb.OleDbCommand(sqla, DBCon)
        Dim dra As OleDb.OleDbDataReader
        Dim flaga As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            dra = cma.ExecuteReader
            If dra.HasRows Then
                While dra.Read

                    somat += dra.Item("valor_i")
                    TextBox13.Text = somat



                    flag = True



                End While
            End If
            'If flag = False Then
            '    MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()


        Dim sqlb As String = "Select preco_total from produtos"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda

        Dim Preco As Double
        Dim cmb As New OleDb.OleDbCommand(sqlb, DBCon)
        Dim drb As OleDb.OleDbDataReader
        Dim flagb As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            drb = cmb.ExecuteReader
            If drb.HasRows Then
                While drb.Read


                    Preco += drb.Item("preco_total")

                    TextBox2.Text = Preco




                    flag = True



                End While
            End If
            'If flagb = False Then
            '    MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()

        Dim sqld As String = "Select valor_compra from produtos"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda
        Dim somad As Double
        Dim cmd As New OleDb.OleDbCommand(sqld, DBCon)
        Dim drd As OleDb.OleDbDataReader
        Dim flagd As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            drd = cmd.ExecuteReader
            If drd.HasRows Then
                While drd.Read

                    somad += drd.Item("valor_compra")
                    TextBox7.Text = somad



                    flag = True



                End While
            End If
            'If flagb = False Then
            '    MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()


        Dim sqlc As String = "Select salario from funcionarios"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda
        Dim somar As Double
        Dim cmc As New OleDb.OleDbCommand(sqlc, DBCon)
        Dim drc As OleDb.OleDbDataReader
        Dim flagc As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            drc = cmc.ExecuteReader
            If drc.HasRows Then
                While drc.Read

                    somar += drc.Item("salario")
                    TextBox8.Text = somar




                    flag = True



                End While
            End If
            'If flag = False Then
            'MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()

        Dim sqle As String = "Select Preco_Transp from funcionarios"

        'SELECT SUM(valor_f) AS SOMA, valor_f FROM gravacao WHERE valor_f='$$linha[valor_f]' Group By valor_
        'SELECT SUM(VALOR_PEDIDO),ID_CLIENTE FROM TAB_PEDIDO(NOLOCK) GROUP BY ID_CLIENT
        '        SELECT nome_banda , SUM(valor_f) FROM gravacao Group by nome_banda
        Dim somarei As Double
        Dim cme As New OleDb.OleDbCommand(sqle, DBCon)
        Dim dre As OleDb.OleDbDataReader
        Dim flage As Boolean
        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            dre = cme.ExecuteReader
            If dre.HasRows Then
                While dre.Read

                    somarei += dre.Item("Preco_Transp")
                    TextBox10.Text = somarei



                    flag = True



                End While
            End If
            'If flag = False Then
            'MessageBox.Show("Pessoa não cadastrada no sistema", "Atenção")

            'End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()



    End Sub

    'Dim SomaColuna As Integer

    'Private Sub ObtemSomaColunas()

    '    Try

    '        'Se o valor for nulo...

    '        If cm.Parameters.AddWithValue("valor_f", TextBox1.Text).Tables(0).Compute("SUM(valor_f)", "") Is DBNull.Value Then
    '            SomaColuna = "0"

    '        Else

    '            'Utilizo o compute para somar os valores da coluna...

    '            SomaColuna = cm.Parameters.AddWithValue("valor_f", TextBox1.Text).Tables(0).Compute("SUM(valor_f)", "") Is DBNull.Value 

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    Private Sub Form25_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If IsNumeric(TextBox4.Text) = False Then
            TextBox4.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Dim A, B, C, D, F As Double
            A = TextBox1.Text

            C = TextBox13.Text

            F = TextBox2.Text

            TextBox4.Text = A + C + F
        Catch ex As Exception

        End Try

        Try
           

            Label16.Text = TextBox4.Text - TextBox9.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) = False Then
            TextBox9.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Dim G, H, I, J, K, L, M, N As Double
            G = TextBox5.Text
            H = TextBox6.Text
            I = TextBox12.Text
            J = TextBox7.Text
            K = TextBox8.Text
            L = TextBox10.Text
            M = TextBox11.Text
            N = TextBox13.Text

            TextBox9.Text = G + H + I + J + K + L + M + N
        Catch ex As Exception

        End Try

        Try


            Label16.Text = TextBox4.Text - TextBox9.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) = False Then
            TextBox5.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        If (TextBox5.Text = "") Then
            TextBox5.Text = 0
        End If
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox6.Text) = False Then
            TextBox6.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        If (TextBox6.Text = "") Then
            TextBox6.Text = 0
        End If
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        If IsNumeric(TextBox12.Text) = False Then
            TextBox12.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        If (TextBox12.Text = "") Then
            TextBox12.Text = 0
        End If
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        If IsNumeric(TextBox11.Text) = False Then
            TextBox11.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        If (TextBox11.Text = "") Then
            TextBox11.Text = 0
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form26.Show()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If IsNumeric(TextBox3.Text) = False Then
            TextBox3.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        If (TextBox3.Text = "") Then
            TextBox3.Text = 0
        End If
    End Sub
End Class