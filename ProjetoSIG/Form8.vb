Imports System.Data.OleDb
Public Class Form8
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Cod_prod, Codigo, RG, CPF, Telefone, Celular As String
    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    Dim login As String
        login = TextBox6.Text

        If login = "" Then
            MessageBox.Show("Informação Incompleta. Busque os campos que estão faltando.", "Informação incompleta")
            Exit Sub
        End If

        Dim A, B, C, J As String

        Dim sqll As String = "Select * from Ensaio where cod_cliente"
        Dim cml As New OleDb.OleDbCommand(sqll, DBCon)
        Dim drl As OleDb.OleDbDataReader

        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            drl = cml.ExecuteReader
            If drl.HasRows Then
                While drl.Read

                    A = drl.Item("inicio")
                    B = drl.Item("termino")
                    C = drl.Item("dia")
                    J = drl.Item("Estudio")


                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()

        If C = DateTimePicker1.Text And J = TextBox4.Text Then
            MessageBox.Show("Já possui agendamentos para esse Dia e Estúdio, Selecione outro Estúdio e tente novamente!", "Aviso")
            If A = TextBox2.Text And B = TextBox3.Text Then
                MessageBox.Show("Já possui agendamento nesse Horário!", "Aviso")
                Exit Sub

            ElseIf A <> TextBox2.Text And B <> TextBox3.Text Then


                Dim sql As String
                sql = "Insert into Ensaio(cod_cliente, nome_banda, inicio, termino, dia, violao, guitarra, baixo, pratos, valor_f, valor_i, valor_e, Statu, Estudio, teclado) Values (@cod_cliente, @nome_banda, @inicio, @termino, @dia, @violao, @guitarra, @baixo, @pratos, @valor_f, @valor_i, @valor_e,@Statu,@Estudio,@teclado)"
                Dim cm As New OleDb.OleDbCommand(sql, DBCon)
                cm.Parameters.AddWithValue("@cod_cliente", TextBox1.Text)
                cm.Parameters.AddWithValue("@nome_banda", TextBox6.Text)
                'cm.Parameters.AddWithValue("@status", ComboBox1.Text)
                cm.Parameters.AddWithValue("@inicio", TextBox2.Text)
                cm.Parameters.AddWithValue("@termino", TextBox3.Text)
                cm.Parameters.AddWithValue("@dia", DateTimePicker1.Text)
                'cm.Parameters.AddWithValue("@estudioB", ComboBox2.Text)
                cm.Parameters.AddWithValue("@violao", TextBox5.Text)
                cm.Parameters.AddWithValue("@guitarra", TextBox8.Text)
                cm.Parameters.AddWithValue("@baixo", TextBox9.Text)
                cm.Parameters.AddWithValue("@pratos", TextBox10.Text)
                cm.Parameters.AddWithValue("@valor_i", Label20.Text)
                cm.Parameters.AddWithValue("@valor_e", Label14.Text)
                cm.Parameters.AddWithValue("@valor_f", TextBox11.Text)
                cm.Parameters.AddWithValue("@Statu", ComboBox3.Text)
                cm.Parameters.AddWithValue("@Estudio", TextBox4.Text)
                cm.Parameters.AddWithValue("@teclado", TextBox13.Text)

                If DBCon.State = ConnectionState.Open Then
                    'MessageBox.Show("Obrigado!")
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
            End If

        ElseIf C <> DateTimePicker1.Text And J <> TextBox4.Text Then
            Dim sql As String
            sql = "Insert into Ensaio(cod_cliente, nome_banda, inicio, termino, dia, violao, guitarra, baixo, pratos, valor_f, valor_i, valor_e, Statu, Estudio, teclado) Values (@cod_cliente, @nome_banda, @inicio, @termino, @dia, @violao, @guitarra, @baixo, @pratos, @valor_f, @valor_i, @valor_e,@Statu,@Estudio,@teclado)"
            Dim cm As New OleDb.OleDbCommand(sql, DBCon)
            cm.Parameters.AddWithValue("@cod_cliente", TextBox1.Text)
            cm.Parameters.AddWithValue("@nome_banda", TextBox6.Text)
            'cm.Parameters.AddWithValue("@status", ComboBox1.Text)
            cm.Parameters.AddWithValue("@inicio", TextBox2.Text)
            cm.Parameters.AddWithValue("@termino", TextBox3.Text)
            cm.Parameters.AddWithValue("@dia", DateTimePicker1.Text)
            'cm.Parameters.AddWithValue("@estudioB", ComboBox2.Text)
            cm.Parameters.AddWithValue("@violao", TextBox5.Text)
            cm.Parameters.AddWithValue("@guitarra", TextBox8.Text)
            cm.Parameters.AddWithValue("@baixo", TextBox9.Text)
            cm.Parameters.AddWithValue("@pratos", TextBox10.Text)
            cm.Parameters.AddWithValue("@valor_i", Label20.Text)
            cm.Parameters.AddWithValue("@valor_e", Label14.Text)
            cm.Parameters.AddWithValue("@valor_f", TextBox11.Text)
            cm.Parameters.AddWithValue("@Statu", ComboBox3.Text)
            cm.Parameters.AddWithValue("@Estudio", TextBox4.Text)
            cm.Parameters.AddWithValue("@teclado", TextBox13.Text)

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

        ElseIf C = DateTimePicker1.Text And J <> TextBox4.Text Then

            Dim sql As String
            sql = "Insert into Ensaio(cod_cliente, nome_banda, inicio, termino, dia, violao, guitarra, baixo, pratos, valor_f, valor_i, valor_e, Statu, Estudio, teclado) Values (@cod_cliente, @nome_banda, @inicio, @termino, @dia, @violao, @guitarra, @baixo, @pratos, @valor_f, @valor_i, @valor_e,@Statu,@Estudio,@teclado)"
            Dim cm As New OleDb.OleDbCommand(sql, DBCon)
            cm.Parameters.AddWithValue("@cod_cliente", TextBox1.Text)
            cm.Parameters.AddWithValue("@nome_banda", TextBox6.Text)
            'cm.Parameters.AddWithValue("@status", ComboBox1.Text)
            cm.Parameters.AddWithValue("@inicio", TextBox2.Text)
            cm.Parameters.AddWithValue("@termino", TextBox3.Text)
            cm.Parameters.AddWithValue("@dia", DateTimePicker1.Text)
            'cm.Parameters.AddWithValue("@estudioB", ComboBox2.Text)
            cm.Parameters.AddWithValue("@violao", TextBox5.Text)
            cm.Parameters.AddWithValue("@guitarra", TextBox8.Text)
            cm.Parameters.AddWithValue("@baixo", TextBox9.Text)
            cm.Parameters.AddWithValue("@pratos", TextBox10.Text)
            cm.Parameters.AddWithValue("@valor_i", Label20.Text)
            cm.Parameters.AddWithValue("@valor_e", Label14.Text)
            cm.Parameters.AddWithValue("@valor_f", TextBox11.Text)
            cm.Parameters.AddWithValue("@Statu", ComboBox3.Text)
            cm.Parameters.AddWithValue("@Estudio", TextBox4.Text)
            cm.Parameters.AddWithValue("@teclado", TextBox13.Text)

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
        End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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
                        TextBox6.Text = dr.Item("Nome")



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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox13.Text = ""
        TextBox12.Text = ""
        Label12.Text = "0"
        Label16.Text = "0"
        Label17.Text = "0"
        Label18.Text = "0"
        Label20.Text = "0"
        Label14.Text = "0"
        Label25.Text = "0"
        Label40.Text = "0"
        ComboBox3.Text = ""
        TextBox4.Text = "Estúdio "

       
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) = False Then
            TextBox5.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox5.Text = "") Then
                Label12.Text = 0
                Label20.Text = 0
            End If


            Label12.Text = TextBox5.Text * 20
            'Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label14.Text) + Int(Label17.Text) + Int(Label18.Text)

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        If IsNumeric(TextBox8.Text) = False Then
            TextBox8.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox8.Text = "") Then
                Label16.Text = 0
                Label20.Text = 0
            End If

            Label16.Text = TextBox8.Text * 20
            'Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label14.Text) + Int(Label17.Text) + Int(Label18.Text)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        If IsNumeric(TextBox9.Text) = False Then
            TextBox9.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox9.Text = "") Then
                Label17.Text = 0
                Label20.Text = 0
            End If

            Label17.Text = TextBox9.Text * 20
            'Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label14.Text) + Int(Label17.Text) + Int(Label18.Text)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged
        If IsNumeric(TextBox10.Text) = False Then
            TextBox10.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox10.Text = "") Then
                Label18.Text = 0
                Label20.Text = 0
            End If

            Label18.Text = TextBox10.Text * 20
            'Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label14.Text) + Int(Label17.Text) + Int(Label18.Text)

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        If IsNumeric(TextBox12.Text) = False Then
            TextBox12.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox12.Text = "") Then
                Label25.Text = 0

            End If
            Label25.Text = TextBox12.Text - Label14.Text
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        If IsNumeric(TextBox11.Text) = False Then
            TextBox11.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox11.Text = "") Then
                Label14.Text = 0

            End If

            Label14.Text = Int(Label20.Text) + Int(TextBox11.Text)
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
            Try
            Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label17.Text) + Int(Label18.Text) + Int(Label40.Text)
                Label14.Text = Int(Label20.Text) + Int(TextBox11.Text)
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = False Then
            TextBox1.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub
    Private Sub Form8_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        Label12.Text = "0"
        Label16.Text = "0"
        Label17.Text = "0"
        Label18.Text = "0"
        Label20.Text = "0"
        Label14.Text = "0"
        Label25.Text = "0"
        ComboBox3.Text = ""
        TextBox4.Text = "Estúdio "

    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        If IsNumeric(TextBox13.Text) = False Then
            TextBox13.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
        Try
            If (TextBox13.Text = "") Then
                Label40.Text = 0
                Label20.Text = 0
            End If
            Label40.Text = TextBox13.Text * 20
            'Label20.Text = Int(Label12.Text) + Int(Label16.Text) + Int(Label14.Text) + Int(Label17.Text) + Int(Label18.Text)
            'Label14.Text = Label20.Text
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class