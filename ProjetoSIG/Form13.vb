Imports System.Data.OleDb
Public Class Form13
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Codigo, Nome, RG, CPF, Telefone, Celular As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Insira os valores corretamente", "Atenção")
            Exit Sub
        End If
        Codigo = TextBox1.Text

        Dim sql As String = "Select * from show where Codigo ='" + TextBox1.Text + "'"
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
                        TextBox2.Text = dr.Item("nome_banda")
                        TextBox5.Text = dr.Item("telefone")
                        TextBox6.Text = dr.Item("celular")
                        TextBox7.Text = dr.Item("email")
                        TextBox3.Text = dr.Item("inicio")
                        TextBox4.Text = dr.Item("termino")
                        DateTimePicker1.Text = dr.Item("agendamento")


                        flag = True

                    End If

                End While
            End If
            If flag = False Then
                MessageBox.Show("Banda não cadastrada para fazer Show", "Atenção")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim A, B, C As String

        'Dim sqll As String = "Select * from show where Codigo"
        'Dim cml As New OleDb.OleDbCommand(sqll, DBCon)
        'Dim drl As OleDb.OleDbDataReader

        'If DBCon.State = ConnectionState.Open Then
        '    DBCon.Close()

        'End If
        'Try
        '    DBCon.Open()
        '    drl = cml.ExecuteReader
        '    If drl.HasRows Then
        '        While drl.Read

        '            A = drl.Item("inicio")
        '            B = drl.Item("termino")
        '            C = drl.Item("agendamento")



        '        End While
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'DBCon.Close()

        'If C = DateTimePicker1.Text Then
        '    MessageBox.Show("Possui agendamentos nesse Dia!")

        '    If A = TextBox3.Text And B = TextBox4.Text Then
        '        MessageBox.Show("Já possui agendamento nesse Horário!")
        '        Exit Sub

        '    ElseIf A <> TextBox3.Text And B <> TextBox4.Text Then

        Dim sqli As String
        sqli = "update show set Codigo=@Codigo, nome_banda=@nome_banda, telefone=@telefone, celular=@celular, email=@email, inicio=@inicio, termino=@termino, agendamento=@agendamento where Codigo = '" + TextBox1.Text + "'"
        Dim cm As New OleDbCommand(sqli, DBCon)

        cm.Parameters.AddWithValue("@Codigo", TextBox1.Text)
        cm.Parameters.AddWithValue("@nome_banda", TextBox2.Text)
        cm.Parameters.AddWithValue("@telefone", TextBox5.Text)
        cm.Parameters.AddWithValue("@celular", TextBox6.Text)
        cm.Parameters.AddWithValue("@email", TextBox7.Text)
        cm.Parameters.AddWithValue("@inicio", TextBox3.Text)
        cm.Parameters.AddWithValue("@termino", TextBox4.Text)
        cm.Parameters.AddWithValue("@agendamento", DateTimePicker1.Text)

        If DBCon.State = ConnectionState.Open Then
            DBCon.Close()

        End If
        Try
            DBCon.Open()
            cm.ExecuteNonQuery()
            MessageBox.Show("Alterado com Sucesso!")
            DBCon.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "erro")
        End Try
        '    End If

        'ElseIf C <> DateTimePicker1.Text Then
        'Dim sqli As String
        'sqli = "update show set Codigo=@Codigo, nome_banda=@nome_banda, telefone=@telefone, celular=@celular, email=@email, inicio=@inicio, termino=@termino, agendamento=@agendamento where Codigo = '" + TextBox1.Text + "'"
        'Dim cm As New OleDbCommand(sqli, DBCon)

        'cm.Parameters.AddWithValue("@Codigo", TextBox1.Text)
        'cm.Parameters.AddWithValue("@nome_banda", TextBox2.Text)
        'cm.Parameters.AddWithValue("@telefone", TextBox5.Text)
        'cm.Parameters.AddWithValue("@celular", TextBox6.Text)
        'cm.Parameters.AddWithValue("@email", TextBox7.Text)
        'cm.Parameters.AddWithValue("@inicio", TextBox3.Text)
        'cm.Parameters.AddWithValue("@termino", TextBox4.Text)
        'cm.Parameters.AddWithValue("@agendamento", DateTimePicker1.Text)

        'If DBCon.State = ConnectionState.Open Then
        '    DBCon.Close()

        'End If
        'Try
        '    DBCon.Open()
        '    cm.ExecuteNonQuery()
        '    MessageBox.Show("Alterado com Sucesso!")
        '    DBCon.Close()

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "erro")
        'End Try
        'End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        DateTimePicker1.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Digite um código para consulta", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox1.Focus()
            Return
        End If
        Codigo = TextBox1.Text

        Dim flag As Boolean
        Dim sql As String = "Delete from show where Codigo ='" & Codigo & "'"
        Dim cm As New OleDbCommand(sql, DBCon)
        Dim dr As OleDb.OleDbDataReader
        Try
            If MessageBox.Show("Você deseja excluir os dados do sistema?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                DBCon.Open()
                cm.ExecuteNonQuery()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
                DateTimePicker1.Text = ""
                DBCon.Close()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Form21.ShowDialog()
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
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox6.Text) = False Then
            TextBox6.Text = ""
            'MessageBox.Show("Digite um valor númerico", "Aviso")

        End If
    End Sub
End Class