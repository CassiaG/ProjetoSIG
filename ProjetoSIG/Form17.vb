Imports System.Data.OleDb
Public Class Form17
    Dim constring As String = "Provider=Microsoft.ACE.Oledb.12.0;Password="""";User ID=Admin; Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim Codigo, Nome, RG, CPF, Telefone, Celular As String
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
                        TextBox2.Text = dr.Item("Nome")
                        TextBox5.Text = dr.Item("Telefone")
                        TextBox6.Text = dr.Item("Celular")
                        TextBox7.Text = dr.Item("Email")


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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim logi As String
        logi = TextBox2.Text

        If logi = "" Then
            MessageBox.Show("Informação Incompleta. Busque os campos que estão faltando.", "Informação incompleta")
            Exit Sub
        End If
        Dim A, B, C As String

        Dim sqll As String = "Select * from show where Codigo"
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
                    C = drl.Item("agendamento")



                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro genérico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        DBCon.Close()

        If C = DateTimePicker1.Text Then
            MessageBox.Show("Possui agendamentos nesse Dia!")
            If A = TextBox3.Text And B = TextBox4.Text Then
                MessageBox.Show("Já possui agendamento nesse Horário!")
                Exit Sub

            ElseIf A <> TextBox3.Text And B <> TextBox4.Text Then


                Dim sql As String
                sql = "Insert into show(Codigo, nome_banda, telefone, celular, email, inicio, termino, agendamento) Values (@Codigo, @nome_banda, @telefone, @celular, @email, @inicio, @termino, @agendamento)"
                Dim cm As New OleDbCommand(sql, DBCon)
                cm.Parameters.AddWithValue("@Codigo", TextBox1.Text)
                cm.Parameters.AddWithValue("@nome_banda", TextBox2.Text)
                cm.Parameters.AddWithValue("@telefone", TextBox5.Text)
                cm.Parameters.AddWithValue("@celular", TextBox6.Text)
                cm.Parameters.AddWithValue("@email", TextBox7.Text)
                cm.Parameters.AddWithValue("@inicio", TextBox3.Text)
                cm.Parameters.AddWithValue("@termino", TextBox4.Text)
                cm.Parameters.AddWithValue("@agendamento", DateTimePicker1.Text)

                If DBCon.State = ConnectionState.Open Then
                    '    MessageBox.Show("Obrigado!")
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

            ElseIf C <> DateTimePicker1.Text Then
                Dim sql As String
                sql = "Insert into show(Codigo, nome_banda, telefone, celular, email, inicio, termino, agendamento) Values (@Codigo, @nome_banda, @telefone, @celular, @email, @inicio, @termino, @agendamento)"
                Dim cm As New OleDbCommand(sql, DBCon)
                cm.Parameters.AddWithValue("@Codigo", TextBox1.Text)
                cm.Parameters.AddWithValue("@nome_banda", TextBox2.Text)
                cm.Parameters.AddWithValue("@telefone", TextBox5.Text)
                cm.Parameters.AddWithValue("@celular", TextBox6.Text)
                cm.Parameters.AddWithValue("@email", TextBox7.Text)
                cm.Parameters.AddWithValue("@inicio", TextBox3.Text)
                cm.Parameters.AddWithValue("@termino", TextBox4.Text)
                cm.Parameters.AddWithValue("@agendamento", DateTimePicker1.Text)

                If DBCon.State = ConnectionState.Open Then
                ' MessageBox.Show("Obrigado!")
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

        'Dim sql As String
        'sql = ("Select * from show Where agendamento = @agendamento")
        'Dim ds As New DataSet()
    End Sub

    Private Sub DateTimePicker1_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd/MM/yyyy"
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
    Private Sub Form17_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
    End Sub
End Class