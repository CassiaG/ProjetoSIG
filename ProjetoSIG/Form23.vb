Imports System.Data.OleDb
Public Class Form23
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("Nome"), dr.Item("Cargo"), dr.Item("RG"), dr.Item("CPF"), dr.Item("Telefone"), dr.Item("Celular"), dr.Item("Data_nasc"), dr.Item("Sexo"), dr.Item("Estado_civil"), dr.Item("Nome_Pais"), dr.Item("Endereco"), dr.Item("Numero"), dr.Item("Bairro"), dr.Item("Cidade"), dr.Item("CEP"), dr.Item("Estado"), dr.Item("Email"), dr.Item("Grau_esc"), dr.Item("Horario_de_trabalho"), dr.Item("Num_cont"), dr.Item("Digito"), dr.Item("Agencia"))
            End While
        End If
    End Sub
    Private Sub Form23_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub
    Private Sub Form23_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from funcionarios"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)

        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            exibir()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
        DBCon.Close()
    End Sub
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class