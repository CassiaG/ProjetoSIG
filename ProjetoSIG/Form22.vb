Imports System.Data.OleDb
Public Class Form22
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("CNPJ"), dr.Item("Nome_Empresa"), dr.Item("Data_Cad"), dr.Item("Data_Alt"), dr.Item("Telefone"), dr.Item("Celular"), dr.Item("Endereco"), dr.Item("Numero"), dr.Item("Bairro"), dr.Item("Cidade"), dr.Item("Cep"), dr.Item("Estado"), dr.Item("E_mail"), dr.Item("Razao_social"), dr.Item("Atendimento"))
            End While
        End If
    End Sub
    Private Sub Form22_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub
    Private Sub Form22_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from fornecedores"
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
End Class