Imports System.Data.OleDb
Public Class Form21
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("Codigo"), dr.Item("nome_banda"), dr.Item("telefone"), dr.Item("celular"), dr.Item("inicio"), dr.Item("termino"), dr.Item("agendamento"))
            End While
        End If
    End Sub
    Private Sub Form21_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub
    Private Sub Form21_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from show"
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