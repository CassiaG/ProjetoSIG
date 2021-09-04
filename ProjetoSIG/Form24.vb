Imports System.Data.OleDb
Public Class Form24
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("Codigo"), dr.Item("Nome"), dr.Item("Cidade"), dr.Item("Estado"), dr.Item("Telefone"), dr.Item("Celular"), dr.Item("Email"), dr.Item("Agendamento"))
            End While
        End If
    End Sub
    Private Sub Form24_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub
    Private Sub Form24_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from clientes"
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