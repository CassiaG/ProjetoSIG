Imports System.Data.OleDb


Public Class Form26
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader

   
    Private Sub Form26_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from gastolucro"
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
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                DataGridView1.Rows.Add(dr.Item("l_gravacao"), dr.Item("l_ensaio"), dr.Item("l_produtos"), dr.Item("d_luz"), dr.Item("d_agua"), dr.Item("d_tel"), dr.Item("d_produto"), dr.Item("d_salario"), dr.Item("d_transp"), dr.Item("aluguel"), dr.Item("d_outras"), dr.Item("total_ganhos"), dr.Item("total_despesas"), dr.Item("total_lucro"), dr.Item("dia"))
               
            End While
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

   
   
End Class