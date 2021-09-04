Imports System.Data.OleDb
Public Class Form18
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


      

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
      
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Form18_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub

    Private Sub Form18_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "Select * from gravacao"
        Dim cm As New OleDb.OleDbCommand(sql, DBCon)

        Try
            DBCon.Open()
            dr = cm.ExecuteReader
            If dr.HasRows Then
                While dr.Read
                    DataGridView1.Rows.Add(dr.Item("cod_cliente"), dr.Item("nome_banda"), dr.Item("inicio"), dr.Item("termino"), dr.Item("dia"), dr.Item("tempo_g"), dr.Item("valor_g"), dr.Item("valor_f"))
                    '    DataGridView1.Rows.Clear()

                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
        DBCon.Close()

    End Sub
End Class