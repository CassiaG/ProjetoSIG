Imports System.Data.OleDb
Public Class Form20
    Dim ConString As String = "provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" + Application.StartupPath & "\login.accdb"
    Dim DBCon As New OleDb.OleDbConnection(constring)
    Dim dr As OleDb.OleDbDataReader
    Public Sub exibir()
        If dr.HasRows Then
            While dr.Read
                Me.DataGridView1.Rows.Add(dr.Item("cod_produto"), dr.Item("nomeE"), dr.Item("nomeP"), dr.Item("preco"), dr.Item("quantidade"), dr.Item("Quantidade_final"))
            End While
        End If
    End Sub

    Private Sub Form20_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TextBox1.Visible = False
        Dim sql As String = "Select * from produtos"
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
    Private Sub Form20_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        DataGridView1.Rows.Clear()
        Me.Close()

    End Sub
    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'Try
        '    TextBox1.Text = DataGridView1.CurrentRow.Cells(0).Value
        '    'PictureBox1.Load(Application.StartupPath + "\" + DataGridView1.CurrentRow.Cells(3).Value)
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
        ''If (Me.Close = DataGridView1.Rows.Clear()
        'End If
    End Sub
    

    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class