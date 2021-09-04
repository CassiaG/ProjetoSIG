Public Class Form27

    Private Sub ClientesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClientesToolStripMenuItem.Click
        Form2.ShowDialog()
    End Sub

    Private Sub ProdutosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdutosToolStripMenuItem.Click
        Form5.ShowDialog()
    End Sub

    Private Sub ServiçosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ServiçosToolStripMenuItem.Click
        Form6.ShowDialog()
    End Sub

    Private Sub GravaçãoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GravaçãoToolStripMenuItem.Click
        Form7.ShowDialog()
    End Sub

    Private Sub EnsaioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnsaioToolStripMenuItem.Click
        Form8.ShowDialog()
    End Sub

    Private Sub AgendamentoDeShowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgendamentoDeShowToolStripMenuItem.Click
        Form17.ShowDialog()
    End Sub

    Private Sub FornecedoresToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FornecedoresToolStripMenuItem1.Click
        Form12.ShowDialog()

    End Sub

    Private Sub ProdutosToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProdutosToolStripMenuItem1.Click
        Form14.ShowDialog()
    End Sub

    Private Sub GravaçãoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GravaçãoToolStripMenuItem1.Click
        Form15.ShowDialog()
    End Sub

    Private Sub EnsaioToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnsaioToolStripMenuItem1.Click
        Form16.ShowDialog()
    End Sub

    Private Sub VisualizaçãoDeServiçosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisualizaçãoDeServiçosToolStripMenuItem.Click
        Form13.ShowDialog()
    End Sub

    Private Sub AjudaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AjudaToolStripMenuItem.Click

    End Sub

    Private Sub SobreToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SobreToolStripMenuItem1.Click
        MessageBox.Show("Este é um projeto dos Terceiros Modulos de Administração e Informática, idealizado pela Etec Jk")
    End Sub

    Private Sub SairToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SairToolStripMenuItem.Click
        Form9.Show()
    End Sub
End Class