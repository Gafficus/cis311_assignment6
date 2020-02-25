

#Disable Warning IDE1006 ' Naming Styles
Public Class frmMDIParent
#Enable Warning IDE1006 ' Naming Styles

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles mnuFileNew.Click
        ' Create a new instance of the child form.
        ' Make it a child of this MDI form before showing it.
        Dim ChildForm As New frmChild With {
            .MdiParent = Me
        }

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuWindowsCascade.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuWindowsTileVertical.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuWindowsTileHorizantal.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuWindowsCloseAll.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        frmAbout.ShowDialog()
    End Sub
End Class
