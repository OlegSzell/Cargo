Imports System.IO
Public Class Древо
    Dim Path As String


    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim Files() As String = IO.Directory.GetFiles(IO.Path.GetDirectoryName(Path) & "\" & e.Node.FullPath, "*.*", SearchOption.TopDirectoryOnly)
        ListView1.Items.Clear()
        For Each File As String In Files
            ListView1.Items.Add(IO.Path.GetFileName(File)).Tag = File
            ListView1.Items(ListView1.Items.Count - 1).ImageIndex = 1
        Next


        Dim Files2() As String = IO.Directory.GetFiles(IO.Path.GetDirectoryName(Path) & "\" & e.Node.FullPath, "*.*", SearchOption.TopDirectoryOnly)
        ListBox1.Items.Clear()

        For i As Integer = 0 To Files2.Count - 1
            ListBox1.Items.Add(IO.Path.GetFileName(Files(i)))
        Next


    End Sub

    Private Sub Древо_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim chilieNode As TreeNode = New TreeNode
        TreeView1.Nodes.Add("Кадры")
        Path = "U:\Офис\Финансовый\6. Бух.услуги\Кадры"
        Search(Path, TreeView1.Nodes(0))
        TreeView1.Nodes(0).Expand()
        'chilieNode = (U:\Офис\Финансовый\6. Бух.услуги\Кадры\)
        'TreeView1.Nodes(0).Add("U:\Офис\Финансовый\6. Бух.услуги\Кадры\")
    End Sub
    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDoubleClick
        If ListView1.SelectedItems.Count > 0 Then
            MsgBox(ListView1.SelectedItems(0).Tag)
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FBD As New FolderBrowserDialog

        If FBD.ShowDialog = Windows.Forms.DialogResult.OK Then
            Path = FBD.SelectedPath
            TreeView1.Nodes.Add(IO.Path.GetFileName(Path))
            Search(Path, TreeView1.Nodes(0))
            TreeView1.Nodes(0).Expand()
        End If
    End Sub
    Sub Search(ByVal Fol As String, ByVal Node As TreeNode)
        For Each S As String In IO.Directory.GetDirectories(Fol, "*.*", SearchOption.TopDirectoryOnly)
            Dim TmpNode As New TreeNode(IO.Path.GetFileName(S))
            TmpNode.ImageIndex = 0
            Node.Nodes.Add(TmpNode)
            Try
                Search(S, TmpNode)
            Catch ex As Exception
            End Try
        Next
    End Sub
End Class