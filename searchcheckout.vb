Imports System.Data
Public Class searchcheckout
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Private Sub searchcheckout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        rs.Open("select * from checkout", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            Me.ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        rs.Open("select * from checkout where gid = " & Me.ComboBox1.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.Label8.Text = rs(1).Value
        Me.Label9.Text = rs(2).Value
        Me.Label12.Text = rs(3).Value
        Me.Label18.Text = rs(4).Value
        Me.Label19.Text = rs(5).Value
        Me.Label20.Text = rs(6).Value
        Me.Label21.Text = rs(7).Value
        Me.Label22.Text = rs(8).Value
        Me.Label23.Text = rs(9).Value
        Me.Label24.Text = rs(10).Value
        Me.Label25.Text = rs(11).Value
        Me.Label26.Text = rs(12).Value
        rs.Close()
        Me.GroupBox1.Visible = True
        Me.GroupBox2.Visible = True
        Me.Panel2.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class