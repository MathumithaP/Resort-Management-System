Imports System.Data

Public Class editcheckin
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub editcheckin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            Me.ComboBox5.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        rs.Open("select * from guestin where guestid = " & Me.ComboBox5.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.TextBox1.Text = rs(1).Value
        Me.TextBox2.Text = rs(2).Value
        Me.TextBox3.Text = rs(3).Value
        Me.ComboBox1.Text = rs(4).Value
        Me.TextBox4.Text = rs(5).Value
        Me.TextBox5.Text = rs(6).Value
        Me.DateTimePicker1.Value = rs(7).Value
        Me.ComboBox2.Text = rs(8).Value
        Me.ComboBox3.Text = rs(10).Value
        Me.TextBox10.Text = rs(11).Value
        Me.ComboBox4.Text = rs(12).Value
        Me.TextBox9.Text = rs(13).Value
        Me.TextBox8.Text = rs(14).Value
        rs.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        rs.Open("select * from guestin where guestid = " & Me.ComboBox5.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs(1).Value = Me.TextBox1.Text
        rs(2).Value = Me.TextBox2.Text
        rs(3).Value = Me.TextBox3.Text
        rs(4).Value = Me.ComboBox1.Text
        rs(5).Value = Me.TextBox4.Text
        rs(6).Value = Me.TextBox5.Text
        rs(7).Value = Me.DateTimePicker1.Value
        rs(8).Value = Me.ComboBox2.Text
        rs(10).Value = Me.ComboBox3.Text
        rs(11).Value = Me.TextBox10.Text
        rs(12).Value = Me.ComboBox4.Text
        rs(13).Value = Me.TextBox9.Text
        rs(14).Value = Me.TextBox8.Text
        rs.Update()
        rs.Close()
        MsgBox("updated successfully")
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""
        Me.ComboBox3.Text = ""
        Me.ComboBox5.Text = ""
        Me.ComboBox4.Text = ""
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please Enter the Digits")
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

    End Sub
End Class