Imports System.Data

Public Class dinner
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub dinner_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            Me.ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        rs.Open("select * from guestin where guestid = " & Me.ComboBox1.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.TextBox1.Text = rs(1).Value
        Me.TextBox2.Text = rs(8).Value
        Me.TextBox3.Text = rs(10).Value
        rs.Close()
    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        Me.TextBox6.Text = Val(Me.ComboBox2.Text) * Val(Me.TextBox5.Text)
        Me.TextBox6.Enabled = False
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        rs.Open("select * from dinner", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs(0).Value = Me.ComboBox1.Text
        rs(1).Value = Me.TextBox1.Text
        rs(2).Value = Me.TextBox3.Text
        rs(3).Value = Me.TextBox4.Text
        rs(4).Value = Me.ComboBox2.Text
        rs(5).Value = Me.TextBox6.Text
        rs(6).Value = Me.DateTimePicker1.Value
        rs.Update()
        rs.Close()
        MsgBox("Dinner Order Placed")
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.TextBox4.Text = "Poori Sets"
        Me.TextBox5.Text = 25

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.TextBox4.Text = "Barota Sets"
        Me.TextBox5.Text = 35

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.TextBox4.Text = "Special Dhosai"
        Me.TextBox5.Text = 40

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.TextBox4.Text = "Badham Milk"
        Me.TextBox5.Text = 45

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        End
    End Sub
End Class