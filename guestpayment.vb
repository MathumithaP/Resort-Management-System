Imports System.Data

Public Class guestpayment
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub guestpayment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            Me.ComboBox4.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        rs.Open("select * from guestin where guestid = " & Me.ComboBox4.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.TextBox2.Text = rs(1).Value
        Me.TextBox3.Text = rs(2).Value
        Me.TextBox4.Text = rs(5).Value
        Me.TextBox5.Text = rs(7).Value
        Me.TextBox9.Text = rs(8).Value
        Me.TextBox8.Text = rs(11).Value
        Me.TextBox7.Text = rs(12).Value
        Me.TextBox6.Text = rs(13).Value
        Me.TextBox1.Text = rs(15).Value
        rs.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.TextBox11.Text = Val(Me.TextBox1.Text) - Val(Me.TextBox10.Text)
        Me.TextBox11.Enabled = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        rs.Open("select * from guestpayment", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs(0).Value = Me.ComboBox4.Text
        rs(1).Value = Me.TextBox2.Text
        rs(2).Value = Me.TextBox6.Text
        rs(3).Value = Me.TextBox11.Text
        rs(4).Value = Me.DateTimePicker1.Value
        rs.Update()
        rs.Close()
        rs.Open("select * from guestin where guestid = " & Me.ComboBox4.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs(15).Value = Me.TextBox11.Text
        rs.Update()
        rs.Close()
        MsgBox("Details get updated")
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox1.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox10.KeyPress
        If Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please Enter the Digits")
        End If
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        paymentreport.Show()
    End Sub
End Class