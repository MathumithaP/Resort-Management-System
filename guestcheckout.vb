Imports System.Data
Public Class guestcheckout
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub guestcheckout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Me.TextBox2.Text = rs(2).Value
        Me.TextBox3.Text = rs(3).Value
        Me.TextBox6.Text = rs(4).Value
        Me.TextBox4.Text = rs(5).Value
        Me.TextBox5.Text = rs(6).Value
        Me.TextBox12.Text = rs(7).Value
        Me.TextBox10.Text = rs(8).Value
        Me.TextBox9.Text = rs(10).Value
        Me.TextBox7.Text = rs(11).Value
        Me.TextBox8.Text = rs(12).Value
        Me.TextBox11.Text = rs(13).Value
        rs.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        rs.Open("select * from checkout", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs(0).Value = Me.ComboBox1.Text
        rs(1).Value = Me.TextBox1.Text
        rs(2).Value = Me.TextBox2.Text
        rs(3).Value = Me.TextBox3.Text
        rs(4).Value = Me.TextBox6.Text
        rs(5).Value = Me.TextBox4.Text
        rs(6).Value = Me.TextBox5.Text
        rs(7).Value = Me.TextBox12.Text
        rs(8).Value = Me.TextBox10.Text
        rs(9).Value = Me.TextBox9.Text
        rs(10).Value = Me.TextBox7.Text
        rs(11).Value = Me.TextBox8.Text
        rs(12).Value = Me.TextBox11.Text
        rs.Update()
        rs.Close()

        If Me.TextBox10.Text = "Single" Then
            rs.Open("select * from single1 where roomno = " & Me.textbox9.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "yes"
            rs.Update()
        ElseIf Me.TextBox10.Text = "Double" Then
            rs.Open("select * from double1 where roomno = " & Me.textbox9.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "yes"
            rs.Update()
        ElseIf Me.TextBox10.Text = "Deluxe" Then
            rs.Open("select * from deluxe where roomno = " & Me.TextBox9.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "yes"
            rs.Update()
        ElseIf Me.TextBox10.Text = "Suite" Then
            rs.Open("select * from suite where roomno = " & Me.TextBox9.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "yes"
            rs.Update()
        End If
        rs.Close()

        rs.Open("select * from guestin where guestid = " & Me.ComboBox1.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.Delete()
        rs.Close()
        MsgBox("Guest Details check out")
        Me.ComboBox1.Text = ""
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox11.Text = ""


        Me.ComboBox1.Items.Clear()
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            Me.ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class