Imports System.Data

Public Class employee
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub employee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Me.ComboBox4.Enabled = False Then
            rs.Open("select * from employee", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            rs.AddNew()
            rs(0).Value = Me.ComboBox4.Text
            rs(1).Value = Me.TextBox2.Text
            rs(2).Value = Me.TextBox3.Text
            rs(3).Value = Me.TextBox4.Text
            rs(4).Value = Me.TextBox5.Text
            rs(5).Value = Me.ComboBox1.Text
            rs(6).Value = Me.DateTimePicker1.Value
            rs(7).Value = Me.ComboBox2.Text
            rs(8).Value = Me.ComboBox3.Text
            rs(9).Value = Me.TextBox6.Text
            rs.Update()
            rs.Close()
            MsgBox("New Employee Details added successfully")
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.ComboBox1.Text = ""
            Me.ComboBox2.Text = ""
            Me.ComboBox3.Text = ""
            Me.ComboBox4.Text = ""
        Else
            rs.Open("select * from employee where eid = " & Me.ComboBox4.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = Me.TextBox2.Text
            rs(2).Value = Me.TextBox3.Text
            rs(3).Value = Me.TextBox4.Text
            rs(4).Value = Me.TextBox5.Text
            rs(5).Value = Me.ComboBox1.Text
            rs(6).Value = Me.DateTimePicker1.Value
            rs(7).Value = Me.ComboBox2.Text
            rs(8).Value = Me.ComboBox3.Text
            rs(9).Value = Me.TextBox6.Text
            rs.Update()
            rs.Close()
            MsgBox("Updated Successfully")
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.ComboBox1.Text = ""
            Me.ComboBox2.Text = ""
            Me.ComboBox3.Text = ""
            Me.ComboBox4.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a As Integer
        rs.Open("select * from employee", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.MoveLast()
        a = rs(0).Value
        a = a + 1
        Me.ComboBox4.Text = a
        Me.ComboBox4.Enabled = False
        rs.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.ComboBox4.Items.Clear()
        rs.Open("select * from employee", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.ComboBox4.Enabled = True
        Me.ComboBox4.Text = ""
        While (Not rs.EOF)
            Me.ComboBox4.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        rs.Open("select * from employee where eid = " & Me.ComboBox4.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.TextBox2.Text = rs(1).Value
        Me.TextBox3.Text = rs(2).Value
        Me.TextBox4.Text = rs(3).Value
        Me.TextBox5.Text = rs(4).Value
        Me.ComboBox1.Text = rs(5).Value
        Me.DateTimePicker1.Value = rs(6).Value
        Me.ComboBox2.Text = rs(7).Value
        Me.ComboBox3.Text = rs(8).Value
        Me.TextBox6.Text = rs(9).Value
        rs.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        rs.Open("select * from employee where eid = " & Me.ComboBox4.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.Delete()
        rs.Close()
        MsgBox("Deleted Successfully")
        rs.Open("select * from employee", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Me.ComboBox4.Enabled = True
        Me.ComboBox4.Text = ""
        Me.ComboBox4.Items.Clear()
        While (Not rs.EOF)
            Me.ComboBox4.Items.Add(rs(0).Value)
            rs.MoveNext()
        End While
        rs.Close()
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""
        Me.ComboBox3.Text = ""
        Me.ComboBox4.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
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

    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please Enter the Digits")
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        empsalaryreport.Show()
    End Sub
End Class