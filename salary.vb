Imports System.Data

Public Class salary
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub salary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")

        Me.ComboBox4.Items.Clear()
        rs.Open("select * from employee", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

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
        Me.TextBox3.Text = rs(5).Value
        Me.TextBox4.Text = rs(6).Value
        Me.TextBox5.Text = rs(9).Value
        rs.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        rs.Open("select * from payment", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs(0).Value = Me.ComboBox4.Text
        rs(1).Value = Me.TextBox2.Text
        rs(2).Value = Me.TextBox3.Text
        rs(3).Value = Me.TextBox4.Text
        rs(4).Value = Me.TextBox5.Text
        rs(5).Value = Me.DateTimePicker1.Value
        rs.Update()
        rs.Close()
        MsgBox("Paymet Details Get added")
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class