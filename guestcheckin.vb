Imports System.Data

Public Class guestcheckin
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim a As Integer

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub guestcheckin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.AddNew()
        rs(0).Value = Me.TextBox6.Text
        rs(1).Value = Me.TextBox1.Text
        rs(2).Value = Me.TextBox2.Text
        rs(3).Value = Me.TextBox3.Text
        rs(4).Value = Me.ComboBox1.Text
        rs(5).Value = Me.TextBox4.Text
        rs(6).Value = Me.TextBox5.Text
        rs(7).Value = Me.DateTimePicker1.Value
        rs(8).Value = Me.ComboBox2.Text
        If Me.RadioButton1.Checked = True Then
            rs(9).Value = Me.RadioButton1.Text
        Else
            rs(9).Value = Me.RadioButton2.Text
        End If
        rs(10).Value = Me.ComboBox3.Text
        rs(11).Value = Me.TextBox10.Text
        rs(12).Value = Me.ComboBox4.Text
        rs(13).Value = Me.TextBox9.Text
        rs(14).Value = Me.TextBox7.Text
        rs(15).Value = Me.TextBox8.Text
        rs.Update()
        rs.Close()
        If Me.ComboBox2.SelectedIndex = 0 Then
            rs.Open("select * from single1 where roomno = " & Me.ComboBox3.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "no"
            rs.Update()
        ElseIf Me.ComboBox2.SelectedIndex = 1 Then
            rs.Open("select * from double1 where roomno = " & Me.ComboBox3.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "no"
            rs.Update()
        ElseIf Me.ComboBox2.SelectedIndex = 2 Then
            rs.Open("select * from deluxe where roomno = " & Me.ComboBox3.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "no"
            rs.Update()
        ElseIf Me.ComboBox2.SelectedIndex = 3 Then
            rs.Open("select * from suite where roomno = " & Me.ComboBox3.Text & "", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = "no"
            rs.Update()
        End If
        rs.Close()
        MsgBox("added successfully")
        Me.TextBox6.Text = ""
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""
        Me.ComboBox3.Text = ""
        Me.ComboBox4.Text = ""
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Me.ComboBox3.Items.Clear()
        Me.ComboBox3.Text = ""
        If Me.ComboBox2.SelectedIndex = 0 Then
            rs.Open("select * from single1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not rs.EOF)
                If rs(1).Value = "yes" Then
                    Me.ComboBox3.Items.Add(rs(0).Value)
                End If
                rs.MoveNext()
            End While
        ElseIf Me.ComboBox2.SelectedIndex = 1 Then
            rs.Open("select * from double1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not rs.EOF)
                If rs(1).Value = "yes" Then
                    Me.ComboBox3.Items.Add(rs(0).Value)
                End If
                rs.MoveNext()
            End While
        ElseIf Me.ComboBox2.SelectedIndex = 2 Then
            rs.Open("select * from deluxe", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not rs.EOF)
                If rs(1).Value = "yes" Then
                    Me.ComboBox3.Items.Add(rs(0).Value)
                End If
                rs.MoveNext()
            End While
        ElseIf Me.ComboBox2.SelectedIndex = 3 Then
            rs.Open("select * from suite", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not rs.EOF)
                If rs(1).Value = "yes" Then
                    Me.ComboBox3.Items.Add(rs(0).Value)
                End If
                rs.MoveNext()
            End While
        End If
        rs.Close()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        rs.Open("select * from roomcharge", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Me.ComboBox2.SelectedIndex = 0 Then

            Me.TextBox10.Text = rs(0).Value
        ElseIf Me.ComboBox2.SelectedIndex = 1 Then
            Me.TextBox10.Text = rs(1).Value
        ElseIf Me.ComboBox2.SelectedIndex = 2 Then
            Me.TextBox10.Text = rs(2).Value
        ElseIf Me.ComboBox2.SelectedIndex = 3 Then
            Me.TextBox10.Text = rs(3).Value
        End If
        rs.Close()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Me.TextBox9.Text = Val(Me.TextBox10.Text) * Val(Me.ComboBox4.Text)
        Me.TextBox9.Enabled = False

    End Sub

    Private Sub TextBox8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.GotFocus
        Me.TextBox8.Text = Val(Me.TextBox9.Text) - Val(Me.TextBox7.Text)
        Me.TextBox8.Enabled = False
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            MsgBox("Please Enter the Digits")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        rs.Open("select * from guestin", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        rs.MoveLast()
        a = rs(0).Value
        a = a + 1
        Me.TextBox6.Text = a
        Me.TextBox6.Enabled = False
        rs.Close()
    End Sub

End Class