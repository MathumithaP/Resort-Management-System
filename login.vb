Imports System.Data

Public Class login
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
     
        rs.Open("select * from login where username = '" & Me.ComboBox1.Text & "' and password = '" & Me.TextBox1.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If rs.EOF = True Then

            MsgBox("Invalid user/password")
        Else
            home.Label13.Text = Me.ComboBox1.Text
            changepass.Label3.Text = Me.ComboBox1.Text
            changepass.Label4.Text = Me.TextBox1.Text
            home.Show()
            Me.Hide()
        End If
        If Me.ComboBox1.Text = "Personnel" Then
            home.EmployeeDetailsToolStripMenuItem.Enabled = False
            home.PaymentToolStripMenuItem.Enabled = False
            home.GuestInToolStripMenuItem.Enabled = False
            home.GuestOutToolStripMenuItem.Enabled = False
            home.EditGuestToolStripMenuItem.Enabled = False
            home.UtilitiesToolStripMenuItem.Enabled = False
        ElseIf Me.ComboBox1.Text = "Front Officer" Then
            home.EmployeeDetailsToolStripMenuItem.Enabled = False
        End If
        rs.Close()
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        Me.ComboBox1.Text = ""
        Me.TextBox1.Text = ""
    End Sub

    
    
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class