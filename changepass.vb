Imports System.Data

Public Class changepass
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.TextBox1.Text = Me.TextBox2.Text Then
            rs.Open("select * from login where username = '" & Me.Label3.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            rs(1).Value = Me.TextBox1.Text
            rs.Update()
            rs.Close()
            MsgBox("New Password got changed")
        Else
            MsgBox("Your password not match")
        End If
    End Sub

    Private Sub changepass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class