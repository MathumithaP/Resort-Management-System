Imports System.Data

Public Class status1
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    Private Sub status1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub status1_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As Integer
        a = 0
        con.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\resortnet\resort1.mdb")
        rs.Open("select * from single1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            a = a + 1
            rs.MoveNext()
        End While
        rs.Close()
        Me.Label10.Text = a
        Dim b As Integer
        b = 0
        rs.Open("select * from double1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            b = b + 1
            rs.MoveNext()
        End While
        rs.Close()
        Me.Label11.Text = b
        Dim c As Integer
        c = 0
        rs.Open("select * from suite", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            c = c + 1
            rs.MoveNext()
        End While
        rs.Close()
        Me.Label12.Text = c

        Dim d As Integer
        d = 0
        rs.Open("select * from deluxe", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            d = d + 1
            rs.MoveNext()
        End While
        rs.Close()
        Me.Label13.Text = d

        Dim a1 As Integer
        a1 = 0
        rs.Open("select * from single1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "no" Then


                a1 = a1 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label14.Text = a1



        Dim b1 As Integer
        b1 = 0
        rs.Open("select * from double1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "no" Then
                b1 = b1 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label15.Text = b1


        Dim c1 As Integer
        c1 = 0
        rs.Open("select * from suite", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "no" Then
                c1 = c1 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label16.Text = c1


        Dim d1 As Integer
        d1 = 0
        rs.Open("select * from deluxe", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "no" Then


                d1 = d1 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label17.Text = d1



        Dim a2 As Integer
        a2 = 0
        rs.Open("select * from single1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "yes" Then


                a2 = a2 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label18.Text = a2



        Dim b2 As Integer
        b2 = 0
        rs.Open("select * from double1", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "yes" Then
                b2 = b2 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label19.Text = b2


        Dim c2 As Integer
        c2 = 0
        rs.Open("select * from suite", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "yes" Then
                c2 = c2 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label20.Text = c2


        Dim d2 As Integer
        d2 = 0
        rs.Open("select * from deluxe", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While (Not rs.EOF)
            If rs(1).Value = "yes" Then


                d2 = d2 + 1
            End If
            rs.MoveNext()

        End While
        rs.Close()
        Me.Label21.Text = d2
    End Sub



    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        home.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub
End Class