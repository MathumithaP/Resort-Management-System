Imports System.Windows.Forms

Public Class home


    Private Sub GuestInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestInToolStripMenuItem.Click
        guestcheckin.Show()
        Me.Hide()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub GuestOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestOutToolStripMenuItem.Click
        guestcheckout.Show()
        Me.Hide()
    End Sub

    Private Sub EditGuestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditGuestToolStripMenuItem.Click
        editcheckin.Show()
        Me.Hide()
    End Sub

  
    Private Sub EmployeeDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployeeDetailsToolStripMenuItem.Click
        employee.Show()
        Me.Hide()
    End Sub

    Private Sub GuestCheckInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestCheckInToolStripMenuItem.Click
        searchcheckin.Show()
        Me.Hide()
    End Sub

  

    Private Sub home_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GuestCheckOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestCheckOutToolStripMenuItem.Click
        searchcheckout.Show()
        Me.Hide()
    End Sub

    Private Sub StatusOfHotelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusOfHotelToolStripMenuItem.Click
        status1.Show()
        Me.Hide()
    End Sub

    Private Sub RoomChargeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoomChargeToolStripMenuItem.Click
        charge.Show()

        Me.Hide()
    End Sub

    Private Sub ChangeRoomChargeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeRoomChargeToolStripMenuItem.Click
        roomcharge.Show()
        Me.Hide()
    End Sub

    Private Sub EmpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpToolStripMenuItem.Click
        salary.Show()
        Me.Hide()
    End Sub

    Private Sub GuestPaymentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestPaymentToolStripMenuItem.Click
        guestpayment.Show()
        Me.Hide()
    End Sub

    Private Sub GuestReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GuestReportToolStripMenuItem.Click
        breakfast.Show()
        Me.Hide()
    End Sub

    Private Sub EmployeeReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployeeReportToolStripMenuItem.Click
        lunch.Show()
        Me.Hide()
    End Sub

    Private Sub DinnerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DinnerToolStripMenuItem.Click
        dinner.Show()
        Me.Hide()
    End Sub

    Private Sub ChangePasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangePasswordToolStripMenuItem.Click
        changepass.Show()
        Me.Hide()
    End Sub

    Private Sub LogoffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoffToolStripMenuItem.Click
        Me.EmployeeDetailsToolStripMenuItem.Enabled = True
        Me.PaymentToolStripMenuItem.Enabled = True
        Me.GuestInToolStripMenuItem.Enabled = True
        Me.GuestOutToolStripMenuItem.Enabled = True
        Me.EditGuestToolStripMenuItem.Enabled = True
        Me.UtilitiesToolStripMenuItem.Enabled = True
        Form1.Show()
        Me.Hide()
    End Sub

End Class
