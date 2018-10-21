Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Public Class paymentreport
    Private Sub paymentreport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cry As New ReportDocument
        cry.Load("E:\resortnet\resortnet\CrystalReport2.rpt")
        Me.CrystalReportViewer1.ReportSource = cry
        Me.CrystalReportViewer1.Refresh()
        Me.CrystalReportViewer1.Visible = True
    End Sub
End Class