Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine

Public Class empsalaryreport
    Private Sub empsalaryreport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cry As New ReportDocument
        cry.Load("E:\resortnet\resortnet\CrystalReport1.rpt")
        Me.CrystalReportViewer1.ReportSource = cry
        Me.CrystalReportViewer1.Refresh()
        Me.CrystalReportViewer1.Visible = True
    End Sub
End Class