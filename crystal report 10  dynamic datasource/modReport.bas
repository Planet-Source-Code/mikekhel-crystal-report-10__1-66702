Attribute VB_Name = "modReport"
Option Explicit
Public crApp As New CRAXDRT.Application
Public crRep As CRAXDRT.Report
Public dbTable As CRAXDRT.DatabaseTable
Public strSelect As String

Function viewReport(ByVal strSql As String, ByVal strReportFile As String)
Set crRep = New CRAXDRT.Report
Set crApp = CreateObject("crystalruntime.application")
Set crRep = crApp.OpenReport(strReportFile)

For Each dbTable In crRep.Database.Tables
 dbTable.SetLogOnInfo "servername", "databasename", "", ""
Next dbTable

crRep.SQLQueryString = strSql
Form1.CRViewer1.ReportSource = crRep
Form1.CRViewer1.viewReport

Set crRep = Nothing
Set crApp = Nothing

End Function


