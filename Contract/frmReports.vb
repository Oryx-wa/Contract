Imports System.Configuration
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrytsalReportsDemo
Imports SBO.SboAddOnBase

Public Class frmReports

    Private WithEvents rptDoc As New ReportDocument
    Private tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
    Private tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
    Friend parentaddon As SBOAddOn, Level As Integer = 0
    Private Report As ReportDocument

    Public Sub New(ByVal cryRpt As ReportDocument)
        MyBase.New()
        InitializeComponent()
        rptDoc = cryRpt
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Friend Sub showreport(ByVal cReportName As String, ByVal ds As DataSet, _
    ByVal displayGroupTree As Boolean, ByVal displayToolbar As Boolean, _
    ByVal pvCollection As CrystalDecisions.Shared.ParameterValues, ByRef paramfields As CrystalDecisions.Shared.ParameterFields, _
    Optional ByVal NewTitle As String = Nothing, Optional ByVal strFilter As String = Nothing, _
    Optional ByVal groupfield As Integer = -1)
        Try
            DbConnectionInfo.SetConnectionString(ConfigurationManager.ConnectionStrings("SQLConnStr").ConnectionString.ToString)

            Dim logOnInfo As TableLogOnInfo = New TableLogOnInfo
            Dim ConnectionInfo As ConnectionInfo = New ConnectionInfo

            ' Load the report
            rptDoc.Load(cReportName)
            ' Set the connection information for all the tables used in the report
            ' Leave UserID and Password blank for trusted connection
            For Each table In rptDoc.Database.Tables
                logOnInfo = table.LogOnInfo
                ConnectionInfo = logOnInfo.ConnectionInfo
                'Set the Connection parameters.
                ConnectionInfo.DatabaseName = DbConnectionInfo.InitialCatalog
                ConnectionInfo.ServerName = DbConnectionInfo.ServerName
                If (Not DbConnectionInfo.UseIntegratedSecurity) Then
                    ConnectionInfo.Password = DbConnectionInfo.Password
                    ConnectionInfo.UserID = DbConnectionInfo.UserName
                Else
                    ConnectionInfo.IntegratedSecurity = True
                End If

                table.ApplyLogOnInfo(logOnInfo)

            Next table
            ' Test the connection
            If Not rptDoc.Database.Tables.Item(0).TestConnectivity() Then Exit Sub

            If Not pvCollection Is Nothing Then rptDoc.DataDefinition.ParameterFields(0).ApplyCurrentValues(pvCollection)
            'Set the title
            If Not NewTitle Is Nothing Then rptDoc.SummaryInfo.ReportTitle = NewTitle
            ' Set the report source for the crystal reports 
            ' viewer to the report instance.
            If Not paramfields Is Nothing Then crvBasic.ParameterFieldInfo = paramfields

            ' Set the selection criteria
            If Not strFilter Is Nothing Then
                rptDoc.RecordSelectionFormula = strFilter
            Else
                rptDoc.RecordSelectionFormula = ""
            End If

            ' Set the group criteria
            If Not groupfield = -1 Then
                Dim fieldDef As FieldDefinition
                fieldDef = rptDoc.Database.Tables.Item(0).Fields.Item(groupfield)
                rptDoc.DataDefinition.Groups.Item(0).ConditionField = fieldDef
            End If
            rptDoc.Refresh()
            crvBasic.ReportSource = rptDoc
            crvBasic.DisplayGroupTree = displayGroupTree
            crvBasic.DisplayToolbar = displayToolbar
        Catch Exp As LoadSaveReportException
            MsgBox("Incorrect path for loading report.", _
                    MsgBoxStyle.Critical, "Load Report Error")
        Catch Exp As Exception
            MsgBox(Exp.Message, MsgBoxStyle.Critical, "General Error")
        End Try


    End Sub
    Public Sub setReport(ByVal ds As DataSet, ByVal xReport As ReportDocument, ByVal caption As String, ByVal displayGroupTree As Boolean, _
               Optional ByRef paramfields As CrystalDecisions.Shared.ParameterFields = Nothing, _
               Optional ByVal paperOrientation As Integer = 0, Optional ByVal papersize As Integer = 0)

        'oRpt = New CrystalReport1
        rptDoc = xReport
        rptDoc.SetDataSource(ds)
        rptDoc.PrintOptions.PaperSize = papersize
        rptDoc.PrintOptions.PaperOrientation = paperOrientation

        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oCompanyAdminInfo As SAPbobsCOM.AdminInfo

        oCompanyService = parentaddon.SboCompany.GetCompanyService
        oCompanyAdminInfo = oCompanyService.GetAdminInfo

        rptDoc.SummaryInfo.ReportTitle = oCompanyAdminInfo.CompanyName
        rptDoc.SummaryInfo.ReportComments = oCompanyAdminInfo.Address
        rptDoc.SummaryInfo.ReportAuthor = caption

        Me.crvBasic.DisplayGroupTree = displayGroupTree
        If Not paramfields Is Nothing Then crvBasic.ParameterFieldInfo = paramfields
        Me.crvBasic.ReportSource = rptDoc


    End Sub

    

    Private Sub frmReports_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        System.Windows.Forms.Application.ExitThread()
    End Sub

    Private Sub frmReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DbConnectionInfo.SetConnectionString(ConfigurationManager.ConnectionStrings("SQLConnStr").ConnectionString.ToString)

        Dim logOnInfo As TableLogOnInfo = New TableLogOnInfo
        Dim ConnectionInfo As ConnectionInfo = New ConnectionInfo


        For Each table As Table In rptDoc.Database.Tables

            logOnInfo = table.LogOnInfo
            ConnectionInfo = logOnInfo.ConnectionInfo
            'Set the Connection parameters.
            ConnectionInfo.DatabaseName = DbConnectionInfo.InitialCatalog
            ConnectionInfo.ServerName = DbConnectionInfo.ServerName
            If (Not DbConnectionInfo.UseIntegratedSecurity) Then
                ConnectionInfo.Password = DbConnectionInfo.Password
                ConnectionInfo.UserID = DbConnectionInfo.UserName
            Else
                ConnectionInfo.IntegratedSecurity = True
            End If

            table.ApplyLogOnInfo(logOnInfo)
        Next

        crvBasic.ReportSource = rptDoc
        crvBasic.Refresh()
    End Sub
End Class