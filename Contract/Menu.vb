Imports SAPbouiCOM.Framework
Imports System.IO
Imports SBO.SboAddOnBase

Imports System
Imports System.Xml
Imports System.Xml.XPath
Imports System.Windows.Forms
Imports System.Activator
Imports System.Runtime.Remoting
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine


Public Class ContractAddon
    Inherits SboAddon
    Private RunApp As Boolean = True
    Public Sub Main()


    End Sub


    Public Sub New(ByVal StartUpPath As String, ByVal AddonName As String, ByRef pbo_RunApplication As Boolean)

        MyBase.New(StartUpPath, AddonName)
        m_Namespace = "Contract"
        m_AssemblyName = "Contract"
        TablePrefix = "OWA_PQ"
        PermissionPrefix = "OWA_PQ"
        MenuXMLFileName = "Menus.xml"
        'm_MenuImageFile = "truck.jpg"
        If IsNothing(m_SboApplication) Then
            pbo_RunApplication = False
            Exit Sub
        Else

            If Not initialise() Then
                pbo_RunApplication = False
                Exit Sub
            End If

        End If
        oApp.Run()
        pbo_RunApplication = True
        'Me.setFilters(oFilters)

    End Sub

    Private _dataSet As DataSet
    Private _Report As ReportDocument
    Private _ReportCaption As String
    Private _DisplayGroupTree As Boolean
    Private _ParamFields As CrystalDecisions.Shared.ParameterFields
    Private _PaperOrientation As Integer
    Private _PaperSize As Integer

    Public Function ShowReport(ByVal ds As DataSet, ByVal xReport As ReportDocument, ByVal caption As String, ByVal displayGroupTree As Boolean, _
              Optional ByRef paramfields As CrystalDecisions.Shared.ParameterFields = Nothing, _
              Optional ByVal paperOrientation As Integer = 0, Optional ByVal papersize As Integer = 0) As String

        Dim ShowReportWindowThread As Threading.Thread
        Try
            ShowReportWindowThread = New Threading.Thread(AddressOf ShowReportWindow)
            _dataSet = ds
            _Report = xReport
            _ReportCaption = caption
            _DisplayGroupTree = displayGroupTree
            _ParamFields = paramfields
            _PaperOrientation = paperOrientation
            _PaperSize = papersize

            If ShowReportWindowThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowReportWindowThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowReportWindowThread.Start()
            ElseIf ShowReportWindowThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowReportWindowThread.Start()
                ShowReportWindowThread.Join()

            End If
            While ShowReportWindowThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While

        Catch ex As Exception
            SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return ""

    End Function

    Public Sub ShowReportWindow()
        Dim MyProcs() As System.Diagnostics.Process
        Dim ReportForm As New frmReports(_Report)

        Try

            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oCompanyAdminInfo As SAPbobsCOM.AdminInfo

            oCompanyService = Me.SboCompany.GetCompanyService
            oCompanyAdminInfo = oCompanyService.GetAdminInfo
            With _Report
                .SetDataSource(_dataSet)
                .PrintOptions.PaperSize = _PaperSize
                .PrintOptions.PaperOrientation = _PaperOrientation
                .SummaryInfo.ReportTitle = oCompanyAdminInfo.CompanyName
                .SummaryInfo.ReportComments = oCompanyAdminInfo.Address
                .SummaryInfo.ReportAuthor = _ReportCaption

            End With

            With ReportForm
                .crvBasic.DisplayGroupTree = _DisplayGroupTree
                If Not _ParamFields Is Nothing Then .crvBasic.ParameterFieldInfo = _ParamFields
                .crvBasic.ReportSource = _Report
            End With


            MyProcs = Process.GetProcessesByName("SAP Business One")

            If MyProcs.Length = 1 Then
                For i As Integer = 0 To MyProcs.Length - 1

                    Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                    ReportForm.ShowDialog()

                Next
            End If

        Catch ex As Exception
            SboApplication.StatusBar.SetText(ex.Message)

        Finally
            ReportForm.Dispose()
        End Try

    End Sub


    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class


End Class