Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase
Imports System.IO
Imports System.Windows.Forms

Namespace Contract
    <FormAttribute("Contract.Contract", "SBOForms/Contract.b1f")>
    Friend Class Contract
        Inherits UserFormBaseClass

        Public Sub New()

            LoadFilters()
            GetData()

        End Sub



        Private Sub LoadFilters()
            Try
                'filter for Vendor
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add
                With oCondition
                    .Alias = "CardType"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "S"
                End With
                cflVendor = oForm.ChooseFromLists.Item("cflVend")
                cflVendor.SetConditions(oConditions)


            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
            End Try
        End Sub

        Private Sub GetData()
            Try

                fillCombo("Code", "Name", "@OWA_PQCONTRINSTYPE", txtInsType, , False)
                'fillCombo("Code", "Name", "@OWA_PQCONTRINSTYPE", txtInsType2, , False)
                'fillCombo("Code", "Name", "@OWA_PQCONTRINSTYPE", txtInsType3, , False)
                fillCombo("Code", "Name", "@OWA_PQCONTRAMD", txtAmdType, , False)
                fillCombo("Code", "Name", "@OWA_PQCONTRBU", cboBizUnit, , False)
                fillCombo("Code", "Name", "@OWA_PQCONTRSPONSOR", cboSponsor, , False)
                fillCombo("Code", "Name", "@OWA_PQCONTRSTATUS", cboStatus, , False)

                fillCombo("CurrCode", "CurrName", "OCRN", cboCurr, , False)


            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Contract_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Reset()
        End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "1288", "1289", "1290", "1291"
                        oForm = SBO_Application.Forms.ActiveForm
                        If oForm.TypeEx = "Contract.Contract" Then
                            HandleNavigation()
                        End If

                    Case "1282", "1283"
                        Reset()
                    Case "519"

                        'PrintRequest()
                        'BubbleEvent = False

                End Select
            Else

            End If

        End Sub

        Private Sub HandleNavigation()
            Try

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                oDataSource2 = oForm.DataSources.DBDataSources.Item("OCRD")
                oDataSource3 = oForm.DataSources.DBDataSources.Item("OHEM")

                If getOffset(oDataSource.GetValue("U_ContractorCode", 0), "CardCode", oDataSource2) Then
                    getOffset(oDataSource2.GetValue("CardCode", 0).Trim, "U_ContractorCode", oDataSource)
                End If


                If getOffset(oDataSource.GetValue("U_Employee", 0), "empID", oDataSource3) Then
                    getOffset(oDataSource3.GetValue("empID", 0).Trim, "U_Employee", oDataSource)
                End If

                txtContrCode = oForm.Items.Item("ContrCode").Specific
                getDocuments(txtContrCode.Value)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub Reset()
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
                oDataSource.SetValue("CardName", 0, "")

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub ContractCode_ChooseFromListAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtContrCode.ChooseFromListAfter
            Dim val As String, oRec As SAPbobsCOM.Recordset
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            'checks if the code has already be used in the contract before
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(String.Format("select * from [@OWA_PQCONTRACT] where U_ContractCode = '" & val & "'"))
            If oRec.RecordCount > 0 Then
                SBO_Application.StatusBar.SetText("Contract Code already used. Select another one!!!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
            oDataSource.SetValue("U_ContractCode", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OPRJ")
            If getOffset(val, "PrjCode", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub Contractor_ChooseFromListAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtContr.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
            oDataSource.SetValue("U_ContractorCode", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
            If getOffset(val, "CardCode", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub Employee_ChooseFromListAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtEmp.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
            oDataSource.SetValue("U_Employee", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
            If getOffset(val, "empID", oDataSource) Then
                oDataSource.Offset = 0
                getOffset(CInt(val), "empID", oDataSource)
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                oDataSource2.SetValue("U_EmployeeName", 0, Trim(oDataSource.GetValue("lastName", 0)) & ", " & Trim(oDataSource.GetValue("firstName", 0)))
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End Sub



        Private Sub matAmd_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matAmd.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRAMEND").Clear()
            matAmd = oForm.Items.Item("matAmd").Specific

            If pVal.Row = matAmd.RowCount + 1 Then
                If pVal.Row = 1 Then
                    matAmd.AddRow(1)
                Else
                    matAmd.AddRow(1, matAmd.RowCount)
                End If
                matAmd.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub

        Private Sub matInsr_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matInsr.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRINSUR").Clear()
            matInsr = oForm.Items.Item("matInsr").Specific

            If pVal.Row = matInsr.RowCount + 1 Then
                If pVal.Row = 1 Then
                    matInsr.AddRow(1)
                Else
                    matInsr.AddRow(1, matInsr.RowCount)
                End If
                matInsr.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub

        Private Sub matAttach_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matAttach.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRATTACH").Clear()
            matAttach = oForm.Items.Item("matAttach").Specific

            If pVal.Row = matAttach.RowCount + 1 Then
                If pVal.Row = 1 Then
                    matAttach.AddRow(1)
                Else
                    matAttach.AddRow(1, matAttach.RowCount)
                End If
                matAttach.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub


        Private Sub txtprjcode_ValidateAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtprjcode.ValidateAfter
            ' txtprjcode = oForm.Items.Item("txtprjcode").Specific
            ' getInvoiceAndPayment(txtprjcode.Value)
        End Sub

        Private Sub getDocuments(ByVal ContractCode As String)
            Dim sSQL As String, oRec As SAPbobsCOM.Recordset
            Try
                'get the current interest rate from the header
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                Dim exchRate As Decimal = oDataSource.GetValue("U_ExchangeRate", 0)

                'Invoice
                sSQL = "select cast(docentry as varchar(15)) InvoiceNo,docdate InvoiceSubmitDate,docduedate InvoiceDueDate,DocTotal "
                sSQL += " InvoiceAmount,case DocTotalFC when 0 then DocTotal/exchRate else DocTotalFC end  InvoiceAmountFC,DocCur InvoiceCurr  from opch "
                sSQL += " WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oDataTable = oForm.DataSources.DataTables.Item("DTInv")
                oDataTable.ExecuteQuery(sSQL)

                'Invoice Total
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sSQL = "select sum(DocTotal) InvoiceAmount,case sum(DocTotalFC) when 0 then sum(DocTotal)/exchRate else sum(DocTotalFC) end  "
                sSQL += " InvoiceAmountFC from opch WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oRec.DoQuery(String.Format(sSQL))
                oDataSource.SetValue("U_TotInvAmt", 0, oRec.Fields.Item("InvoiceAmount").Value)
                oDataSource.SetValue("U_TotInvAmtFC", 0, oRec.Fields.Item("InvoiceAmountFC").Value)

                'Order
                sSQL = "select cast(docentry as varchar(15)) OrderNo,docdate OrderSubmitDate,docduedate OrderDueDate,DocTotal "
                sSQL += " OrderAmount,case DocTotalFC when 0 then DocTotal/exchRate else DocTotalFC end OrderAmountFC,DocCur OrderCurr  from opor "
                sSQL += " WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oDataTable = oForm.DataSources.DataTables.Item("DTOdr")
                oDataTable.ExecuteQuery(sSQL)

                'Order Total
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sSQL = "select sum(DocTotal) OrderAmount,case sum(DocTotalFC) when 0 then sum(DocTotal)/exchRate else sum(DocTotalFC) end  "
                sSQL += " OrderAmountFC from opor WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oRec.DoQuery(String.Format(sSQL))
                oDataSource.SetValue("U_TotOrdAmt", 0, oRec.Fields.Item("OrderAmount").Value)
                oDataSource.SetValue("U_TotOrdAmtFC", 0, oRec.Fields.Item("OrderAmountFC").Value)

                'Payment
                sSQL = "select cast(docentry as varchar(15)) PaymentNo,DocTotal PaymentAmount,case DocTotalFC when 0 then DocTotal/exchRate"
                sSQL += " else DocTotalFC end PaymentAmountFC,DocCurr PaymentCurr from ovpm WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oDataTable = oForm.DataSources.DataTables.Item("DTPmt")
                oDataTable.ExecuteQuery(sSQL)

                'Payment Total
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sSQL = "select sum(DocTotal) PaymentAmount,case sum(DocTotalFC) when 0 then sum(DocTotal)/exchRate else sum(DocTotalFC) end"
                sSQL += " PaymentAmountFC from ovpm WHERE U_ContractCode='" + ContractCode + "'"
                sSQL = Replace(sSQL, "exchRate", exchRate.ToString())
                oRec.DoQuery(String.Format(sSQL))
                oDataSource.SetValue("U_TotPmtAmt", 0, oRec.Fields.Item("PaymentAmount").Value)
                oDataSource.SetValue("U_TotPmtAmtFC", 0, oRec.Fields.Item("PaymentAmountFC").Value)

                FormatGrid()

            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                'm_oForm.Freeze(False)
            End Try
        End Sub

        Private Sub FormatGrid()
            Try
                oForm.Freeze(True)

                'Purchase Invoice
                ogrdInv = oForm.Items.Item("grdInv").Specific
                With ogrdInv
                    .Columns.Item(0).TitleObject.Caption = "Invoice #"
                    .Columns.Item(1).TitleObject.Caption = "Invoice Submit Date"
                    .Columns.Item(2).TitleObject.Caption = "Invoice Due Date"
                    .Columns.Item(3).TitleObject.Caption = "Invoice Amount"
                    .Columns.Item(4).TitleObject.Caption = "Invoice Amount FC"
                    .Columns.Item(5).TitleObject.Caption = "Invoice Currency"

                    .Columns.Item(0).Type = SAPbouiCOM.BoFormItemTypes.it_EDIT

                    oEditText = .Columns.Item("InvoiceNo")
                    oEditText.LinkedObjectType = "18"
                    oEditText.ChooseFromListUID = "CFLInv"
                    oEditText.ChooseFromListAlias = "DocEntry"

                    .Columns.Item(0).Editable = False
                    .Columns.Item(1).Editable = False
                    .Columns.Item(2).Editable = False
                    .Columns.Item(3).Editable = False
                    .Columns.Item(4).Editable = False
                    .Columns.Item(5).Editable = False
                End With

                'Purchase Order
                ogrdOdr = oForm.Items.Item("grdOdr").Specific
                With ogrdOdr
                    .Columns.Item(0).TitleObject.Caption = "Order #"
                    .Columns.Item(1).TitleObject.Caption = "Order Submit Date"
                    .Columns.Item(2).TitleObject.Caption = "Order Due Date"
                    .Columns.Item(3).TitleObject.Caption = "Order Amount"
                    .Columns.Item(4).TitleObject.Caption = "Order Amount FC"
                    .Columns.Item(5).TitleObject.Caption = "Order Currency"

                    .Columns.Item(0).Type = SAPbouiCOM.BoFormItemTypes.it_EDIT

                    oEditText = .Columns.Item("OrderNo")
                    oEditText.LinkedObjectType = "22"
                    oEditText.ChooseFromListUID = "CFLOdr"
                    oEditText.ChooseFromListAlias = "DocEntry"

                    .Columns.Item(0).Editable = False
                    .Columns.Item(1).Editable = False
                    .Columns.Item(2).Editable = False
                    .Columns.Item(3).Editable = False
                    .Columns.Item(4).Editable = False
                    .Columns.Item(5).Editable = False
                End With

                'Outgoing Payment
                ogrdPmt = oForm.Items.Item("grdPmt").Specific

                With ogrdPmt
                    .Columns.Item(0).TitleObject.Caption = "Payment #"
                    .Columns.Item(1).TitleObject.Caption = "Payment Amount"
                    .Columns.Item(2).TitleObject.Caption = "Payment Amount FC"
                    .Columns.Item(3).TitleObject.Caption = "Payment Currency"

                    .Columns.Item(0).Type = SAPbouiCOM.BoFormItemTypes.it_EDIT

                    oEditText = .Columns.Item("PaymentNo")
                    oEditText.LinkedObjectType = "46"
                    oEditText.ChooseFromListUID = "CFLPmt"
                    oEditText.ChooseFromListAlias = "DocEntry"

                    .Columns.Item(0).Editable = False
                    .Columns.Item(1).Editable = False
                    .Columns.Item(2).Editable = False
                    .Columns.Item(3).Editable = False
                End With

                ogrdInv.AutoResizeColumns()
                ogrdOdr.AutoResizeColumns()
                ogrdPmt.AutoResizeColumns()

            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                oForm.Freeze(False)
            End Try

        End Sub

        

        Public Overrides Sub OnInitializeComponent()
            Me.matInsr = CType(Me.GetItem("matInsr").Specific, SAPbouiCOM.Matrix)
            Me.matAmd = CType(Me.GetItem("matAmd").Specific, SAPbouiCOM.Matrix)
            Me.matAttach = CType(Me.GetItem("matAttach").Specific, SAPbouiCOM.Matrix)
            Me.ogrdInv = CType(Me.GetItem("grdInv").Specific, SAPbouiCOM.Grid)
            Me.ogrdPmt = CType(Me.GetItem("grdPmt").Specific, SAPbouiCOM.Grid)
            Me.ogrdOdr = CType(Me.GetItem("grdOdr").Specific, SAPbouiCOM.Grid)
            Me.txtAmdType = CType(Me.GetItem("matAmd").Specific, SAPbouiCOM.Matrix).Columns.Item("colAmdType")
            Me.txtInsType = CType(Me.GetItem("matInsr").Specific, SAPbouiCOM.Matrix).Columns.Item("colInsType")
            Me.colAmdVal = CType(Me.GetItem("matAmd").Specific, SAPbouiCOM.Matrix).Columns.Item("colAmdVal")
            Me.cboAdMtd = CType(Me.GetItem("cboAdMtd").Specific, SAPbouiCOM.ComboBox)
            Me.cboStatus = CType(Me.GetItem("cboStatus").Specific, SAPbouiCOM.ComboBox)
            Me.txtContr = CType(Me.GetItem("txtContr").Specific, SAPbouiCOM.EditText)
            Me.txtEmp = CType(Me.GetItem("txtEmp").Specific, SAPbouiCOM.EditText)
            Me.cboSponsor = CType(Me.GetItem("spons").Specific, SAPbouiCOM.ComboBox)
            Me.txtprjcode = CType(Me.GetItem("txtprjcode").Specific, SAPbouiCOM.EditText)
            Me.cboBizUnit = CType(Me.GetItem("bizunit").Specific, SAPbouiCOM.ComboBox)
            Me.cboCurr = CType(Me.GetItem("curr").Specific, SAPbouiCOM.ComboBox)
            Me.txtExchRate = CType(Me.GetItem("ExchRate").Specific, SAPbouiCOM.EditText)
            Me.txtValueFC = CType(Me.GetItem("ValueFC").Specific, SAPbouiCOM.EditText)
            Me.txtValueLC = CType(Me.GetItem("ValueLC").Specific, SAPbouiCOM.EditText)
            Me.txtAmdVal = CType(Me.GetItem("AmdVal").Specific, SAPbouiCOM.EditText)
            Me.txtTotInv = CType(Me.GetItem("totinvval").Specific, SAPbouiCOM.EditText)
            Me.txtTotOrd = CType(Me.GetItem("totordval").Specific, SAPbouiCOM.EditText)
            Me.txtTotPmt = CType(Me.GetItem("totpmtval").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("Item_36").Specific, SAPbouiCOM.StaticText)
            Me.StaticText1 = CType(Me.GetItem("Item_23").Specific, SAPbouiCOM.StaticText)
            Me.StaticText2 = CType(Me.GetItem("Item_32").Specific, SAPbouiCOM.StaticText)
            Me.txtContrCode = CType(Me.GetItem("ContrCode").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_34").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("Item_38").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("Item_39").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_41").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("Item_42").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("Item_48").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()

        End Sub

        Private WithEvents cflVendor As SAPbouiCOM.ChooseFromList
        Private WithEvents cflEmp As SAPbouiCOM.ChooseFromList
        Private WithEvents matInsr As SAPbouiCOM.Matrix
        Private WithEvents txtAmdType As SAPbouiCOM.Column
        Private WithEvents txtInsType As SAPbouiCOM.Column
        Private WithEvents txtInsType2 As SAPbouiCOM.Column
        Private WithEvents txtInsType3 As SAPbouiCOM.Column
        Private WithEvents colAmdVal As SAPbouiCOM.Column
        Private WithEvents matAmd As SAPbouiCOM.Matrix
        Private WithEvents cboAdMtd As SAPbouiCOM.ComboBox
        Private WithEvents cboStatus As SAPbouiCOM.ComboBox
        Private WithEvents ogrdInv As SAPbouiCOM.Grid
        Private WithEvents ogrdPmt As SAPbouiCOM.Grid
        Private WithEvents oEditText As SAPbouiCOM.EditTextColumn
        Private WithEvents oEditText2 As SAPbouiCOM.EditText
        Private WithEvents ogrdOdr As SAPbouiCOM.Grid
        Private WithEvents matAttach As SAPbouiCOM.Matrix
        Private WithEvents btnAttach As SAPbouiCOM.Button
        Private WithEvents txtContr As SAPbouiCOM.EditText
        Private WithEvents txtEmp As SAPbouiCOM.EditText
        Private WithEvents txtprjcode As SAPbouiCOM.EditText
        Private WithEvents cboSponsor As SAPbouiCOM.ComboBox
        Private WithEvents cboBizUnit As SAPbouiCOM.ComboBox
        Private WithEvents cboCurr As SAPbouiCOM.ComboBox
        Private WithEvents txtExchRate As SAPbouiCOM.EditText
        Private WithEvents txtValueFC As SAPbouiCOM.EditText
        Private WithEvents txtValueLC As SAPbouiCOM.EditText
        Private WithEvents txtAmdVal As SAPbouiCOM.EditText
        Private WithEvents txtTotInv As SAPbouiCOM.EditText
        Private WithEvents txtTotOrd As SAPbouiCOM.EditText
        Private WithEvents txtTotPmt As SAPbouiCOM.EditText
        Private Sub txtExchRate_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtExchRate.ValidateBefore
            Try
                txtExchRate = oForm.Items.Item("ExchRate").Specific
                txtValueLC = oForm.Items.Item("ValueLC").Specific
                txtAmdVal = oForm.Items.Item("AmdVal").Specific

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                oDataSource.SetValue("U_ValueFC", 0, If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value)) / If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value)))
                'oDataSource.SetValue("U_TotValueFC", 0, (If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value)) + If(txtAmdVal.Value.Length = 0, 0, CDbl(txtAmdVal.Value))) * If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value)))

                oDataSource.SetValue("U_TotValueFC", 0, ((If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value)) / If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value))) + If(txtAmdVal.Value.Length = 0, 0, CDbl(txtAmdVal.Value))))

            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub

        Private Sub txtValueLC_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtValueLC.ValidateBefore
            Try

                txtExchRate = oForm.Items.Item("ExchRate").Specific
                txtValueLC = oForm.Items.Item("ValueLC").Specific
                txtAmdVal = oForm.Items.Item("AmdVal").Specific

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                oDataSource.SetValue("U_ValueFC", 0, If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value) / If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value))))
                'oDataSource.SetValue("U_TotValueFC", 0, (If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value)) + If(txtAmdVal.Value.Length = 0, 0, CDbl(txtAmdVal.Value))) * If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value)))
                oDataSource.SetValue("U_TotValueFC", 0, ((If(txtValueLC.Value.Length = 0, 0, CDbl(txtValueLC.Value)) / If(txtExchRate.Value.Length = 0, 0, CDbl(txtExchRate.Value))) + If(txtAmdVal.Value.Length = 0, 0, CDbl(txtAmdVal.Value))))



            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub txtAmdVal_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles colAmdVal.ValidateBefore
            Try
                Dim oEditAmdVal As SAPbouiCOM.EditText    ' Amendment Value

                Dim CalcTotal As Double
                Dim i As Integer

                txtExchRate = oForm.Items.Item("ExchRate").Specific

                CalcTotal = 0
                ' Iterate all the matrix rows
                For i = 1 To matAmd.RowCount
                    oEditAmdVal = colAmdVal.Cells.Item(i).Specific
                    CalcTotal += oEditAmdVal.Value
                Next
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTRACT")
                oDataSource.SetValue("U_TotAmendVal", oDataSource.Offset, CalcTotal)
                oDataSource.SetValue("U_TotValueFC", oDataSource.Offset, (oDataSource.GetValue("U_ValueFC", 0) + CalcTotal))
                'If(txtExchRate.Value.Length = 0, 0, txtExchRate.Value)
                oForm.Update()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
       
        Private Sub cboBizUnit_ComboSelectAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles cboBizUnit.ComboSelectAfter

            'If cboBizUnit.Selected.Value = "Add New" Then
            '    SBO_Application.ActivateMenuItem("Contract.BU")
            'End If
        End Sub
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents txtContrCode As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText

    End Class
End Namespace
