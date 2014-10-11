Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase


Namespace Contract
    <FormAttribute("Contract.ContractType", "SBOForms/ContractType.b1f")>
    Friend Class ContractType
        Inherits UserFormBaseClass

        Dim oNavInicio As Boolean, oRec As SAPbobsCOM.Recordset

        Public Sub New()

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            FilterRecordSet()

        End Sub

        Private Sub FilterRecordSet()
            Try
                If Not IsNothing(oRec) Then
                    oNavInicio = False
                    oRec.DoQuery(String.Format("select * from [@OWA_PQCONTROPTION] where U_Type = 'T' order by docentry"))
                    If oRec.RecordCount > 0 Then
                        LoadByDocEntry(oRec.Fields.Item("Code").Value)
                    End If
                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub LoadByDocEntry(ByVal code As String)
            '
            If oForm.TypeEx = "Contract.ContractType" Then
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add
                With oCondition
                    .Alias = "Code"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = Convert.ToString(code).Trim
                End With
                oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION").Query(oConditions)
                oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION").Offset = 0

            End If

        End Sub

        Public Overrides Sub OnInitializeComponent()

            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()


        End Sub

        Private Sub OnCustomInitialize()

        End Sub



        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            oForm = SBO_Application.Forms.ActiveForm

            If oForm.TypeEx = "Contract.ContractType" Then
                If (pVal.BeforeAction = False) Then
                    Select Case pVal.MenuUID

                        Case "1288" 'Next
                            If (oNavInicio) Then

                                oRec.MoveFirst()
                                oNavInicio = False
                            Else
                                If (Not oRec.EoF) Then

                                    oRec.MoveNext()
                                    If (oRec.EoF) Then
                                        oRec.MoveFirst()

                                    End If
                                End If
                            End If

                        Case "1289"
                            If (oNavInicio) Then
                                oRec.MoveLast()
                                oNavInicio = False

                            Else
                                If (Not oRec.BoF) Then
                                    oRec.MovePrevious()
                                Else
                                    oRec.MoveLast()
                                End If
                            End If

                        Case "1290"
                            oRec.MoveFirst()
                            'Exit Select

                        Case "1291"
                            oRec.MoveLast()
                            'Exit Select
                    End Select

                    oForm.Freeze(True)

                    ' Put form in OK Mode  
                    ' oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                    ' Filter the DBDataSource with RecordSet's current record.
                    If oRec.RecordCount > 0 Then
                        LoadByDocEntry(oRec.Fields.Item("Code").Value)
                    Else
                        LoadByDocEntry("")
                    End If


                    oForm.Freeze(False)

                    'HandleNavigation()
                    ' Don't let the normal flow of SBO event  
                    BubbleEvent = False

                End If 'end of condition for BeforeAction

            End If ' end of condition for the type of form
            BubbleEvent = True
        End Sub



        Private Sub ContractAmend_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataAddAfter
            FilterRecordSet()
        End Sub

        Private Sub ContractAmend_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataAddBefore
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION")
                oDataSource.SetValue("U_Type", oDataSource.Offset, "T")

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub ContractAmend_DataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataUpdateAfter
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION")
                oDataSource.SetValue("U_Type", oDataSource.Offset, "T")

                oForm.Update()

                FilterRecordSet()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

    End Class
End Namespace
