Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace Contract
    <FormAttribute("Contract.BU", "SBOForms/BU.b1f")>
    Friend Class BU
        Inherits UserFormBaseClass

        Public Sub New()
            Try
 
                Matrix0.Clear()
                Matrix0.AutoResizeColumns()
                LoadBU("B")

                Matrix0.LoadFromDataSource()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub


        Public Overrides Sub OnInitializeComponent()
            Me.Matrix0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Private Sub LoadBU(ByVal Type As String)
            '
            If oForm.TypeEx = "Contract.BU" Then
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add
                With oCondition
                    .Alias = "U_Type"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = Convert.ToString(Type)
                End With
                oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION").Query(oConditions)
                oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION").Offset = 0
            End If

        End Sub

        Private Sub Matrix0_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_PQCONTROPTION").Clear()
            Matrix0 = oForm.Items.Item("Item_0").Specific

            If pVal.Row = Matrix0.RowCount + 1 Then
                If pVal.Row = 1 Then
                    Matrix0.AddRow(1)
                Else
                    Matrix0.AddRow(1, Matrix0.RowCount)
                End If
                Matrix0.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub OnCustomInitialize()

        End Sub
        
    End Class
End Namespace
