Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase


Namespace Contract
    <FormAttribute("Contract.ContractInsurType", "SBOForms/ContractInsurType.b1f")>
    Friend Class ContractInsurType
        Inherits UserFormBaseClass

        Dim oNavInicio As Boolean, oRec As SAPbobsCOM.Recordset

        Public Sub New()

        End Sub


        Public Overrides Sub OnInitializeComponent()
            Me.EditText2 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
          

        End Sub

        Private Sub OnCustomInitialize()

        End Sub


    
        Private WithEvents EditText2 As SAPbouiCOM.EditText

    End Class
End Namespace
