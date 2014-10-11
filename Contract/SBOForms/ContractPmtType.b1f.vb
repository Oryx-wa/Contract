Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase


Namespace Contract
    <FormAttribute("Contract.ContractPmtType", "SBOForms/ContractPmtType.b1f")>
    Friend Class ContractPmtType
        Inherits UserFormBaseClass



        Public Sub New()

           
        End Sub


        Public Overrides Sub OnInitializeComponent()

            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()


        End Sub

        Private Sub OnCustomInitialize()

        End Sub




    End Class
End Namespace
