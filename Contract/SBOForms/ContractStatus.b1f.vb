Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase


Namespace Contract
    <FormAttribute("Contract.ContractStatus", "SBOForms/ContractStatus.b1f")>
    Friend Class ContractStatus
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
