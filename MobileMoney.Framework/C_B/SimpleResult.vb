Namespace C_B
    Public Class SimpleResult
        Public Property serviceDate As Date = Date.Today

        ''' <summary>
        ''' Third party's unique identifier of the transaction
        ''' </summary>
        ''' <returns></returns>
        Public Property serviceID As String
        ''' <summary>
        ''' Third party's unique identifier of the transaction
        ''' </summary>
        ''' <returns></returns>
        Public Property serviceReceipt As String
        ''' <summary>
        ''' 'Completed' when processing is completed
        ''' </summary>
        ''' <returns></returns>
        Public Property resultType As String
        ''' <summary>
        ''' '0' when service has been successfully provisioned And '999' when service has not been provisioned successfully
        ''' </summary>
        ''' <returns></returns>
        Public Property resultCode As Long


    End Class
End Namespace
