Namespace C_B
    Public Class BrokerRequest
        Public Property resultURL As String

        <Newtonsoft.Json.JsonConverter(GetType(Converters.IRequestTransactionConverter))>
        Public Property transaction As IRequestTransaction

        Public Property serviceProvider As ServiceProvider

    End Class
End Namespace
