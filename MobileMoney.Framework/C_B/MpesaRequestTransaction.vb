Imports MobileMoney.Framework
Namespace C_B
    Public Class MpesaRequestTransaction
        Implements IRequestTransaction, IRequestTransactionMetaData

        Public Property accountReference As String Implements IRequestTransaction.accountReference


        Public Property amount As Decimal Implements IRequestTransaction.amount


        Public Property commandID As String Implements IRequestTransaction.commandID

        Public Property conversationID As String Implements IRequestTransactionMetaData.conversationID


        Public Property initiator As String Implements IRequestTransaction.initiator


        Public Property originatorConversationID As String Implements IRequestTransactionMetaData.originatorConversationID


        Public Property mpesaReceipt As String Implements IRequestTransaction.Receipt


        Public Property recipient As Integer Implements IRequestTransaction.recipient


        Public Property transactionDate As Date Implements IRequestTransactionMetaData.transactionDate


        Public Property transactionID As String Implements IRequestTransactionMetaData.transactionID

    End Class
End Namespace
