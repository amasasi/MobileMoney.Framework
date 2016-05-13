Imports MobileMoney.Framework

Public Class MpesaRequestTransaction
    Implements IRequestTransaction

    Public Property accountReference As String Implements IRequestTransaction.accountReference


    Public Property amount As Decimal Implements IRequestTransaction.amount


    Public Property commandID As String Implements IRequestTransaction.commandID

    Public Property conversationID As String Implements IRequestTransaction.conversationID


    Public Property initiator As String Implements IRequestTransaction.initiator


    Public Property originatorConversationID As String Implements IRequestTransaction.originatorConversationID


    Public Property mpesaReceipt As String Implements IRequestTransaction.Receipt


    Public Property recipient As Integer Implements IRequestTransaction.recipient


    Public Property transactionDate As Date Implements IRequestTransaction.transactionDate


    Public Property transactionID As String Implements IRequestTransaction.transactionID

End Class
