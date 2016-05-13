Namespace C_B
    Public Interface IRequestTransaction
        Property commandID As String
        Property amount As Decimal
        Property initiator As String
        Property conversationID As String
        Property recipient As Integer
        Property transactionDate As Date
        Property accountReference As String
        Property transactionID As String
        Property originatorConversationID As String

        Property Receipt As String

    End Interface
End Namespace
