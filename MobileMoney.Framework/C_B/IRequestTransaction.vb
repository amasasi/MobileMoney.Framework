Namespace C_B
    Public Interface IRequestTransaction
        Property commandID As String
        Property amount As Decimal
        Property initiator As String
        ''' <summary>
        ''' Short code of the service provider (receiving party), length should be 6 digits
        '''[Example] 123000
        ''' </summary>
        ''' <returns></returns>
        Property recipient As Integer

        Property accountReference As String



        Property Receipt As String

    End Interface
End Namespace
