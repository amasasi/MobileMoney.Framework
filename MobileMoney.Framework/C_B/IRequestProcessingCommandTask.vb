Namespace C_B
    Public Interface IRequestProcessingCommandTask

        Function Commands() As List(Of String)
        Function Execute(ByVal taskInfo As IRequestTransaction) As SimpleResult

    End Interface
End Namespace
