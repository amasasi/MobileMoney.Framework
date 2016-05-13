
Imports System.Net
    Imports System.Web.Http

Namespace Controllers
    Public Class CbServiceProviderController
        Inherits ApiController

        ' GET: api/CbServiceProviderController
        Public Function GetValues() As IEnumerable(Of String)
            Return New String() {"value1", "value2"}
        End Function

        ' GET: api/CbServiceProviderController/5
        Public Function GetValue(ByVal id As Integer) As String
            Return "value"
        End Function



        ' POST: api/CbServiceProviderController
        Public Sub PostValue(ByVal value As BrokerRequest)




        End Sub

        ' PUT: api/CbServiceProviderController/5
        Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

        End Sub

        ' DELETE: api/CbServiceProviderController/5
        Public Sub DeleteValue(ByVal id As Integer)

        End Sub

    End Class
End Namespace
