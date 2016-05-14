
Imports System.Net
    Imports System.Web.Http

Namespace C_B
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
        Public Sub PostValue(ByVal value As C_B.BrokerRequest)


            FrameworkFactory.Initilize()


            Dim job = Quartz.JobBuilder.Create(Of RequestProcessJob)().Build()

            Dim trigger = Quartz.TriggerBuilder.Create()
            trigger.StartNow()


            Dim serializesettings = New Newtonsoft.Json.JsonSerializerSettings()

            serializesettings.TypeNameHandling = Newtonsoft.Json.TypeNameHandling.All

            Dim serialized_data = Newtonsoft.Json.JsonConvert.SerializeObject(value, serializesettings)
            job.JobDataMap.Add("c_bdata", serialized_data)
            Quartz.Impl.StdSchedulerFactory.GetDefaultScheduler().ScheduleJob(job, trigger.Build)
        End Sub

        ' PUT: api/CbServiceProviderController/5
        Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

        End Sub

        ' DELETE: api/CbServiceProviderController/5
        Public Sub DeleteValue(ByVal id As Integer)

        End Sub

    End Class
End Namespace
