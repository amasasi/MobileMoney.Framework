Imports Quartz

Namespace C_B
    Public Class RequestProcessJob
        Implements Quartz.IJob

        Public Sub Execute(context As IJobExecutionContext) Implements IJob.Execute


            Dim data_to_deserialize = context.MergedJobDataMap.Get("c_bdata").ToString
            Dim serializesettings = New Newtonsoft.Json.JsonSerializerSettings()

            serializesettings.TypeNameHandling = Newtonsoft.Json.TypeNameHandling.All

            Dim serialized_data = Newtonsoft.Json.JsonConvert.DeserializeObject(Of C_B.BrokerRequest)(data_to_deserialize)

            Dim resolve_assemblies As System.Web.Http.Dispatcher.IAssembliesResolver = System.Web.Http.GlobalConfiguration.Configuration.Services.GetService(GetType(System.Web.Http.Dispatcher.IAssembliesResolver))

            Dim command_based As New List(Of Type)
            Dim generic_based As New List(Of Type)
            For Each v In resolve_assemblies.GetAssemblies
                command_based.AddRange((From ass In v.GetTypes Where GetType(IRequestProcessingCommandTask).IsAssignableFrom(ass) And Not ass.IsInterface Select ass).ToList)
                generic_based.AddRange((From ass In v.GetTypes Where GetType(IRequestProcessingGenericTask).IsAssignableFrom(ass) And Not ass.IsInterface Select ass).ToList)
            Next

            Dim handler As Object = Nothing
            For Each comm In command_based
                handler = Activator.CreateInstance(comm)
                If Not CType(handler, IRequestProcessingCommandTask).Commands.Contains(serialized_data.transaction.commandID) Then
                    handler = Nothing
                End If
            Next
            If handler Is Nothing Then
                For Each comm In generic_based
                    handler = Activator.CreateInstance(comm)
                    Exit For
                Next
            End If
            If handler IsNot Nothing Then

                Dim result As SimpleResult = Reflection.MethodCaller.CallMethod(handler, "Execute", serialized_data.transaction)
                ' handler.GetType.InvokeMember("",)
            End If
        End Sub
    End Class
End Namespace
