Public Class FrameworkFactory



    Public Shared Sub Initilize()
        Dim scheduler = Quartz.Impl.StdSchedulerFactory.GetDefaultScheduler()
        If Not scheduler.IsStarted Then
            scheduler.Start()
        End If


    End Sub
End Class
