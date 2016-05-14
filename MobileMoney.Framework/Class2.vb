Imports System.Reflection
Namespace Reflection
    Friend Class DynamicMethodHandle
        Public Property DynamicMethod As DynamicMethodDelegate

        Public Property FinalArrayElementType As Type

        Public Property HasFinalArrayParam As Boolean

        Public Property MethodName As String

        Public Property MethodParamsLength As Integer

        Public Sub New(ByVal info As MethodInfo, ByVal ParamArray parameters As Object())
            MyBase.New()
            If (info Is Nothing) Then
                Me.DynamicMethod = Nothing
                Return
            End If
            Me.MethodName = info.Name
            Dim infoParams As ParameterInfo() = info.GetParameters()
            Dim inParams As Object() = Nothing
            inParams = If(parameters IsNot Nothing, parameters, New Object(0) {})
            Dim pCount As Integer = CInt(infoParams.Length)
            If (pCount > 0 AndAlso (pCount = 1 AndAlso infoParams(0).ParameterType.IsArray OrElse infoParams(pCount - 1).GetCustomAttributes(GetType(ParamArrayAttribute), True).Length <> 0)) Then
                Me.HasFinalArrayParam = True
                Me.MethodParamsLength = pCount
                Me.FinalArrayElementType = infoParams(pCount - 1).ParameterType
            End If
            Me.DynamicMethod = DynamicMethodHandlerFactory.CreateMethod(info)
        End Sub
    End Class

End Namespace
