Namespace Reflection
    Friend Class MethodCacheKey
        Private _hashKey As Integer

        Public Property MethodName As String

        Public Property ParamTypes As Type()

        Public Property TypeName As String

        Public Sub New(ByVal typeName As String, ByVal methodName As String, ByVal paramTypes As Type())
            MyBase.New()
            Me.TypeName = typeName
            Me.MethodName = methodName
            Me.ParamTypes = paramTypes
            Me._hashKey = typeName.GetHashCode()
            Me._hashKey = Me._hashKey Xor methodName.GetHashCode()
            Dim typeArray As Type() = paramTypes
            For i As Integer = 0 To CInt(typeArray.Length)
                Dim item As Type = typeArray(i)
                Me._hashKey = Me._hashKey Xor item.Name.GetHashCode()
            Next

        End Sub

        Private Function ArrayEquals(ByVal a1 As Type(), ByVal a2 As Type()) As Boolean
            If (CInt(a1.Length) <> CInt(a2.Length)) Then
                Return False
            End If
            Dim pos As Integer = 0
            Do
                If (a1(pos) <> a2(pos)) Then
                    Return False
                End If
                pos = pos + 1
            Loop While pos < CInt(a1.Length)
            Return True
        End Function

        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            Dim key As MethodCacheKey = TryCast(obj, MethodCacheKey)
            If (key IsNot Nothing AndAlso key.TypeName = Me.TypeName AndAlso key.MethodName = Me.MethodName AndAlso Me.ArrayEquals(key.ParamTypes, Me.ParamTypes)) Then
                Return True
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Me._hashKey
        End Function
    End Class

End Namespace
