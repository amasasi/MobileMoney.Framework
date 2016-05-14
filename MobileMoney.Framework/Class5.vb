
Imports System
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Reflection.Emit

Namespace Reflection
    Friend Class DynamicMethodHandlerFactory
        Public Shared Function CreateConstructor(ByVal constructor As ConstructorInfo) As DynamicCtorDelegate
            If (constructor Is Nothing) Then
                Throw New ArgumentNullException("constructor")
            End If
            If (constructor.GetParameters().Length <> 0) Then
                Throw New NotSupportedException("Constructors With ParametersNotSupported")
            End If
            Dim body As Expression = Expression.[New](constructor)
            If (constructor.DeclaringType.IsValueType) Then
                body = Expression.Convert(body, GetType(Object))
            End If
            Return Expression.Lambda(Of DynamicCtorDelegate)(body, Array.Empty(Of ParameterExpression)()).Compile()
        End Function

        Public Shared Function CreateFieldGetter(ByVal field As FieldInfo) As DynamicMemberGetDelegate
            If (field Is Nothing) Then
                Throw New ArgumentNullException("field")
            End If
            Dim target As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim body As Expression = Expression.Field(Expression.Convert(target, field.DeclaringType), field)
            If (field.FieldType.IsValueType) Then
                body = Expression.Convert(body, GetType(Object))
            End If
            Return Expression.Lambda(Of DynamicMemberGetDelegate)(body, New ParameterExpression() {target}).Compile()
        End Function

        Public Shared Function CreateFieldSetter(ByVal field As FieldInfo) As DynamicMemberSetDelegate
            If (field Is Nothing) Then
                Throw New ArgumentNullException("property")
            End If
            Dim target As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim val As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim body As Expression = Expression.Assign(Expression.Field(Expression.Convert(target, field.DeclaringType), field), Expression.Convert(val, field.FieldType))
            Return Expression.Lambda(Of DynamicMemberSetDelegate)(body, New ParameterExpression() {target, val}).Compile()
        End Function

        Public Shared Function CreateMethod(ByVal method As MethodInfo) As DynamicMethodDelegate
            If (method Is Nothing) Then
                Throw New ArgumentNullException("method")
            End If
            Dim pi As ParameterInfo() = method.GetParameters()
            Dim targetExpression As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim parametersExpression As ParameterExpression = Expression.Parameter(GetType(Object()))
            Dim callParametrs(CInt(pi.Length) - 1) As Expression
            Dim x As Integer = 0
            Do
                callParametrs(x) = Expression.Convert(Expression.ArrayIndex(parametersExpression, Expression.Constant(x)), pi(x).ParameterType)
                x = x + 1
            Loop While x < CInt(pi.Length)
            Dim instance As Expression = Expression.Convert(targetExpression, method.DeclaringType)
            Dim body As Expression = If(pi.Length <> 0, Expression.[Call](instance, method, callParametrs), Expression.[Call](instance, method))
            If (method.ReturnType = GetType(Void)) Then
                Dim target As LabelTarget = Expression.Label(GetType(Object))
                Dim nullRef As ConstantExpression = Expression.Constant(Nothing)
                body = Expression.Block(body, Expression.[Return](target, nullRef), Expression.Label(target, nullRef))
            ElseIf (method.ReturnType.IsValueType) Then
                body = Expression.Convert(body, GetType(Object))
            End If
            Return Expression.Lambda(Of DynamicMethodDelegate)(body, New ParameterExpression() {targetExpression, parametersExpression}).Compile()
        End Function

        Public Shared Function CreatePropertyGetter(ByVal [property] As PropertyInfo) As DynamicMemberGetDelegate
            If ([property] Is Nothing) Then
                Throw New ArgumentNullException("property")
            End If
            If (Not [property].CanRead) Then
                Return Nothing
            End If
            Dim target As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim body As Expression = Expression.[Property](Expression.Convert(target, [property].DeclaringType), [property])
            If ([property].PropertyType.IsValueType) Then
                body = Expression.Convert(body, GetType(Object))
            End If
            Return Expression.Lambda(Of DynamicMemberGetDelegate)(body, New ParameterExpression() {target}).Compile()
        End Function

        Public Shared Function CreatePropertySetter(ByVal [property] As PropertyInfo) As DynamicMemberSetDelegate
            If ([property] Is Nothing) Then
                Throw New ArgumentNullException("property")
            End If
            If (Not [property].CanWrite) Then
                Return Nothing
            End If
            Dim target As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim val As ParameterExpression = Expression.Parameter(GetType(Object))
            Dim body As Expression = Expression.Assign(Expression.[Property](Expression.Convert(target, [property].DeclaringType), [property]), Expression.Convert(val, [property].PropertyType))
            Return Expression.Lambda(Of DynamicMemberSetDelegate)(body, New ParameterExpression() {target, val}).Compile()
        End Function

        Private Shared Sub EmitCastToReference(ByVal il As ILGenerator, ByVal type As System.Type)
            If (type.IsValueType) Then
                il.Emit(OpCodes.Unbox_Any, type)
                Return
            End If
            il.Emit(OpCodes.Castclass, type)
        End Sub
    End Class
End Namespace
