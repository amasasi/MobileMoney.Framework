Imports System.Reflection
Namespace Reflection
    Friend Class DynamicMemberHandle
        Public Property DynamicMemberGet As DynamicMemberGetDelegate

        Public Property DynamicMemberSet As DynamicMemberSetDelegate

        Public Property MemberName As String

        Public Property MemberType As Type

        Public Sub New(ByVal memberName As String, ByVal memberType As Type, ByVal dynamicMemberGet As DynamicMemberGetDelegate, ByVal dynamicMemberSet As DynamicMemberSetDelegate)
            MyBase.New()
            Me.MemberName = memberName
            Me.MemberType = memberType
            Me.DynamicMemberGet = dynamicMemberGet
            Me.DynamicMemberSet = dynamicMemberSet
        End Sub

        Public Sub New(ByVal info As PropertyInfo)
            MyClass.New(info.Name, info.PropertyType, DynamicMethodHandlerFactory.CreatePropertyGetter(info), DynamicMethodHandlerFactory.CreatePropertySetter(info))
        End Sub

        Public Sub New(ByVal info As FieldInfo)
            MyClass.New(info.Name, info.FieldType, DynamicMethodHandlerFactory.CreateFieldGetter(info), DynamicMethodHandlerFactory.CreateFieldSetter(info))
        End Sub
    End Class
End Namespace
