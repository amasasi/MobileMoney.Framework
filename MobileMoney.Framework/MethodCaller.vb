
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection

Namespace Reflection
    ''' <summary>
    ''' Provides methods to dynamically find and call methods.
    ''' </summary>
    Public Class MethodCaller
        Private Const allLevelFlags As BindingFlags = BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic Or BindingFlags.FlattenHierarchy

        Private Const oneLevelFlags As BindingFlags = BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic

        Private Const ctorFlags As BindingFlags = BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic

        Private Const factoryFlags As BindingFlags = BindingFlags.[Static] Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy

        Private Const privateMethodFlags As BindingFlags = BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic Or BindingFlags.FlattenHierarchy

        Private Shared _methodCache As Dictionary(Of MethodCacheKey, DynamicMethodHandle)

        Private Shared _ctorCache As Dictionary(Of Type, DynamicCtorDelegate)

        Private Const propertyFlags As BindingFlags = BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy

        Private Const fieldFlags As BindingFlags = BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic

        Private Shared ReadOnly _memberCache As Dictionary(Of MethodCacheKey, DynamicMemberHandle)

        Shared Sub New()
            MethodCaller._methodCache = New Dictionary(Of MethodCacheKey, DynamicMethodHandle)()
            MethodCaller._ctorCache = New Dictionary(Of Type, DynamicCtorDelegate)()
            MethodCaller._memberCache = New Dictionary(Of MethodCacheKey, DynamicMemberHandle)()
        End Sub

        ''' <summary>
        ''' Invokes a static factory method.
        ''' </summary>
        ''' <param name="objectType">Business class where the factory is defined.</param>
        ''' <param name="method">Name of the factory method</param>
        ''' <param name="parameters">Parameters passed to factory method.</param>
        ''' <returns>Result of the factory method invocation.</returns>
        Public Shared Function CallFactoryMethod(ByVal objectType As Type, ByVal method As String, ByVal ParamArray parameters As Object()) As Object
            Dim returnValue As Object
            Dim factory As MethodInfo = objectType.GetMethod(method, BindingFlags.[Static] Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy, Nothing, MethodCaller.GetParameterTypes(parameters), Nothing)
            If (factory Is Nothing) Then
                Dim parameterCount As Integer = CInt(parameters.Length)
                Dim methods As MethodInfo() = objectType.GetMethods(BindingFlags.[Static] Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy)
                Dim num As Integer = 0
                While num < CInt(methods.Length)
                    Dim oneMethod As MethodInfo = methods(num)
                    If (Not (oneMethod.Name = method) OrElse CInt(oneMethod.GetParameters().Length) <> parameterCount) Then
                        num = num + 1
                    Else
                        factory = oneMethod
                        Exit While
                    End If
                End While
            End If
            If (factory Is Nothing) Then
                Throw New InvalidOperationException(String.Format("No Such Factory Method", method))
            End If
            Try
                returnValue = factory.Invoke(Nothing, parameters)
            Catch exception As System.Exception
                Dim ex As System.Exception = exception
                Dim inner As System.Exception = Nothing
                inner = If(ex.InnerException IsNot Nothing, ex.InnerException, ex)
                Throw New CallMethodException(String.Concat(New String() {objectType.Name, ".", factory.Name, " ", " Method Call Failed"}), inner)
            End Try
            Return returnValue
        End Function

        ''' <summary>
        ''' Uses reflection to dynamically invoke a method,
        ''' throwing an exception if it is not
        ''' implemented on the target object.
        ''' </summary>
        ''' <param name="obj">
        ''' Object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        Public Shared Function CallMethod(ByVal obj As Object, ByVal method As String) As Object
            Return MethodCaller.CallMethod(obj, method, False, Nothing)
        End Function

        ''' <summary>
        ''' Uses reflection to dynamically invoke a method,
        ''' throwing an exception if it is not
        ''' implemented on the target object.
        ''' </summary>
        ''' <param name="obj">
        ''' Object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        ''' <param name="parameters">
        ''' Parameters to pass to method.
        ''' </param>
        Public Shared Function CallMethod(ByVal obj As Object, ByVal method As String, ByVal ParamArray parameters As Object()) As Object
            Return MethodCaller.CallMethod(obj, method, True, parameters)
        End Function

        Private Shared Function CallMethod(ByVal obj As Object, ByVal method As String, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As Object
            Dim mh As DynamicMethodHandle = MethodCaller.GetCachedMethod(obj, method, hasParameters, parameters)
            If (mh Is Nothing OrElse mh.DynamicMethod Is Nothing) Then
                Throw New NotImplementedException(String.Concat(New String() {obj.[GetType]().Name, ".", method, " ", " MethodNotImplemented"}))
            End If
            Return MethodCaller.CallMethod(obj, mh, hasParameters, parameters)
        End Function

        ''' <summary>
        ''' Uses reflection to dynamically invoke a method,
        ''' throwing an exception if it is not
        ''' implemented on the target object.
        ''' </summary>
        ''' <param name="obj">
        ''' Object containing method.
        ''' </param>
        ''' <param name="info">
        ''' System.Reflection.MethodInfo for the method.
        ''' </param>
        ''' <param name="parameters">
        ''' Parameters to pass to method.
        ''' </param>
        Public Shared Function CallMethod(ByVal obj As Object, ByVal info As MethodInfo, ByVal ParamArray parameters As Object()) As Object
            Return MethodCaller.CallMethod(obj, info, True, parameters)
        End Function

        Private Shared Function CallMethod(ByVal obj As Object, ByVal info As MethodInfo, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As Object
            Dim mh As DynamicMethodHandle = MethodCaller.GetCachedMethod(obj, info, parameters)
            If (mh Is Nothing OrElse mh.DynamicMethod Is Nothing) Then
                Throw New NotImplementedException(String.Concat(New String() {obj.[GetType]().Name, ".", info.Name, " ", "Method Not Implemented"}))
            End If
            Return MethodCaller.CallMethod(obj, mh, hasParameters, parameters)
        End Function

        Private Shared Function CallMethod(ByVal obj As Object, ByVal methodHandle As DynamicMethodHandle, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As Object
            Dim obj1 As Object
            Dim result As Object = Nothing
            Dim method As DynamicMethodDelegate = methodHandle.DynamicMethod
            Dim inParams As Object() = Nothing
            inParams = If(parameters IsNot Nothing, parameters, New Object(0) {})
            If (methodHandle.HasFinalArrayParam) Then
                Dim pCount As Integer = methodHandle.MethodParamsLength
                Dim inCount As Integer = CInt(inParams.Length)
                If (inCount = pCount - 1) Then
                    Dim objArray(pCount - 1) As Object
                    Dim num As Integer = 0
                    Do
                        objArray(num) = parameters(num)
                        num = num + 1
                    Loop While num <= pCount - 2
                    Dim objArray1 As Object() = objArray
                    Dim length As Integer = CInt(objArray.Length) - 1
                    If (Not hasParameters OrElse inParams.Length <> 0) Then
                        obj1 = Nothing
                    Else
                        obj1 = inParams
                    End If
                    objArray1(length) = obj1
                    inParams = objArray
                ElseIf (inCount = pCount AndAlso inParams(inCount - 1) IsNot Nothing AndAlso Not inParams(inCount - 1).[GetType]().IsArray OrElse inCount > pCount) Then
                    Dim extras As Integer = CInt(inParams.Length) - (pCount - 1)
                    Dim extraArray As Object() = MethodCaller.GetExtrasArray(extras, methodHandle.FinalArrayElementType)
                    Array.Copy(inParams, pCount - 1, extraArray, 0, extras)
                    Dim paramList(pCount - 1) As Object
                    Dim pos As Integer = 0
                    Do
                        paramList(pos) = parameters(pos)
                        pos = pos + 1
                    Loop While pos <= pCount - 2
                    paramList(CInt(paramList.Length) - 1) = extraArray
                    inParams = paramList
                End If
            End If
            Try
                result = methodHandle.DynamicMethod(obj, inParams)
            Catch exception As System.Exception
                Dim ex As System.Exception = exception
                Throw New CallMethodException(String.Concat(New String() {obj.[GetType]().Name, ".", methodHandle.MethodName, " ", "MethodCallFailed "}), ex)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' Invokes an instance method on an object.
        ''' </summary>
        ''' <param name="obj">Object containing method.</param>
        ''' <param name="info">Method info object.</param>
        ''' <returns>Any value returned from the method.</returns>
        Public Shared Function CallMethod(ByVal obj As Object, ByVal info As MethodInfo) As Object
            Dim result As Object = Nothing
            Try
                result = info.Invoke(obj, Nothing)
            Catch exception As System.Exception
                Dim e As System.Exception = exception
                Dim inner As System.Exception = Nothing
                inner = If(e.InnerException IsNot Nothing, e.InnerException, e)
                Throw New CallMethodException(String.Concat(New String() {obj.[GetType]().Name, ".", info.Name, " ", " methodCallFailed"}), inner)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' Uses reflection to dynamically invoke a method
        ''' if that method is implemented on the target object.
        ''' </summary>
        ''' <param name="obj">
        ''' Object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        Public Shared Function CallMethodIfImplemented(ByVal obj As Object, ByVal method As String) As Object
            Return MethodCaller.CallMethodIfImplemented(obj, method, False, Nothing)
        End Function

        ''' <summary>
        ''' Uses reflection to dynamically invoke a method
        ''' if that method is implemented on the target object.
        ''' </summary>
        ''' <param name="obj">
        ''' Object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        ''' <param name="parameters">
        ''' Parameters to pass to method.
        ''' </param>
        Public Shared Function CallMethodIfImplemented(ByVal obj As Object, ByVal method As String, ByVal ParamArray parameters As Object()) As Object
            Return MethodCaller.CallMethodIfImplemented(obj, method, True, parameters)
        End Function

        Private Shared Function CallMethodIfImplemented(ByVal obj As Object, ByVal method As String, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As Object
            Dim mh As DynamicMethodHandle = MethodCaller.GetCachedMethod(obj, method, parameters)
            If (mh Is Nothing OrElse mh.DynamicMethod Is Nothing) Then
                Return Nothing
            End If
            Return MethodCaller.CallMethod(obj, mh, hasParameters, parameters)
        End Function

        ''' <summary>
        ''' Invokes a property getter using dynamic
        ''' method invocation.
        ''' </summary>
        ''' <param name="obj">Target object.</param>
        ''' <param name="property">Property to invoke.</param>
        ''' <returns></returns>
        Public Shared Function CallPropertyGetter(ByVal obj As Object, ByVal [property] As String) As Object
            If (obj Is Nothing) Then
                Throw New ArgumentNullException("obj")
            End If
            If (String.IsNullOrEmpty([property])) Then
                Throw New ArgumentException("Argument Is null Or empty.", "property")
            End If
            Dim mh As DynamicMemberHandle = MethodCaller.GetCachedProperty(obj.[GetType](), [property])
            If (mh.DynamicMemberGet Is Nothing) Then
                Throw New NotSupportedException(String.Format(CultureInfo.CurrentCulture, "The property '{0}' on Type '{1}' does not have a public getter.", [property], obj.[GetType]()))
            End If
            Return mh.DynamicMemberGet(obj)
        End Function

        ''' <summary>
        ''' Invokes a property setter using dynamic
        ''' method invocation.
        ''' </summary>
        ''' <param name="obj">Target object.</param>
        ''' <param name="property">Property to invoke.</param>
        ''' <param name="value">New value for property.</param>
        Public Shared Sub CallPropertySetter(ByVal obj As Object, ByVal [property] As String, ByVal value As Object)
            If (obj Is Nothing) Then
                Throw New ArgumentNullException("obj")
            End If
            If (String.IsNullOrEmpty([property])) Then
                Throw New ArgumentException("Argument is null or empty.", "property")
            End If
            Dim mh As DynamicMemberHandle = MethodCaller.GetCachedProperty(obj.[GetType](), [property])
            If (mh.DynamicMemberSet Is Nothing) Then
                Throw New NotSupportedException(String.Format(CultureInfo.CurrentCulture, "The property '{0}' on Type '{1}' does not have a public setter.", [property], obj.[GetType]()))
            End If
            mh.DynamicMemberSet.Invoke(obj, value)
        End Sub

        ''' <summary>
        ''' Uses reflection to create an object using its 
        ''' default constructor.
        ''' </summary>
        ''' <param name="objectType">Type of object to create.</param>
        Public Shared Function CreateInstance(ByVal objectType As Type) As Object
            Dim ctor As DynamicCtorDelegate = MethodCaller.GetCachedConstructor(objectType)
            If (ctor Is Nothing) Then
                Throw New NotImplementedException()
            End If
            Return ctor()
        End Function

        ''' <summary>
        ''' Returns information about the specified
        ''' method, even if the parameter types are
        ''' generic and are located in an abstract
        ''' generic base class.
        ''' </summary>
        ''' <param name="objectType">
        ''' Type of object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        ''' <param name="types">
        ''' Parameter types to pass to method.
        ''' </param>
        Public Shared Function FindMethod(ByVal objectType As Type, ByVal method As String, ByVal types As Type()) As MethodInfo
            Dim info As MethodInfo = Nothing
            Do
                info = objectType.GetMethod(method, BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic, Nothing, types, Nothing)
                If (info IsNot Nothing) Then
                    Exit Do
                End If
                objectType = objectType.BaseType
            Loop While objectType IsNot Nothing
            Return info
        End Function

        ''' <summary>
        ''' Returns information about the specified
        ''' method, finding the method based purely
        ''' on the method name and number of parameters.
        ''' </summary>
        ''' <param name="objectType">
        ''' Type of object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        ''' <param name="parameterCount">
        ''' Number of parameters to pass to method.
        ''' </param>
        Public Shared Function FindMethod(ByVal objectType As Type, ByVal method As String, ByVal parameterCount As Integer) As MethodInfo
            Dim result As MethodInfo = Nothing
            Dim currentType As Type = objectType
            Do
                Dim info As MethodInfo = currentType.GetMethod(method, BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic)
                If (info IsNot Nothing) Then
                    Dim infoParams As ParameterInfo() = info.GetParameters()
                    Dim pCount As Integer = CInt(infoParams.Length)
                    If (pCount > 0 AndAlso (pCount = 1 AndAlso infoParams(0).ParameterType.IsArray OrElse infoParams(pCount - 1).GetCustomAttributes(GetType(ParamArrayAttribute), True).Length <> 0)) Then
                        If (parameterCount >= pCount - 1) Then
                            result = info
                            Exit Do
                        End If
                    ElseIf (pCount = parameterCount) Then
                        result = info
                        Exit Do
                    End If
                End If
                currentType = currentType.BaseType
            Loop While currentType IsNot Nothing
            Return result
        End Function

        ''' <summary>
        ''' Returns information about the specified
        ''' method.
        ''' </summary>
        ''' <param name="objType">Type of object.</param>
        ''' <param name="method">Name of the method.</param>
        ''' <param name="flags">Flag values.</param>
        Public Shared Function FindMethod(ByVal objType As Type, ByVal method As String, ByVal flags As BindingFlags) As MethodInfo
            Dim info As MethodInfo = Nothing
            Do
                info = objType.GetMethod(method, flags)
                If (info IsNot Nothing) Then
                    Exit Do
                End If
                objType = objType.BaseType
            Loop While objType IsNot Nothing
            Return info
        End Function

        Private Shared Function FindMethodUsingFuzzyMatching(ByVal objectType As Type, ByVal method As String, ByVal parameters As Object()) As System.Reflection.MethodInfo
            Dim result As System.Reflection.MethodInfo = Nothing
            Dim currentType As Type = objectType
            Do
                Dim methods As System.Reflection.MethodInfo() = currentType.GetMethods(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic)
                Dim parameterCount As Integer = CInt(parameters.Length)
                Dim methodInfoArray As System.Reflection.MethodInfo() = methods
                Dim num As Integer = 0
                Do
                    Dim methodInfo As System.Reflection.MethodInfo = methodInfoArray(num)
                    If (methodInfo.Name = method) Then
                        Dim infoParams As ParameterInfo() = methodInfo.GetParameters()
                        Dim pCount As Integer = CInt(infoParams.Length)
                        If (pCount > 0) Then
                            If (pCount = 1 AndAlso infoParams(0).ParameterType.IsArray AndAlso parameters.[GetType]().Equals(infoParams(0).ParameterType)) Then
                                result = methodInfo
                                Exit Do
                            ElseIf (infoParams(pCount - 1).GetCustomAttributes(GetType(ParamArrayAttribute), True).Length <> 0 AndAlso parameterCount = pCount AndAlso parameters(pCount - 1).[GetType]().Equals(infoParams(pCount - 1).ParameterType)) Then
                                result = methodInfo
                                Exit Do
                            End If
                        End If
                    End If
                    num = num + 1
                Loop While num < CInt(methodInfoArray.Length)
                If (result Is Nothing) Then
                    Dim methodInfoArray1 As System.Reflection.MethodInfo() = methods
                    Dim num1 As Integer = 0
                    While num1 < CInt(methodInfoArray1.Length)
                        Dim m As System.Reflection.MethodInfo = methodInfoArray1(num1)
                        If (Not (m.Name = method) OrElse CInt(m.GetParameters().Length) <> parameterCount) Then
                            num1 = num1 + 1
                        Else
                            result = m
                            Exit While
                        End If
                    End While
                End If
                If (result IsNot Nothing) Then
                    Exit Do
                End If
                currentType = currentType.BaseType
            Loop While currentType IsNot Nothing
            Return result
        End Function

        Private Shared Function GetCachedConstructor(ByVal objectType As Type) As DynamicCtorDelegate
            Dim result As DynamicCtorDelegate = Nothing
            Dim found As Boolean = False
            Try
                found = MethodCaller._ctorCache.TryGetValue(objectType, result)
            Catch
            End Try
            If (Not found) Then
                SyncLock MethodCaller._ctorCache
                    If (Not MethodCaller._ctorCache.TryGetValue(objectType, result)) Then
                        Dim info As ConstructorInfo = objectType.GetConstructor(BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic, Nothing, Type.EmptyTypes, Nothing)
                        If (info Is Nothing) Then
                            Throw New NotSupportedException(String.Format(CultureInfo.CurrentCulture, "Cannot create instance of Type '{0}'. No public parameterless constructor found.", objectType))
                        End If
                        result = DynamicMethodHandlerFactory.CreateConstructor(info)
                        MethodCaller._ctorCache.Add(objectType, result)
                    End If
                End SyncLock
            End If
            Return result
        End Function

        Friend Shared Function GetCachedField(ByVal objectType As Type, ByVal fieldName As String) As DynamicMemberHandle
            Dim key As MethodCacheKey = New MethodCacheKey(objectType.FullName, fieldName, MethodCaller.GetParameterTypes(Nothing))
            Dim mh As DynamicMemberHandle = Nothing
            If (Not MethodCaller._memberCache.TryGetValue(key, mh)) Then
                SyncLock MethodCaller._memberCache
                    If (Not MethodCaller._memberCache.TryGetValue(key, mh)) Then
                        Dim info As FieldInfo = objectType.GetField(fieldName, BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic)
                        If (info Is Nothing) Then
                            Throw New InvalidOperationException(String.Format("MemberNotFoundException", fieldName))
                        End If
                        mh = New DynamicMemberHandle(info)
                        MethodCaller._memberCache.Add(key, mh)
                    End If
                End SyncLock
            End If
            Return mh
        End Function

        Private Shared Function GetCachedMethod(ByVal obj As Object, ByVal info As MethodInfo, ByVal ParamArray parameters As Object()) As DynamicMethodHandle
            Dim key As MethodCacheKey = New MethodCacheKey(obj.[GetType]().FullName, info.Name, MethodCaller.GetParameterTypes(parameters))
            Dim mh As DynamicMethodHandle = Nothing
            Dim found As Boolean = False
            Try
                found = MethodCaller._methodCache.TryGetValue(key, mh)
            Catch
            End Try
            If (Not found) Then
                SyncLock MethodCaller._methodCache
                    If (Not MethodCaller._methodCache.TryGetValue(key, mh)) Then
                        mh = New DynamicMethodHandle(info, parameters)
                        MethodCaller._methodCache.Add(key, mh)
                    End If
                End SyncLock
            End If
            Return mh
        End Function

        Private Shared Function GetCachedMethod(ByVal obj As Object, ByVal method As String) As DynamicMethodHandle
            Return MethodCaller.GetCachedMethod(obj, method, False, Nothing)
        End Function

        Private Shared Function GetCachedMethod(ByVal obj As Object, ByVal method As String, ByVal ParamArray parameters As Object()) As DynamicMethodHandle
            Return MethodCaller.GetCachedMethod(obj, method, True, parameters)
        End Function

        Private Shared Function GetCachedMethod(ByVal obj As Object, ByVal method As String, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As DynamicMethodHandle
            Dim key As MethodCacheKey = New MethodCacheKey(obj.[GetType]().FullName, method, MethodCaller.GetParameterTypes(hasParameters, parameters))
            Dim mh As DynamicMethodHandle = Nothing
            If (Not MethodCaller._methodCache.TryGetValue(key, mh)) Then
                SyncLock MethodCaller._methodCache
                    If (Not MethodCaller._methodCache.TryGetValue(key, mh)) Then
                        Dim info As MethodInfo = MethodCaller.GetMethod(obj.[GetType](), method, hasParameters, parameters)
                        mh = New DynamicMethodHandle(info, parameters)
                        MethodCaller._methodCache.Add(key, mh)
                    End If
                End SyncLock
            End If
            Return mh
        End Function

        Friend Shared Function GetCachedProperty(ByVal objectType As Type, ByVal propertyName As String) As DynamicMemberHandle
            Dim key As MethodCacheKey = New MethodCacheKey(objectType.FullName, propertyName, MethodCaller.GetParameterTypes(Nothing))
            Dim mh As DynamicMemberHandle = Nothing
            If (Not MethodCaller._memberCache.TryGetValue(key, mh)) Then
                SyncLock MethodCaller._memberCache
                    If (Not MethodCaller._memberCache.TryGetValue(key, mh)) Then
                        Dim info As PropertyInfo = objectType.GetProperty(propertyName, BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy)
                        If (info Is Nothing) Then
                            Throw New InvalidOperationException(String.Format("member not found", propertyName))
                        End If
                        mh = New DynamicMemberHandle(info)
                        MethodCaller._memberCache.Add(key, mh)
                    End If
                End SyncLock
            End If
            Return mh
        End Function

        Private Shared Function GetExtrasArray(ByVal count As Integer, ByVal arrayType As Type) As Object()
            Return DirectCast(Array.CreateInstance(arrayType.GetElementType(), count), Object())
        End Function

        ''' <summary>
        ''' Uses reflection to locate a matching method
        ''' on the target object.
        ''' </summary>
        ''' <param name="objectType">
        ''' Type of object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        Public Shared Function GetMethod(ByVal objectType As Type, ByVal method As String) As MethodInfo
            Return MethodCaller.GetMethod(objectType, method, True, New Object() {False, Nothing})
        End Function

        ''' <summary>
        ''' Uses reflection to locate a matching method
        ''' on the target object.
        ''' </summary>
        ''' <param name="objectType">
        ''' Type of object containing method.
        ''' </param>
        ''' <param name="method">
        ''' Name of the method.
        ''' </param>
        ''' <param name="parameters">
        ''' Parameters to pass to method.
        ''' </param>
        Public Shared Function GetMethod(ByVal objectType As Type, ByVal method As String, ByVal ParamArray parameters As Object()) As MethodInfo
            Return MethodCaller.GetMethod(objectType, method, True, parameters)
        End Function

        Private Shared Function GetMethod(ByVal objectType As Type, ByVal method As String, ByVal hasParameters As Boolean, ByVal ParamArray parameters As Object()) As MethodInfo
            Dim result As MethodInfo = Nothing
            Dim inParams As Object() = Nothing
            If (hasParameters) Then
                inParams = If(parameters IsNot Nothing, parameters, New Object(0) {})
            Else
                ReDim inParams(-1)
            End If
            result = MethodCaller.FindMethod(objectType, method, MethodCaller.GetParameterTypes(hasParameters, inParams))
            If (result Is Nothing) Then
                Try
                    result = MethodCaller.FindMethod(objectType, method, CInt(inParams.Length))
                Catch ambiguousMatchException As System.Reflection.AmbiguousMatchException
                    result = MethodCaller.FindMethodUsingFuzzyMatching(objectType, method, inParams)
                End Try
            End If
            If (result Is Nothing) Then
                result = objectType.GetMethod(method, BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic Or BindingFlags.FlattenHierarchy)
            End If
            Return result
        End Function

        ''' <summary>
        ''' Gets a System.Reflection.MethodInfo object corresponding to a
        ''' non-public method.
        ''' </summary>
        ''' <param name="objectType">Object containing the method.</param>
        ''' <param name="method">Name of the method.</param>
        Public Shared Function GetNonPublicMethod(ByVal objectType As Type, ByVal method As String) As MethodInfo
            Return MethodCaller.FindMethod(objectType, method, BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.NonPublic Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' Returns an array of Type objects corresponding
        ''' to the type of parameters provided.
        ''' </summary>
        Public Shared Function GetParameterTypes() As Type()
            Return MethodCaller.GetParameterTypes(False, Nothing)
        End Function

        ''' <summary>
        ''' Returns an array of Type objects corresponding
        ''' to the type of parameters provided.
        ''' </summary>
        ''' <param name="parameters">
        ''' Parameter values.
        ''' </param>
        Public Shared Function GetParameterTypes(ByVal parameters As Object()) As Type()
            Return MethodCaller.GetParameterTypes(True, parameters)
        End Function

        Private Shared Function GetParameterTypes(ByVal hasParameters As Boolean, ByVal parameters As Object()) As Type()
            If (Not hasParameters) Then
                Return New Type(-1) {}
            End If
            Dim result As List(Of Type) = New List(Of Type)()
            If (parameters IsNot Nothing) Then
                Dim objArray As Object() = parameters
                For i As Integer = 0 To CInt(objArray.Length)
                    Dim item As Object = objArray(i)
                    If (item IsNot Nothing) Then
                        result.Add(item.[GetType]())
                    Else
                        result.Add(GetType(Object))
                    End If
                Next

            Else
                result.Add(GetType(Object))
            End If
            Return result.ToArray()
        End Function

        ''' <summary>
        ''' Gets information about a property.
        ''' </summary>
        ''' <param name="objectType">Object containing the property.</param>
        ''' <param name="propertyName">Name of the property.</param>
        Public Shared Function GetProperty(ByVal objectType As Type, ByVal propertyName As String) As PropertyInfo
            Return objectType.GetProperty(propertyName, BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.FlattenHierarchy)
        End Function

        ''' <summary>
        ''' Gets a property type descriptor by name.
        ''' </summary>
        ''' <param name="t">Type of object containing the property.</param>
        ''' <param name="propertyName">Name of the property.</param>
        Public Shared Function GetPropertyDescriptor(ByVal t As Type, ByVal propertyName As String) As PropertyDescriptor
            Dim flag As Boolean = False
            Dim propertyDescriptors As PropertyDescriptorCollection = TypeDescriptor.GetProperties(t)
            Dim result As PropertyDescriptor = Nothing
            For Each desc As PropertyDescriptor In propertyDescriptors
                If (desc.Name <> propertyName) Then
                    Continue For
                End If
                result = desc
                flag = True
                If (flag) Then
                    Exit For
                End If
            Next
            flag = False
            Return result
        End Function

        ''' <summary>
        ''' Gets a property value.
        ''' </summary>
        ''' <param name="obj">Object containing the property.</param>
        ''' <param name="info">Property info object for the property.</param>
        ''' <returns>The value of the property.</returns>
        Public Shared Function GetPropertyValue(ByVal obj As Object, ByVal info As PropertyInfo) As Object
            Dim result As Object = Nothing
            Try
                result = info.GetValue(obj, Nothing)
            Catch exception As System.Exception
                Dim e As System.Exception = exception
                Dim inner As System.Exception = Nothing
                inner = If(e.InnerException IsNot Nothing, e.InnerException, e)
                Throw New CallMethodException(String.Concat(New String() {obj.[GetType]().Name, ".", info.Name, " ", " method fasiled"}), inner)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' Gets a Type object based on the type name.
        ''' </summary>
        ''' <param name="typeName">Type name including assembly name.</param>
        ''' <param name="throwOnError">true to throw an exception if the type can't be found.</param>
        ''' <param name="ignoreCase">true for a case-insensitive comparison of the type name.</param>
        Public Shared Function [GetType](ByVal typeName As String, ByVal throwOnError As Boolean, ByVal ignoreCase As Boolean) As Type
            Return Type.[GetType](typeName, throwOnError, ignoreCase)
        End Function

        ''' <summary>
        ''' Gets a Type object based on the type name.
        ''' </summary>
        ''' <param name="typeName">Type name including assembly name.</param>
        ''' <param name="throwOnError">true to throw an exception if the type can't be found.</param>
        Public Shared Function [GetType](ByVal typeName As String, ByVal throwOnError As Boolean) As Type
            Return MethodCaller.[GetType](typeName, throwOnError, False)
        End Function

        ''' <summary>
        ''' Gets a Type object based on the type name.
        ''' </summary>
        ''' <param name="typeName">Type name including assembly name.</param>
        Public Shared Function [GetType](ByVal typeName As String) As Type
            Return MethodCaller.[GetType](typeName, True, False)
        End Function

        ''' <summary>
        ''' Detects if a method matching the name and parameters is implemented on the provided object.
        ''' </summary>
        ''' <param name="obj">The object implementing the method.</param>
        ''' <param name="method">The name of the method to find.</param>
        ''' <param name="parameters">The parameters matching the parameters types of the method to match.</param>
        ''' <returns>True obj implements a matching method.</returns>
        Public Shared Function IsMethodImplemented(ByVal obj As Object, ByVal method As String, ByVal ParamArray parameters As Object()) As Boolean
            Dim mh As DynamicMethodHandle = MethodCaller.GetCachedMethod(obj, method, parameters)
            If (mh Is Nothing) Then
                Return False
            End If
            Return mh.DynamicMethod IsNot Nothing
        End Function
    End Class
End Namespace
