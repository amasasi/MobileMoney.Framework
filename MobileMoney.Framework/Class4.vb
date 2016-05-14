Imports System.Runtime.Serialization
Imports System.Security.Permissions
Namespace Reflection
    Public Delegate Function DynamicCtorDelegate() As Object


    Public Delegate Function DynamicMemberGetDelegate(ByVal target As Object) As Object



    Public Delegate Function DynamicMethodDelegate(ByVal target As Object, ByVal args As Object()) As Object



    Public Delegate Sub DynamicMemberSetDelegate(ByVal target As Object, ByVal arg As Object)

    <Serializable>
    Public Class CallMethodException
        Inherits Exception
        Private _innerStackTrace As String

        ''' <summary>
        ''' Get the stack trace from the original
        ''' exception.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property StackTrace As String
            Get
                Return String.Format("{0}{1}{2}", Me._innerStackTrace, Environment.NewLine, MyBase.StackTrace)
            End Get
        End Property

        ''' <summary>
        ''' Creates an instance of the object.
        ''' </summary>
        ''' <param name="message">Message text describing the exception.</param>
        ''' <param name="ex">Inner exception object.</param>
        Public Sub New(ByVal message As String, ByVal ex As Exception)
            MyBase.New(message, ex)
            Me._innerStackTrace = ex.StackTrace
        End Sub

        ''' <summary>
        ''' Creates an instance of the object for deserialization.
        ''' </summary>
        ''' <param name="info">Serialization info.</param>
        ''' <param name="context">Serialiation context.</param>
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
            Me._innerStackTrace = info.GetString("_innerStackTrace")
        End Sub

        ''' <summary>
        ''' Serializes the object.
        ''' </summary>
        ''' <param name="info">Serialization info.</param>
        ''' <param name="context">Serialization context.</param>
        <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.SerializationFormatter)>
        <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.SerializationFormatter)>
        Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.GetObjectData(info, context)
            info.AddValue("_innerStackTrace", Me._innerStackTrace)
        End Sub
    End Class
End Namespace
