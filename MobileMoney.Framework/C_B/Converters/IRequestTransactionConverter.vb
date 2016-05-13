Imports Newtonsoft.Json
Namespace C_B.Converters
    Public Class IRequestTransactionConverter
        Inherits Newtonsoft.Json.JsonConverter

        Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
            serializer.Serialize(writer, value)
        End Sub

        Public Overrides Function CanConvert(objectType As Type) As Boolean
            Return GetType(IRequestTransaction).IsAssignableFrom(objectType)
        End Function

        Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
            Return serializer.Deserialize(Of MpesaRequestTransaction)(reader)
        End Function
    End Class
End Namespace