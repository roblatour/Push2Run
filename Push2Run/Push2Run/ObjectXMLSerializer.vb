
' For serialization of an object to an XML Binary file.
Imports System.IO
' For reading/writing data to an XML file.
Imports System.IO.IsolatedStorage
' For serialization of an object to an XML Document file.
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Xml.Serialization
' For accessing user isolated data.

Namespace XML
    ''' <summary>
    ''' Serialization format types.
    ''' </summary>
    Public Enum SerializedFormat
        ''' <summary>
        ''' Binary serialization format.
        ''' </summary>
        Binary

        ''' <summary>
        ''' Document serialization format.
        ''' </summary>
        Document
    End Enum

    ''' <summary>
    ''' Facade to XML serialization and deserialization of strongly typed objects to/from an XML file.
    ''' 
    ''' References: XML Serialization at http://samples.gotdotnet.com/:
    ''' http://samples.gotdotnet.com/QuickStart/howto/default.aspx?url=/quickstart/howto/doc/xmlserialization/rwobjfromxml.aspx
    ''' </summary>
    Public NotInheritable Class ObjectXMLSerializer(Of T As Class)
        Private Sub New()
        End Sub
        ' Specify that T must be a class.
#Region "Load methods"

        ''' <summary>
        ''' Loads an object from an XML file in Document format.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load(@"C:\XMLObjects.xml");
        ''' </code>
        ''' </example>
        ''' <param name="path">Path of the file to load the object from.</param>
        ''' <returns>Object loaded from an XML file in Document format.</returns>
        Public Shared Function Load(ByVal path As String) As T
            Dim serializableObject As T = LoadFromDocumentFormat(Nothing, path, Nothing)
            Return serializableObject
        End Function

        ''' <summary>
        ''' Loads an object from an XML file using a specified serialized format.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load(@"C:\XMLObjects.xml", SerializedFormat.Binary);
        ''' </code>
        ''' </example>		
        ''' <param name="path">Path of the file to load the object from.</param>
        ''' <param name="serializedFormat">XML serialized format used to load the object.</param>
        ''' <returns>Object loaded from an XML file using the specified serialized format.</returns>
        Public Shared Function Load(ByVal path As String, ByVal serializedFormat__1 As SerializedFormat) As T
            Dim serializableObject As T = Nothing

            Select Case serializedFormat__1
                Case SerializedFormat.Binary
                    serializableObject = LoadFromBinaryFormat(path, Nothing)
                    Exit Select

                Case SerializedFormat.Document ', Else
                    serializableObject = LoadFromDocumentFormat(Nothing, path, Nothing)
                    Exit Select
            End Select

            Return serializableObject
        End Function

        ''' <summary>
        ''' Loads an object from an XML file in Document format, supplying extra data types to enable deserialization of custom types within the object.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load(@"C:\XMLObjects.xml", new Type[] { typeof(MyCustomType) });
        ''' </code>
        ''' </example>
        ''' <param name="path">Path of the file to load the object from.</param>
        ''' <param name="extraTypes">Extra data types to enable deserialization of custom types within the object.</param>
        ''' <returns>Object loaded from an XML file in Document format.</returns>
        Public Shared Function Load(ByVal path As String, ByVal extraTypes As System.Type()) As T
            Dim serializableObject As T = LoadFromDocumentFormat(extraTypes, path, Nothing)
            Return serializableObject
        End Function

        ''' <summary>
        ''' Loads an object from an XML file in Document format, located in a specified isolated storage area.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load("XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly());
        ''' </code>
        ''' </example>
        ''' <param name="fileName">Name of the file in the isolated storage area to load the object from.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to load the object from.</param>
        ''' <returns>Object loaded from an XML file in Document format located in a specified isolated storage area.</returns>
        Public Shared Function Load(ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile) As T
            Dim serializableObject As T = LoadFromDocumentFormat(Nothing, fileName, isolatedStorageDirectory)
            Return serializableObject
        End Function

        ''' <summary>
        ''' Loads an object from an XML file located in a specified isolated storage area, using a specified serialized format.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load("XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly(), SerializedFormat.Binary);
        ''' </code>
        ''' </example>		
        ''' <param name="fileName">Name of the file in the isolated storage area to load the object from.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to load the object from.</param>
        ''' <param name="serializedFormat">XML serialized format used to load the object.</param>        
        ''' <returns>Object loaded from an XML file located in a specified isolated storage area, using a specified serialized format.</returns>
        Public Shared Function Load(ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile, ByVal serializedFormat__1 As SerializedFormat) As T
            Dim serializableObject As T = Nothing

            Select Case serializedFormat__1
                Case SerializedFormat.Binary
                    serializableObject = LoadFromBinaryFormat(fileName, isolatedStorageDirectory)
                    Exit Select

                Case SerializedFormat.Document ', Else
                    serializableObject = LoadFromDocumentFormat(Nothing, fileName, isolatedStorageDirectory)
                    Exit Select
            End Select

            Return serializableObject
        End Function

        ''' <summary>
        ''' Loads an object from an XML file in Document format, located in a specified isolated storage area, and supplying extra data types to enable deserialization of custom types within the object.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' serializableObject = ObjectXMLSerializer&lt;SerializableObject&gt;.Load("XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly(), new Type[] { typeof(MyCustomType) });
        ''' </code>
        ''' </example>		
        ''' <param name="fileName">Name of the file in the isolated storage area to load the object from.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to load the object from.</param>
        ''' <param name="extraTypes">Extra data types to enable deserialization of custom types within the object.</param>
        ''' <returns>Object loaded from an XML file located in a specified isolated storage area, using a specified serialized format.</returns>
        Public Shared Function Load(ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile, ByVal extraTypes As System.Type()) As T
            Dim serializableObject As T = LoadFromDocumentFormat(Nothing, fileName, isolatedStorageDirectory)
            Return serializableObject
        End Function

#End Region

#Region "Save methods"

        ''' <summary>
        ''' Saves an object to an XML file in Document format.
        ''' </summary>
        ''' <example>
        ''' <code>        
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, @"C:\XMLObjects.xml");
        ''' </code>
        ''' </example>
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="path">Path of the file to save the object to.</param>
        Public Shared Sub Save(ByVal serializableObject As T, ByVal path As String)
            SaveToDocumentFormat(serializableObject, Nothing, path, Nothing)
        End Sub

        ''' <summary>
        ''' Saves an object to an XML file using a specified serialized format.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, @"C:\XMLObjects.xml", SerializedFormat.Binary);
        ''' </code>
        ''' </example>
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="path">Path of the file to save the object to.</param>
        ''' <param name="serializedFormat">XML serialized format used to save the object.</param>
        Public Shared Sub Save(ByVal serializableObject As T, ByVal path As String, ByVal serializedFormat__1 As SerializedFormat)
            Select Case serializedFormat__1
                Case SerializedFormat.Binary
                    SaveToBinaryFormat(serializableObject, path, Nothing)
                    Exit Select

                Case SerializedFormat.Document ', Else
                    SaveToDocumentFormat(serializableObject, Nothing, path, Nothing)
                    Exit Select
            End Select
        End Sub

        ''' <summary>
        ''' Saves an object to an XML file in Document format, supplying extra data types to enable serialization of custom types within the object.
        ''' </summary>
        ''' <example>
        ''' <code>        
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, @"C:\XMLObjects.xml", new Type[] { typeof(MyCustomType) });
        ''' </code>
        ''' </example>
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="path">Path of the file to save the object to.</param>
        ''' <param name="extraTypes">Extra data types to enable serialization of custom types within the object.</param>
        Public Shared Sub Save(ByVal serializableObject As T, ByVal path As String, ByVal extraTypes As System.Type())
            SaveToDocumentFormat(serializableObject, extraTypes, path, Nothing)
        End Sub

        ''' <summary>
        ''' Saves an object to an XML file in Document format, located in a specified isolated storage area.
        ''' </summary>
        ''' <example>
        ''' <code>        
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, "XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly());
        ''' </code>
        ''' </example>
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="fileName">Name of the file in the isolated storage area to save the object to.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to save the object to.</param>
        Public Shared Sub Save(ByVal serializableObject As T, ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile)
            SaveToDocumentFormat(serializableObject, Nothing, fileName, isolatedStorageDirectory)
        End Sub

        ''' <summary>
        ''' Saves an object to an XML file located in a specified isolated storage area, using a specified serialized format.
        ''' </summary>
        ''' <example>
        ''' <code>        
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, "XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly(), SerializedFormat.Binary);
        ''' </code>
        ''' </example>
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="fileName">Name of the file in the isolated storage area to save the object to.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to save the object to.</param>
        ''' <param name="serializedFormat">XML serialized format used to save the object.</param>        
        Public Shared Sub Save(ByVal serializableObject As T, ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile, ByVal serializedFormat__1 As SerializedFormat)
            Select Case serializedFormat__1
                Case SerializedFormat.Binary
                    SaveToBinaryFormat(serializableObject, fileName, isolatedStorageDirectory)
                    Exit Select

                Case SerializedFormat.Document ', Else
                    SaveToDocumentFormat(serializableObject, Nothing, fileName, isolatedStorageDirectory)
                    Exit Select
            End Select
        End Sub

        ''' <summary>
        ''' Saves an object to an XML file in Document format, located in a specified isolated storage area, and supplying extra data types to enable serialization of custom types within the object.
        ''' </summary>
        ''' <example>
        ''' <code>
        ''' SerializableObject serializableObject = new SerializableObject();
        ''' 
        ''' ObjectXMLSerializer&lt;SerializableObject&gt;.Save(serializableObject, "XMLObjects.xml", IsolatedStorageFile.GetUserStoreForAssembly(), new Type[] { typeof(MyCustomType) });
        ''' </code>
        ''' </example>		
        ''' <param name="serializableObject">Serializable object to be saved to file.</param>
        ''' <param name="fileName">Name of the file in the isolated storage area to save the object to.</param>
        ''' <param name="isolatedStorageDirectory">Isolated storage area directory containing the XML file to save the object to.</param>
        ''' <param name="extraTypes">Extra data types to enable serialization of custom types within the object.</param>
        Public Shared Sub Save(ByVal serializableObject As T, ByVal fileName As String, ByVal isolatedStorageDirectory As IsolatedStorageFile, ByVal extraTypes As System.Type())
            SaveToDocumentFormat(serializableObject, Nothing, fileName, isolatedStorageDirectory)
        End Sub

#End Region

#Region "Private"

        Private Shared Function CreateFileStream(ByVal isolatedStorageFolder As IsolatedStorageFile, ByVal path As String) As FileStream
            Dim fileStream As FileStream = Nothing

            If isolatedStorageFolder Is Nothing Then
                fileStream = New FileStream(path, FileMode.OpenOrCreate)
            Else
                fileStream = New IsolatedStorageFileStream(path, FileMode.OpenOrCreate, isolatedStorageFolder)
            End If

            Return fileStream
        End Function

        Private Shared Function LoadFromBinaryFormat(ByVal path As String, ByVal isolatedStorageFolder As IsolatedStorageFile) As T
            Dim serializableObject As T = Nothing

            Using fileStream As FileStream = CreateFileStream(isolatedStorageFolder, path)
                Dim binaryFormatter As New BinaryFormatter()
                serializableObject = TryCast(binaryFormatter.Deserialize(fileStream), T)
            End Using

            Return serializableObject
        End Function

        Private Shared Function LoadFromDocumentFormat(ByVal extraTypes As System.Type(), ByVal path As String, ByVal isolatedStorageFolder As IsolatedStorageFile) As T
            Dim serializableObject As T = Nothing

            Using textReader As TextReader = CreateTextReader(isolatedStorageFolder, path)
                Dim xmlSerializer As XmlSerializer = CreateXmlSerializer(extraTypes)

                serializableObject = TryCast(xmlSerializer.Deserialize(textReader), T)
            End Using

            Return serializableObject
        End Function

        Private Shared Function CreateTextReader(ByVal isolatedStorageFolder As IsolatedStorageFile, ByVal path As String) As TextReader
            Dim textReader As TextReader = Nothing

            If isolatedStorageFolder Is Nothing Then
                textReader = New StreamReader(path)
            Else
                textReader = New StreamReader(New IsolatedStorageFileStream(path, FileMode.Open, isolatedStorageFolder))
            End If

            Return textReader
        End Function

        Private Shared Function CreateTextWriter(ByVal isolatedStorageFolder As IsolatedStorageFile, ByVal path As String) As TextWriter
            Dim textWriter As TextWriter = Nothing

            If isolatedStorageFolder Is Nothing Then
                textWriter = New StreamWriter(path)
            Else
                textWriter = New StreamWriter(New IsolatedStorageFileStream(path, FileMode.OpenOrCreate, isolatedStorageFolder))
            End If

            Return textWriter
        End Function

        Private Shared Function CreateXmlSerializer(ByVal extraTypes As System.Type()) As XmlSerializer
            Dim ObjectType As Type = GetType(T)

            Dim xmlSerializer As XmlSerializer = Nothing

            If extraTypes IsNot Nothing Then
                xmlSerializer = New XmlSerializer(ObjectType, extraTypes)
            Else
                xmlSerializer = New XmlSerializer(ObjectType)
            End If

            Return xmlSerializer
        End Function

        Private Shared Sub SaveToDocumentFormat(ByVal serializableObject As T, ByVal extraTypes As System.Type(), ByVal path As String, ByVal isolatedStorageFolder As IsolatedStorageFile)
            Using textWriter As TextWriter = CreateTextWriter(isolatedStorageFolder, path)
                Dim xmlSerializer As XmlSerializer = CreateXmlSerializer(extraTypes)
                xmlSerializer.Serialize(textWriter, serializableObject)
            End Using
        End Sub

        Private Shared Sub SaveToBinaryFormat(ByVal serializableObject As T, ByVal path As String, ByVal isolatedStorageFolder As IsolatedStorageFile)
            Using fileStream As FileStream = CreateFileStream(isolatedStorageFolder, path)
                Dim binaryFormatter As New BinaryFormatter()
                binaryFormatter.Serialize(fileStream, serializableObject)
            End Using
        End Sub

#End Region
    End Class

End Namespace
