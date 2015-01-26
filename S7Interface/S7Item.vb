Public Class S7Item
    Friend m_strName As String
    Friend m_strAddress As String
    Friend m_objTag As Object
    Friend m_objValue As Object
    Friend m_iQualities As Integer
    Friend m_dtTimeStamp As DateTime

    Public Sub New()

    End Sub

    Public Sub New(ByVal Address As String)
        If Not Address Is Nothing Then
            If Address.Count > 0 Then
                m_strName = Address
                m_strAddress = Address
                m_objValue = InitializeValue(Address)
                m_iQualities = 0
                m_dtTimeStamp = Nothing
            End If
        End If
    End Sub

    Property Name() As String
        Get
            Return m_strName
        End Get

        Set(ByVal Name As String)
            m_strName = Name
        End Set
    End Property

    Property Address() As String
        Get
            Dim str As String = m_strAddress
            If str Is Nothing Then
                Return String.Empty
            Else
                Return str
            End If
        End Get

        Set(ByVal Address As String)
            If Not Address Is Nothing Then
                If Address.Count > 0 Then
                    m_strAddress = Address
                    m_objValue = InitializeValue(Address)
                    m_iQualities = 0
                    m_dtTimeStamp = Nothing
                End If
            End If

        End Set
    End Property

    Property Value() As Object
        Get
            Return m_objValue
        End Get

        Set(ByVal Value As Object)
            m_objValue = Value
        End Set
    End Property

    ReadOnly Property Qualities() As Integer
        Get
            Return m_iQualities
        End Get
    End Property

    ReadOnly Property TimeStamp() As DateTime
        Get
            Return m_dtTimeStamp
        End Get
    End Property

    Public Sub InitializeValue()
        If Not m_strAddress Is Nothing Then
            If m_strAddress.Count > 0 Then
                m_objValue = InitializeValue(Address)
                m_iQualities = 0
                m_dtTimeStamp = Nothing
            Else
                m_strAddress = ""
                m_objValue = Nothing
                m_iQualities = 0
                m_dtTimeStamp = Nothing
            End If
        Else
            m_strAddress = ""
            m_objValue = Nothing
            m_iQualities = 0
            m_dtTimeStamp = Nothing
        End If
    End Sub

    Private Function InitializeValue(ByVal strAddress As String) As Object
        Dim obj As Object = Nothing

        ' Quando imposto l'indirizzo, inizializzo anche la variabile Value.
        Dim charSep(1) As Char
        Dim fORG_ID As String()
        Dim strDataType As String
        Dim shortAddress As Short
        Dim byteDBNr As Byte
        Dim iArrayLenght As Integer

        charSep(0) = ","
        charSep(1) = "."

        fORG_ID = strAddress.Split(charSep)

        strDataType = ""
        If fORG_ID.Length > 1 Then
            If SplitDataTypeAndStartAddress(fORG_ID(0), strDataType, shortAddress) = True Then
                If strDataType.Contains("DB") Or strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO") Then
                    ' Numero DB
                    Try
                        byteDBNr = CByte(shortAddress)
                    Catch ex As Exception
                        byteDBNr = 0
                    End Try
                    If (byteDBNr >= 1 And byteDBNr <= 255 And strDataType.Contains("DB")) Or (byteDBNr = 0 And strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO")) Then
                        If SplitDataTypeAndStartAddress(fORG_ID(1), strDataType, shortAddress) = True Then
                            Select Case strDataType
                                Case "B"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Byte), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CByte(0)
                                    End If

                                Case "CHAR"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Char), iArrayLenght)
                                            For i As Integer = 0 To iArrayLenght - 1
                                                obj(i) = " "
                                            Next
                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CChar(" ")
                                    End If

                                Case "W"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If

                                Case "INT"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If

                                Case "DWORD"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If
                                Case "DINT"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If

                                Case "X"
                                    If fORG_ID.Length > 3 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(3).ToUpper())
                                            obj = Array.CreateInstance(GetType(Boolean), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = False
                                    End If

                                Case "REAL"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Double), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CDbl(0.0)
                                    End If

                                Case "STRING"
                                    ' Lunghezza dati
                                    ' Nr di Stringhe e Lunghezza dati
                                    If fORG_ID.Length > 3 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(3))
                                            obj = Array.CreateInstance(GetType(String), iArrayLenght)
                                            For iIndice = 0 To iArrayLenght - 1
                                                obj(iIndice) = ""
                                            Next iIndice

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = ""
                                    End If
                                Case "TIME"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If

                                Case "TOD"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Integer), iArrayLenght)

                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = CInt(0)
                                    End If

                                Case "DT"
                                    ' Lunghezza dati
                                    If fORG_ID.Length > 2 Then
                                        Try
                                            iArrayLenght = CShort(fORG_ID(2))
                                            obj = Array.CreateInstance(GetType(Date), iArrayLenght)
                                        Catch ex As Exception
                                        End Try
                                    Else
                                        iArrayLenght = 1
                                        obj = #12:00:00 AM#
                                    End If

                            End Select
                        End If
                    End If
                End If
            End If
        End If

        Return obj

    End Function

End Class
