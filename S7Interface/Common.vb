Imports System.Net.Sockets
Imports System.Management

Module Common

    Sub ReadItems(ByRef tc As Net.Sockets.TcpClient, ByVal Items As S7Item())

        Dim byteReceive As Array
        Dim byteQualities As Byte
        Dim iaValueIndex As Integer

        For Each Item As S7Item In Items
            Try
                Dim bTemp As Boolean
                Dim byteTemp As Byte
                Dim aTemp As Array
                Dim strTemp As String

                ' Verifico il tipo di area da cui leggere
                Dim fORG_ID As String()
                Dim byteORG_ID As Byte

                ' Numero del DB su cui eseguire le operazioni
                Dim strDataType As String
                Dim shortAddress As Short

                Dim byteDBNr As Byte
                Dim shortStartAddress As Short

                Dim iArrayLenght As Integer
                Dim iByteLenght As Integer
                Dim iBitsLenght As Integer

                ' Serve per la codifica delle stringhe
                Dim enc As New System.Text.ASCIIEncoding

                strDataType = ""
                iaValueIndex = 0
                If Not Item Is Nothing Then

                    Item.m_iQualities = 0
                    Item.m_dtTimeStamp = Now

                    Dim charSep(1) As Char
                    charSep(0) = ","
                    charSep(1) = "."
                    fORG_ID = Item.Address.Split(charSep)

                    If fORG_ID.Length > 1 Then
                        If SplitDataTypeAndStartAddress(fORG_ID(0), strDataType, shortAddress) = True Then
                            If strDataType.Contains("DB") Or strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO") Then
                                If strDataType.Contains("DB") Then
                                    byteORG_ID = 1
                                End If
                                If strDataType.Contains("M") Then
                                    byteORG_ID = 2
                                End If
                                If strDataType.Contains("PI") Then
                                    byteORG_ID = 3
                                End If
                                If strDataType.Contains("PO") Then
                                    byteORG_ID = 4
                                End If
                                ' Numero DB
                                Try
                                    byteDBNr = CByte(shortAddress)

                                    If (byteDBNr >= 1 And byteDBNr <= 255 And strDataType.Contains("DB")) Or (byteDBNr = 0 And (strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO"))) Then
                                        If SplitDataTypeAndStartAddress(fORG_ID(1), strDataType, shortAddress) = True Then

                                            Select Case strDataType
                                                Case "B"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iByteLenght = CShort(fORG_ID(2))
                                                            iArrayLenght = (iByteLenght \ 2)
                                                            If iByteLenght Mod (2) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iByteLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iByteLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iByteLenght)

                                                            iaValueIndex = (iByteLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1)
                                                                iaValue(iaValueIndex) = byteReceive(iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = byteReceive(1)
                                                        End If
                                                    End If

                                                Case "CHAR"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iByteLenght = CShort(fORG_ID(2))
                                                            iArrayLenght = (iByteLenght \ 2)
                                                            If iByteLenght Mod (2) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iByteLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iByteLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iByteLenght)

                                                            iaValueIndex = (iByteLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1)
                                                                iaValue(iaValueIndex) = Convert.ToChar(byteReceive(iIndice_1))
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = Convert.ToChar(byteReceive(1))
                                                        End If
                                                    End If

                                                Case "W"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 2
                                                                iaValue(iaValueIndex) = BitConverter.ToUInt16(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToUInt16(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "INT"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 2
                                                                iaValue(iaValueIndex) = BitConverter.ToInt16(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToInt16(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "DWORD"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 4
                                                                iaValue(iaValueIndex) = BitConverter.ToUInt32(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToUInt32(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "DINT"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 4
                                                                iaValue(iaValueIndex) = BitConverter.ToInt32(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToInt32(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "X"
                                                    Dim shortBitNr As Short
                                                    ' Indirizzo di partenza
                                                    If fORG_ID.Length > 2 Then
                                                        shortStartAddress = shortAddress
                                                        Try
                                                            shortBitNr = CShort(fORG_ID(2))
                                                            If Not (shortBitNr >= 0 And shortBitNr) <= 7 Then
                                                                ' Errore
                                                                Item.m_iQualities = 4
                                                                Exit Select
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 3 Then
                                                        If shortBitNr = 0 Then
                                                            iBitsLenght = CShort(fORG_ID(3).ToUpper())
                                                            iArrayLenght = (iBitsLenght \ 16)
                                                            If iBitsLenght Mod (16) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End If
                                                    Else
                                                        iBitsLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iBitsLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iBitsLenght)

                                                            strTemp = ""
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1)
                                                                strTemp = strTemp + Convert.ToString(byteReceive(iIndice_1), 2).PadLeft(8, "0")
                                                            Next iIndice_1
                                                            aTemp = enc.GetBytes(strTemp)
                                                            Array.Reverse(aTemp)
                                                            strTemp = enc.GetString(aTemp)
                                                            For iIndice_1 As Integer = 0 To (iBitsLenght - 1)
                                                                If strTemp.Chars(iIndice_1) = "0" Then
                                                                    bTemp = False
                                                                Else
                                                                    bTemp = True
                                                                End If
                                                                iaValue(iIndice_1) = bTemp
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            strTemp = Convert.ToString(byteReceive(1), 2).PadLeft(8, "0")
                                                            aTemp = enc.GetBytes(strTemp)
                                                            Array.Reverse(aTemp)
                                                            strTemp = enc.GetString(aTemp)
                                                            If strTemp.Chars(shortBitNr) = "0" Then
                                                                bTemp = False
                                                            Else
                                                                bTemp = True
                                                            End If
                                                            Item.m_objValue = bTemp
                                                        End If
                                                    End If

                                                Case "REAL"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 4
                                                                iaValue(iaValueIndex) = BitConverter.ToSingle(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToSingle(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "STRING"
                                                    Dim shortStrLenght As Short
                                                    ' Indirizzo di partenza
                                                    If fORG_ID.Length > 2 Then
                                                        shortStartAddress = shortAddress
                                                        Try
                                                            shortStrLenght = CShort(fORG_ID(2)) + 2 'Sono i byte di controllo
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Nr di Stringhe e Lunghezza dati
                                                    If fORG_ID.Length > 3 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(3))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * shortStrLenght))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)
                                                    Array.Reverse(byteReceive)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        ' Attenzione, li devo girare
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = 0
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step shortStrLenght
                                                                strTemp = enc.GetString(byteReceive, 2 + iIndice_1, byteReceive(1 + iIndice_1))
                                                                iaValue(iaValueIndex) = strTemp
                                                                iaValueIndex = iaValueIndex + 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            strTemp = enc.GetString(byteReceive, 2, byteReceive(1))
                                                            Item.m_objValue = strTemp
                                                        End If
                                                    End If

                                                Case "TIME"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 4
                                                                iaValue(iaValueIndex) = BitConverter.ToInt32(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToInt32(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "TOD"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 4
                                                                iaValue(iaValueIndex) = BitConverter.ToUInt32(byteReceive, iIndice_1)
                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            Item.m_objValue = BitConverter.ToUInt32(byteReceive, 0)
                                                        End If
                                                    End If

                                                Case "DT"
                                                    Dim byteaYearTemp(1) As Byte
                                                    Dim strDT As String
                                                    Dim dtTemp As DateTime
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteReceive = Array.CreateInstance(GetType(Byte), (iArrayLenght * 8))
                                                    SendFetch(tc, byteORG_ID, byteDBNr, shortStartAddress, byteReceive.Length, byteReceive, byteQualities)

                                                    ' Elaboro i dati
                                                    Item.m_iQualities = byteQualities
                                                    If byteQualities = 192 Then
                                                        ' Nessun errore, prelevo i dati
                                                        If iArrayLenght > 1 Then
                                                            ' Array
                                                            Dim iaValue As Array = Array.CreateInstance(GetType(Object), iArrayLenght)

                                                            iaValueIndex = (iArrayLenght - 1)
                                                            For iIndice_1 As Integer = 0 To (byteReceive.Length - 1) Step 8
                                                                strDT = ""
                                                                ' Anno
                                                                byteaYearTemp(0) = byteReceive(iIndice_1 + 7) >> 4
                                                                byteaYearTemp(1) = (byteReceive(iIndice_1 + 7) << 4) >> 4
                                                                byteTemp = (byteaYearTemp(0) * 10) + byteaYearTemp(1) ' Anno 90...89 1990-2089
                                                                If byteTemp >= 90 And byteTemp <= 99 Then
                                                                    strDT = strDT + (1900 + byteTemp).ToString()
                                                                Else
                                                                    strDT = strDT + (2000 + byteTemp).ToString()
                                                                End If
                                                                strDT = strDT + "-"
                                                                ' Mese
                                                                strDT = strDT + (byteReceive(iIndice_1 + 6) >> 4).ToString()
                                                                strDT = strDT + ((byteReceive(iIndice_1 + 6) << 4) >> 4).ToString()
                                                                strDT = strDT + "-"
                                                                ' Giorno
                                                                strDT = strDT + (byteReceive(iIndice_1 + 5) >> 4).ToString()
                                                                strDT = strDT + ((byteReceive(iIndice_1 + 5) << 4) >> 4).ToString()
                                                                strDT = strDT + " "
                                                                ' Ore
                                                                strDT = strDT + (byteReceive(iIndice_1 + 4) >> 4).ToString()
                                                                strDT = strDT + ((byteReceive(iIndice_1 + 4) << 4) >> 4).ToString()
                                                                strDT = strDT + ":"
                                                                ' Minuti
                                                                strDT = strDT + (byteReceive(iIndice_1 + 3) >> 4).ToString()
                                                                strDT = strDT + ((byteReceive(iIndice_1 + 3) << 4) >> 4).ToString()
                                                                strDT = strDT + ":"
                                                                ' Secondi
                                                                strDT = strDT + (byteReceive(iIndice_1 + 2) >> 4).ToString()
                                                                strDT = strDT + ((byteReceive(iIndice_1 + 2) << 4) >> 4).ToString()
                                                                Try
                                                                    dtTemp = strDT
                                                                    iaValue(iaValueIndex) = dtTemp
                                                                Catch ex As Exception
                                                                    ' Errore
                                                                    Item.m_iQualities = 84
                                                                    Exit Select
                                                                End Try

                                                                iaValueIndex = iaValueIndex - 1
                                                            Next iIndice_1
                                                            Item.m_objValue = iaValue
                                                        Else
                                                            ' Singolo valore
                                                            strDT = ""
                                                            ' Anno
                                                            byteaYearTemp(0) = byteReceive(7) >> 4
                                                            byteaYearTemp(1) = (byteReceive(7) << 4) >> 4
                                                            byteTemp = (byteaYearTemp(0) * 10) + byteaYearTemp(1) ' Anno 90...89 1990-2089
                                                            If byteTemp >= 90 And byteTemp <= 99 Then
                                                                strDT = strDT + (1900 + byteTemp).ToString()
                                                            Else
                                                                strDT = strDT + (2000 + byteTemp).ToString()
                                                            End If
                                                            strDT = strDT + "-"
                                                            ' Mese
                                                            strDT = strDT + (byteReceive(6) >> 4).ToString()
                                                            strDT = strDT + ((byteReceive(6) << 4) >> 4).ToString()
                                                            strDT = strDT + "-"
                                                            ' Giorno
                                                            strDT = strDT + (byteReceive(5) >> 4).ToString()
                                                            strDT = strDT + ((byteReceive(5) << 4) >> 4).ToString()
                                                            strDT = strDT + " "
                                                            ' Ore
                                                            strDT = strDT + (byteReceive(4) >> 4).ToString()
                                                            strDT = strDT + ((byteReceive(4) << 4) >> 4).ToString()
                                                            strDT = strDT + ":"
                                                            ' Minuti
                                                            strDT = strDT + (byteReceive(3) >> 4).ToString()
                                                            strDT = strDT + ((byteReceive(3) << 4) >> 4).ToString()
                                                            strDT = strDT + ":"
                                                            ' Secondi
                                                            strDT = strDT + (byteReceive(2) >> 4).ToString()
                                                            strDT = strDT + ((byteReceive(2) << 4) >> 4).ToString()
                                                            Try
                                                                dtTemp = strDT
                                                            Catch ex As Exception
                                                                ' Errore
                                                                Item.m_iQualities = 84
                                                                Exit Select
                                                            End Try
                                                            Item.m_objValue = dtTemp
                                                        End If
                                                    End If

                                                Case Else
                                                    ' Errore
                                                    Exit Select
                                            End Select
                                        Else
                                            ' Errore
                                            Item.m_iQualities = 4
                                        End If
                                    Else
                                        ' Errore
                                        Item.m_iQualities = 4
                                    End If
                                Catch ex As Exception
                                    ' Errore
                                    Item.m_iQualities = 4
                                End Try
                            Else
                                ' Errore
                                Item.m_iQualities = 4
                            End If
                        Else
                            ' Errore
                            Item.m_iQualities = 4
                        End If

                    Else
                        ' Errore
                        Item.m_iQualities = 4
                    End If
                End If

            Catch ex As Exception
                Item.m_iQualities = 4
            End Try
        Next Item

    End Sub

    Private Sub SendFetch(ByRef tc As Net.Sockets.TcpClient, ByVal byteORG_ID As Byte, ByVal byteDB As Byte, ByVal shortStartAddr As Short, ByVal shortByteLength As Short, ByRef byteReceive As Array, ByRef byteQualities As Byte)
        ' Invio Fetch
        Dim iMaxByteSize As Integer = 2000
        Dim iIndice As Integer

        If Not tc Is Nothing Then

            Dim shortByteLengthRemainingToRead As Short ' Byte rimasti da leggere
            Dim shortByteLengthToRead As Short          ' Byte da leggere per ogni chiamata
            Dim shortByteLengthRead As Short            ' Byte totali letti
            Dim byteTemp As Byte
            Dim byteSendFetch(15) As Byte

            Dim byteaTemp As Byte()

            Dim byteStartHigh As Byte
            Dim byteStartLow As Byte

            shortByteLengthRemainingToRead = shortByteLength
            Do
                ' Devo richiedere al massimo "iMaxByteSize" byte alla volta
                If shortByteLengthRemainingToRead > iMaxByteSize Then
                    shortByteLengthToRead = iMaxByteSize
                Else
                    shortByteLengthToRead = shortByteLengthRemainingToRead
                End If

                Dim byteReceiveFetch As Array = Array.CreateInstance(GetType(Byte), (16 + shortByteLengthToRead))

                byteaTemp = BitConverter.GetBytes(shortStartAddr + shortByteLengthRead)
                If byteaTemp.Length = 2 Then
                    byteStartLow = byteaTemp(0)
                    byteStartHigh = byteaTemp(1)

                    Dim byteLenghtHigh As Byte
                    Dim byteLenghtLow As Byte

                    If Not shortByteLengthToRead < 2 And (byteORG_ID = 1 Or byteORG_ID = 2 Or byteORG_ID = 3 Or byteORG_ID = 4) Then
                        If byteORG_ID = 1 Then
                            byteaTemp = BitConverter.GetBytes(shortByteLengthToRead >> 1)
                        End If
                        If byteORG_ID = 2 Or byteORG_ID = 3 Or byteORG_ID = 4 Then
                            byteaTemp = BitConverter.GetBytes(shortByteLengthToRead)
                        End If

                        If byteaTemp.Length = 2 Then
                            byteLenghtLow = byteaTemp(0)
                            byteLenghtHigh = byteaTemp(1)

                            'Invio il telegramma
                            byteSendFetch(0) = 83     'Codice sistema
                            byteSendFetch(1) = 53     'Codice sistema
                            byteSendFetch(2) = 16     'Lunghezza Header
                            byteSendFetch(3) = 1      'ID OP
                            byteSendFetch(4) = 3      'Lunghezza OP
                            byteSendFetch(5) = 5      'Codice OP 5=Fetch 3=Write
                            byteSendFetch(6) = 3      'Blocco ORG
                            byteSendFetch(7) = 8      'Lunghezza Blocco ORG
                            byteSendFetch(8) = byteORG_ID      'Identificatore ORG 1=DB, 2=M, 3=E, 4=A
                            byteSendFetch(9) = byteDB  'DB da leggere (0-255)
                            byteSendFetch(10) = byteStartHigh     'Start Ind. High
                            byteSendFetch(11) = byteStartLow     'Start Ind. Low
                            byteSendFetch(12) = byteLenghtHigh     'Lungh. High
                            byteSendFetch(13) = byteLenghtLow   'Lungh. Low ORG 1=DB->(1 word = 2 byte), 2=M, 3=E, 4=A->(1 byte = 1 byte)
                            byteSendFetch(14) = 255   'Blocco Libero
                            byteSendFetch(15) = 2     'Lunghezza Blocco Libero
                            Try
                                ' Non ci devono essere altri dati disponibili, se ci sono, li elimino
                                While tc.GetStream().DataAvailable = True
                                    tc.GetStream().ReadByte()
                                End While

                                ' Invio la richiesta al PLC
                                tc.GetStream().Write(byteSendFetch, 0, 16)
                                ' Prelevo i dati
                                iIndice = 0
                                Do
                                    Try
                                        byteTemp = tc.GetStream().ReadByte()
                                        If byteTemp <> -1 Then
                                            byteReceiveFetch(iIndice) = byteTemp
                                            iIndice = iIndice + 1
                                            byteQualities = 192
                                        Else
                                            ' Errore
                                            byteQualities = 24
                                            Exit Do
                                        End If
                                    Catch ex As Exception
                                        ' Timeout
                                        ' Errore
                                        byteQualities = 24
                                        Exit Do
                                    End Try
                                Loop While iIndice < (16 + shortByteLengthToRead)

                                '' Prelevo intestazione
                                'For indice As Integer = 0 To (16) - 1
                                '    Try
                                '        byteTemp = tc.GetStream().ReadByte()
                                '        If byteTemp <> -1 Then
                                '            byteReceiveFetch(indice) = byteTemp
                                '            byteQualities = 192
                                '        Else
                                '            ' Errore
                                '            byteQualities = 24
                                '            Exit For
                                '        End If
                                '    Catch ex As Exception
                                '        ' Timeout
                                '        ' Errore
                                '        byteQualities = 24
                                '        Exit For
                                '    End Try
                                'Next indice

                                ' Controllo intestazione
                                If byteQualities = 192 Then
                                    If byteReceiveFetch(0) = 83 And byteReceiveFetch(1) = 53 And _
                                    byteReceiveFetch(2) = 16 And byteReceiveFetch(3) = 1 And _
                                    byteReceiveFetch(4) = 3 And byteReceiveFetch(5) = 6 And _
                                    byteReceiveFetch(6) = 15 And byteReceiveFetch(7) = 3 And _
                                    byteReceiveFetch(9) = 255 And _
                                    byteReceiveFetch(10) = 7 And byteReceiveFetch(11) = 0 And _
                                    byteReceiveFetch(12) = 0 And byteReceiveFetch(13) = 0 And _
                                    byteReceiveFetch(14) = 0 And byteReceiveFetch(15) = 0 Then
                                        ' Intestazione OK, controllo errore
                                        byteTemp = byteReceiveFetch(8)
                                        If byteTemp = 0 Then
                                            '' Nessun errore, prelevo il resto dei dati
                                            'For indice As Integer = 16 To (16 + shortByteLengthToRead) - 1
                                            '    Try
                                            '        byteTemp = tc.GetStream().ReadByte()
                                            '        If byteTemp <> -1 Then
                                            '            byteReceiveFetch(indice) = byteTemp
                                            '            byteQualities = 192
                                            '        Else
                                            '            ' Errore
                                            '            byteQualities = 24
                                            '            Exit For
                                            '        End If
                                            '    Catch ex As Exception
                                            '        ' Timeout
                                            '        ' Errore
                                            '        byteQualities = 24
                                            '        Exit For
                                            '    End Try
                                            'Next indice
                                            ' Elaboro i dati ricevuti
                                            If byteQualities = 192 Then

                                                Array.ConstrainedCopy(byteReceiveFetch, 16, byteReceive, shortByteLengthRead, shortByteLengthToRead)
                                                ' Incremento il conteggio dei byte ricevuti
                                                shortByteLengthRead = shortByteLengthRead + shortByteLengthToRead

                                            End If
                                        Else
                                            ' Errore
                                            byteQualities = byteTemp
                                        End If
                                    Else
                                        ' Errore
                                        byteQualities = 24
                                    End If
                                End If

                            Catch ex As Exception
                                ' Timeout
                                ' Errore
                                byteQualities = 24
                            End Try
                        Else
                            ' Errore
                            byteQualities = 24
                        End If
                    Else
                        ' Errore
                        byteQualities = 24
                    End If
                Else
                    ' Errore
                    byteQualities = 24
                End If

                ' Incremento il contatore dei byte letti
                shortByteLengthRemainingToRead = shortByteLength - shortByteLengthRead

            Loop While shortByteLengthRemainingToRead > 0 And byteQualities = 192
            If byteQualities = 192 Then
                Array.Reverse(byteReceive)
            Else
                Array.Clear(byteReceive, 0, byteReceive.Length)
            End If
        Else
            ' Errore
            byteQualities = 24
        End If

    End Sub

    Sub WriteItems(ByRef tc As Net.Sockets.TcpClient, ByVal Items As S7Item())

        Dim byteSend As Array
        Dim byteQualities As Byte
        Dim iaValueIndex As Integer

        For Each Item As S7Item In Items

            Try
                Dim bTemp As Boolean
                Dim baTemp() As Boolean
                Dim byteaTemp() As Byte
                Dim ushortTemp As UShort
                Dim iTemp As Integer
                Dim strTemp As String

                Dim iIndice_1 As Integer
                Dim iIndice_2 As Integer

                ' Verifico il tipo di area da cui leggere
                Dim fORG_ID() As String
                Dim byteORG_ID As Byte

                ' Numero del DB su cui eseguire le operazioni
                Dim strDataType As String
                Dim shortAddress As Short

                Dim byteDBNr As Byte
                Dim shortStartAddress As Short

                Dim iArrayLenght As Integer
                Dim iByteLenght As Integer
                Dim iBitsLenght As Integer

                ' Serve per la codifica delle stringhe
                Dim enc As New System.Text.ASCIIEncoding

                strDataType = ""
                iaValueIndex = 0
                If Not Item Is Nothing Then

                    Item.m_iQualities = 0
                    Item.m_dtTimeStamp = Now

                    Dim charSep(1) As Char
                    charSep(0) = ","
                    charSep(1) = "."
                    fORG_ID = Item.Address.Split(charSep)

                    If fORG_ID.Length > 1 Then
                        If SplitDataTypeAndStartAddress(fORG_ID(0), strDataType, shortAddress) = True Then
                            If strDataType.Contains("DB") Or strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO") Then
                                If strDataType.Contains("DB") Then
                                    byteORG_ID = 1
                                End If
                                If strDataType.Contains("M") Then
                                    byteORG_ID = 2
                                End If
                                If strDataType.Contains("PI") Then
                                    byteORG_ID = 3
                                End If
                                If strDataType.Contains("PO") Then
                                    byteORG_ID = 4
                                End If
                                ' Numero DB
                                Try
                                    byteDBNr = CByte(shortAddress)
                                    If (byteDBNr >= 1 And byteDBNr <= 255 And strDataType.Contains("DB")) Or (byteDBNr = 0 And (strDataType.Contains("M") Or strDataType.Contains("PI") Or strDataType.Contains("PO"))) Then
                                        If SplitDataTypeAndStartAddress(fORG_ID(1), strDataType, shortAddress) = True Then
                                            Select Case strDataType
                                                Case "B"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iByteLenght = CShort(fORG_ID(2))
                                                            iArrayLenght = (iByteLenght \ 2)
                                                            If iByteLenght Mod (2) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iByteLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))

                                                    ' Elaboro i dati
                                                    If iByteLenght > 1 Then
                                                        ' Array
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iByteLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = 0
                                                        For iIndice_1 = 0 To (byteSend.Length - 1)
                                                            byteSend(iIndice_1) = iaValue(iaValueIndex)
                                                            iaValueIndex = iaValueIndex + 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        ' Visto che il minimo sono 2 byte, devo prima leggere i due byte e quindi
                                                        ' modificare il byte interessato e riscrivere entrambi
                                                        Dim itemTemp As New S7Item(Item.m_strAddress + ",2")
                                                        Dim aitemTemp(0) As S7Item
                                                        aitemTemp(0) = itemTemp
                                                        ReadItems(tc, aitemTemp)
                                                        If aitemTemp(0).m_iQualities = 192 Then
                                                            byteSend(0) = Item.m_objValue
                                                            byteSend(1) = aitemTemp(0).m_objValue(1)
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = aitemTemp(0).m_iQualities
                                                            Exit Select
                                                        End If
                                                    End If

                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "CHAR"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iByteLenght = CShort(fORG_ID(2))
                                                            iArrayLenght = (iByteLenght \ 2)
                                                            If iByteLenght Mod (2) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iByteLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))

                                                    ' Elaboro i dati
                                                    If iByteLenght > 1 Then
                                                        ' Array
                                                        Dim iaValue As Array = Array.CreateInstance(GetType(Char), iByteLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = 0
                                                        For iIndice_1 = 0 To (iTemp - 1)
                                                            byteaTemp = BitConverter.GetBytes(CChar(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            iaValueIndex = iaValueIndex + 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        ' Visto che il minimo sono 2 byte, devo prima leggere i due byte e quindi
                                                        ' modificare il byte interessato e riscrivere entrambi
                                                        Dim byteaTemp_1() As Byte
                                                        Dim byteaTemp_2() As Byte
                                                        Dim itemTemp As New S7Item(Item.m_strAddress + ",2")
                                                        Dim aitemTemp(0) As S7Item
                                                        aitemTemp(0) = itemTemp
                                                        ReadItems(tc, aitemTemp)
                                                        If aitemTemp(0).m_iQualities = 192 Then
                                                            If Not Item.m_objValue Is Nothing Then
                                                                byteaTemp_1 = BitConverter.GetBytes(CChar(Item.m_objValue))
                                                            Else
                                                                byteaTemp_1 = Array.CreateInstance(GetType(Char), 2)
                                                            End If
                                                            If Not aitemTemp(0).m_objValue Is Nothing Then
                                                                byteaTemp_2 = BitConverter.GetBytes(CChar(aitemTemp(0).m_objValue(1)))
                                                            Else
                                                                byteaTemp_2 = Array.CreateInstance(GetType(Byte), 2)
                                                            End If

                                                            byteSend(0) = byteaTemp_1(0)
                                                            byteSend(1) = byteaTemp_2(0)
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = aitemTemp(0).m_iQualities
                                                            Exit Select
                                                        End If
                                                    End If

                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "W"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 2)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 2
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CUShort(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CUShort(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "INT"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))
                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 2)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 2
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If

                                                            byteaTemp = BitConverter.GetBytes(CShort(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CShort(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "DWORD"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))

                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 4
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CUInt(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            byteSend(iIndice_1 + 2) = byteaTemp(2)
                                                            byteSend(iIndice_1 + 3) = byteaTemp(3)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CUInt(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "DINT"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))

                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 4
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CInt(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            byteSend(iIndice_1 + 2) = byteaTemp(2)
                                                            byteSend(iIndice_1 + 3) = byteaTemp(3)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CInt(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "X"

                                                    Dim shortBitNr As Short
                                                    ' Indirizzo di partenza
                                                    If fORG_ID.Length > 2 Then
                                                        shortStartAddress = shortAddress
                                                        Try
                                                            shortBitNr = CShort(fORG_ID(2))
                                                            If Not (shortBitNr >= 0 And shortBitNr) <= 7 Then
                                                                ' Errore
                                                                Item.m_iQualities = 4
                                                                Exit Select
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 3 Then
                                                        If shortBitNr = 0 Then
                                                            iBitsLenght = CShort(fORG_ID(3).ToUpper())
                                                            iArrayLenght = (iBitsLenght \ 16)
                                                            If iBitsLenght Mod (16) <> 0 Then
                                                                iArrayLenght = iArrayLenght + 1
                                                            End If
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End If
                                                    Else
                                                        iBitsLenght = 1
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 2))

                                                    ' Elaboro i dati
                                                    If iBitsLenght > 1 Then
                                                        ' Devo leggere sempre multipli di 2 byte
                                                        Dim itemTemp As New S7Item(Item.m_strAddress.Replace("," + iBitsLenght.ToString(), "," + ((iArrayLenght * 2) * 8).ToString()))
                                                        Dim aitemTemp(0) As S7Item
                                                        aitemTemp(0) = itemTemp
                                                        ReadItems(tc, aitemTemp)
                                                        If aitemTemp(0).m_iQualities = 192 Then
                                                            If Not Item.m_objValue Is Nothing Then
                                                                If Item.m_objValue.Length > iBitsLenght Then
                                                                    iTemp = iBitsLenght
                                                                Else
                                                                    iTemp = Item.m_objValue.Length
                                                                End If
                                                                For iIndice_1 = 0 To iBitsLenght - 1
                                                                    aitemTemp(0).m_objValue(iIndice_1) = False
                                                                Next iIndice_1
                                                                For iIndice_1 = 0 To iTemp - 1
                                                                    aitemTemp(0).m_objValue(iIndice_1) = Item.m_objValue(iIndice_1)
                                                                Next iIndice_1
                                                            Else
                                                                baTemp = Array.CreateInstance(GetType(Boolean), iBitsLenght)
                                                                For iIndice_1 = 0 To iBitsLenght - 1
                                                                    aitemTemp(0).m_objValue(iIndice_1) = Item.m_objValue(iIndice_1)
                                                                Next iIndice_1
                                                            End If

                                                            For iIndice_1 = 0 To ((iArrayLenght * 2) - 1) Step 2
                                                                ushortTemp = 0
                                                                For iIndice_2 = 0 To 15
                                                                    bTemp = aitemTemp(0).m_objValue((iIndice_1 * 8) + iIndice_2)
                                                                    If bTemp = True Then
                                                                        ushortTemp = ushortTemp + (2 ^ iIndice_2)
                                                                    End If
                                                                Next iIndice_2
                                                                byteaTemp = BitConverter.GetBytes(ushortTemp)
                                                                Array.ConstrainedCopy(byteaTemp, 0, byteSend, iIndice_1, 2)
                                                            Next iIndice_1
                                                            iIndice_1 = iIndice_1
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = aitemTemp(0).m_iQualities
                                                            Exit Select
                                                        End If
                                                    Else
                                                        ' Singolo valore
                                                        ' Visto che il minimo sono 2 byte, devo prima leggere i due byte e quindi
                                                        ' modificare il byte interessato e riscrivere entrambi
                                                        Dim itemTemp As New S7Item(Item.Address.Replace("." + shortBitNr.ToString(), ".0,16"))
                                                        Dim aitemTemp(0) As S7Item
                                                        aitemTemp(0) = itemTemp
                                                        ReadItems(tc, aitemTemp)
                                                        If aitemTemp(0).m_iQualities = 192 Then
                                                            If Not Item.m_objValue Is Nothing Then
                                                                aitemTemp(0).m_objValue(shortBitNr) = Item.m_objValue
                                                            Else
                                                                aitemTemp(0).m_objValue(shortBitNr) = CBool(0)
                                                            End If

                                                            ushortTemp = 0
                                                            iIndice_1 = 0
                                                            For Each bTemp In aitemTemp(0).m_objValue
                                                                If bTemp = True Then
                                                                    ushortTemp = ushortTemp + (2 ^ iIndice_1)
                                                                End If
                                                                iIndice_1 = iIndice_1 + 1
                                                            Next bTemp
                                                            byteSend = BitConverter.GetBytes(ushortTemp)
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = aitemTemp(0).m_iQualities
                                                            Exit Select
                                                        End If
                                                    End If

                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "REAL"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))

                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Item.m_objValue.Length > iaValue.Length Then
                                                            iTemp = iaValue.Length
                                                        Else
                                                            iTemp = Item.m_objValue.Length
                                                        End If
                                                        Array.Copy(Item.m_objValue, iaValue, iTemp)

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 4
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CSng(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            byteSend(iIndice_1 + 2) = byteaTemp(2)
                                                            byteSend(iIndice_1 + 3) = byteaTemp(3)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CSng(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "STRING"
                                                    Dim shortStrLenght As Short
                                                    ' Indirizzo di partenza
                                                    If fORG_ID.Length > 2 Then
                                                        shortStartAddress = shortAddress
                                                        Try
                                                            ' La lunghezza deve sempre essere pari
                                                            shortStrLenght = CShort(fORG_ID(2)) + 2 'Sono i byte di controllo
                                                            If shortStrLenght Mod 2 <> 0 Then
                                                                shortStrLenght = shortStrLenght + 1
                                                            End If
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Nr di Stringhe e Lunghezza dati
                                                    If fORG_ID.Length > 3 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(3))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * shortStrLenght))

                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        Dim iaValue As Array = Array.CreateInstance(GetType(String), iArrayLenght)
                                                        If Item.m_objValue.Length > iaValue.Length Then
                                                            iTemp = iaValue.Length
                                                        Else
                                                            iTemp = Item.m_objValue.Length
                                                        End If
                                                        Array.Copy(Item.m_objValue, iaValue, iTemp)

                                                        iaValueIndex = 0
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step shortStrLenght
                                                            If iaValueIndex < iaValue.Length Then
                                                                If Not iaValue(iaValueIndex) Is Nothing Then
                                                                    byteaTemp = enc.GetBytes(iaValue(iaValueIndex))
                                                                    iTemp = byteaTemp.Length
                                                                    If iTemp > (shortStrLenght - 2) Then
                                                                        iTemp = (shortStrLenght - 2)
                                                                    End If
                                                                    byteSend(iIndice_1) = shortStrLenght - 2
                                                                    byteSend(1 + iIndice_1) = iTemp
                                                                    Array.ConstrainedCopy(byteaTemp, 0, byteSend, (2 + iIndice_1), iTemp)
                                                                Else
                                                                    byteSend(iIndice_1) = shortStrLenght - 2
                                                                    byteSend(1 + iIndice_1) = 0
                                                                End If
                                                            End If
                                                            iaValueIndex = iaValueIndex + 1
                                                        Next iIndice_1

                                                    Else
                                                        ' Singolo valore
                                                        If Not Item.m_objValue Is Nothing Then
                                                            byteaTemp = enc.GetBytes(Item.m_objValue)
                                                            iTemp = byteaTemp.Length
                                                            If iTemp > (shortStrLenght - 2) Then
                                                                iTemp = (shortStrLenght - 2)
                                                            End If
                                                            byteSend(0) = shortStrLenght - 2
                                                            byteSend(1) = iTemp
                                                            Array.ConstrainedCopy(byteaTemp, 0, byteSend, 2, iTemp)
                                                        End If
                                                    End If

                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "TIME"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 4
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CInt(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            byteSend(iIndice_1 + 2) = byteaTemp(2)
                                                            byteSend(iIndice_1 + 3) = byteaTemp(3)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CInt(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "TOD"
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 4))
                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 4
                                                            If Not IsNumeric(iaValue(iaValueIndex)) Then
                                                                iaValue(iaValueIndex) = 0
                                                            End If
                                                            byteaTemp = BitConverter.GetBytes(CUInt(iaValue(iaValueIndex)))
                                                            byteSend(iIndice_1) = byteaTemp(0)
                                                            byteSend(iIndice_1 + 1) = byteaTemp(1)
                                                            byteSend(iIndice_1 + 2) = byteaTemp(2)
                                                            byteSend(iIndice_1 + 3) = byteaTemp(3)
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not IsNumeric(Item.m_objValue) Then
                                                            Item.m_objValue = 0
                                                        End If
                                                        byteSend = BitConverter.GetBytes(CUInt(Item.m_objValue))
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case "DT"
                                                    Dim byteaYearTemp(1) As Byte
                                                    ' Indirizzo di partenza
                                                    shortStartAddress = shortAddress

                                                    ' Lunghezza dati
                                                    If fORG_ID.Length > 2 Then
                                                        Try
                                                            iArrayLenght = CShort(fORG_ID(2))
                                                        Catch ex As Exception
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End Try
                                                    Else
                                                        iArrayLenght = 1
                                                    End If
                                                    If iArrayLenght <= 0 Then
                                                        ' Errore
                                                        Item.m_iQualities = 4
                                                        Exit Select
                                                    End If

                                                    ' Imposto il buffer di ricezione
                                                    byteSend = Array.CreateInstance(GetType(Byte), (iArrayLenght * 8))

                                                    ' Elaboro i dati
                                                    ' Nessun errore, prelevo i dati
                                                    If iArrayLenght > 1 Then
                                                        ' Array
                                                        byteaTemp = Array.CreateInstance(GetType(Byte), 4)
                                                        Dim iaValue As Array = Array.CreateInstance(Item.m_objValue(0).GetType, iArrayLenght)
                                                        If Not Item.m_objValue Is Nothing Then
                                                            If Item.m_objValue.Length > iaValue.Length Then
                                                                iTemp = iaValue.Length
                                                            Else
                                                                iTemp = Item.m_objValue.Length
                                                            End If
                                                            Array.Copy(Item.m_objValue, iaValue, iTemp)
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End If

                                                        iaValueIndex = (iArrayLenght - 1)
                                                        For iIndice_1 = 0 To (byteSend.Length - 1) Step 8
                                                            If Not iaValue(iaValueIndex) Is Nothing Then
                                                                ' Anno
                                                                If iaValue(iaValueIndex).Year >= 1 And iaValue(iaValueIndex).Year <= 2089 Then
                                                                    ' Anno
                                                                    strTemp = iaValue(iaValueIndex).Year.ToString().PadLeft(4, "0")
                                                                    byteSend(7 + iIndice_1) = 0
                                                                    byteSend(7 + iIndice_1) = byteSend(7 + iIndice_1) + CByte(strTemp.Substring(2, 1))
                                                                    byteSend(7 + iIndice_1) = byteSend(7 + iIndice_1) << 4
                                                                    byteSend(7 + iIndice_1) = byteSend(7 + iIndice_1) + CByte(strTemp.Substring(3, 1))
                                                                    ' Mese
                                                                    strTemp = iaValue(iaValueIndex).Month.ToString().PadLeft(2, "0")
                                                                    byteSend(6 + iIndice_1) = 0
                                                                    byteSend(6 + iIndice_1) = byteSend(6 + iIndice_1) + CByte(strTemp.Substring(0, 1))
                                                                    byteSend(6 + iIndice_1) = byteSend(6 + iIndice_1) << 4
                                                                    byteSend(6 + iIndice_1) = byteSend(6 + iIndice_1) + CByte(strTemp.Substring(1, 1))
                                                                    ' Giorno
                                                                    strTemp = iaValue(iaValueIndex).Day.ToString().PadLeft(2, "0")
                                                                    byteSend(5 + iIndice_1) = 0
                                                                    byteSend(5 + iIndice_1) = byteSend(5 + iIndice_1) + CByte(strTemp.Substring(0, 1))
                                                                    byteSend(5 + iIndice_1) = byteSend(5 + iIndice_1) << 4
                                                                    byteSend(5 + iIndice_1) = byteSend(5 + iIndice_1) + CByte(strTemp.Substring(1, 1))
                                                                    ' Ore
                                                                    strTemp = iaValue(iaValueIndex).Hour.ToString().PadLeft(2, "0")
                                                                    byteSend(4 + iIndice_1) = 0
                                                                    byteSend(4 + iIndice_1) = byteSend(4 + iIndice_1) + CByte(strTemp.Substring(0, 1))
                                                                    byteSend(4 + iIndice_1) = byteSend(4 + iIndice_1) << 4
                                                                    byteSend(4 + iIndice_1) = byteSend(4 + iIndice_1) + CByte(strTemp.Substring(1, 1))
                                                                    ' Minuti
                                                                    strTemp = iaValue(iaValueIndex).Minute.ToString().PadLeft(2, "0")
                                                                    byteSend(3 + iIndice_1) = 0
                                                                    byteSend(3 + iIndice_1) = byteSend(3 + iIndice_1) + CByte(strTemp.Substring(0, 1))
                                                                    byteSend(3 + iIndice_1) = byteSend(3 + iIndice_1) << 4
                                                                    byteSend(3 + iIndice_1) = byteSend(3 + iIndice_1) + CByte(strTemp.Substring(1, 1))
                                                                    ' Secondi
                                                                    strTemp = iaValue(iaValueIndex).Second.ToString().PadLeft(2, "0")
                                                                    byteSend(2 + iIndice_1) = 0
                                                                    byteSend(2 + iIndice_1) = byteSend(2 + iIndice_1) + CByte(strTemp.Substring(0, 1))
                                                                    byteSend(2 + iIndice_1) = byteSend(2 + iIndice_1) << 4
                                                                    byteSend(2 + iIndice_1) = byteSend(2 + iIndice_1) + CByte(strTemp.Substring(1, 1))
                                                                    ' Millisec 2 cifre piu' significative
                                                                    byteSend(1 + iIndice_1) = 0
                                                                    byteSend(1 + iIndice_1) = byteSend(1 + iIndice_1) + 0
                                                                    byteSend(1 + iIndice_1) = byteSend(1 + iIndice_1) << 4
                                                                    byteSend(1 + iIndice_1) = byteSend(1 + iIndice_1) + 0
                                                                    ' Millisec cifra meno significativa
                                                                    byteSend(0 + iIndice_1) = 0
                                                                    byteSend(0 + iIndice_1) = byteSend(0 + iIndice_1) + 0
                                                                    ' Giorno dell'anno
                                                                    byteSend(0 + iIndice_1) = byteSend(0 + iIndice_1) << 4
                                                                    byteSend(0 + iIndice_1) = byteSend(0 + iIndice_1) + CByte(CByte(Item.Value.DayOfWeek) + 1)
                                                                Else
                                                                    ' Errore
                                                                    Item.m_iQualities = 4
                                                                    Exit Select
                                                                End If
                                                            Else
                                                                ' Errore
                                                                Item.m_iQualities = 4
                                                                Exit Select
                                                            End If
                                                            iaValueIndex = iaValueIndex - 1
                                                        Next iIndice_1
                                                    Else
                                                        ' Singolo valore
                                                        If Not Item.m_objValue Is Nothing Then
                                                            ' Anno
                                                            If Item.m_objValue.Year >= 1990 And Item.m_objValue.Year <= 2089 Then
                                                                ' Anno
                                                                strTemp = Item.m_objValue.Year.ToString().PadLeft(4, "0")
                                                                byteSend(7) = 0
                                                                byteSend(7) = byteSend(7) + CByte(strTemp.Substring(2, 1))
                                                                byteSend(7) = byteSend(7) << 4
                                                                byteSend(7) = byteSend(7) + CByte(strTemp.Substring(3, 1))
                                                                ' Mese
                                                                strTemp = Item.m_objValue.Month.ToString().PadLeft(2, "0")
                                                                byteSend(6) = 0
                                                                byteSend(6) = byteSend(6) + CByte(strTemp.Substring(0, 1))
                                                                byteSend(6) = byteSend(6) << 4
                                                                byteSend(6) = byteSend(6) + CByte(strTemp.Substring(1, 1))
                                                                ' Giorno
                                                                strTemp = Item.m_objValue.Day.ToString().PadLeft(2, "0")
                                                                byteSend(5) = 0
                                                                byteSend(5) = byteSend(5) + CByte(strTemp.Substring(0, 1))
                                                                byteSend(5) = byteSend(5) << 4
                                                                byteSend(5) = byteSend(5) + CByte(strTemp.Substring(1, 1))
                                                                ' Ore
                                                                strTemp = Item.m_objValue.Hour.ToString().PadLeft(2, "0")
                                                                byteSend(4) = 0
                                                                byteSend(4) = byteSend(4) + CByte(strTemp.Substring(0, 1))
                                                                byteSend(4) = byteSend(4) << 4
                                                                byteSend(4) = byteSend(4) + CByte(strTemp.Substring(1, 1))
                                                                ' Minuti
                                                                strTemp = Item.m_objValue.Minute.ToString().PadLeft(2, "0")
                                                                byteSend(3) = 0
                                                                byteSend(3) = byteSend(3) + CByte(strTemp.Substring(0, 1))
                                                                byteSend(3) = byteSend(3) << 4
                                                                byteSend(3) = byteSend(3) + CByte(strTemp.Substring(1, 1))
                                                                ' Secondi
                                                                strTemp = Item.m_objValue.Second.ToString().PadLeft(2, "0")
                                                                byteSend(2) = 0
                                                                byteSend(2) = byteSend(2) + CByte(strTemp.Substring(0, 1))
                                                                byteSend(2) = byteSend(2) << 4
                                                                byteSend(2) = byteSend(2) + CByte(strTemp.Substring(1, 1))
                                                                ' Millisec 2 cifre piu' significative
                                                                byteSend(1) = 0
                                                                byteSend(1) = byteSend(1) + 0
                                                                byteSend(1) = byteSend(1) << 4
                                                                byteSend(1) = byteSend(1) + 0
                                                                ' Millisec cifra meno significativa
                                                                byteSend(0) = 0
                                                                byteSend(0) = byteSend(0) + 0
                                                                ' Giorno dell'anno
                                                                byteSend(0) = byteSend(0) << 4
                                                                byteSend(0) = byteSend(0) + CByte(CByte(Item.Value.DayOfWeek) + 1)
                                                            Else
                                                                ' Errore
                                                                Item.m_iQualities = 4
                                                                Exit Select
                                                            End If
                                                        Else
                                                            ' Errore
                                                            Item.m_iQualities = 4
                                                            Exit Select
                                                        End If
                                                    End If

                                                    Array.Reverse(byteSend)
                                                    SendWrite(tc, byteORG_ID, byteDBNr, shortStartAddress, byteSend.Length, byteSend, byteQualities)

                                                    'Verifico errore
                                                    Item.m_iQualities = byteQualities

                                                Case Else
                                                    ' Errore
                                                    Exit Select
                                            End Select
                                        Else
                                            ' Errore
                                            Item.m_iQualities = 4
                                        End If
                                    Else
                                        ' Errore
                                        Item.m_iQualities = 4
                                    End If
                                Catch ex As Exception
                                    ' Errore
                                    Item.m_iQualities = 4
                                End Try
                            Else
                                ' Errore
                                Item.m_iQualities = 4
                            End If
                        Else
                            ' Errore
                            Item.m_iQualities = 4
                        End If
                    Else
                        ' Errore
                        Item.m_iQualities = 4
                    End If
                End If

            Catch ex As Exception
                Item.m_iQualities = 4
            End Try
        Next Item

    End Sub

    Private Sub SendWrite(ByRef tc As Net.Sockets.TcpClient, ByVal byteORG_ID As Byte, ByVal byteDB As Byte, ByVal shortStartAddr As Short, ByVal shortByteLength As Short, ByVal byteSend As Array, ByRef byteQualities As Byte)

        ' Invio Fetch
        Dim iMaxByteSize As Integer = 2000
        Dim iIndice As Integer

        If Not tc Is Nothing Then

            Dim shortByteLengthRemainingToWrite As Short    ' Byte rimasti da scrivere
            Dim shortByteLengthToWrite As Short             ' Byte da scrivere per ogni chiamata
            Dim shortByteLengthWritten As Short             ' Byte totali scritti
            Dim byteTemp As Byte
            Dim byteReceiveWrite(15) As Byte

            Dim byteaTemp As Byte()

            Dim byteStartHigh As Byte
            Dim byteStartLow As Byte

            shortByteLengthRemainingToWrite = shortByteLength
            Do
                ' Devo richiedere al massimo "iMaxByteSize" byte alla volta
                If shortByteLengthRemainingToWrite > iMaxByteSize Then
                    shortByteLengthToWrite = iMaxByteSize
                Else
                    shortByteLengthToWrite = shortByteLengthRemainingToWrite
                End If

                Dim byteSendWrite As Array = Array.CreateInstance(GetType(Byte), (16 + shortByteLengthToWrite))

                byteaTemp = BitConverter.GetBytes(shortStartAddr + shortByteLengthWritten)
                If byteaTemp.Length = 2 Then
                    byteStartLow = byteaTemp(0)
                    byteStartHigh = byteaTemp(1)

                    Dim byteLenghtHigh As Byte
                    Dim byteLenghtLow As Byte

                    If Not shortByteLengthToWrite < 2 And (byteORG_ID = 1 Or byteORG_ID = 2 Or byteORG_ID = 3 Or byteORG_ID = 4) Then
                        If byteORG_ID = 1 Then
                            byteaTemp = BitConverter.GetBytes(shortByteLengthToWrite >> 1)
                        End If
                        If byteORG_ID = 2 Or byteORG_ID = 3 Or byteORG_ID = 4 Then
                            byteaTemp = BitConverter.GetBytes(shortByteLengthToWrite)
                        End If

                        If byteaTemp.Length = 2 Then
                            byteLenghtLow = byteaTemp(0)
                            byteLenghtHigh = byteaTemp(1)

                            'Invio il telegramma
                            byteSendWrite(0) = 83     'Codice sistema
                            byteSendWrite(1) = 53     'Codice sistema
                            byteSendWrite(2) = 16     'Lunghezza Header
                            byteSendWrite(3) = 1      'ID OP
                            byteSendWrite(4) = 3      'Lunghezza OP
                            byteSendWrite(5) = 3      'Codice OP 5=Fetch 3=Write
                            byteSendWrite(6) = 3      'Blocco ORG
                            byteSendWrite(7) = 8      'Lunghezza Blocco ORG
                            byteSendWrite(8) = byteORG_ID      'Identificatore ORG 1=DB, 2=M, 3=E, 4=A
                            byteSendWrite(9) = byteDB  'DB da leggere (0-255)
                            byteSendWrite(10) = byteStartHigh     'Start Ind. High
                            byteSendWrite(11) = byteStartLow     'Start Ind. Low
                            byteSendWrite(12) = byteLenghtHigh     'Lungh. High
                            byteSendWrite(13) = byteLenghtLow     'Lungh. Low (1 word = 2 byte)
                            byteSendWrite(14) = 255   'Blocco Libero
                            byteSendWrite(15) = 2     'Lunghezza Blocco Libero
                            ' Copio i dati da scrivere
                            Array.ConstrainedCopy(byteSend, shortByteLengthWritten, byteSendWrite, 16, shortByteLengthToWrite)
                            Try
                                ' Non ci devono essere altri dati disponibili, se ci sono, li elimino
                                While tc.GetStream().DataAvailable = True
                                    tc.GetStream().ReadByte()
                                End While

                                ' Invio la richiesta al PLC
                                tc.GetStream().Write(byteSendWrite, 0, byteSendWrite.Length)

                                ' Prelevo i dati
                                byteSendWrite = Array.CreateInstance(GetType(Byte), 16)

                                ' Prelevo i dati
                                iIndice = 0
                                Do
                                    Try
                                        byteTemp = tc.GetStream().ReadByte()
                                        If byteTemp <> -1 Then
                                            byteSendWrite(iIndice) = byteTemp
                                            iIndice = iIndice + 1
                                            byteQualities = 192
                                        Else
                                            ' Errore
                                            byteQualities = 24
                                            Exit Do
                                        End If
                                    Catch ex As Exception
                                        ' Timeout
                                        ' Errore
                                        byteQualities = 24
                                        Exit Do
                                    End Try
                                Loop While iIndice < (16)

                                '' Prelevo intestazione
                                'For indice As Integer = 0 To 15
                                '    Try
                                '        byteTemp = tc.GetStream().ReadByte()
                                '        If byteTemp <> -1 Then
                                '            byteSendWrite(indice) = byteTemp
                                '            byteQualities = 192
                                '        Else
                                '            ' Errore
                                '            byteQualities = 24
                                '            Exit For
                                '        End If
                                '    Catch ex As Exception
                                '        ' Timeout
                                '        ' Errore
                                '        byteQualities = 24
                                '        Exit For
                                '    End Try
                                'Next indice

                                ' Controllo intestazione
                                If byteQualities = 192 Then
                                    If byteSendWrite(0) = 83 And byteSendWrite(1) = 53 And _
                                    byteSendWrite(2) = 16 And byteSendWrite(3) = 1 And _
                                    byteSendWrite(4) = 3 And byteSendWrite(5) = 4 And _
                                    byteSendWrite(6) = 15 And byteSendWrite(7) = 3 And _
                                    byteSendWrite(9) = 255 And _
                                    byteSendWrite(10) = 7 And byteSendWrite(11) = 0 And _
                                    byteSendWrite(12) = 0 And byteSendWrite(13) = 0 And _
                                    byteSendWrite(14) = 0 And byteSendWrite(15) = 0 Then
                                        ' Intestazione OK, controllo errore
                                        byteTemp = byteSendWrite(8)
                                        If byteTemp = 0 Then
                                            ' Nessun errore
                                            ' Incremento il conteggio dei byte inviati
                                            shortByteLengthWritten = shortByteLengthWritten + shortByteLengthToWrite
                                        Else
                                            ' Errore
                                            byteQualities = 24
                                        End If
                                    Else
                                        ' Errore
                                        byteQualities = 24
                                    End If
                                End If
                            Catch ex As Exception
                                ' Timeout
                                ' Errore
                                byteQualities = 24
                            End Try
                        Else
                            ' Errore
                            byteQualities = 24
                        End If
                    Else
                        ' Errore
                        byteQualities = 24
                    End If
                Else
                    ' Errore
                    byteQualities = 24
                End If

                ' Incremento il contatore dei byte letti
                shortByteLengthRemainingToWrite = shortByteLength - shortByteLengthWritten

            Loop While shortByteLengthRemainingToWrite > 0 And byteQualities = 192
        Else
            ' Errore
            byteQualities = 24
        End If

    End Sub

    Public Function SplitDataTypeAndStartAddress(ByVal strToSplit As String, ByRef strDataType As String, ByRef shortAddress As Short)
        Dim bRes As Boolean
        Dim strTemp As String
        strTemp = ""
        strDataType = ""
        shortAddress = 0
        For Each chTemp As Char In strToSplit
            If Char.IsNumber(chTemp) = True Then
                strTemp = strTemp + chTemp
            Else
                strDataType = strDataType + chTemp
            End If
        Next chTemp

        strDataType = strDataType.ToUpper()

        If strTemp.Length > 0 Then
            Try
                shortAddress = CShort(strTemp)
                bRes = True
            Catch ex As Exception

            End Try
        Else
            bRes = True
        End If

        Return bRes
    End Function
End Module
