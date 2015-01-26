Imports S7Interface
Imports System.Xml

Public Class CepiS7Interface
    Private m_s7r As New S7Read
    Private m_s7w As New S7Write
    Private m_ai() As S7Item

    Private Sub CepiS7Interface_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F9 Then
            Button_Write_Click(sender, e)
        End If
    End Sub

    Private Sub TestSiemensFWInterface_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dgrv As New DataGridViewRow
        Try
            ' Carico le impostazioni
            Button_Connect.Enabled = True
            Button_Disconnect.Enabled = False
            Button_Write.Enabled = False

            TextBox_Ip_Addr_1.Text = My.Settings.PLC_Ip_Addr_1
            TextBox_Port_1.Text = My.Settings.PLC_Port_1
            TextBox_Port_2.Text = My.Settings.PLC_Port_2

            DataGridView_OPC.Columns.Add("Address", "Address")
            DataGridView_OPC.Columns.Item("Address").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            DataGridView_OPC.Columns.Item("Address").CellTemplate.Style.BackColor = Color.AntiqueWhite

            DataGridView_OPC.Columns.Add("Type", "Type")
            DataGridView_OPC.Columns.Item("Type").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            DataGridView_OPC.Columns.Item("Type").ReadOnly = True

            DataGridView_OPC.Columns.Add("GetValue", "GetValue")
            'DataGridView_OPC.Columns.Item("GetValue").ReadOnly = True

            DataGridView_OPC.Columns.Add("GetValueQualities", "GetValueQualities")
            DataGridView_OPC.Columns.Item("GetValueQualities").ReadOnly = True

            DataGridView_OPC.Columns.Add("GetValueTimeStamp", "GetValueTimeStamp")
            DataGridView_OPC.Columns.Item("GetValueTimeStamp").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            DataGridView_OPC.Columns.Item("GetValueTimeStamp").ReadOnly = True

            DataGridView_OPC.Columns.Add("SetValue", "SetValue")
            DataGridView_OPC.Columns.Item("SetValue").CellTemplate.Style.BackColor = Color.AntiqueWhite

            DataGridView_OPC.Columns.Add("SetValueQualities", "SetValueQualities")
            DataGridView_OPC.Columns.Item("SetValueQualities").ReadOnly = True

            DataGridView_OPC.Columns.Add("SetValueTimeStamp", "SetValueTimeStamp")
            DataGridView_OPC.Columns.Item("SetValueTimeStamp").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            DataGridView_OPC.Columns.Item("SetValueTimeStamp").ReadOnly = True

            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(0).Cells("Address").Value = "DB1,B0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(1).Cells("Address").Value = "DB1,B0,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(2).Cells("Address").Value = "DB1,CHAR130"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(3).Cells("Address").Value = "DB1,CHAR130,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(4).Cells("Address").Value = "DB1,W260"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(5).Cells("Address").Value = "DB1,W260,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(6).Cells("Address").Value = "DB1,INT2310"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(7).Cells("Address").Value = "DB1,INT2310,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(8).Cells("Address").Value = "DB1,DWORD4360"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(9).Cells("Address").Value = "DB1,DWORD4360,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(10).Cells("Address").Value = "DB1,DINT4490"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(11).Cells("Address").Value = "DB1,DINT4490,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(12).Cells("Address").Value = "DB1,X4620.0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(13).Cells("Address").Value = "DB1,X4620.0,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(14).Cells("Address").Value = "DB1,REAL4750"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(15).Cells("Address").Value = "DB1,REAL4750,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(16).Cells("Address").Value = "DB1,STRING8850.16"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(17).Cells("Address").Value = "DB1,STRING8850.16,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(18).Cells("Address").Value = "DB1,TIME27300"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(19).Cells("Address").Value = "DB1,TIME27300,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(20).Cells("Address").Value = "DB1,TOD29352"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(21).Cells("Address").Value = "DB1,TOD29352,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(22).Cells("Address").Value = "DB1,DT31404"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(23).Cells("Address").Value = "DB1,DT31404,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(24).Cells("Address").Value = "M,B255"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(25).Cells("Address").Value = "M,B255,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(26).Cells("Address").Value = "M,W255"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(27).Cells("Address").Value = "M,W255,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(28).Cells("Address").Value = "PI,B0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(29).Cells("Address").Value = "PI,B0,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(30).Cells("Address").Value = "PI,W0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(31).Cells("Address").Value = "PI,W0,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(32).Cells("Address").Value = "PO,B0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(33).Cells("Address").Value = "PO,B0,8"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(34).Cells("Address").Value = "PO,W0"
            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(35).Cells("Address").Value = "PO,W0,8"

            'DataGridView_OPC.Rows.Add()
            'DataGridView_OPC.Rows.Item(0).Cells("Address").Value = "PI,B0"

            DataGridView_OPC.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right

            ProgressBar_OPC.BackColor = Color.Red

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DataGridView_OPC_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_OPC.CellValueChanged
        'Dim s7i(0) As S7Item
        If sender.GetType().Name = "DataGridView" Then
            If e.ColumnIndex = 0 Then
                AddItem()

                If DataGridView_OPC.Rows.Count > 0 Then
                    For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                        If dgvr.Cells.Item("Address").Value Is Nothing Then
                            dgvr.Cells.Item("Type").Value = Nothing
                            dgvr.Cells.Item("GetValue").Value = Nothing
                            dgvr.Cells.Item("GetValueQualities").Value = Nothing
                            dgvr.Cells.Item("GetValueTimeStamp").Value = Nothing

                            dgvr.Cells.Item("SetValue").Value = Nothing
                            dgvr.Cells.Item("SetValueQualities").Value = Nothing
                            dgvr.Cells.Item("SetValueTimeStamp").Value = Nothing
                        End If
                    Next
                End If

            End If

            If e.ColumnIndex = 5 Then
                If DataGridView_OPC.Rows.Count > 0 Then
                    For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                        If dgvr.Cells.Item("SetValue").Value Is Nothing Then
                            dgvr.Cells.Item("SetValueQualities").Value = Nothing
                            dgvr.Cells.Item("SetValueTimeStamp").Value = Nothing
                        End If
                    Next
                End If
            End If

        End If
    End Sub

    Private Sub Timer_OPC_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer_OPC.Tick
        Timer_OPC.Enabled = False


        If Not m_ai Is Nothing Then
            m_s7r.Read(m_ai)
            FillDataGridView()

            If ProgressBar_OPC.Value >= ProgressBar_OPC.Maximum Then
                ProgressBar_OPC.Value = 0
            End If
            ProgressBar_OPC.PerformStep()
        End If

        Timer_OPC.Interval = 100
        Timer_OPC.Enabled = True

    End Sub

    Private Sub Button_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Connect.Click
        Try
            Button_Connect.Enabled = False
            TextBox_Ip_Addr_1.ReadOnly = True
            TextBox_Port_1.ReadOnly = True
            TextBox_Port_2.ReadOnly = True
            TextBox_Ip_Addr_1.Update()
            TextBox_Port_1.Update()
            TextBox_Port_2.Update()

            ResetItem()
            FillDataGridView()

            ' Connessione Read
            m_s7r.Connect(TextBox_Ip_Addr_1.Text, CInt(TextBox_Port_1.Text))

            ' Connessione Write
            m_s7w.Connect(TextBox_Ip_Addr_1.Text, CInt(TextBox_Port_2.Text))

            ProgressBar_OPC.Value = 0

            ' Memorizzo le impostazioni
            My.Settings.PLC_Ip_Addr_1 = TextBox_Ip_Addr_1.Text
            My.Settings.PLC_Port_1 = TextBox_Port_1.Text
            My.Settings.PLC_Port_2 = TextBox_Port_2.Text
            My.Settings.Save()

            ' Abilito Timer
            Timer_OPC.Enabled = True

            Button_Disconnect.Enabled = True
            Button_Write.Enabled = True
        Catch ex As Exception
            Button_Connect.Enabled = True
            Button_Disconnect.Enabled = False
            TextBox_Ip_Addr_1.ReadOnly = False
            TextBox_Port_1.ReadOnly = False
            TextBox_Port_2.ReadOnly = False

            MsgBox("Errore:" & ex.Message)
        End Try


    End Sub

    Private Sub Button_Disconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Disconnect.Click
        Try
            Button_Disconnect.Enabled = False
            Button_Write.Enabled = False
            TextBox_Ip_Addr_1.ReadOnly = False
            TextBox_Port_1.ReadOnly = False
            TextBox_Port_2.ReadOnly = False

            Timer_OPC.Enabled = False

            m_s7r.Close()
            m_s7w.Close()

            ResetItem()
            FillDataGridView()

            ProgressBar_OPC.Value = 0

            Button_Connect.Enabled = True

        Catch ex As Exception
            Button_Connect.Enabled = True
            Button_Disconnect.Enabled = True
            TextBox_Ip_Addr_1.ReadOnly = False
            TextBox_Port_1.ReadOnly = False
            TextBox_Port_2.ReadOnly = False

            MsgBox("Errore:" & ex.Message)
        End Try

    End Sub

    Private Sub Button_Write_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Write.Click
        Dim s7i() As S7Item

        If DataGridView_OPC.SelectedRows.Count = 0 Then
            If DataGridView_OPC.Rows.Count > 0 Then
                s7i = Array.CreateInstance(GetType(S7Item), 0)
                If Not s7i Is Nothing Then
                    For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                        If Not dgvr.Cells.Item("Address").Value Is Nothing And Not dgvr.Cells.Item("SetValue").Value Is Nothing Then
                            Array.Resize(s7i, s7i.Length + 1)
                            s7i(s7i.Length - 1) = New S7Item(dgvr.Cells("Address").Value)

                            If IsArray(s7i(s7i.Length - 1).Value) = True Then
                                Dim iIndice_1 As Integer
                                Dim str_Orig As String
                                Dim stra_Dest() As String
                                Dim charTrim(1) As Char
                                Dim charSep(0) As Char

                                charTrim(0) = "{"
                                charTrim(1) = "}"
                                str_Orig = dgvr.Cells("SetValue").Value.ToString().Trim(charTrim)

                                charSep(0) = ","
                                stra_Dest = str_Orig.Split(charSep)

                                iIndice_1 = 0
                                For Each str As String In stra_Dest
                                    If iIndice_1 < s7i(s7i.Length - 1).Value.Length Then
                                        Try
                                            s7i(s7i.Length - 1).Value(iIndice_1) = ConvertValue(str, s7i(s7i.Length - 1).Value(iIndice_1).GetType())
                                        Catch ex As Exception
                                            MsgBox("Errore di conversione del valore: '" & str & "' in: '" & s7i(s7i.Length - 1).Value(iIndice_1).GetType().Name & "'.")
                                            Exit Sub
                                        End Try
                                    End If
                                    iIndice_1 = iIndice_1 + 1
                                Next
                            Else
                                Dim str_Orig As String

                                str_Orig = dgvr.Cells("SetValue").Value.ToString()
                                Try
                                    s7i(s7i.Length - 1).Value = ConvertValue(str_Orig, s7i(s7i.Length - 1).Value.GetType())
                                Catch ex As Exception
                                    MsgBox("Errore di conversione del valore : " & str_Orig & " in : " & s7i(s7i.Length - 1).Value.GetType().Name)
                                    Exit Sub
                                End Try

                            End If
                        End If
                        If Not dgvr.Cells.Item("SetValueQualities") Is Nothing Then
                            dgvr.Cells.Item("SetValueQualities").Value = Nothing
                        End If
                        If Not dgvr.Cells.Item("SetValueTimeStamp") Is Nothing Then
                            dgvr.Cells.Item("SetValueTimeStamp").Value = Nothing
                        End If
                    Next
                    m_s7w.Write(s7i)
                    For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                        If Not dgvr.Cells.Item("Address") Is Nothing Then
                            For Each s7iTemp As S7Item In s7i
                                If Not s7iTemp Is Nothing Then
                                    If s7iTemp.Address = dgvr.Cells.Item("Address").Value Then

                                        If Not dgvr.Cells.Item("SetValueQualities") Is Nothing Then
                                            dgvr.Cells.Item("SetValueQualities").Value = s7iTemp.Qualities
                                        End If
                                        If Not dgvr.Cells.Item("SetValueTimeStamp") Is Nothing Then
                                            dgvr.Cells.Item("SetValueTimeStamp").Value = s7iTemp.TimeStamp
                                        End If

                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Else
            s7i = Array.CreateInstance(GetType(S7Item), 0)
            If Not s7i Is Nothing Then
                For Each dgvr As DataGridViewRow In DataGridView_OPC.SelectedRows
                    If Not dgvr.Cells.Item("Address").Value Is Nothing And Not dgvr.Cells.Item("SetValue").Value Is Nothing Then
                        Array.Resize(s7i, s7i.Length + 1)
                        s7i(s7i.Length - 1) = New S7Item(dgvr.Cells("Address").Value)

                        If IsArray(s7i(s7i.Length - 1).Value) = True Then
                            Dim iIndice_1 As Integer
                            Dim str_Orig As String
                            Dim stra_Dest() As String
                            Dim charSep(0) As Char

                            str_Orig = dgvr.Cells("SetValue").Value.ToString()

                            charSep(0) = ","
                            stra_Dest = str_Orig.Split(charSep)

                            iIndice_1 = 0
                            For Each str As String In stra_Dest
                                If iIndice_1 < s7i(s7i.Length - 1).Value.Length Then
                                    Try
                                        s7i(s7i.Length - 1).Value(iIndice_1) = ConvertValue(str, s7i(s7i.Length - 1).Value(iIndice_1).GetType())
                                    Catch ex As Exception
                                        MsgBox("Errore di conversione del valore: '" & str & "' in: '" & s7i(s7i.Length - 1).Value(iIndice_1).GetType().Name & "'.")
                                        Exit Sub
                                    End Try
                                End If
                                iIndice_1 = iIndice_1 + 1
                            Next
                        Else
                            Dim str_Orig As String

                            str_Orig = dgvr.Cells("SetValue").Value.ToString()
                            Try
                                s7i(s7i.Length - 1).Value = ConvertValue(str_Orig, s7i(s7i.Length - 1).Value.GetType())
                            Catch ex As Exception
                                MsgBox("Errore di conversione del valore : " & str_Orig & " in : " & s7i(s7i.Length - 1).Value.GetType().Name)
                                Exit Sub
                            End Try

                        End If
                    End If
                    If Not dgvr.Cells.Item("SetValueQualities") Is Nothing Then
                        dgvr.Cells.Item("SetValueQualities").Value = Nothing
                    End If
                    If Not dgvr.Cells.Item("SetValueTimeStamp") Is Nothing Then
                        dgvr.Cells.Item("SetValueTimeStamp").Value = Nothing
                    End If
                Next
                m_s7w.Write(s7i)
                For Each dgvr As DataGridViewRow In DataGridView_OPC.SelectedRows
                    If Not dgvr.Cells.Item("Address") Is Nothing Then
                        For Each s7iTemp As S7Item In s7i
                            If Not s7iTemp Is Nothing Then
                                If s7iTemp.Address = dgvr.Cells.Item("Address").Value Then

                                    If Not dgvr.Cells.Item("SetValueQualities") Is Nothing Then
                                        dgvr.Cells.Item("SetValueQualities").Value = s7iTemp.Qualities
                                    End If
                                    If Not dgvr.Cells.Item("SetValueTimeStamp") Is Nothing Then
                                        dgvr.Cells.Item("SetValueTimeStamp").Value = s7iTemp.TimeStamp
                                    End If

                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub AddItem()
        Dim iIndice_1 As Integer

        If DataGridView_OPC.Rows.Count > 0 Then
            If m_ai Is Nothing Then
                m_ai = Array.CreateInstance(GetType(S7Item), DataGridView_OPC.Rows.Count)
            Else
                Array.Clear(m_ai, 0, m_ai.Length)
                Array.Resize(m_ai, DataGridView_OPC.Rows.Count)
            End If
            iIndice_1 = 0
            For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                If Not dgvr.Cells.Item("Address").Value Is Nothing Then
                    m_ai(iIndice_1) = New S7Item(dgvr.Cells.Item("Address").Value)
                End If
                iIndice_1 = iIndice_1 + 1
            Next
        End If
    End Sub

    Private Sub ResetItem()
        If Not m_ai Is Nothing Then
            For Each itm As S7Item In m_ai
                If Not itm Is Nothing Then
                    itm.InitializeValue()
                End If
            Next
        End If
    End Sub

    Private Sub FillDataGridView()
        If Not m_ai Is Nothing Then
            For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                If Not dgvr.Cells.Item("Address") Is Nothing Then
                    For Each s7i As S7Item In m_ai
                        If Not s7i Is Nothing Then

                            If s7i.Address = dgvr.Cells.Item("Address").Value Then

                                If s7i.TimeStamp.Millisecond = 0 Then
                                    If Not dgvr.Cells.Item("GetValue") Is Nothing Then
                                        dgvr.Cells.Item("GetValue").Value = Nothing
                                    End If

                                    If Not dgvr.Cells.Item("Type") Is Nothing Then
                                        dgvr.Cells.Item("Type").Value = Nothing
                                    End If

                                    If Not dgvr.Cells.Item("GetValueQualities") Is Nothing Then
                                        dgvr.Cells.Item("GetValueQualities").Value = Nothing
                                    End If

                                    If Not dgvr.Cells.Item("GetValueTimeStamp") Is Nothing Then
                                        dgvr.Cells.Item("GetValueTimeStamp").Value = Nothing
                                    End If
                                Else
                                    If Not dgvr.Cells.Item("GetValue") Is Nothing And Not dgvr.Cells.Item("Type") Is Nothing And Not s7i.Value Is Nothing Then
                                        If IsArray(s7i.Value) = True Then
                                            Dim str As String
                                            str = ""
                                            For Each objValue As Object In s7i.Value
                                                If str.Count = 0 Then
                                                    str = str & "{"
                                                Else
                                                    str = str & ","
                                                End If
                                                str = str + objValue.ToString()
                                            Next
                                            str = str & "}"
                                            dgvr.Cells.Item("GetValue").Value = str

                                            dgvr.Cells.Item("Type").Value = s7i.Value(0).GetType().Name.ToString()
                                        Else
                                            dgvr.Cells.Item("GetValue").Value = s7i.Value

                                            dgvr.Cells.Item("Type").Value = s7i.Value.GetType().Name.ToString()
                                        End If
                                    End If

                                    If Not dgvr.Cells.Item("GetValueQualities") Is Nothing Then
                                        dgvr.Cells.Item("GetValueQualities").Value = s7i.Qualities
                                    End If

                                    If Not dgvr.Cells.Item("GetValueTimeStamp") Is Nothing Then
                                        dgvr.Cells.Item("GetValueTimeStamp").Value = s7i.TimeStamp
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub DataGridView_OPC_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView_OPC.RowsRemoved
        AddItem()
    End Sub

    Private Function ConvertValue(ByVal strValue As String, ByVal t As Type)
        Dim obj As Object = Nothing

        Select Case t.Name
            Case Is = "Byte"
                obj = CByte(strValue)

            Case Is = "Char"
                obj = CChar(strValue)

            Case Is = "String"
                obj = CStr(strValue)

            Case Is = "UInt16"
                obj = CUInt(strValue)

            Case Is = "Int16"
                obj = CInt(strValue)

            Case Is = "UInt32"
                obj = CUInt(strValue)

            Case Is = "Int32"
                obj = CInt(strValue)

            Case Is = "Boolean"
                obj = CBool(strValue)

            Case Is = "Single"
                obj = CSng(strValue)

            Case Is = "DateTime"
                obj = CDate(strValue)

        End Select

        Return obj
    End Function

    Private Sub About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles About.Click
        S7Help.About()
    End Sub

    Private Sub Button_Salva_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Salva.Click

        'Dim myXmlTextWriter As XmlTextWriter = New XmlTextWriter("newbooks.xml", System.Text.Encoding.UTF8)
        'myXmlTextWriter.Formatting = System.Xml.Formatting.Indented
        'myXmlTextWriter.WriteStartDocument(False)
        'myXmlTextWriter.WriteComment("This is a comment")
        ''Create the main document element.
        'myXmlTextWriter.WriteStartElement("bookstore")
        'myXmlTextWriter.WriteStartElement("book")

        ''Create an element named 'title' with a text node
        '' and then close the element.
        'myXmlTextWriter.WriteStartElement("title")
        'myXmlTextWriter.WriteString("The Autobiography of Mark Twain")
        'myXmlTextWriter.WriteEndElement()

        ''Create an element named 'Author'.
        'myXmlTextWriter.WriteStartElement("Author")

        ''Create an element named 'first-name' with a text node
        '' and close it in one line.
        'myXmlTextWriter.WriteElementString("first-name", "Mark")

        ''Create an element named 'first-name' with a text node.
        'myXmlTextWriter.WriteElementString("last-name", "Twain")

        ''Close off the parent element.
        'myXmlTextWriter.WriteEndElement()

        ''Create an element named 'price' with a text node
        '' and close it in one line.
        'myXmlTextWriter.WriteElementString("price", "7.99")

        ''Close off the book element.
        'myXmlTextWriter.WriteEndElement()

        'myXmlTextWriter.WriteStartElement("book")
        'myXmlTextWriter.WriteAttributeString("genre", "autobiography")
        'myXmlTextWriter.WriteAttributeString("publicationdate", "1979")
        'myXmlTextWriter.WriteAttributeString("ISBN", "0-7356-0562-9")

        ''Close off the book element.   
        'myXmlTextWriter.WriteEndElement()
        ''Close off the Parent Element bookstore. 
        'myXmlTextWriter.WriteEndElement()

        'myXmlTextWriter.Flush()
        'myXmlTextWriter.Close()
        ''Waits for user to press enter before exiting the program.
        'Console.ReadLine()
        Dim sfd As New SaveFileDialog()

        'sfd.InitialDirectory = My.Settings.OpenSavePath
        sfd.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
        sfd.FilterIndex = 1
        sfd.RestoreDirectory = True
        'sfd.FileName = "S7Items"
        sfd.FileName = My.Settings.OpenSavePath

        Try
            If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                My.Settings.OpenSavePath = sfd.FileName
                My.Settings.Save()

                Try
                    Dim myXmlTextWriter As XmlTextWriter = New XmlTextWriter(sfd.FileName, System.Text.Encoding.UTF8)
                    'Dim myXmlTextWriter As XmlTextWriter = New XmlTextWriter("c:\S7Item.xml", System.Text.Encoding.UTF8)

                    Try
                        Dim bAddressFound As Boolean

                        myXmlTextWriter.Formatting = System.Xml.Formatting.Indented
                        myXmlTextWriter.WriteStartDocument(False)
                        myXmlTextWriter.WriteComment("Elenco Indirizzi")
                        myXmlTextWriter.WriteStartElement("S7Item")

                        bAddressFound = False
                        For Each dgvr As DataGridViewRow In DataGridView_OPC.Rows
                            If Not dgvr.Cells.Item("Address") Is Nothing Then
                                If Not dgvr.Cells.Item("Address").Value Is Nothing Then
                                    myXmlTextWriter.WriteStartElement("Address")

                                    myXmlTextWriter.WriteString(dgvr.Cells.Item("Address").Value.ToString())

                                    myXmlTextWriter.WriteEndElement()

                                    bAddressFound = True
                                End If
                            End If
                        Next

                        If bAddressFound = False Then
                            myXmlTextWriter.WriteString("-No S7Item-")
                        End If

                        myXmlTextWriter.WriteEndElement()

                    Catch ex As Exception
                        MsgBox("Errore:" & ex.Message)
                    End Try

                    myXmlTextWriter.Flush()
                    myXmlTextWriter.Close()

                Catch ex_1 As Exception
                    MsgBox("Errore:" & ex_1.Message)
                    Try
                        My.Settings.OpenSavePath = ""
                        My.Settings.Save()
                    Catch ex_2 As Exception
                        MsgBox("Errore:" & ex_2.Message)
                    End Try
                End Try

            End If

        Catch Ex As Exception
            MessageBox.Show("Errore:" & Ex.Message)
        End Try

    End Sub

    Private Sub Button_Carica_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Carica.Click
        'Dim reader As XmlTextReader = New XmlTextReader("c:\newbooks.xml")

        'Do While (reader.Read())
        '    Select Case reader.NodeType
        '        Case XmlNodeType.Element 'Display beginning of element.
        '            Console.Write("<" + reader.Name)
        '            If reader.HasAttributes Then 'If attributes exist
        '                While reader.MoveToNextAttribute()
        '                    'Display attribute name and value.
        '                    Console.Write(" {0}='{1}'", reader.Name, reader.Value)
        '                End While
        '            End If
        '            Console.WriteLine(">")
        '        Case XmlNodeType.Text 'Display the text in each element.
        '            Console.WriteLine(reader.Value)
        '        Case XmlNodeType.EndElement 'Display end of element.
        '            Console.Write("</" + reader.Name)
        '            Console.WriteLine(">")
        '    End Select
        'Loop
        'Console.ReadLine()
        Dim ofd As New OpenFileDialog()

        'ofd.InitialDirectory = My.Settings.OpenSavePath
        ofd.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
        ofd.FilterIndex = 1
        ofd.RestoreDirectory = True
        'ofd.FileName = "S7Items"
        ofd.FileName = My.Settings.OpenSavePath

        Try

            If ofd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                My.Settings.OpenSavePath = ofd.FileName
                My.Settings.Save()

                Try
                    Dim reader As XmlTextReader = New XmlTextReader(ofd.FileName)
                    'Dim reader As XmlTextReader = New XmlTextReader("c:\S7Item.xml")
                    Try

                        Dim bS7ItemFound As Boolean
                        Dim bAddressFound As Boolean
                        Dim dgr As DataGridViewRow

                        DataGridView_OPC.Rows.Clear()

                        Do While (reader.Read())
                            Select Case reader.NodeType
                                Case XmlNodeType.Element 'Display beginning of element.
                                    If reader.Name = "S7Item" Then
                                        bS7ItemFound = True
                                    End If
                                    If reader.Name = "Address" Then
                                        bAddressFound = True
                                    End If
                                Case XmlNodeType.Text 'Display the text in each element.
                                    If bS7ItemFound = True And bAddressFound = True Then
                                        dgr = New DataGridViewRow
                                        dgr.CreateCells(DataGridView_OPC)
                                        dgr.Cells.Item(0).Value = reader.Value
                                        DataGridView_OPC.Rows.Add(dgr)
                                        dgr.Dispose()
                                    End If
                                Case XmlNodeType.EndElement 'Display end of element.
                                    If reader.Name = "S7Item" Then
                                        bS7ItemFound = False
                                    End If
                                    If reader.Name = "Address" Then
                                        bAddressFound = False
                                    End If
                            End Select
                        Loop

                    Catch ex As Exception
                        MsgBox("Errore:" & ex.Message)
                    End Try

                    reader.Close()

                Catch ex_1 As Exception
                    MsgBox("Errore:" & ex_1.Message)

                    Try
                        My.Settings.OpenSavePath = ""
                        My.Settings.Save()
                    Catch ex_2 As Exception
                        MsgBox("Errore:" & ex_2.Message)
                    End Try

                End Try

                AddItem()
            End If

        Catch Ex As Exception
            MessageBox.Show("Errore:" & Ex.Message)
        End Try

    End Sub

End Class
