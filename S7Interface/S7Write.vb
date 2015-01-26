Imports System.Net.Sockets

Public Class S7Write
    Private m_tcSendWrite As TcpClient
    Private m_strCPUHostNameOrAddress As String
    Private m_iCPUPortFetch As Integer
    Private m_iTimeout As Integer

    Sub Connect(ByVal CPUHostNameOrAddress As String, ByVal CPUPortFetch As Integer, Optional ByVal Timeout As Integer = 1000)
        m_strCPUHostNameOrAddress = CPUHostNameOrAddress
        m_iCPUPortFetch = CPUPortFetch
        m_iTimeout = Timeout
        Close()
        If m_tcSendWrite Is Nothing Then
            m_tcSendWrite = New Net.Sockets.TcpClient
            m_tcSendWrite.NoDelay = True
            m_tcSendWrite.Connect(CPUHostNameOrAddress, CPUPortFetch)
        End If
    End Sub

    Sub Close()
        If Not m_tcSendWrite Is Nothing Then
            If Not m_tcSendWrite.Client Is Nothing Then
                m_tcSendWrite.Client.Close()
                m_tcSendWrite.Close()
                m_tcSendWrite = Nothing
            End If
        End If
    End Sub

    Public Sub Write(ByVal Items As S7Item(), Optional ByVal Timeout As Integer = 5000)
        If Not m_tcSendWrite Is Nothing Then
            If m_tcSendWrite.Connected = False Then
                Connect(m_strCPUHostNameOrAddress, m_iCPUPortFetch, m_iTimeout)
            End If
            WriteItems(m_tcSendWrite, Items)
        End If

    End Sub
End Class
