Imports System.Net.Sockets
Imports System.Windows.Forms
Imports System.Drawing

Public Class S7Read
    Private m_tcSendFetch As TcpClient
    Private m_strCPUHostNameOrAddress As String
    Private m_iCPUPortFetch As Integer
    Private m_iTimeout As Integer

    Sub Connect(ByVal CPUHostNameOrAddress As String, ByVal CPUPortFetch As Integer, Optional ByVal Timeout As Integer = 1000)
        m_strCPUHostNameOrAddress = CPUHostNameOrAddress
        m_iCPUPortFetch = CPUPortFetch
        m_iTimeout = Timeout
        Close()
        If m_tcSendFetch Is Nothing Then
            m_tcSendFetch = New Net.Sockets.TcpClient
            m_tcSendFetch.NoDelay = True
            m_tcSendFetch.Client.DontFragment = True
            m_tcSendFetch.ReceiveTimeout = m_iTimeout
            m_tcSendFetch.SendTimeout = m_iTimeout
            m_tcSendFetch.Connect(m_strCPUHostNameOrAddress, m_iCPUPortFetch)
        End If
    End Sub

    Sub Close()
        If Not m_tcSendFetch Is Nothing Then
            If Not m_tcSendFetch.Client Is Nothing Then
                m_tcSendFetch.Client.Close()
                m_tcSendFetch.Close()
                m_tcSendFetch = Nothing
            End If
        End If
    End Sub

    Public Sub Read(ByVal Items As S7Item(), Optional ByVal Timeout As Integer = 1000)
        If Not m_tcSendFetch Is Nothing Then
            If m_tcSendFetch.Connected = False Then
                Connect(m_strCPUHostNameOrAddress, m_iCPUPortFetch, m_iTimeout)
            Else
                ReadItems(m_tcSendFetch, Items)
            End If
        End If

    End Sub

End Class
