Imports System.Threading  ' for threads
Imports System.Net.Sockets  ' for TcpClient and TcpServer
Imports System.Net
Imports System.IO

Public Class clsServerWinsock
    ' This object is created to serve many TCPClients and to receive Unicode messages 
    ' A thread is created to serve each TCPClient
    Const MaxThread As Integer = 500  ' Maximum number of threads that ServerWinsock can handle
    Private oListener As TcpListener ' Variable for TcpListener
    Private bStopListener As Boolean  ' Flag indicating that user wants to dispose this ServerWinsock
    Private ActiveThreads As Integer  ' Number of active threads, i.e. threads that are serving TCPClients
    Private oSocket As Socket
    ' Event that returns the incoming message
    Public Event OnMessage(ByVal IncomingMessage As String, ByVal strIpClient As String)
    Public Function GetSocketServer() As Socket
        Return oSocket
    End Function


    Public Sub New(ByVal PortNo As Integer)
        ' Instantiate a TcpListener on given PortNo
        oListener = New TcpListener(GetIP(), PortNo)
        oListener.Start() ' Start the TcpListener

        ' Create a thread for Sub AcceptConnection
        Dim ServerThread As Thread
        ServerThread = New Thread(AddressOf AcceptConnection)
        ServerThread.Start() ' Run Sub AcceptConnection like a light weight child process
    End Sub
    Public Sub New(ByVal PortNo As Integer, strIPAddress As String)
        ' Instantiate a TcpListener on given PortNo
        oListener = New TcpListener(GetIP(), PortNo)
        oListener.Start() ' Start the TcpListener

        ' Create a thread for Sub AcceptConnection
        Dim ServerThread As Thread
        ServerThread = New Thread(AddressOf AcceptConnection)
        ServerThread.Start() ' Run Sub AcceptConnection like a light weight child process
    End Sub

    Private Function GetIP() As IPAddress
        Dim strHostName As String = Dns.GetHostName
        Dim IpHostEntry As IPHostEntry = Dns.GetHostByName(strHostName)
        Dim IpAddress As IPAddress = IpHostEntry.AddressList(1)
        Return IpAddress
    End Function

    Protected Sub ProcessRequest()
        Dim Buffer(100) As Byte ' used to receive incoming message from TcpClient

        Dim bytes As Integer ' Actual number of bytes read
        Dim RecvMessage As String ' UTF8 text string of the buffer (array of bytes)
        Dim strIPClient As String
        ' Use oThread to reference the thread of this Sub
        'Create a socket
        oSocket = oListener.AcceptSocket
        Dim oThread As Thread
        oThread = Thread.CurrentThread()
        ' Keep looping until user wants to stop
        While Not bStopListener
            If oSocket.Available > 0 Then ' A message has arrived from TcpClient
                ' read the incoming message into a buffer
                'Try

                '    bytes = oSocket.Receive(Buffer, Buffer.Length, SocketFlags.None)

                'Catch ex As SocketException
                '    If ex.SocketErrorCode = SocketError.ConnectionReset OrElse ex.SocketErrorCode = SocketError.WouldBlock OrElse ex.SocketErrorCode = SocketError.IOPending OrElse ex.SocketErrorCode = SocketError.NoBufferSpaceAvailable Then
                '        Thread.Sleep(30)
                '    Else
                '        Throw ex
                '    End If
                'Catch exEx As Exception
                '    '    oSocket.Close()
                '    '    WriteSip2Log("send:" & ex.Message)
                'End Try
                Try
                    bytes = oSocket.Receive(Buffer, Buffer.Length, SocketFlags.None)
                Catch ex As Exception
                    WriteSip2Log("Error:" & ex.Message)
                End Try

                Dim byReturn As Byte() = New Byte(bytes - 1) {}
                Array.Copy(Buffer, byReturn, bytes)
                SyncLock oThread  ' Lock oThread
                    ' Convert the array of bytes (i.e. the buffer) to UTF8 text string
                    RecvMessage = System.Text.Encoding.UTF8.GetString(byReturn)
                    ' Get Ip client
                    strIPClient = CType(oSocket.RemoteEndPoint, IPEndPoint).Address.ToString
                    '' Raise an event to return the message to the program that owns this ServerWinsock
                    RaiseEvent OnMessage(RecvMessage, strIPClient)
                    '  WriteSip2Log("ServerSock:" & RecvMessage.Trim())
                End SyncLock  ' unlock oThread
                '  Exit While
            End If
            ' get out of while loop if TcpClient has disconnected
            If Not oSocket.Connected Then Exit While
        End While
        '  oSocket.Close() ' Close the TcpServer socket
        SyncLock oThread ' Lock oThread
            ActiveThreads -= 1  ' Decrement number of Active Threads
        End SyncLock ' unlock oThread


    End Sub

    Private Sub AcceptConnection()
        'This is the main Sub of ServerWinsock
        ' Keep looping until user wants to stop
        '    WriteSip2Log("Server Accept1")
        Do While Not bStopListener
            Thread.Sleep(100)  ' Sleep 100 msec.
            If Not oListener Is Nothing AndAlso oListener.Pending() Then  ' received a request for connection from a TCPClient
                If ActiveThreads <= MaxThread Then
                    ' WriteSip2Log("Server Accept2")
                    ' create a thread to handle this Client Request
                    Dim oThread As Thread
                    oThread = New Thread(AddressOf ProcessRequest)
                    oThread.Start()  ' Run Sub ProcessRequest
                    SyncLock oThread  ' Lock oThread so that ActiveThreads value is not changed while we're adding 1 to it
                        ActiveThreads += 1  ' Increment number of active threads
                    End SyncLock  ' Release the lock on oThread
                End If
            End If
        Loop
    End Sub

    Public Overloads Sub Dispose(ByVal disposing As Boolean)
        bStopListener = False  ' Indicate user wants to dispose this ServerWinsock object
        ' WriteSip2Log("Dispose:")
        oListener.Stop() ' Stop the TcpListener
        oListener = Nothing
    End Sub

    Private Sub WriteSip2Log(ByVal strMessage As String, Optional ByVal blnCreate As Boolean = False)
        Try
            Dim fsLog As StreamWriter
            If blnCreate Then
                fsLog = File.CreateText("C:\Sip2Server.log")
            Else
                fsLog = File.AppendText("C:\Sip2Server.log")
            End If
            fsLog.WriteLine(strMessage)
            fsLog.Close()
            fsLog = Nothing
        Catch ex As Exception

        End Try



    End Sub

End Class

