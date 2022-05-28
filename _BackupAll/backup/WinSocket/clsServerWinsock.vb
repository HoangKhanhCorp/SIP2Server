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
    Public Event OnMessage(soc As Socket, ByVal IncomingMessage As String, ByVal strIpClient As String)
    Private m_aryClients As ArrayList = New ArrayList()
    Public Shared allDone As New ManualResetEvent(False)
    Public Function GetSocketServer() As Socket
        Return oSocket
    End Function
    Public Sub New()

    End Sub

    Public Sub New(ByVal PortNo As Integer)
        '================================ Method 1
        '' Instantiate a TcpListener on given PortNo
        'oListener = New TcpListener(GetIP(), PortNo)
        'oListener.Start() ' Start the TcpListener

        '' Create a thread for Sub AcceptConnection
        'Dim ServerThread As Thread
        'ServerThread = New Thread(AddressOf AcceptConnection)
        'ServerThread.Start() ' Run Sub AcceptConnection like a light weight child process
        WriteSip2Log("Connect new service")

        Try



            '=====================================Method 2
            ' Dim clsServer As New clsServerWinsock()
            ' Create the listener socket in this machines IP address
            ' Socket listener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            oSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            oSocket.Bind(New IPEndPoint(GetIP(), PortNo))
            'oSocket.Bind( New IPEndPoint( IPAddress.Loopback, PortNo ) );	// For use with localhost 127.0.0.1
            oSocket.Listen(10)
            '    While True

            '   oSocket.BeginAccept(New AsyncCallback(AddressOf OnConnectRequest), oSocket)

            ' While True
            ' Set the event to nonsignaled state.  
            '   allDone.Reset()

            ' Start an asynchronous socket to listen for connections.  
            '  Console.WriteLine("Waiting for a connection...")
            '  While True

            oSocket.BeginAccept(New AsyncCallback(AddressOf OnConnectRequest), oSocket)
            '  End While

            ' Wait until a connection is made and processed before continuing.  
            '    allDone.WaitOne()
            '  End While

            '  End While
            'Setup a callback to be notified of connection requests


            'Console.WriteLine("Press Enter to exit")
            'Console.ReadLine()
            'Console.WriteLine("OK that does it! Screw you guys I'm going home...")

            ' Clean up before we go home
            '    oSocket.Close()
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
        Catch ex As Exception
            WriteSip2Log("EX:" & ex.Message)
        End Try
    End Sub
    Public Sub OnConnectRequest(ByVal ar As IAsyncResult)
        Dim listener As Socket = CType(ar.AsyncState, Socket)
        NewConnection(listener.EndAccept(ar))
        listener.BeginAccept(New AsyncCallback(AddressOf OnConnectRequest), listener)
    End Sub
    Public Sub NewConnection(ByVal sockClient As Socket)
        Dim client As SocketChatClient = New SocketChatClient(sockClient)
        m_aryClients.Add(client)
        '  Console.WriteLine("Client {0}, joined", client.Sock.RemoteEndPoint)
        Dim now As DateTime = DateTime.Now
        Dim strDateLine As String = "96AZFFFFFEF6" ' "Welcome " & now.ToString("G") & vbLf & vbCr
        Dim byteDateLine As Byte() = System.Text.Encoding.ASCII.GetBytes(strDateLine.ToCharArray())
        client.Sock.Send(byteDateLine, byteDateLine.Length, 0)
        client.SetupRecieveCallback(Me)
    End Sub
    Public Sub OnRecievedData(ByVal ar As IAsyncResult)
        '  Thread.Sleep(100)
        Dim client As SocketChatClient = ar.AsyncState ' CType(ar.AsyncState, SocketChatClient)
        Dim aryRet As Byte() = client.GetRecievedData(ar)

        If aryRet.Length < 1 Then
            '        Console.WriteLine("Client {0}, disconnected", client.Sock.RemoteEndPoint)
            client.Sock.Close()
            m_aryClients.Remove(client)
            Return
        End If

        For Each clientSend As SocketChatClient In m_aryClients

            Try
                'clientSend.Sock.Send(aryRet)
                Dim RecvMessage As String ' UTF8 text string of the buffer (array of bytes)
                Dim strIPClient As String
                RecvMessage = System.Text.Encoding.UTF8.GetString(aryRet)
                strIPClient = CType(client.Sock.RemoteEndPoint, IPEndPoint).Address.ToString
                oSocket = client.Sock
                RaiseEvent OnMessage(Nothing, RecvMessage, strIPClient)
                client.SetupRecieveCallback(Me)
            Catch ex As Exception
                '  Console.WriteLine("Send to client {0} failed", client.Sock.RemoteEndPoint)

                clientSend.Sock.Close()
                m_aryClients.Remove(client)
                Return
            End Try
        Next



        ' Get Ip client

    End Sub

    Private Function GetIP() As IPAddress
        Dim strHostName As String = Dns.GetHostName
        Dim IpHostEntry As IPHostEntry = Dns.GetHostByName(strHostName)
        Dim IpAddress As IPAddress = IpHostEntry.AddressList(1)
        Return IpAddress
    End Function

    Protected Sub Unsed_ProcessRequest()
        Dim Buffer(5000) As Byte ' used to receive incoming message from TcpClient
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
                Try
                    bytes = oSocket.Receive(Buffer, Buffer.Length, 0)
                    SyncLock oThread  ' Lock oThread
                        ' Convert the array of bytes (i.e. the buffer) to UTF8 text string
                        RecvMessage = System.Text.Encoding.UTF8.GetString(Buffer)
                        ' Get Ip client
                        strIPClient = CType(oSocket.RemoteEndPoint, IPEndPoint).Address.ToString
                        '' Raise an event to return the message to the program that owns this ServerWinsock
                        '        RaiseEvent OnMessage(RecvMessage, strIPClient)

                        WriteSip2Log("ServerSock:" & RecvMessage)
                    End SyncLock  ' unlock oThread

                Catch ex As Exception

                End Try


                Exit While
            End If
            ' get out of while loop if TcpClient has disconnected
            If Not oSocket.Connected Then Exit While
        End While
        '    oSocket.Close() ' Close the TcpServer socket
        SyncLock oThread ' Lock oThread
            ActiveThreads -= 1  ' Decrement number of Active Threads
        End SyncLock ' unlock oThread

        WriteSip2Log("Process close:" & oSocket.Connected) 'oSocket.Connected.ToString()
    End Sub

    Private Sub AcceptConnection()
        'This is the main Sub of ServerWinsock
        ' Keep looping until user wants to stop
        WriteSip2Log("Server Accept1")
        Do While Not bStopListener
            Thread.Sleep(3000)  ' Sleep 100 msec. '5000
            If Not oListener Is Nothing AndAlso oListener.Pending() Then  ' received a request for connection from a TCPClient
                If ActiveThreads <= MaxThread Then
                    WriteSip2Log("Server Accept2")
                    ' create a thread to handle this Client Request
                    Dim oThread As Thread
                    oThread = New Thread(AddressOf Unsed_ProcessRequest) 'Unsed_
                    oThread.Start()  ' Run Sub ProcessRequest
                    SyncLock oThread  ' Lock oThread so that ActiveThreads value is not changed while we're adding 1 to it
                        ActiveThreads += 1  ' Increment number of active threads
                    End SyncLock  ' Release the lock on oThread
                End If
            End If
        Loop
    End Sub

    Public Overloads Sub Dispose(ByVal disposing As Boolean)
        bStopListener = True  ' Indicate user wants to dispose this ServerWinsock object
        WriteSip2Log("Dispose:")
        oListener.Stop() ' Stop the TcpListener
        oListener = Nothing
    End Sub

    Private Sub WriteSip2Log(ByVal strMessage As String, Optional ByVal blnCreate As Boolean = False)
        Dim fsLog As StreamWriter
        If blnCreate Then
            fsLog = File.CreateText("C:\Sip2Server.log")
        Else
            fsLog = File.AppendText("C:\Sip2Server.log")
        End If
        fsLog.WriteLine(strMessage)
        fsLog.Close()
        fsLog = Nothing
    End Sub

End Class

Friend Class SocketChatClient
    Private m_sock As Socket
    Private m_byBuff As Byte() = New Byte(50) {}

    Public Sub New(ByVal sock As Socket)
        m_sock = sock
    End Sub

    Public ReadOnly Property Sock As Socket
        Get
            Return m_sock
        End Get
    End Property

    Public Sub SetupRecieveCallback(ByVal app As clsServerWinsock)
        Try
            Dim recieveData As AsyncCallback = New AsyncCallback(AddressOf app.OnRecievedData)
            m_sock.BeginReceive(m_byBuff, 0, m_byBuff.Length, SocketFlags.None, recieveData, Me)

            'Dim RecvMessage As String = System.Text.Encoding.UTF8.GetString(m_byBuff)
            'If RecvMessage = "" Then
            '    RecvMessage ="khong nhan dc"
            'End If
        Catch ex As Exception
            '  Console.WriteLine("Recieve callback setup failed! {0}", ex.Message)
        End Try
    End Sub

    Public Function GetRecievedData(ByVal ar As IAsyncResult) As Byte()
        Dim nBytesRec As Integer = 0

        Try
            nBytesRec = m_sock.EndReceive(ar)
        Catch eSocket As SocketException
        Catch ex As Exception

        End Try

        Dim byReturn As Byte() = New Byte(nBytesRec - 1) {}
        Array.Copy(m_byBuff, byReturn, nBytesRec)
        Return byReturn
    End Function
End Class

