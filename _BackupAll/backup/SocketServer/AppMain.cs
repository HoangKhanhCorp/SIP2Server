using System;
using System.Threading;								// Sleeping
using System.Net;									// Used to local machine info
using System.Net.Sockets;							// Socket namespace
using System.Collections;                           // Access to the Array list
using ProcessData;


	/// <summary>
	/// Main class from which all objects are created
	/// </summary>
	public class AppMain
	{
		// Attributes
		private ArrayList	m_aryClients = new ArrayList(); 
        // List of Client Connections
                                                            /// <summary>
                                                            /// Application starts here. Create an instance of this class and use it
                                                            /// as the main object.
                                                            /// </summary>
                                                            /// <param name="args"></param>
                                                            /// OnMessage
                                                            /// 
        public delegate void OnMessageHandler(Socket clientsc, byte[] msg,string  strIPClient);
        public event OnMessageHandler OnMessage;
        static void Main(string[] args)
		{
			AppMain app = new AppMain();
			// Welcome and Start listening
			Console.WriteLine( "*** Chat Server Started {0} *** ", DateTime.Now.ToString( "G" ) );


			

			//
			// Method 2 
			//
			const int nPortListen = 6001;
			// Determine the IPAddress of this machine
			IPAddress [] aryLocalAddr = null;
			String strHostName = "";
			try
			{
				// NOTE: DNS lookups are nice and all but quite time consuming.
				strHostName = Dns.GetHostName();
				IPHostEntry ipEntry = Dns.GetHostByName( strHostName );
				aryLocalAddr = ipEntry.AddressList;
			}
			catch( Exception ex )
			{
				Console.WriteLine ("Error trying to get local address {0} ", ex.Message );
			}
	
			// Verify we got an IP address. Tell the user if we did
			if( aryLocalAddr == null || aryLocalAddr.Length < 1 )
			{
				Console.WriteLine( "Unable to get local address" );
				return;
			}
			Console.WriteLine( "Listening on : [{0}] {1}:{2}", strHostName, aryLocalAddr[1], nPortListen );

			// Create the listener socket in this machines IP address
			Socket listener = new Socket( AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp );
			listener.Bind( new IPEndPoint( aryLocalAddr[1], 6001 ) );
			//listener.Bind( new IPEndPoint( IPAddress.Loopback, 399 ) );	// For use with localhost 127.0.0.1
			listener.Listen( 10 );

			// Setup a callback to be notified of connection requests
			listener.BeginAccept( new AsyncCallback( app.OnConnectRequest ), listener );

			Console.WriteLine ("Press Enter to exit" );
			Console.ReadLine();
			Console.WriteLine ("OK that does it! Screw you guys I'm going home..." );

			// Clean up before we go home
			listener.Close();
			GC.Collect();
			GC.WaitForPendingFinalizers();		
		}
       
            private string Checksum(string strIn)
            {
            //int intDecChecksum, i;
            //string strHexChecksum;
            //intDecChecksum = 0;
            //for (i = 1; i <= strIn.Length; i++)
            //    intDecChecksum = intDecChecksum + Strings.Asc(Strings.Mid(strIn, i, 1));
            //strHexChecksum = Conversion.Hex((!intDecChecksum) + 1);
            //Checksum = strHexChecksum.PadLeft(4, "0");
            return strIn;
            }

            private bool VerifyChecksum(string strIn)
            {
            //string strHexChecksum, strIncludedChecksum;
            //strHexChecksum = Checksum(Strings.Left(strIn, Strings.Len(strIn) - 4));
            //strIncludedChecksum = Strings.Right(strIn, 4);
            //if (strIncludedChecksum == strHexChecksum)
            //    VerifyChecksum = true;
            //else
            //    VerifyChecksum = false;
            return true;
            }
        clsSip2Patron[] objSip2Patron = new clsSip2Patron[10 + 1];
        private string strChecksumEnabled ="yes";
        private string ProcessRequest(string req, int idx)
        {
           
                string strCommandId = string.Empty;
            string strControl = string.Empty;
            string strChecksum = string.Empty;
            string strSequence = string.Empty;
            string strOutput = string.Empty;

            string strReturn = "";

            strChecksum = strChecksumEnabled;
            if (strChecksum.ToLower() == "yes")
            {
                strControl = req.Substring(req.Length-9, 9);
                if (strControl.IndexOf("AY") == 1 & strControl.IndexOf("AZ") == 4)
                {
                    if (!VerifyChecksum(req))
                    {
                        // ProcessRequest = "96AZ" & Checksum("96AZ")
                        strReturn = "98YYYYYN3000052020022407001156032.00AO|AM|BXYYYYYYYYYYYYYYYY|AN|AF|AG";
                        return strReturn;
                    }
                    else
                    {
                        strSequence = req.Substring(req.Length - 6, 1); 
                        req = req.Substring(0,req.Length-9);
                    }
                }
                else
                    strChecksum = "no";
            }
            if (req == "")
            {
                return "";
               
            }
            strCommandId = req.Substring(0, 2);// Strings.Left(req, 2);
            req = req.Substring(req.Length-2,2);// Strings.Right(req, Strings.Len(req) - 2);
            //objSip2Patron = new clsSip2Patron[idx+1];
            //objSip2Patron[idx].Connectionstring = "SERVER=.\\SQL2017;UID=sa;PWD=123;DATABASE=LibolTDM";
            //objSip2Patron[idx].DBServer = "SQLSERVER";
            //objSip2Patron[idx].Initialize();
            if (strCommandId != "93" )//& not objSip2Patron[idx] is null & objSip2Patron[idx].UserID == 0)
                strOutput = "940";
            else
                switch (strCommandId)
                {
                    case "23":
                        {
                            strOutput = objSip2Patron[idx].PatronStatusRequest(req, true);
                            break;
                        }

                    case "11":
                        {
                            strOutput = objSip2Patron[idx].CheckOut(req);
                            break;
                        }

                    case "09":
                        {
                            strOutput = objSip2Patron[idx].CheckIn(req);
                            break;
                        }

                    case "01":
                        {
                            strOutput = objSip2Patron[idx].BlockPatron(req);
                            break;
                        }

                    case "99":
                        {
                            strOutput = objSip2Patron[idx].SCStatus(req);
                            break;
                        }

                    case "97":
                        {
                            strOutput = "arrLastResponse";// arrLastResponse(idx);
                            break;
                        }

                    case "93":
                        {
                            strOutput = objSip2Patron[idx].Login(req);
                            break;
                        }

                    case "63":
                        {
                            strOutput = objSip2Patron[idx].PatronInformation(req);
                            break;
                        }

                    case "35":
                        {
                            strOutput = objSip2Patron[idx].EndPatronSession(req);
                            break;
                        }

                    case "37":
                        {
                            strOutput = objSip2Patron[idx].FeePaid(req);
                            break;
                        }

                    case "17":
                        {
                            strOutput = objSip2Patron[idx].ItemInformation(req);
                            break;
                        }

                    case "19":
                        {
                            strOutput = objSip2Patron[idx].ItemStatusUpdate(req);
                            break;
                        }

                    case "25":
                        {
                            strOutput = objSip2Patron[idx].PatronEnable(req);
                            break;
                        }

                    case "15":
                        {
                            strOutput = objSip2Patron[idx].Hold(req);
                            break;
                        }

                    case "29":
                        {
                            strOutput = objSip2Patron[idx].Renew(req);
                            break;
                        }

                    case "65":
                        {
                            strOutput = objSip2Patron[idx].RenewAll(req);
                            break;
                        }
                }
            if (strChecksum == "yes")
            {
                strOutput = strOutput + "AY" + strSequence + "AZ";
                strOutput = strOutput + Checksum(strOutput);
            }
            return strOutput;
        }


        /// <summary>
        /// Callback used when a client requests a connection. 
        /// Accpet the connection, adding it to our list and setup to 
        /// accept more connections.
        /// </summary>
        /// <param name="ar"></param>
        public void OnConnectRequest( IAsyncResult ar )
		{
			Socket listener = (Socket)ar.AsyncState;
			NewConnection( listener.EndAccept( ar ) );
			listener.BeginAccept( new AsyncCallback( OnConnectRequest ), listener );
		}

		/// <summary>
		/// Add the given connection to our list of clients
		/// Note we have a new friend
		/// Send a welcome to the new client
		/// Setup a callback to recieve data
		/// </summary>
		/// <param name="sockClient">Connection to keep</param>
		//public void NewConnection( TcpListener listener )
		public void NewConnection( Socket sockClient )
		{
			// Program blocks on Accept() until a client connects.
			//SocketChatClient client = new SocketChatClient( listener.AcceptSocket() );
			SocketChatClient client = new SocketChatClient( sockClient );
			m_aryClients.Add( client );
			Console.WriteLine( "Client {0}, joined", client.Sock.RemoteEndPoint );
 
			// Get current date and time.
			DateTime now = DateTime.Now;
        String strDateLine = "98YYYYYY50000320200304    0943462.00AO|BXYYYYYYYYYYYYYYYY|AY9AZEF2C";//Welcome " + now.ToString("G") + "\n\r";

			// Convert to byte array and send.
			Byte[] byteDateLine = System.Text.Encoding.ASCII.GetBytes( strDateLine.ToCharArray() );
			client.Sock.Send( byteDateLine, byteDateLine.Length, 0 );

			client.SetupRecieveCallback( this );
		}

		/// <summary>
		/// Get the new data and send it out to all other connections. 
		/// Note: If not data was recieved the connection has probably 
		/// died.
		/// </summary>
		/// <param name="ar"></param>
		public void OnRecievedData( IAsyncResult ar )
		{
			SocketChatClient client = (SocketChatClient)ar.AsyncState;
			byte [] aryRet = client.GetRecievedData( ar );

			// If no data was recieved then the connection is probably dead
			if( aryRet.Length < 1 )
			{
				Console.WriteLine( "Client {0}, disconnected", client.Sock.RemoteEndPoint );
				client.Sock.Close();
				m_aryClients.Remove( client );      				
				return;
			}
            int i = 0;
			// Send the recieved data to all clients (including sender for echo)
			foreach( SocketChatClient clientSend in m_aryClients )
			{
				try
				{
                  string  RecvMessage = System.Text.Encoding.UTF8.GetString(aryRet);
                string strIPClient = ((IPEndPoint)clientSend.Sock.RemoteEndPoint).Address.ToString();
                //  this.OnMessage += handler(clientSend.Sock,aryRet,strIPClient);
                RecvMessage = ProcessRequest(RecvMessage, i);
                byte[] RBuffer = System.Text.Encoding.ASCII.GetBytes(RecvMessage.ToCharArray());



                clientSend.Sock.Send(RBuffer);
            }
				catch(Exception e)
				{
					// If the send fails the close the connection
					Console.WriteLine( "Send to client {0} failed" + e.Message, client.Sock.RemoteEndPoint );
					clientSend.Sock.Close();
					m_aryClients.Remove( client );
					return;
				}
                i += 1;
			}
			client.SetupRecieveCallback( this );
		}

    private OnMessageHandler handler(Socket sock, byte[] aryRet, string strIPClient)
    {
        sock.Send(aryRet);
        return null;
     //   throw new NotImplementedException();
    }

    //private OnMessageHandler onsend(Socket sock, byte[] aryRet, string strIPClient)
    //{
    //    sock.Send(aryRet)
    //}
}

	/// <summary>
	/// Class holding information and buffers for the Client socket connection
	/// </summary>
	internal class SocketChatClient
	{
		private Socket m_sock;						// Connection to the client
		private byte[] m_byBuff = new byte[50];		// Receive data buffer
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="sock">client socket conneciton this object represents</param>
		public SocketChatClient( Socket sock )
		{
			m_sock = sock;
		}

		// Readonly access
		public Socket Sock
		{
			get{ return m_sock; }
		}

		/// <summary>
		/// Setup the callback for recieved data and loss of conneciton
		/// </summary>
		/// <param name="app"></param>
		public void SetupRecieveCallback( AppMain app )
		{
			try
			{
				AsyncCallback recieveData = new AsyncCallback(app.OnRecievedData);
				m_sock.BeginReceive( m_byBuff, 0, m_byBuff.Length, SocketFlags.None, recieveData, this );
			}
			catch( Exception ex )
			{
				Console.WriteLine( "Recieve callback setup failed! {0}", ex.Message );
			}
		}

		/// <summary>
		/// Data has been recieved so we shall put it in an array and
		/// return it.
		/// </summary>
		/// <param name="ar"></param>
		/// <returns>Array of bytes containing the received data</returns>
		public byte [] GetRecievedData( IAsyncResult ar )
		{
            int nBytesRec = 0;
			try
			{
				nBytesRec = m_sock.EndReceive( ar );
			}
			catch{}
			byte [] byReturn = new byte[nBytesRec];
			Array.Copy( m_byBuff, byReturn, nBytesRec );
			
			/*
			// Check for any remaining data and display it
			// This will improve performance for large packets 
			// but adds nothing to readability and is not essential
			int nToBeRead = m_sock.Available;
			if( nToBeRead > 0 )
			{
				byte [] byData = new byte[nToBeRead];
				m_sock.Receive( byData );
				// Append byData to byReturn here
			}
			*/
			return byReturn;
		}
	}

