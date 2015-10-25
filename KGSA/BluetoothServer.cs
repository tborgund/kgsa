using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using KGSA.Bluetooth;

namespace KGSA
{
    public class BluetoothServer : IDisposable
    {
        private FormMain main;
        public static string CMD_UPDATE_INVENTORY = "UPDATE_INVENTORY";
        public static string CMD_UPDATE_PRODUCT = "UPDATE_PRODUCT";
        public static string CMD_START_TRANSFER = "START_TRANSFER";
        public static string CMD_GOODBYE = "GOODBYE";
        public static string CMD_OK = "OK";
        public static string RESP_ACCEPTED = "ACCEPTED";
        public static string RESP_REJECTED = "REJECTED";
        public static string RESP_SENDING_FILE = "SENDING_FILE";
        public static string APP_PACKAGE_NAME = "borgund.mscanner";

        private readonly Guid mUUID = new Guid("00001101-5500-1000-8000-00805F9834FB");
        public static string inventoryFilename = "Auto_AppInventory.sqlite";
        public static string dataFilename = "Auto_AppProductData.sqlite";

        private bool serverStarted = false;
        private BluetoothListener blueListener;
        private Thread bluetoothServerThread;

        private string clientName = "Unknown";
        private bool clientAccepted = false;
        private string clientAppName = "Unknown";
        private int clientAppVersion = 0;

        private bool noBlockingMode = false;

        public BluetoothServer(FormMain form)
        {
            this.main = form;
        }

        public void StartServer(bool noBlocking)
        {
            this.noBlockingMode = noBlocking;
            if (hasBluetoothSupport(true))
            {
                if (main.appConfig.blueServerDatabaseUpdated.Date != DateTime.Now.Date)
                    Log.BtServer("App-databasene trenger oppdatering. Sist oppdatert: " + main.appConfig.blueServerDatabaseUpdated.ToShortDateString(), true);
                runServer();
            }
        }

        public void StopServer()
        {
            this.serverStarted = false;
            Dispose();
        }

        public bool IsOnline()
        {
            return this.serverStarted;
        }

        internal bool hasBluetoothSupport(bool noBlocking)
        {
            BluetoothRadio bluetoothRadio = BluetoothRadio.PrimaryRadio;
            if (bluetoothRadio == null)
            {
                if (noBlocking)
                    Log.BtServer("Ingen støttet bluetooth maskinvare er tilkoblet eller slått på.", true);
                else
                    Log.Alert("Ingen støttet bluetooth maskinvare er tilkoblet eller slått på.", "Bluetooth maskinvare mangler", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return false;
            }
            else
            {
                Log.BtServer("Bluetooth enhet funnet: " + bluetoothRadio.Name + " (" + bluetoothRadio.LocalAddress + ") Modus: " + bluetoothRadio.Mode.ToString());
                if (!bluetoothRadio.Mode.ToString().Equals("Discoverable"))
                {
                    if (noBlocking)
                        Log.BtServer("Advarsel: Denne maskinen er ikke synlig for andre Bluetooth enheter."
                            + " Det anbefales at den gjøres synlig for at App'en kan oppdateres automatisk.", true);
                    else
                        Log.Alert("Advarsel: Denne maskinen er ikke synlig for andre Bluetooth enheter."
                            + " Det anbefales at den gjøres synlig for at App'en kan oppdateres automatisk.",
                            "Bluetooth ikke synlig", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }
                return true;
            }
        }

        internal void runServer()
        {
            bluetoothServerThread = new Thread(new ThreadStart(ServerListeningThread));
            bluetoothServerThread.Start();
        }

        internal void ServerListeningThread()
        {
            try
            {
                serverStarted = false; // just making sure server start is in sync
                try
                {
                    blueListener = new BluetoothListener(mUUID);
                    blueListener.Start();
                    serverStarted = true;
                }
                catch (Exception ex)
                {
                    serverStarted = false;
                    Log.Unhandled(ex);
                    if (noBlockingMode)
                    {
                        Log.BtServer("Bluetooth enhet er ikke støttet, er slått av, eller mangler." + ex.Message, true);
                    }
                    else
                    {
                        Log.ErrorDialog(ex, "Problem oppstod ved start av Bluetooth listener.\nBluetooth enhet er ikke støttet, er slått av, eller mangler. Kan også oppstå når der er problem med programvaren/driver til Bluetooth maskinvaren.", "Bluetooth serveren kan ikke starte.");
                    }
                }

                Log.BtServer("Bluetooth server har startet. Aksepterer tilkoblinger fra alle MScanner mobiler fra versjon " + main.appConfig.blueServerMinimumAcceptedVersion);
                while (true)
                {
                    try
                    {
                        if (!serverStarted)
                            break;

                        if (blueListener.Pending())
                        {
                            using (BluetoothClient client = blueListener.AcceptBluetoothClient())
                            {
                                ClientConnected(client);
                            }
                        }
                    }
                    catch (SocketException ex) { Log.BtServer("Lese/skrive feil: " + ex.Message); break; }
                    catch (Exception ex) { Log.Unhandled(ex); break; }
                }

                if (blueListener != null)
                    blueListener.Stop();

                if (blueListener != null)
                    Log.BtServer("Server avslått");

                serverStarted = false;
            }
            catch (Exception ex)
            {
                Log.BtServer("Generell feil: " + ex.Message);
            }
            finally
            {
                if (blueListener != null)
                {
                    blueListener.Stop();
                    blueListener = null;
                }
                serverStarted = false;
            }
        }

        internal void ClientConnected(BluetoothClient client)
        {
            Stream stream = null;
            try
            {
                Log.BtServer("MScanner med navn " + client.RemoteMachineName + " kobler til..");
                this.clientName = client.RemoteMachineName;
                this.clientAccepted = false;
                this.clientAppName = "Unknown";
                this.clientAppVersion = 0;
                stream = client.GetStream();
                stream.ReadTimeout = 10 * 1000;

                if (main.appConfig.blueServerDatabaseUpdated.Date != DateTime.Now.Date)
                {
                    Log.BtServer("App-databasene trenger oppdatering. Sist oppdatert: " + main.appConfig.blueServerDatabaseUpdated.ToShortDateString());
                    if (FormMain.appManagerIsBusy)
                        Log.BtServer("AppManager er opptatt: App-databasene er iferd med å bli oppdatert");
                    else
                    {
                        AppManager app = new AppManager(main);
                        app.UpdateAllAsync();
                        return;
                    }
                }

                while (true)
                {
                    try
                    {
                        if (!client.Connected || !serverStarted)
                            break;

                        if (!clientAccepted) // Client is not initialized, run init routine..
                            this.clientAccepted = serverInitConnectionRoutine(stream);

                        string strServerMsg = waitForMessage(stream);
                        if (strServerMsg.StartsWith(CMD_UPDATE_PRODUCT))
                        {
                            if (!serverUpdateRoutine(stream, strServerMsg, dataFilename))
                                break;
                        }
                        if (strServerMsg.StartsWith(CMD_UPDATE_INVENTORY))
                        {
                            if (!serverUpdateRoutine(stream, strServerMsg, inventoryFilename))
                                break;
                        }
                        if (strServerMsg.StartsWith(CMD_GOODBYE))
                        {
                            Log.BtServer("(" + this.clientName + ") Goodbye!");
                            break;
                        }

                        Thread.Sleep(300); // Used to limit the number of loops per second
                    }
                    catch (Exception ex)
                    {
                        Log.BtServer("Kommunikasjons feil med " + this.clientName + ". Kobler fra..", true);
                        Log.Unhandled(ex);
                        break;
                    }
                }

                if (stream != null)
                {
                    stream.Flush();
                    stream.Close();
                }

                if (client != null)
                    client.Close();

                Log.BtServer("(" + this.clientName + ") Koblet fra.");
            }
            catch (Exception ex)
            {
                Log.BtServer("Generell feil: " + ex.Message);
            }
            finally
            {
                if (stream != null)
                {
                    stream.Flush();
                    stream.Close();
                    stream = null;
                }
                if (client != null)
                {
                    client.Close();
                    client = null;
                }
            }
        }

        internal bool serverInitConnectionRoutine(Stream stream)
        {
            try
            {
                string strClientMsg = waitForMessage(stream);
                string[] lines = strClientMsg.Split(';');
                if (lines == null || !lines[0].Equals(APP_PACKAGE_NAME) || Convert.ToDouble(lines[1]) < main.appConfig.blueServerMinimumAcceptedVersion)
                {
                    Log.BtServer(this.clientName + " har ikke en kompatibel versjon av MScanner. Kobles fra..");
                    sendMessage(stream, RESP_REJECTED);
                    return false;
                }

                this.clientAppName = lines[0];
                this.clientAppVersion = Convert.ToInt32(lines[1]);

                Log.BtServer(this.clientName + " koblet til. Navn: " + this.clientAppName + " App versjon: " + this.clientAppVersion);
                if (!sendMessage(stream, RESP_ACCEPTED + ";" + main.appConfig.blueProductExportDate.ToString("dd.MM.yy") + ";" + main.appConfig.blueInventoryExportDate.ToString("dd.MM.yy")))
                {
                    Log.BtServer("Kommunikasjons feil med " + this.clientName + ". Kobles fra..");
                    return false;
                }

                clientAccepted = true;
                Log.BtServer(this.clientName + " er klar til å sende kommandoer.");

                return true;
            }
            catch (Exception ex)
            {
                Log.BtServer("Feil under initialisering av tilkoblingen til " + this.clientName + ". Kobles fra..");
                Log.Unhandled(ex);
            }
            this.clientAccepted = false;
            return false;
        }

        internal bool serverUpdateRoutine(Stream stream, string strClientMsg, string filenameArg)
        {
            try
            {
                string[] lines = strClientMsg.Split(';');
                if (lines == null || lines.Length != 2 || String.IsNullOrEmpty(lines[1]))
                {
                    Log.BtServer("(" + this.clientName + ") Kommunikasjons problem. Kobler fra..");
                    return false;
                }

                string dateFormat = "dd.MM.yy";
                DateTime date = DateTime.ParseExact(lines[1], dateFormat, FormMain.norway);

                Log.BtServer("(" + this.clientName + ") Er sist oppdatert " + date.ToShortDateString());

                try { Thread.Sleep(1500); } // We may disconnect devices if we reply too fast. Most communication methods has otherwise mandatory delays
                catch (Exception ex) { Log.Unhandled(ex); }

                string filepath = FormMain.settingsPath + @"\" + filenameArg;

                // Load file to memory. Its usually ~2-20 MB, nothing to worry about
                MemoryStream memoryStream = null;
                FileStream fsSource = null;
                try
                {
                    memoryStream = new MemoryStream();
                    fsSource = new FileStream(filepath, FileMode.Open, FileAccess.Read);
                    fsSource.CopyTo(memoryStream);
                    fsSource.Close();
                }
                catch (IOException ex)
                {
                    Log.n("Kan ikke lese database! Feilmelding: " + ex.Message, Color.Red);
                    if (fsSource != null)
                        fsSource.Close();
                    if (memoryStream != null)
                        memoryStream.Close();
                    return false;
                }

                long dataLength = memoryStream.Length;
                memoryStream.Seek(0, SeekOrigin.Begin);

                Log.BtServer("Forbereder sending.. Størrelse: " + BytesToString((long)dataLength));

                if (!sendMessage(stream, RESP_SENDING_FILE + ";" + filenameArg + ";" + dataLength.ToString()))
                {
                    Log.BtServer("(" + this.clientName + ") Kommunikasjons problem. Kobler fra..", true);
                    return false;
                }

                strClientMsg = waitForMessage(stream);
                if (!strClientMsg.Equals(CMD_START_TRANSFER))
                {
                    Log.BtServer("Ukjent kommando sendt fra " + this.clientName + ": '" + strClientMsg + "'", true);
                    return false;
                }

                Log.BtServer("Sender fil til " + this.clientName + "..");
                if (!sendFile(stream, memoryStream))
                {
                    Log.BtServer("(" + this.clientName + ") Filoverføring uventet avbrutt.", true);
                    return false;
                }

                Log.BtServer("(" + this.clientName + ") Filoverføring ferdig. Venter på verifikasjon..");
                strClientMsg = waitForMessage(stream);
                if (!strClientMsg.Equals(CMD_OK))
                {
                    Log.BtServer("Ukjent kommando sendt fra " + this.clientName + ": '" + strClientMsg + "'", true);
                    return false;
                }
                Log.BtServer("(" + this.clientName + ") Mottat fil uten feil");

                return true;
            }
            catch (Exception ex)
            {
                Log.BtServer("Feil under tilkobling av " + this.clientName + ". Kobles fra..");
                Log.d("Unntak ved serverUpdateRoutine(): " + ex.Message);
            }
            this.clientAccepted = false;
            return false;
        }

        internal String BytesToString(long byteCount)
        {
            string[] suf = { " B", " KB", " MB", " GB", " TB", " PB", " EB" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }

        internal bool sendMessage(Stream stream, string message)
        {
            try
            {
                try { Thread.Sleep(2000); }
                catch (Exception ex) { Log.Unhandled(ex); }

                byte[] send = Encoding.UTF8.GetBytes(message);
                stream.Write(send, 0, send.Length);
                Log.d("btserver: Melding sendt til '" + message + "'");
                return true;
            }
            catch (Exception)
            {
                Log.d("btserver: Kunne ikke sende melding til klient: " + message);
            }
            return false;
        }

        internal bool sendFile(Stream outStream, MemoryStream inStream)
        {
            byte[] buffer = new byte[1024];
            int bytes = 0;

            try
            {
                try { Thread.Sleep(2000); }
                catch (Exception ex) { Log.Unhandled(ex); }

                while ((bytes = inStream.Read(buffer, 0, buffer.Length)) != 0)
                {
                    if (!serverStarted)
                        break;

                    outStream.Write(buffer, 0, bytes);
                }

                return true;
            }
            catch (TimeoutException) { Log.d("btserver: klient tok for lang tid til å svare."); }
            catch (Exception ex)
            {
                Log.d("btserver: Unntak ved lesing/skriving til bluetooth klient. Melding: " + ex.Message);
                if (inStream != null)
                    inStream.Close();
            }
            return false;
        }

        internal string waitForMessage(Stream stream)
        {
            byte[] buffer = new byte[1024];
            int bytes = 0;

            while (true)
            {
                try
                {
                    if (!serverStarted)
                        break;

                    bytes = stream.Read(buffer, 0, buffer.Length);
                    if (bytes > 0)
                    {
                        String str = Encoding.UTF8.GetString(buffer).TrimEnd('\0');
                        if (str != null && str.Length > 0)
                        {
                            Log.d("btserver: Melding fra klient: '" + str + "'");
                            return str;
                        }
                        else
                            Log.d("btserver: Melding fra klient var tom (ikke Null)!");
                    }
                }
                catch (TimeoutException) { Log.d("btserver: klient tok for lang tid til å svare."); break; }
                catch (Exception ex)
                {
                    Log.d("btserver: Unntak ved lesing/skriving til bluetooth klient. Melding: " + ex.Message);
                    break;
                }
            }
            return "";
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    if (blueListener != null)
                    {
                        blueListener.Stop();
                        blueListener = null;
                    }
                }
                catch (Exception) { }
            }
        }
    }
}
