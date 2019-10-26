using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace WindowsService1
{
    public partial class BackUpService : ServiceBase
    {
        private EventLog eventLog;
        private bool stopping;
        private ManualResetEvent stoppedEvent;

        public BackUpService()
        {
            InitializeComponent();
            eventLog = new EventLog("backUp.log", "remotehost");
        }
        protected override void OnStart(string[] args)
        {
            //Тушим порт на котором у нас NAS, вдруг он включен
            Utility.DoCommand("shutdown", "interface GE 1/10");
            // Пишем в логи, шлём письма
            this.eventLog.WriteEntry("BackupService is OnStart.");
            Utility.SendMail("Служба BCP запущенна", "Служба BCP успешно запущенна");
            Utility.WriteLog("Служба BCP успешно запущенна" + DateTime.Now + "\n");
            // Запускаем потоки на исполнение, один из которых делает бэкапы второй следит за статусом службы
            ThreadPool.QueueUserWorkItem(new WaitCallback(ServiceBackupThread));
            // ThreadPool.QueueUserWorkItem(new WaitCallback(ServiceWorkerThread));
        }
        protected override void OnStop()
        {
        }
        private void ServiceBackupThread(object state)
        {
            while (!this.stopping)
            {
                int Errors = 0;
                string files = "\n";

                while (true)
                {
                    // Исполнение задания в 6 утра по времени локального хоста
                    if (DateTime.Now.Hour == 6 && DateTime.Now.Minute == 0)
                    {
                        string Year = DateTime.Now.Year.ToString();
                        string Month = DateTime.Now.Month.ToString();
                        string Day = DateTime.Now.Day.ToString();
                        string BCPath = @"\\Dropbox\";
                        string BCPath2 = @"\Dropbox\";

                        // Создаём папки в сетевом хранилище для бэкапов
                        try
                        {
                            Directory.CreateDirectory(BCPath2 + Year + Month + Day + @"\");
                        }
                        catch (Exception ex)
                        {
                            Utility.WriteLog("ОШИБКА: " + ex.Message + "\n");
                        }

                        Utility.WriteLog("\n" + "НАЧАЛО РЕЗЕВНОГО КОПИРОВАНИЯ " + DateTime.Now + "\n");
                        try
                        {
                            Utility.DoCommand("no shutdown", "interface FE 1/10");
                            Utility.WriteLog("ШАГ 1: Порт успешно открыт" + "\n");
                        }
                        catch (Exception ex)
                        {
                            Utility.WriteLog("ОШИБКА НА ШАГЕ 1: " + ex.Message + "\n");
                            Utility.SendMail("ОШИБКА НА ШАГЕ 1", "ОШИБКА НА ШАГЕ 1: " + ex.Message + "\n");
                            Errors++;
                        }

                        try
                        {
                            Utility.WriteLog("ШАГ 2: Начало копированя файлов" + "\n");

                            List<FileInfo> FileList = Utility.GetFileList(@"D:" + BCPath2);
                            System.Collections.IList list = FileList;
                            for (int i = 0; i < list.Count; i++)
                            {
                                string file = list[i].ToString();
                                files = files + file + "\n";
                                string SoucePath = @"D:"+ BCPath2 + file;
                                string DestPath = BCPath + Year + @"\" + Month + @"\" + file;
                                Utility.WriteLog("Copy file from " + SoucePath + " to " + DestPath + "\n");
                                Utility.MoveTime(SoucePath, DestPath);
                            }
                            Utility.WriteLog("ШАГ 2: Файлы успешно скопированы" + "\n" + files);
                        }
                        catch (Exception ex)
                        {
                            Utility.WriteLog("ОШИБКА НА ШАГЕ 2: " + ex.Message + "\n" + files);
                            Utility.SendMail("ОШИБКА НА ШАГЕ 2", "ОШИБКА НА ШАГЕ 2: " + ex.Message + "\n");
                            Errors++;
                        }

                        try
                        {
                            Utility.DoCommand("shutdown", "interface GE 1/10");
                            Utility.WriteLog("ШАГ 3: Порт успешно закрыт" + "\n");
                        }
                        catch (Exception ex)
                        {
                            Utility.WriteLog("ОШИБКА НА ШАГЕ 3: " + ex.Message + "\n");
                            Utility.SendMail("ОШИБКА НА ШАГЕ 3", "ОШИБКА НА ШАГЕ 3: " + ex.Message + "\n");
                            Errors++;
                        }

                        if (Errors == 0)
                        {
                            Utility.WriteLog("РЕЗЕВНОЕ КОПИРОВАНИЕ ВЫПОЛНЕНО БЕЗ ОШИБОК " + DateTime.Now + "\n");
                            Utility.SendMail("РК ВЫПОЛНЕНО", "РЕЗЕРВНОЕ КОПИРОВАНИЕ ФАЙЛОВ " + files + " ВЫПОЛНЕНО БЕЗ ОШИБОК \n");
                        }
                        else
                        {
                            Utility.WriteLog("РЕЗЕВНОЕ КОПИРОВАНИЕ ВЫПОЛНЕНО С " + Errors + " ОШИБКАМИ " + DateTime.Now + "\n");
                            Utility.SendMail("РК НЕ ВЫПОЛНЕНО", "РЕЗЕРВНОЕ КОПИРОВАНИЕ ВЫПОЛНЕНО С " + Errors + " ОШИБКАМИ \n" + files);
                            Errors = 0;
                        }
                        files = "\n";
                    }
                    Thread.Sleep(60000);
                }
            }

            this.stoppedEvent.Set();
        }
    /*
        if (Get-Service BackUpItNow -ErrorAction SilentlyContinue) {
            $service = Get-WmiObject -Class Win32_Service -Filter "name='BackUpItNow'"
            $service.StopService()
            Start-Sleep -s 1
            $service.delete()
        }

        $workdir = Split-Path -parent $MyInvocation.MyCommand.Path

        New-Service -name BackUpItNow `
        -displayName BackUpItNow `
        -binaryPathName "`"a:\BackUpItNow.exe`""
     */
    }
}

public static class Utility
{
    public static List<FileInfo> GetFileList(string directoryPath)
    {
        DirectoryInfo dir = new DirectoryInfo(directoryPath);
        FileInfo[] theFiles = dir.GetFiles("*", SearchOption.TopDirectoryOnly);
        return theFiles.Where(fl => fl.CreationTime.Date == DateTime.Today).ToList();
    }
    public static void SendMail(string Body, string Text)
    {
        try
        {
            MailMessage mail = new MailMessage("backups@company.ru", "youtname@company.ru");
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Host = "mailserv.company.ru";
            mail.Subject = Body;
            mail.Body = Text;
            client.Send(mail);
        }
        catch (Exception ex)
        {
            WriteLog("ОШИБКА ОТПРАВКИ ПИСЬМА: " + "\n" + ex + "\n");
        }
    }
    public static void WriteLog(string Text)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(Text);
        File.AppendAllText("C:\\servicepath\\LogBackup.txt", sb.ToString());
        sb.Clear();
    }
    public static void DoCommand(string Command, string Port)
    {
        TelnetConnection tc = new TelnetConnection("192.168.0.200", true);

        string s = tc.Login("cisco", "SuperSecretPASS", 1000);
        Console.Write(s);

        string prompt = s.TrimEnd();

        prompt = "";

        System.Threading.Thread.Sleep(1000);

        prompt = "conf";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = Port;
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = Command;
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = "end";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = "write";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = "Y";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(1000);

        prompt = "exit";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(10000);

        prompt = "";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        System.Threading.Thread.Sleep(10000);

        prompt = "exit";
        tc.WriteLine(prompt);
        Console.Write(tc.Read());

        Console.WriteLine("***DISCONNECTED" + "\n");
        WriteLog("Command '" + Command + "' was executed for port: " + Port + "\n");
    }
    public static void MoveTime(string source, string destination)
    {
        DateTime start_time = DateTime.Now;
        FMove(source, destination);
        long size = new FileInfo(destination).Length;
        int milliseconds = 1 + (int)((DateTime.Now - start_time).TotalMilliseconds);
        long tsize = size * 3600000 / milliseconds;
        tsize = tsize / (int)Math.Pow(2, 30);
        Console.WriteLine("Speed of copy " + tsize + "GB/hour" + "\n");
        WriteLog("Speed of copy " + tsize + "GB/hour" + "\n");
    }
    static void FMove(string source, string destination)
    {
        int array_length = (int)Math.Pow(2, 19);
        byte[] dataArray = new byte[array_length];
        using (FileStream fsread = new FileStream
        (source, FileMode.Open, FileAccess.Read, FileShare.None, array_length))
        {
            using (BinaryReader bwread = new BinaryReader(fsread))
            {
                using (FileStream fswrite = new FileStream
                (destination, FileMode.Create, FileAccess.Write, FileShare.None, array_length))
                {
                    using (BinaryWriter bwwrite = new BinaryWriter(fswrite))
                    {
                        for (; ; )
                        {
                            int read = bwread.Read(dataArray, 0, array_length);
                            if (0 == read)
                                break;
                            bwwrite.Write(dataArray, 0, read);
                        }
                    }
                }
            }
        }
    }
    public static string GetMonth(string Month)
    {
        if (Month == "1")
            Month = "01 JAN";
        if (Month == "2")
            Month = "02 FEB";
        if (Month == "3")
            Month = "03 MRT";
        if (Month == "4")
            Month = "04 APR";
        if (Month == "5")
            Month = "05 MAY";
        if (Month == "6")
            Month = "06 JUNE";
        if (Month == "7")
            Month = "07 JULE";
        if (Month == "8")
            Month = "08 AUG";
        if (Month == "9")
            Month = "09 SEP";
        if (Month == "10")
            Month = "10 OKT";
        if (Month == "11")
            Month = "11 NOV";
        if (Month == "12")
            Month = "12 DEC";
        return (Month);
    }
}

public class TelnetConnection : IDisposable
{
    TcpClient m_Client;
    NetworkStream m_Stream;
    bool m_IsOpen = false;
    string m_Hostname;
    int m_ReadTimeout = 1000;
    public delegate void ConnectionDelegate();
    public event ConnectionDelegate Opened;
    public event ConnectionDelegate Closed;
    public bool IsOpen { get { return m_IsOpen; } }
    public TelnetConnection() { }
    public TelnetConnection(bool open) : this("localhost", true) { }
    public TelnetConnection(string host, bool open)
    {
        if (open)
            Open(host);
    }
    void CheckOpen()
    {
        if (!IsOpen)
            throw new Exception("Connection not open.");
    }
    public string Hostname
    {
        get { return m_Hostname; }
    }
    public int ReadTimeout
    {
        set { m_ReadTimeout = value; if (IsOpen) m_Stream.ReadTimeout = value; }
        get { return m_ReadTimeout; }
    }
    public void Write(string str)
    {
        CheckOpen();
        byte[] bytes = System.Text.ASCIIEncoding.ASCII.GetBytes(str);
        m_Stream.Write(bytes, 0, bytes.Length);
        m_Stream.Flush();
    }
    public void WriteLine(string str)
    {
        CheckOpen();
        byte[] bytes = System.Text.ASCIIEncoding.ASCII.GetBytes(str);
        m_Stream.Write(bytes, 0, bytes.Length);
        WriteTerminator();
    }
    void WriteTerminator()
    {
        byte[] bytes = System.Text.ASCIIEncoding.ASCII.GetBytes("\r\n\0");
        m_Stream.Write(bytes, 0, bytes.Length);
        m_Stream.Flush();
    }
    public string Read()
    {
        CheckOpen();
        return System.Text.ASCIIEncoding.ASCII.GetString(ReadBytes());
    }
    public string Read(String login, String password, int timeout)
    {
        CheckOpen();
        Open($"connect {login} {password} /t {timeout}");
        return this.Read();
    }
    public byte[] ReadBytes()
    {
        int i = m_Stream.ReadByte();
        byte b = (byte)i;
        int bytesToRead = 0;
        var bytes = new List<byte>();
        if ((char)b == '#')
        {
            bytesToRead = ReadLengthHeader();
            if (bytesToRead > 0)
            {
                i = m_Stream.ReadByte();
                if ((char)i != '\n')
                    bytes.Add((byte)i);
            }
        }
        if (bytesToRead == 0)
        {
            while (i != -1 && b != (byte)'\n')
            {
                bytes.Add(b);
                i = m_Stream.ReadByte();
                b = (byte)i;
            }
        }
        else
        {
            int bytesRead = 0;
            while (bytesRead < bytesToRead && i != -1)
            {
                i = m_Stream.ReadByte();
                if (i != -1)
                {
                    bytesRead++;
                    // record all bytes except \n if it is the last char.
                    if (bytesRead < bytesToRead || (char)i != '\n')
                        bytes.Add((byte)i);
                }
            }
        }
        return bytes.ToArray();
    }
    int ReadLengthHeader()
    {
        int numDigits = Convert.ToInt32(new string(new char[] { (char)m_Stream.ReadByte() }));
        string bytes = "";
        for (int i = 0; i < numDigits; ++i)
            bytes = bytes + (char)m_Stream.ReadByte();

        return Convert.ToInt32(bytes);
    }
    public void Open(string hostname)
    {
        if (IsOpen)
            Close();
        m_Hostname = hostname;
        m_Client = new TcpClient(hostname, 5025);
        m_Stream = m_Client.GetStream();
        m_Stream.ReadTimeout = ReadTimeout;
        m_IsOpen = true;
        if (Opened != null)
            Opened();
    }
    public void Close()
    {
        if (!m_IsOpen)
            return;
        m_Stream.Close();
        m_Client.Close();
        m_IsOpen = false;
        if (Closed != null)
            Closed();
    }
    public void Dispose()
    {
        Close();
    }
    public string Login(string v1, string v2, int v3)
    {
        return this.Read();
    }
}
