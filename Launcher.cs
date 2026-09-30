// Lanzador con licencia - Toto Tools. Compatible con el compilador de Windows (C# 5).
using System;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace TotoLic
{
    public static class Gate
    {
        static volatile bool _ok;
        static long _expTicks;
        static string _dir = "";
        static string _cli = "";
        const string RegPath = @"SOFTWARE\TotoTools";

        public static void Init(string dir) { _dir = dir; }
        public static bool IsOpen() { return _ok && DateTime.Now.Ticks < Interlocked.Read(ref _expTicks); }
        public static string VenceTexto()
        {
            long t = Interlocked.Read(ref _expTicks);
            return t > 0 ? new DateTime(t).ToString("dd/MM/yyyy") : "-";
        }
        public static int Dias()
        {
            long t = Interlocked.Read(ref _expTicks);
            if (t <= 0) return 0;
            return (int)Math.Ceiling((new DateTime(t) - DateTime.Now).TotalDays);
        }
        static void Cerrar() { _ok = false; Interlocked.Exchange(ref _expTicks, 0); }

        // ---------- codigo de esta PC ----------
        static string MachineRaw()
        {
            string g = "";
            try
            {
                using (RegistryKey k = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    if (k != null) g = Convert.ToString(k.GetValue("MachineGuid"));
                }
            }
            catch { }
            if (string.IsNullOrEmpty(g)) g = Environment.MachineName;
            byte[] h;
            using (SHA256 sha = SHA256.Create()) h = sha.ComputeHash(Encoding.UTF8.GetBytes("TT1|" + g));
            const string abc = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
            StringBuilder sb = new StringBuilder();
            int buf = 0, bits = 0;
            for (int i = 0; i < 10; i++)
            {
                buf = (buf << 8) | h[i]; bits += 8;
                while (bits >= 5) { sb.Append(abc[(buf >> (bits - 5)) & 31]); bits -= 5; }
            }
            return sb.ToString();
        }
        public static string MachineCode()
        {
            string c = MachineRaw();
            return c.Substring(0, 4) + "-" + c.Substring(4, 4) + "-" + c.Substring(8, 4) + "-" + c.Substring(12, 4);
        }

        // ---------- verificacion de la licencia (solo clave PUBLICA en el programa) ----------
        static byte[] UnB64(string s)
        {
            s = s.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4) { case 2: s += "=="; break; case 3: s += "="; break; }
            return Convert.FromBase64String(s);
        }
        static bool Firma(byte[] data, byte[] sig)
        {
            try
            {
                byte[] blob = new byte[72];
                blob[0] = 0x45; blob[1] = 0x43; blob[2] = 0x53; blob[3] = 0x31; blob[4] = 32; // ECS1 + 32
                Buffer.BlockCopy(Secrets.Pub, 1, blob, 8, 64);
                using (CngKey key = CngKey.Import(blob, CngKeyBlobFormat.EccPublicBlob))
                using (ECDsaCng ec = new ECDsaCng(key))
                {
                    ec.HashAlgorithm = CngAlgorithm.Sha256;
                    return ec.VerifyData(data, sig);
                }
            }
            catch { return false; }
        }
        static bool Leer(string token, out DateTime exp, out string cli, out string err)
        {
            exp = DateTime.MinValue; cli = ""; err = "";
            try
            {
                token = (token ?? "").Trim().Replace("\r", "").Replace("\n", "").Replace(" ", "");
                string[] p = token.Split('.');
                if (p.Length != 2) { err = "Formato de licencia no valido."; return false; }
                byte[] data = UnB64(p[0]);
                byte[] sig = UnB64(p[1]);
                if (!Firma(data, sig)) { err = "Licencia no valida (firma incorrecta)."; return false; }
                string[] f = Encoding.UTF8.GetString(data).Split('|');
                if (f.Length < 5 || f[0] != "TT1") { err = "Licencia no valida."; return false; }
                if (f[1] != MachineRaw()) { err = "Esta licencia es de otra PC."; return false; }
                exp = DateTime.ParseExact(f[3], "yyyyMMdd", CultureInfo.InvariantCulture);
                cli = f[4];
                return true;
            }
            catch { err = "Licencia ilegible."; return false; }
        }

        // ---------- proteccion contra atrasar el reloj ----------
        static string StampFile()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "TotoTools", "sys.dat");
        }
        static string Mac(string s)
        {
            using (HMACSHA256 h = new HMACSHA256(Secrets.Hm()))
                return Convert.ToBase64String(h.ComputeHash(Encoding.UTF8.GetBytes(s + "|" + MachineRaw())));
        }
        static bool LeerStamp(string raw, out long utc, out string maxExp)
        {
            utc = 0; maxExp = "";
            try
            {
                if (string.IsNullOrEmpty(raw)) return false;
                string[] p = raw.Trim().Split(';');
                if (p.Length != 3 || Mac(p[0] + ";" + p[1]) != p[2]) return false;
                utc = long.Parse(p[0]); maxExp = p[1];
                return true;
            }
            catch { return false; }
        }
        static void StampGet(out long utc, out string maxExp)
        {
            utc = 0; maxExp = "";
            string[] raws = new string[2];
            try { using (RegistryKey k = Registry.LocalMachine.OpenSubKey(RegPath)) { if (k != null) raws[0] = k.GetValue("ts") as string; } } catch { }
            try { if (File.Exists(StampFile())) raws[1] = File.ReadAllText(StampFile()); } catch { }
            foreach (string r in raws)
            {
                long u; string m;
                if (LeerStamp(r, out u, out m))
                {
                    if (string.CompareOrdinal(m, maxExp) > 0 || (m == maxExp && u > utc)) { maxExp = m; utc = u; }
                }
            }
        }
        static void StampSet(long utc, string maxExp)
        {
            string body = utc.ToString() + ";" + maxExp;
            string raw = body + ";" + Mac(body);
            try { using (RegistryKey k = Registry.LocalMachine.CreateSubKey(RegPath)) k.SetValue("ts", raw); } catch { }
            try { string f = StampFile(); Directory.CreateDirectory(Path.GetDirectoryName(f)); File.WriteAllText(f, raw); } catch { }
        }

        // ---------- estado de la licencia ----------
        public static bool Check(out string msg)
        {
            msg = "";
            string token = "";
            try { string f = Path.Combine(_dir, "licencia.lic"); if (File.Exists(f)) token = File.ReadAllText(f); } catch { }
            if (token.Trim().Length == 0) { Cerrar(); msg = "No hay licencia activada en esta PC."; return false; }
            DateTime exp; string cli, err;
            if (!Leer(token, out exp, out cli, out err)) { Cerrar(); msg = err; return false; }
            if (DateTime.Now >= exp) { Cerrar(); msg = "La licencia vencio el " + exp.ToString("dd/MM/yyyy") + "."; return false; }
            long su; string sm;
            StampGet(out su, out sm);
            string e8 = exp.ToString("yyyyMMdd");
            long now = DateTime.UtcNow.Ticks;
            int cmp = string.CompareOrdinal(e8, sm);
            if (cmp < 0) { Cerrar(); msg = "Esta licencia es anterior a otra ya usada en esta PC."; return false; }
            if (cmp == 0 && now < su - TimeSpan.FromMinutes(30).Ticks) { Cerrar(); msg = "El reloj de Windows fue atrasado. Corrige la fecha y hora."; return false; }
            if (cmp > 0 || now > su + TimeSpan.FromMinutes(1).Ticks) StampSet(now, e8);
            _cli = cli;
            Interlocked.Exchange(ref _expTicks, exp.Ticks);
            _ok = true;
            return true;
        }
        static bool Activate(string token, out string msg)
        {
            DateTime exp; string cli, err;
            msg = "";
            if (!Leer(token, out exp, out cli, out err)) { msg = err; return false; }
            if (DateTime.Now >= exp) { msg = "Esa licencia ya vencio (" + exp.ToString("dd/MM/yyyy") + ")."; return false; }
            string f = Path.Combine(_dir, "licencia.lic");
            string old = null;
            try { if (File.Exists(f)) old = File.ReadAllText(f); } catch { }
            File.WriteAllText(f, token.Trim());
            if (Check(out msg)) return true;
            try { if (old != null) File.WriteAllText(f, old); else File.Delete(f); } catch { }
            Check(out err);
            return false;
        }

        static void CopyClip(string t)
        {
            try
            {
                Thread th = new Thread(delegate() { try { System.Windows.Forms.Clipboard.SetText(t); } catch { } });
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
                th.Join(3000);
            }
            catch { }
        }

        // Pide licencia hasta que haya una valida (o el usuario escribe SALIR)
        public static void AskLoop()
        {
            while (true)
            {
                string msg;
                if (Check(out msg))
                {
                    Console.WriteLine(" Licencia OK - " + _cli + " - vence el " + VenceTexto());
                    return;
                }
                CopyClip(MachineCode());
                Console.WriteLine();
                Console.WriteLine("==================================================");
                Console.WriteLine(" LICENCIA REQUERIDA - el sistema esta BLOQUEADO");
                Console.WriteLine(" " + msg);
                Console.WriteLine("");
                Console.WriteLine(" Codigo de esta PC:  " + MachineCode());
                Console.WriteLine(" (ya esta copiado; enviaselo al proveedor)");
                Console.WriteLine("");
                Console.WriteLine(" Pega aqui la licencia y presiona ENTER");
                Console.WriteLine(" (o escribe SALIR para cerrar):");
                Console.WriteLine("==================================================");
                string line = null;
                try { line = Console.ReadLine(); } catch { }
                if (line == null) { Thread.Sleep(2000); continue; }
                line = line.Trim();
                if (line.Equals("SALIR", StringComparison.OrdinalIgnoreCase)) Environment.Exit(0);
                if (line.Length == 0) continue;
                string m2;
                if (!Activate(line, out m2)) Console.WriteLine(" >> " + m2);
            }
        }

        public static void StartMonitor()
        {
            Thread t = new Thread(delegate()
            {
                DateTime lastWarn = DateTime.MinValue;
                while (true)
                {
                    Thread.Sleep(20000);
                    string msg;
                    if (!Check(out msg)) { AskLoop(); lastWarn = DateTime.MinValue; }
                    else if (Dias() <= 5 && (DateTime.Now - lastWarn).TotalMinutes > 60)
                    {
                        Console.WriteLine(" AVISO: la licencia vence el " + VenceTexto() + " (quedan " + Dias() + " dias).");
                        lastWarn = DateTime.Now;
                    }
                }
            });
            t.IsBackground = true;
            t.Start();
        }
    }

    static class Program
    {
        static bool Payload(out string ps, out string js)
        {
            ps = null; js = null;
            try
            {
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload"))
                using (MemoryStream ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    byte[] all = ms.ToArray();
                    byte[] iv = new byte[16];
                    Buffer.BlockCopy(all, 0, iv, 0, 16);
                    using (Aes aes = Aes.Create())
                    {
                        aes.Key = Secrets.Key(); aes.IV = iv;
                        aes.Mode = CipherMode.CBC; aes.Padding = PaddingMode.PKCS7;
                        byte[] plain;
                        using (ICryptoTransform d = aes.CreateDecryptor()) plain = d.TransformFinalBlock(all, 16, all.Length - 16);
                        int n = BitConverter.ToInt32(plain, 0);
                        UTF8Encoding u = new UTF8Encoding(false);
                        ps = u.GetString(plain, 4, n);
                        js = u.GetString(plain, 4 + n, plain.Length - 4 - n);
                    }
                }
                return true;
            }
            catch { return false; }
        }

        static void Run(string dir, string script, string js)
        {
            Runspace rs = RunspaceFactory.CreateRunspace();
            rs.ThreadOptions = PSThreadOptions.UseCurrentThread;
            rs.Open();
            rs.SessionStateProxy.SetVariable("TT_APPDIR", dir);
            rs.SessionStateProxy.SetVariable("TT_XLSX", js);
            using (PowerShell ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                ps.Streams.Information.DataAdded += delegate(object sender, DataAddedEventArgs e)
                {
                    PSDataCollection<InformationRecord> c = (PSDataCollection<InformationRecord>)sender;
                    try { Console.WriteLine(c[e.Index].MessageData); c.Clear(); } catch { }
                };
                ps.Streams.Warning.DataAdded += delegate(object sender, DataAddedEventArgs e)
                {
                    PSDataCollection<WarningRecord> c = (PSDataCollection<WarningRecord>)sender;
                    try { Console.WriteLine("AVISO: " + c[e.Index].Message); c.Clear(); } catch { }
                };
                ps.Streams.Error.DataAdded += delegate(object sender, DataAddedEventArgs e)
                {
                    PSDataCollection<ErrorRecord> c = (PSDataCollection<ErrorRecord>)sender;
                    try { Console.WriteLine("ERROR: " + c[e.Index]); c.Clear(); } catch { }
                };
                ps.AddScript(script);
                ps.Invoke();
            }
            rs.Close();
        }

        [STAThread]
        static int Main()
        {
            Console.Title = "Servidor de Pedidos - Toto Tools";
            try { Console.SetIn(new StreamReader(Console.OpenStandardInput(8192), Console.InputEncoding, false, 8192)); } catch { }
            string dir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            Gate.Init(dir);
            Gate.AskLoop();
            Gate.StartMonitor();
            string ps, js;
            if (!Payload(out ps, out js)) { Console.WriteLine("Programa danado. Pide uno nuevo al proveedor."); Thread.Sleep(15000); return 2; }
            try { Run(dir, ps, js); }
            catch (Exception ex) { Console.WriteLine("Error: " + ex.Message); }
            Console.WriteLine();
            Console.WriteLine("El servidor se detuvo. Presiona una tecla para cerrar.");
            try { Console.ReadKey(); } catch { }
            return 0;
        }
    }
}
