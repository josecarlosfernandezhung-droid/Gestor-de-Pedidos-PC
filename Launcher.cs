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
        static string _modo = "";          // "licencia", "prueba" o "vencida"
        static string _motivo = "";        // por que esta cerrado (para mostrarlo en pantalla)
        static string _licMsg = "";        // problema de la licencia escrita (si la hay)
        static DateTime _trialWrote = DateTime.MinValue;
        static readonly object _lk = new object();
        const string RegPath = @"SOFTWARE\TotoTools";
        public const int TrialDays = 7;    // dias de prueba gratis

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
        public static string Modo() { return _modo; }
        public static string Motivo() { return _motivo; }
        public static string Cliente() { return _cli; }

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
            DateTime ini;
            return LeerFull(token, out exp, out ini, out cli, out err);
        }
        static string NormTok(string t) { return (t ?? "").Trim().Replace("\r", "").Replace("\n", "").Replace(" ", ""); }
        static bool LeerFull(string token, out DateTime exp, out DateTime ini, out string cli, out string err)
        {
            exp = DateTime.MinValue; ini = DateTime.MinValue; cli = ""; err = "";
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
                try { ini = DateTime.ParseExact(f[2], "yyyyMMdd", CultureInfo.InvariantCulture); } catch { ini = DateTime.MinValue; }
                cli = f[4];
                return true;
            }
            catch { err = "Licencia ilegible."; return false; }
        }

        // ---------- licencia.lic: la clave y, si se renovo antes de vencer, los dias que se sumaron ----------
        // Linea 1: la clave. Linea 2 (opcional): +DIAS;MAC  (la MAC la liga a esta clave y a esta PC; si se toca, no vale).
        static void LeerArchivoLic(string raw, out string token, out int bonus)
        {
            token = ""; bonus = 0;
            if (string.IsNullOrEmpty(raw)) return;
            string[] ls = raw.Replace("\r", "").Split('\n');
            for (int i = 0; i < ls.Length; i++)
            {
                string t = ls[i].Trim();
                if (t.Length == 0) continue;
                if (token.Length == 0 && !t.StartsWith("+")) { token = NormTok(t); continue; }
                if (token.Length > 0 && t.StartsWith("+"))
                {
                    string[] p = t.Substring(1).Split(';');
                    int n;
                    if (p.Length == 2 && int.TryParse(p[0], out n) && n > 0 && n <= 3650 && Mac("BONO|" + token + "|" + p[0]) == p[1]) bonus = n;
                }
            }
        }
        static string ArchivoLic(string token, int bonus)
        {
            string s = token;
            if (bonus > 0) s += "\r\n+" + bonus.ToString() + ";" + Mac("BONO|" + token + "|" + bonus.ToString());
            return s;
        }
        static string UsadasFile() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "TotoTools", "lu.dat"); }
        static string UsadaId(string token) { return Mac("USADA|" + token); }
        static bool UsadaTiene(string token)
        {
            string id = UsadaId(token);
            string[] raws = new string[2];
            try { using (RegistryKey k = Registry.LocalMachine.OpenSubKey(RegPath)) { if (k != null) raws[0] = k.GetValue("lu") as string; } } catch { }
            try { if (File.Exists(UsadasFile())) raws[1] = File.ReadAllText(UsadasFile()); } catch { }
            foreach (string r in raws)
            {
                if (string.IsNullOrEmpty(r)) continue;
                foreach (string l in r.Replace("\r", "").Split('\n')) { if (l.Trim() == id) return true; }
            }
            return false;
        }
        static void UsadaPon(string token)
        {
            if (string.IsNullOrEmpty(token) || UsadaTiene(token)) return;
            string id = UsadaId(token);
            string r0 = null, r1 = null;
            try { using (RegistryKey k = Registry.LocalMachine.OpenSubKey(RegPath)) { if (k != null) r0 = k.GetValue("lu") as string; } } catch { }
            try { if (File.Exists(UsadasFile())) r1 = File.ReadAllText(UsadasFile()); } catch { }
            string n0 = (string.IsNullOrEmpty(r0) ? "" : r0.Trim() + "\n") + id;
            string n1 = (string.IsNullOrEmpty(r1) ? "" : r1.Trim() + "\n") + id;
            try { using (RegistryKey k = Registry.LocalMachine.CreateSubKey(RegPath)) k.SetValue("lu", n0); } catch { }
            try { string f = UsadasFile(); Directory.CreateDirectory(Path.GetDirectoryName(f)); File.WriteAllText(f, n1); } catch { }
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

        // ---------- prueba gratis (7 dias desde el primer arranque) ----------
        // La fecha de inicio se guarda firmada en 3 sitios (registro, ProgramData y carpeta del programa).
        // Si falta alguno se repone con los otros; si se borran los 3 a la vez, empieza otra vez (no es a prueba de todo).
        static string TrialMac(string s) { return Mac("TRIAL|" + s); }
        static string TrialFile() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "TotoTools", "tp.dat"); }
        static string TrialFile2() { return Path.Combine(_dir, ".tt_prueba"); }
        static bool TrialParse(string raw, out long start, out long last)
        {
            start = 0; last = 0;
            try
            {
                if (string.IsNullOrEmpty(raw)) return false;
                string[] p = raw.Trim().Split(';');
                if (p.Length != 3 || TrialMac(p[0] + ";" + p[1]) != p[2]) return false;
                start = long.Parse(p[0]); last = long.Parse(p[1]);
                return true;
            }
            catch { return false; }
        }
        static void TrialWrite(long start, long last)
        {
            string body = start.ToString() + ";" + last.ToString();
            string raw = body + ";" + TrialMac(body);
            try { using (RegistryKey k = Registry.LocalMachine.CreateSubKey(RegPath)) k.SetValue("tp", raw); } catch { }
            try { string f = TrialFile(); Directory.CreateDirectory(Path.GetDirectoryName(f)); File.WriteAllText(f, raw); } catch { }
            try
            {
                string f2 = TrialFile2();
                if (File.Exists(f2)) File.SetAttributes(f2, FileAttributes.Normal);
                File.WriteAllText(f2, raw);
                File.SetAttributes(f2, FileAttributes.Hidden);
            }
            catch { }
            _trialWrote = DateTime.Now;
        }
        // true = la prueba sigue abierta (end = cuando termina)
        static bool Trial(out DateTime end, out string msg)
        {
            end = DateTime.MinValue; msg = "";
            string[] raws = new string[3];
            try { using (RegistryKey k = Registry.LocalMachine.OpenSubKey(RegPath)) { if (k != null) raws[0] = k.GetValue("tp") as string; } } catch { }
            try { if (File.Exists(TrialFile())) raws[1] = File.ReadAllText(TrialFile()); } catch { }
            try { if (File.Exists(TrialFile2())) raws[2] = File.ReadAllText(TrialFile2()); } catch { }
            long now = DateTime.UtcNow.Ticks;
            int found = 0, valid = 0;
            long start = long.MaxValue, last = 0;
            foreach (string r in raws)
            {
                if (string.IsNullOrEmpty(r)) continue;
                found++;
                long s, l;
                if (TrialParse(r, out s, out l)) { valid++; if (s < start) start = s; if (l > last) last = l; }
            }
            if (found == 0)
            {
                start = now; last = now;
                TrialWrite(start, last);
            }
            else if (valid == 0)
            {
                msg = "No se pudo comprobar la prueba gratis de esta PC.";
                return false;
            }
            else
            {
                if (now < last - TimeSpan.FromHours(36).Ticks)
                {
                    msg = "El reloj de Windows fue atrasado. Corrige la fecha y hora.";
                    return false;
                }
                bool faltan = (found < 3) || (valid < found);
                if ((faltan && (DateTime.Now - _trialWrote).TotalSeconds > 60) || now > last + TimeSpan.FromMinutes(10).Ticks)
                {
                    if (now > last) last = now;
                    TrialWrite(start, last);
                }
            }
            end = new DateTime(start, DateTimeKind.Utc).AddDays(TrialDays).ToLocalTime();
            if (DateTime.Now >= end) { msg = "La prueba gratis de " + TrialDays.ToString() + " dias termino."; return false; }
            return true;
        }

        // ---------- estado de la licencia ----------
        // Orden: licencia valida > prueba gratis > bloqueado.
        public static bool Check(out string msg)
        {
            lock (_lk) { return CheckLocked(out msg); }
        }
        static bool CheckLocked(out string msg)
        {
            msg = "";
            DateTime tEnd; string tMsg;
            bool trialOk = Trial(out tEnd, out tMsg);
            string token = "";
            int bonus = 0;
            try { string f = Path.Combine(_dir, "licencia.lic"); if (File.Exists(f)) LeerArchivoLic(File.ReadAllText(f), out token, out bonus); } catch { }
            bool hayToken = token.Trim().Length > 0;
            string lmsg = "";
            bool licOk = false;
            DateTime exp = DateTime.MinValue; string cli = "";
            if (!hayToken) lmsg = "No hay licencia activada en esta PC.";
            else
            {
                string err;
                bool leido = Leer(token, out exp, out cli, out err);
                if (leido && bonus > 0) exp = exp.AddDays(bonus);
                if (!leido) lmsg = err;
                else if (DateTime.Now >= exp) lmsg = "La licencia vencio el " + exp.ToString("dd/MM/yyyy") + ".";
                else
                {
                    long su; string sm;
                    StampGet(out su, out sm);
                    string e8 = exp.ToString("yyyyMMdd");
                    long now = DateTime.UtcNow.Ticks;
                    int cmp = string.CompareOrdinal(e8, sm);
                    if (cmp < 0) lmsg = "Esta licencia es anterior a otra ya usada en esta PC.";
                    else if (cmp == 0 && now < su - TimeSpan.FromMinutes(30).Ticks) lmsg = "El reloj de Windows fue atrasado. Corrige la fecha y hora.";
                    else
                    {
                        if (cmp > 0 || now > su + TimeSpan.FromMinutes(1).Ticks) StampSet(now, e8);
                        licOk = true;
                    }
                }
            }
            _licMsg = lmsg;
            if (licOk)
            {
                _cli = cli; _modo = "licencia"; _motivo = "";
                Interlocked.Exchange(ref _expTicks, exp.Ticks);
                _ok = true;
                return true;
            }
            if (trialOk)
            {
                _cli = "Prueba gratis"; _modo = "prueba"; _motivo = hayToken ? lmsg : "";
                Interlocked.Exchange(ref _expTicks, tEnd.Ticks);
                _ok = true;
                return true;
            }
            Cerrar();
            _modo = "vencida";
            msg = hayToken ? (lmsg + " " + tMsg).Trim() : tMsg;
            _motivo = msg;
            return false;
        }
        static bool Activate(string token, out string msg)
        {
            DateTime exp, ini; string cli, err;
            msg = "";
            token = NormTok(token);
            if (!LeerFull(token, out exp, out ini, out cli, out err)) { msg = err; return false; }
            if (DateTime.Now >= exp) { msg = "Esa licencia ya vencio (" + exp.ToString("dd/MM/yyyy") + ")."; return false; }
            string f = Path.Combine(_dir, "licencia.lic");
            string old = null;
            try { if (File.Exists(f)) old = File.ReadAllText(f); } catch { }
            // licencia que esta vigente ahora (con sus dias sumados), si la hay
            string oldTok; int oldBonus;
            LeerArchivoLic(old, out oldTok, out oldBonus);
            DateTime oldExp = DateTime.MinValue; bool oldVig = false;
            if (oldTok.Length > 0)
            {
                DateTime e0, i0; string c0, er0;
                if (LeerFull(oldTok, out e0, out i0, out c0, out er0)) { oldExp = e0.AddDays(oldBonus); oldVig = DateTime.Now < oldExp; }
            }
            if (oldTok == token) { msg = oldVig ? "Esa licencia ya esta activada en esta PC." : "Esa licencia ya se uso en esta PC."; return false; }
            if (UsadaTiene(token)) { msg = "Esa licencia ya se uso en esta PC."; return false; }
            // Renovar antes de vencer: los dias que quedan se suman a los de la licencia nueva.
            int bonus = 0;
            if (oldVig && ini > DateTime.MinValue && exp > ini)
            {
                int dias = (int)Math.Round((exp - ini).TotalDays);
                DateTime sumada = oldExp.AddDays(dias);
                int b = (int)Math.Ceiling((sumada - exp).TotalDays);
                if (b > 0 && b <= 3650) bonus = b;
            }
            File.WriteAllText(f, ArchivoLic(token, bonus));
            string m1;
            bool ok = Check(out m1);
            if (ok && _modo == "licencia")
            {
                UsadaPon(token);
                if (oldTok.Length > 0) UsadaPon(oldTok);
                return true;
            }
            msg = string.IsNullOrEmpty(_licMsg) ? m1 : _licMsg;
            try { if (old != null) File.WriteAllText(f, old); else File.Delete(f); } catch { }
            Check(out err);
            return false;
        }
        // Para activar desde la pagina /licencia del programa: devuelve "" si salio bien, o el motivo del error.
        public static string Activar(string token)
        {
            try
            {
                lock (_lk)
                {
                    string m;
                    if (Activate(token ?? "", out m)) return "";
                    return string.IsNullOrEmpty(m) ? "No se pudo activar la licencia." : m;
                }
            }
            catch (Exception ex) { return "No se pudo activar: " + ex.Message; }
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
                if (Check(out msg) && _modo == "licencia")
                {
                    Console.WriteLine(" Licencia OK - " + _cli + " - vence el " + VenceTexto());
                    return;
                }
                if (_ok) return;
                CopyClip(MachineCode());
                Console.WriteLine();
                Console.WriteLine("==================================================");
                Console.WriteLine(" LICENCIA REQUERIDA - el sistema esta BLOQUEADO");
                Console.WriteLine(" (tambien puedes activarla desde el navegador: pagina /licencia)");
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

        static int _ask = 0;
        // Pide la licencia por la consola (opcional: tambien se puede activar desde la pagina /licencia).
        public static void StartAsk()
        {
            if (Interlocked.CompareExchange(ref _ask, 1, 0) != 0) return;
            Thread t = new Thread(delegate() { try { AskLoop(); } finally { Interlocked.Exchange(ref _ask, 0); } });
            t.IsBackground = true;
            t.Start();
        }
        public static void Aviso()
        {
            if (_modo == "licencia") Console.WriteLine(" Licencia OK - " + _cli + " - vence el " + VenceTexto());
            else if (_modo == "prueba") Console.WriteLine(" PRUEBA GRATIS: quedan " + Dias() + " dias (hasta el " + VenceTexto() + "). Activa la licencia desde la pantalla de Ayuda o la pagina /licencia.");
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
                    if (!Check(out msg)) { StartAsk(); lastWarn = DateTime.MinValue; }
                    else if (((_modo == "licencia" && Dias() <= 5) || _modo == "prueba") && (DateTime.Now - lastWarn).TotalMinutes > 60)
                    {
                        if (_modo == "prueba") Console.WriteLine(" AVISO: prueba gratis, quedan " + Dias() + " dias (hasta el " + VenceTexto() + ").");
                        else Console.WriteLine(" AVISO: la licencia vence el " + VenceTexto() + " (quedan " + Dias() + " dias).");
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
            string m0;
            if (Gate.Check(out m0)) Gate.Aviso();
            else
            {
                Console.WriteLine(" " + m0);
                Console.WriteLine(" Codigo de esta PC: " + Gate.MachineCode());
                Gate.StartAsk();
            }
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
