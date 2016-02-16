using Microsoft.Win32;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace HomeConnect
{
    public partial class HomeConnectService : ServiceBase
    {
        private const string Pattern = "\\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b";

        private static readonly Regex IpRegex = new Regex(Pattern, RegexOptions.ExplicitCapture | RegexOptions.Compiled);

        private readonly string _oneDriveLocation;

        private readonly Timer _timer;

        private WebClient _client;

        private string _lastIp;

        private readonly TimeSpan _interval;

        public HomeConnectService()
        {
            InitializeComponent();
            _oneDriveLocation = GetOneDriveLocation();
            _timer = new Timer(Callback, null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            _interval = TimeSpan.Parse(ConfigurationManager.AppSettings["Interval"]);
            Log($"Создана задача обновления скрипта по пути {_oneDriveLocation} с интервалом {_interval}", EventLogEntryType.Information);
        }

        private static string TryGetOneDrivePath()
        {
            SecurityIdentifier securityIdentifier = (SecurityIdentifier)new NTAccount(ConfigurationManager.AppSettings["UserName"]).Translate(typeof(SecurityIdentifier));
            return "HKEY_USERS\\" + securityIdentifier.Value + "\\Software\\Microsoft\\OneDrive";
        }

        protected override void OnStart(string[] args)
        {
            _client = new WebClient();
            _timer.Change(TimeSpan.Zero, _interval);
        }

        protected override void OnStop()
        {
            _timer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            _client.Dispose();
        }

        private void Callback(object state)
        {
            try
            {
                if (_client.IsBusy)
                {
                    Log("Выполняется предыдущий запрос. Пропускаем обработку", EventLogEntryType.Error);
                }
                else
                {
                    string text = _client.DownloadString("http://checkip.dyndns.org");
                    Match match = IpRegex.Match(text);
                    if (!match.Success)
                    {
                        Log($"Регулярное выражение не нашло IP-адреса в ответе сервера: {text}", EventLogEntryType.Error);
                    }
                    else
                    {
                        string value = match.Value;
                        if (_lastIp == value)
                        {
                            Log($"IP-адрес не изменился. Время следующего пробуждения: {DateTime.Now.Add(_interval).ToShortTimeString()}", EventLogEntryType.Information);
                        }
                        else
                        {
                            Log($"IP-адрес изменился на {value}, обновляем скрипт. Ответ сервера: {text}", EventLogEntryType.Warning);
                            _lastIp = value;
                            string contents = $"start \"\" mstsc /v:\"{value}\"";
                            File.WriteAllText(Path.Combine(_oneDriveLocation, "homeconnect.bat"), contents, Encoding.ASCII);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, EventLogEntryType.Error);
            }
        }

        private void Log(string entry, EventLogEntryType type)
        {
            EventLog.WriteEntry(entry, type);
        }

        public string GetOneDriveLocation()
        {
            return (string)Registry.GetValue(TryGetOneDrivePath(), "UserFolder", null);
        }
    }
}
