using Microsoft.Win32;
using System;
using System.ComponentModel;
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

        private static readonly Regex IpRegex = new Regex("\\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b", RegexOptions.ExplicitCapture | RegexOptions.Compiled);

        private readonly string _oneDriveLocation;

        private readonly Timer _timer;

        private WebClient _client;

        private string _lastIp;

        private readonly TimeSpan _interval;

        private EventLog eventLog;

        public HomeConnectService()
        {
            this.InitializeComponent();
            this._oneDriveLocation = this.GetOneDriveLocation();
            this._timer = new Timer(new TimerCallback(this.Callback), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            this._interval = TimeSpan.Parse(ConfigurationManager.AppSettings["Interval"]);
            this.Log(string.Format("Создана задача обновления скрипта по пути {0} с интервалом {1}", this._oneDriveLocation, this._interval), EventLogEntryType.Information);
        }

        private static string TryGetOneDrivePath()
        {
            SecurityIdentifier securityIdentifier = (SecurityIdentifier)new NTAccount(ConfigurationManager.AppSettings["UserName"]).Translate(typeof(SecurityIdentifier));
            return "HKEY_USERS\\" + securityIdentifier.Value + "\\Software\\Microsoft\\OneDrive";
        }

        protected override void OnStart(string[] args)
        {
            this._client = new WebClient();
            this._timer.Change(TimeSpan.Zero, this._interval);
        }

        protected override void OnStop()
        {
            this._timer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            this._client.Dispose();
        }

        private void Callback(object state)
        {
            try
            {
                if (this._client.IsBusy)
                {
                    this.Log("Выполняется предыдущий запрос. Пропускаем обработку", EventLogEntryType.Error);
                }
                else
                {
                    string text = this._client.DownloadString("http://checkip.dyndns.org");
                    Match match = HomeConnectService.IpRegex.Match(text);
                    if (!match.Success)
                    {
                        this.Log(string.Format("Регулярное выражение не нашло IP-адреса в ответе сервера: {0}", text), EventLogEntryType.Error);
                    }
                    else
                    {
                        string value = match.Value;
                        if (this._lastIp == value)
                        {
                            this.Log(string.Format("IP-адрес не изменился. Время следующего пробуждения: {0}", DateTime.Now.Add(this._interval).ToShortTimeString()), EventLogEntryType.Information);
                        }
                        else
                        {
                            this.Log(string.Format("IP-адрес изменился на {0}, обновляем скрипт. Ответ сервера: {1}", value, text), EventLogEntryType.Warning);
                            this._lastIp = value;
                            string contents = string.Format("start \"\" mstsc /v:\"{0}\"", value);
                            File.WriteAllText(Path.Combine(this._oneDriveLocation, "homeconnect.bat"), contents, Encoding.ASCII);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.Log(ex.Message, EventLogEntryType.Error);
            }
        }

        private void Log(string entry, EventLogEntryType type)
        {
            this.EventLog.WriteEntry(entry, type);
        }

        public string GetOneDriveLocation()
        {
            return (string)Registry.GetValue(HomeConnectService.TryGetOneDrivePath(), "UserFolder", null);
        }
    }
}
