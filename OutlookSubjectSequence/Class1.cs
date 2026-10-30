
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

internal sealed class MailSubjectSequenceWatcher : IDisposable
{
    private readonly Outlook.Inspectors _inspectors;
    private readonly List<MailInspectorWatcher> _watchers =
        new List<MailInspectorWatcher>();

    public MailSubjectSequenceWatcher(Outlook.Application application)
    {
        _inspectors = application.Inspectors;
        _inspectors.NewInspector += OnNewInspector;
    }

    private void OnNewInspector(Outlook.Inspector inspector)
    {
        try
        {
            if (!(inspector.CurrentItem is Outlook.MailItem mail))
                return;

            var watcher = new MailInspectorWatcher(inspector, mail);
            _watchers.Add(watcher);
            watcher.Closed += OnWatcherClosed;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                "Outlook watcher error: " + ex.Message);
        }
    }

    private void OnWatcherClosed(MailInspectorWatcher watcher)
    {
        watcher.Closed -= OnWatcherClosed;
        _watchers.Remove(watcher);
        watcher.Dispose();
    }

    public void Dispose()
    {
        _inspectors.NewInspector -= OnNewInspector;

        foreach (var watcher in _watchers.ToArray())
            watcher.Dispose();

        _watchers.Clear();
    }
}

internal sealed class MailInspectorWatcher : IDisposable
{
    private static readonly HttpClient Http = new HttpClient
    {
        Timeout = TimeSpan.FromSeconds(20)
    };

    private readonly Outlook.Inspector _inspector;
    private readonly Outlook.MailItem _mail;
    private readonly Outlook.ItemEvents_10_Event _itemEvents;
    private readonly Outlook.InspectorEvents_10_Event _inspectorEvents;
    private readonly Timer _debounceTimer;

    // Satu idempotency key untuk satu draft dan satu proses request.
    private readonly string _idempotencyKey =
        Guid.NewGuid().ToString("N");

    private bool _requestInProgress;
    private bool _numberIssued;
    private bool _disposed;

    public event Action<MailInspectorWatcher> Closed;

    public MailInspectorWatcher(
        Outlook.Inspector inspector,
        Outlook.MailItem mail)
    {
        _inspector = inspector;
        _mail = mail;

        _itemEvents = (Outlook.ItemEvents_10_Event)mail;
        _inspectorEvents = (Outlook.InspectorEvents_10_Event)inspector;

        _itemEvents.PropertyChange += OnPropertyChange;
        _inspectorEvents.Close += OnInspectorClose;

        _debounceTimer = new Timer { Interval = 500 };
        _debounceTimer.Tick += OnDebounceTick;
    }

    private void OnPropertyChange(string propertyName)
    {
        if (_disposed || propertyName != "Subject")
            return;

        _debounceTimer.Stop();
        _debounceTimer.Start();
    }

    private async void OnDebounceTick(object sender, EventArgs e)
    {
        _debounceTimer.Stop();

        if (_disposed || _numberIssued || _requestInProgress)
            return;

        string subject;

        try
        {
            subject = _mail.Subject ?? "";
        }
        catch
        {
            return;
        }

        // Contoh prefix yang diterima: IT OUT, BD INQ.
        // Prefix harus berada di awal Subject.
        Match prefixMatch = Regex.Match(
            subject,
            @"^\s*([A-Z0-9]+)\s+([A-Z0-9]+)(?:\s|$)",
            RegexOptions.IgnoreCase);

        if (!prefixMatch.Success)
            return;

        string departmentCode =
            prefixMatch.Groups[1].Value.ToUpperInvariant();

        string documentTypeCode =
            prefixMatch.Groups[2].Value.ToUpperInvariant();

        _requestInProgress = true;

        try
        {
            string number = await RequestNumberAsync(
                departmentCode,
                documentTypeCode);

            if (_disposed)
                return;

            // Ambil Subject terbaru agar teks yang diketik saat request
            // berlangsung tidak hilang.
            string latestSubject = _mail.Subject ?? "";

            Match latestPrefix = Regex.Match(
                latestSubject,
                @"^\s*([A-Z0-9]+)\s+([A-Z0-9]+)(?:\s|$)",
                RegexOptions.IgnoreCase);

            // Jangan ubah Subject jika prefix sudah dihapus atau diganti
            // ketika request masih berlangsung.
            if (!latestPrefix.Success ||
                !string.Equals(
                    latestPrefix.Groups[1].Value,
                    departmentCode,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    latestPrefix.Groups[2].Value,
                    documentTypeCode,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string remainder =
                latestSubject.Substring(latestPrefix.Length)
                    .TrimStart();

            _numberIssued = true;

            _mail.Subject = "[" + number + "]" +
                (remainder.Length > 0 ? " " + remainder : " ");

            // Draft tidak disimpan otomatis.
        }
        catch (Exception ex)
        {
            if (!_disposed)
            {
                MessageBox.Show(
                    "Gagal meminta nomor dari Odoo.\n\n" +
                    ex.Message,
                    "Outlook Subject Sequence",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        finally
        {
            _requestInProgress = false;
        }
    }

    private async Task<string> RequestNumberAsync(
        string departmentCode,
        string documentTypeCode)
    {
        ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12;

        string url =
            ConfigurationManager.AppSettings["OdooApiUrl"];

        string token =
            ConfigurationManager.AppSettings["OdooApiToken"];

        if (string.IsNullOrWhiteSpace(url) ||
            string.IsNullOrWhiteSpace(token) ||
            token == "REPLACE_WITH_TEST_TOKEN")
        {
            throw new InvalidOperationException(
                "OdooApiUrl atau OdooApiToken belum dikonfigurasi.");
        }

        if (!url.StartsWith(
            "https://",
            StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "API harus menggunakan HTTPS.");
        }

        var payload = new Dictionary<string, object>
        {
            { "department_code", departmentCode },
            { "document_type_code", documentTypeCode },
            { "idempotency_key", _idempotencyKey }
        };

        var serializer = new JavaScriptSerializer();
        string json = serializer.Serialize(payload);

        using (var request = new HttpRequestMessage(
            HttpMethod.Post, url))
        {
            request.Headers.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            request.Content = new StringContent(
                json, Encoding.UTF8, "application/json");

            using (HttpResponseMessage response =
                await Http.SendAsync(request))
            {
                string body =
                    await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(
                        "API Odoo mengembalikan HTTP " +
                        (int)response.StatusCode + ": " + body);
                }

                var result =
                    serializer.Deserialize<Dictionary<string, object>>(
                        body);

                if (result == null ||
                    !result.ContainsKey("ok") ||
                    !Convert.ToBoolean(result["ok"]) ||
                    !result.ContainsKey("subject_number"))
                {
                    throw new InvalidOperationException(
                        "Respons API tidak sesuai format: " + body);
                }

                string number =
                    Convert.ToString(result["subject_number"]);

                if (string.IsNullOrWhiteSpace(number))
                {
                    throw new InvalidOperationException(
                        "API tidak mengembalikan subject_number.");
                }

                return number;
            }
        }
    }

    private void OnInspectorClose()
    {
        Closed?.Invoke(this);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        _debounceTimer.Stop();
        _debounceTimer.Tick -= OnDebounceTick;
        _debounceTimer.Dispose();

        _itemEvents.PropertyChange -= OnPropertyChange;
        _inspectorEvents.Close -= OnInspectorClose;
    }
}
