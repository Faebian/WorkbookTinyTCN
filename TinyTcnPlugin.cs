using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using UnicornOne.Abstractions.Workbook;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Threading;

namespace WorkbookTinyTCN;

// VS Code useful shortcuts :)
// Ctrl + K, Ctrl + 0 = collapse all
// Shift + Alt + F = fortmat all

sealed class InferResponse
{
    public string name { get; set; }

    public double t0 { get; set; }
    public double t1 { get; set; }
    public double t2 { get; set; }
    public double t3 { get; set; }

    public double[] conf { get; set; }
}

public sealed class TinyTcnPlugin :
    IWorkbookPlugin,
    IWorkbookController,
    IDisposable
{
    private IWorkbookHost _host;
    private Panel _panel;
    private Chart _chart;
    private Button _openButton;
    private Button _inferButton;
    private Button _liveButton;
    private CancellationTokenSource _liveCts;
    private DateTime _liveStartTime;


    public string Name => "TinyTCN";
    public string Version => "0.3";

    private readonly HttpClient _httpClient = new HttpClient();
    private InferResponse _lastInference;

    public void Initialize(string jsonConfig, IWorkbookHost host)
    {
        _host = host;
        _host.AppendLog("TinyTCN", "Info", "TinyTCN plugin initialized");

        // you can process any json fields here...

    }

    public void BuildUI()
    {
        _panel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(8)
        };

        _host.Theme.ApplyBaseStyles(_panel);

        var topPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 45,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(4)
        };

        _host.Theme.ApplyBaseStyles(topPanel);

        _openButton = new Button
        {
            Text = "Open",
            Width = 90,
            Height = 30
        };

        _inferButton = new Button
        {
            Text = "Infer",
            Width = 90,
            Height = 30
        };

        _liveButton = new Button
        {
            Text = "Live",
            Width = 90,
            Height = 30
        };

        _openButton.Click += OpenButton_Click;
        _inferButton.Click += InferButton_Click;
        _liveButton.Click += LiveButton_Click;

        topPanel.Controls.Add(_openButton);
        topPanel.Controls.Add(_inferButton);
        topPanel.Controls.Add(_liveButton);

        _chart = CreateChart();

        _panel.Controls.Add(_chart);
        _panel.Controls.Add(topPanel);

        _host.AddWidget(_panel, "TinyTCN");
    }

    private Chart CreateChart()
    {
        var chart = new Chart
        {
            Dock = DockStyle.Fill,
            BackColor = _host.Theme.Back
        };

        var area = new ChartArea("Main");

        area.BackColor = _host.Theme.Back;
        area.AxisX.Title = "Time / X";
        area.AxisY.Title = "Signal";

        area.AxisX.LineColor = _host.Theme.Fore;
        area.AxisY.LineColor = _host.Theme.Fore;
        area.AxisX.LabelStyle.ForeColor = _host.Theme.Fore;
        area.AxisY.LabelStyle.ForeColor = _host.Theme.Fore;
        area.AxisX.TitleForeColor = _host.Theme.Fore;
        area.AxisY.TitleForeColor = _host.Theme.Fore;

        area.AxisX.MajorGrid.LineColor = Color.FromArgb(80, 80, 80);
        area.AxisY.MajorGrid.LineColor = Color.FromArgb(80, 80, 80);

        area.AxisX.Minimum = 0;
        area.AxisX.Maximum = 30;
        area.AxisY.Minimum = 0;
        area.AxisY.Maximum = 1;

        chart.ChartAreas.Add(area);

        var csvSeries = new Series("CSV")
        {
            ChartType = SeriesChartType.FastLine,
            XValueType = ChartValueType.Double,
            BorderWidth = 2,
            Color = Color.LimeGreen,
            IsXValueIndexed = false
        };

        chart.Series.Add(csvSeries);

        chart.Series[0].Points.AddXY(0.0, 0.0);

        UpdateUiState();

        return chart;
    }

    #region CSV

    private void OpenButton_Click(object sender, EventArgs e)
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Filter = "CSV files|*.csv;*.txt|All files|*.*";
            ofd.Title = "Open CSV";

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            LoadCsvToChart(ofd.FileName);
        }
    }

    private void LoadCsvToChart(string path)
    {
        var lines = File.ReadAllLines(path);

        if (lines.Length < 2)
            return;

        char delimiter = DetectDelimiter(lines[0]);
        string[] headers = lines[0].Split(delimiter);

        int xIndex = 0;
        int yIndex = headers.Length > 1 ? 1 : 0;

        var series = _chart.Series["CSV"];
        series.Points.Clear();

        for (int i = 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
                continue;

            string[] parts = lines[i].Split(delimiter);

            if (parts.Length <= Math.Max(xIndex, yIndex))
                continue;

            if (!double.TryParse(parts[xIndex], NumberStyles.Float, CultureInfo.InvariantCulture, out double x))
                continue;

            if (!double.TryParse(parts[yIndex], NumberStyles.Float, CultureInfo.InvariantCulture, out double y))
                continue;

            series.Points.AddXY(x, y);
        }

        AutoScaleChart();

        UpdateUiState();

        _host.AppendLog("TinyTCN", "Info", $"Loaded CSV: {Path.GetFileName(path)}");
    }

    private char DetectDelimiter(string line)
    {
        char[] candidates = { ',', ';', '\t' };

        return candidates
            .OrderByDescending(c => line.Count(ch => ch == c))
            .First();
    }

    private void AutoScaleChart()
    {
        var series = _chart.Series["CSV"];

        if (series.Points.Count == 0)
            return;

        double minX = series.Points.Min(p => p.XValue);
        double maxX = series.Points.Max(p => p.XValue);
        double minY = series.Points.Min(p => p.YValues[0]);
        double maxY = series.Points.Max(p => p.YValues[0]);

        if (minX == maxX)
            maxX = minX + 1;

        if (minY == maxY)
            maxY = minY + 1;

        var area = _chart.ChartAreas[0];

        area.AxisX.Minimum = Math.Floor(minX);
        area.AxisX.Maximum = Math.Ceiling(maxX);

        area.AxisY.Minimum = Math.Floor(minY);
        area.AxisY.Maximum = Math.Ceiling(maxY);

        _chart.Invalidate();
    }

    #endregion

    #region Infer

    private void UpdateUiState()
    {
        bool hasData =
            _chart != null &&
            _chart.Series.IndexOf("CSV") >= 0 &&
            _chart.Series["CSV"].Points.Count > 10;

        _inferButton.Enabled = hasData;
    }

    private async void InferButton_Click(object sender, EventArgs e)
    {
        await RunInference();
    }

    private async Task RunInference()
    {
        try
        {
            _inferButton.Enabled = false;

            string csv = BuildInferenceCsv();

            using (var content = new MultipartFormDataContent())
            {
                content.Add(
                    new StringContent("WorkbookTinyTCN"),
                    "name");

                content.Add(
                    new StringContent(csv),
                    "csv");

                var response = await _httpClient.PostAsync(
                    "http://127.0.0.1:11880/infer_csv",
                    content);

                string json = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    MessageBox.Show(json, "Inference Error");
                    return;
                }

                ApplyInference(json);

                _host.AppendLog(
                    "TinyTCN",
                    "Info",
                    "Inference complete");
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        finally
        {
            _inferButton.Enabled = true;
        }
    }

    private string BuildInferenceCsv()
    {
        var ordered = _chart.Series["CSV"].Points
            .Select(p => new { T = p.XValue, I = p.YValues[0] })
            .OrderBy(p => p.T)
            .ToList();

        double[] times = ordered.Select(p => p.T).ToArray();
        double[] intensity = ordered.Select(p => p.I).ToArray();

        double[] smooth = SmoothEmaZeroPhaseTime(times, intensity, 1.0); // build a local version that uses double[] throughout
        double[] slope = RollingSlope(times, smooth, windowSamples: 11);
        slope = MovingAverage(slope, span: 7);

        var sb = new StringBuilder();
        sb.AppendLine("t,I_fixed,I_smooth,slope,dx,dy");

        for (int i = 0; i < ordered.Count; i++)
        {
            sb.AppendLine(string.Format(
                CultureInfo.InvariantCulture,
                "{0:F6},{1:F6},{2:F6},{3:F6},0,0",
                times[i],
                intensity[i],
                smooth[i],
                slope[i]));
        }

        string csvContents = sb.ToString();

        // Debug dump to application root
        string dumpPath = Path.Combine(Application.StartupPath, "dump.csv");
        File.WriteAllText(dumpPath, csvContents, Encoding.UTF8);

        return csvContents;
    }

    private static double[] SmoothEmaZeroPhaseTime(double[] times, double[] values, double tauSec)
    {
        if (times == null)
            throw new ArgumentNullException(nameof(times));

        if (values == null)
            throw new ArgumentNullException(nameof(values));

        if (times.Length != values.Length)
            throw new ArgumentException("times and values must have the same length");

        int n = values.Length;

        if (n == 0 || tauSec <= 0)
            return values.ToArray();

        // Forward EMA
        var yf = new double[n];
        yf[0] = values[0];

        for (int i = 1; i < n; i++)
        {
            double dt = Math.Max(0.0, times[i] - times[i - 1]);
            double a = 1.0 - Math.Exp(-dt / tauSec);

            yf[i] = a * values[i] + (1.0 - a) * yf[i - 1];
        }

        // Backward EMA (zero phase)
        var yb = new double[n];
        yb[n - 1] = yf[n - 1];

        for (int i = n - 2; i >= 0; i--)
        {
            double dt = Math.Max(0.0, times[i + 1] - times[i]);
            double a = 1.0 - Math.Exp(-dt / tauSec);

            yb[i] = a * yf[i] + (1.0 - a) * yb[i + 1];
        }

        return yb;
    }

    private double[] RollingSlope(double[] x, double[] y, int windowSamples)
    {
        var slope = new double[y.Length];
        int half = Math.Max(1, windowSamples / 2);

        for (int i = 0; i < y.Length; i++)
        {
            int a = Math.Max(0, i - half);
            int b = Math.Min(y.Length - 1, i + half);

            int n = b - a + 1;
            if (n < 2)
            {
                slope[i] = 0;
                continue;
            }

            double sx = 0, sy = 0, sxx = 0, sxy = 0;

            for (int j = a; j <= b; j++)
            {
                sx += x[j];
                sy += y[j];
                sxx += x[j] * x[j];
                sxy += x[j] * y[j];
            }

            double denom = n * sxx - sx * sx;
            slope[i] = Math.Abs(denom) < 1e-12 ? 0 : (n * sxy - sx * sy) / denom;
        }

        return slope;
    }

    private double[] MovingAverage(double[] values, int span)
    {
        var output = new double[values.Length];
        int half = Math.Max(1, span / 2);

        for (int i = 0; i < values.Length; i++)
        {
            int a = Math.Max(0, i - half);
            int b = Math.Min(values.Length - 1, i + half);

            double sum = 0;
            int count = 0;

            for (int j = a; j <= b; j++)
            {
                sum += values[j];
                count++;
            }

            output[i] = count > 0 ? sum / count : values[i];
        }

        return output;
    }

    private void ApplyInference(string json)
    {
        var result = JsonConvert.DeserializeObject<InferResponse>(json);

        if (result == null)
            throw new InvalidOperationException("Inference response was empty.");

        _lastInference = result;

        DrawMarker("t0", result.t0, Color.Yellow);
        DrawMarker("t1", result.t1, Color.Red);
        DrawMarker("t2", result.t2, Color.Yellow);
        DrawMarker("t3", result.t3, Color.Red);

        _chart.Invalidate();

        string confText = result.conf == null
            ? ""
            : $" conf=[{string.Join(", ", result.conf.Select(c => c.ToString("0.00")))}]";

        _host.AppendLog(
            "TinyTCN",
            "Info",
            $"Inference: t0={result.t0:0.###}, t1={result.t1:0.###}, t2={result.t2:0.###}, t3={result.t3:0.###}{confText}");
    }

    private void DrawMarker(string name, double x, Color color)
    {
        var area = _chart.ChartAreas[0];

        var existing =
            area.AxisX.StripLines
                .FirstOrDefault(s =>
                    string.Equals(
                        s.Text,
                        name,
                        StringComparison.OrdinalIgnoreCase));

        if (existing != null)
            area.AxisX.StripLines.Remove(existing);

        area.AxisX.StripLines.Add(
            new StripLine
            {
                Text = name,
                IntervalOffset = x,
                StripWidth = 0.01,
                BorderColor = color,
                BorderWidth = 2,
                BorderDashStyle = ChartDashStyle.Dash
            });
    }

    #endregion

    #region Live

    private void LiveButton_Click(object sender, EventArgs e)
    {
        if (_liveCts != null)
        {
            StopLive();
            return;
        }

        StartLive();
    }

    private void StartLive()
    {
        var s = _chart.Series["CSV"];
        s.Points.Clear();

        ClearMarkers();

        _liveStartTime = DateTime.Now;
        _liveCts = new CancellationTokenSource();

        _liveButton.Text = "Stop Live";
        _inferButton.Enabled = false;

        _host.AppendLog("TinyTCN", "Info", "Started live laser trace");

        _ = LiveLoopAsync(_liveCts.Token);
    }

    private void StopLive()
    {
        try
        {
            _liveCts?.Cancel();
            _liveCts?.Dispose();
        }
        catch { }

        _liveCts = null;

        _liveButton.Text = "Live";
        UpdateUiState();

        _host.AppendLog("TinyTCN", "Info", "Stopped live laser trace");
    }

    private async Task LiveLoopAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                object valueObj = _host.GetValue("M1.Laser.Sensor.MV");

                if (TryGetDouble(valueObj, out double y))
                {
                    double x = (DateTime.Now - _liveStartTime).TotalSeconds;

                    _host.InvokeUI(() =>
                    {
                        var s = _chart.Series["CSV"];
                        s.Points.AddXY(x, y);

                        UpdateChartAxesFromSeries();
                        _chart.Invalidate();
                    });
                }

                await Task.Delay(200, token);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _host.AppendLog("TinyTCN", "Warning", $"Live sample failed: {ex.Message}");
                await Task.Delay(1000, token);
            }
        }
    }

    private bool TryGetDouble(object valueObj, out double value)
    {
        value = 0;

        if (valueObj == null)
            return false;

        if (valueObj is double d)
        {
            value = d;
            return true;
        }

        if (valueObj is float f)
        {
            value = f;
            return true;
        }

        if (valueObj is int i)
        {
            value = i;
            return true;
        }

        return double.TryParse(
            valueObj.ToString(),
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out value);
    }

    private void UpdateChartAxesFromSeries()
    {
        var s = _chart.Series["CSV"];

        if (s.Points.Count == 0)
            return;

        double maxX = s.Points.Max(p => p.XValue);
        double minY = s.Points.Min(p => p.YValues[0]);
        double maxY = s.Points.Max(p => p.YValues[0]);

        if (minY == maxY)
            maxY = minY + 1;

        var area = _chart.ChartAreas[0];

        area.AxisX.Minimum = 0;
        area.AxisX.Maximum = Math.Max(30, Math.Ceiling(maxX + 1));

        area.AxisY.Minimum = Math.Floor(minY);
        area.AxisY.Maximum = Math.Ceiling(maxY);
    }

    private void ClearMarkers()
    {
        var area = _chart.ChartAreas[0];
        area.AxisX.StripLines.Clear();
    }

    #endregion

    #region Commands

    private void OpenShutter(string name)
    {
        // name = Ga1, In1, etc   
        _host.ExecuteCommand($"Open({name})");
    }

    private void CloseShutter(string name)
    {
        // name = Ga1, In1, etc      
        _host.ExecuteCommand($"Close({name})");
    }

    #endregion

    public void Run()
    {
        // Fire and forget
        // This starts the task and returns immediately
        // The workbook is "complete" when the _host.MarkComplete() is called in RunAsync
        _ = RunAsync();
    }

    private async Task RunAsync()
    {
        string path = @"D:\UnicornOne\DataSetGenerator\CSV\flux_0.68.csv";

        _inferButton.Enabled = false;

        try
        {
            LoadCsvToChart(path);   // UI thread
            await RunInference();   // async; resumes on UI thread unless configured otherwise
            _host.MarkComplete();   // actually completes the workbook 
        }
        catch (Exception ex)
        {
            _host.AppendLog("TinyTCN", "Warning", ex.ToString());
            _host.MarkComplete();   // unblock a failed workbook execution 
        }
        finally
        {
            _inferButton.Enabled = true;
        }
    }


    public void RequestCancel()
    {
    }

    public void OnShown()
    {
    }

    public void OnHidden()
    {
    }

    public void Dispose()
    {
        StopLive();
        _httpClient?.Dispose();
    }
}