using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;
using Tesseract;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ControlFacturacion;
   
public sealed class MainForm : Form
{
    private const string ExcelPath = @"C:\Users\Lalá\Downloads\Siesta.xlsx";
    private readonly Button selectButton = new() { Text = "1. Seleccionar área FACTURADO", AutoSize = true };
    private readonly Button readButton = new() { Text = "2. Leer importe", AutoSize = true, Enabled = false };
    private readonly Button save13Button = new() { Text = "Guardar como 13:00", AutoSize = true, Enabled = false };
    private readonly Button save16Button = new() { Text = "Guardar como 16:00", AutoSize = true, Enabled = false };
    private readonly Label regionLabel = new() { AutoSize = true, Text = "Área: todavía no seleccionada" };
    private readonly TextBox rawText = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
    private readonly Label resultLabel = new() { AutoSize = true, Font = new Font("Segoe UI", 13, FontStyle.Bold), Text = "Importe detectado: —" };
    private readonly Label statusLabel = new() { AutoSize = true, Text = "Primero seleccioná el área del importe." };
    private readonly System.Windows.Forms.Timer scheduleTimer = new() { Interval = 15000 };
    private Rectangle selectedRegion;
    private readonly string settingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ControlFacturacion", "region.txt");
    private readonly string[] startupArgs = Environment.GetCommandLineArgs();
    private bool automaticMode;
    private int automaticColumn;
    private string automaticTimeLabel = "";
    private decimal? lastAmount;
    private DateTime lastReadDate;
    private bool saved13Today;
    private bool saved16Today;

    public MainForm()
    {
        Text = "ControlFacturacion — OCR y Excel";
        Width = 760;
        Height = 500;
        StartPosition = FormStartPosition.CenterScreen;

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12), WrapContents = true };
        top.Controls.Add(selectButton);
        top.Controls.Add(readButton);
        top.Controls.Add(save13Button);
        top.Controls.Add(save16Button);

        var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(12), WrapContents = true };
        bottom.Controls.Add(resultLabel);
        bottom.Controls.Add(statusLabel);

        Controls.Add(rawText);
        Controls.Add(top);
        Controls.Add(bottom);

        selectButton.Click += (_, _) => SelectRegion();
        readButton.Click += (_, _) => ReadAmount();
        save13Button.Click += (_, _) => SaveLastAmount(3, "13:00");
        save16Button.Click += (_, _) => SaveLastAmount(4, "16:00");
        scheduleTimer.Tick += (_, _) => CheckSchedule();
        scheduleTimer.Start();
        FormClosed += (_, _) => scheduleTimer.Stop();

        LoadSavedRegion();
        automaticMode = startupArgs.Any(a => a.Equals("--auto13", StringComparison.OrdinalIgnoreCase)
                                          || a.Equals("--auto16", StringComparison.OrdinalIgnoreCase));
        if (startupArgs.Any(a => a.Equals("--auto13", StringComparison.OrdinalIgnoreCase)))
        {
            automaticColumn = 3;
            automaticTimeLabel = "13:00";
        }
        else if (startupArgs.Any(a => a.Equals("--auto16", StringComparison.OrdinalIgnoreCase)))
        {
            automaticColumn = 4;
            automaticTimeLabel = "16:00";
        }

        if (automaticMode)
        {
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Shown += async (_, _) =>
            {
                Hide();
                await System.Threading.Tasks.Task.Delay(1800);
                RunAutomaticCapture(automaticColumn, automaticTimeLabel);
            };
        }
    }


    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private static IntPtr FindServerWindow()
    {
        foreach (var process in Process.GetProcessesByName("servidor"))
        {
            try
            {
                if (process.MainWindowHandle != IntPtr.Zero)
                    return process.MainWindowHandle;
            }
            catch { }
        }

        return IntPtr.Zero;
    }

    private Bitmap CaptureServerArea()
    {
        IntPtr hwnd = FindServerWindow();

        if (hwnd == IntPtr.Zero)
            throw new InvalidOperationException(
                "No encontré servidor.exe. Abrilo y entrá en Cierres → Preliminar.");

        if (!GetWindowRect(hwnd, out RECT rect))
            throw new InvalidOperationException(
                "No pude obtener las dimensiones de servidor.exe.");

        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        if (width <= 0 || height <= 0)
            throw new InvalidOperationException(
                "La ventana de servidor.exe no tiene dimensiones válidas.");

        // selectedRegion ahora representa coordenadas relativas
        // al interior de la ventana capturada.
        if (selectedRegion.X < 0 || selectedRegion.Y < 0 ||
            selectedRegion.Right > width ||
            selectedRegion.Bottom > height)
        {
            throw new InvalidOperationException(
                "El área seleccionada no coincide con la ventana. " +
                "Volvé a seleccionar el importe.");
        }

        using var fullWindow = new Bitmap(width, height);

        using (var graphics = Graphics.FromImage(fullWindow))
        {
            IntPtr hdc = graphics.GetHdc();
            bool success;

            try
            {
                // 2 = PW_RENDERFULLCONTENT
                success = PrintWindow(hwnd, hdc, 2);
            }
            finally
            {
                graphics.ReleaseHdc(hdc);
            }

            if (!success)
                throw new InvalidOperationException(
                    "Windows no pudo capturar servidor.exe directamente.");
        }

        return fullWindow.Clone(
            selectedRegion,
            System.Drawing.Imaging.PixelFormat.Format24bppRgb);
    }


    private void SaveRegion()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
            File.WriteAllText(settingsPath, $"{selectedRegion.X},{selectedRegion.Y},{selectedRegion.Width},{selectedRegion.Height}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No pude guardar el área seleccionada: {ex.Message}", "Configuración", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void LoadSavedRegion()
    {
        try
        {
            if (!File.Exists(settingsPath)) return;
            var parts = File.ReadAllText(settingsPath).Split(',');
            if (parts.Length != 4) return;
            selectedRegion = new Rectangle(
                int.Parse(parts[0], CultureInfo.InvariantCulture),
                int.Parse(parts[1], CultureInfo.InvariantCulture),
                int.Parse(parts[2], CultureInfo.InvariantCulture),
                int.Parse(parts[3], CultureInfo.InvariantCulture));
            if (selectedRegion.Width > 5 && selectedRegion.Height > 5)
            {
                regionLabel.Text = $"Área guardada: X={selectedRegion.X}, Y={selectedRegion.Y}, {selectedRegion.Width}×{selectedRegion.Height}";
                readButton.Enabled = true;
                statusLabel.Text = "Área guardada cargada.";
            }
        }
        catch { /* Una configuración inválida se puede reemplazar seleccionando el área de nuevo. */ }
    }

    private void RunAutomaticCapture(int column, string timeLabel)
    {
        try
        {
            if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
                throw new InvalidOperationException("No hay un área guardada. Abrí ControlFacturacion una vez y seleccioná el número FACTURADO.");

            // Capture and OCR without showing message boxes in unattended mode.
            decimal amount = ReadAmountForAutomation();
            SaveToExcel(amount, column, DateTime.Today);
            WriteAutomationLog($"OK | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | {timeLabel} | {amount.ToString(CultureInfo.InvariantCulture)}");
        }
        catch (Exception ex)
        {
            WriteAutomationLog($"ERROR | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | {timeLabel} | {ex.Message}");
        }
        finally
        {
            Close();
        }
    }

    private decimal ReadAmountForAutomation()
    {
        string dataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        if (!File.Exists(Path.Combine(dataPath, "spa.traineddata")))
            throw new FileNotFoundException("Falta tessdata\\spa.traineddata.", Path.Combine(dataPath, "spa.traineddata"));

        using var bitmap = CaptureServerArea();

        string capturePath = Path.Combine(AppContext.BaseDirectory, "ultima_captura.png");
        bitmap.Save(capturePath, System.Drawing.Imaging.ImageFormat.Png);
        using var engine = new TesseractEngine(dataPath, "spa", EngineMode.LstmOnly);
        engine.SetVariable("tessedit_char_whitelist", "0123456789.,$ ");
        engine.DefaultPageSegMode = PageSegMode.SingleLine;
        using var pix = Pix.LoadFromFile(capturePath);
        using var page = engine.Process(pix);
        string? parsed = ParseAmount(page.GetText().Trim());
        if (page.GetMeanConfidence() < 0.45f || parsed is null)
            throw new InvalidOperationException($"OCR no validado (confianza {page.GetMeanConfidence():P0}). No se guardó ningún importe.");
        if (!decimal.TryParse(parsed, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal amount))
            throw new InvalidOperationException("El OCR no devolvió un importe numérico válido.");
        return amount;
    }

    private static void WriteAutomationLog(string line)
    {
        string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ControlFacturacion");
        Directory.CreateDirectory(folder);
        File.AppendAllText(Path.Combine(folder, "automatizacion.log"), line + Environment.NewLine);
    }


    private void SelectRegion()
    {
        Hide();

        using var overlay = new RegionSelector();

        if (overlay.ShowDialog() == DialogResult.OK &&
            overlay.SelectedRectangle.Width > 5 &&
            overlay.SelectedRectangle.Height > 5)
        {
            IntPtr hwnd = FindServerWindow();

            if (hwnd == IntPtr.Zero ||
                !GetWindowRect(hwnd, out RECT rect))
            {
                MessageBox.Show(
                    "No encontré servidor.exe. Abrilo en Cierres → Preliminar e intentá de nuevo.",
                    "Ventana no encontrada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else
            {
                Rectangle screenArea = overlay.SelectedRectangle;

                // Convertir coordenadas de pantalla a coordenadas
                // relativas a la ventana de servidor.exe.
                Rectangle relativeArea = new Rectangle(
                    screenArea.X - rect.Left,
                    screenArea.Y - rect.Top,
                    screenArea.Width,
                    screenArea.Height);

                int windowWidth = rect.Right - rect.Left;
                int windowHeight = rect.Bottom - rect.Top;

                if (relativeArea.X < 0 ||
                    relativeArea.Y < 0 ||
                    relativeArea.Right > windowWidth ||
                    relativeArea.Bottom > windowHeight)
                {
                    MessageBox.Show(
                        "La selección debe estar completamente dentro de servidor.exe.",
                        "Área fuera de la ventana",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    selectedRegion = relativeArea;

                    regionLabel.Text =
                        $"Área relativa: X={selectedRegion.X}, " +
                        $"Y={selectedRegion.Y}, " +
                        $"{selectedRegion.Width}×{selectedRegion.Height}";

                    readButton.Enabled = true;
                    SaveRegion();

                    statusLabel.Text =
                        "Área de servidor.exe guardada correctamente.";
                }
            }
        }

        Show();
        Activate();
    }


    private void ReadAmount()
    {
        if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
        {
            MessageBox.Show("Primero seleccioná el área que contiene el importe FACTURADO.", "Falta seleccionar área", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string dataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        string languageFile = Path.Combine(dataPath, "spa.traineddata");
        if (!File.Exists(languageFile))
        {
            MessageBox.Show("Falta tessdata\\spa.traineddata. Copiá ese archivo a la carpeta tessdata del proyecto y volvé a ejecutar.", "Falta el modelo de español", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            using var bitmap = CaptureServerArea();

            string capturePath = Path.Combine(AppContext.BaseDirectory, "ultima_captura.png");
            bitmap.Save(capturePath, System.Drawing.Imaging.ImageFormat.Png);

            using var engine = new TesseractEngine(dataPath, "spa", EngineMode.LstmOnly);
            engine.SetVariable("tessedit_char_whitelist", "0123456789.,$ ");
            engine.DefaultPageSegMode = PageSegMode.SingleLine;
            using var pix = Pix.LoadFromFile(capturePath);
            using var page = engine.Process(pix);
            string text = page.GetText().Trim();
            rawText.Text = $"Texto OCR:\r\n{text}\r\n\r\nConfianza: {page.GetMeanConfidence():P0}\r\nCaptura guardada en: {capturePath}";

            string? normalized = ParseAmount(text);
            if (normalized is null)
            {
                lastAmount = null;
                save13Button.Enabled = save16Button.Enabled = false;
                resultLabel.Text = "Importe detectado: no reconocido";
                statusLabel.Text = "No se guardó nada. Volvé a seleccionar solo el número.";
                MessageBox.Show("No pude identificar un importe. Revisá ultima_captura.png y seleccioná solamente el número.", "Lectura no validada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out var value))
            {
                MessageBox.Show("No pude convertir el número reconocido.", "Lectura no validada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            lastAmount = value;
            lastReadDate = DateTime.Today;
            save13Button.Enabled = save16Button.Enabled = true;
            resultLabel.Text = $"Importe detectado: {value.ToString("N2", CultureInfo.GetCultureInfo("es-AR"))}";
            statusLabel.Text = "Lectura lista. Podés probar guardar en Excel o esperar la hora programada.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error durante la lectura OCR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static string? ParseAmount(string input)
    {
        var match = Regex.Match(input, @"(?<!\d)\d{1,3}(?:[,.]\d{3})*(?:[,.]\d{2})?(?!\d)|(?<!\d)\d+(?:[,.]\d{2})(?!\d)");
        if (!match.Success) return null;
        string candidate = match.Value.Replace("$", "").Replace(" ", "");
        if (candidate.Contains(',') && candidate.Contains('.'))
            candidate = candidate.Replace(",", "");
        else if (candidate.Contains(','))
        {
            int decimals = candidate.Length - candidate.LastIndexOf(',') - 1;
            candidate = decimals == 2 ? candidate.Replace(",", ".") : candidate.Replace(",", "");
        }
        if (decimal.TryParse(candidate, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal value))
            return value.ToString(CultureInfo.InvariantCulture);
        return null;
    }

    private void SaveLastAmount(int column, string timeLabel)
    {
        if (lastAmount is null || lastReadDate.Date != DateTime.Today)
        {
            MessageBox.Show("Primero leé el importe de la pantalla. Para evitar guardar un dato viejo, la lectura debe ser de hoy.", "No hay importe válido", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            SaveToExcel(lastAmount.Value, column, DateTime.Today);
            statusLabel.Text = $"Guardado en Siesta.xlsx para hoy, columna {timeLabel}.";
            MessageBox.Show($"Se guardó {lastAmount.Value.ToString("N2", CultureInfo.GetCultureInfo("es-AR"))} en Siesta.xlsx, columna {timeLabel}.", "Guardado correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No pude guardar en Excel:\n{ex.Message}\n\nSi Siesta.xlsx está abierto en Excel, cerralo y volvé a intentar.", "Error al guardar", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void SaveToExcel(decimal amount, int column, DateTime date)
    {
        if (!File.Exists(ExcelPath))
            throw new FileNotFoundException("No encontré el archivo en la ruta configurada.", ExcelPath);

        using var workbook = new XLWorkbook(ExcelPath);
        var sheet = workbook.Worksheets.FirstOrDefault(w => w.Name == "Hoja1")
            ?? throw new InvalidOperationException("No encontré la hoja Hoja1 dentro del Excel.");
        int row = 0;
        for (int r = 1; r <= sheet.LastRowUsed().RowNumber(); r++)
        {
            if (sheet.Cell(r, 2).TryGetValue<DateTime>(out var sheetDate) && sheetDate.Date == date.Date)
            {
                row = r;
                break;
            }
        }
        if (row == 0)
            throw new InvalidOperationException($"No hay una fila para la fecha {date:dd/MM/yyyy} en la columna B de Hoja1.");

        sheet.Cell(row, column).Value = amount;
        sheet.Cell(row, column).Style.NumberFormat.Format = "#,##0.00";
        workbook.Save();
    }

    private void CheckSchedule()
    {
        var now = DateTime.Now;
        if (now.Date != lastReadDate.Date && lastReadDate != default)
        {
            lastAmount = null;
            save13Button.Enabled = save16Button.Enabled = false;
        }
        if (now.Hour == 13 && now.Minute == 0 && !saved13Today)
        {
            saved13Today = true; TryScheduledSave(3, "13:00");
        }
        if (now.Hour == 16 && now.Minute == 0 && !saved16Today)
        {
            saved16Today = true; TryScheduledSave(4, "16:00");
        }
        if (now.Hour == 0 && now.Minute == 0)
        {
            saved13Today = false;
            saved16Today = false;
        }
    }

    private bool TryScheduledSave(int column, string timeLabel)
    {
        try
        {
            if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
                throw new InvalidOperationException("No seleccionaste el área FACTURADO.");
            // At the scheduled time, capture and OCR the current screen directly.
            ReadAmount();
            if (lastAmount is null || lastReadDate.Date != DateTime.Today)
                throw new InvalidOperationException("No se obtuvo una lectura válida para hoy.");
            SaveToExcel(lastAmount.Value, column, DateTime.Today);
            statusLabel.Text = $"Guardado automático de {timeLabel} completado.";
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No se pudo guardar automáticamente a las {timeLabel}:\n{ex.Message}", "Error de guardado automático", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
    }
}

internal sealed class RegionSelector : Form
{
    private Point start;
    private Point current;
    private bool dragging;
    public Rectangle SelectedRectangle { get; private set; }

    public RegionSelector()
    {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        Bounds = SystemInformation.VirtualScreen;
        TopMost = true;
        ShowInTaskbar = false;
        DoubleBuffered = true;
        BackColor = Color.Black;
        Opacity = 0.28;
        Cursor = Cursors.Cross;
        KeyPreview = true;
        KeyDown += (_, e) => { if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); } };
        MouseDown += OnMouseDown;
        MouseMove += OnMouseMove;
        MouseUp += OnMouseUp;
        Paint += OnPaint;
    }

    private void OnMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        start = e.Location; current = e.Location; dragging = true;
    }
    private void OnMouseMove(object? sender, MouseEventArgs e)
    {
        if (!dragging) return;
        current = e.Location; Invalidate();
    }
    private void OnMouseUp(object? sender, MouseEventArgs e)
    {
        if (!dragging) return;
        dragging = false; current = e.Location;
        var local = Rectangle.FromLTRB(Math.Min(start.X, current.X), Math.Min(start.Y, current.Y), Math.Max(start.X, current.X), Math.Max(start.Y, current.Y));
        SelectedRectangle = new Rectangle(local.X + Bounds.X, local.Y + Bounds.Y, local.Width, local.Height);
        DialogResult = SelectedRectangle.Width > 5 && SelectedRectangle.Height > 5 ? DialogResult.OK : DialogResult.Cancel;
        Close();
    }
    private void OnPaint(object? sender, PaintEventArgs e)
    {
        if (dragging)
        {
            var rect = Rectangle.FromLTRB(Math.Min(start.X, current.X), Math.Min(start.Y, current.Y), Math.Max(start.X, current.X), Math.Max(start.Y, current.Y));
            using var pen = new Pen(Color.Lime, 3);
            e.Graphics.DrawRectangle(pen, rect);
        }
        using var brush = new SolidBrush(Color.White);
        e.Graphics.DrawString("Arrastrá para seleccionar solamente el número FACTURADO. ESC para cancelar.", Font, brush, 20, 20);
    }
}