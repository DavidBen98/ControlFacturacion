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
using Excel = Microsoft.Office.Interop.Excel;

namespace ControlFacturacion;
   
public sealed class MainForm : Form
{
    private const string ExcelPath = @"C:\Users\Servidor\Documents\Siestas 2026.xlsx";
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

           

Shown += async (_, _) =>
{
    if (automaticMode)
    {
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
        Hide();
    }

    try
    {
        // Esperar a que Servidor de sucursales esté listo.
        await System.Threading.Tasks.Task.Delay(1800);

        // Actualizar el servidor al iniciar ControlFacturacion.
        await RefreshServerAsync();

        // Si se inició en modo automático, continuar con OCR y Excel.
        if (automaticMode)
        {
            RunAutomaticCaptureAfterRefresh(
                automaticColumn,
                automaticTimeLabel);
        }
        else
        {
            statusLabel.Text =
                "Servidor actualizado. Ya podés leer el importe.";
        }
    }
    catch (Exception ex)
    {
        WriteAutomationLog(
            $"ERROR | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Inicio | {ex.Message}");

        if (automaticMode)
            Close();
        else
            MessageBox.Show(
                $"No se pudo actualizar Servidor de sucursales:\n{ex.Message}",
                "Error de actualización",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
    }
};

 
}


    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

[DllImport("oleaut32.dll", PreserveSig = true)]
private static extern int GetActiveObject(
    ref Guid rclsid,
    IntPtr reserved,
    [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

[DllImport("user32.dll", SetLastError = true)]
private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

private const int SW_RESTORE = 9;

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


[DllImport("user32.dll", SetLastError = true)]
private static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("user32.dll")]
private static extern bool SetCursorPos(int X, int Y);

[DllImport("user32.dll")]
private static extern void mouse_event(
    uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
private const uint MOUSEEVENTF_LEFTUP = 0x0004;



private static async System.Threading.Tasks.Task RefreshServerAsync()
{
    IntPtr hwnd = FindServerWindow();

    if (hwnd == IntPtr.Zero)
        throw new InvalidOperationException(
            "No encontré Servidor de sucursales.");

    // Minimizar Servidor.
ShowWindow(hwnd, 6); // SW_MINIMIZE

await System.Threading.Tasks.Task.Delay(1000);

// Restaurar la ventana.
ShowWindow(hwnd, SW_RESTORE);

await System.Threading.Tasks.Task.Delay(1000);

// Intentar poner Servidor en primer plano.
SetForegroundWindow(hwnd);

await System.Threading.Tasks.Task.Delay(700);

    // Obtener las coordenadas DESPUÉS de restaurarla.
    if (!GetWindowRect(hwnd, out RECT rect))
        throw new InvalidOperationException(
            "No pude obtener las dimensiones de Servidor de sucursales.");

    // A partir de acá mantené tus tres clics actuales.
// Coordenadas relativas del botón Actualizar.
const int relativeX = 512;
const int relativeY = 279;

   // CLICK 1: primera acción previa.
SetCursorPos(rect.Left + 200, rect.Top + 50);
await System.Threading.Tasks.Task.Delay(200);
mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
await System.Threading.Tasks.Task.Delay(100);
mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);

await System.Threading.Tasks.Task.Delay(500);

// CLICK 2: segunda acción previa.
SetCursorPos(rect.Left + 50, rect.Top + 80);
await System.Threading.Tasks.Task.Delay(200);
mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
await System.Threading.Tasks.Task.Delay(100);
mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);

await System.Threading.Tasks.Task.Delay(500);

// CLICK 3: botón Actualizar.
SetCursorPos(rect.Left + relativeX, rect.Top + relativeY);
await System.Threading.Tasks.Task.Delay(200);
mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
await System.Threading.Tasks.Task.Delay(100);
mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
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

   
private async void RunAutomaticCapture(int column, string timeLabel)
{
    try
    {
        if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
            throw new InvalidOperationException(
                "No hay un área guardada. Abrí ControlFacturacion una vez y seleccioná el importe FACTURADO.");

        // 1. Actualizar Servidor de sucursales.
        await RefreshServerAsync();

        // 2. Leer el importe actualizado mediante OCR.
        decimal amount = ReadAmountForAutomation();

        // 3. Guardar el importe en Excel.
        SaveToExcel(amount, column, DateTime.Today);

        WriteAutomationLog(
            $"OK | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | {amount.ToString(CultureInfo.InvariantCulture)}");
    }
    catch (Exception ex)
    {
        WriteAutomationLog(
            $"ERROR | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | {ex.Message}");
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


private void RunAutomaticCaptureAfterRefresh(
    int column, string timeLabel)
{
    try
    {
        if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
            throw new InvalidOperationException(
                "No hay un área OCR guardada.");

        decimal amount = ReadAmountForAutomation();

        SaveToExcel(amount, column, DateTime.Today);

        WriteAutomationLog(
            $"OK | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | " +
            $"{amount.ToString(CultureInfo.InvariantCulture)}");
    }
    catch (Exception ex)
    {
        WriteAutomationLog(
            $"ERROR | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | {ex.Message}");
    }
    finally
    {
        Close();
    }
}

    




private static void SaveToExcel(decimal amount, int column, DateTime date)
{
    if (!File.Exists(ExcelPath))
        throw new FileNotFoundException(
            "No encontré el archivo Excel configurado.",
            ExcelPath);

    object? excelObject = null;
    object? booksObject = null;
    object? workbookObject = null;
    object? sheetsObject = null;
    object? sheetObject = null;
    object? usedRangeObject = null;
    object? targetCellObject = null;

    bool createdExcelInstance = false;
    bool openedWorkbook = false;

    try
    {
        // 1. Intentar conectarse a una instancia existente de Excel.
        Guid excelClsid = new Guid(
            "00024500-0000-0000-C000-000000000046");

        int hr = GetActiveObject(
            ref excelClsid,
            IntPtr.Zero,
            out excelObject!);

        if (hr == 0 && excelObject != null)
        {
            // Excel ya estaba ejecutándose.
            // Esta instancia no debe cerrarse al terminar.
            dynamic app = excelObject;
            booksObject = app.Workbooks;
        }
        else
        {
            // No se encontró una instancia accesible.
            // Crear una instancia propia.
            Type? excelType = Type.GetTypeFromProgID(
                "Excel.Application");

            if (excelType == null)
                throw new InvalidOperationException(
                    "No se encontró Microsoft Excel instalado.");

            excelObject = Activator.CreateInstance(excelType);

            if (excelObject == null)
                throw new InvalidOperationException(
                    "No se pudo iniciar Excel.");

            createdExcelInstance = true;

            dynamic app = excelObject;
            app.Visible = false;
            booksObject = app.Workbooks;
        }

        dynamic books = booksObject!;

        string expectedPath = Path.GetFullPath(ExcelPath);

        // 2. Buscar el archivo en la instancia de Excel elegida.
        int bookCount = (int)books.Count;

        for (int i = 1; i <= bookCount; i++)
        {
            object? candidateObject = null;

            try
            {
                candidateObject = books.Item(i);

                dynamic candidate = candidateObject!;

                string fullName = Path.GetFullPath(
                    (string)candidate.FullName);

                if (string.Equals(
                    fullName,
                    expectedPath,
                    StringComparison.OrdinalIgnoreCase))
                {
                    workbookObject = candidateObject;
                    candidateObject = null;
                    break;
                }
            }
            finally
            {
                // Liberar solamente los libros que no son el objetivo.
                if (candidateObject != null)
                    Marshal.ReleaseComObject(candidateObject);
            }
        }

        // 3. Abrir el archivo únicamente si no estaba abierto
        // en esta instancia de Excel.
        if (workbookObject == null)
        {
            workbookObject = books.Open(ExcelPath);
            openedWorkbook = true;
        }

        dynamic workbook = workbookObject!;

        // 4. Obtener Hoja1 y su rango utilizado.
        sheetsObject = workbook.Worksheets;

        dynamic sheets = sheetsObject!;

        sheetObject = sheets.Item("Hoja1");

        dynamic sheet = sheetObject!;

        usedRangeObject = sheet.UsedRange;

        dynamic usedRange = usedRangeObject!;

        int firstRow = (int)usedRange.Row;
        int firstCol = (int)usedRange.Column;

        int lastRow =
            firstRow + (int)usedRange.Rows.Count - 1;

        int lastCol =
            firstCol + (int)usedRange.Columns.Count - 1;

        int dateRow = 0;
        int dateColumn = 0;

        // 5. Buscar la fecha indicada en toda la hoja.
        for (int r = firstRow;
             r <= lastRow && dateRow == 0;
             r++)
        {
            for (int c = firstCol; c <= lastCol; c++)
            {
                object? value = sheet.Cells[r, c].Value2;

                if (value is not double serial ||
                    serial < 1 ||
                    serial > DateTime.MaxValue.ToOADate())
                {
                    continue;
                }

                DateTime cellDate;

                try
                {
                    cellDate = DateTime.FromOADate(serial);
                }
                catch (ArgumentException)
                {
                    continue;
                }

                if (cellDate.Date == date.Date)
                {
                    dateRow = r;
                    dateColumn = c;
                    break;
                }
            }
        }

        if (dateRow == 0)
        {
            throw new InvalidOperationException(
                $"No se encontró la fecha {date:dd/MM/yyyy} en Hoja1.");
        }

        // 6. Calcular la columna de destino.
        int offset = column switch
        {
            3 => 1, // 13:00
            4 => 2, // 16:00
            _ => throw new InvalidOperationException(
                $"Horario no válido: {column}.")
        };

        targetCellObject = sheet.Cells[
            dateRow, dateColumn + offset];

        dynamic targetCell = targetCellObject!;

        // 7. Escribir el importe y aplicar el formato monetario.
        targetCell.Value2 = (double)amount;

        targetCell.NumberFormat =
            "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";

        // 8. Guardar los cambios.
        workbook.Save();
    }
    finally
    {
        // Liberar las referencias auxiliares.
        if (targetCellObject != null)
            Marshal.ReleaseComObject(targetCellObject);

        if (usedRangeObject != null)
            Marshal.ReleaseComObject(usedRangeObject);

        if (sheetObject != null)
            Marshal.ReleaseComObject(sheetObject);

        if (sheetsObject != null)
            Marshal.ReleaseComObject(sheetsObject);

        // Cerrar solo el libro que abrió este método.
        if (openedWorkbook && workbookObject != null)
        {
            try
            {
                ((dynamic)workbookObject).Close(false);
            }
            finally
            {
                Marshal.ReleaseComObject(workbookObject);
                workbookObject = null;
            }
        }
        else if (workbookObject != null)
        {
            // Si ya estaba abierto, liberar la referencia COM,
            // sin cerrar el libro.
            Marshal.ReleaseComObject(workbookObject);
            workbookObject = null;
        }

        if (booksObject != null)
        {
            Marshal.ReleaseComObject(booksObject);
            booksObject = null;
        }

        // Cerrar Excel solo si esta ejecución creó la instancia.
        if (createdExcelInstance && excelObject != null)
        {
            try
            {
                ((dynamic)excelObject).Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(excelObject);
                excelObject = null;
            }
        }
        else if (excelObject != null)
        {
            // Si Excel ya estaba abierto, no ejecutar Quit().
            Marshal.ReleaseComObject(excelObject);
            excelObject = null;
        }
    }
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
            saved13Today = true;
_ = TryScheduledSave(3, "13:00");
        }
        if (now.Hour == 16 && now.Minute == 0 && !saved16Today)
        {
            saved16Today = true;
_ = TryScheduledSave(4, "16:00");
        }
        if (now.Hour == 0 && now.Minute == 0)
        {
            saved13Today = false;
            saved16Today = false;
        }
    }

    
private async System.Threading.Tasks.Task<bool> TryScheduledSave(
    int column, string timeLabel)
{
    try
    {
        if (selectedRegion.Width <= 0 || selectedRegion.Height <= 0)
            throw new InvalidOperationException(
                "No seleccionaste el área FACTURADO.");

        // 1. Actualizar Servidor de sucursales.
        await RefreshServerAsync();

        // 2. Leer el importe actualizado mediante OCR.
        decimal amount = ReadAmountForAutomation();

        // 3. Guardar el importe en Excel.
        SaveToExcel(amount, column, DateTime.Today);

        lastAmount = amount;
        lastReadDate = DateTime.Today;

        statusLabel.Text =
            $"Guardado automático de {timeLabel} completado.";

        WriteAutomationLog(
            $"OK | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | {amount.ToString(CultureInfo.InvariantCulture)}");

        return true;
    }
    catch (Exception ex)
    {
        WriteAutomationLog(
            $"ERROR | {DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
            $"{timeLabel} | {ex.Message}");

        MessageBox.Show(
            $"No se pudo guardar automáticamente a las {timeLabel}:\n{ex.Message}",
            "Error de guardado automático",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);

        return false;
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
}