using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using ZXing;
using ZXing.Common;
using ZXing.Windows.Compatibility;

namespace BarcodePrinterApp
{
    public partial class MainWindow : Window
    {
        private List<string> _itemsToPrint = new List<string>();
        private int _currentPrintIndex = 0;

        public MainWindow()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            LoadPrinters();
        }

        // Loads installed printers into the ComboBox
        private void LoadPrinters()
        {
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                cmbPrinters.Items.Add(printer);
            }

            // Select default system printer if available
            PrinterSettings settings = new PrinterSettings();
            if (!string.IsNullOrEmpty(settings.PrinterName) && cmbPrinters.Items.Contains(settings.PrinterName))
            {
                cmbPrinters.SelectedItem = settings.PrinterName;
            }
            else if (cmbPrinters.Items.Count > 0)
            {
                cmbPrinters.SelectedIndex = 0;
            }
        }

        private void BtnBrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select Bulk List"
            };

            if (ofd.ShowDialog() == true)
            {
                txtExcelPath.Text = ofd.FileName;
            }
        }

        private List<string> ReadExcelRange(int startRow, int endRow)
        {
            List<string> codes = new List<string>();
            if (string.IsNullOrEmpty(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
            {
                MessageBox.Show("Please select a valid Excel file first.");
                return codes;
            }

            using (var stream = File.Open(txtExcelPath.Text, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                DataTable table = result.Tables[0];

                for (int i = startRow - 1; i < endRow && i < table.Rows.Count; i++)
                {
                    if (table.Rows[i][0] != null)
                    {
                        string code = table.Rows[i][0].ToString();
                        if (!string.IsNullOrWhiteSpace(code))
                            codes.Add(code);
                    }
                }
            }
            return codes;
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtManualCode.Text)) return;
            Bitmap previewBmp = GeneratePreviewBitmap(txtManualCode.Text);
            imgPreview.Source = ConvertToWpfImage(previewBmp);
        }

        private void BtnPrintManual_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtManualCode.Text)) return;
            _itemsToPrint = new List<string> { txtManualCode.Text };
            StartPrintJob();
        }

        private void BtnPrintRange_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(txtRangeStart.Text, out int start) && int.TryParse(txtRangeEnd.Text, out int end))
            {
                _itemsToPrint = ReadExcelRange(start, end);
                if (_itemsToPrint.Count > 0)
                    StartPrintJob();
            }
            else
            {
                MessageBox.Show("Invalid range numbers.");
            }
        }

        private void BtnPrintTest_Click(object sender, RoutedEventArgs e)
        {
            _itemsToPrint = new List<string> { "TEST-1234", "TEST-5678" };
            StartPrintJob();
        }

        // Direct printing method (No Forms Dialog)
        //private void StartPrintJob()
        //{
        //    if (cmbPrinters.SelectedItem == null)
        //    {
        //        MessageBox.Show("Please select a printer from the dropdown.");
        //        return;
        //    }

        //    _currentPrintIndex = 0;
        //    PrintDocument pd = new PrintDocument();

        //    // Apply selected printer from ComboBox
        //    pd.PrinterSettings.PrinterName = cmbPrinters.SelectedItem.ToString();
        //    pd.PrintPage += Pd_PrintPage;

        //    try
        //    {
        //        pd.Print(); // Sends directly to the selected printer
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Print error: " + ex.Message);
        //    }
        //}
        private void StartPrintJob()
        {
            if (cmbPrinters.SelectedItem == null)
            {
                MessageBox.Show("Please select a printer from the dropdown.");
                return;
            }

            _currentPrintIndex = 0;
            PrintDocument pd = new PrintDocument();

            // Apply selected printer from ComboBox
            pd.PrinterSettings.PrinterName = cmbPrinters.SelectedItem.ToString();

            // ---------------------------------------------------------
            // CRITICAL FIX: OVERRIDE WINDOWS PAPER SIZE AND MARGINS
            // PrintDocument expects sizes in hundredths of an inch.
            // 10.6cm = 106mm. (106 / 25.4) * 100 = ~417 hundredths of an inch.
            // ---------------------------------------------------------

            float labelH_mm = float.Parse(txtLabelHeight.Text) * 10f; // e.g., 25mm
            float totalHeight_mm = labelH_mm + 4f; // Label + top/bottom margins (29mm)

            int paperWidthHundredths = (int)(106f / 25.4f * 100f);
            int paperHeightHundredths = (int)(totalHeight_mm / 25.4f * 100f);

            pd.DefaultPageSettings.PaperSize = new PaperSize("CustomStickerRoll", paperWidthHundredths, paperHeightHundredths);

            // Force margins to 0 so Windows doesn't artificially clip the print area
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            // ---------------------------------------------------------

            pd.PrintPage += Pd_PrintPage;

            try
            {
                pd.Print(); // Sends directly to the selected printer
            }
            catch (Exception ex)
            {
                MessageBox.Show("Print error: " + ex.Message);
            }
        }

        private void Pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.PageUnit = GraphicsUnit.Millimeter;

            float labelW = float.Parse(txtLabelWidth.Text) * 10f;
            float margin = 2f;
            float gap = 2f;

            if (_currentPrintIndex < _itemsToPrint.Count)
            {
                DrawSingleLabel(e.Graphics, margin, margin, _itemsToPrint[_currentPrintIndex]);
                _currentPrintIndex++;
            }

            if (_currentPrintIndex < _itemsToPrint.Count)
            {
                float xOffset2 = margin + labelW + gap;
                DrawSingleLabel(e.Graphics, xOffset2, margin, _itemsToPrint[_currentPrintIndex]);
                _currentPrintIndex++;
            }

            e.HasMorePages = _currentPrintIndex < _itemsToPrint.Count;
        }

        //private void DrawSingleLabel(Graphics g, float startX, float startY, string text)
        //{
        //    float labelW = float.Parse(txtLabelWidth.Text) * 10f;
        //    float bcW = float.Parse(txtBcWidth.Text) * 10f;
        //    float bcH = float.Parse(txtBcHeight.Text);
        //    float fontMm = float.Parse(txtFontSize.Text);

        //    BarcodeWriter writer = new BarcodeWriter
        //    {
        //        Format = BarcodeFormat.CODE_128,
        //        Options = new EncodingOptions
        //        {
        //            Width = 300,
        //            Height = (int)(bcH * 10),
        //            Margin = 0,
        //            PureBarcode = true
        //        }
        //    };

        //    using (Bitmap bcBitmap = writer.Write(text))
        //    {
        //        float barcodeX = startX + ((labelW - bcW) / 2f);
        //        float barcodeY = startY + 3f;

        //        g.DrawImage(bcBitmap, barcodeX, barcodeY, bcW, bcH);

        //        using (Font font = new Font("Arial", fontMm, GraphicsUnit.Millimeter))
        //        {
        //            SizeF textSize = g.MeasureString(text, font);
        //            float textX = startX + ((labelW - textSize.Width) / 2f);
        //            float textY = barcodeY + bcH + 1f;

        //            g.DrawString(text, font, System.Drawing.Brushes.Black, textX, textY);
        //        }
        //    }
        //}

        //private Bitmap GeneratePreviewBitmap(string text)
        //{
        //    Bitmap bmp = new Bitmap(400, 200);
        //    bmp.SetResolution(100, 100);
        //    using (Graphics g = Graphics.FromImage(bmp))
        //    {
        //        g.Clear(System.Drawing.Color.White);
        //        g.PageUnit = GraphicsUnit.Millimeter;

        //        float labelW = float.Parse(txtLabelWidth.Text) * 10f;
        //        float labelH = float.Parse(txtLabelHeight.Text) * 10f;
        //        g.DrawRectangle(Pens.LightGray, 1, 1, labelW - 2, labelH - 2);

        //        DrawSingleLabel(g, 0, 0, text);
        //    }
        //    return bmp;
        //}
        //private Bitmap GeneratePreviewBitmap(string text)
        //{
        //    float labelW = float.Parse(txtLabelWidth.Text) * 10f;  // 50mm
        //    float labelH = float.Parse(txtLabelHeight.Text) * 10f; // 25mm
        //    float margin = 2f;                                     // 2mm margin
        //    float gap = 2f;                                        // 2mm center gap

        //    // Total roll width: 2mm left margin + 50mm label + 2mm gap + 50mm label + 2mm right margin = 106mm
        //    float totalWidthMm = margin + labelW + gap + labelW + margin;
        //    float totalHeightMm = margin + labelH + margin;

        //    // Convert mm to pixels for the on-screen bitmap (approx 4 pixels per mm for 100 DPI)
        //    int pixelW = (int)(totalWidthMm * 4);
        //    int pixelH = (int)(totalHeightMm * 4);

        //    Bitmap bmp = new Bitmap(pixelW, pixelH);
        //    bmp.SetResolution(100, 100);

        //    using (Graphics g = Graphics.FromImage(bmp))
        //    {
        //        // 1. Fill the background to simulate the white backing roll
        //        g.Clear(System.Drawing.Color.WhiteSmoke);
        //        g.PageUnit = GraphicsUnit.Millimeter;

        //        // 2. Draw border representing the LEFT sticker edge
        //        g.FillRectangle(System.Drawing.Brushes.White, margin, margin, labelW, labelH);
        //        g.DrawRectangle(Pens.LightGray, margin, margin, labelW, labelH);

        //        // 3. Draw border representing the RIGHT sticker edge
        //        float xOffset2 = margin + labelW + gap;
        //        g.FillRectangle(System.Drawing.Brushes.White, xOffset2, margin, labelW, labelH);
        //        g.DrawRectangle(Pens.LightGray, xOffset2, margin, labelW, labelH);

        //        // 4. Draw the barcode and text on both stickers
        //        DrawSingleLabel(g, margin, margin, text);
        //        DrawSingleLabel(g, xOffset2, margin, text);
        //    }
        //    return bmp;
        //}

        private void DrawSingleLabel(Graphics g, float startX, float startY, string text)
        {
            // Parse dimensions in millimeters
            float labelW = float.Parse(txtLabelWidth.Text) * 10f;  // e.g., 50mm
            float labelH = float.Parse(txtLabelHeight.Text) * 10f; // e.g., 25mm
            float bcW = float.Parse(txtBcWidth.Text) * 10f;        // e.g., 30mm
            float bcH = float.Parse(txtBcHeight.Text);             // e.g., 9mm
            float fontMm = float.Parse(txtFontSize.Text);          // e.g., 4mm

            BarcodeWriter writer = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Width = 300,
                    Height = (int)(bcH * 10),
                    Margin = 0,
                    PureBarcode = true
                }
            };

            using (Bitmap bcBitmap = writer.Write(text))
            {
                using (Font font = new Font("Arial", fontMm, System.Drawing.FontStyle.Bold, GraphicsUnit.Millimeter))
                {
                    // Measure the exact physical size of the text
                    SizeF textSize = g.MeasureString(text, font);

                    // Gap between the barcode and the text below it (1.5mm)
                    float gapBetween = 1.5f;

                    // Calculate total height of all content to center it vertically
                    float totalContentHeight = bcH + gapBetween + textSize.Height;

                    // X and Y coordinates for PERFECT centering
                    float barcodeX = startX + ((labelW - bcW) / 2f);
                    float barcodeY = startY + ((labelH - totalContentHeight) / 2f);

                    // Draw Barcode
                    g.DrawImage(bcBitmap, barcodeX, barcodeY, bcW, bcH);

                    // Draw Text exactly centered below the barcode
                    float textX = startX + ((labelW - textSize.Width) / 2f);
                    float textY = barcodeY + bcH + gapBetween;

                    g.DrawString(text, font, System.Drawing.Brushes.Black, textX, textY);
                }
            }
        }

        private Bitmap GeneratePreviewBitmap(string text)
        {
            float labelW = float.Parse(txtLabelWidth.Text) * 10f;
            float labelH = float.Parse(txtLabelHeight.Text) * 10f;
            float margin = 2f;
            float gap = 2f;

            // Calculate full roll width (10.6cm -> 106mm)
            float totalWidthMm = margin + labelW + gap + labelW + margin;
            float totalHeightMm = margin + labelH + margin;

            // Convert mm to pixels for a high-quality screen preview
            int pixelW = (int)(totalWidthMm * 5);
            int pixelH = (int)(totalHeightMm * 5);

            Bitmap bmp = new Bitmap(pixelW, pixelH);
            bmp.SetResolution(120, 120);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Smoother rendering for preview
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                // Background (Roll paper color)
                g.Clear(System.Drawing.Color.LightGray);
                g.PageUnit = GraphicsUnit.Millimeter;

                // LEFT STICKER: Draw white background and faint cut-line
                g.FillRectangle(System.Drawing.Brushes.White, margin, margin, labelW, labelH);
                g.DrawRectangle(Pens.Gray, margin, margin, labelW, labelH);

                // RIGHT STICKER: Draw white background and faint cut-line
                float xOffset2 = margin + labelW + gap;
                g.FillRectangle(System.Drawing.Brushes.White, xOffset2, margin, labelW, labelH);
                g.DrawRectangle(Pens.Gray, xOffset2, margin, labelW, labelH);

                // Draw the barcode/text perfectly centered on both stickers
                DrawSingleLabel(g, margin, margin, text);
                DrawSingleLabel(g, xOffset2, margin, text);
            }

            return bmp;
        }

        private BitmapImage ConvertToWpfImage(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }
    }
}