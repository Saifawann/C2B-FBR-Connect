using C2B_FBR_Connect.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.IO.Image;
using QRCoder;

namespace C2B_FBR_Connect.Services
{
    public class PDFService
    {
        private readonly PdfFont _boldFont;
        private readonly PdfFont _normalFont;

        public PDFService()
        {
            // Initialize fonts once
            _boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            _normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
        }

        public void GenerateInvoicePDF(Invoice invoice, FBRInvoicePayload details, string outputPath)
        {
            if (invoice == null)
                throw new ArgumentNullException(nameof(invoice));
            if (details == null)
                throw new ArgumentNullException(nameof(details));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be empty", nameof(outputPath));

            // Ensure directory exists
            var directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            try
            {
                using (var writer = new PdfWriter(outputPath))
                using (var pdf = new PdfDocument(writer))
                using (var document = new Document(pdf))
                {
                    // Set margins
                    document.SetMargins(36, 36, 36, 36);

                    AddTitle(document);
                    AddCompanyInfo(document, invoice, details);
                    AddItemsTable(document, details);
                    AddTotals(document, details);
                    AddFBRInfo(document, invoice);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to generate PDF: {ex.Message}", ex);
            }
        }

        private void AddTitle(Document document)
        {
            var title = new Paragraph("TAX INVOICE")
                .SetFont(_boldFont)
                .SetFontSize(18)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetMarginBottom(20);

            document.Add(title);
        }

        private void AddCompanyInfo(Document document, Invoice invoice, FBRInvoicePayload details)
        {
            document.Add(new Paragraph($"Company: {invoice.CompanyName ?? "N/A"}")
                .SetFont(_boldFont)
                .SetFontSize(10));

            document.Add(new Paragraph($"Invoice No: {invoice.InvoiceNumber ?? "N/A"}")
                .SetFont(_normalFont)
                .SetFontSize(10));

            document.Add(new Paragraph($"Date: {details.InvoiceDate:dd-MMM-yyyy}")
                .SetFont(_normalFont)
                .SetFontSize(10));

            document.Add(new Paragraph($"Customer: {invoice.CustomerName ?? "N/A"}")
                .SetFont(_normalFont)
                .SetFontSize(10)
                .SetMarginBottom(15));
        }

        private void AddItemsTable(Document document, FBRInvoicePayload details)
        {
            if (details.Items == null || details.Items.Count == 0)
            {
                document.Add(new Paragraph("No items found")
                    .SetFont(_normalFont)
                    .SetFontSize(10));
                return;
            }

            Table table = new Table(UnitValue.CreatePercentArray(new float[] { 3f, 1f, 1.5f, 1.5f, 1.5f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(15);

            // Add headers
            AddTableHeader(table, "Description");
            AddTableHeader(table, "Qty");
            AddTableHeader(table, "Rate");
            AddTableHeader(table, "Tax");
            AddTableHeader(table, "Amount");

            // Add items
            foreach (var item in details.Items)
            {
                AddTableCell(table, item.ItemName ?? "N/A");
                AddTableCell(table, item.Quantity.ToString("N0"));
                AddTableCell(table, $"Rs. {item.UnitPrice:N2}");
                AddTableCell(table, $"{item.TaxRate:N2}%");
                AddTableCell(table, $"Rs. {item.TotalPrice:N2}");
            }

            document.Add(table);
        }

        private void AddTableHeader(Table table, string text)
        {
            var cell = new Cell()
                .Add(new Paragraph(text).SetFont(_boldFont).SetFontSize(10))
                .SetBackgroundColor(new DeviceRgb(240, 240, 240))
                .SetPadding(5)
                .SetTextAlignment(TextAlignment.CENTER);

            table.AddHeaderCell(cell);
        }

        private void AddTableCell(Table table, string text)
        {
            var cell = new Cell()
                .Add(new Paragraph(text).SetFont(_normalFont).SetFontSize(10))
                .SetPadding(5);

            table.AddCell(cell);
        }

        private void AddTotals(Document document, FBRInvoicePayload details)
        {
            var totalWithTax = details.TotalAmount + details.TaxAmount;

            // Create right-aligned totals table
            Table totalsTable = new Table(UnitValue.CreatePercentArray(new float[] { 1f, 1f }))
                .SetWidth(UnitValue.CreatePercentValue(40))
                .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.RIGHT)
                .SetMarginBottom(15);

            // Subtotal
            totalsTable.AddCell(CreateTotalCell("Subtotal:", false));
            totalsTable.AddCell(CreateTotalCell($"Rs. {details.TotalAmount:N2}", false));

            // Tax
            totalsTable.AddCell(CreateTotalCell("Tax:", false));
            totalsTable.AddCell(CreateTotalCell($"Rs. {details.TaxAmount:N2}", false));

            // Total
            totalsTable.AddCell(CreateTotalCell("Total:", true));
            totalsTable.AddCell(CreateTotalCell($"Rs. {totalWithTax:N2}", true));

            document.Add(totalsTable);
        }

        private Cell CreateTotalCell(string text, bool isBold)
        {
            return new Cell()
                .Add(new Paragraph(text)
                    .SetFont(isBold ? _boldFont : _normalFont)
                    .SetFontSize(10))
                .SetBorder(iText.Layout.Borders.Border.NO_BORDER)
                .SetPadding(3)
                .SetTextAlignment(TextAlignment.RIGHT);
        }

        private void AddFBRInfo(Document document, Invoice invoice)
        {
            document.Add(new Paragraph("FBR Digital Invoice")
                .SetFont(_boldFont)
                .SetFontSize(10)
                .SetMarginTop(10));

            document.Add(new Paragraph($"IRN: {invoice.FBR_IRN ?? "N/A"}")
                .SetFont(_normalFont)
                .SetFontSize(10));

            // Generate and add QR Code
            if (!string.IsNullOrEmpty(invoice.FBR_QRCode))
            {
                try
                {
                    var qrImageData = GenerateQRCodeImage(invoice.FBR_QRCode);
                    if (qrImageData != null)
                    {
                        var qrImage = new iText.Layout.Element.Image(qrImageData)
                            .SetWidth(150)
                            .SetHeight(150)
                            .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER)
                            .SetMarginTop(10);

                        document.Add(qrImage);
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue - QR code is not critical
                    Console.WriteLine($"Failed to generate QR code: {ex.Message}");
                    document.Add(new Paragraph("QR Code unavailable")
                        .SetFont(_normalFont)
                        .SetFontSize(10));
                }
            }
        }

        private ImageData GenerateQRCodeImage(string qrContent)
        {
            using (var qrGenerator = new QRCodeGenerator())
            {
                using (var qrCodeData = qrGenerator.CreateQrCode(qrContent, QRCodeGenerator.ECCLevel.Q))
                {
                    using (var qrCode = new QRCode(qrCodeData))
                    {
                        using (var qrBitmap = qrCode.GetGraphic(5))
                        {
                            using (var ms = new MemoryStream())
                            {
                                qrBitmap.Save(ms, ImageFormat.Png);
                                return ImageDataFactory.Create(ms.ToArray());
                            }
                        }
                    }
                }
            }
        }
    }
}