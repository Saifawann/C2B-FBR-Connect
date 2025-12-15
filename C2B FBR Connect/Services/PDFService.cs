using C2B_FBR_Connect.Models;
using C2B_FBR_Connect.Utilities;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace C2B_FBR_Connect.Services
{
    public class PDFService
    {
        // Simplified color scheme
        private static readonly DeviceRgb PRIMARY_COLOR = new DeviceRgb(41, 128, 185);
        private static readonly DeviceRgb TABLE_HEADER = new DeviceRgb(52, 152, 219);
        private static readonly DeviceRgb TABLE_ALT_ROW = new DeviceRgb(250, 250, 250);
        private static readonly DeviceRgb BORDER_COLOR = new DeviceRgb(189, 195, 199);
        private static readonly DeviceRgb TEXT_PRIMARY = new DeviceRgb(44, 62, 80);
        private static readonly DeviceRgb TEXT_SECONDARY = new DeviceRgb(127, 140, 141);
        private static readonly DeviceRgb ACCENT_GREEN = new DeviceRgb(39, 174, 96);
        private static readonly DeviceRgb LIGHT_BG = new DeviceRgb(236, 240, 241);

        public void GenerateInvoicePDF(Invoice invoice, Company company, string outputPath)
        {
            ValidateInputs(invoice, company, outputPath);
            EnsureDirectoryExists(outputPath);

            PdfWriter writer = null;
            PdfDocument pdfDoc = null;
            Document document = null;

            try
            {
                writer = new PdfWriter(outputPath);
                pdfDoc = new PdfDocument(writer);
                pdfDoc.SetDefaultPageSize(PageSize.A4);

                // Minimal margins for maximum space
                document = new Document(pdfDoc);
                document.SetMargins(15, 15, 100, 15);  // Bottom margin for FBR section

                // Create fonts
                var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                // Add watermark if needed (you can add Status property to Invoice model)
                // AddWatermark(pdfDoc, "DRAFT"); // Uncomment if you want watermark

                BuildPDFContent(document, invoice, company, boldFont, normalFont);
            }
            finally
            {
                document?.Close();
                pdfDoc?.Close();
                writer?.Close();
            }
        }

        private void ValidateInputs(Invoice invoice, Company company, string outputPath)
        {
            if (invoice == null) throw new ArgumentNullException(nameof(invoice));
            if (company == null) throw new ArgumentNullException(nameof(company));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path cannot be empty");
        }

        private void EnsureDirectoryExists(string path)
        {
            var dir = System.IO.Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        private void AddWatermark(PdfDocument pdfDoc, string text)
        {
            // This will be applied to all pages
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                PdfCanvas canvas = new PdfCanvas(page);
                canvas.SaveState();

                // Set transparency
                var gs = new iText.Kernel.Pdf.Extgstate.PdfExtGState();
                gs.SetFillOpacity(0.3f);
                canvas.SetExtGState(gs);

                canvas.SetFillColor(new DeviceRgb(200, 200, 200));
                canvas.BeginText();
                canvas.SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD), 60);
                canvas.MoveText(page.GetPageSize().GetWidth() / 2 - 100, page.GetPageSize().GetHeight() / 2);
                canvas.ShowText(text);
                canvas.EndText();
                canvas.RestoreState();
            }
        }

        private void BuildPDFContent(Document doc, Invoice invoice, Company company,
            PdfFont boldFont, PdfFont normalFont)
        {
            AddCompactHeader(doc, company, invoice, boldFont, normalFont);
            AddPartyDetailsCompact(doc, invoice, company, boldFont, normalFont);
            AddItemsTableCompact(doc, invoice, boldFont, normalFont);
            AddCompactSummary(doc, invoice, boldFont, normalFont);
            AddFBRSectionAboveFooter(doc, invoice, boldFont, normalFont);
            AddCompactFooter(doc, normalFont);
        }

        #region Compact Header

        private void AddCompactHeader(Document doc, Company company, Invoice invoice,
            PdfFont boldFont, PdfFont normalFont)
        {
            Table headerTable = new Table(UnitValue.CreatePercentArray(new float[] { 1f, 2f, 1f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(8);

            // LEFT: Logo
            Cell logoCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPadding(0);

            if (company.LogoImage != null && company.LogoImage.Length > 0)
            {
                try
                {
                    var logoData = ImageDataFactory.Create(company.LogoImage);
                    var logo = new iText.Layout.Element.Image(logoData);

                    float originalWidth = logo.GetImageWidth();
                    float originalHeight = logo.GetImageHeight();
                    float maxWidth = 70f;
                    float maxHeight = 40f;

                    float widthRatio = maxWidth / originalWidth;
                    float heightRatio = maxHeight / originalHeight;
                    float scaleFactor = Math.Min(widthRatio, heightRatio);

                    logo.SetWidth(originalWidth * scaleFactor);
                    logo.SetHeight(originalHeight * scaleFactor);
                    logoCell.Add(logo);
                }
                catch
                {
                    string initials = "CO";
                    if (!string.IsNullOrEmpty(company.CompanyName) && company.CompanyName.Length > 0)
                    {
                        initials = company.CompanyName.Substring(0, Math.Min(2, company.CompanyName.Length)).ToUpper();
                    }

                    logoCell.Add(new Paragraph(initials)
                        .SetFont(boldFont)
                        .SetFontSize(14)
                        .SetFontColor(PRIMARY_COLOR));
                }
            }

            headerTable.AddCell(logoCell);

            // CENTER: Company Info
            Cell companyCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetPadding(0);

            companyCell.Add(new Paragraph(company.CompanyName ?? "Company Name")
                .SetFont(boldFont)
                .SetFontSize(12)
                .SetFontColor(PRIMARY_COLOR)
                .SetMarginBottom(2));

            if (!string.IsNullOrEmpty(company.SellerAddress))
            {
                companyCell.Add(new Paragraph(company.SellerAddress)
                    .SetFont(normalFont)
                    .SetFontSize(7)
                    .SetFontColor(TEXT_SECONDARY)
                    .SetMarginBottom(1));
            }

            var contactLine = "";
            if (!string.IsNullOrEmpty(company.SellerPhone))
                contactLine = $"Tel: {company.SellerPhone}";
            if (!string.IsNullOrEmpty(company.SellerEmail))
                contactLine += (string.IsNullOrEmpty(contactLine) ? "" : " | ") + company.SellerEmail;

            if (!string.IsNullOrEmpty(contactLine))
            {
                companyCell.Add(new Paragraph(contactLine)
                    .SetFont(normalFont)
                    .SetFontSize(7)
                    .SetFontColor(TEXT_SECONDARY));
            }

            headerTable.AddCell(companyCell);

            // RIGHT: NTN Box
            Cell ntnCell = new Cell()
                .SetBorder(new SolidBorder(PRIMARY_COLOR, 1))
                .SetBackgroundColor(LIGHT_BG)
                .SetPadding(5)
                .SetTextAlignment(TextAlignment.CENTER);

            ntnCell.Add(new Paragraph("NTN")
                .SetFont(normalFont)
                .SetFontSize(7)
                .SetFontColor(TEXT_SECONDARY)
                .SetMarginBottom(2));

            ntnCell.Add(new Paragraph(company.SellerNTN ?? "N/A")
                .SetFont(boldFont)
                .SetFontSize(9)
                .SetFontColor(PRIMARY_COLOR));

            headerTable.AddCell(ntnCell);

            doc.Add(headerTable);

            // Title Bar
            Table titleBar = new Table(1).UseAllAvailableWidth().SetMarginBottom(5);
            Cell titleCell = new Cell()
                .SetBackgroundColor(PRIMARY_COLOR)
                .SetBorder(Border.NO_BORDER)
                .SetPadding(5)
                .SetTextAlignment(TextAlignment.CENTER);

            titleCell.Add(new Paragraph("SALES TAX INVOICE")
                .SetFont(boldFont)
                .SetFontSize(11)
                .SetFontColor(ColorConstants.WHITE));

            titleBar.AddCell(titleCell);
            doc.Add(titleBar);
        }

        #endregion

        #region Compact Party Details

        private void AddPartyDetailsCompact(Document doc, Invoice invoice, Company company,
            PdfFont boldFont, PdfFont normalFont)
        {
            // ✅ 2-Column Layout: Customer Details (LEFT) | Invoice Details (RIGHT)
            Table partyTable = new Table(UnitValue.CreatePercentArray(new float[] { 1.5f, 1f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(8);

            // LEFT: Customer Details (with all information)
            Cell customerCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetPadding(8)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            // Customer Name
            customerCell.Add(new Paragraph($"To: M/s {invoice.CustomerName ?? "N/A"}")
                .SetFont(boldFont)
                .SetFontSize(8)
                .SetFontColor(TEXT_PRIMARY)
                .SetMarginBottom(4));

            // Address
            if (!string.IsNullOrEmpty(invoice.CustomerAddress))
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("📍 ").SetFontSize(7))
                    .Add(new Text(invoice.CustomerAddress)
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }
            else
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("📍 ").SetFontSize(7))
                    .Add(new Text("N/A")
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }

            // Phone
            if (!string.IsNullOrEmpty(invoice.CustomerPhone))
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("📞 ").SetFontSize(7))
                    .Add(new Text(invoice.CustomerPhone)
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }
            else
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("📞 ").SetFontSize(7))
                    .Add(new Text("N/A")
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }

            // Email
            if (!string.IsNullOrEmpty(invoice.CustomerEmail))
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("✉ ").SetFontSize(7))
                    .Add(new Text(invoice.CustomerEmail)
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }
            else
            {
                customerCell.Add(new Paragraph()
                    .Add(new Text("✉ ").SetFontSize(7))
                    .Add(new Text("N/A")
                        .SetFont(normalFont)
                        .SetFontSize(7)
                        .SetFontColor(TEXT_SECONDARY))
                    .SetMarginBottom(3));
            }

            // NTN
            customerCell.Add(new Paragraph()
                .Add(new Text("NTN: ")
                    .SetFont(normalFont)
                    .SetFontSize(7)
                    .SetFontColor(TEXT_SECONDARY))
                .Add(new Text(invoice.CustomerNTN ?? "N/A")
                    .SetFont(boldFont)
                    .SetFontSize(8)
                    .SetFontColor(PRIMARY_COLOR)));

            partyTable.AddCell(customerCell);

            // RIGHT: Invoice Details
            Cell invoiceCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetPadding(8)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            invoiceCell.Add(new Paragraph($"Invoice #: {invoice.InvoiceNumber ?? "N/A"}")
                .SetFont(boldFont)
                .SetFontSize(8)
                .SetFontColor(TEXT_PRIMARY)
                .SetMarginBottom(4));

            invoiceCell.Add(new Paragraph($"Date: {(invoice.InvoiceDate != DateTime.MinValue ? invoice.InvoiceDate.ToString("dd-MM-yyyy") : "N/A")}")
                .SetFont(normalFont)
                .SetFontSize(8)
                .SetFontColor(TEXT_SECONDARY));

            partyTable.AddCell(invoiceCell);

            doc.Add(partyTable);
        }

        #endregion

        #region Compact Items Table

        private void AddItemsTableCompact(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            if (invoice.Items == null || invoice.Items.Count == 0)
            {
                doc.Add(new Paragraph("No items")
                    .SetFont(normalFont)
                    .SetFontSize(8)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetMarginTop(10)
                    .SetMarginBottom(10));
                return;
            }

            // Ultra-compact table for 20+ items
            Table itemsTable = new Table(UnitValue.CreatePercentArray(
                new float[] { 0.4f, 2.2f, 1f, 0.6f, 0.9f, 0.9f, 0.9f, 1f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(5);

            // Headers - very compact
            string[] headers = { "#", "Description", "HS Code", "Qty", "Rate", "Amount", "Disc", "Net" };

            foreach (var header in headers)
            {
                Cell headerCell = new Cell()
                    .Add(new Paragraph(header)
                        .SetFont(boldFont)
                        .SetFontSize(7)
                        .SetFontColor(ColorConstants.WHITE))
                    .SetBackgroundColor(TABLE_HEADER)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                    .SetPadding(3)
                    .SetBorder(new SolidBorder(ColorConstants.WHITE, 0.5f));

                itemsTable.AddHeaderCell(headerCell);
            }

            // Items - ultra compact rows
            int index = 1;
            foreach (var item in invoice.Items)
            {
                decimal grossAmount = item.TotalPrice > 0 ? item.TotalPrice : (item.Quantity * item.UnitPrice);
                decimal netAmount = Math.Max(0, grossAmount - item.Discount);
                decimal unitRate = item.Quantity > 0 ? grossAmount / item.Quantity : item.UnitPrice;

                var rowColor = index % 2 == 0 ? TABLE_ALT_ROW : ColorConstants.WHITE;

                // Ultra-compact cells with 6pt font and 2pt padding
                AddCompactCell(itemsTable, index.ToString(), TextAlignment.CENTER, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, item.ItemName ?? "Item", TextAlignment.LEFT, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, item.HSCode ?? "-", TextAlignment.CENTER, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, item.Quantity.ToString("N0"), TextAlignment.CENTER, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, unitRate.ToString("N2"), TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, grossAmount.ToString("N2"), TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, item.Discount.ToString("N2"), TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddCompactCell(itemsTable, netAmount.ToString("N2"), TextAlignment.RIGHT, boldFont, 6, rowColor);

                index++;
            }

            doc.Add(itemsTable);
        }

        private void AddCompactCell(Table table, string text, TextAlignment align, PdfFont font,
            float fontSize, iText.Kernel.Colors.Color bgColor)
        {
            Cell cell = new Cell()
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(TEXT_PRIMARY))
                .SetTextAlignment(align)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPadding(2)
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.3f));

            table.AddCell(cell);
        }

        #endregion

        #region Compact Summary

        private void AddCompactSummary(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            // Calculate totals
            decimal grossAmount = 0;
            decimal discount = 0;
            decimal salesTax = 0;

            if (invoice.Items != null)
            {
                foreach (var item in invoice.Items)
                {
                    decimal itemGross = item.TotalPrice > 0 ? item.TotalPrice : (item.Quantity * item.UnitPrice);
                    grossAmount += itemGross;
                    discount += item.Discount;
                    salesTax += item.SalesTaxAmount + item.FurtherTax;
                }
            }

            decimal subtotal = grossAmount - discount;
            decimal grandTotal = subtotal + salesTax;

            // Amount in words - ultra compact (full width)
            string amountInWords = NumberToWords((long)Math.Round(grandTotal));

            doc.Add(new Paragraph($"Amount in Words: Rupees {amountInWords} only")
                .SetFont(normalFont)
                .SetFontSize(7)
                .SetFontColor(TEXT_SECONDARY)
                .SetBackgroundColor(new DeviceRgb(255, 251, 230))
                .SetPadding(4)
                .SetMarginBottom(5));

            // Split layout: Empty left space, Totals on right
            Table mainTable = new Table(UnitValue.CreatePercentArray(new float[] { 1f, 1.2f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(10);

            // LEFT: Empty space
            Cell leftCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(0);

            leftCell.Add(new Paragraph("")
                .SetFont(normalFont)
                .SetFontSize(7)
                .SetFontColor(TEXT_SECONDARY));

            mainTable.AddCell(leftCell);

            // RIGHT: Financial Summary
            Cell totalsCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                .SetPadding(6);

            Table totalsTable = new Table(UnitValue.CreatePercentArray(new float[] { 1.5f, 1f }))
                .UseAllAvailableWidth();

            AddCompactSummaryRow(totalsTable, "Gross Amount:", grossAmount.ToString("N2"), normalFont, false);
            AddCompactSummaryRow(totalsTable, "Discount:", discount.ToString("N2"), normalFont, false);
            AddCompactSummaryRow(totalsTable, "Subtotal:", subtotal.ToString("N2"), normalFont, false);
            AddCompactSummaryRow(totalsTable, "Sales Tax:", salesTax.ToString("N2"), normalFont, false);
            AddCompactSummaryRow(totalsTable, "GRAND TOTAL:", grandTotal.ToString("N2"), boldFont, true);

            totalsCell.Add(totalsTable);
            mainTable.AddCell(totalsCell);

            doc.Add(mainTable);
        }

        private void AddCompactSummaryRow(Table table, string label, string value, PdfFont font, bool isTotal)
        {
            var bgColor = isTotal ? ACCENT_GREEN : ColorConstants.WHITE;
            var textColor = isTotal ? ColorConstants.WHITE : TEXT_PRIMARY;
            var fontSize = isTotal ? 8f : 7f;

            Cell labelCell = new Cell()
                .Add(new Paragraph(label)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(textColor))
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.3f))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(3);

            Cell valueCell = new Cell()
                .Add(new Paragraph(value)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(textColor))
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.3f))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetPadding(3);

            table.AddCell(labelCell);
            table.AddCell(valueCell);
        }

        #endregion

        #region FBR Section Above Footer

        private void AddFBRSectionAboveFooter(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            PdfDocument pdfDoc = doc.GetPdfDocument();
            int numberOfPages = pdfDoc.GetNumberOfPages();

            for (int i = 1; i <= numberOfPages; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                iText.Kernel.Geom.Rectangle pageSize = page.GetPageSize();
                PdfCanvas canvas = new PdfCanvas(page);

                // ✅ Increased dimensions for larger FBR logo and QR code
                float fbrLogoSize = 60f;  // Increased from 30 to 60
                float qrCodeSize = 72f;   // 1.0 inch (as per specification)
                float maxElementSize = Math.Max(fbrLogoSize, qrCodeSize);  // 72

                // ✅ Calculate border height based on largest element + padding
                float borderPadding = 10f;
                float borderHeight = maxElementSize + (borderPadding * 2);  // 72 + 20 = 92

                float fbrY = 60;
                float leftX = pageSize.GetLeft() + 15;
                float rightX = pageSize.GetRight() - 15;
                float borderWidth = pageSize.GetWidth() - 30;

                // ✅ Draw FBR section border (expanded)
                canvas.SaveState();
                canvas.SetStrokeColor(PRIMARY_COLOR);
                canvas.SetLineWidth(1.5f);
                canvas.Rectangle(leftX, fbrY, borderWidth, borderHeight);
                canvas.Stroke();
                canvas.RestoreState();

                // ✅ LEFT: FBR Logo (60x60 - doubled in size)
                string fbrLogoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "fbr_logo.png");
                if (File.Exists(fbrLogoPath))
                {
                    try
                    {
                        var fbrLogoData = ImageDataFactory.Create(fbrLogoPath);
                        var fbrLogoImage = new iText.Layout.Element.Image(fbrLogoData);

                        // Get original dimensions and scale proportionally
                        float originalWidth = fbrLogoImage.GetImageWidth();
                        float originalHeight = fbrLogoImage.GetImageHeight();
                        float maxLogoSize = fbrLogoSize;

                        float widthRatio = maxLogoSize / originalWidth;
                        float heightRatio = maxLogoSize / originalHeight;
                        float logoScaleFactor = Math.Min(widthRatio, heightRatio);

                        float scaledLogoWidth = originalWidth * logoScaleFactor;
                        float scaledLogoHeight = originalHeight * logoScaleFactor;

                        var fbrLogo = new iText.Layout.Element.Image(fbrLogoData)
                            .SetWidth(scaledLogoWidth)
                            .SetHeight(scaledLogoHeight);

                        // Center vertically within border
                        float logoYPosition = fbrY + (borderHeight - scaledLogoHeight) / 2;
                        fbrLogo.SetFixedPosition(i, leftX + borderPadding, logoYPosition);
                        doc.Add(fbrLogo);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading FBR logo: {ex.Message}");

                        // Fallback: Draw "FBR" text
                        canvas.SaveState();
                        canvas.BeginText();
                        canvas.SetFontAndSize(boldFont, 16);
                        canvas.SetFillColor(PRIMARY_COLOR);
                        canvas.MoveText(leftX + borderPadding + 10, fbrY + borderHeight / 2);
                        canvas.ShowText("FBR");
                        canvas.EndText();
                        canvas.RestoreState();
                    }
                }
                else
                {
                    // Fallback: Draw "FBR" text if logo not found
                    canvas.SaveState();
                    canvas.BeginText();
                    canvas.SetFontAndSize(boldFont, 16);
                    canvas.SetFillColor(PRIMARY_COLOR);
                    canvas.MoveText(leftX + borderPadding + 10, fbrY + borderHeight / 2);
                    canvas.ShowText("FBR");
                    canvas.EndText();
                    canvas.RestoreState();
                }

                // ✅ CENTER: FBR Invoice text and number (adjusted for new height)
                float centerTextY = fbrY + (borderHeight / 2) + 10;

                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(boldFont, 9);
                canvas.SetFillColor(PRIMARY_COLOR);
                canvas.MoveText(leftX + fbrLogoSize + borderPadding + 20, centerTextY);
                canvas.ShowText("FBR Invoice #");
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(normalFont, 8);
                canvas.SetFillColor(TEXT_SECONDARY);
                string fbrIRN = invoice.FBR_IRN ?? "PENDING";
                canvas.MoveText(leftX + fbrLogoSize + borderPadding + 20, centerTextY - 15);
                canvas.ShowText(fbrIRN);
                canvas.EndText();
                canvas.RestoreState();

                // ✅ RIGHT: QR Code - Version 2.0 (1.0 x 1.0 inch = 72x72 points)
                if (!string.IsNullOrEmpty(invoice.FBR_IRN))
                {
                    try
                    {
                        var qrData = GenerateQRCodeVersion2(invoice.FBR_IRN);
                        var qrImage = new iText.Layout.Element.Image(qrData)
                            .SetWidth(qrCodeSize)
                            .SetHeight(qrCodeSize);

                        // Center QR code vertically within border
                        float qrYPosition = fbrY + (borderHeight - qrCodeSize) / 2;
                        qrImage.SetFixedPosition(i, rightX - qrCodeSize - borderPadding, qrYPosition);
                        doc.Add(qrImage);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error generating QR code: {ex.Message}");

                        // Fallback: Draw placeholder
                        canvas.SaveState();
                        canvas.SetStrokeColor(BORDER_COLOR);
                        canvas.SetLineWidth(1);
                        float qrYPosition = fbrY + (borderHeight - qrCodeSize) / 2;
                        canvas.Rectangle(rightX - qrCodeSize - borderPadding, qrYPosition, qrCodeSize, qrCodeSize);
                        canvas.Stroke();

                        canvas.BeginText();
                        canvas.SetFontAndSize(normalFont, 8);
                        canvas.SetFillColor(TEXT_SECONDARY);
                        canvas.MoveText(rightX - qrCodeSize - borderPadding + 20, qrYPosition + qrCodeSize / 2);
                        canvas.ShowText("QR Pending");
                        canvas.EndText();
                        canvas.RestoreState();
                    }
                }
                else
                {
                    // Draw "QR Pending" placeholder
                    canvas.SaveState();
                    canvas.SetStrokeColor(BORDER_COLOR);
                    canvas.SetLineWidth(1);
                    float qrYPosition = fbrY + (borderHeight - qrCodeSize) / 2;
                    canvas.Rectangle(rightX - qrCodeSize - borderPadding, qrYPosition, qrCodeSize, qrCodeSize);
                    canvas.Stroke();

                    canvas.BeginText();
                    canvas.SetFontAndSize(normalFont, 8);
                    canvas.SetFillColor(TEXT_SECONDARY);
                    canvas.MoveText(rightX - qrCodeSize - borderPadding + 20, qrYPosition + qrCodeSize / 2);
                    canvas.ShowText("QR Pending");
                    canvas.EndText();
                    canvas.RestoreState();
                }
            }
        }

        #endregion

        #region QR Code Generation - Version 2.0

        /// <summary>
        /// Generates QR Code Version 2.0 (25×25 modules) at 1.0 x 1.0 inch size
        /// </summary>
        private ImageData GenerateQRCodeVersion2(string content)
        {
            using var qrGenerator = new QRCodeGenerator();

            // QR Code Version 2.0 (25×25 modules) with Error Correction Level M
            using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
            using var qrCode = new QRCode(qrCodeData);

            // Calculate pixels per module for 1 inch at 300 DPI
            // 1 inch = 72 points in PDF
            // At 300 DPI: 1 inch = 300 pixels
            // Version 2.0 has 25 modules
            // Pixels per module = 300 / 25 = 12 pixels per module
            int pixelsPerModule = 12;

            using var bitmap = qrCode.GetGraphic(
                pixelsPerModule: pixelsPerModule,
                darkColor: System.Drawing.Color.Black,
                lightColor: System.Drawing.Color.White,
                drawQuietZones: true);

            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            return ImageDataFactory.Create(ms.ToArray());
        }

        #endregion

        #region Compact Footer

        private void AddCompactFooter(Document doc, PdfFont normalFont)
        {
            PdfDocument pdfDoc = doc.GetPdfDocument();
            int numberOfPages = pdfDoc.GetNumberOfPages();

            for (int i = 1; i <= numberOfPages; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                iText.Kernel.Geom.Rectangle pageSize = page.GetPageSize();
                PdfCanvas canvas = new PdfCanvas(page);

                float footerY = 20;
                float centerX = pageSize.GetWidth() / 2;

                // Footer background
                canvas.SaveState();
                canvas.SetFillColor(PRIMARY_COLOR);
                canvas.Rectangle(pageSize.GetLeft(), footerY - 5, pageSize.GetWidth(), 25);
                canvas.Fill();
                canvas.RestoreState();

                // Footer text
                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(normalFont, 6);
                canvas.SetFillColor(ColorConstants.WHITE);

                string footerText = $"Powered By: C2B Smart App | Page {i} | Generated: {DateTime.Now:dd-MMM-yyyy HH:mm}";
                float textWidth = normalFont.GetWidth(footerText, 6);
                canvas.MoveText(centerX - (textWidth / 2), footerY + 5);
                canvas.ShowText(footerText);

                canvas.EndText();
                canvas.RestoreState();
            }
        }

        #endregion

        #region Helper Methods

        private string NumberToWords(long number)
        {
            if (number == 0) return "Zero";
            if (number < 0) return "Minus " + NumberToWords(Math.Abs(number));

            string[] units = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
                               "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
                               "Eighteen", "Nineteen" };
            string[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

            string words = "";

            if ((number / 10000000) > 0)
            {
                words += NumberToWords(number / 10000000) + " Crore ";
                number %= 10000000;
            }

            if ((number / 100000) > 0)
            {
                words += NumberToWords(number / 100000) + " Lakh ";
                number %= 100000;
            }

            if ((number / 1000) > 0)
            {
                words += NumberToWords(number / 1000) + " Thousand ";
                number %= 1000;
            }

            if ((number / 100) > 0)
            {
                words += NumberToWords(number / 100) + " Hundred ";
                number %= 100;
            }

            if (number > 0)
            {
                if (number < 20)
                    words += units[number];
                else
                {
                    words += tens[number / 10];
                    if ((number % 10) > 0)
                        words += "-" + units[number % 10];
                }
            }

            return words.Trim();
        }

        #endregion
    }

    // Extension class for additional border
    public class BottomBorder : SolidBorder
    {
        public BottomBorder(iText.Kernel.Colors.Color color, float width) : base(color, width) { }

        public override void Draw(PdfCanvas canvas, float x1, float y1, float x2, float y2,
            Border.Side defaultSide, float borderWidthBefore, float borderWidthAfter)
        {
            if (defaultSide == Border.Side.BOTTOM)
            {
                base.Draw(canvas, x1, y1, x2, y2, defaultSide, borderWidthBefore, borderWidthAfter);
            }
        }
    }
}