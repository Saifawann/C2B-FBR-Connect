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
        // Clean color scheme
        private static readonly DeviceRgb PRIMARY_COLOR = new DeviceRgb(41, 128, 185);
        private static readonly DeviceRgb TABLE_HEADER = new DeviceRgb(52, 73, 94);
        private static readonly DeviceRgb TABLE_ALT_ROW = new DeviceRgb(249, 250, 251);
        private static readonly DeviceRgb BORDER_COLOR = new DeviceRgb(206, 212, 218);
        private static readonly DeviceRgb TEXT_PRIMARY = new DeviceRgb(33, 37, 41);
        private static readonly DeviceRgb TEXT_SECONDARY = new DeviceRgb(108, 117, 125);
        private static readonly DeviceRgb ACCENT_GREEN = new DeviceRgb(40, 167, 69);
        private static readonly DeviceRgb LIGHT_BG = new DeviceRgb(248, 249, 250);
        private static readonly DeviceRgb HIGHLIGHT_BG = new DeviceRgb(255, 243, 205);

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

                // Minimal margins for maximum content space
                document = new Document(pdfDoc);
                document.SetMargins(12, 12, 95, 12);  // Top, Right, Bottom (for FBR), Left

                var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                var normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

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

        private void BuildPDFContent(Document doc, Invoice invoice, Company company,
            PdfFont boldFont, PdfFont normalFont)
        {
            AddHeader(doc, company, invoice, boldFont, normalFont);
            AddPartyDetails(doc, invoice, company, boldFont, normalFont);
            AddItemsTable(doc, invoice, boldFont, normalFont);
            AddSummarySection(doc, invoice, boldFont, normalFont);
            AddFBRSection(doc, invoice, boldFont, normalFont);
            AddFooter(doc, normalFont);
        }

        #region Header Section

        private void AddHeader(Document doc, Company company, Invoice invoice,
            PdfFont boldFont, PdfFont normalFont)
        {
            // Main header table: Logo | Company Info | Invoice Info
            Table headerTable = new Table(UnitValue.CreatePercentArray(new float[] { 0.8f, 2.2f, 1.2f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(6);

            // LEFT: Company Logo
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

                    float maxWidth = 65f;
                    float maxHeight = 38f;
                    float widthRatio = maxWidth / logo.GetImageWidth();
                    float heightRatio = maxHeight / logo.GetImageHeight();
                    float scaleFactor = Math.Min(widthRatio, heightRatio);

                    logo.SetWidth(logo.GetImageWidth() * scaleFactor);
                    logo.SetHeight(logo.GetImageHeight() * scaleFactor);
                    logoCell.Add(logo);
                }
                catch
                {
                    AddFallbackLogo(logoCell, company, boldFont);
                }
            }
            else
            {
                AddFallbackLogo(logoCell, company, boldFont);
            }

            headerTable.AddCell(logoCell);

            // CENTER: Company Information
            Cell companyCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPaddingLeft(8)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            companyCell.Add(new Paragraph(company.CompanyName ?? "Company Name")
                .SetFont(boldFont)
                .SetFontSize(11)
                .SetFontColor(PRIMARY_COLOR)
                .SetMarginBottom(1));

            if (!string.IsNullOrEmpty(company.SellerAddress))
            {
                companyCell.Add(new Paragraph(company.SellerAddress)
                    .SetFont(normalFont)
                    .SetFontSize(6.5f)
                    .SetFontColor(TEXT_SECONDARY)
                    .SetMarginBottom(1));
            }

            // Contact line
            var contactParts = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(company.SellerPhone)) contactParts.Add($"📞 {company.SellerPhone}");
            if (!string.IsNullOrEmpty(company.SellerEmail)) contactParts.Add($"✉ {company.SellerEmail}");

            if (contactParts.Count > 0)
            {
                companyCell.Add(new Paragraph(string.Join("  |  ", contactParts))
                    .SetFont(normalFont)
                    .SetFontSize(6)
                    .SetFontColor(TEXT_SECONDARY));
            }

            headerTable.AddCell(companyCell);

            // RIGHT: NTN & STR Box
            Cell ntnCell = new Cell()
                .SetBorder(new SolidBorder(PRIMARY_COLOR, 1))
                .SetBackgroundColor(LIGHT_BG)
                .SetPadding(5)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            // NTN Row
            Table ntnTable = new Table(UnitValue.CreatePercentArray(new float[] { 1f, 1.5f }))
                .UseAllAvailableWidth();

            ntnTable.AddCell(new Cell().SetBorder(Border.NO_BORDER)
                .Add(new Paragraph("NTN:").SetFont(normalFont).SetFontSize(7).SetFontColor(TEXT_SECONDARY)));
            ntnTable.AddCell(new Cell().SetBorder(Border.NO_BORDER)
                .Add(new Paragraph(company.SellerNTN ?? "N/A").SetFont(boldFont).SetFontSize(8).SetFontColor(PRIMARY_COLOR)));

            // STR Row
            ntnTable.AddCell(new Cell().SetBorder(Border.NO_BORDER)
                .Add(new Paragraph("STR #:").SetFont(normalFont).SetFontSize(7).SetFontColor(TEXT_SECONDARY)));
            ntnTable.AddCell(new Cell().SetBorder(Border.NO_BORDER)
                .Add(new Paragraph(company.StrNo ?? "N/A").SetFont(boldFont).SetFontSize(8).SetFontColor(PRIMARY_COLOR)));

            ntnCell.Add(ntnTable);
            headerTable.AddCell(ntnCell);

            doc.Add(headerTable);

            // Invoice Title Bar
            Table titleBar = new Table(UnitValue.CreatePercentArray(new float[] { 4f, 0f, 0f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(5);

            // Title
            Cell titleCell = new Cell()
                .SetBackgroundColor(TABLE_HEADER)
                .SetBorder(Border.NO_BORDER)
                .SetPadding(4)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            string invoiceTitle = invoice.InvoiceType == "Credit Memo" ? "CREDIT NOTE" : "SALES TAX INVOICE";
            titleCell.Add(new Paragraph(invoiceTitle)
                .SetFont(boldFont)
                .SetFontSize(10)
                .SetFontColor(ColorConstants.WHITE));

            titleBar.AddCell(titleCell);

            doc.Add(titleBar);
        }

        private void AddFallbackLogo(Cell logoCell, Company company, PdfFont boldFont)
        {
            string initials = "CO";
            if (!string.IsNullOrEmpty(company.CompanyName) && company.CompanyName.Length > 0)
            {
                var words = company.CompanyName.Split(' ');
                if (words.Length >= 2)
                    initials = $"{words[0][0]}{words[1][0]}".ToUpper();
                else
                    initials = company.CompanyName.Substring(0, Math.Min(2, company.CompanyName.Length)).ToUpper();
            }

            logoCell.Add(new Paragraph(initials)
                .SetFont(boldFont)
                .SetFontSize(16)
                .SetFontColor(PRIMARY_COLOR)
                .SetBackgroundColor(LIGHT_BG)
                .SetPadding(8)
                .SetTextAlignment(TextAlignment.CENTER));
        }

        #endregion

        #region Party Details Section

        private void AddPartyDetails(Document doc, Invoice invoice, Company company,
            PdfFont boldFont, PdfFont normalFont)
        {
            Customer customer = invoice.Customer;

            // Two column layout: Seller | Buyer
            Table partyTable = new Table(UnitValue.CreatePercentArray(new float[] { 1f, 1f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(6);

            // LEFT: Buyer Details
            Cell buyerCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetPadding(6)
                .SetBackgroundColor(ColorConstants.WHITE);

            buyerCell.Add(new Paragraph("BUYER DETAILS")
                .SetFont(boldFont)
                .SetFontSize(7)
                .SetFontColor(TEXT_SECONDARY)
                .SetMarginBottom(3));

            buyerCell.Add(new Paragraph($"M/s {customer?.CustomerName ?? "N/A"}")
                .SetFont(boldFont)
                .SetFontSize(8)
                .SetFontColor(TEXT_PRIMARY)
                .SetMarginBottom(2));

            // Buyer details
            AddDetailRow(buyerCell, "NTN", customer?.CustomerNTN, normalFont, boldFont);
            AddDetailRow(buyerCell, "STR #", customer?.CustomerStrNo, normalFont, boldFont);

            // Address (if exists)
            if (!string.IsNullOrEmpty(customer?.CustomerAddress))
            {
                buyerCell.Add(new Paragraph($"📍 {customer.CustomerAddress}")
                    .SetFont(normalFont)
                    .SetFontSize(6)
                    .SetFontColor(TEXT_SECONDARY)
                    .SetMarginTop(2));
            }

            // Contact info
            var contactInfo = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(customer?.CustomerPhone)) contactInfo.Add($"📞 {customer.CustomerPhone}");
            if (!string.IsNullOrEmpty(customer?.CustomerEmail)) contactInfo.Add($"✉ {customer.CustomerEmail}");

            if (contactInfo.Count > 0)
            {
                buyerCell.Add(new Paragraph(string.Join("  ", contactInfo))
                    .SetFont(normalFont)
                    .SetFontSize(6)
                    .SetFontColor(TEXT_SECONDARY));
            }

            partyTable.AddCell(buyerCell);

            // Right: invoice Details
            Cell sellerCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetPadding(6)
                .SetBackgroundColor(ColorConstants.WHITE);

            string dateStr = invoice.InvoiceDate != DateTime.MinValue
                ? invoice.InvoiceDate.ToString("dd-MMM-yyyy")
                : "N/A";

            AddDetailRow(sellerCell, "Invoice No", invoice.InvoiceNumber, normalFont, boldFont);
            AddDetailRow(sellerCell, "Date :", dateStr, normalFont, boldFont);

            partyTable.AddCell(sellerCell);

            doc.Add(partyTable);
        }

        private void AddDetailRow(Cell cell, string label, string value, PdfFont normalFont, PdfFont boldFont)
        {
            var para = new Paragraph()
                .Add(new Text($"{label}: ").SetFont(normalFont).SetFontSize(6.5f).SetFontColor(TEXT_SECONDARY))
                .Add(new Text(value ?? "N/A").SetFont(boldFont ?? normalFont).SetFontSize(7).SetFontColor(TEXT_PRIMARY))
                .SetMarginBottom(1);

            cell.Add(para);
        }

        #endregion

        #region Items Table

        private void AddItemsTable(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            if (invoice.Items == null || invoice.Items.Count == 0)
            {
                doc.Add(new Paragraph("No items")
                    .SetFont(normalFont)
                    .SetFontSize(8)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetPadding(10));
                return;
            }

            // Optimized column widths for maximum items
            Table itemsTable = new Table(UnitValue.CreatePercentArray(
                new float[] { 0.35f, 2.4f, 0.9f, 0.5f, 0.8f, 0.85f, 0.7f, 0.9f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(4);

            // Header row
            string[] headers = { "#", "Description", "HS Code", "Qty", "Rate", "Amount", "Disc", "Net" };

            foreach (var header in headers)
            {
                Cell headerCell = new Cell()
                    .Add(new Paragraph(header)
                        .SetFont(boldFont)
                        .SetFontSize(6.5f)
                        .SetFontColor(ColorConstants.WHITE))
                    .SetBackgroundColor(TABLE_HEADER)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                    .SetPadding(2.5f)
                    .SetBorder(new SolidBorder(ColorConstants.WHITE, 0.3f));

                itemsTable.AddHeaderCell(headerCell);
            }

            // Item rows
            int index = 1;
            foreach (var item in invoice.Items)
            {
                decimal grossAmount = item.TotalPrice > 0 ? item.TotalPrice : (item.Quantity * item.UnitPrice);
                decimal netAmount = Math.Max(0, grossAmount - item.Discount);
                decimal unitRate = item.Quantity > 0 ? grossAmount / item.Quantity : item.UnitPrice;

                var rowColor = index % 2 == 0 ? TABLE_ALT_ROW : ColorConstants.WHITE;

                // Ultra compact cells
                AddTableCell(itemsTable, index.ToString(), TextAlignment.CENTER, normalFont, 6, rowColor);
                AddTableCell(itemsTable, TruncateText(item.ItemName ?? "Item", 35), TextAlignment.LEFT, normalFont, 6, rowColor);
                AddTableCell(itemsTable, item.HSCode ?? "-", TextAlignment.CENTER, normalFont, 6, rowColor);
                AddTableCell(itemsTable, item.Quantity.ToString("N0"), TextAlignment.CENTER, normalFont, 6, rowColor);
                AddTableCell(itemsTable, unitRate.ToString("N2"), TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddTableCell(itemsTable, grossAmount.ToString("N2"), TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddTableCell(itemsTable, item.Discount > 0 ? item.Discount.ToString("N2") : "-", TextAlignment.RIGHT, normalFont, 6, rowColor);
                AddTableCell(itemsTable, netAmount.ToString("N2"), TextAlignment.RIGHT, boldFont, 6, rowColor);

                index++;
            }

            doc.Add(itemsTable);
        }

        private void AddTableCell(Table table, string text, TextAlignment align, PdfFont font,
            float fontSize, iText.Kernel.Colors.Color bgColor)
        {
            Cell cell = new Cell()
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(TEXT_PRIMARY))
                .SetTextAlignment(align)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPadding(1.5f)
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.25f));

            table.AddCell(cell);
        }

        private string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength - 2) + "..";
        }

        #endregion

        #region Summary Section

        private void AddSummarySection(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            // Calculate totals
            decimal grossAmount = 0, discount = 0, salesTax = 0;

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

            // Main layout: Amount in words (left) | Totals (right)
            Table summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 1.5f, 1f }))
                .UseAllAvailableWidth()
                .SetMarginBottom(8);

            // LEFT: Amount in words
            Cell wordsCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetBackgroundColor(HIGHLIGHT_BG)
                .SetPadding(6)
                .SetVerticalAlignment(VerticalAlignment.TOP);

            wordsCell.Add(new Paragraph("Amount in Words:")
                .SetFont(boldFont)
                .SetFontSize(6.5f)
                .SetFontColor(TEXT_SECONDARY)
                .SetMarginBottom(2));

            string amountInWords = NumberToWords((long)Math.Round(grandTotal));
            wordsCell.Add(new Paragraph($"Rupees {amountInWords} Only")
                .SetFont(normalFont)
                .SetFontSize(7)
                .SetFontColor(TEXT_PRIMARY));

            summaryTable.AddCell(wordsCell);

            // RIGHT: Financial Summary
            Cell totalsCell = new Cell()
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f))
                .SetPadding(4);

            Table totalsTable = new Table(UnitValue.CreatePercentArray(new float[] { 1.2f, 1f }))
                .UseAllAvailableWidth();

            AddSummaryRow(totalsTable, "Gross Amount", grossAmount.ToString("N2"), normalFont, false);
            AddSummaryRow(totalsTable, "Discount", $"({discount.ToString("N2")})", normalFont, false);
            AddSummaryRow(totalsTable, "Subtotal", subtotal.ToString("N2"), normalFont, false);
            AddSummaryRow(totalsTable, "Sales Tax", salesTax.ToString("N2"), normalFont, false);
            AddSummaryRow(totalsTable, "GRAND TOTAL", $"Rs. {grandTotal.ToString("N2")}", boldFont, true);

            totalsCell.Add(totalsTable);
            summaryTable.AddCell(totalsCell);

            doc.Add(summaryTable);
        }

        private void AddSummaryRow(Table table, string label, string value, PdfFont font, bool isTotal)
        {
            var bgColor = isTotal ? ACCENT_GREEN : ColorConstants.WHITE;
            var textColor = isTotal ? ColorConstants.WHITE : TEXT_PRIMARY;
            var fontSize = isTotal ? 7.5f : 6.5f;
            var borderColor = isTotal ? ACCENT_GREEN : BORDER_COLOR;

            Cell labelCell = new Cell()
                .Add(new Paragraph(label)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(textColor))
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(borderColor, 0.3f))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(2.5f);

            Cell valueCell = new Cell()
                .Add(new Paragraph(value)
                    .SetFont(font)
                    .SetFontSize(fontSize)
                    .SetFontColor(textColor))
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(borderColor, 0.3f))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetPadding(2.5f);

            table.AddCell(labelCell);
            table.AddCell(valueCell);
        }

        #endregion

        #region FBR Section

        private void AddFBRSection(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            PdfDocument pdfDoc = doc.GetPdfDocument();
            int numberOfPages = pdfDoc.GetNumberOfPages();

            for (int i = 1; i <= numberOfPages; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                iText.Kernel.Geom.Rectangle pageSize = page.GetPageSize();
                PdfCanvas canvas = new PdfCanvas(page);

                float qrCodeSize = 72f;  // 1.0 inch as per FBR spec
                float fbrLogoSize = 55f;
                float borderHeight = 85f;
                float borderPadding = 8f;

                float fbrY = 55;
                float leftX = pageSize.GetLeft() + 12;
                float rightX = pageSize.GetRight() - 12;
                float borderWidth = pageSize.GetWidth() - 24;

                // FBR section border
                canvas.SaveState();
                canvas.SetStrokeColor(PRIMARY_COLOR);
                canvas.SetLineWidth(1.5f);
                canvas.Rectangle(leftX, fbrY, borderWidth, borderHeight);
                canvas.Stroke();

                // FBR label background
                canvas.SetFillColor(PRIMARY_COLOR);
                canvas.Rectangle(leftX, fbrY + borderHeight - 15, 80, 15);
                canvas.Fill();
                canvas.RestoreState();

                // FBR label text
                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(boldFont, 7);
                canvas.SetFillColor(ColorConstants.WHITE);
                canvas.MoveText(leftX + 8, fbrY + borderHeight - 11);
                canvas.ShowText("FBR VERIFICATION");
                canvas.EndText();
                canvas.RestoreState();

                // FBR Logo
                AddFBRLogo(doc, canvas, leftX + borderPadding, fbrY, borderHeight, fbrLogoSize, boldFont, i);

                // FBR Invoice details (center)
                float centerX = leftX + fbrLogoSize + 25;
                float textY = fbrY + borderHeight - 35;

                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(normalFont, 7);
                canvas.SetFillColor(TEXT_SECONDARY);
                canvas.MoveText(centerX, textY);
                canvas.ShowText("FBR Invoice Reference Number (IRN):");
                canvas.EndText();
                canvas.RestoreState();

                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(boldFont, 9);
                canvas.SetFillColor(PRIMARY_COLOR);
                string fbrIRN = invoice.FBR_IRN ?? "PENDING UPLOAD";
                canvas.MoveText(centerX, textY - 14);
                canvas.ShowText(fbrIRN);
                canvas.EndText();
                canvas.RestoreState();

                // Verification instruction
                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(normalFont, 6);
                canvas.SetFillColor(TEXT_SECONDARY);
                canvas.MoveText(centerX, textY - 30);
                canvas.ShowText("Scan QR code or visit fbr.gov.pk to verify");
                canvas.EndText();
                canvas.RestoreState();

                // QR Code
                AddQRCode(doc, canvas, invoice, rightX - qrCodeSize - borderPadding, fbrY, borderHeight, qrCodeSize, normalFont, i);
            }
        }

        private void AddFBRLogo(Document doc, PdfCanvas canvas, float x, float fbrY, float borderHeight,
            float logoSize, PdfFont boldFont, int pageNum)
        {
            string fbrLogoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "fbr_logo.png");

            if (File.Exists(fbrLogoPath))
            {
                try
                {
                    var fbrLogoData = ImageDataFactory.Create(fbrLogoPath);
                    var fbrLogoImage = new iText.Layout.Element.Image(fbrLogoData);

                    float widthRatio = logoSize / fbrLogoImage.GetImageWidth();
                    float heightRatio = logoSize / fbrLogoImage.GetImageHeight();
                    float scaleFactor = Math.Min(widthRatio, heightRatio);

                    float scaledWidth = fbrLogoImage.GetImageWidth() * scaleFactor;
                    float scaledHeight = fbrLogoImage.GetImageHeight() * scaleFactor;

                    var logo = new iText.Layout.Element.Image(fbrLogoData)
                        .SetWidth(scaledWidth)
                        .SetHeight(scaledHeight);

                    float logoY = fbrY + (borderHeight - scaledHeight) / 2 - 5;
                    logo.SetFixedPosition(pageNum, x, logoY);
                    doc.Add(logo);
                }
                catch
                {
                    DrawFBRFallback(canvas, x, fbrY, borderHeight, boldFont);
                }
            }
            else
            {
                DrawFBRFallback(canvas, x, fbrY, borderHeight, boldFont);
            }
        }

        private void DrawFBRFallback(PdfCanvas canvas, float x, float fbrY, float borderHeight, PdfFont boldFont)
        {
            canvas.SaveState();
            canvas.BeginText();
            canvas.SetFontAndSize(boldFont, 18);
            canvas.SetFillColor(PRIMARY_COLOR);
            canvas.MoveText(x + 5, fbrY + borderHeight / 2 - 5);
            canvas.ShowText("FBR");
            canvas.EndText();
            canvas.RestoreState();
        }

        private void AddQRCode(Document doc, PdfCanvas canvas, Invoice invoice, float x, float fbrY,
            float borderHeight, float qrSize, PdfFont normalFont, int pageNum)
        {
            float qrY = fbrY + (borderHeight - qrSize) / 2 - 3;

            if (!string.IsNullOrEmpty(invoice.FBR_IRN))
            {
                try
                {
                    var qrData = GenerateQRCode(invoice.FBR_IRN);
                    var qrImage = new iText.Layout.Element.Image(qrData)
                        .SetWidth(qrSize)
                        .SetHeight(qrSize);

                    qrImage.SetFixedPosition(pageNum, x, qrY);
                    doc.Add(qrImage);
                }
                catch
                {
                    DrawQRPlaceholder(canvas, x, qrY, qrSize, normalFont, "QR Error");
                }
            }
            else
            {
                DrawQRPlaceholder(canvas, x, qrY, qrSize, normalFont, "QR Pending");
            }
        }

        private void DrawQRPlaceholder(PdfCanvas canvas, float x, float y, float size, PdfFont font, string text)
        {
            canvas.SaveState();
            canvas.SetStrokeColor(BORDER_COLOR);
            canvas.SetFillColor(LIGHT_BG);
            canvas.SetLineWidth(1);
            canvas.Rectangle(x, y, size, size);
            canvas.FillStroke();

            canvas.BeginText();
            canvas.SetFontAndSize(font, 8);
            canvas.SetFillColor(TEXT_SECONDARY);
            canvas.MoveText(x + size / 2 - 20, y + size / 2);
            canvas.ShowText(text);
            canvas.EndText();
            canvas.RestoreState();
        }

        private ImageData GenerateQRCode(string content)
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
            using var qrCode = new QRCode(qrCodeData);

            // Version 2.0: 12 pixels per module for 300 DPI at 1 inch
            using var bitmap = qrCode.GetGraphic(
                pixelsPerModule: 12,
                darkColor: System.Drawing.Color.Black,
                lightColor: System.Drawing.Color.White,
                drawQuietZones: true);

            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            return ImageDataFactory.Create(ms.ToArray());
        }

        #endregion

        #region Footer

        private void AddFooter(Document doc, PdfFont normalFont)
        {
            PdfDocument pdfDoc = doc.GetPdfDocument();
            int numberOfPages = pdfDoc.GetNumberOfPages();

            for (int i = 1; i <= numberOfPages; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                iText.Kernel.Geom.Rectangle pageSize = page.GetPageSize();
                PdfCanvas canvas = new PdfCanvas(page);

                float footerY = 18;
                float centerX = pageSize.GetWidth() / 2;

                // Footer background
                canvas.SaveState();
                canvas.SetFillColor(TABLE_HEADER);
                canvas.Rectangle(pageSize.GetLeft(), footerY - 3, pageSize.GetWidth(), 22);
                canvas.Fill();
                canvas.RestoreState();

                // Footer text
                canvas.SaveState();
                canvas.BeginText();
                canvas.SetFontAndSize(normalFont, 6);
                canvas.SetFillColor(ColorConstants.WHITE);

                string footerText = $"Powered by C2B Smart App  |  Page {i} of {numberOfPages}  |  Generated: {DateTime.Now:dd-MMM-yyyy HH:mm}";
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
}