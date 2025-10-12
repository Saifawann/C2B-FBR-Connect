using C2B_FBR_Connect.Models;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace C2B_FBR_Connect.Services
{
    public class PDFService
    {
        private readonly PdfFont _boldFont;
        private readonly PdfFont _normalFont;

        public PDFService()
        {
            _boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            _normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
        }

        public void GenerateInvoicePDF(Invoice invoice, FBRInvoicePayload details, string outputPath)
        {
            ValidateInputs(invoice, details, outputPath);
            EnsureDirectoryExists(outputPath);

            using (var writer = new PdfWriter(outputPath))
            using (var pdfDoc = new PdfDocument(writer))
            using (var document = new Document(pdfDoc))
            {
                document.SetMargins(30, 30, 40, 30);
                BuildPDFContent(document, invoice, details);
            }
        }

        private void ValidateInputs(Invoice invoice, FBRInvoicePayload details, string outputPath)
        {
            if (invoice == null) throw new ArgumentNullException(nameof(invoice));
            if (details == null) throw new ArgumentNullException(nameof(details));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path cannot be empty");
        }

        private void EnsureDirectoryExists(string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        // 🔹 Core Layout Builder
        private void BuildPDFContent(Document doc, Invoice invoice, FBRInvoicePayload details)
        {
            AddHeaderSection(doc, details);
            AddInvoiceTitle(doc);
            AddCustomerAndPaymentSection(doc, invoice, details);
            AddItemsSection(doc, details);
            AddAmountInWords(doc, details);
            AddTotalsSection(doc, details);
            AddFBRSection(doc, invoice);
            AddFooter(doc);
        }

        #region Header
        private void AddHeaderSection(Document doc, FBRInvoicePayload details)
        {
            Table header = new Table(UnitValue.CreatePercentArray(new float[] { 2, 1 }))
                .UseAllAvailableWidth();

            // Left: Company Info
            var companyInfo = new Paragraph(details.SellerBusinessName ?? "N/A")
                .SetFont(_boldFont).SetFontSize(12).SetMarginBottom(2);
            companyInfo.Add("\n" + (details.SellerAddress ?? ""))
                        .Add("\nContact: " + (""))
                        .Add("\nEmail: " + (""));
            header.AddCell(new Cell().Add(companyInfo)
                .SetBorder(Border.NO_BORDER));

            // Right: Tax Info
            var taxInfo = new Paragraph($"NTN: {details.SellerNTN ?? "N/A"}")
                .SetFont(_normalFont).SetFontSize(10);
            taxInfo.Add("\nSTRN: " + ("N/A"))
                   .Add("\nPOS ID: " + ("N/A"));
            header.AddCell(new Cell().Add(taxInfo)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.RIGHT));

            doc.Add(header);
            doc.Add(new Paragraph("\n"));
        }

        private void AddInvoiceTitle(Document doc)
        {
            var title = new Paragraph("SALES TAX INVOICE")
                .SetFont(_boldFont)
                .SetFontSize(16)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetMarginBottom(10);
            doc.Add(title);
        }
        #endregion

        #region Customer + Payment
        private void AddCustomerAndPaymentSection(Document doc, Invoice invoice, FBRInvoicePayload details)
        {
            Table info = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .UseAllAvailableWidth()
                .SetMarginBottom(10);

            // Left: Customer Info
            var customer = new Paragraph("To: M/s " + (invoice.CustomerName ?? "N/A"))
                .SetFont(_boldFont).SetFontSize(10);
            customer.Add("\nAddress: " + (""))
                    .Add("\nPhone: " + (""));
            info.AddCell(new Cell().Add(customer).SetPadding(8)
                .SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f)));

            // Right: Payment Terms
            var payment = new Paragraph("Terms of Payment")
                .SetFont(_boldFont).SetFontSize(10);
            payment.Add("\nInvoice #: " + (invoice.InvoiceNumber ?? "N/A"))
                   .Add("\nDate: " + details.InvoiceDate.ToString("dd-MMM-yyyy"))
                   .Add("\nPayment Mode: " + ("Cash"));
            info.AddCell(new Cell().Add(payment).SetPadding(8)
                .SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f)));

            doc.Add(info);
        }
        #endregion

        #region Items Table
        private void AddItemsSection(Document doc, FBRInvoicePayload details)
        {
            if (details.Items == null || details.Items.Count == 0)
            {
                doc.Add(new Paragraph("No items to display").SetFont(_normalFont).SetFontSize(10));
                return;
            }

            Table table = new Table(UnitValue.CreatePercentArray(new float[]
                { 1, 3, 2, 1.5f, 1.5f, 1.5f, 1.5f, 1.5f }))
                .UseAllAvailableWidth()
                .SetMarginTop(10);

            string[] headers = { "S#", "Description", "HS Code", "Qty", "Rate", "Amount", "Discount", "Net Amt" };
            foreach (var h in headers)
                table.AddHeaderCell(new Cell().Add(new Paragraph(h).SetFont(_boldFont).SetFontSize(9))
                    .SetBackgroundColor(new DeviceRgb(230, 230, 230))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetPadding(5));

            int index = 1;
            foreach (var item in details.Items)
            {
                table.AddCell(CreateItemCell(index.ToString(), TextAlignment.CENTER));
                table.AddCell(CreateItemCell(item.ItemName ?? "", TextAlignment.LEFT));
                table.AddCell(CreateItemCell(item.HSCode ?? "-", TextAlignment.CENTER));
                table.AddCell(CreateItemCell(item.Quantity.ToString("N2"), TextAlignment.CENTER));
                table.AddCell(CreateItemCell($"Rs. {item.UnitPrice:N2}", TextAlignment.RIGHT));
                table.AddCell(CreateItemCell($"Rs. {item.TotalPrice:N2}", TextAlignment.RIGHT));
                table.AddCell(CreateItemCell($"Rs. {item.Discount:N2}", TextAlignment.RIGHT));
                table.AddCell(CreateItemCell($"Rs. {(item.TotalPrice - item.Discount):N2}", TextAlignment.RIGHT));
                index++;
            }

            doc.Add(table);
        }

        private Cell CreateItemCell(string text, TextAlignment align)
        {
            return new Cell().Add(new Paragraph(text).SetFont(_normalFont).SetFontSize(9))
                .SetTextAlignment(align).SetPadding(5);
        }
        #endregion

        #region Amount in Words
        private void AddAmountInWords(Document doc, FBRInvoicePayload details)
        {
            var total = details.TotalAmount + details.TaxAmount;
            string words = NumberToWords((long)Math.Round(total));
            var para = new Paragraph("Amount in Words: " + words + " Rupees Only")
                .SetFont(_boldFont)
                .SetFontSize(9)
                .SetMarginTop(10);
            doc.Add(para);
        }

        private string NumberToWords(long number)
        {
            if (number == 0) return "Zero";
            if (number < 0) return "Minus " + NumberToWords(Math.Abs(number));

            string[] units = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
                               "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
                               "Eighteen", "Nineteen" };
            string[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

            string words = "";
            if ((number / 1000000) > 0)
            {
                words += NumberToWords(number / 1000000) + " Million ";
                number %= 1000000;
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

        #region Totals Section
        private void AddTotalsSection(Document doc, FBRInvoicePayload details)
        {
            var table = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .SetWidth(UnitValue.CreatePercentValue(40))
                .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.RIGHT)
                .SetMarginTop(10);

            AddTotalRow(table, "Gross Total", $"Rs. {details.TotalAmount:N2}");
            AddTotalRow(table, "Discount", $"Rs. ");
            AddTotalRow(table, "Sales Tax", $"Rs. {details.TaxAmount:N2}");

            var total = details.TotalAmount + details.TaxAmount;
            AddTotalRow(table, "Total (Incl. Tax)", $"Rs. {total:N2}", true);

            doc.Add(table);
        }

        private void AddTotalRow(Table table, string label, string value, bool bold = false)
        {
            table.AddCell(new Cell().Add(new Paragraph(label).SetFont(bold ? _boldFont : _normalFont).SetFontSize(9))
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(3));
            table.AddCell(new Cell().Add(new Paragraph(value).SetFont(bold ? _boldFont : _normalFont).SetFontSize(9))
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetPadding(3));
        }
        #endregion

        #region FBR Section
        private void AddFBRSection(Document doc, Invoice invoice)
        {
            doc.Add(new Paragraph("\n"));

            Table fbr = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1, 1 }))
                .UseAllAvailableWidth();

            // Left: FBR Logo
            string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "fbr_logo.png");
            if (File.Exists(logoPath))
            {
                var logo = new iText.Layout.Element.Image(ImageDataFactory.Create(logoPath))
                    .SetWidth(80).SetHeight(80).SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT);
                fbr.AddCell(new Cell().Add(logo).SetBorder(Border.NO_BORDER));
            }
            else
            {
                fbr.AddCell(new Cell().Add(new Paragraph("[FBR Logo Missing]").SetFont(_normalFont).SetFontSize(8))
                    .SetBorder(Border.NO_BORDER));
            }

            // Center: IRN
            var irnCell = new Cell().SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            irnCell.Add(new Paragraph("FBR INVOICE REFERENCE NUMBER")
                .SetFont(_boldFont).SetFontSize(9));
            irnCell.Add(new Paragraph(invoice.FBR_IRN ?? "N/A")
                .SetFont(_normalFont).SetFontSize(9));
            fbr.AddCell(irnCell);

            // Right: QR Code
            if (!string.IsNullOrEmpty(invoice.FBR_IRN))
            {
                var qrData = GenerateQRCodeImage(invoice.FBR_IRN);
                var qrImage = new iText.Layout.Element.Image(qrData)
                    .SetWidth(80).SetHeight(80)
                    .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.RIGHT);
                fbr.AddCell(new Cell().Add(qrImage).SetBorder(Border.NO_BORDER));
            }
            else
            {
                fbr.AddCell(new Cell().Add(new Paragraph("QR Not Available").SetFont(_normalFont).SetFontSize(8))
                    .SetBorder(Border.NO_BORDER)
                    .SetTextAlignment(TextAlignment.RIGHT));
            }

            doc.Add(fbr);
        }

        private ImageData GenerateQRCodeImage(string content)
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);
            using var qrCode = new QRCode(qrCodeData);
            using var bitmap = qrCode.GetGraphic(20, System.Drawing.Color.Black, System.Drawing.Color.White, true);
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            return ImageDataFactory.Create(ms.ToArray());
        }
        #endregion

        #region Footer
        private void AddFooter(Document doc)
        {
            doc.Add(new Paragraph("\n"));
            doc.Add(new Paragraph("Powered by C2B FBR Connect")
                .SetFont(_boldFont)
                .SetFontSize(9)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontColor(ColorConstants.GRAY)
                .SetMarginTop(20));
        }
        #endregion
    }
}
