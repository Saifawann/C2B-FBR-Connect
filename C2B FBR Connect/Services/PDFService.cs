using C2B_FBR_Connect.Models;
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
        // Define color scheme for consistency
        private static readonly DeviceRgb PRIMARY_COLOR = new DeviceRgb(0, 123, 255);  // Professional blue
        private static readonly DeviceRgb SECONDARY_COLOR = new DeviceRgb(52, 58, 64);  // Dark gray
        private static readonly DeviceRgb LIGHT_GRAY = new DeviceRgb(248, 249, 250);
        private static readonly DeviceRgb BORDER_COLOR = new DeviceRgb(222, 226, 230);
        private static readonly DeviceRgb SUCCESS_COLOR = new DeviceRgb(40, 167, 69);  // Green for totals

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

                // Set page size and create document
                pdfDoc.SetDefaultPageSize(PageSize.A4);
                document = new Document(pdfDoc);
                document.SetMargins(25, 25, 35, 25);

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
            AddHeaderSection(doc, invoice, company, boldFont, normalFont);
            AddDivider(doc);
            AddInvoiceTitle(doc, boldFont, invoice);
            AddCustomerAndPaymentSection(doc, invoice, boldFont, normalFont);
            AddItemsSection(doc, invoice, boldFont, normalFont);
            AddAmountInWords(doc, invoice, boldFont, normalFont);
            AddTotalsSection(doc, invoice, boldFont, normalFont);
            AddDivider(doc);
            AddFBRSection(doc, invoice, boldFont, normalFont);
            AddSignatureSection(doc, boldFont, normalFont);
            AddFooter(doc, normalFont);
        }

        #region Header Section

        private void AddHeaderSection(Document doc, Invoice invoice, Company company,
            PdfFont boldFont, PdfFont normalFont)
        {
            Table header = new Table(UnitValue.CreatePercentArray(new float[] { 2, 1 }))
                .UseAllAvailableWidth();

            // Left: Company Info with better styling
            Cell companyCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(10);

            // Company Name with larger font and color
            var companyName = new Paragraph(company.CompanyName ?? "Company Name")
                .SetFont(boldFont)
                .SetFontSize(16)
                .SetFontColor(PRIMARY_COLOR)
                .SetMarginBottom(5);
            companyCell.Add(companyName);

            // Company details
            var addressPara = new Paragraph()
                .SetFont(normalFont)
                .SetFontSize(10)
                .SetFontColor(SECONDARY_COLOR);

            if (!string.IsNullOrEmpty(company.SellerAddress))
                addressPara.Add("📍 " + company.SellerAddress + "\n");
            if (!string.IsNullOrEmpty(company.SellerPhone))
                addressPara.Add("📞 " + company.SellerPhone + "\n");
            if (!string.IsNullOrEmpty(company.SellerEmail))
                addressPara.Add("✉ " + company.SellerEmail);

            companyCell.Add(addressPara);
            header.AddCell(companyCell);

            // Right: Tax Info in a styled box
            Cell taxCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(10);

            // Create a nested table for tax info box
            Table taxBox = new Table(UnitValue.CreatePercentArray(new float[] { 1 }))
                .UseAllAvailableWidth();

            Cell taxInfoCell = new Cell()
                .SetBackgroundColor(LIGHT_GRAY)
                .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                .SetBorderRadius(new BorderRadius(5))
                .SetPadding(8);

            var taxInfo = new Paragraph("TAX INFORMATION")
                .SetFont(boldFont)
                .SetFontSize(9)
                .SetFontColor(SECONDARY_COLOR)
                .SetMarginBottom(3);
            taxInfoCell.Add(taxInfo);

            taxInfoCell.Add(new Paragraph($"NTN: {company.SellerNTN ?? "N/A"}")
                .SetFont(normalFont)
                .SetFontSize(10)
                .SetFontColor(SECONDARY_COLOR));

            taxInfoCell.Add(new Paragraph($"Province: {company.SellerProvince ?? "N/A"}")
                .SetFont(normalFont)
                .SetFontSize(10)
                .SetFontColor(SECONDARY_COLOR));

            taxBox.AddCell(taxInfoCell);
            taxCell.Add(taxBox);
            header.AddCell(taxCell);

            doc.Add(header);
        }

        private void AddDivider(Document doc)
        {
            LineSeparator line = new LineSeparator(new SolidLine(1f));
            line.SetStrokeColor(BORDER_COLOR);
            doc.Add(line);
            doc.Add(new Paragraph("\n").SetFontSize(5));
        }

        private void AddInvoiceTitle(Document doc, PdfFont boldFont, Invoice invoice)
        {
            // Main title
            var title = new Paragraph("SALES TAX INVOICE")
                .SetFont(boldFont)
                .SetFontSize(18)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontColor(SECONDARY_COLOR)
                .SetMarginBottom(10);
            doc.Add(title);
        }

        #endregion

        #region Customer + Payment Section

        private void AddCustomerAndPaymentSection(Document doc, Invoice invoice,
            PdfFont boldFont, PdfFont normalFont)
        {
            Table info = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .UseAllAvailableWidth()
                .SetMarginBottom(15);

            // Left: Customer Info with better styling
            Cell customerCell = new Cell()
                .SetBackgroundColor(new DeviceRgb(255, 255, 255))
                .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                .SetBorderRadius(new BorderRadius(5))
                .SetPadding(12);

            var customerTitle = new Paragraph("BILL TO")
                .SetFont(boldFont)
                .SetFontSize(10)
                .SetFontColor(PRIMARY_COLOR)
                .SetMarginBottom(8);
            customerCell.Add(customerTitle);

            var customerInfo = new Paragraph()
                .SetFont(normalFont)
                .SetFontSize(10);

            customerInfo.Add(new Text(invoice.CustomerName ?? "N/A")
                .SetFont(boldFont)
                .SetFontSize(11));

            if (!string.IsNullOrEmpty(invoice.CustomerAddress))
                customerInfo.Add("\n" + invoice.CustomerAddress);

            if (!string.IsNullOrEmpty(invoice.CustomerPhone))
                customerInfo.Add("\nPhone: " + invoice.CustomerPhone);

            if (!string.IsNullOrEmpty(invoice.CustomerNTN))
                customerInfo.Add("\nNTN: " + invoice.CustomerNTN);

            customerCell.Add(customerInfo);
            info.AddCell(customerCell);

            // Right: Invoice Details
            Cell invoiceCell = new Cell()
                .SetBackgroundColor(LIGHT_GRAY)
                .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                .SetBorderRadius(new BorderRadius(5))
                .SetPadding(12);

            var invoiceTitle = new Paragraph("INVOICE DETAILS")
                .SetFont(boldFont)
                .SetFontSize(10)
                .SetFontColor(PRIMARY_COLOR)
                .SetMarginBottom(8);
            invoiceCell.Add(invoiceTitle);

            // Create a small table for invoice details
            Table detailsTable = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .UseAllAvailableWidth();

            AddDetailRow(detailsTable, "Invoice #:", invoice.InvoiceNumber ?? "N/A", normalFont, boldFont);
            AddDetailRow(detailsTable, "Date:", invoice.InvoiceDate != DateTime.MinValue
                ? invoice.InvoiceDate.ToString("dd-MMM-yyyy") : "N/A", normalFont, normalFont);
            AddDetailRow(detailsTable, "Payment:", "Cash", normalFont, normalFont); // Default to Cash

            invoiceCell.Add(detailsTable);
            info.AddCell(invoiceCell);

            doc.Add(info);
        }

        private void AddDetailRow(Table table, string label, string value, PdfFont labelFont, PdfFont valueFont)
        {
            table.AddCell(new Cell()
                .Add(new Paragraph(label)
                    .SetFont(labelFont)
                    .SetFontSize(9)
                    .SetFontColor(SECONDARY_COLOR))
                .SetBorder(Border.NO_BORDER)
                .SetPaddingBottom(3));

            table.AddCell(new Cell()
                .Add(new Paragraph(value)
                    .SetFont(valueFont)
                    .SetFontSize(9))
                .SetBorder(Border.NO_BORDER)
                .SetPaddingBottom(3)
                .SetTextAlignment(TextAlignment.RIGHT));
        }

        #endregion

        #region Items Table

        private void AddItemsSection(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            if (invoice.Items == null || invoice.Items.Count == 0)
            {
                var noItems = new Paragraph("No items to display")
                    .SetFont(normalFont)
                    .SetFontSize(10)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontColor(SECONDARY_COLOR)
                    .SetMarginTop(20)
                    .SetMarginBottom(20);
                doc.Add(noItems);
                return;
            }

            var sectionTitle = new Paragraph("ITEMS")
                .SetFont(boldFont)
                .SetFontSize(11)
                .SetFontColor(SECONDARY_COLOR)
                .SetMarginTop(10)
                .SetMarginBottom(5);
            doc.Add(sectionTitle);

            Table table = new Table(UnitValue.CreatePercentArray(new float[]
                { 0.5f, 2.5f, 1.2f, 0.8f, 1f, 1.2f, 1f, 1.3f }))
                .UseAllAvailableWidth()
                .SetBorderRadius(new BorderRadius(5));

            string[] headers = { "#", "Description", "HS Code", "Qty", "Rate", "Amount", "Discount", "Net Amount" };
            foreach (var h in headers)
            {
                Cell headerCell = new Cell()
                    .Add(new Paragraph(h)
                        .SetFont(boldFont)
                        .SetFontSize(9)
                        .SetFontColor(ColorConstants.WHITE))
                    .SetBackgroundColor(PRIMARY_COLOR)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetPadding(8)
                    .SetBorder(Border.NO_BORDER);

                if (h == "#")
                    headerCell.SetBorderTopLeftRadius(new BorderRadius(5));
                else if (h == "Net Amount")
                    headerCell.SetBorderTopRightRadius(new BorderRadius(5));

                table.AddHeaderCell(headerCell);
            }

            // ✅ FIXED LOGIC: Handle both normal items and 3rd schedule items
            int index = 1;
            foreach (var item in invoice.Items)
            {
                decimal netAmountAfterDiscount;
                decimal originalAmount;

                // For normal items: NetAmount > 0 means it's properly set
                // For 3rd schedule items: NetAmount = 0 (tax is on retail price)

                if (item.NetAmount > 0)
                {
                    // ✅ Normal case: NetAmount is properly set (after discount, before tax)
                    netAmountAfterDiscount = item.NetAmount;
                    originalAmount = netAmountAfterDiscount + item.Discount;
                }
                else if (item.TotalPrice > 0)
                {
                    // ✅ 3rd Schedule or other special case
                    // TotalPrice has the original amount before discount
                    originalAmount = item.TotalPrice;
                    netAmountAfterDiscount = originalAmount - item.Discount;

                    if (netAmountAfterDiscount < 0)
                        netAmountAfterDiscount = 0;
                }
                else
                {
                    // ✅ Fallback for edge cases
                    originalAmount = 0;
                    netAmountAfterDiscount = 0;
                }

                var unitPrice = item.Quantity > 0 ? originalAmount / item.Quantity : item.UnitPrice;
                var rowColor = index % 2 == 0 ? LIGHT_GRAY : ColorConstants.WHITE;

                table.AddCell(CreateStyledItemCell(index.ToString(), TextAlignment.CENTER, normalFont, rowColor));
                table.AddCell(CreateStyledItemCell(item.ItemName ?? item.HSCode ?? "", TextAlignment.LEFT, normalFont, rowColor));
                table.AddCell(CreateStyledItemCell(item.HSCode ?? "-", TextAlignment.CENTER, normalFont, rowColor));
                table.AddCell(CreateStyledItemCell(item.Quantity.ToString("N2"), TextAlignment.CENTER, normalFont, rowColor));
                table.AddCell(CreateStyledItemCell($"{unitPrice:N2}", TextAlignment.RIGHT, normalFont, rowColor));
                table.AddCell(CreateStyledItemCell($"{originalAmount:N2}", TextAlignment.RIGHT, normalFont, rowColor)); // Shows 10,000

                var discountCell = CreateStyledItemCell($"{item.Discount:N2}", TextAlignment.RIGHT, normalFont, rowColor);
                if (item.Discount > 0)
                    discountCell.SetFontColor(new DeviceRgb(220, 53, 69));
                table.AddCell(discountCell); // Shows 1,000

                table.AddCell(CreateStyledItemCell($"{netAmountAfterDiscount:N2}", TextAlignment.RIGHT, boldFont, rowColor)); // Shows 9,000
                index++;
            }

            doc.Add(table);
        }

        private Cell CreateStyledItemCell(string text, TextAlignment align, PdfFont font, iText.Kernel.Colors.Color bgColor)
        {
            return new Cell()
                .Add(new Paragraph(text)
                    .SetFont(font)
                    .SetFontSize(9))
                .SetTextAlignment(align)
                .SetPadding(6)
                .SetBackgroundColor(bgColor)
                .SetBorder(new SolidBorder(BORDER_COLOR, 0.5f));
        }

        #endregion

        #region Amount in Words

        private void AddAmountInWords(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            // Calculate the correct total
            var subtotal = invoice.Items?.Sum(i => {
                var baseAmount = i.NetAmount > 0 ? i.NetAmount :
                                (i.TotalPrice > 0 && i.SalesTaxAmount > 0 ?
                                 i.TotalPrice - i.SalesTaxAmount : i.TotalPrice);
                return baseAmount; // After discount
            }) ?? 0;

            var tax = invoice.Items?.Sum(i => i.SalesTaxAmount) ?? 0;
            var total = subtotal + tax;

            string words = NumberToWords((long)Math.Round(total));

            // Create a highlighted box for amount in words
            Table amountBox = new Table(UnitValue.CreatePercentArray(new float[] { 1 }))
                .UseAllAvailableWidth()
                .SetMarginTop(10)
                .SetMarginBottom(10);

            Cell amountCell = new Cell()
                .SetBackgroundColor(new DeviceRgb(255, 253, 235))  // Light yellow background
                .SetBorder(new SolidBorder(new DeviceRgb(255, 193, 7), 1))  // Yellow border
                .SetBorderRadius(new BorderRadius(5))
                .SetPadding(10);

            var amountPara = new Paragraph()
                .Add(new Text("Amount in Words: ")
                    .SetFont(boldFont)
                    .SetFontSize(10))
                .Add(new Text(words + " Rupees Only")
                    .SetFont(normalFont)  // Changed from italicFont
                    .SetFontSize(10)
                    .SetFontColor(SECONDARY_COLOR));

            amountCell.Add(amountPara);
            amountBox.AddCell(amountCell);
            doc.Add(amountBox);
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

            // Handle Crores (10 million)
            if ((number / 10000000) > 0)
            {
                words += NumberToWords(number / 10000000) + " Crore ";
                number %= 10000000;
            }

            // Handle Lakhs (100 thousand)
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

        #region Totals Section

        private void AddTotalsSection(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            Table container = new Table(UnitValue.CreatePercentArray(new float[] { 1.5f, 1 }))
                .UseAllAvailableWidth()
                .SetMarginTop(10);

            container.AddCell(new Cell().SetBorder(Border.NO_BORDER));

            Cell totalsCell = new Cell().SetBorder(Border.NO_BORDER);
            Table totalsTable = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .UseAllAvailableWidth();

            // ✅ Calculate totals using clear logic
            decimal grossTotal = 0m;
            decimal totalDiscount = 0m;
            decimal totalTax = 0m;

            if (invoice.Items != null)
            {
                foreach (var item in invoice.Items)
                {
                    // Calculate original amount before discount
                    decimal originalAmount;

                    if (item.NetAmount > 0)
                    {
                        // ✅ STANDARD CASE: NetAmount is after discount
                        // Your data: NetAmount=9000, Discount=1000
                        // Original = 9000 + 1000 = 10000 ✅
                        originalAmount = item.NetAmount + item.Discount;
                    }
                    else if (item.TotalPrice > 0)
                    {
                        // ✅ 3RD SCHEDULE CASE: NetAmount=0
                        // Your data: TotalPrice=2000, Discount=0
                        // Original = 2000 ✅
                        originalAmount = item.TotalPrice;
                    }
                    else
                    {
                        // Fallback
                        originalAmount = 0m;
                    }

                    grossTotal += originalAmount;
                    totalDiscount += item.Discount;
                    totalTax += item.SalesTaxAmount;
                }
            }

            // Calculate derived amounts
            decimal subtotal = grossTotal - totalDiscount;  // 10000 - 1000 = 9000
            decimal grandTotal = subtotal + totalTax;       // 9000 + 1620 = 10620

            // Display rows
            AddStyledTotalRow(totalsTable, "Gross Total:", $"Rs. {grossTotal:N2}",
                normalFont, normalFont, LIGHT_GRAY, null);

            if (totalDiscount > 0)
            {
                AddStyledTotalRow(totalsTable, "Total Discount:", $"Rs. {totalDiscount:N2}",
                    normalFont, normalFont,
                    new DeviceRgb(255, 235, 235), new DeviceRgb(220, 53, 69));
            }

            AddStyledTotalRow(totalsTable, "Subtotal:", $"Rs. {subtotal:N2}",
                normalFont, normalFont, LIGHT_GRAY, null);

            AddStyledTotalRow(totalsTable, "Sales Tax:", $"Rs. {totalTax:N2}",
                normalFont, normalFont, LIGHT_GRAY, null);

            // Grand Total with styling
            Cell totalLabelCell = new Cell()
                .Add(new Paragraph("GRAND TOTAL:")
                    .SetFont(boldFont)
                    .SetFontSize(11)
                    .SetFontColor(ColorConstants.WHITE))
                .SetBackgroundColor(SUCCESS_COLOR)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(8);

            Cell totalValueCell = new Cell()
                .Add(new Paragraph($"Rs. {grandTotal:N2}")
                    .SetFont(boldFont)
                    .SetFontSize(11)
                    .SetFontColor(ColorConstants.WHITE))
                .SetBackgroundColor(SUCCESS_COLOR)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetPadding(8);

            totalsTable.AddCell(totalLabelCell);
            totalsTable.AddCell(totalValueCell);

            totalsCell.Add(totalsTable);
            container.AddCell(totalsCell);

            doc.Add(container);
        }

        private void AddStyledTotalRow(Table table, string label, string value, PdfFont labelFont,
            PdfFont valueFont, iText.Kernel.Colors.Color bgColor, iText.Kernel.Colors.Color textColor)
        {
            var labelCell = new Cell()
                .Add(new Paragraph(label)
                    .SetFont(labelFont)
                    .SetFontSize(10))
                .SetBackgroundColor(bgColor)
                .SetBorder(new BottomBorder(BORDER_COLOR, 0.5f))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetPadding(5);

            var valueCell = new Cell()
                .Add(new Paragraph(value)
                    .SetFont(valueFont)
                    .SetFontSize(10))
                .SetBackgroundColor(bgColor)
                .SetBorder(new BottomBorder(BORDER_COLOR, 0.5f))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetPadding(5);

            if (textColor != null)
            {
                labelCell.SetFontColor(textColor);
                valueCell.SetFontColor(textColor);
            }

            table.AddCell(labelCell);
            table.AddCell(valueCell);
        }

        #endregion

        #region FBR Section

        private void AddFBRSection(Document doc, Invoice invoice, PdfFont boldFont, PdfFont normalFont)
        {
            // FBR Section Title
            var fbrTitle = new Paragraph("FBR VERIFICATION")
                .SetFont(boldFont)
                .SetFontSize(11)
                .SetFontColor(SECONDARY_COLOR)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetMarginTop(15)
                .SetMarginBottom(10);
            doc.Add(fbrTitle);

            Table fbr = new Table(UnitValue.CreatePercentArray(new float[] { 1, 2, 1 }))
                .UseAllAvailableWidth();

            // Left: FBR Logo with border
            Cell logoCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER);

            string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "fbr_logo.png");
            if (File.Exists(logoPath))
            {
                var logo = new iText.Layout.Element.Image(ImageDataFactory.Create(logoPath))
                    .SetWidth(70)
                    .SetHeight(70)
                    .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
                logoCell.Add(logo);
            }
            else
            {
                // Placeholder for missing logo
                Cell placeholderCell = new Cell()
                    .SetWidth(70)
                    .SetHeight(70)
                    .SetBackgroundColor(LIGHT_GRAY)
                    .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                    .Add(new Paragraph("FBR\nLOGO")
                        .SetFont(boldFont)
                        .SetFontSize(12)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetFontColor(SECONDARY_COLOR));
                logoCell.Add(new Table(1).AddCell(placeholderCell));
            }
            fbr.AddCell(logoCell);

            // Center: IRN with better styling
            Cell irnCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            var irnBox = new Table(UnitValue.CreatePercentArray(new float[] { 1 }))
                .UseAllAvailableWidth();

            Cell irnContent = new Cell()
                .SetBackgroundColor(LIGHT_GRAY)
                .SetBorder(new SolidBorder(PRIMARY_COLOR, 1))
                .SetBorderRadius(new BorderRadius(5))
                .SetPadding(10);

            irnContent.Add(new Paragraph("FBR INVOICE REFERENCE NUMBER")
                .SetFont(boldFont)
                .SetFontSize(8)
                .SetFontColor(SECONDARY_COLOR)
                .SetMarginBottom(5));

            irnContent.Add(new Paragraph(invoice.FBR_IRN ?? "PENDING")
                .SetFont(boldFont)
                .SetFontSize(10)
                .SetFontColor(PRIMARY_COLOR));

            irnBox.AddCell(irnContent);
            irnCell.Add(irnBox);
            fbr.AddCell(irnCell);

            // Right: QR Code with border
            Cell qrCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER);

            if (!string.IsNullOrEmpty(invoice.FBR_IRN))
            {
                var qrData = GenerateQRCodeImage(invoice.FBR_IRN);
                var qrImage = new iText.Layout.Element.Image(qrData)
                    .SetWidth(70)
                    .SetHeight(70)
                    .SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);

                // Add QR in a bordered container
                Table qrContainer = new Table(1);
                qrContainer.AddCell(new Cell()
                    .Add(qrImage)
                    .SetBorder(new SolidBorder(BORDER_COLOR, 1))
                    .SetPadding(5));
                qrCell.Add(qrContainer);

                qrCell.Add(new Paragraph("Scan for Verification")
                    .SetFont(normalFont)
                    .SetFontSize(7)
                    .SetFontColor(SECONDARY_COLOR)
                    .SetMarginTop(3));
            }
            else
            {
                // Placeholder for QR
                Cell placeholderCell = new Cell()
                    .SetWidth(70)
                    .SetHeight(70)
                    .SetBackgroundColor(LIGHT_GRAY)
                    .SetBorder(new DashedBorder(BORDER_COLOR, 1))
                    .Add(new Paragraph("QR Code\nPending")
                        .SetFont(normalFont)
                        .SetFontSize(10)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetFontColor(SECONDARY_COLOR));
                qrCell.Add(new Table(1).AddCell(placeholderCell));
            }
            fbr.AddCell(qrCell);

            doc.Add(fbr);
        }

        private ImageData GenerateQRCodeImage(string content)
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);
            using var qrCode = new QRCode(qrCodeData);
            using var bitmap = qrCode.GetGraphic(20,
                System.Drawing.Color.Black,
                System.Drawing.Color.White,
                drawQuietZones: true);

            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            return ImageDataFactory.Create(ms.ToArray());
        }

        #endregion

        #region Signature Section

        private void AddSignatureSection(Document doc, PdfFont boldFont, PdfFont normalFont)
        {
            doc.Add(new Paragraph("\n"));

            Table signatureTable = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 }))
                .UseAllAvailableWidth()
                .SetMarginTop(20);

            // Authorized Signatory
            Cell authCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(10);

            authCell.Add(new Paragraph("_______________________")
                .SetFont(normalFont)
                .SetFontSize(10)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetMarginBottom(3));

            authCell.Add(new Paragraph("Authorized Signatory")
                .SetFont(boldFont)
                .SetFontSize(9)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontColor(SECONDARY_COLOR));

            signatureTable.AddCell(authCell);

            // Customer Signature
            Cell custCell = new Cell()
                .SetBorder(Border.NO_BORDER)
                .SetPadding(10);

            custCell.Add(new Paragraph("_______________________")
                .SetFont(normalFont)
                .SetFontSize(10)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetMarginBottom(3));

            custCell.Add(new Paragraph("Customer Signature")
                .SetFont(boldFont)
                .SetFontSize(9)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontColor(SECONDARY_COLOR));

            signatureTable.AddCell(custCell);

            doc.Add(signatureTable);
        }

        #endregion

        #region Footer

        private void AddFooter(Document doc, PdfFont normalFont)
        {
            doc.Add(new Paragraph("\n"));

            // Terms and Conditions (if any)
            var terms = new Paragraph("Terms & Conditions: Goods once sold will not be taken back. " +
                                    "Payment should be made within the due date to avoid late fees.")
                .SetFont(normalFont)  // Changed from italicFont
                .SetFontSize(8)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontColor(SECONDARY_COLOR)
                .SetMarginBottom(10);
            doc.Add(terms);

            // Powered by section with better styling
            Table footerTable = new Table(UnitValue.CreatePercentArray(new float[] { 1 }))
                .UseAllAvailableWidth();

            Cell footerCell = new Cell()
                .SetBackgroundColor(SECONDARY_COLOR)
                .SetBorder(Border.NO_BORDER)
                .SetPadding(8);

            var footerText = new Paragraph()
                .Add(new Text("Powered by ")
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetFontColor(ColorConstants.WHITE))
                .Add(new Text("C2B Smart App")
                    .SetFont(normalFont)
                    .SetFontSize(9)
                    .SetFontColor(new DeviceRgb(255, 193, 7)))  // Yellow accent
                .Add(new Text(" | Generated on " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm"))
                    .SetFont(normalFont)
                    .SetFontSize(8)
                    .SetFontColor(new DeviceRgb(200, 200, 200)))
                .SetTextAlignment(TextAlignment.CENTER);

            footerCell.Add(footerText);
            footerTable.AddCell(footerCell);

            doc.Add(footerTable);
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