using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Printing;
using QRCoder;
using System.IO;
using Newtonsoft.Json;
using System.Text;
using System.Runtime.InteropServices;

namespace PrintLabelThermal
{
    public partial class Form1 : Form
    {

        #region khai báo biến
        private const float CM_TO_INCH = 0.393701f;  // Chuyển đổi từ cm sang inch
        private const string CONFIG_FILE_PATH = "data\\config.ini";

        private PrintDocument printDocument;
        private PrintPreviewDialog printPreviewDialog;
        private PrintPreviewDialog printPreview = new PrintPreviewDialog();

        private List<Order> orders;
        private List<Page> pages;
        private Order selectedOrder;
        private Page selectedPage;


        #endregion

        public Form1()
        {
            InitializeComponent();
            InitializePrintComponents();
            LoadOrders();
            LoadPageSize();
            InitializeDataGridView();
            InitializePageSizeComboBox();
        }


        #region khởi tạo và load data
        private void InitializePrintComponents()
        {
            printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintOrderPage);


            // Set paper size
            PaperSize paperSize = new PaperSize("customPaper",0,0);

            printDocument.DefaultPageSettings.Landscape = false;
            printDocument.DefaultPageSettings.PaperSize = paperSize;

            printPreviewControl.Document = printDocument;
            printPreviewControl.Zoom = 1.5; // Điều chỉnh mức phóng to cho xem trước tốt hơn (tùy chọn)

            printPreview.Document = printDocument;
            printPreviewControl.InvalidatePreview(); // Cập nhật xem trước
        }
        private void LoadOrders()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\data";
            string filePath = Path.Combine(baseDirectory, "data.json");
            orders = LoadDataFromJson<Order>(filePath);
        }
        private void LoadPageSize()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\data";
            string filePath = Path.Combine(baseDirectory, "pages.json");
            pages = LoadDataFromJson<Page>(Path.Combine(baseDirectory, filePath));

        }
        private void InitializeDataGridView()
        {
            // Configure orderDataGridView
            dataGridView.AutoGenerateColumns = false;
            dataGridView.Columns.Add("IdColumn", "ID");
            dataGridView.Columns.Add("DateColumn", "Date");
            dataGridView.Columns.Add("OrdersColumn", "Orders");
            dataGridView.Columns.Add("NotesColumn", "Notes");
            dataGridView.Columns.Add("TotalPriceColumn", "Total Price");

            // Populate orderDataGridView
            foreach (var order in orders)
            {
                dataGridView.Rows.Add(order.OrderID, 
                    order.Date,
                    order.Orders != null && order.Orders.Length > 0 ? (order.Orders.Length > 1 ? string.Join(", ", order.Orders) : order.Orders[0]) : "",
                    order.Notes != null && order.Notes.Length > 0 ? (order.Notes.Length > 1 ? string.Join(", ", order.Notes) : (string.IsNullOrEmpty(order.Notes[0]) ? "" : order.Notes[0])) : "",
                    order.TotalPrice);
            }

            // Select the first row by default
            if (dataGridView.Rows.Count > 0)
            {
                dataGridView.CurrentCell = dataGridView.Rows[0].Cells[0];
                dataGridView.Rows[0].Selected = true;

                // Trigger the cell click event to load the first order
                DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(0, 0);
                dataGridView_CellClick(this.dataGridView, args);
            }


            // Handle cell click event for order selection
            dataGridView.CellClick += new DataGridViewCellEventHandler(dataGridView_CellClick);

            // Calculate column widths
            int totalWidth = dataGridView.Width;
            int idColumnWidth = totalWidth * 50 / totalWidth;
            int dateColumnWidth = totalWidth * 100 / totalWidth;
            int ordersColumnWidth = totalWidth * 300 / totalWidth;
            int notesColumnWidth = totalWidth * 300 / totalWidth;
            int totalPriceColumnWidth = totalWidth * 60 / totalWidth;

            // Set column widths
            dataGridView.Columns["IdColumn"].Width = idColumnWidth;
            dataGridView.Columns["DateColumn"].Width = dateColumnWidth;
            dataGridView.Columns["OrdersColumn"].Width = ordersColumnWidth;
            dataGridView.Columns["NotesColumn"].Width = notesColumnWidth;
            dataGridView.Columns["TotalPriceColumn"].Width = totalPriceColumnWidth;

            // Set row style
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        private void InitializePageSizeComboBox()
        {
            pageSizeComboBox.SelectedIndexChanged += new EventHandler(pageSizeComboBox_SelectedIndexChanged);
            this.Controls.Add(pageSizeComboBox);
            foreach (var page in pages)
            {
                pageSizeComboBox.Items.Add(page.PageSizeName);
            }

            // Read pageSize from config.ini
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\data";
            string filePath = Path.Combine(baseDirectory, "config.ini");
            string pageSizeConfig = ReadConfigValue(filePath, "pageSize");

            // Set selected index based on pageSizeConfig
            if (!string.IsNullOrEmpty(pageSizeConfig))
            {
                Page selectedPage = pages.Find(p => p.id ==int.Parse(pageSizeConfig));
                if (selectedPage != null)
                {
                    int selectedIndex = pages.IndexOf(selectedPage);
                    pageSizeComboBox.SelectedIndex = selectedIndex;
                }
            }
        }
        #endregion

        #region handle button
        private void btnPrint_Click(object sender, EventArgs e)
        {
            //printPreview.ShowDialog();
            printDocument.Print();
        }
        private void orderListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox listBox = sender as ListBox;
            if (listBox.SelectedItem != null)
            {
                Order selectedOrder = listBox.SelectedItem as Order;
                UpdatePrintOrderPage(selectedOrder);
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Đảm bảo rằng khi ứng dụng đóng, giá trị pageSize đã được lưu lại vào config.ini
            if (pageSizeComboBox.SelectedItem != null)
            {
                string selectedPageSize = pageSizeComboBox.SelectedItem.ToString();
                Page selectedPage = pages.Find(p => p.PageSizeName == selectedPageSize);
                SaveConfigValue("pageSize", selectedPage.id.ToString());
            }
        }
        private void pageSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedPageSizeName = pageSizeComboBox.SelectedItem.ToString();
            var selectedPage = pages.FirstOrDefault(p => p.PageSizeName == selectedPageSizeName);
            if (selectedPage != null)
            {
                UpdatePaperSize(selectedPage);
            }
        }
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Handle cell click event for order selection
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView.Rows[e.RowIndex];
                string orderId = row.Cells["IdColumn"].Value.ToString();
                Order selectedOrder = orders.Find(o => o.OrderID == orderId);
                if (selectedOrder != null)
                {
                    UpdatePrintOrderPage(selectedOrder);
                }
            }
        }
        #endregion

        #region print 
        private void PrintOrderPage(object sender, PrintPageEventArgs e)
        {
            if (selectedOrder == null) return;
            
            string sz = selectedPage.PageSizeName;
            
            if (sz.Contains("50x30mm"))
            {
                Graphics graphics = e.Graphics;
                graphics.PageUnit = GraphicsUnit.Millimeter; 

                float sX = 3f;   // Starting X position
                float sY = 3f;   // Starting Y position

                const float pH = 30; 
                const float pW = 50;

                float fsz = 7; //font size

                Font mfont = new Font("Arial", fsz);

                //render id
                string orderID = selectedOrder.OrderID;
                graphics.DrawString(orderID, mfont, Brushes.Black, sX, sY);

                //render date
                string Date = FormatTime(selectedOrder.Date,0) + " " + FormatTime(selectedOrder.Date, 1);
                SizeF FontSize = graphics.MeasureString(Date, mfont);
                graphics.DrawString(Date, mfont, Brushes.Black,pW - sX - FontSize.Width, sY);
                sY += (float)FontSize.Height + 1;

                //render order
                bool check_num = selectedOrder.Orders.Length + selectedOrder.Notes.Length > 4;
                for (int i = 0; i < selectedOrder.Orders.Length; i++)
                {

                    if (check_num)
                    {
                        mfont = new Font("Arial", 5);
                    }

                    string orderName = selectedOrder.Orders[i].Split('X')[0];
                    SizeF orderNameSize = graphics.MeasureString(orderName, mfont);
                    float num = sX + orderNameSize.Width;
                    float posX =  pW - sX - 3;
                    if (num > posX)
                    {
                        List<string> lines = new List<string>();
                        StringBuilder currentLine = new StringBuilder();
                        foreach (char c in orderName)
                        {
                            currentLine.Append(c);
                            SizeF currentLineSize = graphics.MeasureString(currentLine.ToString(), mfont);
                            if (sX + currentLineSize.Width > posX)
                            {
                                lines.Add(currentLine.ToString(0, currentLine.Length - 1));
                                currentLine.Clear();
                                currentLine.Append(c);
                            }
                        }
                        if (currentLine.Length > 0)
                        {
                            lines.Add(currentLine.ToString());
                        }

                        float currentY = sY;
                        foreach (string line in lines)
                        {
                            graphics.DrawString(line, mfont, Brushes.Black, sX, currentY);
                            currentY += check_num ? 1.7f : (float)FontSize.Height;
                        }
                        graphics.DrawString(selectedOrder.Orders[i].Split('X')[1], mfont, Brushes.Black, posX, sY);
                        sY += (check_num ? 1.7f : (float)FontSize.Height) * lines.Count;

                    }
                    else
                    {
                        graphics.DrawString(selectedOrder.Orders[i].Split('X')[0], mfont, Brushes.Black, sX, sY);
                        graphics.DrawString(selectedOrder.Orders[i].Split('X')[1], mfont, Brushes.Black, posX, sY);
                        sY += check_num ? 1.7f : (float)FontSize.Height;
                    }
                }

                //render notes
                string[] notes = selectedOrder.Notes;
                if(notes.Length > 0)
                {
                    graphics.DrawString("- note: ", mfont, Brushes.Black, sX, sY);
                    sY += check_num ? 1.7f : (float)FontSize.Height;
                }
                   

                for (int i = 0; i < selectedOrder.Notes.Length; i++)
                {

                    if (check_num)
                        mfont = new Font("Arial", 5);

                    string orderName = selectedOrder.Orders[i].Split('X')[0];
                    SizeF orderNameSize = graphics.MeasureString(orderName, mfont);
                    float num = sX + orderNameSize.Width;
                    float posX = pW  - 13;
                    if (num > posX)
                    {
                        List<string> lines = new List<string>();
                        StringBuilder currentLine = new StringBuilder();
                        foreach (char c in orderName)
                        {
                            currentLine.Append(c);
                            SizeF currentLineSize = graphics.MeasureString(currentLine.ToString(), mfont);
                            if (sX + currentLineSize.Width > posX)
                            {
                                lines.Add(currentLine.ToString(0, currentLine.Length - 1));
                                currentLine.Clear();
                                currentLine.Append(c);
                            }
                        }
                        if (currentLine.Length > 0)
                        {
                            lines.Add(currentLine.ToString());
                        }

                        float currentY = sY;
                        foreach (string line in lines)
                        {
                            graphics.DrawString(line, mfont, Brushes.Black, sX+ 3, sY);
                            sY += check_num ? 1.7f : (float)FontSize.Height;
                        }

                    }
                    else
                    {
                        graphics.DrawString(selectedOrder.Notes[i], mfont, Brushes.Black, sX+5, sY);
                        sY += check_num ? 1.7f : (float)FontSize.Height;
                    }
                }




                // Generate and print QR code
                string qrData = $"{orderID} {Date}";

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                float qrCodeSize = 100 * 0.1f;
                float qrCodeX = pW - qrCodeSize - 2f;
                float qrCodeY = pH - qrCodeSize;

                graphics.DrawImage(qrCodeImage, qrCodeX, qrCodeY, qrCodeSize, qrCodeSize);

                return;
            }

            if(sz.Contains("57x38mm"))
            {
                Graphics graphics = e.Graphics;
                graphics.PageUnit = GraphicsUnit.Millimeter; 

                float startX = 5f;   
                float startY = 5f;   

                // Print order details
                //!bug
                string orderDetails = $"{selectedOrder.OrderID}\nDate: {selectedOrder.Date}\nOrder: {selectedOrder.Orders.ToString()}\nQuantity: {5}\nPrice: {selectedOrder.TotalPrice}";
                Font font = new Font("Arial", 7);
                SizeF textSize = graphics.MeasureString(orderDetails, font);
                float textX = startX;
                float textY = startY;
                graphics.DrawString(orderDetails, font, Brushes.Black, textX, textY);

                // Generate and print QR code
                string qrData = orderDetails;
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                float qrCodeSize = 100 * 0.1f; // QR code size relative to page size
                float qrCodeX = 40f;
                float qrCodeY = textSize.Height+7f; // Position below the text with some margin

                graphics.DrawImage(qrCodeImage, qrCodeX, qrCodeY, qrCodeSize, qrCodeSize);
                return;
            }

            if (sz.Contains("80x45mm"))
            {
                Graphics graphics = e.Graphics;
                graphics.PageUnit = GraphicsUnit.Millimeter; // Set unit to millimeter

                const float pageW = 80;
                const float pageH = 45;

                float startX = 5f;   // Margin from left
                float startY = 5f;   // Margin from top

                string timeOrder = $"{FormatTime(selectedOrder.Date,0)}  {FormatTime(selectedOrder.Date, 1)}";
                Font timeFont = new Font("Arial", 7);
                SizeF timeSize = graphics.MeasureString(timeOrder, timeFont);
                float timeX = startX + 5f;   // Right aligned
                float timeY = startY;  // Same vertical position as order name
                graphics.DrawString(timeOrder, timeFont, Brushes.Black, timeX, timeY);

                // Print order ID (centered horizontally)
                string orderId = $"{selectedOrder.OrderID}";
                Font idFont = new Font("Arial", 10, FontStyle.Bold);
                SizeF idSize = graphics.MeasureString(orderId, idFont);
                float idX =(pageW  - idSize.Width)/2; // Center horizontally
                float idY = startY + 2f;  // 5mm from top margin
                graphics.DrawString(orderId, idFont, Brushes.Black, idX, idY);
                
                //line
                float lineY = startY +  idSize.Height + 3f; // Ví dụ: vị trí y để vẽ đường line
                float lineWidth = pageW - 30; // Chiều rộng của đường line

                using (Pen pen = new Pen(Color.Black, 0.1f))
                {
                    e.Graphics.DrawLine(pen, new Point(10, (int)lineY), new Point((int)(20 + lineWidth), (int)lineY));
                }
                //orderName
                string orderName = $"{selectedOrder.Orders.ToString()}";
                Font nameFont = new Font("Arial", 9);
                SizeF nameSize = graphics.MeasureString(orderName, nameFont);
                float nameX = startX + 5f;
                float nameY = idY + idSize.Height + 3f;  // Below id, with 5mm spacing
                graphics.DrawString(orderName, nameFont, Brushes.Black, nameX, nameY);

                // quanlity
                //!bug
                string quanlity = $"{1}";
                Font quanlityFont = new Font("Arial", 9);
                SizeF quanSize = graphics.MeasureString(quanlity, quanlityFont);

                float quanX = pageW - 10f - quanSize.Width;
                graphics.DrawString(quanlity, nameFont, Brushes.Black, quanX, nameY);


                //line
                float lineY2 = pageH - 15f; // Ví dụ: vị trí y để vẽ đường line
                using (Pen pen = new Pen(Color.Black, 0.1f))
                {
                    e.Graphics.DrawLine(pen, new Point(10, (int)lineY2), new Point((int)(20 + lineWidth), (int)lineY2));
                }  
                
                
                // Generate and print QR code
                string qrData = $"{selectedOrder.OrderID} {selectedOrder.Orders.ToString()} {selectedOrder.Date}";
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                float qrCodeSize = 12f; // QR code size relative to page size
                float qrCodeX = pageW - 10f - qrCodeSize;  // Center horizontally
                float qrCodeY = pageH - 13f;  // 5mm from bottom margin
                graphics.DrawImage(qrCodeImage, qrCodeX, qrCodeY, qrCodeSize, qrCodeSize);
                return;

            }


        }
        #endregion

        #region extenstions
        private string ReadConfigValue(string filePath, string key)
        {
            // Example method to read configuration value from config.ini
            string value = string.Empty;
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    string[] parts = line.Split('=');
                    if (parts.Length == 2 && parts[0].Trim() == key)
                    {
                        value = parts[1].Trim();
                        break;
                    }
                }
            }
            return value;
        }
        private void SaveConfigValue(string key, string value)
        {
            // Phương thức lưu giá trị vào config.ini
            try
            {
                using (StreamWriter sw = new StreamWriter(CONFIG_FILE_PATH))
                {
                    sw.WriteLine($"{key}={value}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu cấu hình: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdatePaperSize(Page page)
        {
            selectedPage = page;
            float pageWidth = page.w * CM_TO_INCH * 100f; // Chuyển đổi cm sang inch và sau đó sang millimet
            float pageHeight = page.h * CM_TO_INCH * 100f; // Chuyển đổi cm sang inch và sau đó sang millimet

            PaperSize paperSize = new PaperSize(page.PageSizeName, (int)pageWidth, (int)pageHeight);
            printDocument.DefaultPageSettings.PaperSize = paperSize;
            printPreviewControl.InvalidatePreview();
        }
        private void UpdatePrintOrderPage(Order order)
        {
            selectedOrder = order;
            printPreviewControl.InvalidatePreview();
        }
        private List<T> LoadDataFromJson<T>(string filePath)
        {
            using (StreamReader r = new StreamReader(filePath))
            {
                string json = r.ReadToEnd();
                return JsonConvert.DeserializeObject<List<T>>(json);
            }
        }
        private string FormatTime(string timeString,int type)
        {
            DateTime parsedDateTime;
            if (DateTime.TryParse(timeString, out parsedDateTime))
            {
                if(type == 0)
                    return parsedDateTime.ToString("dd/MM");
                if (type == 1)
                    return parsedDateTime.ToString("HH:mm");
            }
            return timeString;
        }
        #endregion

  


        
    }
}
