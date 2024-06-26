using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using QRCoder;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace PrintLabelThermal
{
    public partial class Form1 : Form
    {
        private PrintDocument printDocument;
        private PrintPreviewDialog printPreviewDialog;
        private PrintPreviewDialog printPreview = new PrintPreviewDialog();

        private const float CM_TO_INCH = 0.393701f;  // Chuyển đổi từ cm sang inch

        private const string CONFIG_FILE_PATH = "data\\config.ini";

        private List<Order> orders;
        private List<Page> pages;

        private Order selectedOrder;
        private Page selectedPage;

      
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
            PaperSize paperSize = new PaperSize("3x5",
                (int)0,
                (int)0);

            printDocument.DefaultPageSettings.Landscape = false;

            printDocument.DefaultPageSettings.PaperSize = paperSize;
            // Set landscape orientation

            printPreviewControl.Document = printDocument;
            printPreviewControl.Zoom = 1.5; // Điều chỉnh mức phóng to cho xem trước tốt hơn (tùy chọn)

            printPreview.Document = printDocument;
            printPreviewControl.InvalidatePreview(); // Cập nhật xem trước
        }
        private void LoadOrders()
        {
            //đọc dữ liệu từ file json
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\data";
            string filePath = Path.Combine(baseDirectory, "data.json");
            orders = LoadDataFromJson<Order>(filePath);

        }
        private void LoadPageSize()
        {
            //đọc dữ liệu từ file json
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\data";
            string filePath = Path.Combine(baseDirectory, "pages.json");
            pages = LoadDataFromJson<Page>(Path.Combine(baseDirectory, filePath));

        }
        private void InitializeDataGridView()
        {
            // Configure orderDataGridView
            dataGridView.AutoGenerateColumns = false;
            dataGridView.Columns.Add("IdColumn", "ID");
            dataGridView.Columns.Add("NameColumn", "Name");
            dataGridView.Columns.Add("TimeOrderColumn", "Time Order");

            // Populate orderDataGridView
            foreach (var order in orders)
            {
                dataGridView.Rows.Add(order.OrderID, order.OrderName, order.Date);
            }

            // Handle cell click event for order selection
            dataGridView.CellClick += new DataGridViewCellEventHandler(dataGridView_CellClick);

            // Calculate column widths
            int totalWidth = dataGridView.Width;
            int idColumnWidth = totalWidth * 50 / 400; 
            int nameColumnWidth = totalWidth * 200 / 400; 
            int timeOrderColumnWidth = totalWidth * 100 / 400;

            // Set column widths
            this.dataGridView.Columns["IdColumn"].Width = idColumnWidth;
            this.dataGridView.Columns["NameColumn"].Width = nameColumnWidth;
            this.dataGridView.Columns["TimeOrderColumn"].Width = timeOrderColumnWidth;

            //set row style
            this.dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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

            if(sz.Contains("50x30mm"))
            {
                Graphics graphics = e.Graphics;
                graphics.PageUnit = GraphicsUnit.Millimeter; // Set unit to millimeter

                float startX = 5f;   // Starting X position
                float startY = 5f;   // Starting Y position

                // Print order details
                string orderDetails = $"{selectedOrder.OrderID}\nDate: {selectedOrder.Date}\nOrder: {selectedOrder.OrderName}\nQuantity: {selectedOrder.Quantity}\nPrice: {selectedOrder.Price}";
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
                float qrCodeX = 35f;
                float qrCodeY = textSize.Height +2f; // Position below the text with some margin

                graphics.DrawImage(qrCodeImage, qrCodeX, qrCodeY, qrCodeSize, qrCodeSize);
                return;
            }

            if(sz.Contains("57x38mm"))
            {
                Graphics graphics = e.Graphics;
                graphics.PageUnit = GraphicsUnit.Millimeter; // Set unit to millimeter

                float startX = 5f;   // Starting X position
                float startY = 5f;   // Starting Y position

                // Print order details
                string orderDetails = $"{selectedOrder.OrderID}\nDate: {selectedOrder.Date}\nOrder: {selectedOrder.OrderName}\nQuantity: {selectedOrder.Quantity}\nPrice: {selectedOrder.Price}";
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
                string orderName = $"{selectedOrder.OrderName}";
                Font nameFont = new Font("Arial", 9);
                SizeF nameSize = graphics.MeasureString(orderName, nameFont);
                float nameX = startX + 5f;
                float nameY = idY + idSize.Height + 3f;  // Below id, with 5mm spacing
                graphics.DrawString(orderName, nameFont, Brushes.Black, nameX, nameY);

                // quanlity
                string quanlity = $"{selectedOrder.Quantity}";
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
                return;

            }


        }
        #endregion


        #region extenstion
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
            // Store the selected order in a field to be used in PrintOrderPage
            selectedOrder = order;
            // Optionally, update the preview control if necessary
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
                // Định dạng lại thời gian theo định dạng mong muốn, ví dụ: "dd-MM-yyyy HH:mm:ss"
                    return parsedDateTime.ToString("dd/MM");
                if (type == 1)
                    return parsedDateTime.ToString("HH:mm");
            }
            return timeString; // Nếu không thể phân tích thời gian, trả về chuỗi gốc
        }




        #endregion

     
    }
}
