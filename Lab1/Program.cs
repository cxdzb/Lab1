using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Gma.QrCodeNet.Encoding;
using Gma.QrCodeNet.Encoding.Windows.Render;
using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Lab1
{
    class Program
    {
        // Creating Encoder and Converting String to QrCode
        public QrCode StringToQrCode(string str)
        {
            QrEncoder qrencoder = new QrEncoder(ErrorCorrectionLevel.M);    ///// Definition coder
            QrCode qrcode = qrencoder.Encode(str);    ///// Definition qrcode
            return qrcode;    /////Return qrcode
        }

        // Output two-dimensional code
        public void PrintQrCode(QrCode qrcode)
        {
            for (int i = 0; i < qrcode.Matrix.Width; i++)    ///// Traveling through each point
            {
                for (int j = 0; j < qrcode.Matrix.Width; j++)
                    Console.Write(qrcode.Matrix[j, i] ? '　' : '█');    ///// Output black and white squares
                Console.WriteLine();
            }
        }

        // Generating PNG pictures
        public void DumpPng(QrCode qrcode,int row,string name="LengthBelow4")
        {
            const int modelSizeInPixels = 16;
            GraphicsRenderer render = new GraphicsRenderer(new FixedModuleSize(modelSizeInPixels, QuietZoneModules.Two), Brushes.Black, Brushes.White);// Define the render

            DrawingSize dSize = render.SizeCalculator.GetSize(qrcode.Matrix.Width);    ///// Get the size of QrCode
            Bitmap map = new Bitmap(dSize.CodeWidth, dSize.CodeWidth);    ///// Definition qrcode bitmap
            Graphics g = Graphics.FromImage(map);    ///// definition diagram
            render.Draw(g, qrcode.Matrix);    ///// Rendering pictures

            Bitmap background = (Bitmap)Image.FromFile(@"..\..\..\resource\background.jpg");    ///// Define background bitmap
            Graphics gh = Graphics.FromImage(background);    ///// Define Background Map
            Point qrcodePoint = new Point((background.Width - 400) / 2, (background.Height - 400) / 2);    ///// Define the upper left corner position of qrcode
            gh.FillRectangle(Brushes.Green, qrcodePoint.X-5, qrcodePoint.Y-5, 410, 410);    ///// Fill in a green area
            gh.DrawImage(map, qrcodePoint.X, qrcodePoint.Y, 400, 400);    ///// Draw qrcode in the green area

            Image img = Image.FromFile(@"..\..\..\resource\logo.png");    ///// Load logo images
            Point imgPoint = new Point((background.Width - img.Width/4) / 2, (background.Height - img.Height/4) / 2);    ///// Define the starting point coordinates of the upper left corner of logo
            gh.DrawImage(img, imgPoint.X, imgPoint.Y, img.Width/4, img.Height/4);    ///// Draw logo in the center of the picture

            string fileNumber = row.ToString();    ///// Line Number Converted to String
            while (fileNumber.Length < 3) fileNumber = "0" + fileNumber;    ///// Fill forward 0 to 3 bits

            background.Save(@"..\..\..\results\" + fileNumber + "-" + name + ".png", ImageFormat.Png);    ///// Save as PNG picture
        }

        // Read MySQL database
        public List<string> ReadMysql(string table)
        {
            string connStr = "Database=mydata;datasource=192.168.142.130;port=3306;user=lsp;pwd=1005968086;";
            MySqlConnection conn = new MySqlConnection(connStr);     ///// Connect to the database
            conn.Open();    ///// Open the connection
            MySqlCommand cmd = new MySqlCommand("select * from "+table, conn);
            MySqlDataReader reader = cmd.ExecuteReader();     ///// Create commands and execute them
            List<string> qrcodes = new List<string>();     ///// Declare storage QRcode container
            while (reader.Read())     ///// Read data line by line and store it
                qrcodes.Add(reader.GetString("code"));
            reader.Close();     ///// Close the connection
            return qrcodes;
        }

        // Read Excel tables
        public List<string> ReadExcel(string path)
        {
            IWorkbook workbook;    ///// Define the workbook to save the data of excel
            string fileExt = Path.GetExtension(path).ToLower();    ///// Get the Extension(.xls/.xlsx)
            List<string> contents = new List<string>();    ///// Create a list to save every line
            path = Path.GetFullPath(path);
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))    ///// Open the file
            {
                if (fileExt == ".xlsx") workbook = new XSSFWorkbook(fs);    ///// Judge which format the file is
                else if (fileExt == ".xls") workbook = new HSSFWorkbook(fs);
                else workbook = null;
                if (workbook == null) return null;

                ISheet sheet = workbook.GetSheetAt(0);    ///// Get sheet1
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                    contents.Add(sheet.GetRow(i).GetCell(0).ToString());     ///// Read data line by line and store it
            }
            return contents;
        }

        // main program
        static void Main(string[] args)
        {
            // Control input length
            if (args.Length == 1 && args[0].Length <= 64)
            {
                Program program = new Program();
                string str = args[0];
                // If there is a -f parameter, open the txt file
                if (str.StartsWith("-f"))
                    // Open and read files
                    using (StreamReader file = new StreamReader(args[0].Substring(2)))
                    {
                        // Read and display rows from the file until the end of the file
                        string line;
                        int row = 0;
                        while ((line = file.ReadLine()) != null)
                        {
                            // Describe line as the line of the file
                            row++;
                            // If the number of characters is less than 4
                            if (line.Length < 4) program.DumpPng(program.StringToQrCode(line), row);
                            // If the number of characters is greater than or equal to 4
                            else program.DumpPng(program.StringToQrCode(line), row, line.Substring(0, 4));
                        }
                    }
                // If there is a -m parameter, open MySQL
                else if (str.StartsWith("-m"))
                {
                    // Connect and read the database
                    List<string> qrcodes = program.ReadMysql(args[0].Substring(2));
                    // read by line
                    int row = 0;
                    foreach (string qrcode in qrcodes)
                    {
                        // Describe QRcode as the row in the table
                        row++;
                        // If the number of characters is less than 4
                        if (qrcode.Length < 4) program.DumpPng(program.StringToQrCode(qrcode), row);
                        // If the number of characters is greater than or equal to 4
                        else program.DumpPng(program.StringToQrCode(qrcode), row, qrcode.Substring(0, 4));
                    }
                }
                // If there is a -e parameter, open the Excel
                else if (str.StartsWith("-e"))
                {
                    // Read the excel
                    List<string> qrcodes = program.ReadExcel(args[0].Substring(2));
                    // read by line
                    int row = 0;
                    foreach (string qrcode in qrcodes)
                    {
                        // Describe QRcode as the row in the table
                        row++;
                        // If the number of characters is less than 4
                        if (qrcode.Length < 4) program.DumpPng(program.StringToQrCode(qrcode), row);
                        // If the number of characters is greater than or equal to 4
                        else program.DumpPng(program.StringToQrCode(qrcode), row, qrcode.Substring(0, 4));
                    }
                }
                // If there is no - parameter
                else program.PrintQrCode(program.StringToQrCode(str));
            }
            // Input Format Error
            else if (args.Length != 1) Console.WriteLine("The number of arg is too many!");
            else Console.WriteLine("The length of arg is too long!");
        }
    }
}
