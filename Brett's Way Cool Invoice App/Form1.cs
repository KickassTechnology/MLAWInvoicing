using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;
using PdfSharp.Pdf.Printing;
using System.IO;

namespace Brett_s_Way_Cool_Invoice_App
{
    public partial class Form1 : Form
    {
        public String strAcrobat = "";
        //Path to the databse
        public String strConn = "Server=mlawdb.cja22lachoyz.us-west-2.rds.amazonaws.com;Database=MLAW_MS;User Id=sa;Password=!sd2the2power!;";

        public Form1()
        {
            //This is maybe the only straightforward project we did and it should be dead simple to understand


            InitializeComponent();
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "X";
            checkColumn.HeaderText = "X";
            checkColumn.Width = 50;
            checkColumn.ReadOnly = false;
            checkColumn.FillWeight = 10; //if the datagridview is resized (on form resize) the checkbox won't take up too much; value is relative to the other columns' fill values
           
            dataGridView1.Columns.Add(checkColumn);
            loadData();
            
        }

        //loads our UI with everything that is marked with a status of "Delivered" - that status id is 10
        public void loadData()
        {
            
            dataGridView1.DataSource = null;

            
            using (SqlConnection conn = new SqlConnection(strConn))
            {

                conn.Open();

                String strSql = "Get_Foundation_Invoiceable";

                SqlCommand sqlComm = new SqlCommand(strSql, conn);
                sqlComm.CommandType = CommandType.Text;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = sqlComm;

                DataSet ds = new DataSet();
                da.Fill(ds);

                strSql = "Get_Foundation_Revisions_Invoiceable";

                sqlComm = new SqlCommand(strSql, conn);
                sqlComm.CommandType = CommandType.Text;

                da = new SqlDataAdapter();
                da.SelectCommand = sqlComm;

                DataSet ds2 = new DataSet();
                da.Fill(ds2);

                ds.Merge(ds2.Tables[0]);

                dataGridView1.AutoGenerateColumns = true;
                dataGridView1.DataSource = ds; // dataset
                dataGridView1.DataMember = ds.Tables[0].TableName;

            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //This is our check all 
            if (checkBox1.Checked == true)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                    chk.Value = true;
                } 
            }
            else
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                    chk.Value = false;
                } 
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Get our path to Acrobat
            PdfFilePrinter.AdobeReaderPath = strAcrobat.Replace("\\", "\\\\");

            // Present a Printer settings dialog to the user so the can select the printer
            // to use.
            PrinterSettings settings = new PrinterSettings();
            settings.Collate = false;
            PrintDialog printerDialog = new PrintDialog();
            printerDialog.AllowSomePages = false;
            printerDialog.ShowHelp = false;
            printerDialog.PrinterSettings = settings;
            printerDialog.AllowPrintToFile = true;
            printerDialog.PrinterSettings.PrintToFile = true;
            DialogResult result = printerDialog.ShowDialog();

            //If the user doesn't cancel, do something
            if (result == DialogResult.OK)
            {
                //Loop through the DataGrid
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    //If the row is checked, let's work with it.
                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                    if (Convert.ToBoolean(chk.Value) == true)
                    {

                        Int32 iOrderId = Convert.ToInt32(row.Cells[1].Value.ToString());

                        using (SqlConnection conn = new SqlConnection(strConn))
                        {
                            conn.Open();
                            DataSet ds = new DataSet();
                            
                            //Get info about the Order
                            SqlCommand sqlComm = new SqlCommand("Get_Order_By_Id", conn);
                            sqlComm.Parameters.AddWithValue("@Order_Id", iOrderId);
                            sqlComm.CommandType = CommandType.StoredProcedure;

                            SqlDataAdapter da = new SqlDataAdapter();
                            da.SelectCommand = sqlComm;

                            da.Fill(ds);
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                DataSet dsInvoiceNum = new DataSet();
                                
                                //Returns in invoice number 
                                SqlCommand sqlCommInv = new SqlCommand("Generate_Invoice", conn);
                                sqlCommInv.Parameters.AddWithValue("@Order_Id", iOrderId);
                                sqlCommInv.CommandType = CommandType.StoredProcedure;

                                SqlDataAdapter daInv = new SqlDataAdapter();
                                daInv.SelectCommand = sqlCommInv;

                                daInv.Fill(dsInvoiceNum);

                                int iInvoiceNum = Convert.ToInt32(dsInvoiceNum.Tables[0].Rows[0][0]);

                                DataRow dr = ds.Tables[0].Rows[0];

                                //Create the PDF
                                PdfDocument document = new PdfDocument();
                                document.Info.Author = "MLAW Engineers";
                                document.Info.Keywords = "";

                                PdfPage page = document.AddPage();
                                page.Size = PdfSharp.PageSize.A4;

                                // Obtain an XGraphics object to render to
                                XGraphics gfx = XGraphics.FromPdfPage(page);

                                // Create a font
                                double fontHeight = 12;
                                XFont font = new XFont("Times New Roman", fontHeight, XFontStyle.BoldItalic);

                                // Get the centre of the page
                                double y = 20;
                                int lineCount = 0;
                                double linePadding = 10;

                                // Create a rectangle to draw the text in and draw in it
                                XRect rect = new XRect(0, y, page.Width, fontHeight);
                                
                                //This is all PDF formatting/printing/layout
                                lineCount++;
                                y += fontHeight;

                                String imageLoc = "mlaw_logo.png";
                                DrawImage(gfx, imageLoc, 28, 20, 111, 28);
                                
                                PointF pt = new PointF(144, 48);
                                XFont fontEng = new XFont("Arial", 15, XFontStyle.Regular);
                                gfx.DrawString("ENGINEERS" , fontEng, XBrushes.Navy, pt);

                                XPen pen = new XPen(XColors.Black, 2);
                                gfx.DrawLine(pen, 20, 50, 580, 50);

                                pt = new PointF(24, 61);
                                XFont fontServLine = new XFont("Arial Narrow", 9, XFontStyle.Regular);
                                gfx.DrawString("FOUNDATION | FRAMING | INSPECTIONS | ENERGY | GEOSTRUCTURAL", fontServLine, XBrushes.Black, pt);

                                pt = new PointF(520, 60);
                                XFont fontPE = new XFont("Arial", 8, XFontStyle.Regular);
                                gfx.DrawString("TX PE #002685", fontPE, XBrushes.Black, pt);
                                

                                XFont fontNormal = new XFont("Arial", 10, XFontStyle.Regular);
                                pt = new PointF(40, 180);
                                gfx.DrawString(dr["Client_Full_Name"].ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(40, 191);
                                gfx.DrawString(dr["Billing_Address_1"].ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(40, 202);
                                gfx.DrawString(dr["Billing_City"].ToString() + ", " + dr["Billing_State"].ToString() + " " + dr["Billing_Postal_Code"].ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(340, 180);

                                DateTime dtNow = DateTime.Now;


                                gfx.DrawString(dtNow.ToString("MMMM dd, yyyy"), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(340, 191);
                                gfx.DrawString("Invoice Number: " + iInvoiceNum.ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(340, 202);

                                /* Removed per Janet
                                gfx.DrawString("CC: 001240", fontNormal, XBrushes.Black, pt);
                                pt = new PointF(340, 223);
                                gfx.DrawString("Comments:", fontNormal, XBrushes.Black, pt);
                                pt = new PointF(340, 234);
                                gfx.DrawString(dr["Comments"].ToString(), fontNormal, XBrushes.Black, pt);
                                 * */

                                XFont fontBold = new XFont("Arial", 10, XFontStyle.Bold);
                                pt = new PointF(40, 250);
                                gfx.DrawString("Address: " + dr["Address"].ToString(), fontBold, XBrushes.Black, pt);
                                pt = new PointF(91, 261);
                                gfx.DrawString("Lot: " + dr["Lot"].ToString(), fontBold, XBrushes.Black, pt);
                                pt = new PointF(170, 261);
                                gfx.DrawString("Block: " + dr["Block"].ToString(), fontBold, XBrushes.Black, pt);
                                pt = new PointF(91, 272);
                                gfx.DrawString("Subdivision: " + dr["Subdivision_Name"].ToString(), fontBold, XBrushes.Black, pt);


                                pt = new PointF(40, 300);
                                gfx.DrawString("Engineers Project No: " + dr["MLAW_Number"].ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(40, 311);

                                /*Removed Per Janet
                                gfx.DrawString("Date Received: " + dr["Received_Date_String"].ToString(), fontNormal, XBrushes.Black, pt);
                                pt = new PointF(40, 322);
                                gfx.DrawString("Date Delivered: ", fontNormal, XBrushes.Black, pt);
                                */

                                //Figure out what to bill the customer

                                Decimal dAmount = 0;
                                Decimal dDiscount = 0;
                                Decimal dTotal = 0;
                                Decimal number;

                                if (Decimal.TryParse(dr["Amount"].ToString(), out number))
                                {
                                    dAmount = number;
                                }

                                if (Decimal.TryParse(dr["Discount"].ToString(), out number))
                                {
                                    dDiscount = number;
                                }

                                dTotal = dAmount - dDiscount;

                                PointF pt1 = new PointF(40, 400);
                                PointF pt2 = new PointF(540, 400);
                                gfx.DrawLine(XPens.Black, pt1, pt2);

                                pt = new PointF(40, 430);
                                gfx.DrawString("Foundation Design Services: ", fontNormal, XBrushes.Black, pt);

                                pt = new PointF(40, 460);
                                gfx.DrawString("Base Charge: ..........................................................................................................................", fontNormal, XBrushes.Black, pt);

                                pt = new PointF(500, 460);
                                gfx.DrawString(dAmount.ToString("C2"), fontNormal, XBrushes.Black, pt);

                                if (dDiscount > 0)
                                {
                                    pt = new PointF(500, 480);
                                    gfx.DrawString(dDiscount.ToString("C2"), fontNormal, XBrushes.Black, pt);

                                }
                                pt = new PointF(140, 480);
                                String strSqFt = "X";

                                /* Removed per Janet
                                if (dSlabSqFt > 0 || iLevel1 > 0)
                                {
                                    strSqFt = Convert.ToInt32(dSlabSqFt).ToString();
                                }


                                gfx.DrawString("(         Sq Ft:  " + strSqFt + "             )", fontNormal, XBrushes.Black, pt);
                                

                                pt = new PointF(140, 500);
                                gfx.DrawString("Copies of Document Included", fontNormal, XBrushes.Black, pt);
                                */

                                pt1 = new PointF(40, 540);
                                pt2 = new PointF(540, 540);
                                gfx.DrawLine(XPens.Black, pt1, pt2);

                                pt = new PointF(300, 600);
                                gfx.DrawString("Total Charge:", fontNormal, XBrushes.Black, pt);

                                pt = new PointF(500, 600);
                                gfx.DrawString(dTotal.ToString("C2"), fontNormal, XBrushes.Black, pt);
                                
                                /* Remove per Janet
                                pt = new PointF(300, 530);
                                gfx.DrawString("Approved:     ____________________________", fontNormal, XBrushes.Black, pt);
                                */

                                pt = new PointF(60, 830);
                                XFont fontSmall = new XFont("Arial", 6, XFontStyle.Regular);
                                gfx.DrawString("2804 LONGHORN BLVD.", fontSmall, XBrushes.Black, pt);

                                pt = new PointF(182, 830);
                                gfx.DrawString("AUSTIN, TX 78758", fontSmall, XBrushes.Black, pt);

                                pt = new PointF(287, 830);
                                gfx.DrawString("512.835.7000", fontSmall, XBrushes.Black, pt);

                                pt = new PointF(380, 830);
                                gfx.DrawString("FAX 512.835.4975", fontSmall, XBrushes.Black, pt);

                                pt = new PointF(480, 830);
                                gfx.DrawString("TOLL FREE 877.855.3041", fontSmall, XBrushes.Black, pt);

                                //The invoice is Save in an Invoices directory
                                
                                /****** MUST CREATE THIS DIRECTORY AHEAD OF TIME - REFACTOR TO AUTO-CREATE IF NOT THERE *****/
                                String strFileName = "Invoices/Invoice_" + dr["MLAW_Number"].ToString() + ".pdf";
                                document.Save(strFileName);

                                try
                                {
                                    FileInfo f = new FileInfo(strFileName);
                                    string fullname = f.FullName;
                                    Process process = new Process();
                                    
                                    // pdf file to print 
                                    process.StartInfo.FileName = fullname;

                                    //print to specified printer
                                    process.StartInfo.Verb = "Print";
                                    process.StartInfo.CreateNoWindow = true;

                                    process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                                    process.StartInfo.UseShellExecute = true;
                                    

                                    //Printer name
                                    process.StartInfo.Arguments = settings.PrinterName;
                                    process.Start();
                                    process.WaitForExit(3000);

                                    //Start Acrobat and print it out
                                    Process[] procs = Process.GetProcessesByName("Acrobat");

                                    foreach (Process proc in procs)
                                    {
                                       if (!process.HasExited)
                                           {
                                              proc.Kill();
                                              proc.WaitForExit();
                                            }
                                     }

                                     process.Close();

                                    
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Error: " + ex.Message);
                                }
                            }
                        }
                    }
                }
            }
            loadData();
        }

        void DrawImage(XGraphics gfx, string jpegSamplePath, int x, int y, int width, int height)
        {
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, x, y, width, height);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                strAcrobat = openFileDialog1.FileName;
            }
        }
    }
}
