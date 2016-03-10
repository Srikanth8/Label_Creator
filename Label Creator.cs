using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Desktop_Application
{
    public partial class Form1 : Form
    {
        public class CommonData
        {
            public CommonData(int n_cut, string n_buyer, string n_style, string n_po, char n_fabric, string n_colour)
            {
                cut_no = n_cut;
                buyer = n_buyer;
                style = n_style;
                po = n_po;
                fabric_lot = n_fabric;
                colour = n_colour;
            }
            public int cut_no { get; set; }
            public string buyer { get; set; }
            public string style { get; set; }
            public string po { get; set; }
            public char fabric_lot { get; set; }
            public string colour { get; set; }
        }

        public class BundleData
        {
            public BundleData(int BundleNo, int BundleSize, string BundleSerial)
            {
                Bundle_No = BundleNo;
                Bundle_Size = BundleSize;
                Bundle_Serial = BundleSerial;
            }
            public int Bundle_No { get; set; }
            public int Bundle_Size { get; set; }
            public string Bundle_Serial { get; set; }
        }

        private void loadData(int x)
        {
            a.Text = Common.cut_no.ToString();
            b.Text = Common.buyer;
            c.Text = Common.style;
            d.Text = Common.po;
            e.Text = Common.fabric_lot.ToString();
            f.Text = Common.colour;
            g.Text = Tags[x].size;
            h.Text = Tags[x].component;
            i.Text = Tags[x].bundle.Bundle_No.ToString();
            j.Text = Tags[x].bundle.Bundle_Size.ToString();
            k.Text = Tags[x].bundle.Bundle_Serial;

            TagNo.Text = (x+1) + " out of " + Tags.Count() + " labels";
            if (Tags.Count()%12 == 0)
                Pages.Text = (Tags.Count()/12) + " page(s) in total";
            else
                Pages.Text = (Tags.Count() / 12 + 1) + " page(s) in total";
        }

        public class Tag
        {
            public Tag(string n_size, string n_component, BundleData n_bundle)
            {
                size = n_size;
                component = n_component;
                bundle = n_bundle;
            }
            public string size { get; set; }
            public string component { get; set; }
            public BundleData bundle { get; set; }
        }

        private CommonData Common;
        private List<string> Sizes = new List<string>();
        private List<string> Components = new List<string>();
        private List<BundleData> Bundle = new List<BundleData>();
        private List<Tag> Tags = new List<Tag>();

        private int currentTag = 0;
        private int count = 1;

        public Form1()
        {
            InitializeComponent();
            //tabControl1.Appearance = TabAppearance.FlatButtons; 
            tabControl1.ItemSize = new Size(0, 1); 
            tabControl1.SizeMode = TabSizeMode.Fixed;
        }

        public class xl
        {
            public const int startColumn = 1;
            public const int startRow = 1;
            public const int maxColumn = 3;
            public const int maxRow = 4;
        }

        public class Sizes_Counter
        {
            public static int newSizes = 0;
            public static List<ComboBox> Sizes_Combo = new List<ComboBox>();
            public static List<TextBox> Sizes_Text = new List<TextBox>();
        }

        public class Components_Counter
        {
            public static int newComponents = 0;
            public static List<CheckBox> Components_Check = new List<CheckBox>();
            public static List<TextBox> Components_Text = new List<TextBox>();
        }

        private void ExcelEditor(Excel.Worksheet xlWorkSheet, int columnNo, int rowNo, int x, bool changeRowSize)
        {
            int cellColumn = columnNo*3 - 2 + 1;
            int cellRow = ((rowNo-1)/xl.maxRow)*49 + 2 + ((rowNo-1) % xl.maxRow)*12;

            if (changeRowSize)
            {
                for (int i=0; i<11; i++)
                    xlWorkSheet.Rows[cellRow + i].RowHeight = 15;
            }

            xlWorkSheet.Cells[cellRow, cellColumn + 1].NumberFormat = "@";
            xlWorkSheet.Cells[cellRow + 2, cellColumn + 1].NumberFormat = "@";
            xlWorkSheet.Cells[cellRow + 3, cellColumn + 1].NumberFormat = "@";
            xlWorkSheet.Cells[cellRow + 8, cellColumn + 1].NumberFormat = "@";
            xlWorkSheet.Cells[cellRow + 9, cellColumn + 1].NumberFormat = "@";
            xlWorkSheet.Cells[cellRow + 10, cellColumn + 1].NumberFormat = "@";

            xlWorkSheet.Cells[cellRow, cellColumn] = "Cut No";
            xlWorkSheet.Cells[cellRow + 1, cellColumn] = "Buyer";
            xlWorkSheet.Cells[cellRow + 2, cellColumn] = "Style";
            xlWorkSheet.Cells[cellRow + 3, cellColumn] = "PO";
            xlWorkSheet.Cells[cellRow + 4, cellColumn] = "Fabric Lot";
            xlWorkSheet.Cells[cellRow + 5, cellColumn] = "Colour";
            xlWorkSheet.Cells[cellRow + 6, cellColumn] = "Size";
            xlWorkSheet.Cells[cellRow + 7, cellColumn] = "Component";
            xlWorkSheet.Cells[cellRow + 8, cellColumn] = "Bundle No";
            xlWorkSheet.Cells[cellRow + 9, cellColumn] = "Quantity";
            xlWorkSheet.Cells[cellRow + 10, cellColumn] = "Serial No";

            xlWorkSheet.Cells[cellRow, cellColumn + 1] = Common.cut_no.ToString();
            xlWorkSheet.Cells[cellRow + 1, cellColumn + 1] = Common.buyer;
            xlWorkSheet.Cells[cellRow + 2, cellColumn + 1] = Common.style;
            xlWorkSheet.Cells[cellRow + 3, cellColumn + 1] = Common.po;
            xlWorkSheet.Cells[cellRow + 4, cellColumn + 1] = Common.fabric_lot.ToString();
            xlWorkSheet.Cells[cellRow + 5, cellColumn + 1] = Common.colour;
            xlWorkSheet.Cells[cellRow + 6, cellColumn + 1] = Tags[x].size;
            xlWorkSheet.Cells[cellRow + 7, cellColumn + 1] = Tags[x].component;
            xlWorkSheet.Cells[cellRow + 8, cellColumn + 1] = Tags[x].bundle.Bundle_No.ToString();
            xlWorkSheet.Cells[cellRow + 9, cellColumn + 1] = Tags[x].bundle.Bundle_Size.ToString();
            xlWorkSheet.Cells[cellRow + 10, cellColumn + 1] = Tags[x].bundle.Bundle_Serial;
        }

        private void SubmitPrint(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to print?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string fileName = submitToExcel();
                if (fileName != null)
                    print(fileName);
            }
        }

        private void SubmitOnly(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to submit?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                submitToExcel();
            }
        }

        private string submitToExcel()
        {
             Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return null;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 1; i < 10; i++)
            {
                xlWorkSheet.Columns[i].ColumnWidth = 13;
            }

            xlWorkSheet.Columns[1].ColumnWidth = 0.44;
            xlWorkSheet.Columns[4].ColumnWidth = 0.44;
            xlWorkSheet.Columns[7].ColumnWidth = 0.44;
            xlWorkSheet.Columns[10].ColumnWidth = 0.44;
            xlWorkSheet.Columns[1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            xlWorkSheet.Columns[4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            xlWorkSheet.Columns[7].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            xlWorkSheet.Columns[10].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
           
            int tags = Tags.Count();
            int pages = tags / 12;

            int columnNo = 1;
            int rowNo = 1;

            for (int x = 0; x < Tags.Count(); x++)
            {
                if (columnNo == xl.maxColumn)
                {
                    ExcelEditor(xlWorkSheet, columnNo, rowNo, x, true);

                    xlWorkSheet.Rows[((rowNo-1) / xl.maxRow) * 49 + 1 + ((rowNo-1) % xl.maxRow) * 12].RowHeight = 4.2;
                    xlWorkSheet.Rows[((rowNo-1) / xl.maxRow) * 49 + 1 + ((rowNo-1) % xl.maxRow) * 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                    if (rowNo % xl.maxRow == 0)
                    {
                        xlWorkSheet.Rows[(rowNo/xl.maxRow)*49].RowHeight = 4.2;
                        xlWorkSheet.Rows[(rowNo/xl.maxRow)*49].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                        xlWorkSheet.HPageBreaks.Add(xlWorkSheet.Range["A" + ((rowNo / xl.maxRow)*49 + 1)]);
                    }
                    else if (x == Tags.Count() - 1)
                    {          
                        xlWorkSheet.Rows[(rowNo / xl.maxRow) * 49 + 1 + (rowNo % xl.maxRow) * 12].RowHeight = 4.2;
                        xlWorkSheet.Rows[(rowNo / xl.maxRow) * 49 + 1 + (rowNo % xl.maxRow) * 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    }

                    columnNo = xl.startColumn;
                    rowNo++;
                }
                else
                {
                    ExcelEditor(xlWorkSheet, columnNo, rowNo, x, false);
                    columnNo++;

                    if (x == Tags.Count() - 1)
                    {
                        xlWorkSheet.Rows[((rowNo - 1) / xl.maxRow) * 49 + 1 + ((rowNo - 1) % xl.maxRow) * 12].RowHeight = 4.2;
                        xlWorkSheet.Rows[((rowNo - 1) / xl.maxRow) * 49 + 1 + ((rowNo - 1) % xl.maxRow) * 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                        xlWorkSheet.Rows[(rowNo / xl.maxRow) * 49 + 1 + (rowNo % xl.maxRow) * 12].RowHeight = 4.2;
                        xlWorkSheet.Rows[(rowNo / xl.maxRow) * 49 + 1 + (rowNo % xl.maxRow) * 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    }
                }
            }

            xlWorkSheet.Cells[2, 11] = "P.T. Muara Krakatau";
            xlWorkSheet.Columns[11].ColumnWidth = 0;
 
            string name = "Labels" + count + ".xls";
            count++;

            try
            {
                string printPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                xlWorkBook.SaveAs(printPath + @"/" + name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch
            {
                MessageBox.Show("Excel File has not been Created. Please try Again");
                return null;
            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file Created! You can find the file in the Desktop");

            tabControl1.SelectedTab = tabPage1;

            Tags.Clear();
            Sizes.Clear();
            Components.Clear();
            Bundle.Clear();

            dataGrid1.DataSource = null;
            dataGrid1.Rows.Clear();
            dataGrid1.Visible = false;

            return name;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //System.IO.StreamReader fileToPrint; @"C:\Users\User\Desktop\" 
        //System.Drawing.Font printFont;

        private void print(string fileName)
        {
            Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Visible = false;
            ExcelApp.DisplayAlerts = false;

            string printPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            Excel.Workbook WBook = ExcelApp.Workbooks.Open
            (printPath + @"/" + fileName, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value);

            WBook.PrintOut(Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value);

            WBook.Close(false, Missing.Value, Missing.Value);

            ExcelApp.Quit();

            /*string printPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            fileToPrint = new System.IO.StreamReader(printPath + @"/" + fileName);
            printFont = new System.Drawing.Font("Arial", 10);
            printDocument1.Print();
            fileToPrint.Close();*/
        }

        /*private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            float yPos = 0f;
            int count = 0;
            float leftMargin = e.MarginBounds.Left;
            float topMargin = e.MarginBounds.Top;
            string line = null;
            float linesPerPage = e.MarginBounds.Height / printFont.GetHeight(e.Graphics);
            while (count < linesPerPage)
            {
                line = fileToPrint.ReadLine();
                if (line == null)
                {
                    break;
                }
                yPos = topMargin + count * printFont.GetHeight(e.Graphics);
                e.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, yPos, new StringFormat());
                count++;
            }
            if (line != null)
            {
                e.HasMorePages = true;
            }
        }*/

        private void Size_Back(object sender, EventArgs e)
        {
            MessageBox.Show(Sizes_Counter.Sizes_Text[4].Text + " " + Sizes_Counter.Sizes_Combo[4].Text);
        }

        private void Add_Size(object sender, EventArgs e)
        {
            if (Sizes_Counter.newSizes >= 6)
            {
                return;
            }

            button12.Visible = true;

            int i = Sizes_Counter.newSizes;
            int displacement = Sizes_Counter.newSizes * 39;

            Sizes_Counter.Sizes_Combo.Add(new ComboBox());
            groupBox2.Controls.Add(Sizes_Counter.Sizes_Combo[i]);
            Sizes_Counter.Sizes_Combo[i].FormattingEnabled = true;
            Sizes_Counter.Sizes_Combo[i].Items.AddRange(new object[] {
            "0",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8"});
            Sizes_Counter.Sizes_Combo[i].Location = new System.Drawing.Point(387, 30 + displacement);
            Sizes_Counter.Sizes_Combo[i].Name = "comboBox10" + Sizes_Counter.newSizes;
            Sizes_Counter.Sizes_Combo[i].Size = new System.Drawing.Size(128, 24);
            Sizes_Counter.Sizes_Combo[i].TabIndex = 19;
            Sizes_Counter.Sizes_Combo[i].Text = "0";


            Sizes_Counter.Sizes_Text.Add(new TextBox());
            groupBox2.Controls.Add(Sizes_Counter.Sizes_Text[i]);
            Sizes_Counter.Sizes_Text[i].AutoSize = true;
            Sizes_Counter.Sizes_Text[i].Font = new System.Drawing.Font("Verdana", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Sizes_Counter.Sizes_Text[i].ForeColor = System.Drawing.SystemColors.ControlText;
            Sizes_Counter.Sizes_Text[i].Location = new System.Drawing.Point(315, 28 + displacement);
            Sizes_Counter.Sizes_Text[i].Name = "textBox10" + Sizes_Counter.newSizes;
            Sizes_Counter.Sizes_Text[i].Size = new System.Drawing.Size(40, 25);
            Sizes_Counter.Sizes_Text[i].TabIndex = 18;

            Sizes_Counter.newSizes++;
        }

        private void Delete_Size(object sender, EventArgs e)
        {
            if (Sizes_Counter.Sizes_Combo.Count() != 0)
            {
                Sizes_Counter.Sizes_Combo[Sizes_Counter.Sizes_Combo.Count() - 1].Dispose();
                Sizes_Counter.Sizes_Text[Sizes_Counter.Sizes_Text.Count() - 1].Dispose();
                Sizes_Counter.Sizes_Combo.RemoveAt(Sizes_Counter.Sizes_Combo.Count() - 1);
                Sizes_Counter.Sizes_Text.RemoveAt(Sizes_Counter.Sizes_Text.Count() - 1);
                Sizes_Counter.newSizes--;
            }
            if (Sizes_Counter.Sizes_Combo.Count() == 0)
            {
                button12.Visible = false;
            }
        }

        private void Add_Component(object sender, EventArgs e)
        {
            if (Components_Counter.newComponents >= 6)
            {
                return;
            }

            button18.Visible = true;

            int i = Components_Counter.newComponents;
            int displacement = Components_Counter.newComponents * 39;

            Components_Counter.Components_Check.Add(new CheckBox());
            groupBox4.Controls.Add(Components_Counter.Components_Check[i]);

            Components_Counter.Components_Check[i].AutoSize = true;
            Components_Counter.Components_Check[i].Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Components_Counter.Components_Check[i].ForeColor = System.Drawing.SystemColors.ControlText;
            Components_Counter.Components_Check[i].Location = new System.Drawing.Point(310, 48 + displacement);
            Components_Counter.Components_Check[i].Name = "checkBox20" + Components_Counter.newComponents;
            Components_Counter.Components_Check[i].Size = new System.Drawing.Size(70, 24);
            Components_Counter.Components_Check[i].TabIndex = 32;
            Components_Counter.Components_Check[i].Text = "";
            Components_Counter.Components_Check[i].UseVisualStyleBackColor = true;

            Components_Counter.Components_Text.Add(new TextBox());
            groupBox4.Controls.Add(Components_Counter.Components_Text[i]);

            Components_Counter.Components_Text[i].AutoSize = true;
            Components_Counter.Components_Text[i].Font = new System.Drawing.Font("Verdana", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Components_Counter.Components_Text[i].ForeColor = System.Drawing.SystemColors.ControlText;
            Components_Counter.Components_Text[i].Location = new System.Drawing.Point(330, 43 + displacement);
            Components_Counter.Components_Text[i].Name = "textBox10" + Sizes_Counter.newSizes;
            Components_Counter.Components_Text[i].Size = new System.Drawing.Size(70, 25);
            Components_Counter.Components_Text[i].TabIndex = 18;

            Components_Counter.newComponents++;
        }

        private void Delete_Component(object sender, EventArgs e)
        {
            if (Components_Counter.Components_Check.Count() != 0)
            {
                Components_Counter.Components_Check[Components_Counter.Components_Check.Count() - 1].Dispose();
                Components_Counter.Components_Text[Components_Counter.Components_Text.Count() - 1].Dispose();
                Components_Counter.Components_Check.RemoveAt(Components_Counter.Components_Check.Count() - 1);
                Components_Counter.Components_Text.RemoveAt(Components_Counter.Components_Text.Count() - 1);
                Components_Counter.newComponents--;
            }
            if (Components_Counter.Components_Check.Count() == 0)
            {
                button18.Visible = false;
            }
        }

        private bool BundleProcessor()
        {
            Bundle.Clear();

            if (textBox4.Text.Length == 0 || textBox3.Text.Length == 0)
            {
                MessageBox.Show("Please enter data in all fields");
                return false;
            }

            for (int i = 0; i < textBox4.Text.Length; i++)
            {
                if (!char.IsNumber(textBox4.Text[i]))
                {
                    MessageBox.Show("No of Layers must be a Positive Integer Value");
                    return false;
                }
            }
            for (int i = 0; i < textBox3.Text.Length; i++)
            {
                if (!char.IsNumber(textBox3.Text[i]))
                {
                    MessageBox.Show("No of Bundles Required must be a Positive Integer Value");
                    return false;
                }
            }

            int layers = Convert.ToInt32(textBox4.Text);
            int bundles = Convert.ToInt32(textBox3.Text);

            if (bundles == 0 || layers == 0)
            {
                MessageBox.Show("Please enter Positive Integer Values in both Fields");
                return false;
            }

            if (bundles > layers)
            {
                MessageBox.Show("No of Bundles should not be Greater than no of Layers");
                return false;
            }

            int bundleSize = layers / bundles;
            int remainingLayers = layers % bundles;
            int startNo = 1;
            int endNo;

            for (int i = 0; i < bundles; i++)
            {
                if (i < remainingLayers)
                {
                    endNo = startNo + bundleSize;
                    Bundle.Add(new BundleData(i + 1, bundleSize + 1, startNo + "-" + endNo));
                    startNo = endNo + 1;
                }
                else
                {
                    endNo = startNo + bundleSize - 1;
                    Bundle.Add(new BundleData(i + 1, bundleSize, startNo + "-" + endNo));
                    startNo = endNo + 1;
                }
            }

            return true;
        }

        private void Preview(object sender, EventArgs e)
        {
            if (!BundleProcessor())
                return;

            dataGrid1.DataSource = null;
            dataGrid1.Rows.Clear();
            dataGrid1.DataSource = Bundle;
            dataGrid1.Visible = true;
        }

        private void Next_1(object sender, EventArgs e)
        {
            foreach (Control control in groupBox1.Controls)
            {
                if (control is TextBox)
                {
                    TextBox textBox = (TextBox)control;
                    if (textBox.Text == "")
                    {
                        MessageBox.Show("Please enter values in all fields to proceed");
                        return;
                    }
                }

                if (control is ComboBox)
                {
                    ComboBox comboBox = (ComboBox)control;
                    if (comboBox.Text == "")
                    {
                        MessageBox.Show("Please enter values in all fields to proceed");
                        return;
                    }
                }
            }

            for (int i = 0; i < CutNo.Text.Length; i++)
            {
                if (!char.IsNumber(CutNo.Text[i]))
                {
                    MessageBox.Show("Cut No should have a Positive Integer Value");
                    return;
                }
            }

            tabControl1.SelectedTab = tabPage2;

            Common = new CommonData(Convert.ToInt32(CutNo.Text), Buyer.Text, Style.Text, PO.Text, FabricLot.Text.ToUpper()[0], Colour.Text);

            foreach (Control control in groupBox1.Controls)
            {
                if (control is TextBox)
                {
                    TextBox textBox = (TextBox)control;
                    textBox.Text = null;
                }

                if (control is ComboBox)
                {
                    ComboBox comboBox = (ComboBox)control;
                    if (comboBox.Items.Count > 0)
                        comboBox.Text = null;
                }
            }
        }

        private void Next_2(object sender, EventArgs e)
        {
            bool check = false;
            foreach (Control control in groupBox2.Controls)
            {
                if (control is ComboBox)
                {
                    ComboBox comboBox = (ComboBox)control;

                    if (comboBox.Text == "")
                    {
                        comboBox.Text = "0";
                    }

                    if (comboBox.Text != "0")
                    {
                        check = true;   
                        
                        for (int i = 0; i < comboBox.Text.Length; i++)
                        {
                            if (!char.IsNumber(comboBox.Text[i]))
                            {
                                MessageBox.Show("Please enter Non-negative Integer Values for all Fields");
                                return;
                            }
                        }                                         

                        if (Convert.ToInt32(comboBox.Text) < 0)
                        {
                            MessageBox.Show("Please enter Non-negative Integer Values for all Fields");
                            return;
                        }                        
                    }
                }
            }

            if (check == false)
            {
                MessageBox.Show("Please select Quantity for atleast one Size");
                return;
            }

            Sizes.Clear();

            List<ComboBox> comboBoxes = new List<ComboBox>();
            comboBoxes.Add(XS); comboBoxes.Add(S); comboBoxes.Add(M); comboBoxes.Add(L);
            comboBoxes.Add(XL); comboBoxes.Add(XXL); comboBoxes.Add(XXXL);

            for (int i = 0; i < 7; i++)
            {
                if (Convert.ToInt32(comboBoxes[i].Text) > 0)
                {
                    if (Convert.ToInt32(comboBoxes[i].Text) == 1)
                        Sizes.Add(comboBoxes[i].Name);
                    else
                        for (int j = 0; j < Convert.ToInt32(comboBoxes[i].Text); j++)
                            Sizes.Add(comboBoxes[i].Name + (j + 1));
                }
            }

            for (int i = 0; i < Sizes_Counter.Sizes_Combo.Count; i++)
            {
                if (Convert.ToInt32(Sizes_Counter.Sizes_Combo[i].Text) > 0)
                {
                    if (Sizes_Counter.Sizes_Text[i].Text == "")
                    {
                        MessageBox.Show("Please enter the Name(s) of the new Size(s)");
                        return;
                    }

                    if (Convert.ToInt32(Sizes_Counter.Sizes_Combo[i].Text) == 1)
                        Sizes.Add(Sizes_Counter.Sizes_Text[i].Text);
                    else
                        for (int j = 0; j < Convert.ToInt32(Sizes_Counter.Sizes_Combo[i].Text); j++)
                            Sizes.Add(Sizes_Counter.Sizes_Text[i].Text + (j + 1));
                }
            }

            tabControl1.SelectedTab = tabPage3;
            
            //Sizes_Counter.Sizes_Combo.Clear();
            //Sizes_Counter.Sizes_Text.Clear();

            foreach (Control control in groupBox2.Controls)
            {
                if (control is ComboBox)
                {
                    ComboBox comboBox = (ComboBox)control;
                    if (comboBox.Items.Count > 0)
                        comboBox.Text = "0";
                }
            }
        }

        private void Next_3(object sender, EventArgs e)
        {
            bool check = false;
            foreach (Control control in groupBox4.Controls)
            {
                if (control is CheckBox)
                {
                    CheckBox checkBox = (CheckBox)control;
                    if (checkBox.Checked)
                    {
                        check = true;
                    }
                }
            }

            if (check == false)
            {
                MessageBox.Show("Please select atleast one Component");
                return;
            }

            Components.Clear();
            
            List<CheckBox> checkBoxes = new List<CheckBox>();
            checkBoxes.Add(Front); checkBoxes.Add(Back); checkBoxes.Add(Sleeve);
            checkBoxes.Add(Collar); checkBoxes.Add(Cuff); checkBoxes.Add(Placket);
            checkBoxes.Add(Pocket); checkBoxes.Add(Band); checkBoxes.Add(Yoke);

            for (int i = 0; i < 9; i++)
            {
                if (checkBoxes[i].Checked == true)
                    Components.Add(checkBoxes[i].Name);
            }

            for (int i = 0; i < Components_Counter.Components_Check.Count(); i++)
            {
                if (Components_Counter.Components_Check[i].Checked == true)
                {
                    if (Components_Counter.Components_Text[i].Text == "")
                    {
                        MessageBox.Show("Please enter the Name(s) of the new Component(s)");
                        return;
                    }

                    Components.Add(Components_Counter.Components_Text[i].Text);
                }
            }

            tabControl1.SelectedTab = tabPage4;

            //Components_Counter.Components_Check.Clear();
            //Components_Counter.Components_Text.Clear();

            foreach (Control control in groupBox4.Controls)
            {
                if (control is CheckBox)
                {
                    CheckBox checkBox = (CheckBox)control;
                    checkBox.Checked = false;
                }
            }
        }

        private void Next_4(object sender, EventArgs e)
        {
            currentTag = 0;
            if (BundleProcessor())
            {
                TagCreator();
                loadData(currentTag);
                tabControl1.SelectedTab = tabPage5;
                foreach (Control control in groupBox5.Controls)
                {
                    if (control is TextBox)
                    {
                        TextBox textBox = (TextBox)control;
                        textBox.Text = null;
                    }
                }
            }            
        }

        private void TagCreator()
        {
            Tags.Clear();
            foreach (string m_size in Sizes)
                foreach (string m_component in Components)
                    foreach (BundleData m_bundle in Bundle)
                    {
                        Tags.Add(new Tag(m_size, m_component, m_bundle));
                    }
        }

        private void Back_2(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
        }

        private void Back_3(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
        }

        private void Back_4(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
        }

        private void Back_5(object sender, EventArgs e)
        {
            dataGrid1.DataSource = null;
            dataGrid1.Rows.Clear();
            dataGrid1.Visible = false;
            tabControl1.SelectedTab = tabPage4;
        }

        private void n_Next(object sender, EventArgs e)
        {
            if (currentTag < Tags.Count() - 1)
            {
                currentTag++;
                loadData(currentTag);
            }
        }

        private void n_Prev(object sender, EventArgs e)
        {
            if (currentTag > 0)
            {
                currentTag = currentTag - 1;
                loadData(currentTag);
            }
        }

        private void n_First(object sender, EventArgs e)
        {
            currentTag = 0;
            loadData(currentTag);
        }

        private void n_Last(object sender, EventArgs e)
        {
            currentTag = Tags.Count() - 1;
            loadData(currentTag);
        }
    }
}
