using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;  

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        string pathName = "_properties.xls";//file path address
        DataSet ds = new DataSet();
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        List<DeviceProperty> dp_list = new List<DeviceProperty>();
        public Form1()
        {
            InitializeComponent();
            //Create COM Objects. Create a COM object for everything that is referenced
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"" + pathName);
            ds.Reset();
            DataTable _myDataTable = new DataTable();
            
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                _myDataTable.Columns.Add(new DataColumn(xlRange.Cells[1, i + 1].Value2.ToString()));
            }
            for (int i = 1; i < rowCount; i++)
            {
                // create a DataRow using .NewRow()
                DataRow row = _myDataTable.NewRow();

                // iterate over all columns to fill the row
                for (int j = 0; j < colCount; j++)
                {
                    if (xlRange.Cells[i + 1, j + 1].Value != null)
                        row[j] = xlRange.Cells[i + 1, j + 1].Value.ToString();
                }

                // add the current row to the DataTable
                _myDataTable.Rows.Add(row);
            }
            ds.Tables.Add(_myDataTable);
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DeviceProperty new_dp = new DeviceProperty();
                if (!dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("breake"))//"breake" is null row.
                {

                    new_dp.custom_id = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    new_dp.deviceId = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    new_dp.propertyName = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    new_dp.propertyId = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    new_dp.propertyType = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    dp_list.Add(new_dp);
                }
                else { break; }

            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(@"" + pathName);
            ds.Reset();
            DataTable _myDataTable = new DataTable();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                _myDataTable.Columns.Add(new DataColumn(xlRange.Cells[1, i + 1].Value2.ToString()));
            }
            for (int i = 1; i < rowCount; i++)
            {
                // create a DataRow using .NewRow()
                DataRow row = _myDataTable.NewRow();

                // iterate over all columns to fill the row
                for (int j = 0; j < colCount; j++)
                {
                    if (xlRange.Cells[i + 1, j + 1].Value != null)
                        row[j] = xlRange.Cells[i + 1, j + 1].Value.ToString();
                }

                // add the current row to the DataTable
                _myDataTable.Rows.Add(row);
            }
            ds.Tables.Add(_myDataTable);
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataGridView filter = new DataGridView();
            DataView dv = ds.Tables[0].DefaultView;

            dv.RowFilter = string.Format("" + dataGridView1.Columns[dataGridView1.SelectedCells[0].ColumnIndex].Name + " = '" + dataGridView1.SelectedCells[0].Value.ToString() + "'");
            filter.DataSource = dv.Table;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();  //create openfileDialog Object
                openFileDialog1.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
                openFileDialog1.FilterIndex = 3;

                openFileDialog1.Multiselect = false;        //not allow multiline selection at the file selection level
                openFileDialog1.Title = "Open Text File-R13";   //define the name of openfileDialog
                openFileDialog1.InitialDirectory = @"Desktop"; //define the initial directory


                if (openFileDialog1.ShowDialog() == DialogResult.OK)        //executing when file open
                {
                    pathName = openFileDialog1.FileName;
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(@"" + pathName);
                    if (xlWorkbook.Sheets.Count < 2) 
                    {
                        button2.Enabled = false;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //rule of thumb for releasing com objects:
                    //  never use two dots, all COM objects must be referenced and released individually
                    //  ex: [somthing].[something].[something] is bad

                    //release com objects to fully kill excel process from running in the background
                    

                    //close and release
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception exp) { MessageBox.Show(exp.Message); }
        }
        public struct DeviceInfo
        {
            public string device_custom_id;
            public string device_mac;
            public string device_name;
            public string lat;
            public string lun;
            public string device_type;
            public string move_type;
            public string autURL;
            public string feedURL;
            public string url;
            public string method;
            public string username;
            public string password;
            public string domainRef;
        }
        public struct DeviceProperty
        {
            public string custom_id;
            public string deviceId;
            public string propertyName;
            public string propertyId;
            public string propertyType;
        }
        

    }
}
