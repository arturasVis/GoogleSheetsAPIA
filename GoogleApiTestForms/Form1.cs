using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GoogleApiTestForms
{
    public partial class Form1 : Form
    {
        static private string path;
        static private int time;
        Database1DataSet sheet = new Database1DataSet();
        

        DataTable newOrders=new DataTable();
        GoogleApi sheetLine= new GoogleApi("Desktop application 1", "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI", path);
        //GoogleApi sheetTesting = new GoogleApi("Desktop application 1", "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI",path);
        public Form1()
        {
            InitializeComponent();
            readPath();
            sheetLine.Authentication();       
            //sheetTesting.Authentication();
            Timer tmr = new Timer();
            tmr.Interval = 1000;   // milliseconds
            tmr.Tick += Tmr_Tick;  // set handler
            tmr.Start();
            sheetLine.Testing();
        }
        private void Tmr_Tick(object sender, EventArgs e)
        {
            var csvFile = readCSV(path + "openorder.csv");
            //newOrders = FindNew(sheet.Tables["History"],csvFile);
            
            AppendToDataSet(csvFile);
            
            for (int i=0;i<sheet.Tables["Table"].Rows.Count;i++)
            {
                richTextBox1.AppendText(sheet.Tables["Table"].Rows[i][0] + "     "+sheet.Tables["Table"].Rows[i]["1"]+"\n");
            }
            richTextBox1.AppendText( "\n");
            richTextBox1.AppendText("\n");
            richTextBox1.AppendText("\n");
            /*if (CheckForInternetConnection())
            {
                var sheet = sheetLine.ReadEntries("!A", "E", "History");
                //var sheet2 = sheetTesting.ReadEntries("!A", "E", "History Testing");
                var openSheet = sheetLine.ReadEntries("!A", "E", "Open");
                //var openSheet2= sheetTesting.ReadEntries("!A", "E", "Open Testing");
                var source = readCSV(path + "openorder.csv");
                //var source2 = readCSV(path + "openorderProgram.csv");

                Update("History", source, sheet, openSheet.Rows.Count, sheetLine);
                Update("Open", source, sheet, sheet.Rows.Count, sheetLine);
                //Update("Open Testing", source2, sheet2, openSheet2.Rows.Count,sheetTesting);
                //Update("History Testing", source2, sheet2, sheet2.Rows.Count, sheetTesting);
                richTextBox1.AppendText("Updated sheet at: " + DateTime.Now + "\n");
            }
            else {
                richTextBox1.AppendText("No internet connection\n");
            }*/

        }
        public void index(int lastIndex,string sheet,GoogleApi sheets)
        {
            int j = lastIndex;
            j++;
            int cell = j + 1;
            List<object> list = new List<object>(); list.Add(j);
            sheets.UpdateSheet(list, sheet, "!A" + cell, "A" + cell);
            list.Clear();
        }

        public DataTable readCSV(string filePath)
        {
            var dt = new DataTable();
            DataColumn column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.AutoIncrement = true;
            column.AutoIncrementSeed = 1;
            column.AutoIncrementStep = 1;
            column.ColumnName = "Index #";
            dt.Columns.Add(column);
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            try
            {
                foreach (var line in File.ReadLines(filePath))
                {
                    string[] vs = line.Split(',');
                    dt.Rows.Add(null, vs[0].Trim('"'), vs[1].Trim('"'), vs[2].Trim('"'), vs[3].Trim('"'));
                }
                dt.DefaultView.Sort = "Channel ASC";
                return dt;
            }
            catch {
                return null;
            }
        }
        private void AppendToDataSet(DataTable DataTable)
        {
            bool isDuplicate=false;
            DataTable dt = sheet.Tables["Table"];
            if (dt!= null && DataTable != null)
            {
                for (int i = 0; i < DataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {

                        if (DataTable.Rows[i][1].ToString() == dt.Rows[j][1].ToString())
                        {
                            isDuplicate = true;
                        }
                    }
                    if (isDuplicate == false)
                    {
                        dt.Rows.Add(DataTable.Rows[i]["Order #"],DataTable.Rows[i]["SKU"],
                            DataTable.Rows[i]["QTY"],DataTable.Rows[i]["Channel"]);
                    }
                    isDuplicate = false;
                }
                sheet.Tables["Table"].Merge(dt);
                sheet.AcceptChanges();
            }

        }
        private DataTable FindNew(DataTable FirstDataTable, DataTable SecondDataTable)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Index #");
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            bool isDuplicate = false;
            if (FirstDataTable != null && SecondDataTable != null)
            {
                for (int i = 0; i < SecondDataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < FirstDataTable.Rows.Count; j++)
                    {

                        if (SecondDataTable.Rows[i][1].ToString() == FirstDataTable.Rows[j][1].ToString())
                        {
                            isDuplicate = true;
                        }
                    }
                    if (isDuplicate == false)
                    {
                        dt.Rows.Add(null, SecondDataTable.Rows[i][1], SecondDataTable.Rows[i][2], SecondDataTable.Rows[i][3], SecondDataTable.Rows[i][4]);
                    }
                    isDuplicate = false;
                }
                
            }
            return dt;
        }
        public void Update(string sheet,DataTable source,DataTable history,int lastIndex,GoogleApi sheets)
        {
            List<object> list = new List<object>();
            var append = FindNew(history, source);
            for (int i = 0; i < append.Rows.Count; i++)
            {
                for (int j = 0; j < append.Columns.Count; j++)
                {
                    list.Add(append.Rows[i][j]);
                }
                sheets.UpdateSheet(list, sheet, "!A", "E");
                list.Clear();
                index(lastIndex, sheet,sheets);
                lastIndex++;
            }
        }
        public void readPath()
        {
            StreamReader sr = new StreamReader(Application.StartupPath+"//path.txt");
            string line;
            line = sr.ReadLine();
            path=line;
            line=sr.ReadLine();
            time = Int32.Parse(line);
        }
        public bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (client.OpenRead("http://google.com/generate_204"))
                    return true;
            }
            catch
            {
                return false;
            }
        }
        public void GetDBTAble()
        {
            
        }


    }
}
