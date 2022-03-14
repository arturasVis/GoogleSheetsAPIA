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
        DataSet sheet = new DataSet();
        DBManager dBManager;
        DataTable csvFile1, csvFile2;

        DataTable newOrders=new DataTable();
        GoogleApi sheetLine= new GoogleApi("Desktop application 1", "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI", path);
        //GoogleApi sheetTesting = new GoogleApi("Desktop application 1", "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI",path);
        public Form1()
        {
            InitializeComponent();
            readPath();
            sheetLine.Authentication();
            dBManager = new DBManager();
            //sheetTesting.Authentication();
            csvFile1=readCSV(path + "openorder.csv");
            csvFile2 = readCSV(path + "openorderProgram.csv");
            var newOrders = FindNew(dBManager.Select("History"), csvFile1);
            if (newOrders.Rows.Count>0)
            {
                Run(newOrders);
            }
            dataGridView1.DataSource = dBManager.SelectDate("History", DateTime.Now.ToString("MM/dd/yyyy"));
            dataGridView2.DataSource = dBManager.SelectDate("childsku", DateTime.Now.ToString("MM/dd/yyyy"));
            Timer tmr = new Timer();
            tmr.Interval = 1000;   // milliseconds
            tmr.Tick += Tmr_Tick;  // set handler
            tmr.Start();
            sheetLine.Testing();
        }
        private void Run(DataTable newOrders)
        {
            for (int i=0;i<newOrders.Rows.Count;i++)
            {
                string querry1 = $"INSERT INTO History(Orderid,SKU,QTY,Channel) VALUES " +
                    $"('{newOrders.Rows[i][1].ToString()}','{newOrders.Rows[i][2].ToString()}','{newOrders.Rows[i][3].ToString()}','{newOrders.Rows[i][4].ToString()}')";
                dBManager.InsertQuerry(querry1);
            }
            DataTable db = dBManager.Select("History");

            foreach (DataRow rowHistory in db.Rows)
            {
                foreach (DataRow rowCSV2 in csvFile2.Rows)
                {
                    if (Int32.Parse(rowHistory[1].ToString())==Int32.Parse(rowCSV2[1].ToString()))
                    {
                        string querry1 = $"INSERT INTO childsku(Id,Orderid,ChildSKU) VALUES " +
                            $"('{Int32.Parse(rowHistory[0].ToString())}','{rowHistory[1].ToString()}','{rowCSV2[2].ToString()}')";
                        dBManager.InsertQuerry(querry1);
                    }
                }
            }

            int lastIndex = sheetLine.ReadEntries("!A", "E", "Open").Rows.Count;
            UpdateSheet("Open", newOrders, lastIndex, sheetLine);
            richTextBox1.AppendText("Updated sheet at: " + DateTime.Now + "\n");
            dataGridView1.DataSource = dBManager.SelectDate("History", DateTime.Now.ToString("MM/dd/yyyy"));
            dataGridView2.DataSource = dBManager.SelectDate("childsku", DateTime.Now.ToString("MM/dd/yyyy"));
        }
        
        private void Tmr_Tick(object sender, EventArgs e)
        {
            readCSV(path + "openorder.csv");
            var newOrders = FindNew(dBManager.SelectDate("History",DateTime.Now.ToString("MM/dd/yyyy")), csvFile1);
            if(newOrders.Rows.Count > 0)
            {
                Run(newOrders);
            }
            

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
            DataTable dt = new DataTable();
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
                foreach (var lines in File.ReadLines(filePath))
                {
                    string[] line = lines.Split(',');
                    dt.Rows.Add(null, line[0].Trim('"'), line[1].Trim('"'), line[2].Trim('"'), line[3].Trim('"'));
                }
            }
            catch {
                return null;
            }
            return dt;
        }
        /*private void AppendToDataSet(DataTable DataTable,string querry)
        {
            for (int i=0;i<DataTable.Rows.Count;i++)
            {
                string querry1 = $"INSERT INTO History(Orderid,SKU,QTY,Channel) VALUES ('{value[1].ToString()}','{value[2].ToString()}','{value[3].ToString()}','{value[4].ToString()}')";
                dBManager.Insert(querry);
            }
        }*/
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

                        if (Int32.Parse(SecondDataTable.Rows[i][1].ToString()) == Int32.Parse(FirstDataTable.Rows[j][1].ToString()))
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
        public void UpdateSheet(string sheet,DataTable append,int lastIndex,GoogleApi sheets)
        {
            List<object> list = new List<object>();
            
            for (int i = 0; i < append.Rows.Count; i++)
            {
                if (Int32.Parse(append.Rows[i][3].ToString()) > 1)
                {
                    for (int j = 0; j < Int32.Parse(append.Rows[i][3].ToString()); j++)
                    {
                        for (int k = 0; k < append.Columns.Count; k++)
                        {
                            if (k == 3)
                            {
                                list.Add(1);
                            }
                            else
                            {
                                list.Add(append.Rows[i][k]);
                            }
                        }
                        sheets.UpdateSheet(list, sheet, "!A", "E");
                        list.Clear();
                        index(lastIndex, sheet, sheets);
                        lastIndex++;
                    }
                }
                else
                {
                    for (int j = 0; j < append.Columns.Count; j++)
                    {
                        list.Add(append.Rows[i][j]);
                    }
                    sheets.UpdateSheet(list, sheet, "!A", "E");
                    list.Clear();
                    index(lastIndex, sheet, sheets);
                    lastIndex++;
                }
                
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
    }
}
