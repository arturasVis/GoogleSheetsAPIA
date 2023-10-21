using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;

namespace GoogleApiTestForms
{
    public partial class Form1 : Form
    {
        static private string path;
        static private int time;
        DataSet sheet = new DataSet();
        DBManager dBManager;
        DataTable csvFile1;
        DataTable newOrders = new DataTable();

        public Form1()
        {
            InitializeComponent();
            readPath();
            dBManager = new DBManager();
            csvFile1 = readCSV(path + "openorder.csv");

            var newOrders = FindNew(dBManager.Select("History"), csvFile1);
            if (newOrders.Rows.Count > 0)
            {
                Run(newOrders);
            }
            dataGridView1.DataSource = dBManager.SelectDate("History", DateTime.Now.ToString("MM/dd/yyyy"));

            Timer tmr = new Timer();
            tmr.Interval = 1000;   // milliseconds
            tmr.Tick += Tmr_Tick;  // set handler
            tmr.Start();
        }

        private void Run(DataTable newOrders)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Index ");
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            dt.Columns.Add("Date");
            dt.Columns.Add("Tracking Id");

            for (int i = 0; i < newOrders.Rows.Count; i++)
            {
                if (Int32.Parse(newOrders.Rows[i][3].ToString()) > 0)
                {
                    for (int k = 0; k < Int32.Parse(newOrders.Rows[i][3].ToString()); k++)
                    {
                        string querry1 = $"INSERT INTO History(Orderid,SKU,QTY,Channel) VALUES " +
                            $"('{newOrders.Rows[i][1].ToString()}','{newOrders.Rows[i][2].ToString()}','1','{newOrders.Rows[i][4].ToString()}')";
                        dBManager.InsertQuerry(querry1);
                        DataTable NewOrder = dBManager.SelectOrder("History", newOrders.Rows[i][1].ToString());
                        dt.Rows.Add(NewOrder.Rows[0][0].ToString(), newOrders.Rows[i][1].ToString(), newOrders.Rows[i][2].ToString(), newOrders.Rows[i][3].ToString(), newOrders.Rows[i][4].ToString(), " ", newOrders.Rows[i][5].ToString());
                    }
                }
                else
                {
                    string querry1 = $"INSERT INTO History(Orderid,SKU,QTY,Channel) VALUES " +
                        $"('{newOrders.Rows[i][1].ToString()}','{newOrders.Rows[i][2].ToString()}','{newOrders.Rows[i][3].ToString()}','{newOrders.Rows[i][4].ToString()}')";
                    dBManager.InsertQuerry(querry1);
                }
            }
            richTextBox1.AppendText("Updated sheet at: " + DateTime.Now + "\n");
            dataGridView1.DataSource = dBManager.SelectDate("History", DateTime.Now.ToString("MM/dd/yyyy"));
            dt.Clear();
        }

        private void Tmr_Tick(object sender, EventArgs e)
        {
            csvFile1 = readCSV(path + "openorder.csv");
            var newOrders = FindNew(dBManager.Select("History"), csvFile1);
            if (newOrders.Rows.Count > 0)
            {
                Run(newOrders);
            }
        }

        public DataTable readCSV(string filePath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Index #");
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            dt.Columns.Add("Tracking Id");

            try
            {
                foreach (var lines in File.ReadLines(filePath))
                {
                    string[] line = lines.Split(',');
                    dt.Rows.Add(null, line[0].Trim('"'), line[1].Trim('"'), line[2].Trim('"'), line[3].Trim('"'), line[4].Trim('"'));
                }
            }
            catch
            {
                return null;
            }
            return dt;
        }

        private DataTable FindNew(DataTable FirstDataTable, DataTable SecondDataTable)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Index #");
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            dt.Columns.Add("Tracking Id");

            // Create a HashSet and populate it with order numbers from FirstDataTable
            HashSet<int> existingOrderNumbers = new HashSet<int>();
            foreach (DataRow row in FirstDataTable.Rows)
            {
                if (row[4].ToString().Trim(' ') != "Prebuilt")
                {
                    existingOrderNumbers.Add(Int32.Parse(row[1].ToString()));
                }
            }

            // Loop through SecondDataTable and check against the HashSet for duplicates
            foreach (DataRow row in SecondDataTable.Rows)
            {
                int orderNumber = Int32.Parse(row[1].ToString());
                bool isDuplicate = existingOrderNumbers.Contains(orderNumber);

                if ((!isDuplicate && !String.IsNullOrEmpty(row[5].ToString())) ||
                    (!isDuplicate && row[4].ToString().Contains("FBA")))
                {
                    dt.Rows.Add(null, row[1], row[2], row[3], row[4], row[5]);
                }
            }

            return dt;
        }


        public void readPath()
        {
            StreamReader sr = new StreamReader(Application.StartupPath + "//path.txt");
            string line;
            line = sr.ReadLine();
            path = line;
            line = sr.ReadLine();
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

        public void Write(string text, string path)
        {
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.WriteLine(text);
            }
        }
    }
}
