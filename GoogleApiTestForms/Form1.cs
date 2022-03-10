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

namespace GoogleApiTestForms
{
    public partial class Form1 : Form
    {
        static private string path;
        static private int time;
        //GoogleApi sheetLine= new GoogleApi("Desktop application 1", "1ze7XiL7g5Vf_dNg7_AycXgMC2ekZQ1nnSFk9YynTii4", path);
        GoogleApi sheetTesting = new GoogleApi("Desktop application 1", "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI",path);
        public Form1()
        {
            InitializeComponent();
            readPath();
            //sheetLine.Authentication();       
            sheetTesting.Authentication();
            Timer tmr = new Timer();
            tmr.Interval = time;   // milliseconds
            tmr.Tick += Tmr_Tick;  // set handler
            tmr.Start();
        }
        private void Tmr_Tick(object sender, EventArgs e)
        {
            //var sheet = sheetLine.ReadEntries("!A", "D", "History(Server)");
            var sheet2 = sheetTesting.ReadEntries("!A", "E", "History Testing");
            //var openSheet = sheetLine.ReadEntries("!A", "D", "Orders");
            var openSheet2= sheetTesting.ReadEntries("!A", "E", "Open Testing");
            //var source = readCSV(path + "openorder.csv");
            var source2 = readCSV(path + "openorderProgram.csv");

            //Update("Orders", source, sheet, openSheet.Rows.Count,sheetLine);
            //Update("History(Server)", source, sheet, sheet.Rows.Count,sheetLine);
            Update("Open Testing", source2, sheet2, openSheet2.Rows.Count,sheetTesting);
            Update("History Testing", source2, sheet2, sheet2.Rows.Count, sheetTesting);
            richTextBox1.AppendText("Updated sheet at: " + DateTime.Now + "\n");
        }
        /*public void index(int lastIndex,string sheet,GoogleApi sheets)
        {
            int j = lastIndex;
            j++;
            int cell = j + 1;
            List<object> list = new List<object>(); list.Add(j);
            sheets.UpdateSheet(list, sheet, "!A" + cell, "A" + cell);
            list.Clear();
        }*/

        public DataTable readCSV(string filePath)
        {
            var dt = new DataTable();
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            try
            {
                foreach (var line in File.ReadLines(filePath))
                {
                    string[] vs = line.Split(',');
                    dt.Rows.Add(vs[0].Trim('"'), vs[1].Trim('"'), vs[2].Trim('"'), vs[3].Trim('"'));
                }
                return dt;
            }
            catch {
                return null;
            }
        }
        private DataTable Compare(DataTable FirstDataTable, DataTable SecondDataTable)
        {
            DataTable dt = new DataTable();
            //dt.Columns.Add("Index #");
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

                        if (SecondDataTable.Rows[i][0].ToString() == FirstDataTable.Rows[j][0].ToString())
                        {
                            isDuplicate = true;
                        }
                    }
                    if (isDuplicate == false)
                    {
                        dt.Rows.Add(SecondDataTable.Rows[i][0], SecondDataTable.Rows[i][1], SecondDataTable.Rows[i][2], SecondDataTable.Rows[i][3]);
                    }
                    isDuplicate = false;
                }

            }
            return dt;
        }
        public void Update(string sheet,DataTable source,DataTable history,int lastIndex,GoogleApi sheets)
        {
            List<object> list = new List<object>();
            var append = Compare(history, source);
            for (int i = 0; i < append.Rows.Count; i++)
            {
                for (int j = 0; j < append.Columns.Count; j++)
                {
                    list.Add(append.Rows[i][j]);
                }
                sheets.UpdateSheet(list, sheet, "!A", "D");
                list.Clear();
                //index(lastIndex, sheet,sheets);
                //lastIndex++;
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
    }
}
