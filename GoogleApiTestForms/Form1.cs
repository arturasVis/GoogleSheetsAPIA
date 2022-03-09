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
        static private string path = "C:\\Users\\artel\\OneDrive\\Desktop\\C# Shit\\GoogleSheetsAPIA\\ConsoleApp1\\";
        GoogleApi sheets=new GoogleApi("Desktop application 1","1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI", path);
        public Form1()
        {
            InitializeComponent();
            sheets.Authentication();
            var sheet=sheets.ReadEntries("!A","E","openorder");
            //sheets.CreateEntry();

            var source = readCSV(path + "openorder.csv");
            dataGridView1.DataSource = sheet;
            

            List<object> list = new List<object>();
            var append = Compare(sheet,source);
            int lastIndex=sheet.Rows.Count;
            dataGridView2.DataSource = append;
            for (int i=0;i< append.Rows.Count;i++)
            {
                for (int j=0;j< append.Columns.Count;j++)
                {
                    list.Add(append.Rows[i][j]);
                }
                sheets.UpdateSheet(list, "openorder", "!A", "E");
                list.Clear();
                index(lastIndex);
                lastIndex++;
            }
            Application.Exit();
        }
        public void index(int lastIndex)
        {
            int j = lastIndex;
            j++;
            int cell = j + 1;
            List<object> list = new List<object>(); list.Add(j);
            sheets.UpdateSheet(list,"openorder","!A"+cell,"A"+cell);
            list.Clear();
        }
        
        public DataTable readCSV(string filePath)
        {
            var dt=new DataTable();
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
            foreach (var line in File.ReadLines(filePath))
            {
                string[] vs = line.Split(',');
                dt.Rows.Add(null,vs[0].Trim('"'),vs[1].Trim('"'), vs[2].Trim('"'), vs[3].Trim('"'));
            }
            return dt;
        }
        private DataTable Compare(DataTable FirstDataTable,DataTable SecondDataTable)
        {
            DataTable dt= new DataTable();
            dt.Columns.Add("Index #");
            dt.Columns.Add("Order #");
            dt.Columns.Add("SKU");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Channel");
            bool isDuplicate=false;
            for (int i=0;i<SecondDataTable.Rows.Count;i++)
            {
                for (int j=0;j<FirstDataTable.Rows.Count;j++)
                {
                    
                    if (SecondDataTable.Rows[i][1].ToString()==FirstDataTable.Rows[j][1].ToString())
                    {
                        isDuplicate=true;
                    }
                }
                if (isDuplicate==false)
                {
                    dt.Rows.Add(null, SecondDataTable.Rows[i][1], SecondDataTable.Rows[i][2], SecondDataTable.Rows[i][3], SecondDataTable.Rows[i][4]);
                }
                isDuplicate = false;
            }
            
            return dt;
        }
    }
}
