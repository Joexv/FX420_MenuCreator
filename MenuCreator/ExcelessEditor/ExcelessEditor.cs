using IniParser;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace MenuCreator.ExcelessEditor
{
    public partial class ExcelessEditor : Form
    {
        public ExcelessEditor()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.WriteLine("Setting strings");
            XLSX = Path.Combine(System.Windows.Forms.Application.StartupPath, MenuName + ".xlsx");
            INI = Path.Combine(System.Windows.Forms.Application.StartupPath, MenuName + "_Products.ini");
            CompanyList = Path.Combine(@System.Windows.Forms.Application.StartupPath, MenuName + "_Companys.txt");
            this.Text = "Exceless Editor :: " + MenuName;
        }

        private string XLSX;
        private string INI;
        private string CompanyList;

        public string MenuName { get; set; }

        private List<string> Companys = new List<string> { };
        private List<string> Products = new List<string> { };

        private void RefreshCompanys()
        {
            Companys.Clear();
            Console.WriteLine("Refreshing Company List");
            Companys = File.ReadAllLines(CompanyList).ToList();
            Companys.Sort();
            companyGrid.DataSource = CompanyData();
        }

        private char[] alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        private void RefreshProducts(string company)
        {
            Products.Clear();
            Console.WriteLine("Refreshing Products");
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(INI);
            foreach (char c in alph)
            {
                Console.WriteLine(c);
                try
                {
                    string value = data[company][c.ToString().ToUpper()] ?? "";
                    if (value != null || value != "")
                    {
                        Products.Add(value);
                    }

                    Console.WriteLine(value);
                }
                catch { }
            }
            Products.Sort();
        }

        private void SaveCompanys()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)companyGrid.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(String.Format("{0}:{1}", DataRow[0].ToString(), DataRow[1].ToString()));
            }

            File.WriteAllText(CompanyList, String.Join("\n", Values.ToArray()));
        }

        private DataTable CompanyData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Number");
            dt.Columns.Add("Company");
            foreach (string c in Companys)
            {
                DataRow dr = dt.NewRow();
                string[] car = c.Split(':');
                dr["Number"] = car[0];
                dr["Company"] = car[1];
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private DataTable ProductData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Letter");
            dt.Columns.Add("Product");
            dt.Columns.Add("Cost");
            foreach (string p in Products)
            {
                try
                {
                    Console.WriteLine(p);
                    DataRow dr = dt.NewRow();
                    string[] par = p.Split(':');
                    dr["Letter"] = par[0];
                    dr["Product"] = par[1];
                    dr["Cost"] = par[2];
                    dt.Rows.Add(dr);
                }
                catch { }
            }
            return dt;
        }

        private bool initialLoad = false;
        private string CurrentCompany;

        private void companyGrid_SelectionChanged(object sender, EventArgs e)
        {
            //DataRowView dvr = (DataRowView)companyGrid.SelectedCells[1].Value;
            //MessageBox.Show(companyGrid.SelectedCells[0].Value.ToString());
            //MessageBox.Show(dvr[1].ToString());
            if (initialLoad)
            {
                try
                {
                    Console.WriteLine(companyGrid.SelectedRows[0].Index);
                    CurrentCompany = companyGrid.SelectedCells[1].Value.ToString();
                    RefreshProducts(CurrentCompany);
                    productGrid.DataSource = ProductData();
                }
                catch { }
            }
        }

        private void ExcelessEditor_Shown(object sender, EventArgs e)
        {
            RefreshCompanys();
            initialLoad = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string CompanyToDelete = companyGrid.SelectedCells[1].Value.ToString();
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete the company " + CompanyToDelete + "?", "Confirm", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Companys.Remove(CompanyToDelete);
                SaveCompanys();
            }
        }

        private void SaveProducts(string Company)
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(INI);
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)companyGrid.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                data[Company][DataRow[0].ToString().ToUpper()] = String.Format("{0}:{1}:{2}", DataRow[0].ToString().ToUpper(), DataRow[1].ToString(), DataRow[2].ToString()) ?? "";
            }

            parser.WriteFile("Settings.ini", data);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveProducts(CurrentCompany);
        }
    }
}