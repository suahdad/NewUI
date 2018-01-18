using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


//UPDATED JC CODE 1/18/2018


namespace WindowsFormsApplication1
{

    public partial class MainUI_frm : Form
    {
        string[] files = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe522");// interchange these
        string[] files2 = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe523");// path's to change to the
        string[] files3 = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe526");//necessary paths needed 
        public string ProcDBPath;
        public string BIPath;
        static void Main()
        {
            Application.Run(new MainUI_frm());
        }
        public MainUI_frm()
        {

            InitializeComponent();
            ProcDBPath = @"C:\Users\Arandov\Desktop\Database5.accdb";
            BIPath = @"C:\Users\Arandov\Desktop\Billing Invoice 2018";
        }
        private void MainUI_frm_Load(object sender, EventArgs e)
        {
            dt_lbl.Text = DateTime.Now.ToString("MMMM dd, yyyy");

            //setup the Data Grid Views
            OleDbConnection oleAccessCon = new OleDbConnection();
            Setup DGView = new Setup();
            DGView.datvewSetup(true, ProcDBPath, ref oleAccessCon, dataGridView1, "TotalProcData");
            //close Connection
            DataTable tempData = new DataTable();
            DGView.BIsource(BIPath, ref tempData);
            dataGridView2.DataSource = tempData;
            {
                //this whole section should be able to view file in the directory "C:\Users\REYNOSO\Desktop\cpe522"
                string[] files = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\cpe522");// change directory to one needed
                DataTable table = new DataTable(); //not sure about this line
                for (int i = 0; i < files.Length; i++)
                {
                    FileInfo file = new FileInfo(files[i]);
                    table.Rows.Add(file.Name);
                }
                dataGridView2.DataSource = table;
            }

            


        }
        private void optionsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
           Options_frm ldOpt = new Options_frm();
           ldOpt.Visible = true;

        }

        private void fndJO_txbx_TextChanged(object sender, EventArgs e)
        {

           
        }

        private void crBI_btn_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

    }

    public class Setup:MainUI_frm
    {
        //PPART OF JC CODE ( FIRST 3 CODES STARTING WITH STRING[]) USED TO LOCATE THE PATH THAT WILL BE USED
        string[] files = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe522");// interchange these
        string[] files2 = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe523");// path's to change to the
        string[] files3 = Directory.GetFiles(@"C:\Users\REYNOSO\Desktop\FUSO\cpe526");//necessary paths needed 
        public DataTable dtcompany = new DataTable();
        public void BIsource(string Path, ref DataTable dtcompany)
         {
            //creates a datatable describing the relation of the billing invoice transactions
             string[] arrtemp = Directory.GetDirectories(Path);
            foreach (string s in arrtemp)
            {
                dtcompany.Columns.Add(s.Substring(Path.Length));
                //jc code start
                //addcolumn; column = folder
                dtcompany.Columns.Add("NAME OF FOLDER1");
                dtcompany.Columns.Add("NAME OF FOLDER2");
                dtcompany.Columns.Add("NAME OF FOLDER3");
                //addrows
                for (int i = 0; i < files.LongLength; i++)
                {
                    FileInfo file = new FileInfo(files[i]);//file is a variable representing
                    DataRow row1 = dtcompany.NewRow();      //the contents of a certain 
                    row1["NAME OF FOLDER1"] = file.Name;      //path necessary
                    dtcompany.Rows.Add(row1);
                }
                for (int a = 0; a < files2.Length; a++)
                {
                    FileInfo file2 = new FileInfo(files2[a]);
                    DataRow row2 = dtcompany.NewRow();
                    row2["NAME OF FOLDER2"] = file2.Name;
                    dtcompany.Rows.Add(row2);
                }
                for (int b = 0; b < files3.Length; b++)
                {
                    FileInfo file3 = new FileInfo(files3[b]);
                    DataRow row3 = dtcompany.NewRow();
                    row3["NAME OF FOLDER3"] = file3.Name;
                    dtcompany.Rows.Add(row3);
                }
                //JC CODE END
            }
            
         }
        public void OledbCon(bool access, string Setuppath , ref OleDbConnection acsCon)
        {
            //this function opens a connection
            string acs;
            if (access == true)
            {
                acs = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Setuppath + ";Persist Security Info=False;";
            }
            else
            {
                acs = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Setuppath + ";Extended Properties='Excel 12.0 Xml;HDR=YES';";
            }
            acsCon = new OleDbConnection(acs);
            acsCon.Open();


        }
        public void datGrViewDS(DataGridView dat, ref OleDbConnection datcon, string dattable)
        {//this function connects a data source to a datagridview
            string adapt = "Select * from " + dattable;
            OleDbDataAdapter datadapt = new OleDbDataAdapter(adapt, datcon);
            DataTable datdt = new DataTable();
            datadapt.Fill(datdt);
            dat.DataSource = datdt;
            
        }
        public void datvewSetup(bool access, string setuppath, ref OleDbConnection acsCon, DataGridView dat, string dattable)
        {
            OledbCon(access: access, Setuppath: setuppath, acsCon: ref acsCon);
            datGrViewDS(dat:dat,datcon: ref acsCon, dattable: dattable);
        }

        
    }
    public class Operation
    {
        Excel.Application objExcel = new Excel.Application();
        Excel.Workbook xlWkbook = new Excel.Workbook();
        Excel.Worksheet xlWksheet = new Excel.Worksheet();


    public void crBillingInvoice(string BIpath, string Company, string Transaction, string BInumber)
    {
        
    }

    }


}
