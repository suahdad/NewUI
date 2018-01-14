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





namespace WindowsFormsApplication1
{

    public partial class MainUI_frm : Form
    {
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

    }

    public class Setup:MainUI_frm
    {

        public DataTable dtcompany = new DataTable();
        public void BIsource(string Path, ref DataTable dtcompany)
         {
            //creates a datatable describing the relation of the billing invoice transactions
             string[] arrtemp = Directory.GetDirectories(Path);
            foreach (string s in arrtemp)
            {
                dtcompany.Columns.Add(s.Substring(Path.Length));
                
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
