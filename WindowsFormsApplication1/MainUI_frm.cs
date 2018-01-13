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
            DGView.OledbCon(true,ProcDBPath, ref oleAccessCon);
            DGView.datGrViewDS(dataGridView1, ref oleAccessCon, "TotalProcData");
            oleAccessCon.Close();
            //close Connection

            //populate datatable with company info
            DGView.BIsource(BIPath);
            this.dataGridView2.DataSource = DGView.dtcompany;



        }
        private void optionsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
           Options_frm ldOpt = new Options_frm();
           ldOpt.Visible = true;

        }

    }

    public class Setup:MainUI_frm
    {

        public DataTable dtcompany = new DataTable();
        public void BIsource(string Path)
         {
            //creates a datatable describing the relation of the billing invoice transactions
             string[] arrtemp = Directory.GetDirectories(Path);
            List<string> compList = new List<string>();
            List<string> transactlist = new List<string>();
             foreach (string s in arrtemp)
             {
                compList.Add(s.Substring(Path.Length+1));
             }
            foreach (string s in compList)
            {
                dtcompany.Columns.Add(s);
                string sPath = Path+s;
                string[] trtemp = Directory.GetDirectories(sPath);
                foreach (string t in trtemp)
                {
                    transactlist.Add(t.Substring(sPath.Length+1));
                }
                foreach (string t in transactlist)
                {
                    
                }
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
    }
    public class Operation
    { }


}
