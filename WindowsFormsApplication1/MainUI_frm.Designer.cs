namespace WindowsFormsApplication1
{
    partial class MainUI_frm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainUI_frm));
            this.crBI_btn = new System.Windows.Forms.Button();
            this.Logo = new System.Windows.Forms.PictureBox();
            this.dt_lbl = new System.Windows.Forms.Label();
            this.fndJO_txbx = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.sheet2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ctlDB_tbctl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.createBillingInvoiceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editBillingInvoiceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sheet2BindingSource)).BeginInit();
            this.ctlDB_tbctl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // crBI_btn
            // 
            this.crBI_btn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.crBI_btn.Location = new System.Drawing.Point(553, 445);
            this.crBI_btn.Name = "crBI_btn";
            this.crBI_btn.Size = new System.Drawing.Size(124, 67);
            this.crBI_btn.TabIndex = 1;
            this.crBI_btn.Text = "Create Billing Invoice";
            this.crBI_btn.UseVisualStyleBackColor = true;
            this.crBI_btn.Click += new System.EventHandler(this.crBI_btn_Click);
            // 
            // Logo
            // 
            this.Logo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.Logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Logo.Image = ((System.Drawing.Image)(resources.GetObject("Logo.Image")));
            this.Logo.Location = new System.Drawing.Point(219, 40);
            this.Logo.Name = "Logo";
            this.Logo.Size = new System.Drawing.Size(250, 64);
            this.Logo.TabIndex = 3;
            this.Logo.TabStop = false;
            // 
            // dt_lbl
            // 
            this.dt_lbl.Dock = System.Windows.Forms.DockStyle.Top;
            this.dt_lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dt_lbl.Location = new System.Drawing.Point(0, 24);
            this.dt_lbl.Name = "dt_lbl";
            this.dt_lbl.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.dt_lbl.Size = new System.Drawing.Size(689, 26);
            this.dt_lbl.TabIndex = 4;
            this.dt_lbl.Text = "Date";
            this.dt_lbl.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // fndJO_txbx
            // 
            this.fndJO_txbx.Location = new System.Drawing.Point(81, 123);
            this.fndJO_txbx.Name = "fndJO_txbx";
            this.fndJO_txbx.Size = new System.Drawing.Size(168, 20);
            this.fndJO_txbx.TabIndex = 6;
            this.fndJO_txbx.TextChanged += new System.EventHandler(this.fndJO_txbx_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Search J.O.";
            // 
            // sheet2BindingSource
            // 
            this.sheet2BindingSource.DataMember = "Sheet2";
            // 
            // ctlDB_tbctl
            // 
            this.ctlDB_tbctl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.ctlDB_tbctl.Controls.Add(this.tabPage1);
            this.ctlDB_tbctl.Controls.Add(this.tabPage3);
            this.ctlDB_tbctl.Location = new System.Drawing.Point(12, 149);
            this.ctlDB_tbctl.Name = "ctlDB_tbctl";
            this.ctlDB_tbctl.SelectedIndex = 0;
            this.ctlDB_tbctl.Size = new System.Drawing.Size(665, 290);
            this.ctlDB_tbctl.TabIndex = 9;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(657, 264);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "NEW";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(3, 3);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(651, 258);
            this.dataGridView2.TabIndex = 0;
            this.dataGridView2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataGridView1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(657, 264);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "ALL";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(657, 264);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.optionsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(689, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.createBillingInvoiceToolStripMenuItem,
            this.editBillingInvoiceToolStripMenuItem,
            this.optionsToolStripMenuItem1,
            this.toolStripMenuItem3,
            this.exitToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.optionsToolStripMenuItem.Text = "File";
            // 
            // createBillingInvoiceToolStripMenuItem
            // 
            this.createBillingInvoiceToolStripMenuItem.Name = "createBillingInvoiceToolStripMenuItem";
            this.createBillingInvoiceToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
            this.createBillingInvoiceToolStripMenuItem.Text = "Create/Edit...";
            // 
            // editBillingInvoiceToolStripMenuItem
            // 
            this.editBillingInvoiceToolStripMenuItem.Name = "editBillingInvoiceToolStripMenuItem";
            this.editBillingInvoiceToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
            this.editBillingInvoiceToolStripMenuItem.Text = "Rescan Database";
            // 
            // optionsToolStripMenuItem1
            // 
            this.optionsToolStripMenuItem1.Name = "optionsToolStripMenuItem1";
            this.optionsToolStripMenuItem1.Size = new System.Drawing.Size(162, 22);
            this.optionsToolStripMenuItem1.Text = "Options";
            this.optionsToolStripMenuItem1.Click += new System.EventHandler(this.optionsToolStripMenuItem1_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(159, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.aboutToolStripMenuItem.Text = "About";
            // 
            // MainUI_frm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(689, 528);
            this.Controls.Add(this.ctlDB_tbctl);
            this.Controls.Add(this.Logo);
            this.Controls.Add(this.dt_lbl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.fndJO_txbx);
            this.Controls.Add(this.crBI_btn);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainUI_frm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Biling Invoice";
            this.Load += new System.EventHandler(this.MainUI_frm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Logo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sheet2BindingSource)).EndInit();
            this.ctlDB_tbctl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button crBI_btn;
        private System.Windows.Forms.PictureBox Logo;
        private System.Windows.Forms.Label dt_lbl;
        private System.Windows.Forms.TextBox fndJO_txbx;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.BindingSource sheet2BindingSource;
        private System.Windows.Forms.TabControl ctlDB_tbctl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem createBillingInvoiceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editBillingInvoiceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.EventHandler createBillingInvoiceToolStripMenuItem_Click;
        private System.EventHandler toolStripMenuItem3_Click;
        private System.Windows.Forms.DataGridView dataGridView2;
    }
}

