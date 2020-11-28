namespace DataGridViewVirtualModeTest
{
    partial class Form1
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
            System.Windows.Forms.Button cmdTestA;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.cmdPromote = new System.Windows.Forms.Button();
            this.cmdDemote = new System.Windows.Forms.Button();
            this.dgvSheetCategoriesX = new System.Windows.Forms.DataGridView();
            this.dgvSheetCategoriesY = new System.Windows.Forms.DataGridView();
            this.cmdRender = new System.Windows.Forms.Button();
            this.dgCells = new System.Windows.Forms.DataGrid();
            this.cmdTestB = new System.Windows.Forms.Button();
            this.dgvSheetCategoriesFilter = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            cmdTestA = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesX)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesY)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgCells)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesFilter)).BeginInit();
            this.SuspendLayout();
            // 
            // cmdTestA
            // 
            cmdTestA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            cmdTestA.Location = new System.Drawing.Point(525, 401);
            cmdTestA.Name = "cmdTestA";
            cmdTestA.Size = new System.Drawing.Size(64, 23);
            cmdTestA.TabIndex = 25;
            cmdTestA.Text = "TestA";
            cmdTestA.UseVisualStyleBackColor = true;
            cmdTestA.Click += new System.EventHandler(this.cmdTest1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.ColumnHeadersVisible = false;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Location = new System.Drawing.Point(204, 61);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridView1.Size = new System.Drawing.Size(233, 334);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.VirtualMode = true;
            // 
            // dataGridView2
            // 
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(12, 28);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(186, 367);
            this.dataGridView2.TabIndex = 1;
            // 
            // cmdPromote
            // 
            this.cmdPromote.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdPromote.Location = new System.Drawing.Point(12, 404);
            this.cmdPromote.Name = "cmdPromote";
            this.cmdPromote.Size = new System.Drawing.Size(75, 23);
            this.cmdPromote.TabIndex = 2;
            this.cmdPromote.Text = "Promote";
            this.cmdPromote.UseVisualStyleBackColor = true;
            this.cmdPromote.Click += new System.EventHandler(this.cmdPromote_Click);
            // 
            // cmdDemote
            // 
            this.cmdDemote.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdDemote.Location = new System.Drawing.Point(93, 404);
            this.cmdDemote.Name = "cmdDemote";
            this.cmdDemote.Size = new System.Drawing.Size(75, 23);
            this.cmdDemote.TabIndex = 3;
            this.cmdDemote.Text = "Demote";
            this.cmdDemote.UseVisualStyleBackColor = true;
            this.cmdDemote.Click += new System.EventHandler(this.cmdDemote_Click);
            // 
            // dgvSheetCategoriesX
            // 
            this.dgvSheetCategoriesX.AllowUserToAddRows = false;
            this.dgvSheetCategoriesX.AllowUserToDeleteRows = false;
            this.dgvSheetCategoriesX.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSheetCategoriesX.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dgvSheetCategoriesX.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSheetCategoriesX.ColumnHeadersVisible = false;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvSheetCategoriesX.DefaultCellStyle = dataGridViewCellStyle5;
            this.dgvSheetCategoriesX.Location = new System.Drawing.Point(443, 61);
            this.dgvSheetCategoriesX.MultiSelect = false;
            this.dgvSheetCategoriesX.Name = "dgvSheetCategoriesX";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSheetCategoriesX.RowHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.dgvSheetCategoriesX.RowHeadersVisible = false;
            this.dgvSheetCategoriesX.Size = new System.Drawing.Size(76, 334);
            this.dgvSheetCategoriesX.TabIndex = 24;
            this.dgvSheetCategoriesX.VirtualMode = true;
            this.dgvSheetCategoriesX.CellValueNeeded += new System.Windows.Forms.DataGridViewCellValueEventHandler(this.dgvSheetCategoriesX_CellValueNeeded);
            // 
            // dgvSheetCategoriesY
            // 
            this.dgvSheetCategoriesY.AllowUserToAddRows = false;
            this.dgvSheetCategoriesY.AllowUserToDeleteRows = false;
            this.dgvSheetCategoriesY.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSheetCategoriesY.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSheetCategoriesY.ColumnHeadersVisible = false;
            this.dgvSheetCategoriesY.Location = new System.Drawing.Point(204, 401);
            this.dgvSheetCategoriesY.MultiSelect = false;
            this.dgvSheetCategoriesY.Name = "dgvSheetCategoriesY";
            this.dgvSheetCategoriesY.RowHeadersVisible = false;
            this.dgvSheetCategoriesY.Size = new System.Drawing.Size(233, 25);
            this.dgvSheetCategoriesY.TabIndex = 23;
            this.dgvSheetCategoriesY.VirtualMode = true;
            this.dgvSheetCategoriesY.CellValueNeeded += new System.Windows.Forms.DataGridViewCellValueEventHandler(this.dgvSheetCategoriesY_CellValueNeeded);
            // 
            // cmdRender
            // 
            this.cmdRender.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdRender.Location = new System.Drawing.Point(444, 401);
            this.cmdRender.Name = "cmdRender";
            this.cmdRender.Size = new System.Drawing.Size(75, 23);
            this.cmdRender.TabIndex = 25;
            this.cmdRender.Text = "Render";
            this.cmdRender.UseVisualStyleBackColor = true;
            this.cmdRender.Click += new System.EventHandler(this.cmdRender_Click);
            // 
            // dgCells
            // 
            this.dgCells.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgCells.DataMember = "";
            this.dgCells.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dgCells.Location = new System.Drawing.Point(525, 28);
            this.dgCells.Name = "dgCells";
            this.dgCells.Size = new System.Drawing.Size(133, 367);
            this.dgCells.TabIndex = 26;
            // 
            // cmdTestB
            // 
            this.cmdTestB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdTestB.Location = new System.Drawing.Point(595, 401);
            this.cmdTestB.Name = "cmdTestB";
            this.cmdTestB.Size = new System.Drawing.Size(63, 23);
            this.cmdTestB.TabIndex = 25;
            this.cmdTestB.Text = "TestB";
            this.cmdTestB.UseVisualStyleBackColor = true;
            this.cmdTestB.Click += new System.EventHandler(this.cmdTest2_Click);
            // 
            // dgvSheetCategoriesFilter
            // 
            this.dgvSheetCategoriesFilter.AllowUserToAddRows = false;
            this.dgvSheetCategoriesFilter.AllowUserToDeleteRows = false;
            this.dgvSheetCategoriesFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSheetCategoriesFilter.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.dgvSheetCategoriesFilter.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSheetCategoriesFilter.ColumnHeadersVisible = false;
            this.dgvSheetCategoriesFilter.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvSheetCategoriesFilter.DefaultCellStyle = dataGridViewCellStyle8;
            this.dgvSheetCategoriesFilter.Location = new System.Drawing.Point(204, 30);
            this.dgvSheetCategoriesFilter.MultiSelect = false;
            this.dgvSheetCategoriesFilter.Name = "dgvSheetCategoriesFilter";
            this.dgvSheetCategoriesFilter.ReadOnly = true;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSheetCategoriesFilter.RowHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.dgvSheetCategoriesFilter.RowHeadersVisible = false;
            this.dgvSheetCategoriesFilter.Size = new System.Drawing.Size(233, 25);
            this.dgvSheetCategoriesFilter.TabIndex = 27;
            this.dgvSheetCategoriesFilter.VirtualMode = true;
            this.dgvSheetCategoriesFilter.CellValueNeeded += new System.Windows.Forms.DataGridViewCellValueEventHandler(this.dgvSheetCategoriesFilter_CellValueNeeded);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(670, 439);
            this.Controls.Add(this.dgvSheetCategoriesFilter);
            this.Controls.Add(this.dgCells);
            this.Controls.Add(this.cmdTestB);
            this.Controls.Add(cmdTestA);
            this.Controls.Add(this.cmdRender);
            this.Controls.Add(this.dgvSheetCategoriesX);
            this.Controls.Add(this.dgvSheetCategoriesY);
            this.Controls.Add(this.cmdDemote);
            this.Controls.Add(this.cmdPromote);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesX)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesY)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgCells)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSheetCategoriesFilter)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button cmdPromote;
        private System.Windows.Forms.Button cmdDemote;
        private System.Windows.Forms.DataGridView dgvSheetCategoriesX;
        private System.Windows.Forms.DataGridView dgvSheetCategoriesY;
        private System.Windows.Forms.Button cmdRender;
        private System.Windows.Forms.DataGrid dgCells;
        private System.Windows.Forms.Button cmdTestB;
        private System.Windows.Forms.DataGridView dgvSheetCategoriesFilter;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
    }
}

