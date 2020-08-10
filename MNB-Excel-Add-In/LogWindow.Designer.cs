namespace MNB_Excel_Add_In
{
    partial class LogWindow
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.domainUsernameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timestampDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.commentDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mNBButtonLogsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.excelButtonDataSet = new MNB_Excel_Add_In.ExcelButtonDataSet();
            this.saveBtn = new System.Windows.Forms.Button();
            this.exitBtn = new System.Windows.Forms.Button();
            this.mNBButtonLogsTableAdapter = new MNB_Excel_Add_In.ExcelButtonDataSetTableAdapters.MNBButtonLogsTableAdapter();
            this.tableAdapterManager = new MNB_Excel_Add_In.ExcelButtonDataSetTableAdapters.TableAdapterManager();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mNBButtonLogsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelButtonDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.domainUsernameDataGridViewTextBoxColumn,
            this.timestampDataGridViewTextBoxColumn,
            this.commentDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.mNBButtonLogsBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(537, 413);
            this.dataGridView1.TabIndex = 0;
            // 
            // domainUsernameDataGridViewTextBoxColumn
            // 
            this.domainUsernameDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.domainUsernameDataGridViewTextBoxColumn.DataPropertyName = "DomainUsername";
            this.domainUsernameDataGridViewTextBoxColumn.HeaderText = "Domain Username";
            this.domainUsernameDataGridViewTextBoxColumn.Name = "domainUsernameDataGridViewTextBoxColumn";
            this.domainUsernameDataGridViewTextBoxColumn.ReadOnly = true;
            this.domainUsernameDataGridViewTextBoxColumn.Width = 109;
            // 
            // timestampDataGridViewTextBoxColumn
            // 
            this.timestampDataGridViewTextBoxColumn.DataPropertyName = "Timestamp";
            this.timestampDataGridViewTextBoxColumn.HeaderText = "Timestamp";
            this.timestampDataGridViewTextBoxColumn.Name = "timestampDataGridViewTextBoxColumn";
            this.timestampDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // commentDataGridViewTextBoxColumn
            // 
            this.commentDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.commentDataGridViewTextBoxColumn.DataPropertyName = "Comment";
            this.commentDataGridViewTextBoxColumn.HeaderText = "Comment";
            this.commentDataGridViewTextBoxColumn.Name = "commentDataGridViewTextBoxColumn";
            // 
            // mNBButtonLogsBindingSource
            // 
            this.mNBButtonLogsBindingSource.DataMember = "MNBButtonLogs";
            this.mNBButtonLogsBindingSource.DataSource = this.excelButtonDataSet;
            // 
            // excelButtonDataSet
            // 
            this.excelButtonDataSet.DataSetName = "ExcelButtonDataSet";
            this.excelButtonDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // saveBtn
            // 
            this.saveBtn.Location = new System.Drawing.Point(596, 119);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(75, 23);
            this.saveBtn.TabIndex = 1;
            this.saveBtn.Text = "Save";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // exitBtn
            // 
            this.exitBtn.Location = new System.Drawing.Point(586, 215);
            this.exitBtn.Name = "exitBtn";
            this.exitBtn.Size = new System.Drawing.Size(94, 40);
            this.exitBtn.TabIndex = 2;
            this.exitBtn.Text = "Exit";
            this.exitBtn.UseVisualStyleBackColor = true;
            this.exitBtn.Click += new System.EventHandler(this.exitBtn_Click);
            // 
            // mNBButtonLogsTableAdapter
            // 
            this.mNBButtonLogsTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.MNBButtonLogsTableAdapter = this.mNBButtonLogsTableAdapter;
            this.tableAdapterManager.UpdateOrder = MNB_Excel_Add_In.ExcelButtonDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // LogWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 450);
            this.Controls.Add(this.exitBtn);
            this.Controls.Add(this.saveBtn);
            this.Controls.Add(this.dataGridView1);
            this.Name = "LogWindow";
            this.Text = "LogWindow";
            this.Load += new System.EventHandler(this.LogWindow_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mNBButtonLogsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelButtonDataSet)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private ExcelButtonDataSet excelButtonDataSet;
        private System.Windows.Forms.BindingSource mNBButtonLogsBindingSource;
        private ExcelButtonDataSetTableAdapters.MNBButtonLogsTableAdapter mNBButtonLogsTableAdapter;
        private ExcelButtonDataSetTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.DataGridViewTextBoxColumn domainUsernameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn timestampDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn commentDataGridViewTextBoxColumn;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.Button exitBtn;
    }
}