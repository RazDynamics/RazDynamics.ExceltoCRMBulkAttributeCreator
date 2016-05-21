namespace CRMConsultants.CreateAttributes
{
    partial class ImportAttributesToCrm
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportAttributesToCrm));
            this.tsMain = new System.Windows.Forms.ToolStrip();
            this.tsbCloseThisTab = new System.Windows.Forms.ToolStripButton();
            this.tsbLoadEntities = new System.Windows.Forms.ToolStripButton();
            this.tsbCreateAttributes = new System.Windows.Forms.ToolStripButton();
            this.gbFile = new System.Windows.Forms.GroupBox();
            this.btnBrowseFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.cmbEntities = new System.Windows.Forms.ComboBox();
            this.gbEntity = new System.Windows.Forms.GroupBox();
            this.tsMain.SuspendLayout();
            this.gbFile.SuspendLayout();
            this.gbEntity.SuspendLayout();
            this.SuspendLayout();
            // 
            // tsMain
            // 
            this.tsMain.AutoSize = false;
            this.tsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbCloseThisTab,
            this.tsbLoadEntities,
            this.tsbCreateAttributes});
            this.tsMain.Location = new System.Drawing.Point(0, 0);
            this.tsMain.Name = "tsMain";
            this.tsMain.Size = new System.Drawing.Size(738, 25);
            this.tsMain.TabIndex = 100;
            this.tsMain.Text = "toolStrip1";
            // 
            // tsbCloseThisTab
            // 
            this.tsbCloseThisTab.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsbCloseThisTab.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.tsbCloseThisTab.Image = ((System.Drawing.Image)(resources.GetObject("tsbCloseThisTab.Image")));
            this.tsbCloseThisTab.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbCloseThisTab.Name = "tsbCloseThisTab";
            this.tsbCloseThisTab.Size = new System.Drawing.Size(23, 22);
            this.tsbCloseThisTab.Text = "Close this tab";
            // 
            // tsbLoadEntities
            // 
            this.tsbLoadEntities.Image = ((System.Drawing.Image)(resources.GetObject("tsbLoadEntities.Image")));
            this.tsbLoadEntities.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbLoadEntities.Name = "tsbLoadEntities";
            this.tsbLoadEntities.Size = new System.Drawing.Size(94, 22);
            this.tsbLoadEntities.Text = "Load Entities";
            this.tsbLoadEntities.Click += new System.EventHandler(this.tsbLoadEntities_Click);
            // 
            // tsbCreateAttributes
            // 
            this.tsbCreateAttributes.Enabled = false;
            this.tsbCreateAttributes.Image = ((System.Drawing.Image)(resources.GetObject("tsbCreateAttributes.Image")));
            this.tsbCreateAttributes.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbCreateAttributes.Name = "tsbCreateAttributes";
            this.tsbCreateAttributes.Size = new System.Drawing.Size(118, 22);
            this.tsbCreateAttributes.Text = "Import Attributes";
            this.tsbCreateAttributes.Click += new System.EventHandler(this.tsbCreateAttributes_Click);
            // 
            // gbFile
            // 
            this.gbFile.Controls.Add(this.btnBrowseFile);
            this.gbFile.Controls.Add(this.txtFilePath);
            this.gbFile.Location = new System.Drawing.Point(384, 32);
            this.gbFile.Name = "gbFile";
            this.gbFile.Size = new System.Drawing.Size(326, 55);
            this.gbFile.TabIndex = 101;
            this.gbFile.TabStop = false;
            this.gbFile.Text = "Download Template";
            // 
            // btnBrowseFile
            // 
            this.btnBrowseFile.Location = new System.Drawing.Point(205, 21);
            this.btnBrowseFile.Name = "btnBrowseFile";
            this.btnBrowseFile.Size = new System.Drawing.Size(49, 21);
            this.btnBrowseFile.TabIndex = 1;
            this.btnBrowseFile.Text = "......";
            this.btnBrowseFile.UseVisualStyleBackColor = true;
            this.btnBrowseFile.Click += new System.EventHandler(this.btnBrowseFile_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.BackColor = System.Drawing.SystemColors.Control;
            this.txtFilePath.Enabled = false;
            this.txtFilePath.Location = new System.Drawing.Point(35, 22);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(169, 20);
            this.txtFilePath.TabIndex = 0;
            // 
            // cmbEntities
            // 
            this.cmbEntities.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbEntities.Enabled = false;
            this.cmbEntities.FormattingEnabled = true;
            this.cmbEntities.Location = new System.Drawing.Point(10, 25);
            this.cmbEntities.Name = "cmbEntities";
            this.cmbEntities.Size = new System.Drawing.Size(272, 21);
            this.cmbEntities.TabIndex = 9;
            this.cmbEntities.SelectedIndexChanged += new System.EventHandler(this.cmbEntities_SelectedIndexChanged);
            // 
            // gbEntity
            // 
            this.gbEntity.Controls.Add(this.cmbEntities);
            this.gbEntity.Location = new System.Drawing.Point(21, 28);
            this.gbEntity.Name = "gbEntity";
            this.gbEntity.Size = new System.Drawing.Size(320, 59);
            this.gbEntity.TabIndex = 102;
            this.gbEntity.TabStop = false;
            this.gbEntity.Text = "Entity";
            // 
            // ImportAttributesToCrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.gbEntity);
            this.Controls.Add(this.gbFile);
            this.Controls.Add(this.tsMain);
            this.Name = "ImportAttributesToCrm";
            this.Size = new System.Drawing.Size(738, 248);
            this.tsMain.ResumeLayout(false);
            this.tsMain.PerformLayout();
            this.gbFile.ResumeLayout(false);
            this.gbFile.PerformLayout();
            this.gbEntity.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ToolStrip tsMain;
        private System.Windows.Forms.ToolStripButton tsbCloseThisTab;
        private System.Windows.Forms.ToolStripButton tsbCreateAttributes;
        private System.Windows.Forms.GroupBox gbFile;
        private System.Windows.Forms.Button btnBrowseFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.ComboBox cmbEntities;
        private System.Windows.Forms.GroupBox gbEntity;
        private System.Windows.Forms.ToolStripButton tsbLoadEntities;
    }
}
