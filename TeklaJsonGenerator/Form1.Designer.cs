namespace TeklaJsonGenerator
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.linkLabelApp = new System.Windows.Forms.LinkLabel();
            this.linkLabelJson = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.structuresExtender.SetAttributeName(this.cancelButton, null);
            this.structuresExtender.SetAttributeTypeName(this.cancelButton, null);
            this.structuresExtender.SetBindPropertyName(this.cancelButton, null);
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(205, 60);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 0;
            this.cancelButton.Text = "Закрыть";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.structuresExtender.SetAttributeName(this.okButton, null);
            this.structuresExtender.SetAttributeTypeName(this.okButton, null);
            this.structuresExtender.SetBindPropertyName(this.okButton, null);
            this.okButton.Location = new System.Drawing.Point(124, 60);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 1;
            this.okButton.Text = "Экспорт";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // linkLabelApp
            // 
            this.structuresExtender.SetAttributeName(this.linkLabelApp, null);
            this.structuresExtender.SetAttributeTypeName(this.linkLabelApp, null);
            this.linkLabelApp.AutoSize = true;
            this.structuresExtender.SetBindPropertyName(this.linkLabelApp, null);
            this.linkLabelApp.Location = new System.Drawing.Point(12, 9);
            this.linkLabelApp.Name = "linkLabelApp";
            this.linkLabelApp.Size = new System.Drawing.Size(266, 13);
            this.linkLabelApp.TabIndex = 2;
            this.linkLabelApp.TabStop = true;
            this.linkLabelApp.Text = "Открыть программу формирования спецификации";
            this.linkLabelApp.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelApp_LinkClicked);
            // 
            // linkLabelJson
            // 
            this.structuresExtender.SetAttributeName(this.linkLabelJson, null);
            this.structuresExtender.SetAttributeTypeName(this.linkLabelJson, null);
            this.linkLabelJson.AutoSize = true;
            this.structuresExtender.SetBindPropertyName(this.linkLabelJson, null);
            this.linkLabelJson.Location = new System.Drawing.Point(12, 31);
            this.linkLabelJson.Name = "linkLabelJson";
            this.linkLabelJson.Size = new System.Drawing.Size(157, 13);
            this.linkLabelJson.TabIndex = 3;
            this.linkLabelJson.TabStop = true;
            this.linkLabelJson.Text = "Открыть файл спецификации";
            this.linkLabelJson.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelJson_LinkClicked);
            // 
            // Form1
            // 
            this.structuresExtender.SetAttributeName(this, null);
            this.structuresExtender.SetAttributeTypeName(this, null);
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.structuresExtender.SetBindPropertyName(this, null);
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(292, 95);
            this.Controls.Add(this.linkLabelJson);
            this.Controls.Add(this.linkLabelApp);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VNP: Спецификация металлопроката";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.LinkLabel linkLabelApp;
        private System.Windows.Forms.LinkLabel linkLabelJson;
    }
}

