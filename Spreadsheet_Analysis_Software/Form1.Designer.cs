
namespace Spreadsheet_Analysis_Software
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
            this.spreadsheet_dir_textbox = new System.Windows.Forms.TextBox();
            this.import_spreadsheet_button = new System.Windows.Forms.Button();
            this.export_spreadsheet_button = new System.Windows.Forms.Button();
            this.set_export_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // spreadsheet_dir_textbox
            // 
            this.spreadsheet_dir_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.spreadsheet_dir_textbox.Location = new System.Drawing.Point(182, 18);
            this.spreadsheet_dir_textbox.Name = "spreadsheet_dir_textbox";
            this.spreadsheet_dir_textbox.ReadOnly = true;
            this.spreadsheet_dir_textbox.Size = new System.Drawing.Size(286, 26);
            this.spreadsheet_dir_textbox.TabIndex = 0;
            // 
            // import_spreadsheet_button
            // 
            this.import_spreadsheet_button.Location = new System.Drawing.Point(12, 12);
            this.import_spreadsheet_button.Name = "import_spreadsheet_button";
            this.import_spreadsheet_button.Size = new System.Drawing.Size(164, 39);
            this.import_spreadsheet_button.TabIndex = 1;
            this.import_spreadsheet_button.Text = "Import Spreadsheet";
            this.import_spreadsheet_button.UseVisualStyleBackColor = true;
            this.import_spreadsheet_button.Click += new System.EventHandler(this.import_spreadsheet_button_Click);
            // 
            // export_spreadsheet_button
            // 
            this.export_spreadsheet_button.Enabled = false;
            this.export_spreadsheet_button.Location = new System.Drawing.Point(12, 57);
            this.export_spreadsheet_button.Name = "export_spreadsheet_button";
            this.export_spreadsheet_button.Size = new System.Drawing.Size(286, 39);
            this.export_spreadsheet_button.TabIndex = 1;
            this.export_spreadsheet_button.Text = "Export Spreadsheet";
            this.export_spreadsheet_button.UseVisualStyleBackColor = true;
            this.export_spreadsheet_button.Click += new System.EventHandler(this.export_spreadsheet_button_Click);
            // 
            // set_export_button
            // 
            this.set_export_button.Location = new System.Drawing.Point(304, 57);
            this.set_export_button.Name = "set_export_button";
            this.set_export_button.Size = new System.Drawing.Size(164, 39);
            this.set_export_button.TabIndex = 1;
            this.set_export_button.Text = "Set Export Location";
            this.set_export_button.UseVisualStyleBackColor = true;
            this.set_export_button.Click += new System.EventHandler(this.set_export_button_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(481, 107);
            this.Controls.Add(this.set_export_button);
            this.Controls.Add(this.export_spreadsheet_button);
            this.Controls.Add(this.import_spreadsheet_button);
            this.Controls.Add(this.spreadsheet_dir_textbox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "MSS Software Spreadsheets";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox spreadsheet_dir_textbox;
        private System.Windows.Forms.Button import_spreadsheet_button;
        private System.Windows.Forms.Button export_spreadsheet_button;
        private System.Windows.Forms.Button set_export_button;
    }
}

