namespace Meyerhoff_Family_Tree
{
    partial class FamilyForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FamilyForm));
            this.descriptionTextBox = new System.Windows.Forms.TextBox();
            this.clipboardButton = new System.Windows.Forms.Button();
            this.exitButton = new System.Windows.Forms.Button();
            this.titleLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // descriptionTextBox
            // 
            this.descriptionTextBox.AcceptsReturn = true;
            this.descriptionTextBox.Font = new System.Drawing.Font("Courier New", 8.25F);
            this.descriptionTextBox.Location = new System.Drawing.Point(12, 57);
            this.descriptionTextBox.Multiline = true;
            this.descriptionTextBox.Name = "descriptionTextBox";
            this.descriptionTextBox.ReadOnly = true;
            this.descriptionTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.descriptionTextBox.Size = new System.Drawing.Size(459, 408);
            this.descriptionTextBox.TabIndex = 0;
            // 
            // clipboardButton
            // 
            this.clipboardButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clipboardButton.Location = new System.Drawing.Point(12, 471);
            this.clipboardButton.Name = "clipboardButton";
            this.clipboardButton.Size = new System.Drawing.Size(227, 49);
            this.clipboardButton.TabIndex = 1;
            this.clipboardButton.Text = "Copy to Clipboard";
            this.clipboardButton.UseVisualStyleBackColor = true;
            this.clipboardButton.Click += new System.EventHandler(this.clipboardButton_Click);
            // 
            // exitButton
            // 
            this.exitButton.Location = new System.Drawing.Point(244, 471);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(227, 49);
            this.exitButton.TabIndex = 2;
            this.exitButton.Text = "Exit";
            this.exitButton.UseVisualStyleBackColor = true;
            this.exitButton.Click += new System.EventHandler(this.exitButton_Click);
            // 
            // titleLabel
            // 
            this.titleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLabel.Location = new System.Drawing.Point(12, 9);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(459, 45);
            this.titleLabel.TabIndex = 3;
            this.titleLabel.Text = "FAMILY OF JOHN DOE";
            this.titleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FamilyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 532);
            this.Controls.Add(this.titleLabel);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.clipboardButton);
            this.Controls.Add(this.descriptionTextBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FamilyForm";
            this.Text = "Family";
            this.Load += new System.EventHandler(this.FamilyForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox descriptionTextBox;
        private System.Windows.Forms.Button clipboardButton;
        private System.Windows.Forms.Button exitButton;
        private System.Windows.Forms.Label titleLabel;
    }
}