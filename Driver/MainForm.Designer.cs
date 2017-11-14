namespace Driver
{
	partial class MainForm
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
			this.executeButton = new System.Windows.Forms.Button();
			this.closeButton = new System.Windows.Forms.Button();
			this.splitFileButton = new System.Windows.Forms.Button();
			this.xmlToCsvButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// executeButton
			// 
			this.executeButton.Location = new System.Drawing.Point(242, 143);
			this.executeButton.Name = "executeButton";
			this.executeButton.Size = new System.Drawing.Size(75, 23);
			this.executeButton.TabIndex = 0;
			this.executeButton.Text = "&Engage!";
			this.executeButton.UseVisualStyleBackColor = true;
			this.executeButton.Click += new System.EventHandler(this.executeButton_Click);
			// 
			// closeButton
			// 
			this.closeButton.Location = new System.Drawing.Point(465, 315);
			this.closeButton.Name = "closeButton";
			this.closeButton.Size = new System.Drawing.Size(75, 23);
			this.closeButton.TabIndex = 1;
			this.closeButton.Text = "Close";
			this.closeButton.UseVisualStyleBackColor = true;
			this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
			// 
			// splitFileButton
			// 
			this.splitFileButton.Location = new System.Drawing.Point(242, 46);
			this.splitFileButton.Name = "splitFileButton";
			this.splitFileButton.Size = new System.Drawing.Size(75, 23);
			this.splitFileButton.TabIndex = 2;
			this.splitFileButton.Text = "&Split file";
			this.splitFileButton.UseVisualStyleBackColor = true;
			this.splitFileButton.Click += new System.EventHandler(this.splitFileButton_Click);
			// 
			// xmlToCsvButton
			// 
			this.xmlToCsvButton.Location = new System.Drawing.Point(242, 94);
			this.xmlToCsvButton.Name = "xmlToCsvButton";
			this.xmlToCsvButton.Size = new System.Drawing.Size(75, 23);
			this.xmlToCsvButton.TabIndex = 3;
			this.xmlToCsvButton.Text = "&XML to CSV";
			this.xmlToCsvButton.UseVisualStyleBackColor = true;
			this.xmlToCsvButton.Click += new System.EventHandler(this.xmlToCsvButton_Click);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(552, 350);
			this.Controls.Add(this.xmlToCsvButton);
			this.Controls.Add(this.splitFileButton);
			this.Controls.Add(this.closeButton);
			this.Controls.Add(this.executeButton);
			this.Name = "MainForm";
			this.Text = "MainForm";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button executeButton;
		private System.Windows.Forms.Button closeButton;
		private System.Windows.Forms.Button splitFileButton;
		private System.Windows.Forms.Button xmlToCsvButton;
	}
}

