namespace RSGenerate
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnChooseTemplate = new System.Windows.Forms.Button();
            this.btnChooseDefinition = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGenerateOutput = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnChooseTemplate
            // 
            this.btnChooseTemplate.Location = new System.Drawing.Point(12, 12);
            this.btnChooseTemplate.Name = "btnChooseTemplate";
            this.btnChooseTemplate.Size = new System.Drawing.Size(154, 23);
            this.btnChooseTemplate.TabIndex = 1;
            this.btnChooseTemplate.Text = "Load Template...";
            this.btnChooseTemplate.UseVisualStyleBackColor = true;
            this.btnChooseTemplate.Click += new System.EventHandler(this.btnChooseTemplate_Click);
            // 
            // btnChooseDefinition
            // 
            this.btnChooseDefinition.Location = new System.Drawing.Point(12, 41);
            this.btnChooseDefinition.Name = "btnChooseDefinition";
            this.btnChooseDefinition.Size = new System.Drawing.Size(154, 23);
            this.btnChooseDefinition.TabIndex = 3;
            this.btnChooseDefinition.Text = "Choose Definition...";
            this.btnChooseDefinition.UseVisualStyleBackColor = true;
            this.btnChooseDefinition.Click += new System.EventHandler(this.btnChooseDefinition_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(172, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(172, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 5;
            // 
            // btnGenerateOutput
            // 
            this.btnGenerateOutput.Location = new System.Drawing.Point(12, 70);
            this.btnGenerateOutput.Name = "btnGenerateOutput";
            this.btnGenerateOutput.Size = new System.Drawing.Size(154, 23);
            this.btnGenerateOutput.TabIndex = 6;
            this.btnGenerateOutput.Text = "Generate Output";
            this.btnGenerateOutput.UseVisualStyleBackColor = true;
            this.btnGenerateOutput.Click += new System.EventHandler(this.btnGenerateOutput_Click);
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(12, 99);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.Size = new System.Drawing.Size(956, 366);
            this.txtLog.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(980, 477);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnGenerateOutput);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnChooseDefinition);
            this.Controls.Add(this.btnChooseTemplate);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "RSGenerate";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnChooseTemplate;
        private System.Windows.Forms.Button btnChooseDefinition;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnGenerateOutput;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Label label1;
    }
}

