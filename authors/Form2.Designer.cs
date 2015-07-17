namespace authors
{
    partial class Form2
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rbtnExclusiveContributionOfAuthors = new System.Windows.Forms.RadioButton();
            this.rbtnUnigramFrequency = new System.Windows.Forms.RadioButton();
            this.rbtnTrigramFrequency = new System.Windows.Forms.RadioButton();
            this.rbtnBigramFrequency = new System.Windows.Forms.RadioButton();
            this.rbtnExclusiveContributionOfPapers = new System.Windows.Forms.RadioButton();
            this.rbtnCoAuthorFrequency = new System.Windows.Forms.RadioButton();
            this.rbtn2AuthorContribution = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.rbtnCoAuthorFrequencyOfAuthors = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbtnCoAuthorFrequencyOfAuthors);
            this.groupBox1.Controls.Add(this.rbtnExclusiveContributionOfAuthors);
            this.groupBox1.Controls.Add(this.rbtnUnigramFrequency);
            this.groupBox1.Controls.Add(this.rbtnTrigramFrequency);
            this.groupBox1.Controls.Add(this.rbtnBigramFrequency);
            this.groupBox1.Controls.Add(this.rbtnExclusiveContributionOfPapers);
            this.groupBox1.Controls.Add(this.rbtnCoAuthorFrequency);
            this.groupBox1.Controls.Add(this.rbtn2AuthorContribution);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(260, 256);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select an option";
            // 
            // rbtnExclusiveContributionOfAuthors
            // 
            this.rbtnExclusiveContributionOfAuthors.AutoSize = true;
            this.rbtnExclusiveContributionOfAuthors.Location = new System.Drawing.Point(24, 134);
            this.rbtnExclusiveContributionOfAuthors.Name = "rbtnExclusiveContributionOfAuthors";
            this.rbtnExclusiveContributionOfAuthors.Size = new System.Drawing.Size(180, 17);
            this.rbtnExclusiveContributionOfAuthors.TabIndex = 4;
            this.rbtnExclusiveContributionOfAuthors.TabStop = true;
            this.rbtnExclusiveContributionOfAuthors.Text = "Exclusive Contribution of Authors";
            this.rbtnExclusiveContributionOfAuthors.UseVisualStyleBackColor = true;
            // 
            // rbtnUnigramFrequency
            // 
            this.rbtnUnigramFrequency.AutoSize = true;
            this.rbtnUnigramFrequency.Location = new System.Drawing.Point(24, 30);
            this.rbtnUnigramFrequency.Name = "rbtnUnigramFrequency";
            this.rbtnUnigramFrequency.Size = new System.Drawing.Size(117, 17);
            this.rbtnUnigramFrequency.TabIndex = 3;
            this.rbtnUnigramFrequency.TabStop = true;
            this.rbtnUnigramFrequency.Text = "Unigram Frequency";
            this.rbtnUnigramFrequency.UseVisualStyleBackColor = true;
            // 
            // rbtnTrigramFrequency
            // 
            this.rbtnTrigramFrequency.AutoSize = true;
            this.rbtnTrigramFrequency.Location = new System.Drawing.Point(24, 82);
            this.rbtnTrigramFrequency.Name = "rbtnTrigramFrequency";
            this.rbtnTrigramFrequency.Size = new System.Drawing.Size(113, 17);
            this.rbtnTrigramFrequency.TabIndex = 1;
            this.rbtnTrigramFrequency.TabStop = true;
            this.rbtnTrigramFrequency.Text = "Trigram Frequency";
            this.rbtnTrigramFrequency.UseVisualStyleBackColor = true;
            // 
            // rbtnBigramFrequency
            // 
            this.rbtnBigramFrequency.AutoSize = true;
            this.rbtnBigramFrequency.Location = new System.Drawing.Point(24, 56);
            this.rbtnBigramFrequency.Name = "rbtnBigramFrequency";
            this.rbtnBigramFrequency.Size = new System.Drawing.Size(110, 17);
            this.rbtnBigramFrequency.TabIndex = 1;
            this.rbtnBigramFrequency.TabStop = true;
            this.rbtnBigramFrequency.Text = "Bigram Frequency";
            this.rbtnBigramFrequency.UseVisualStyleBackColor = true;
            // 
            // rbtnExclusiveContributionOfPapers
            // 
            this.rbtnExclusiveContributionOfPapers.AutoSize = true;
            this.rbtnExclusiveContributionOfPapers.Location = new System.Drawing.Point(24, 108);
            this.rbtnExclusiveContributionOfPapers.Name = "rbtnExclusiveContributionOfPapers";
            this.rbtnExclusiveContributionOfPapers.Size = new System.Drawing.Size(177, 17);
            this.rbtnExclusiveContributionOfPapers.TabIndex = 1;
            this.rbtnExclusiveContributionOfPapers.TabStop = true;
            this.rbtnExclusiveContributionOfPapers.Text = "Exclusive Contribution of Papers";
            this.rbtnExclusiveContributionOfPapers.UseVisualStyleBackColor = true;
            // 
            // rbtnCoAuthorFrequency
            // 
            this.rbtnCoAuthorFrequency.AutoSize = true;
            this.rbtnCoAuthorFrequency.Location = new System.Drawing.Point(24, 160);
            this.rbtnCoAuthorFrequency.Name = "rbtnCoAuthorFrequency";
            this.rbtnCoAuthorFrequency.Size = new System.Drawing.Size(125, 17);
            this.rbtnCoAuthorFrequency.TabIndex = 2;
            this.rbtnCoAuthorFrequency.TabStop = true;
            this.rbtnCoAuthorFrequency.Text = "Co-Author Frequency";
            this.rbtnCoAuthorFrequency.UseVisualStyleBackColor = true;
            // 
            // rbtn2AuthorContribution
            // 
            this.rbtn2AuthorContribution.AutoSize = true;
            this.rbtn2AuthorContribution.Location = new System.Drawing.Point(24, 212);
            this.rbtn2AuthorContribution.Name = "rbtn2AuthorContribution";
            this.rbtn2AuthorContribution.Size = new System.Drawing.Size(129, 17);
            this.rbtn2AuthorContribution.TabIndex = 3;
            this.rbtn2AuthorContribution.TabStop = true;
            this.rbtn2AuthorContribution.Text = "2-Authors Contribution";
            this.rbtn2AuthorContribution.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(9, 274);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(90, 274);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // rbtnCoAuthorFrequencyOfAuthors
            // 
            this.rbtnCoAuthorFrequencyOfAuthors.AutoSize = true;
            this.rbtnCoAuthorFrequencyOfAuthors.Location = new System.Drawing.Point(24, 186);
            this.rbtnCoAuthorFrequencyOfAuthors.Name = "rbtnCoAuthorFrequencyOfAuthors";
            this.rbtnCoAuthorFrequencyOfAuthors.Size = new System.Drawing.Size(176, 17);
            this.rbtnCoAuthorFrequencyOfAuthors.TabIndex = 5;
            this.rbtnCoAuthorFrequencyOfAuthors.TabStop = true;
            this.rbtnCoAuthorFrequencyOfAuthors.Text = "Co-Author Frequency of Authors";
            this.rbtnCoAuthorFrequencyOfAuthors.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 314);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rbtnExclusiveContributionOfPapers;
        private System.Windows.Forms.RadioButton rbtnCoAuthorFrequency;
        private System.Windows.Forms.RadioButton rbtn2AuthorContribution;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.RadioButton rbtnTrigramFrequency;
        private System.Windows.Forms.RadioButton rbtnBigramFrequency;
        private System.Windows.Forms.RadioButton rbtnUnigramFrequency;
        private System.Windows.Forms.RadioButton rbtnExclusiveContributionOfAuthors;
        private System.Windows.Forms.RadioButton rbtnCoAuthorFrequencyOfAuthors;
    }
}