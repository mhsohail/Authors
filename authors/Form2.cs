using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace authors
{
    public partial class Form2 : Form
    {
        public static RadioButton rbtnExclusiveContributionOfPapersStatic;
        public static RadioButton rbtnCoAuthorFrequencyStatic;
        public static RadioButton rbtnUnigramFrequencyStatic;
        public static RadioButton rbtnBigramFrequencyStatic;
        public static RadioButton rbtnTrigramFrequencyStatic;
        public static RadioButton rbtn2AuthorContributionStatic;
        public static RadioButton rbtnExclusiveContributionOfAuthorsStatic;
        public static RadioButton rbtnCoAuthorFrequencyOfAuthorsStatic;

        public Form2()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {   
            this.DialogResult = DialogResult.OK;

            if (rbtnCoAuthorFrequencyOfAuthors.Checked)
            {
                rbtnCoAuthorFrequencyOfAuthorsStatic.Checked = true;
            }

            if (rbtnExclusiveContributionOfPapers.Checked)
            {
                rbtnExclusiveContributionOfPapersStatic.Checked = true;
            }

            if (rbtnExclusiveContributionOfAuthors.Checked)
            {
                rbtnExclusiveContributionOfAuthorsStatic.Checked = true;
            }

            if (rbtnCoAuthorFrequency.Checked)
            {
                rbtnCoAuthorFrequencyStatic.Checked = true;
            }

            if (rbtnUnigramFrequency.Checked)
            {
                rbtnUnigramFrequencyStatic.Checked = true;
            }

            if (rbtnBigramFrequency.Checked)
            {
                rbtnBigramFrequencyStatic.Checked = true;
            }

            if (rbtnTrigramFrequency.Checked)
            {
                rbtnTrigramFrequencyStatic.Checked = true;
            }

            if (rbtn2AuthorContribution.Checked)
            {
                rbtn2AuthorContributionStatic.Checked = true;
            }

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            /*
             * putthig the below code in constructor will create some issues.
             * For example, if we select rbtnUnigramFrequency radio button and set rbtnUnigramFrequencyStatic to true.
             * Now, if we load this for again and select another radio button (say rbtnExclusiveContribution) and set rbtnExclusiveContributionStatic to true.
             * The problem is that the previsosly set rbtnUnigramFrequencyStatic still remains set to true and while,
             * second time rbtnExclusiveContributionStatic is checked.
             * So in total, now two radio buttons are checked, which will generate problems, because the constructor is called once only, when the form is loaded.
             * Putting the code in form's load event will set these variables to false every time the form is loaded, so there is not problem now. 
             *  :-)
             * 
             */

            rbtnExclusiveContributionOfPapersStatic = new RadioButton();
            rbtnExclusiveContributionOfPapersStatic.Checked = false;

            rbtnCoAuthorFrequencyStatic = new RadioButton();
            rbtnCoAuthorFrequencyStatic.Checked = false;

            rbtnUnigramFrequencyStatic = new RadioButton();
            rbtnUnigramFrequencyStatic.Checked = false;

            rbtnBigramFrequencyStatic = new RadioButton();
            rbtnBigramFrequencyStatic.Checked = false;

            rbtnTrigramFrequencyStatic = new RadioButton();
            rbtnTrigramFrequencyStatic.Checked = false;

            rbtn2AuthorContributionStatic = new RadioButton();
            rbtn2AuthorContributionStatic.Checked = false;

            rbtnExclusiveContributionOfAuthorsStatic = new RadioButton();
            rbtnExclusiveContributionOfAuthorsStatic.Checked = false;

            rbtnCoAuthorFrequencyOfAuthorsStatic = new RadioButton();
            rbtnCoAuthorFrequencyOfAuthorsStatic.Checked = false;
        }
    }
}
