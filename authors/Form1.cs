using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;


namespace authors
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Excel.Application excelApplication;
        List<string> listOfWords;
        List<string> tokens;
        Dictionary<string, Hashtable> unigram;
        Dictionary<string, Hashtable> bigram;
        Dictionary<string, Hashtable> trigram;
        Form2 form2;
        System.Windows.Forms.DataGridView dataGridView;

        public Form1()
        {
            excelApplication = new Microsoft.Office.Interop.Excel.Application();
            listOfWords = new List<string>();
            tokens = new List<string>();
            unigram = new Dictionary<string, Hashtable>();
            bigram = new Dictionary<string, Hashtable>();
            trigram = new Dictionary<string, Hashtable>();
            form2 = new Form2();
            dataGridView = null;
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if ((openFileDialog1.ShowDialog() == DialogResult.OK) && (!string.IsNullOrEmpty(openFileDialog1.FileName)))
            {
                Workbook workBook = excelApplication.Workbooks.Open(openFileDialog1.FileName);
                ProcessEachWorksheet(workBook);
            }
        }

        private void ProcessEachWorksheet(Workbook workBook)
        {
            int numberOfSheets = workBook.Sheets.Count;
            btnExport.Enabled = false;

            // loop through all worksheets of the browsed workbook
            for (int sheetNumber = 1; sheetNumber < numberOfSheets + 1; sheetNumber++)
            {
                Worksheet workSheet = (Worksheet)workBook.Sheets[sheetNumber];
                Range range = workSheet.UsedRange;
                int rows_count = range.Rows.Count;

                // when OK button of form2 dialog box is clicked
                if (form2.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processing... Please wait.";
                    // if any of the N-gram Frequency radio button in form2 is checked
                    if (Form2.rbtnUnigramFrequencyStatic.Checked || Form2.rbtnBigramFrequencyStatic.Checked || Form2.rbtnTrigramFrequencyStatic.Checked)
                    {

                        string Authors;

                        // loop through all rows of worksheet. Start from row 2, first row is for headings
                        for (int row = 2; row <= rows_count; row++)
                        {
                            // clear the list in order to remove the words of title of last iteration if any
                            listOfWords.Clear();

                            // get the title of paper from current row
                            string paper = String.Format("{0}", workSheet.Cells[row, 2].value);
                            
                            // get the authors list of paper from current row
                            Authors = String.Format("{0}", workSheet.Cells[row, 3].value);

                            // split the title into words
                            char[] separaters = { ' ' };
                            string[] words = paper.Split(separaters);

                            // split the authors into authorsArray, one author only per array element
                            char[] authorSeparaters = { ',' };
                            string[] authorsArray = Authors.Split(authorSeparaters);

                            // define the stop words to skip
                            string[] stopWords = { "with", "A", "An", "The", "on", "of", "-", "(", ")", "i", "is", "a", "to", "able", "that", "the", "your", "you", "can", "be", "for", "about", "across", "after", "all", "almost", "also", "am", "among", "an", "and", "any", "are", "as", "at", "be", "because", "in", "or" };

                            // loop through each word of current title and add it to list, ignoring the stop words
                            foreach (string word in words)
                            {
                                // if the current word is not a stop word then add it to list
                                if (Array.IndexOf(stopWords, word) == -1)
                                {
                                    listOfWords.Add(word.ToLower());
                                }
                            }

                            // create tokens
                            for (int i = 0; i < listOfWords.Count; i++)
                            {
                                for (int j = i; j < listOfWords.Count; j++)
                                {
                                    string token = string.Empty;
                                    for (int k = i; k <= j; k++)
                                    {
                                        token = token + " " + listOfWords[k];
                                        token = token.Trim();
                                    }
                                    tokens.Add(token);

                                    string[] wordsInToken = token.Split(separaters);

                                    // if the words in token created above is a single word, add it to unigram count
                                    if (wordsInToken.Count() == 1)
                                    {
                                        if (unigram.ContainsKey(token))
                                        {
                                            // update the unigrams count
                                            int count = (int)unigram[token]["count"];
                                            count++;
                                            unigram[token]["count"] = count;

                                            // update the authors list, first cast it to list
                                            List<string> listOfAuthors = (List<string>)unigram[token]["authors"];
                                            
                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            unigram[token]["authors"] = listOfAuthors;
                                        }
                                        else
                                        {
                                            // create hash table
                                            Hashtable hashTable = new Hashtable();
                                            
                                            // add the authors count
                                            hashTable.Add("count", 1);

                                            // create list of type string
                                            List<string> listOfAuthors = new List<string>();

                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            // put array of strings in hash table
                                            hashTable.Add("authors", listOfAuthors);

                                            // put hastable in dictionary
                                            unigram[token] = hashTable;
                                        }
                                    }

                                    // if the words in token created above are two words, add it to bigram count
                                    if (wordsInToken.Count() == 2)
                                    {
                                        if (bigram.ContainsKey(token))
                                        {
                                            // update the unigrams count
                                            int count = (int)bigram[token]["count"];
                                            count++;
                                            bigram[token]["count"] = count;

                                            // update the authors list, first cast it to list
                                            List<string> listOfAuthors = (List<string>)bigram[token]["authors"];

                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            bigram[token]["authors"] = listOfAuthors;
                                        }
                                        else
                                        {
                                            // create hash table
                                            Hashtable hashTable = new Hashtable();

                                            // add the authors count
                                            hashTable.Add("count", 1);

                                            // create list of type string
                                            List<string> listOfAuthors = new List<string>();

                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            // put array of strings in hash table
                                            hashTable.Add("authors", listOfAuthors);

                                            // put hastable in dictionary
                                            bigram[token] = hashTable;
                                        }
                                    }

                                    // if the words in token created above are three words, add it to trigram count
                                    if (wordsInToken.Count() == 3)
                                    {
                                        if (trigram.ContainsKey(token))
                                        {
                                            // update the unigrams count
                                            int count = (int)trigram[token]["count"];
                                            count++;
                                            trigram[token]["count"] = count;

                                            // update the authors list, first cast it to list
                                            List<string> listOfAuthors = (List<string>)trigram[token]["authors"];

                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            trigram[token]["authors"] = listOfAuthors;
                                        }
                                        else
                                        {
                                            // create hash table
                                            Hashtable hashTable = new Hashtable();

                                            // add the authors count
                                            hashTable.Add("count", 1);

                                            // create list of type string
                                            List<string> listOfAuthors = new List<string>();

                                            // loop through each author of current authors list
                                            foreach (string Author in authorsArray)
                                            {
                                                // if the list dosn't not contain the author, add it
                                                if (!(listOfAuthors.Contains(Author)))
                                                {
                                                    listOfAuthors.Add(Author);
                                                }
                                            }

                                            // put array of strings in hash table
                                            hashTable.Add("authors", listOfAuthors);

                                            // put hastable in dictionary
                                            trigram[token] = hashTable;
                                        }
                                    }
                                }
                            }
                        }

                        
                        // if Unigram Frequency checkbox was selected
                        if (Form2.rbtnUnigramFrequencyStatic.Checked)
                        {
                            CreateDataGridView();
                            CreateColumns();

                            // sort the dictionary by value
                            int serialNumber = 1;
                            foreach (KeyValuePair<string, Hashtable> uniGram in unigram)
                            {
                                //string output = string.Format("Key: {0}, Value: {1}", unigram.Key, unigram.Value);

                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);

                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = uniGram.Key;
                                
                                Hashtable hashTable = uniGram.Value;
                                int count = (int)hashTable["count"];
                                List<string> listOfAuthors = (List<string>)hashTable["authors"];

                                // convert the listOfAuthors to comma delimited string, because, cell of data grid view
                                // cannot store List types...
                                string commaDelimitedListOfAuthors = listOfAuthors.Aggregate((x, y) => x + "," + y);
                                
                                dgvRow.Cells[2].Value = count;
                                dgvRow.Cells[3].Value = commaDelimitedListOfAuthors;

                                /*
                                // loop through all columns of current row
                                for (int column = 0; column < dataGridView1.Columns.Count; column++)
                                {
                                    dgvRow.Cells[column].Value = words[0];
                                }
                                */

                                dataGridView.Rows.Add(dgvRow);
                            }
                        }

                        
                        // if Bigram Frequency checkbox was selected
                        if (Form2.rbtnBigramFrequencyStatic.Checked)
                        {
                            CreateDataGridView();
                            CreateColumns();

                            // sort the dictionary by value
                            int serialNumber = 1;
                            // in case of dictionary, use the below line to loop through it in sorted order
                            // foreach (KeyValuePair<string, int> bigram in bigram.OrderByDescending(key => key.Value))
                            foreach (KeyValuePair<string, Hashtable> biGram in bigram)
                            {
                                //string output = string.Format("Key: {0}, Value: {1}", unigram.Key, unigram.Value);

                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);

                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = biGram.Key;

                                Hashtable hashTable = biGram.Value;
                                int count = (int)hashTable["count"];
                                List<string> listOfAuthors = (List<string>)hashTable["authors"];

                                // convert the listOfAuthors to comma delimited string, because, cell of data grid view
                                // cannot store List types...
                                string commaDelimitedListOfAuthors = listOfAuthors.Aggregate((x, y) => x + "," + y);

                                dgvRow.Cells[2].Value = count;
                                dgvRow.Cells[3].Value = commaDelimitedListOfAuthors;


                                /*
                                // loop through all columns of current row
                                for (int column = 0; column < dataGridView1.Columns.Count; column++)
                                {
                                    dgvRow.Cells[column].Value = words[0];
                                }
                                */

                                dataGridView.Rows.Add(dgvRow);
                            }
                            
                        }
                        
                        // if Trigram Frequency checkbox was selected
                        if (Form2.rbtnTrigramFrequencyStatic.Checked)
                        {
                            CreateDataGridView();
                            CreateColumns();

                            // sort the dictionary by value
                            int serialNumber = 1;
                            // in case of dictionary, use the below line to sort dictionary by value
                            // foreach (KeyValuePair<string, int> trigram in trigram.OrderByDescending(key => key.Value))
                            foreach (KeyValuePair<string, Hashtable> triGram in trigram)
                            {
                                //string output = string.Format("Key: {0}, Value: {1}", unigram.Key, unigram.Value);

                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);

                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = triGram.Key;

                                Hashtable hashTable = triGram.Value;
                                int count = (int)hashTable["count"];
                                List<string> listOfAuthors = (List<string>)hashTable["authors"];

                                // convert the listOfAuthors to comma delimited string, because, cell of data grid view
                                // cannot store List types...
                                string commaDelimitedListOfAuthors = listOfAuthors.Aggregate((x, y) => x + "," + y);

                                dgvRow.Cells[2].Value = count;
                                dgvRow.Cells[3].Value = commaDelimitedListOfAuthors;

                                dataGridView.Rows.Add(dgvRow);
                            }
                        }
                        
                    }


                    // if the Exclusive Contribution of authors radio button in form2 was checked
                    if (Form2.rbtnExclusiveContributionOfAuthorsStatic.Checked)
                    {
                        CreateDataGridView();
                        CreateColumns();

                        // store accumulative sum of each authors points in dictionary
                        // key = name of author, value = another dictionary where keys are accumulativeSum and iterationsCount
                        // authorr["A"]["accumulativeSum"] = 5.67;
                        // authorr["A"]["iterationsCount"] = 7;
                        Dictionary<string, Dictionary<string, double>> authorr = new Dictionary<string, Dictionary<string, double>>();

                        // loop through all rows of worksheet. Start from row 2, first row is for headings
                        int serialNumber = 1;
                        for (int row = 2; row <= rows_count; row++)
                        {
                            // update the success counter
                            UpdateSuccessCounter(serialNumber, rows_count);

                            // get the title and authors of paper from current row
                            string title = null;
                            string authors = null;
                            try
                            {
                                title = String.Format("{0}", workSheet.Cells[row, 2].value);
                                authors = String.Format("{0}", workSheet.Cells[row, 3].value);
                            }
                            catch (ArgumentNullException exc)
                            {
                                title = String.Format("{0}", "Paper title not provided in Excel file");
                                authors = String.Format("{0}", workSheet.Cells[row, 3].value);
                            }

                            // split the authors into one other per array element
                            char[] separaters = { ',' };
                            string[] authorsArray = authors.Split(separaters);
                            double exclusiveContribution = new double();
                            if (authorsArray.Count() > 1)
                            {
                                exclusiveContribution = 1.00 / (authorsArray.Count() - 1);
                            }
                            else
                            {
                                exclusiveContribution = 1;
                            }
                            // create a new row, put data inside it and add the row to data grid view.
                            
                            // iterate through the list of authors for the current paper (row) to add score for each author,
                            // against each paper

                            // the following loop will print the each authors score on data grid view and also will
                            // calculate accumulative sum and iteration counts and will store in dictionary
                            foreach(string author in authorsArray)
                            {
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);

                                dgvRow.Cells[0].Value = serialNumber;
                                dgvRow.Cells[1].Value = title;
                                dgvRow.Cells[2].Value = author;
                                dgvRow.Cells[3].Value = exclusiveContribution;

                                dataGridView.Rows.Add(dgvRow);

                                if (authorr.ContainsKey(author))
                                {
                                    authorr[author]["accumulativeSum"] += exclusiveContribution;
                                    authorr[author]["iterationsCount"]++;
                                }
                                else
                                {
                                    Dictionary<string, double> dictionary = new Dictionary<string, double>();
                                    dictionary["accumulativeSum"] = exclusiveContribution;
                                    dictionary["iterationsCount"] = 1;
                                    authorr[author] = dictionary;
                                }
                                
                            }

                            serialNumber++;
                        }


                        if (MessageBox.Show("Do you want the accumulative sum of points for each author?", "Question", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            CreateDataGridView();

                            // create columns
                            {
                                DataGridViewTextBoxColumn colSerialNo = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colAuthor = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colAccumulativeSum = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colIterationsCount = new DataGridViewTextBoxColumn();

                                // set the properties for columns
                                // colcolAuthors
                                // 
                                colSerialNo.HeaderText = "Serial No";
                                colSerialNo.Name = "colSerialNo";
                                colSerialNo.ReadOnly = true;
                                colSerialNo.Width = 300;
                                //
                                // colcolAuthors
                                // 
                                colAuthor.HeaderText = "Author";
                                colAuthor.Name = "colAuthor";
                                colAuthor.ReadOnly = true;
                                colAuthor.Width = 300;
                                //
                                // colAccumulativeSum
                                // 
                                colAccumulativeSum.HeaderText = "Accumulative Sum";
                                colAccumulativeSum.Name = "colAccumulativeSum";
                                colAccumulativeSum.ReadOnly = true;
                                colAccumulativeSum.Width = 300;
                                //
                                // colIterationsCount
                                // 
                                colIterationsCount.HeaderText = "Iterations Count";
                                colIterationsCount.Name = "colIterationsCount";
                                colIterationsCount.ReadOnly = true;
                                colIterationsCount.Width = 300;
                                //

                                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                                {
                                    colSerialNo,
                                    colAuthor,
                                    colAccumulativeSum,
                                    colIterationsCount
                                });
                            }

                            serialNumber = 1; // reset the serial number so that it starts from 1
                            foreach (KeyValuePair<string, Dictionary<string, double>> keyVal in authorr)
                            {
                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);
                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = keyVal.Key;
                                dgvRow.Cells[2].Value = keyVal.Value["accumulativeSum"];
                                dgvRow.Cells[3].Value = keyVal.Value["iterationsCount"];
                                dataGridView.Rows.Add(dgvRow);
                            }
                        }    
                    }


                    // if the Exclusive Contribution of papers radio button in form2 was checked
                    if (Form2.rbtnExclusiveContributionOfPapersStatic.Checked)
                    {
                        CreateDataGridView();
                        CreateColumns();

                        // loop through all rows of worksheet. Start from row 2, first row is for headings
                        int serialNumber = 1;
                        for (int row = 2; row <= rows_count; row++)
                        {
                            // update success counter by providing the completed rows and total rows
                            UpdateSuccessCounter(serialNumber, rows_count);

                            // get the title and authors of paper from current row
                            string title = null;
                            string authors = null;
                            try
                            {
                                title = String.Format("{0}", workSheet.Cells[row, 2].value);
                                authors = String.Format("{0}", workSheet.Cells[row, 3].value);
                            }
                            catch (ArgumentNullException exc)
                            {
                                title = String.Format("{0}", "Paper title not provided in Excel file");
                                authors = String.Format("{0}", workSheet.Cells[row, 3].value);
                            }

                            // split the authors into one other per array element
                            char[] separaters = { ',' };
                            string[] authorsArray = authors.Split(separaters);
                            double exclusiveContribution = new double();
                            if (authorsArray.Count() > 1)
                            {
                                exclusiveContribution = 1.00 / (authorsArray.Count() - 1);
                            }
                            else
                            {
                                exclusiveContribution = 1;
                            }
                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);
                                
                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = title;
                                dgvRow.Cells[2].Value = authors;   
                                dgvRow.Cells[3].Value = exclusiveContribution;

                                dataGridView.Rows.Add(dgvRow);
                                
                        }
                    }

                    // if the co-author frequency of authors radio button in form2 was checked
                    if (Form2.rbtnCoAuthorFrequencyOfAuthorsStatic.Checked)
                    {
                        CreateDataGridView();
                        CreateColumns();

                        // dictionary to store cu-author frequency
                        // author[A].Add(B)
                        // author[A].Add(C)
                        // author[A].Add(D)
                        // means author A has co-worked with authors B, C and D
                        Dictionary<string, List<string>> author = new Dictionary<string, List<string>>();

                        for (int row = 2; row <= rows_count; row++)
                        {
                            // get the authors list from current row that is separated by comma
                            string authors = String.Format("{0}", workSheet.Cells[row, 3].value);

                            // split the authors into one author per array element
                            char[] separaters = { ',' };
                            string[] authorsArray = authors.Split(separaters);

                            for (int i = 0; i < authorsArray.Length; i++)
                            {
                                for (int j = 0; j < authorsArray.Length; j++)
                                {
                                    if (i != j)
                                    {
                                        // if the dictionary does not has the current author, add it and also add the current co-author
                                        if (!(author.ContainsKey(authorsArray[i])))
                                        {
                                            List<string> list = new List<string>();
                                            list.Add(authorsArray[j]);
                                            author.Add(authorsArray[i], list);
                                        }
                                        else
                                        {
                                            // if the dictionary has the author, but does not has the current co-author in list
                                            if (!(author[authorsArray[i]].Contains(authorsArray[j])))
                                            {
                                                author[authorsArray[i]].Add(authorsArray[j]);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        int serialNumber = 1;
                        foreach (KeyValuePair<string, List<string>> keyVal in author)
                        {
                            // create a new row, put data inside it and add the row to data grid view.
                            DataGridViewRow dgvRow = new DataGridViewRow();
                            dgvRow.CreateCells(dataGridView);

                            dgvRow.Cells[0].Value = serialNumber++;
                            dgvRow.Cells[1].Value = keyVal.Key;
                            dgvRow.Cells[2].Value = keyVal.Value.Count;

                            dataGridView.Rows.Add(dgvRow);

                        }


                    }

                    // if the co-author frequency radio button in form2 was checked
                    if (Form2.rbtnCoAuthorFrequencyStatic.Checked)
                    {
                        CreateDataGridView();
                        CreateColumns();

                        /*
                         * the following authorr dictionary will contain the unique author names, papers written by each author and combined papers written by each author,
                         * in the following format.
                         * author["A"]["numberOfPapers"] = 1;
                         * author["A"]["combinedPapers"] = 2;
                         * 
                         * author["B"]["numberOfPapers"] = 6;
                         * author["B"]["combinedPapers"] = 3;
                         * 
                         * author["C"]["numberOfPapers"] = 2;
                         * author["C"]["combinedPapers"] = 8;
                         */
                        Dictionary<string, Dictionary<string, int>> authorr = new Dictionary<string, Dictionary<string, int>>();
                        
                        for (int row = 2; row <= rows_count; row++)
                        {
                            // get the authors list from current row that is separated by comma
                            string authors = String.Format("{0}", workSheet.Cells[row, 3].value);

                            // split the authors into one author per array element
                            char[] separaters = { ',' };
                            string[] authorsArray = authors.Split(separaters);

                            // now loop through each author in authorsArray and put the author name in uniqueAuthors only if it doesn't contain the author
                            foreach (string author in authorsArray)
                            {
                                string trimedAuthor = author.Trim();

                                Dictionary<string, int> paper = new Dictionary<string, int>();
                                paper = new Dictionary<string, int>();
                                
                                // if the author already exists in the dictionary, just increment the number of papers and number of combined written papers
                                if (authorr.ContainsKey(trimedAuthor))
                                {
                                    authorr[trimedAuthor]["numberOfPapers"]++;
                                    // if current authors list have authors more than one, increment the combined written papers
                                    if (authorsArray.Count() > 1)
                                    {
                                        authorr[trimedAuthor]["combinedPapers"]++;
                                    }
                                }
                                // if the author doesn't exist in the dictionary, add it with the number of papers and number of combined written papers
                                else
                                {
                                    if (!(trimedAuthor.Equals("N/A")))
                                    {
                                        paper["numberOfPapers"] = 1;
                                        // if current authors list have authors more than one, assign 1 to combined written papers
                                        if (authorsArray.Count() > 1)
                                        {
                                            paper["combinedPapers"] = 1;
                                        }
                                        // if current authors list have only one author, assign 0 to combined written papers
                                        else
                                        {
                                            paper["combinedPapers"] = 0;
                                        }

                                        // assign the number of papers and combined written papers to current author
                                        authorr[trimedAuthor] = paper;
                                    }
                                }
                            }
                        }

                        int serialNumber = 1;
                        // loop through the dictionary elements and add the data to dataGridView
                        foreach (KeyValuePair<string, Dictionary<string, int>> keyVal in authorr)
                        {
                            // create a new row, put data inside it and add the row to data grid view.
                            DataGridViewRow dgvRow = new DataGridViewRow();
                            dgvRow.CreateCells(dataGridView);

                            dgvRow.Cells[0].Value = serialNumber++;
                            dgvRow.Cells[1].Value = keyVal.Key;
                            dgvRow.Cells[2].Value = keyVal.Value["numberOfPapers"];
                            dgvRow.Cells[3].Value = keyVal.Value["combinedPapers"];

                            dataGridView.Rows.Add(dgvRow);
                            
                        }
                    }

                    // if the 2-author frequency radio button in form2 was checked
                    if (Form2.rbtn2AuthorContributionStatic.Checked)
                    {
                        CreateDataGridView();
                        CreateColumns();

                        // store unique authors' names and papers written by each author in dictionary
                        Dictionary<string, int> uniqueAuthors = new Dictionary<string, int>();

                        // store two authors mutual contribution in dictionary
                        Dictionary<string, int> twoAuthorsContribution = new Dictionary<string, int>();

                        for (int row = 2; row <= rows_count; row++)
                        {
                            // get the authors list from current row that is separated by comma
                            string authors = String.Format("{0}", workSheet.Cells[row, 3].value);

                            // split the authors into one author per array element
                            char[] separaters = { ',' };
                            string[] authorsArray = authors.Split(separaters);

                            string[] authorsOfExclusivitySum =
                            {
                                "wei wang",
                                "philip s. yu",
                                "tuomas sandholm",
                                "jiawei han",
                                "mahmut t. kandemir",
                                "boi faltings",
                                "kaushik roy",
                                "katsumi tanaka",
                                "ling liu",
                                "reda alhajj",
                                "wei-ying ma",
                                "wei li",
                                "jie wu",
                                "yingxu wang",
                                "barry smyth",
                                "tao li",
                                "thomas a. henzinger",
                                "w. bruce croft",
                                "viktor k. prasanna",
                                "andrew b. kahng"
                            };

                            // now loop through each author in authorsArray and put the author name in uniqueAuthors only if it doesn't contain the author
                            foreach (string author in authorsArray)
                            {
                                string trimedAuthor = author.Trim();
                                
                                // we need results for those authors only, which we have already found the exclusivity sum
                                //if (authorsOfExclusivitySum.Contains<string>(trimedAuthor))
                                {
                                    if (!(trimedAuthor.ToLower().Equals("n/a")))
                                    {
                                        // add each author and number of papers written by each author to dictionary
                                        if (uniqueAuthors.ContainsKey(trimedAuthor))
                                        {
                                            uniqueAuthors[trimedAuthor]++;
                                        }
                                        else
                                        {
                                            uniqueAuthors[trimedAuthor] = 1;
                                        }
                                    }
                                }
                            }
                            
                            for (int i = 0; i < authorsArray.Length; i++)
                            {
                                authorsArray[i] = authorsArray[i].Trim();
                                for (int j = i + 1; j < authorsArray.Length; j++)
                                {
                                    authorsArray[j] = authorsArray[j].Trim();

                                    // we need results for those authors only, which we have already found the exclusivity sum
                                    //if (authorsOfExclusivitySum.Contains<string>(authorsArray[i]) && authorsOfExclusivitySum.Contains<string>(authorsArray[j]))
                                    {               
                                        // check forward: author[A][B]
                                        if (twoAuthorsContribution.ContainsKey(authorsArray[i] + "," + authorsArray[j]))
                                        {
                                            twoAuthorsContribution[authorsArray[i] + "," + authorsArray[j]]++;
                                        }
                                        // check backward: author[B][A]
                                        else if (twoAuthorsContribution.ContainsKey(authorsArray[j] + "," + authorsArray[i]))
                                        {
                                            twoAuthorsContribution[authorsArray[j] + "," + authorsArray[i]]++;
                                        }
                                        else
                                        {
                                            twoAuthorsContribution[authorsArray[i] + "," + authorsArray[j]] = 1;
                                        }
                                    }
                                }
                            }

                        }

                        // store accumulative sum of each authors points in dictionary
                        // key = name of author, value = another dictionary where keys are accumulativeSum and iterationsCount
                        // authorr["A"]["accumulativeSum"] = 5.67;
                        // authorr["A"]["iterationsCount"] = 7;
                        Dictionary<string, Dictionary<string, double>> authorr = new Dictionary<string, Dictionary<string, double>>();

                        // the following loop will print the each authors conbined contrubtion
                        // and also will find accumulative sum of each author's points of individual contribution
                        
                        int serialNumber = 1;
                        foreach (KeyValuePair<string, int> keyVal in twoAuthorsContribution)
                        {   
                            // create a new row, put data inside it and add the row to data grid view.
                            DataGridViewRow dgvRow = new DataGridViewRow();
                            dgvRow.CreateCells(dataGridView);
                            dgvRow.Cells[0].Value = serialNumber++;

                            // pair of authors
                            dgvRow.Cells[1].Value = keyVal.Key;
                            
                            // combined papers written by pair of authors
                            dgvRow.Cells[2].Value = keyVal.Value;

                            char[] separaters = { ',' };
                            string[] twoAuthors = keyVal.Key.Split(separaters);

                            string papersWrittenByEach = null;
                            string eachAuthorContribution = null;

                            // calculate and store each authors contribution for each paper and store in this variable
                            double eachAuthorsContributionPerPaper = 0.0;

                            // iterate through each author of current two authors stored (as key) in twoAuthorsContribution dictionary
                            for (int i = 0; i < twoAuthors.Length; i++)
                            {
                                //MessageBox.Show(twoAuthors[i]);
                                // combinedPapersWrittenByTwoAuthors / TotalPapersWrittenByTheOtherCoAuthor
                                eachAuthorsContributionPerPaper = (Convert.ToDouble(keyVal.Value) / uniqueAuthors[twoAuthors[(i + 1) % twoAuthors.Length]]);
                                papersWrittenByEach += twoAuthors[i] + ": " + uniqueAuthors[twoAuthors[i]] + ", ";
                                eachAuthorContribution += twoAuthors[i] + ": " + eachAuthorsContributionPerPaper + ", ";

                                if (authorr.ContainsKey(twoAuthors[i]))
                                {
                                    authorr[twoAuthors[i]]["accumulativeSum"] += eachAuthorsContributionPerPaper;
                                    authorr[twoAuthors[i]]["iterationsCount"]++;
                                }
                                else
                                {
                                    Dictionary<string, double> dictionary = new Dictionary<string, double>();
                                    dictionary["accumulativeSum"] = eachAuthorsContributionPerPaper;
                                    dictionary["iterationsCount"] = 1;
                                    authorr[twoAuthors[i]] = dictionary;
                                }
                                
                            }

                            // remove the comma and space characters at the end of string
                            papersWrittenByEach = papersWrittenByEach.Remove(papersWrittenByEach.Length - 2);
                            eachAuthorContribution = eachAuthorContribution.Remove(eachAuthorContribution.Length - 2);

                            dgvRow.Cells[3].Value = papersWrittenByEach;
                            dgvRow.Cells[4].Value = eachAuthorContribution;

                            // add the filled row to dataGridView
                            dataGridView.Rows.Add(dgvRow);

                            // after processing the current key and values in the dictionary, remove it.
                            // because the size of this dictionary gets too large, and memory out of range exception occurs,
                            // when more variables are declared, and more variables are not able to declare due to huge size and
                            // memory overload.
                            // twoAuthorsContribution.Remove(keyVal.Key);
                        }

                        if (MessageBox.Show("Do you want to see the unique author's accumulative sum", "Question", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            CreateDataGridView();

                            // create columns
                            {
                                DataGridViewTextBoxColumn colSerialNo = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colAuthor = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colAccumulativeSum = new DataGridViewTextBoxColumn();
                                DataGridViewTextBoxColumn colIterationsCount = new DataGridViewTextBoxColumn();
                                
                                // set the properties for columns
                                // colcolAuthors
                                // 
                                colSerialNo.HeaderText = "Serial No";
                                colSerialNo.Name = "colSerialNo";
                                colSerialNo.ReadOnly = true;
                                colSerialNo.Width = 300;
                                //
                                // colcolAuthors
                                // 
                                colAuthor.HeaderText = "Author";
                                colAuthor.Name = "colAuthor";
                                colAuthor.ReadOnly = true;
                                colAuthor.Width = 300;
                                //
                                // colAccumulativeSum
                                // 
                                colAccumulativeSum.HeaderText = "Accumulative Sum";
                                colAccumulativeSum.Name = "colAccumulativeSum";
                                colAccumulativeSum.ReadOnly = true;
                                colAccumulativeSum.Width = 300;
                                //
                                // colIterationsCount
                                // 
                                colIterationsCount.HeaderText = "Iterations Count";
                                colIterationsCount.Name = "colIterationsCount";
                                colIterationsCount.ReadOnly = true;
                                colIterationsCount.Width = 300;
                                //

                                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                                {
                                    colSerialNo,
                                    colAuthor,
                                    colAccumulativeSum,
                                    colIterationsCount
                                });     
                            }

                            serialNumber = 1; // reset the serial number so that it starts from 1
                            foreach (KeyValuePair<string, Dictionary<string, double>> keyVal in authorr)
                            {
                                // create a new row, put data inside it and add the row to data grid view.
                                DataGridViewRow dgvRow = new DataGridViewRow();
                                dgvRow.CreateCells(dataGridView);
                                dgvRow.Cells[0].Value = serialNumber++;
                                dgvRow.Cells[1].Value = keyVal.Key;
                                dgvRow.Cells[2].Value = keyVal.Value["accumulativeSum"];
                                dgvRow.Cells[3].Value = keyVal.Value["iterationsCount"];
                                dataGridView.Rows.Add(dgvRow);
                            }
                        }
                        
                    }

                    label1.Text = "Done.";
                    btnExport.Enabled = true;
                }
            }
        }

        private void UpdateSuccessCounter(int serialNumber, int rows_count)
        {
            this.lblSuccessCounter.Text = serialNumber + " out of " + rows_count + " rows completed successfully.";
        }

        private void CreateDataGridView()
        {
            if(dataGridView != null)
            {
                dataGridView.Dispose();
                dataGridView = null;
            }
            
            dataGridView = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)(dataGridView)).BeginInit();

            dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            dataGridView.Location = new System.Drawing.Point(3, 40);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new System.Drawing.Size(1074, 480);
            dataGridView.TabIndex = 1;

            // Set allow user to add rows to false, because the new blank row at the end of data grid view is also included in dataGridView.Rows.Count
            dataGridView.AllowUserToAddRows = false;

            this.tableLayoutPanel1.Controls.Add(dataGridView, 0, 1);
            ((System.ComponentModel.ISupportInitialize)(dataGridView)).EndInit();

        }

        private void CreateColumns()
        {
            // create new columns for data grid view
            DataGridViewTextBoxColumn colSerialNo;
            DataGridViewTextBoxColumn colToken;
            DataGridViewTextBoxColumn colFrequency;
            DataGridViewTextBoxColumn colPaperTitle;
            DataGridViewTextBoxColumn colExclusiveContribution;
            DataGridViewTextBoxColumn colCoAuthorFrequency;
            DataGridViewTextBoxColumn colAuthors;
            DataGridViewTextBoxColumn colAuthor;
            DataGridViewTextBoxColumn colNumberOfPapers;
            DataGridViewTextBoxColumn colTwoAuthors;
            DataGridViewTextBoxColumn colCombinedPapers;
            DataGridViewTextBoxColumn colTotalPapersWrittenByEach;
            DataGridViewTextBoxColumn colEachAuthorContribution;
            DataGridViewTextBoxColumn colCoAuthorFrequencyOfAuthors;

            colSerialNo = new DataGridViewTextBoxColumn();
            colToken = new DataGridViewTextBoxColumn();
            colFrequency = new DataGridViewTextBoxColumn();
            colPaperTitle = new DataGridViewTextBoxColumn();
            colExclusiveContribution = new DataGridViewTextBoxColumn();
            colCoAuthorFrequency = new DataGridViewTextBoxColumn();
            colAuthors = new DataGridViewTextBoxColumn();
            colAuthor = new DataGridViewTextBoxColumn();
            colNumberOfPapers = new DataGridViewTextBoxColumn();
            colTwoAuthors = new DataGridViewTextBoxColumn();
            colCombinedPapers = new DataGridViewTextBoxColumn();
            colTotalPapersWrittenByEach = new DataGridViewTextBoxColumn();
            colEachAuthorContribution = new DataGridViewTextBoxColumn();
            colCoAuthorFrequencyOfAuthors = new DataGridViewTextBoxColumn();

            // set the properties for columns
            // colcolAuthors
            // 
            colNumberOfPapers.HeaderText = "Number of Papers";
            colNumberOfPapers.Name = "colnumberOfPapers";
            colNumberOfPapers.ReadOnly = true;
            colNumberOfPapers.Width = 300;
            //
            // colcolAuthors
            // 
            colAuthor.HeaderText = "Author";
            colAuthor.Name = "colAuthor";
            colAuthor.ReadOnly = true;
            colAuthor.Width = 300;
            //
            // colcolAuthors
            // 
            colAuthors.HeaderText = "Authors";
            colAuthors.Name = "colAuthors";
            colAuthors.ReadOnly = true;
            colAuthors.Width = 300;
            //
            // colExclusiveContribution
            // 
            colExclusiveContribution.HeaderText = "Exclusive Contribution";
            colExclusiveContribution.Name = "colExclusiveContribution";
            colExclusiveContribution.ReadOnly = true;
            colExclusiveContribution.Width = 300;
            //
            // colPaperTitle
            // 
            colPaperTitle.HeaderText = "Paper Title";
            colPaperTitle.Name = "colPaperTitle";
            colPaperTitle.ReadOnly = true;
            colPaperTitle.Width = 328;
            //
            // colCoAuthorFrequency
            // 
            colCoAuthorFrequency.HeaderText = "Co-Author Frequency";
            colCoAuthorFrequency.Name = "colCoAuthorFrequency";
            colCoAuthorFrequency.ReadOnly = true;
            colCoAuthorFrequency.Width = 300;
            // 
            // colCoAuthorFrequencyOfAuthors
            // 
            colCoAuthorFrequencyOfAuthors.HeaderText = "Co-Author Frequency of Authors";
            colCoAuthorFrequencyOfAuthors.Name = "colCoAuthorFrequencyOfAuthors";
            colCoAuthorFrequencyOfAuthors.ReadOnly = true;
            colCoAuthorFrequencyOfAuthors.Width = 300;
            //
            // colSerialNo
            // 
            colSerialNo.HeaderText = "Serial No.";
            colSerialNo.Name = "colSerialNo";
            colSerialNo.ReadOnly = true;
            // 
            // colToken
            // 
            colToken.HeaderText = "Token";
            colToken.Name = "colToken";
            colToken.ReadOnly = true;
            colToken.Width = 328;
            // 
            // colFrequency
            // 
            colFrequency.HeaderText = "Frequency";
            colFrequency.Name = "colUnigramFrequency";
            colFrequency.ReadOnly = true;
            colFrequency.Width = 300;
            // 
            // colFrequency
            // 
            colTwoAuthors.HeaderText = "Two Authors";
            colTwoAuthors.Name = "colTwoAuthors";
            colTwoAuthors.ReadOnly = true;
            colTwoAuthors.Width = 200;
            //
            // colCombinedPapers
            //
            colCombinedPapers.HeaderText = "Combined Papers";
            colCombinedPapers.Name = "colCombinedPapers";
            colCombinedPapers.ReadOnly = true;
            colCombinedPapers.Width = 150;
            //
            // colTotalPapersWrittenByEach
            //
            colTotalPapersWrittenByEach.HeaderText = "Total Papers Written By Each";
            colTotalPapersWrittenByEach.Name = "colTotalPapersWrittenByEach";
            colTotalPapersWrittenByEach.ReadOnly = true;
            colTotalPapersWrittenByEach.Width = 250;
            //
            // colEachAuthorContribution
            //
            colEachAuthorContribution.HeaderText = "Each Author's Contribution";
            colEachAuthorContribution.Name = "colEachAuthorContribution";
            colEachAuthorContribution.ReadOnly = true;
            colEachAuthorContribution.Width = 250;
            

            // if the n-gram is selected, add these columns to data grid view
            if (Form2.rbtnUnigramFrequencyStatic.Checked || Form2.rbtnBigramFrequencyStatic.Checked || Form2.rbtnTrigramFrequencyStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colToken,
                    colFrequency,
                    colAuthors
                });
            }

            // if the exclusive contribution of papers is selected, add these columns to data grid view
            if (Form2.rbtnExclusiveContributionOfPapersStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colPaperTitle,
                    colAuthors,
                    colExclusiveContribution
                });
            }

            // if the exclusive contribution of authors is selected, add these columns to data grid view
            if (Form2.rbtnExclusiveContributionOfAuthorsStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colPaperTitle,
                    colAuthor,
                    colExclusiveContribution
                });
            }

            // if the co-author frequency is selected, add these columns to data grid view
            if (Form2.rbtnCoAuthorFrequencyStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colAuthor,
                    colNumberOfPapers,
                    colCoAuthorFrequency
                });
            }

            // if the co-author frequency of authors is selected, add these columns to data grid view
            if (Form2.rbtnCoAuthorFrequencyOfAuthorsStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colAuthor,
                    colCoAuthorFrequencyOfAuthors
                });
            }

            // if the co-author frequency is selected, add these columns to data grid view
            if (Form2.rbtn2AuthorContributionStatic.Checked)
            {
                dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
                {
                    colSerialNo,
                    colTwoAuthors,
                    colCombinedPapers,
                    colTotalPapersWrittenByEach,
                    colEachAuthorContribution
                });
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Do you want to exit?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = string.Empty;
            lblSuccessCounter.Text = string.Empty;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                Microsoft.Office.Interop.Excel.Application myExcelFile = new Microsoft.Office.Interop.Excel.Application();

                sfd.Title = "Save As";
                sfd.Filter = "Microsoft Excel|*.xlsx";
                sfd.DefaultExt = "xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {   
                    Workbook myWorkBook = myExcelFile.Workbooks.Add(XlSheetType.xlWorksheet);
                    Worksheet myWorkSheet = (Worksheet)myExcelFile.ActiveSheet;
                    
                    // don't open excel file in windows during building
                    myExcelFile.Visible = false;

                    label1.Text = "Building the excel file. Please wait...";

                    // set the first row cells as column names according to the names of data grid view columns names
                    foreach (DataGridViewColumn dgvColumn in dataGridView.Columns)
                    {
                        // dataGridView columns is a zero-based array, while excel sheet is a 1-based array
                        // so, first row of excel sheet has index 1
                        // set columns of first row as titles of columns
                        myWorkSheet.Cells[1, dataGridView.Columns.IndexOf(dgvColumn) + 1] = dgvColumn.HeaderText;
                    }

                    // since, first row has titles that are set above, start from row-2 and fill each row of excel file.
                    for (int i = 2; i <= dataGridView.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            myWorkSheet.Cells[i, j + 1] = dataGridView.Rows[i - 2].Cells[j].Value;
                        }
                    }

                    // set the font style of first row as Bold which has titles of each column
                    myWorkSheet.Rows[1].Font.Bold = true;
                    myWorkSheet.Rows[1].Font.Size = 12;

                    // after filling, save the file to the specified location
                    string savePath = sfd.FileName;
                    myWorkBook.SaveCopyAs(savePath);
                }

                label1.Text = "File saved successfully.";

                if (MessageBox.Show("File saved successfully. Do you want to open it now?", "File Saved!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    myExcelFile.Visible = true;
                }
            }
            catch(Exception exc)
            {
                label1.Text = exc.Message;
            }
        }
    }
}
