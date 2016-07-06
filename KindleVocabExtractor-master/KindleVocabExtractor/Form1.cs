using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Data.SQLite;
using System.Net;
using System.Configuration;

namespace KindleVocabExtractor
{
    public partial class frmKVE : Form
    {
        public frmKVE()
        {
            InitializeComponent();
        }

        #region Events
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void openVocabdbToolStripMenuItem_Click(object sender, EventArgs e)
        {                        
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                populateGrid();
            }

        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("There is nothing to export.", "Export Warning");
            }
            else
            {
                // The last checkbox checked needs to be commited.
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);

                Dictionary<string, string> checkedIds = GetCheckedIds();

                if (checkedIds.Count == 0)
                {
                    MessageBox.Show("Please select at least one book before exporting.", "Export Warning");
                    return;
                }

                bool isError = writeClozeFile(checkedIds);

                if (!isError)
                {
                    MessageBox.Show("Finished. Your file(s) are on the desktop.", "Success");
                }
            }
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {            
            StringBuilder version = new StringBuilder();
            version.Append("Kindle Vocab Extractor").Append(Environment.NewLine);
            version.Append("Version: ").Append(getVersionNumber()).Append(Environment.NewLine);
            version.Append("Written by: Bryan Peabody, 2015").Append(Environment.NewLine).Append(Environment.NewLine);
            version.Append("Visit http://www.bryanpeabody.net for more information.").Append(Environment.NewLine);

            MessageBox.Show(version.ToString(), "About");
        }

        private void checkForUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkForUpdates();
        }

        #endregion

        /// <summary>
        /// Show a list of books in the sqlite db.
        /// </summary>
        private void populateGrid()
        {            
            using (SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source=" + openFileDialog1.FileName + ";Version=3;"))
            {
                try
                {
                    m_dbConnection.Open();

                    string sql = "select id, title  from book_info order by authors";

                    SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {                        
                        string id = reader[0].ToString();
                        string title = reader[1].ToString();

                        var index = dataGridView1.Rows.Add();
                        dataGridView1.Rows[index].Cells["Book"].Value = title;
                        dataGridView1.Rows[index].Cells["id"].Value = id;
                    }
                                                           
                    // Resize the DataGridView columns to fit the newly loaded content.
                    dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
                    
                }
                catch (Exception e)
                {
                    MessageBox.Show("There was an error reading the SQLite database: " + e.Message);
                }
                finally
                {
                    m_dbConnection.Close();                    
                }
            }
        }

        /// <summary>
        /// Dictionary is id, book title.
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, string> GetCheckedIds()
        {
            Dictionary<string, string> checkedIds = new Dictionary<string, string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Selected"].Value != null)
                {
                    if (int.Parse(row.Cells["Selected"].Value.ToString()) == 1)
                    {
                        checkedIds.Add(row.Cells["id"].Value.ToString(), row.Cells["Book"].Value.ToString());
                    }
                }
            }

            return checkedIds;
        }

        /// <summary>
        /// For each book, get a list of the vocab and example sentences. Then write out a file with the cloze cards for each book.
        /// </summary>
        /// <param name="bookIDs"></param>
        /// <returns></returns>
        private bool writeClozeFile(Dictionary<string, string> bookIDs)
        {
            bool isError = false;

            using (SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source=" + openFileDialog1.FileName + ";Version=3;"))
            {
                try
                {
                    m_dbConnection.Open();

                    foreach (string bookID in bookIDs.Keys)
                    {
                        string sql = "select word_key, usage from lookups where book_key = '" + bookID + "'";

                        SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                        SQLiteDataReader reader = command.ExecuteReader();

                        string outfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + validateFileName(bookIDs[bookID]) + ".txt";
                        
                        StringBuilder sb = new StringBuilder();

                        while (reader.Read())
                        {
                            string word = cleanWord(reader[0].ToString());
                            string sentence = createClozeSentence(word, reader[1].ToString());

                            sb.Append(sentence).Append(Environment.NewLine);
                        }

                        File.WriteAllText(outfile, sb.ToString());
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("There was an exception writing the file: " + e.Message);
                    isError = true;
                }
                finally
                {
                    m_dbConnection.Close();
                }
            }

            return isError;
        }

        /// <summary>
        /// Make sure the filename doesn't have any special characters.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string validateFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }

        /// <summary>
        /// Incoming format is like this -> es:hola. Strip out the en: and just leave the actual word.
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        private string cleanWord(string word)
        {            
           if (!string.IsNullOrEmpty(word))
            {
                word = word.Substring(word.LastIndexOf(':') + 1);
            }

           return word;
        }

        /// <summary>
        /// Create the cloze card version of each entry.
        /// </summary>
        /// <param name="word"></param>
        /// <param name="sentence"></param>
        /// <returns></returns>
        private string createClozeSentence(string word, string sentence)
        {            
            if (!string.IsNullOrEmpty(word) && !string.IsNullOrEmpty(sentence))
            {
                sentence = sentence.Replace(word, "{{c1::" + word + "}}");
            }

            return sentence;
        }

        /// <summary>
        /// Gets the current version number.
        /// </summary>
        /// <returns></returns>
        private string getVersionNumber()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            return fileVersionInfo.ProductVersion;
        }

        /// <summary>
        /// Check for a new version.
        /// </summary>
        private void checkForUpdates()
        {
            bool hasUpdate = false;
            string url = "http://bryanpeabody.net/versions/KVE/KVE.txt";
            string clientVersion = getVersionNumber();

            using (WebClient client = new WebClient())
            {
                try
                {
                    string hostVersion = client.DownloadString(url);

                    if (hostVersion.Trim() != clientVersion.Trim())
                    {
                        hasUpdate = true;
                    }
                }
                catch (Exception e)
                { 
                    // Just ignore 
                }
            }

            if (hasUpdate)
            {
                if (MessageBox.Show("There is an update available!" + Environment.NewLine + "Would you like to download it?", "Update software", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("http://bryanpeabody.net/kindle-vocab-extractor/"); 
                }
            }  
            else
            {
                MessageBox.Show("Kindle Vocab Extractor is up to date.", "Update software");
            }
        }             
    }
}
