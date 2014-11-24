using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Data;

namespace WpfApplication2

{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        System.Data.DataTable dt;
        string[] Selectedfiles;
        object miss = System.Reflection.Missing.Value;
        object readOnly = true;
        public MainWindow()
        {
            InitializeComponent();
            InitType();
        }
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = new System.Windows.Forms.DialogResult();
            result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string[] files = Directory.GetFiles(fbd.SelectedPath);
                Selectedfiles = files.Where(s => (!s.Contains("~")) && (s.Contains(".doc") || s.Contains(".docx") || s.Contains(".DOC") || s.Contains(".DOCX"))).ToArray();
                txtPath.Text = fbd.SelectedPath;
                btnRead.IsEnabled = true;
            }
        }

        private void btnRead_Click(object sender, RoutedEventArgs e)
        {
            int count = Selectedfiles.Count();
            MyThreadPool t1 = new MyThreadPool(count, Selectedfiles);
            t1.LaunchThreads();
            // read();
            btnExport.IsEnabled = true;
            }

        private void read()
        {
            foreach (string file in Selectedfiles)
            {
                Word.Application app = new Word.Application();
                Word.Document doc = app.Documents.Open(file, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss);
                string[] tempData = new string[doc.FormFields.Count];
                int i = 0;
                foreach (Word.FormField f in doc.FormFields)
                {
                    tempData[i] = f.Result;
                    i++;
                }
                dt.Rows.Add(tempData);
                dgDisp.DataContext = dt.DefaultView;
                object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
                doc.Close(saveOptionsObject, miss, miss);
                app.Quit(saveOptionsObject, miss, miss);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(miss);
            Excel.Worksheet xlSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                xlSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    xlSheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                }
            }
            xlSheet.Columns.AutoFit();
            xlSheet.Rows.AutoFit();
            Excel.Range cellRange = xlSheet.get_Range("A1:AA1", miss);
            cellRange.Interior.Color = Color.LightBlue;
            cellRange.Font.Bold = true;
            xlSheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            xlWorkBook.SaveAs(txtPath.Text + @"\Users.xlsx");
            xlWorkBook.Close(miss, miss, miss);
            xlApp.Quit();
            System.Windows.Forms.MessageBox.Show("Done, File is in " + txtPath.Text + "\\User.xlsx");
        }
    }
    public class MyThreadPool
    {
        private IList<Thread> _threads;
        private readonly int MAX_THREADS = 10;
        private int startIndex = 0;
        private int endIndex = 9;
        private int count;
        private string[] files;
        object miss = System.Reflection.Missing.Value;
        object readOnly = true;
        DataTable dt = new DataTable();


        public MyThreadPool(int Count, string[] selectedFiles)
        {
            _threads = new List<Thread>();
            count = Count;
            files = selectedFiles;
            dt = new System.Data.DataTable();
            dt.Columns.Add("Bell Sponsor Name");
            dt.Columns.Add("Bell Sponsor TXT 10 Digit ID #");
            dt.Columns.Add("Bell Sponsor Contact Phone #");
            dt.Columns.Add("Approving Manager Name");
            dt.Columns.Add("Approving Manager TXT 10 Digit ID #");
            dt.Columns.Add("Approving Manager Contact Phone #");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("First Name");
            dt.Columns.Add("Middle Name");
            dt.Columns.Add("Contact's Company Email Address");
            dt.Columns.Add("Account ID Expiration Date");
            dt.Columns.Add("Company Name");
            dt.Columns.Add("ContactPhone #");
            dt.Columns.Add("Location");
            dt.Columns.Add("Street");
            dt.Columns.Add("City");
            dt.Columns.Add("State");
            dt.Columns.Add("Zip Code");
            dt.Columns.Add("U.S. Citizenship");
            dt.Columns.Add("If Non-Citizen, provide Country of Citizenship");
            dt.Columns.Add("Is Active Directory Account Required");
            dt.Columns.Add("Is Bell Email Account Required");
            dt.Columns.Add("User Type");
            dt.Columns.Add("Bell badge #");
            dt.Columns.Add("Date Entered Into NED");
            dt.Columns.Add("ID Assigend");
            dt.Columns.Add("Textron ID Assigned");
        }

        public void LaunchThreads()
        {
            for (int i = 0; i < MAX_THREADS; i++)
            {
                if (i * 10 >= count)
                    break;
                if (count < endIndex)
                    endIndex = count % 10;
                DataTable dx = new DataTable();
                Thread thread = new Thread(() => ThreadEntry(startIndex, endIndex));
                
                
                thread.IsBackground = true;
                thread.Name = string.Format("MyThread{0}", i);

                _threads.Add(thread);
                thread.Start();
                startIndex = endIndex + 1;
                endIndex += 10;
            }
        }

        public void KillThread(int index)
        {
            string id = string.Format("MyThread{0}", index);
            foreach (Thread thread in _threads)
            {
                if (thread.Name == id)
                    thread.Abort();
            }
        }

        void ThreadEntry(int start, int end)
        {
            for (int x = start; x < end; x++)
            {
                Word.Application app = new Word.Application();
                Word.Document doc = app.Documents.Open(files[x], miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss);
                string[] tempData = new string[doc.FormFields.Count];
                int i = 0;
                foreach (Word.FormField f in doc.FormFields)
                {
                    tempData[i] = f.Result;
                    i++;
                }
                dt.Rows.Add(tempData);
                //dgDisp.DataContext = dt.DefaultView;
                object saveOptionsObject = Word.WdSaveOptions.wdDoNotSaveChanges;
                doc.Close(saveOptionsObject, miss, miss);
                app.Quit(saveOptionsObject, miss, miss);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            }
        }
    }

    #region Def
    public partial class MainWindow
    {
        public void InitType()
        {
            dt = new System.Data.DataTable();
            dt.Columns.Add("Bell Sponsor Name");
            dt.Columns.Add("Bell Sponsor TXT 10 Digit ID #");
            dt.Columns.Add("Bell Sponsor Contact Phone #");
            dt.Columns.Add("Approving Manager Name");
            dt.Columns.Add("Approving Manager TXT 10 Digit ID #");
            dt.Columns.Add("Approving Manager Contact Phone #");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("First Name");
            dt.Columns.Add("Middle Name");
            dt.Columns.Add("Contact's Company Email Address");
            dt.Columns.Add("Account ID Expiration Date");
            dt.Columns.Add("Company Name");
            dt.Columns.Add("ContactPhone #");
            dt.Columns.Add("Location");
            dt.Columns.Add("Street");
            dt.Columns.Add("City");
            dt.Columns.Add("State");
            dt.Columns.Add("Zip Code");
            dt.Columns.Add("U.S. Citizenship");
            dt.Columns.Add("If Non-Citizen, provide Country of Citizenship");
            dt.Columns.Add("Is Active Directory Account Required");
            dt.Columns.Add("Is Bell Email Account Required");
            dt.Columns.Add("User Type");
            dt.Columns.Add("Bell badge #");
            dt.Columns.Add("Date Entered Into NED");
            dt.Columns.Add("ID Assigend");
            dt.Columns.Add("Textron ID Assigned");
        }
    }
    #endregion

}