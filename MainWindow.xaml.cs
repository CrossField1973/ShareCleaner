using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows;
using wf = System.Windows.Forms;

namespace MMIT.ShareCleaner
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Remember archivepath
        private string ArchivePath;
        private String SelectedPath = "C:\\";
        private Node root;

        int t = 2;
        List<string[]> rows;
        private string importpath;
        List<string[]> gesamt = new List<string[]>();



        public MainWindow()
        {

            InitializeComponent();
        }

        //initiating the Treeview
        private void ListDirectory(/*System.Windows.Controls.TreeView treeView, */string path)
        {
            try
            {
                //treeView.Items.Clear();
                path = SelectedPath;
                var rootDirectoryInfo = new DirectoryInfo(path);
                root = Node.GetTree(rootDirectoryInfo);
                //treeView.Items.Add(root);
                root.FullPathDi = SelectedPath;

            }
            catch (System.ArgumentException)
            {

            }
        }




        //Method to select the root folder and initiate the treeview
        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Textmessage("Rufe Ordnerinhalte ab...");
                var fbd = new wf.FolderBrowserDialog();
                fbd.ShowDialog();
                OK.Focusable = false;
                Archiv.Focusable = false;
                Excel_export.Focusable = false;
                Excel_import.Focusable = false;
                Button1.Focusable = false;

                BackgroundWorker worker = new BackgroundWorker();
                worker.DoWork += (o, ea) =>
                                {
                                    SelectedPath = fbd.SelectedPath;
                                    ListDirectory(SelectedPath);
                                };
                worker.RunWorkerCompleted += (o, ea) =>
                            {
                                treeView.Items.Add(root);
                                ProgressIndicator.IsBusy = false;
                                OK.Focusable = true;
                                Archiv.Focusable = true;
                                Excel_export.Focusable = true;
                                Excel_import.Focusable = true;
                                Button1.Focusable = true;

                            };

                treeView.Items.Clear();
                ProgressIndicator.IsBusy = true;
                worker.RunWorkerAsync();
                OK.IsEnabled = true;

            }
            catch (System.ArgumentException)
            {
                throw new Exception("mipmöp");
            }







        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Okay(root);
            Textmessage("Fertig!");
            treeView.Items.Clear();
            ListDirectory(/*treeView, */SelectedPath);
            treeView.Items.Add(root);
        }

        private void Okay(Node item)
        {
            try
            {
                if (item.IsArchiveChecked)
                {
                    if (ArchivePath != null)
                    {
                        //Ckeck if Folder 
                        if (Directory.Exists(item.FullPathDi))
                        {
                            Textmessage("Move Directory: " + item.PathDi);
                            Directory.Move(item.FullPathDi, ArchivePath + "\\" + item.PathDi);
                            
                        }
                        //Check if File
                        if (File.Exists(item.FullPathFi))
                        {
                            Textmessage("Move File: " + item.PathFi);
                            File.Move(item.FullPathFi, ArchivePath + "\\" + item.PathFi);
                        }
                    }
                    else
                    {
                        Textmessage("Es wurde kein Archiv angegeben.");
                    }
                }
                if (item.IsDeleteChecked)
                {
                    if (Directory.Exists(item.FullPathDi))
                    {
                        DeleteDirectory(item.FullPathDi);
                    }
                    if (File.Exists(item.FullPathFi))
                    {
                        Textmessage("Delete File: " + item.FullPathFi);
                        File.Delete(item.FullPathFi);
                    }
                }
                else { }
                foreach (Node n in item.Children)
                {
                    Okay(n);
                }
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                Textmessage("Der angegebene Dateipfad konnte nicht gefunden werden.");

            }
            catch (System.NullReferenceException)
            {
            }
        }


        private void DeleteDirectory(string target_dir) {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                Textmessage("Delete File: " + file);
                File.Delete(file);                
            }

            foreach (string dir in dirs)
            {
                Textmessage("Delete Directory: " + dir);
                DeleteDirectory(dir);
            }
            Textmessage("Delete target_dir: " + target_dir);
            Directory.Delete(target_dir, false);
        }

        private void DeleteRadio_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenDelete(true);


            //    n.Background = "white";
            // warning if more than one radiobutton is selectewd after import
            Methode_1(n);
        }


        private void DeleteRadio_Unchecked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenDelete(false);
            Methode_1(n);
        }

        private void ArchiveRadio_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenArchive(true);
            Methode_1(n);
        }


        private void ArchiveRadio_Unchecked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenArchive(false);
            Methode_1(n);
        }

        private void IgnoreRadio_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenIgnore(true);
            Methode_1(n);
        }

        private void IgnoreRadio_Unchecked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton radio = sender as System.Windows.Controls.RadioButton;
            Node n = radio.DataContext as Node;
            n.CheckAllChildrenIgnore(false);
            Methode_1(n);

        }

        void Methode_1(Node n)
        {
            // checking which radiobutton is checked and initializing the associated task
            if (n.IsDeleteChecked == true && n.IsArchiveChecked == true && n.IsIgnoreChecked == true)
            {
                n.IsDeleteChecked = false;
                n.IsArchiveChecked = false;
                n.IsIgnoreChecked = false;
                n.Background = "red";
            }
            if (n.IsDeleteChecked == true && n.IsArchiveChecked == true)
            {
                n.IsDeleteChecked = false;
                n.IsArchiveChecked = false;
                n.Background = "red";
            }
            if (n.IsArchiveChecked == true && n.IsIgnoreChecked == true)
            {
                n.IsArchiveChecked = false;
                n.IsIgnoreChecked = false;
                n.Background = "red";
            }
            if (n.IsDeleteChecked == true && n.IsIgnoreChecked == true)
            {
                n.IsDeleteChecked = false;
                n.IsIgnoreChecked = false;
                n.Background = "red";
            }
            if (n.IsDeleteChecked == true | n.IsIgnoreChecked == true | n.IsArchiveChecked == true)
            {
                n.Background = "white";
            }
            else
            {

            }
        }
        private void Archiv_Click(object sender, RoutedEventArgs e)
        {
            var afbd = new wf.FolderBrowserDialog();
            afbd.ShowDialog();
            //Save Archive path
            ArchivePath = afbd.SelectedPath;
            string destPath = ArchivePath + "\\";
            Archiv.Content = ArchivePath;

        }

        public void CSV_import_Click(object sender, RoutedEventArgs e)
        {
            var ifd = new wf.OpenFileDialog();
            ifd.ShowDialog();
            importpath = ifd.FileName; if (importpath == "")
            {
                Textmessage("Es wurde keine zu importierende Datei angegeben.");
            }
            else
            {
                treeView.Items.Clear();


                Microsoft.Office.Interop.Excel.ApplicationClass app = new Microsoft.Office.Interop.Excel.ApplicationClass();
                Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(importpath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

                Microsoft.Office.Interop.Excel.Range usedRange = workSheet.UsedRange;

                object[,] valueArray = (object[,])usedRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);

                List<string[]> allLines = new List<string[]>();
                string selectedPath = "";
                if (valueArray.GetLength(0) > 1)
                {
                    selectedPath = valueArray[2, 1].ToString();

                    var rootDirectoryInfo = new DirectoryInfo(@selectedPath);
                    root = Node.GetTree(rootDirectoryInfo);
                    treeView.Items.Add(root);
                    root.FullPathDi = SelectedPath;

                    for (int i = 2; i < valueArray.GetLength(0); i++)
                    {
                        string[] values = new string[6];

                        values[0] = valueArray[i, 1].ToString();
                        values[1] = valueArray[i, 2].ToString();
                        values[2] = valueArray[i, 3].ToString();
                        values[3] = valueArray[i, 4].ToString();
                        values[4] = valueArray[i, 5].ToString();
                        values[5] = valueArray[i, 6].ToString();

                        allLines.Add(values);

                    }                    
                }
                else
                {
                    Textmessage("Fehler beim importieren.");
                    return;
                }
                importcheck2(root, allLines);


                //string targetFile = importpath + "Umwandler.csv";
                //convertExcelToCSV(importpath, "Tabelle1", targetFile);
                //importCSV(root, targetFile);
                OK.IsEnabled = true;
            }
        }

        //public void importCSV(Node item, string targetFile)
        //{
        //    try
        //    {
        //        string SelectedPath = "";
        //        //splitting the csv at ";" and collecting the seperate units
        //        gesamt.Clear();
        //        //Creating a helping path, copying the file to there, opening the helppath to avoid errors

        //        string helppath = targetFile + "Help.csv";
        //        File.Delete(@helppath);
        //        File.Copy(targetFile, helppath);
        //        var lines = File.ReadAllLines(helppath);
        //        int dataRowStart = 0;
        //        for (int i = dataRowStart; i < lines.Length; i++)
        //        {

        //            string line = lines[i];

        //            string[] values = line.Split(';');
        //            gesamt.Add(values);
        //            if (i == 0)
        //            {
        //                SelectedPath = values[0];
        //            }

        //        }
        //        //building new treeview with values[0] as the rootfolder
        //        treeView.Items.Clear();
        //        string path = SelectedPath;
        //        File.Delete(@helppath);
        //        File.Delete(@targetFile);
        //        var rootDirectoryInfo = new DirectoryInfo(@path);
        //        root = Node.GetTree(rootDirectoryInfo);
        //        treeView.Items.Add(root);
        //        root.FullPathDi = SelectedPath;
        //        importcheck(root);
        //        if (gesamt.Any(a => a[0] == root.FullPathDi))
        //        {

        //            string[] values = gesamt.Where(a => a[0] == root.FullPathDi).First();
        //            root.IsDeleteChecked = Boolean.Parse(values[1]);
        //            root.IsArchiveChecked = Boolean.Parse(values[2]);
        //            root.IsIgnoreChecked = Boolean.Parse(values[3]);


        //        }

        //    }
        //    catch (System.ArgumentException)
        //    {
        //        Textmessage("Es wurde keine zu importierende Datei ausgewählt.");

        //    }
        //}

        //static void convertExcelToCSV(string sourceFile, string worksheetName, string targetFile)
        //{

        //    string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties=\"Excel 12.0;readonly=false;\"";
        //    OleDbConnection conn = null;
        //    StreamWriter wrtr = null;
        //    OleDbCommand cmd = null;
        //    OleDbDataAdapter da = null;

        //    try
        //    {
        //        conn = new OleDbConnection(strConn);
        //        conn.Open();
        //        cmd = new OleDbCommand("SELECT * FROM [" + worksheetName + "$]", conn);
        //        cmd.CommandType = CommandType.Text;
        //        wrtr = new StreamWriter(targetFile, false, Encoding.UTF8);
        //        da = new OleDbDataAdapter(cmd);
        //        DataTable dt = new DataTable();
        //        da.Fill(dt);
        //        for (int x = 0; x < dt.Rows.Count; x++)
        //        {
        //            string rowString = "";
        //            for (int y = 0; y < dt.Columns.Count; y++)
        //            {
        //                rowString += dt.Rows[x][y].ToString() + ";";
        //            }
        //            wrtr.WriteLine(rowString);
        //        }
        //    }
        //    catch (Exception exc)
        //    {
        //        Console.WriteLine(exc.ToString());
        //        Console.ReadLine();
        //    }
        //    finally
        //    {
        //        if (conn.State == ConnectionState.Open)
        //            conn.Close();
        //        conn.Dispose();
        //        cmd.Dispose();
        //        da.Dispose();
        //        wrtr.Close();
        //        wrtr.Dispose();
        //    }
        //}

        public void importcheck2(Node root, List<string[]> allLines)
        {

            // going through every node and comparing the filepath to the filepath in the csv, if similar, replace old values with new ones
            foreach (Node n in root.Children)
            {
                if (allLines.Any(a => a[0] == n.FullPathDi))
                {
                    string[] values = allLines.Where(a => a[0] == n.FullPathDi).First();
                    n.IsIgnoreChecked = Boolean.Parse(values[3]);
                    n.IsDeleteChecked = Boolean.Parse(values[1]);
                    n.IsArchiveChecked = Boolean.Parse(values[2]);
                    Methode_1(n);
                    importcheck2(n, allLines);
                }
                if (allLines.Any(a => a[0] == n.FullPathFi))
                {

                    string[] values = allLines.Where(a => a[0] == n.FullPathFi).First();
                    n.IsDeleteChecked = Boolean.Parse(values[1]);
                    n.IsArchiveChecked = Boolean.Parse(values[2]);
                    n.IsIgnoreChecked = Boolean.Parse(values[3]);
                    Methode_1(n);
                    importcheck2(n, allLines);
                }
            }
        }

        //public void importcheck(Node root)
        //{

        //    // going through every node and comparing the filepath to the filepath in the csv, if similar, replace old values with new ones
        //    foreach (Node n in root.Children)
        //    {
        //        if (gesamt.Any(a => a[0] == n.FullPathDi))
        //        {
        //            string[] values = gesamt.Where(a => a[0] == n.FullPathDi).First();
        //            n.IsIgnoreChecked = Boolean.Parse(values[3]);
        //            n.IsDeleteChecked = Boolean.Parse(values[1]);
        //            n.IsArchiveChecked = Boolean.Parse(values[2]);
        //            Methode_1(n);
        //            importcheck(n);
        //        }
        //        if (gesamt.Any(a => a[0] == n.FullPathFi))
        //        {

        //            string[] values = gesamt.Where(a => a[0] == n.FullPathFi).First();
        //            n.IsDeleteChecked = Boolean.Parse(values[1]);
        //            n.IsArchiveChecked = Boolean.Parse(values[2]);
        //            n.IsIgnoreChecked = Boolean.Parse(values[3]);
        //            Methode_1(n);
        //            importcheck(n);
        //        }
        //    }
        //}

        private void CSV_export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sfd = new wf.SaveFileDialog();
                sfd.FileName = "Report_" + DateTime.Now.ToString("yyyyMMdd");
                sfd.Filter = "Excel (*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == wf.DialogResult.OK)
                {

                    string ExportPath = sfd.FileName;

                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    Microsoft.Office.Interop.Excel.Range oRng;

                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);

                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1, 1] = "Path";
                    xlWorkSheet.Cells[1, 2] = "Löschen";
                    xlWorkSheet.Cells[1, 3] = "Archivieren";
                    xlWorkSheet.Cells[1, 4] = "Ignorieren";
                    xlWorkSheet.Cells[1, 5] = "Size";
                    xlWorkSheet.Cells[1, 6] = "Last Access";
                    xlWorkSheet.get_Range("A1", "F1").Font.Bold = true;
                    xlWorkSheet.get_Range("A1", "F1").VerticalAlignment =
                        Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                    xlWorkSheet.get_Range("E1").NumberFormat = "0.00";

                    rows = new List<string[]>();
                    GetData(root);
                    ExportData(rows, xlWorkSheet);



                    oRng = xlWorkSheet.get_Range("A1", "F1");
                    oRng.EntireColumn.AutoFit();
                    xlApp.Visible = false;
                    xlApp.UserControl = false;

                    xlWorkBook.SaveAs(ExportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    Textmessage("Export abgeschlossen.");
                }
                else
                {
                    Textmessage("Es wurde keine Zieldatei für den Export angegeben.");
                }                

            }
            catch (System.ArgumentException)
            {
                Textmessage("Es wurde keine Zieldatei für den Export angegeben.");
            }
            catch (System.IO.IOException)
            {
                Textmessage("Daten konnten nicht Exportiert werden, da die Datei geöffnet ist");
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        private void GetData(Node n)
        {
            if (n != null)
            {
                if (Directory.Exists(n.FullPathDi))
                {
                    rows.Add((new List<string>() { n.FullPathDi, n.IsDeleteChecked.ToString(), n.IsArchiveChecked.ToString(), n.IsIgnoreChecked.ToString(), n.Size.ToString(), n.LastAccess.ToString() }).ToArray());
                }
                else if (File.Exists(n.FullPathFi))
                {
                    rows.Add((new List<string>() { n.FullPathFi, n.IsDeleteChecked.ToString(), n.IsArchiveChecked.ToString(), n.IsIgnoreChecked.ToString(), n.Size.ToString(), n.LastAccess.ToString() }).ToArray());
                }

                if (n.Children.Count > 0)
                {
                    foreach (Node child in n.Children)
                    {
                        GetData(child);
                    }
                }

            }
        }

        private void ExportData(List<string[]> rows, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet)
        {
            object[,] rowrows = new object[rows.Count, 6];

            // string[,] rowrows = new string[rows.Count,6];
            int i = 0;
            foreach (string[] sa in rows)
            {
                int j = 0;
                foreach (string s in sa)
                {
                    if (j == 4)
                    {
                        rowrows[i, j] = float.Parse(s);
                    }
                    else
                    {
                        rowrows[i, j] = s;
                    }

                    j++;
                }

                i++;
            }



            int rowCount = rowrows.GetLength(0);
            int columnCount = rowrows.GetLength(1);

            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, 1];
            range = range.get_Resize(rowCount, 6);


            range.set_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault, rowrows);

            //int rowCount = 2;
            //foreach (string[] s in rows)
            //{

            //    xlWorkSheet.Cells[rowCount, 1] = s[0];
            //    xlWorkSheet.Cells[rowCount, 2] = s[1];
            //    xlWorkSheet.Cells[rowCount, 3] = s[2];
            //    xlWorkSheet.Cells[rowCount, 4] = s[3];
            //    xlWorkSheet.Cells[rowCount, 5] = s[4];
            //    xlWorkSheet.Cells[rowCount, 6] = s[5];
            //    rowCount++;
            //}


        }


        private void exportCSV(Node item, string ExportPath, Microsoft.Office.Interop.Excel.Application oXL, Microsoft.Office.Interop.Excel._Workbook oWB, Microsoft.Office.Interop.Excel._Worksheet oSheet)
        {



            if (Directory.Exists(item.FullPathDi))
            {
                oSheet.Cells[t, 1] = item.FullPathDi;
                oSheet.Cells[t, 2] = item.IsDeleteChecked.ToString();
                oSheet.Cells[t, 3] = item.IsArchiveChecked.ToString();
                oSheet.Cells[t, 4] = item.IsIgnoreChecked.ToString();
                oSheet.Cells[t, 5] = item.Size.ToString();
                oSheet.Cells[t, 6] = item.LastAccess.ToString();
                t = t + 1;
            }
            else if (File.Exists(item.FullPathFi))
            {

                oSheet.Cells[t, 1] = item.FullPathFi;
                oSheet.Cells[t, 2] = item.IsDeleteChecked.ToString();
                oSheet.Cells[t, 3] = item.IsArchiveChecked.ToString();
                oSheet.Cells[t, 4] = item.IsIgnoreChecked.ToString();
                oSheet.Cells[t, 5] = item.Size.ToString();
                oSheet.Cells[t, 6] = item.LastAccess.ToString();
                t = t + 1;
            }
            foreach (Node n in item.Children)
            {
                exportCSV(n, ExportPath, oXL, oWB, oSheet);
            }
        }
        private void Textmessage(string ausgabe)
        {
            textboxerror.AppendText(ausgabe + "\n");
            textboxerror.ScrollToEnd();
        }


    }
}












