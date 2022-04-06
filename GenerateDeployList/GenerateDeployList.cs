using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin.Controls;
using MaterialSkin;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using CommonLibs;
using System.Diagnostics;

namespace GenerateDeployList
{
    public partial class GenerateDeployList : MaterialForm
    {
        public static Excel.Workbook MyBook = null;
        public static Excel.Application MyApp = null;
        public static Excel.Worksheet MySheet = null;
        public List<string> listsP = new List<string>();
        public List<string> listsD = new List<string>();
        public string myLastPath = "";
        
        public GenerateDeployList()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);

            //listView1.Columns.Add("Delete", 80);
            listView1.Columns.Add("Item", 80);
            listView1.Columns.Add("Source Code Path", 480);
            listView1.Columns.Add("Developer", 160);
            
            //listView1.Columns[0].DisplayIndex = listView1.Columns.Count - 1;
            listView1.Invalidate();
            //listView1.Bounds = new Rectangle(new Point(10, 10), new Size(300, 200));
            //listView1.View = View.Details;
            //listView1.CheckBoxes = true;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            
        }

        //Browse
        private void materialRaisedButton1_Click(object sender, EventArgs e)
        {
            DirectoryInfo myInfo;
            FolderBrowserDialog folder = new FolderBrowserDialog();

            if(myLastPath != "")
            {
                folder.SelectedPath = myLastPath;
            }
            else
            {
                folder.SelectedPath = PathText.Text;
            }
            myLastPath = folder.SelectedPath;
            DialogResult result = folder.ShowDialog();
            if (result == DialogResult.OK)
            {
                PathText.Text = folder.SelectedPath;
                //PathText.Text = folder.ShowDialog().ToString().Trim();
                
                //string[] files = Directory.GetFiles(folder.SelectedPath);
                //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                
            }

            //---------------Old
            //var folder = new FolderBrowserDialog();
            //DialogResult result = folder.ShowDialog();
            //if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folder.SelectedPath))
            //{
            //    string[] files = Directory.GetFiles(folder.SelectedPath);
            //    //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
            //    PathText.Text = folder.SelectedPath;
            //}
        }

        //Add
        private void materialRaisedButton2_Click(object sender, EventArgs e)
        {
            listsP.Add(PathText.Text);
            listsD.Add(DevText.Text);
            //listView1.Items.Add(PathText.Text);
            //string[] row1 = { PathText.Text, DevText.Text, "E" };
            //listView1.Items.Add(lists.Count.ToString()).SubItems.AddRange(row1);

            listView1.Items.Add(new ListViewItem(new string[] {(listView1.Items.Count +1).ToString(), PathText.Text, DevText.Text }));
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            var inx = listView1.SelectedIndices;
            int dex = inx[0];
            listView1.Items.RemoveAt(listView1.SelectedIndices[0]);
            listsP.RemoveAt(dex);
            listsD.RemoveAt(dex);
        }

        private void materialRaisedGen_Click(object sender, EventArgs e)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            string path = @"\\doc\DEP_Deploy_GenO1.xls"; //Deploy format
            MyBook = MyApp.Workbooks.Open(path);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            // Project name
            MySheet.Cells[2, 3] = ProjectText.Text;
            // Deploy Date
            MySheet.Cells[3, 3] = dateTimePicker1.Text;
            // Deployer 
            MySheet.Cells[4, 3] = DeplText.Text;
            for(int i =0; i < listsP.Count; i++)
            {
                Console.WriteLine(listsP[i]);
                Console.WriteLine(listsD[i]);
                int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                Console.WriteLine(lastRow);
                if (Directory.Exists(listsP[i]))
                {
                    lastRow++;
                    // No. Col-1
                    MySheet.Cells[lastRow, 1] = i + 1;

                    // App Type Col-2
                    string[] directoryEntries = Directory.GetDirectories(listsP[i]);
                    int len = listsP[i].Length + 1;

                    foreach (string dir in directoryEntries)
                    {
                        int lenalldir = dir.Length;
                        int lendir = lenalldir - len;
                        
                        string dirName = dir.Substring(len, lendir).ToUpper();
                        if (String.Compare(dirName, "DB") == 0 || String.Compare(dirName, "DEF") == 0)
                        {
                            MySheet.Cells[lastRow, 2] = "GenTab";
                        }
                        else if (String.Compare(dirName, "EXE") == 0)
                        {
                            MySheet.Cells[lastRow, 2] = "VB";
                        }
                        else if (String.Compare(dirName, "EXCEL") == 0)
                        {
                            MySheet.Cells[lastRow, 2] = "Excel Macro";
                        }
                    }

                    // Version
                    string ver = "";
                    int flag = 0;
                    string path1 = listsP[i];
                    int pf = 0;
                    int pl = 0;
                    string Type = "";
                    for(int k =0; k < path1.Length; k++)
                    {
                        if(path1[k] == '\\')
                        {
                            if(path1[k+1] == 'V')
                            {
                                try
                                {
                                    //---- Check Type of version / New or Modified by number behind V
                                    int ckver = (int)Char.GetNumericValue(path1[k + 2]);

                                    if(ckver == 1)
                                    {
                                        Type = "New";
                                    }
                                    else
                                    {
                                        Type = "Modified";
                                    }
                                }
                                catch (FormatException)
                                {
                                    DialogResult Error = MessageBox.Show("Directory do not correct format", "Error!", MessageBoxButtons.OK);
                                }
                                //---- Find Last Character of Program Name
                                pl = k - 1;
                                for(int l = pl; l > 1; l--)
                                {
                                    if(path1[l] == '\\')
                                    {
                                        //---- Find First Character of Program Name
                                        pf = l + 1;
                                        break;
                                    }
                                }
                                //---- Flag for find Version
                                flag = 1;
                            }
                        }
                        if(flag == 1 && path1[k+1] != '_')
                        {
                            //---- Add Data version
                            ver = ver + path1[k + 1];
                        }
                        else
                        {
                            flag = 0;
                        }
                        //if(path1[k] == '\\')
                        //{
                        //    flag = 1;
                        //}
                        //if(flag == 1)
                        //{
                        //    if(path1[k] == 'V')
                        //    {
                        //        pl = k - 2;
                        //        for(int l = pl; l > 1; l--)
                        //        {
                        //            if(path1[l] == '\\')
                        //            {
                        //                pf = l + 1;
                        //                break;
                        //            }
                        //        }
                        //        //ver = ver + path1[k];
                        //        flag2 = 1;
                        //        flag = 0;
                        //    }
                        //}
                        //if(flag2 == 1)
                        //{
                        //    if(path1[k] != '_')
                        //    {
                        //        ver = ver + path1[k];
                        //    }
                        //    else
                        //    {
                        //        flag2 = 0;
                        //    }
                        //}
                        
                    }
                    // Program Name Col-3
                    int proLen = pl - pf + 1;
                    MySheet.Cells[lastRow, 3] = path1.Substring(pf, proLen);

                    ProcessDirectory(listsP[i],lastRow,listsD[i],ver,Type);
                }
                else
                {
                    Console.WriteLine("Arkkkkkkkk");
                }
            }

            //MyBook = MyApp.Workbooks.Open(path);
            //Workbook wb = excel.Workbooks.Open(path);
            //string savepath = @"Desktop\DeployData_New.xls";
            //MyBook.SaveAs(savepath);

            DialogResult result = MessageBox.Show("Do you want to save file?", "Generate Complete!", MessageBoxButtons.YesNo);
            
            if(result == System.Windows.Forms.DialogResult.Yes)
            {
                //--- Save file
                MyApp.DisplayAlerts = false;
                MyBook.SaveAs("DeployData_New", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                
                //--- Clear Process
                MyBook.Close();
                ExcelLibs.releaseObject(MySheet, MyBook);
                ExcelLibs.releaseExcelApp(MyApp, true);

                //--- Open Folder
                string MyDoc = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Process.Start(MyDoc);
            }
            else
            {
                //--- Clear Process
                MyBook.Close();
                ExcelLibs.releaseObject(MySheet, MyBook);
                ExcelLibs.releaseExcelApp(MyApp, true);
            }

            //MyApp.Quit();
        }
        
        public static void ProcessDirectory(string directory, int last, string dev, string ver, string type)
        {
            string[] directoryEntries = Directory.GetDirectories(directory);
            int len = directory.Length + 1;
            
            foreach (string dir in directoryEntries)
            {
                int lenalldir = dir.Length;
                int lendir = lenalldir - len;
                //// App Type Col-2
                string dirName = dir.Substring(len, lendir).ToUpper();
                //if (String.Compare(dirName, "DB") == 0 || String.Compare(dirName, "DEF") == 0)
                //{
                //    MySheet.Cells[last, 2] = "GenTab";
                //}
                //else if(String.Compare(dirName,"EXE") == 0)
                //{
                //    MySheet.Cells[last, 2] = "VB";
                //}
                //else if (String.Compare(dirName, "EXCEL") == 0)
                //{
                //    MySheet.Cells[last, 2] = "Excel Macro";
                //}
                //Console.WriteLine(dir);
                //string filedir = directory +"\\"+ dir;
                string[] fileEntries = Directory.GetFiles(dir);
                foreach (string fileName in fileEntries)
                {
                    if(String.Compare(Path.GetFileNameWithoutExtension(fileName), "Thumbs") != 0)
                    {
                        //Console.WriteLine("{0},{1},{2}", directory, dir, fileName);
                        ProcessFile(fileName, last, dev, ver, dirName, type);
                        last++;
                    }
                }
            }
        }

        public static void ProcessFile(string path, int last, string dev, string ver, string dirName, string type)
        {
            Console.WriteLine("Processed file '{0}'.", path);
            //--Col4 Type
            MySheet.Cells[last, 4] = type;
            //--Col5 Object
            switch (dirName)
            {
                case "DEF":
                    MySheet.Cells[last, 5] = "Screen Def";
                    break;
                case "DB":
                    MySheet.Cells[last, 5] = "DB Def";
                    break;
                case "EXE":
                    MySheet.Cells[last, 5] = "Execute File";
                    break;
                case "REPORT":
                    MySheet.Cells[last, 5] = "Excel File";
                    break;
                case "EXCEL":
                    MySheet.Cells[last, 5] = "Excel File";
                    break;
                case "TABLE":
                    MySheet.Cells[last, 5] = "Table";
                    break;
                case "PROC":
                    MySheet.Cells[last, 5] = "Procedure";
                    break;
            }
            
            //--Col6 Method
            if(Path.GetExtension(path).ToString().ToUpper() == ".XLS" || Path.GetExtension(path).ToString().ToUpper() == ".XLSX" || Path.GetExtension(path).ToString().ToUpper() == ".EXE")
            {
                MySheet.Cells[last, 6] = "DbMaker";
            }
            else if(Path.GetExtension(path).ToString().ToUpper() == ".SQL")
            {
                MySheet.Cells[last, 6] = "SQL Import";
            }
            else
            {
                MySheet.Cells[last, 6] = "Other";
            }
            //--Col7 Object name
            MySheet.Cells[last, 7] = Path.GetFileName(path);
            //--Col8 Location
            MySheet.Cells[last, 8] = Path.GetDirectoryName(path);
            //--Col9 Version
            MySheet.Cells[last, 9] = ver;
            //--Col10 Dev
            MySheet.Cells[last, 10] = dev;
        }

        private void materialRaisedButton3_Click(object sender, EventArgs e)
        {
            
            System.Windows.Forms.Application.Exit();
        }
    }
}
