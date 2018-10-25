using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Threading;


namespace WordDocGenerator
{
    public partial class Form1 : Form
    {
        private const int MAX_LOAD_IMG_CNTS = 1024;

        private List<string> lstImgPath;
        private List<string> lstImgDescription;

        private ProgressReport dlgReport;

        private string selectPath;

        private int curLoadedIdx;
        private int totalCheckedCnt;
        private delegate void UpdateList();
        private delegate void AddImgList();

        public Form1()
        {
            InitializeComponent();

            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            lstImgPath = new List<string>();
            lstImgDescription = new List<string>();
            richTextBox1.LostFocus += richTextBox1_LostFocus;

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.ProgressChanged += ProcessChanged_Handler;
            backgroundWorker1.RunWorkerCompleted += RunWorkerCompleted_Handler;
        }

        /// <summary>
        /// progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            while (backgroundWorker3.IsBusy)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    worker.ReportProgress(curLoadedIdx * 100 / totalCheckedCnt);
                    Thread.Sleep(100);
                }
            }
            //dlgReport.Close();
        }

        private void ProcessChanged_Handler(object sender, ProgressChangedEventArgs e)
        {
            dlgReport.SetValue(e.ProgressPercentage);
        }

        private void RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs e)
        {
            dlgReport.Close();
        }

        /// <summary>
        /// Scan folders for all images
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {   
            DirectoryInfo directInfo = new DirectoryInfo(selectPath);
            if (!directInfo.Exists || null == directInfo)
            {
                MessageBox.Show("文件夹不存在");
                return;
            }
            scan(directInfo);
        }

        /// <summary>
        /// Load source folder images
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonLoad_Click(object sender, EventArgs e)
        {
            try
            {
                // choose target folder
                FolderBrowserDialog path = new FolderBrowserDialog();
                if (DialogResult.Cancel == path.ShowDialog())
                {
                    return;
                }
                selectPath = path.SelectedPath;
                lstImgDescription.Clear();
                lstImgPath.Clear();
                listView1.Clear();
                imageList1.Images.Clear();
                curLoadedIdx = 0;
                for (int i = 0; i < MAX_LOAD_IMG_CNTS; i++)
                {
                    lstImgDescription.Add("");
                }

                if (!backgroundWorker2.IsBusy)
                {
                    backgroundWorker2.RunWorkerAsync();
                }

                while (backgroundWorker2.IsBusy)
                {
                    Application.DoEvents();
                }

                // Update listview
                UpdateListviewData();
                toolStripStatusLabel1.Text = "加载完成，共加载" + lstImgPath.Count + "张图片";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Update listview data
        /// </summary>
        private void UpdateListviewData()
        {
            listView1.BeginUpdate();
            for (int i = curLoadedIdx; i < imageList1.Images.Count; i++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.ImageIndex = i;
                lvi.Text = System.IO.Path.GetFileName(lstImgPath[i]);
                this.listView1.Items.Add(lvi);
            }
            curLoadedIdx = imageList1.Images.Count;
            listView1.EndUpdate();
            toolStripStatusLabel1.Text = "已加载" + curLoadedIdx.ToString() + "张图片";
        }

        private void AddImageData()
        {
            imageList1.Images.Add(Image.FromFile(lstImgPath[lstImgPath.Count - 1]));
        }

        /// <summary>
        /// scan directory files
        /// </summary>
        /// <param name="info"></param>
        private void scan(FileSystemInfo info)
        {
            if (!info.Exists)
            {
                return;
            }
            // maximum MAX_LOAD_IMG_CNTS images
            if (imageList1.Images.Count >= MAX_LOAD_IMG_CNTS)
            {
                return;
            }
            DirectoryInfo dir = info as DirectoryInfo;

            if (dir == null)
            {
                return;
            }
            FileSystemInfo[] files = dir.GetFileSystemInfos();
            for(int i = 0; i < files.Length; i++)
            {
                FileInfo file = files[i] as FileInfo;
                if (file != null)
                {
                    // only image files
                    if (file.Extension.ToUpper() == ".JPG" ||
                        file.Extension.ToUpper() == ".JPEG" ||
                        file.Extension.ToUpper() == ".PNG" ||
                        file.Extension.ToUpper() == ".BMP")
                    {
                        lstImgPath.Add(file.DirectoryName + "\\" + file.Name);
                        AddImgList updater = new AddImgList(AddImageData);
                        Invoke(updater);
                    }
                }
                else
                {
                    scan(files[i]);
                }
                if (0 == i % 5)
                {
                    UpdateList updater = new UpdateList(UpdateListviewData);
                    Invoke(updater);
                }
            }
        }

        /// <summary>
        /// Save selected images and word doucument
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSave_Click(object sender, EventArgs e)
        {
            // Get targer folder
            FolderBrowserDialog path = new FolderBrowserDialog();
            if (DialogResult.Cancel == path.ShowDialog())
            {
                return;
            }
            selectPath = path.SelectedPath;

            if (!backgroundWorker3.IsBusy)
            {
                backgroundWorker3.RunWorkerAsync();
            }

            // show progress bar
            curLoadedIdx = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    totalCheckedCnt++;
                }
            }
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
                dlgReport = new ProgressReport();
                dlgReport.ShowDialog();
            }
            while (backgroundWorker1.IsBusy)
            {
                Application.DoEvents();
            }
            dlgReport.Close();
            MessageBox.Show("导出完成");

            return;
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // generate word obejct
                object Nothing = Missing.Value;
                object format = MSWord.WdSaveFormat.wdFormatDocument;
                object EndOfDoc = "\\endofdoc";
                object LinkOfFile = false;
                object SaveDocument = true;
                MSWord.Application wordApp = new MSWord.ApplicationClass();
                MSWord.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

                // look for checked items
                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.Checked)
                    {
                        // save Image and descrptions
                        object range = wordDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        wordDoc.InlineShapes.AddPicture(lstImgPath[item.ImageIndex], ref LinkOfFile, ref SaveDocument, ref range);
                        object what = MSWord.WdGoToItem.wdGoToBookmark;
                        wordApp.Selection.GoTo(what, Nothing, Nothing, EndOfDoc);
                        wordApp.Selection.TypeText("\n" + lstImgDescription[item.ImageIndex] + "\n\n");

                        // copy files
                        string fileName = System.IO.Path.GetFileName(lstImgPath[item.ImageIndex]);
                        if (!File.Exists(selectPath + "\\" + fileName))
                        {
                            File.Copy(lstImgPath[item.ImageIndex], selectPath + "\\" + fileName);
                        }

                        curLoadedIdx++;
                    }
                }

                // Save word doc
                System.DateTime currentTime = new System.DateTime();
                currentTime = System.DateTime.Now;
                object savePath = selectPath + "\\" + currentTime.Year.ToString() + "_" + currentTime.Month.ToString() + "_" +
                    currentTime.Day.ToString() + "_" + currentTime.Hour.ToString() + "_" + currentTime.Minute.ToString() + "_" +
                    currentTime.Second.ToString() + ".doc";
                wordDoc.SaveAs(ref savePath, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                    ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                // close word doc
                wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
                wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Show image descrition
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedItems.Count > 0)
                {
                    richTextBox1.Text = lstImgDescription[listView1.FocusedItem.ImageIndex];
                }
            }
            catch (Exception)
            {
                richTextBox1.Text = "";
            }
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Checked)
            {
                e.Item.Selected = true;
                e.Item.Focused = true;
                try
                {
                    richTextBox1.Text = lstImgDescription[listView1.FocusedItem.ImageIndex];
                    richTextBox1.Focus();
                }
                catch (Exception)
                {
                    richTextBox1.Text = "";
                }
            }
        }

        /// <summary>
        /// Avoid opening image twice
        /// </summary>
        bool inOpenflag = true;

        /// <summary>
        /// Open image with system default image viewer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && inOpenflag)
            {
                
                if (listView1.SelectedItems.Count > 0)
                {
                    inOpenflag = false;
                    string imgPath = lstImgPath[listView1.FocusedItem.ImageIndex];

                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = imgPath;
                    process.StartInfo.Arguments = "rundl132.exe C://WINDOWS//system32//shimgvw.dll";
                    process.StartInfo.UseShellExecute = true;
                    process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    process.Start();
                    process.Close();
                    inOpenflag = true;
                }
            }
        }


        /// <summary>
        /// copy description to list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void richTextBox1_LostFocus(object sender, EventArgs e)
        {
            try
            {
                lstImgDescription[listView1.FocusedItem.ImageIndex] = richTextBox1.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Exit menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Cancel != MessageBox.Show("确定退出程序？", "提示", MessageBoxButtons.OKCancel))
            {
                this.Close();
            }
        }

        /// <summary>
        /// About menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemAbout_Click(object sender, EventArgs e)
        {
            About dlgAbout = new About();
            dlgAbout.StartPosition = FormStartPosition.CenterParent;
            dlgAbout.ShowDialog();
        }

        /// <summary>
        /// Instruction menu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemInstruction_Click(object sender, EventArgs e)
        {
            Instrction dlgInst = new Instrction();
            dlgInst.ShowDialog();
        }

    }
}
