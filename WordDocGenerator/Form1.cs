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
        private const int MAX_LOAD_IMG_CNTS = 4096;

        private List<string> lstImgPath;
        private List<string> lstImgDescription;

        private ProgressReport dlgReport;

        private string selectPath;
        private bool finWriteDocFlag = false;
        private int curLoadedIdx;
        private int totalCheckedCnt;
        private delegate void UpdateList();
        private delegate void AddImgList();
        public delegate void WriteWordDoc();

        public StreamWriter writer;

        public Form1()
        {
            InitializeComponent();

            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            lstImgPath = new List<string>();
            lstImgDescription = new List<string>();
            richTextBox1.LostFocus += richTextBox1_LostFocus;

            bgWorkerProgressBar.WorkerReportsProgress = true;
            bgWorkerProgressBar.WorkerSupportsCancellation = true;
            bgWorkerProgressBar.ProgressChanged += ProgressChanged_Handler;
            bgWorkerProgressBar.RunWorkerCompleted += RunWorkerCompleted_Handler;

            // Create tmp folder
            try
            {
                if (!Directory.Exists(System.Environment.CurrentDirectory + "\\tmp"))
                {
                    Directory.CreateDirectory(System.Environment.CurrentDirectory + "\\tmp");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            if (!File.Exists(".\\log.log"))
            {
                File.Create(".\\log.log").Close();
            }
            writer = File.AppendText(".\\log.log");
        }

        /// <summary>
        /// Load source folder images
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonLoad_Click(object sender, EventArgs e)
        {
            buttonLoad.Enabled = false;
            try
            {
                // choose target folder
                FolderBrowserDialog path = new FolderBrowserDialog();
                path.RootFolder = Environment.SpecialFolder.MyComputer;
                
                if (DialogResult.Cancel == path.ShowDialog())
                {
                    return;
                }
                Environment.SpecialFolder root = path.RootFolder;
                selectPath = path.SelectedPath;

                // Initialization
                lstImgDescription.Clear();
                lstImgPath.Clear();
                listView1.Clear();
                imageList1.Images.Clear();
                curLoadedIdx = 0;

                if (!bgWorkerLoadImgs.IsBusy)
                {
                    bgWorkerLoadImgs.RunWorkerAsync();
                }

                while (bgWorkerLoadImgs.IsBusy)
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
            buttonLoad.Enabled = true;
        }

        /// <summary>
        /// Scan folders for all images
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgWorkerLoadImgs_DoWork(object sender, DoWorkEventArgs e)
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
        /// Update listview data
        /// </summary>
        private void UpdateListviewData()
        {
            listView1.BeginUpdate();
            for (int i = curLoadedIdx; i < imageList1.Images.Count; i++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.ImageIndex = i;
                String fileName = System.IO.Path.GetFileName(lstImgPath[i]);
                lvi.Text = fileName.Length > 16 ? fileName.Substring(0, 16) : fileName;
                this.listView1.Items.Add(lvi);
            }
            curLoadedIdx = imageList1.Images.Count;
            listView1.EndUpdate();
            toolStripStatusLabel1.Text = "已加载" + curLoadedIdx.ToString() + "张图片";
        }

        private void AddImageData()
        {
            //DateTime dt = DateTime.Now;
            //writer.WriteLine(dt.Second.ToString() + ":" + dt.Millisecond.ToString() + " BeginUpdate");
            imageList1.Images.Add(Image.FromFile(lstImgPath[lstImgPath.Count - 1]));
            //dt = DateTime.Now;
            //writer.WriteLine(dt.Second.ToString() + ":" + dt.Millisecond.ToString() + " EndUpdate");
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
            for (int i = 0; i < files.Length; i++)
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
                        lstImgDescription.Add("");

                        AddImageData();
                        //AddImgList updater = new AddImgList(AddImageData);
                        //Invoke(updater);
                    }
                }
                else
                {
                    scan(files[i]);
                }
                if (0 == i % 10)
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

            // Create word doc
            Thread threadDoc = new Thread(new ThreadStart(StartMethod));
            threadDoc.SetApartmentState(ApartmentState.STA);
            threadDoc.Start();

            // show progress bar
            curLoadedIdx = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    totalCheckedCnt++;
                }
            }
            if (!bgWorkerProgressBar.IsBusy) 
            {
                bgWorkerProgressBar.RunWorkerAsync();
                dlgReport = new ProgressReport();
                dlgReport.ShowDialog();
            }
            while (bgWorkerProgressBar.IsBusy)
            {
                Application.DoEvents();
            }
            dlgReport.Close();
            MessageBox.Show("导出完成");

            return;
        }

        private void StartMethod()
        {
            finWriteDocFlag = false;
            WriteDoc();
            finWriteDocFlag = true;
            //this.BeginInvoke(new WriteWordDoc(WriteDoc));
        }

        public void WriteDoc()
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
                object what = MSWord.WdGoToItem.wdGoToBookmark;
                wordApp.Selection.GoTo(what, Nothing, Nothing, EndOfDoc);
                wordApp.Selection.TypeText("\n\n");

                // look for checked items
                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.Checked)
                    {
                        String newFilePath = lstImgPath[item.ImageIndex];
                        // Resize image
                        Image srcImage = Image.FromFile(newFilePath);
                        Image loadImg = srcImage;
                        if (srcImage.Width > 1920 || srcImage.Width > 1080)
                        {
                            double scaleFactor = srcImage.Width >= srcImage.Height ?
                                1920.0 / srcImage.Width : 1080.0 / srcImage.Height;
                            int newWidth = (int)(srcImage.Width * scaleFactor);
                            int newHeight = (int)(srcImage.Height * scaleFactor);
                            loadImg = new Bitmap(newWidth, newHeight);
                            var graphics = Graphics.FromImage(loadImg);
                            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            graphics.DrawImage(srcImage, new Rectangle(0, 0, newWidth, newHeight));

                            var encoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                            var encParams = new System.Drawing.Imaging.EncoderParameters() { Param = new[] { new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 40L) } };
                            try
                            {
                                newFilePath = System.Environment.CurrentDirectory + "\\tmp" + 
                                    newFilePath.Substring(newFilePath.LastIndexOf('\\'));
                                loadImg.Save(newFilePath, encoder, encParams);
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                            }
                        }

                        // Add image
                        object range = wordDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        wordDoc.InlineShapes.AddPicture(newFilePath, ref LinkOfFile, ref SaveDocument, ref range);
                        what = MSWord.WdGoToItem.wdGoToBookmark;
                        wordApp.Selection.GoTo(what, Nothing, Nothing, EndOfDoc);
                        wordApp.Selection.TypeText("\n" + lstImgDescription[item.ImageIndex] + "\n\n");

                        // Add description
                        what = MSWord.WdGoToItem.wdGoToBookmark;
                        wordApp.Selection.GoTo(what, Nothing, Nothing, EndOfDoc);
                        wordApp.Selection.TypeText("\n" + lstImgDescription[item.ImageIndex] + "\n\n");

                        // copy files
#if false
                        string fileName = System.IO.Path.GetFileName(lstImgPath[item.ImageIndex]);
                        if (!File.Exists(selectPath + "\\" + fileName))
                        {
                            File.Copy(lstImgPath[item.ImageIndex], selectPath + "\\" + fileName);
                        }
#endif
                        // delete tmp file
                        try
                        {
                            File.Delete(newFilePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
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

        #region show write doc progress bar
        /// <summary>
        /// progress bar
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void bgWorkerProgressBar_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            while (!finWriteDocFlag)
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
        }

        private void ProgressChanged_Handler(object sender, ProgressChangedEventArgs e)
        {
            dlgReport.SetValue(e.ProgressPercentage);
        }

        private void RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs e)
        {
            dlgReport.Close();
        }
        #endregion

        #region controls behavior
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
            catch (Exception /*ex*/)
            {
                //MessageBox.Show(ex.ToString());
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
                writer.Close();
                this.Close();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.Cancel != MessageBox.Show("确定退出程序？", "提示", MessageBoxButtons.OKCancel))
            {
                writer.Close();
            }
            else
            {
                e.Cancel = true;
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
        #endregion


    }
}
