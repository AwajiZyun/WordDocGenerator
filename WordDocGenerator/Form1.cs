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


namespace WordDocGenerator
{
    public partial class Form1 : Form
    {
        List<string> lstImgPath;
        List<string> lstImgDescription;

        int progressPercent;

        public Form1()
        {
            InitializeComponent();

            imageList1.ColorDepth = ColorDepth.Depth32Bit;

            lstImgPath = new List<string>();
            lstImgDescription = new List<string>();
            for (int i = 0; i < 100; i++)
            {
                lstImgDescription.Add("");
            }
            richTextBox1.LostFocus += richTextBox1_LostFocus;
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
                
                // scan folder
                DirectoryInfo directInfo = new DirectoryInfo(path.SelectedPath);
                if (!directInfo.Exists || null == directInfo)
                {
                    MessageBox.Show("文件夹不存在");
                    return;
                }
                scan(directInfo);

                listView1.BeginUpdate();
                for (int i = 0; i < imageList1.Images.Count; i++)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.ImageIndex = i;
                    lvi.Text = "item" + i;
                    this.listView1.Items.Add(lvi);
                }
                listView1.EndUpdate();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
            DirectoryInfo dir = info as DirectoryInfo;

            if (dir == null)
            {
                return;
            }
            FileSystemInfo[] files = dir.GetFileSystemInfos();
            for(int i =0; i< files.Length;i++)
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
                        imageList1.Images.Add(Image.FromFile(file.DirectoryName + "\\" + file.Name));
                        lstImgPath.Add(file.DirectoryName + "\\" + file.Name);
                    }
                }
                else
                {
                    scan(files[i]);
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

                // Get targer folder
                FolderBrowserDialog path = new FolderBrowserDialog();
                if (DialogResult.Cancel == path.ShowDialog())
                {
                    return;
                }

                // look for checked items
                int totalChecked = 0;
                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.Checked)
                    {
                        totalChecked++;
                    }
                }
                progressPercent = 0;
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
                        File.Copy(lstImgPath[item.ImageIndex], path.SelectedPath + "\\" + fileName);
                    }
                }

                // Save word doc
                System.DateTime currentTime=new System.DateTime();
                currentTime = System.DateTime.Now;
                object savePath = path.SelectedPath + "\\" + currentTime.Year.ToString() + "_" + currentTime.Month.ToString() + "_" +
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
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ProgressReport dlgProgress = new ProgressReport();

            }
            catch(Exception ex)
            {
                ;
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
            catch (Exception ex)
            {
                richTextBox1.Text = "";
                //MessageBox.Show(listView1.FocusedItem.ImageIndex.ToString() + "\n" + ex.ToString());
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


    }
}
