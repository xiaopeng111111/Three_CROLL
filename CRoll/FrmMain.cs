using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MsWord = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Threading;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
namespace CRoll
{

    public partial class FrmMain : Form
    {

        public FrmMain()
        {
            InitializeComponent();
        }

        public IntPtr mWinHandle;
        public IntPtr mComBoHandle;
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int IParam);
        public const int UPGRAD_PROGRESS = 0x500;
        public const int PROGRESSFULL = 0x501;
        public const int FRESHCOMBO = 0x502;
        public static string oneLine, tmpLine;
        //抽取word文件中的学生信息生成考勤数据
        public void fetchStuInfoFromFile()
        {
            SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 3);
            MsWord.ApplicationClass oWordApplic;  //a reference toWordapplication
            MsWord.Document oDoc;  //a reference to thedocument
            try
            {
                oWordApplic = new MsWord.ApplicationClass();
                object missing = System.Reflection.Missing.Value;
                object owdFileName = DataClass.wordFileName;
                object oFalse = false;
                object oTrue = false;
                oDoc = oWordApplic.Documents.Open(ref owdFileName, ref oFalse, ref oTrue,
                   ref oFalse, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                FileInfo wordFinfo = new FileInfo(DataClass.wordFileName);
                string txtFileName1;
                txtFileName1 = wordFinfo.DirectoryName + "\\tempw1.txt";
                if (File.Exists(txtFileName1))
                {
                    File.Delete(txtFileName1);
                }
                //另存为文本
                object ofileName;
                object oEncoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                ofileName = txtFileName1;
                object wordTxtType = MsWord.WdSaveFormat.wdFormatTextLineBreaks;
                oDoc.SaveAs(ref ofileName, ref wordTxtType, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref oEncoding, ref missing, ref missing, ref missing, ref missing);
                oDoc.Close(ref missing, ref missing, ref missing);
                SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 8);
                //另存为网页
                oDoc = oWordApplic.Documents.Open(ref owdFileName, ref oFalse, ref oTrue,
                    ref oFalse, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                object wordHtmlType = MsWord.WdSaveFormat.wdFormatHTML;
                ofileName = wordFinfo.DirectoryName + "\\htm";
                oDoc.SaveAs(ref ofileName, ref wordHtmlType, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref oEncoding, ref missing, ref missing, ref missing, ref missing);
                oDoc.Close(ref missing, ref missing, ref missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                oWordApplic.Quit(ref missing, ref missing, ref missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWordApplic);
                oWordApplic = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 18);


                StreamReader sw1 = new StreamReader(txtFileName1);
                ArrayList alStuId = new ArrayList();
                ArrayList alStuName = new ArrayList();
                Queue qStrStuID = new Queue();
                oneLine = sw1.ReadLine();
                while (oneLine != null)
                {
                    oneLine = oneLine.Trim();
                    if (oneLine.Length > 0)
                    {
                        char oneChar = oneLine.ToCharArray()[0];
                        int oneCharInt = (int)oneChar;
                        if ((oneCharInt < 58) && (oneCharInt > 47))
                        {
                            qStrStuID.Enqueue(oneLine);
                        }
                        else
                        {
                            if (qStrStuID.Count > 0)
                            {
                                tmpLine = (string)qStrStuID.Dequeue();
                                alStuId.Add(tmpLine);
                                alStuName.Add(oneLine);
                            }
                        }
                    }
                    oneLine = sw1.ReadLine();
                }
                sw1.Close();
                SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 21);
                alPicView = new ArrayList();
                DataClass.imgDirName = wordFinfo.DirectoryName + "\\htm.files";
                DirectoryInfo dirInfo = new DirectoryInfo(DataClass.imgDirName);
                FileInfo[] finfo = dirInfo.GetFiles();   //由图片目录获取文件名集合
                string picname;
                string thumbPicname;  //缩略图标
                string bigPicname;    //大图
                string fileExt;    //文件后缀名
                string delFile;   //要删除的文件
                //将所有非jpg格式图片转存成jpg格式
                for (int i = 0; i < finfo.Length; i++)
                {
                    if ((string.Compare(finfo[i].Extension, ".jpg", true) != 0)
                        && (finfo[i].Name.IndexOf("image") != -1))
                    {
                        //文件后缀不为.jpg，并且包含image字符串
                        fileExt = finfo[i].Extension;
                        delFile = finfo[i].FullName;
                        picname = finfo[i].FullName.Replace(fileExt, ".jpg");
                        Bitmap aPic = new Bitmap(finfo[i].FullName);
                        aPic.Save(picname, ImageFormat.Jpeg);
                        aPic.Dispose();
                        File.Delete(delFile);
                    }
                    if (finfo[i].Name.IndexOf("image") == -1)
                    {
                        delFile = finfo[i].FullName;
                        File.Delete(delFile);
                    }
                }
                SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 23);
                //刷新文件名数组
                finfo = dirInfo.GetFiles("*.jpg");
                int iStu = 0;
                for (int i = 0; i < finfo.Length; i++)
                {
                    picname = finfo[i].Name;
                    picname = picname.Replace("image", "");
                    picname = picname.Replace(".jpg", "");
                    if (Int32.Parse(picname) % 2 == 0)
                    {
                        thumbPicname = DataClass.imgDirName + "\\" + finfo[i].Name;
                        //注意生成的字串为003，整数前须加0填充
                        bigPicname = DataClass.imgDirName + "\\image" +
                            (Int32.Parse(picname) - 1).ToString("D3") + ".jpg";
                        PicView picv = new PicView();
                        picv.strThumbpic = thumbPicname;//小图；
                        picv.strBigpic = bigPicname;   //大图
                        //写入学号和姓名
                        picv.strStuID = (string)alStuId[iStu];
                        picv.strStuName = (string)alStuName[iStu];
                        picv.picVImgIndex = iStu;
                        iStu++;
                        alPicView.Add(picv);
                    }
                }
                SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 35);
                //删除临时文件
                if (File.Exists(txtFileName1))
                {
                    File.Delete(txtFileName1);
                }
                if (File.Exists(wordFinfo.DirectoryName + "\\htm.htm"))
                {
                    File.Delete(wordFinfo.DirectoryName + "\\htm.htm");
                }
                try
                {
                    byte[] valBuf;   //用于获取long或字符串字节数组
                    byte[] tempMsBuf;
                    byte[] tempDataBuf;
                    byte[] dataPicBuf;
                    dataPicBuf = DataClass.msFile.GetBuffer();   //总数据内存流
                    int dataLen = 0;
                    tempMsBuf = DataClass.msTmpFile.GetBuffer();  //临时数据内存流
                    DataClass.msFile.Seek(0, SeekOrigin.Begin);
                    int curMsPos = 0;
                    int readLen;
                    FileStream fsTmp;
                    for (int i = 0; i < alPicView.Count; i++)
                    {
                        PicView picV = (PicView)alPicView[i];
                        //strThumbpic
                        tempDataBuf = Encoding.UTF8.GetBytes(picV.strThumbpic);
                        dataLen = tempDataBuf.Length;  //54
                        valBuf = BitConverter.GetBytes(dataLen);
                        //a小图片字符串字节长度
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4;
                        //b小图图片字符串
                        Buffer.BlockCopy(tempDataBuf, 0, dataPicBuf, curMsPos, dataLen);
                        curMsPos += dataLen;

                        //d小图片图像字节值
                        fsTmp = new FileStream(picV.strThumbpic, FileMode.Open, FileAccess.Read);
                        dataLen = 0;
                        readLen = 1000;
                        for (; readLen > 0; )
                        {
                            readLen = fsTmp.Read(dataPicBuf, curMsPos + 4 + dataLen, 1000);
                            dataLen += readLen;
                        }
                        fsTmp.Close();
                        fsTmp.Dispose();

                        valBuf = BitConverter.GetBytes(dataLen);
                        //c小图片长度值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4 + dataLen;


                        //大图名字符串strBigpic
                        tempDataBuf = Encoding.UTF8.GetBytes(picV.strBigpic);
                        dataLen = tempDataBuf.Length;   //54
                        valBuf = BitConverter.GetBytes(dataLen);
                        //e大图名字符串长度值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4;
                        //f大图名字符串字节序列
                        Buffer.BlockCopy(tempDataBuf, 0, dataPicBuf, curMsPos, dataLen);
                        curMsPos += dataLen;
                        //h大图图片像素字节序列
                        fsTmp = new FileStream(picV.strBigpic, FileMode.Open, FileAccess.Read);
                        dataLen = 0;
                        readLen = 1000;
                        for (; readLen > 0; )
                        {
                            readLen = fsTmp.Read(dataPicBuf, curMsPos + 4 + dataLen, 1000);
                            dataLen += readLen;
                        }
                        fsTmp.Close();
                        fsTmp.Dispose();
                        valBuf = BitConverter.GetBytes(dataLen);
                        //g大图图片像素长度值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4 + dataLen;

                        //strStuID
                        tempDataBuf = Encoding.UTF8.GetBytes(picV.strStuID);
                        dataLen = tempDataBuf.Length;
                        valBuf = BitConverter.GetBytes(dataLen);
                        //i学生ID字符串长度值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4;
                        //j学生ID字符串字节序列
                        Buffer.BlockCopy(tempDataBuf, 0, dataPicBuf, curMsPos, dataLen);
                        curMsPos += dataLen;
                        //strStuName
                        tempDataBuf = Encoding.UTF8.GetBytes(picV.strStuName);
                        dataLen = tempDataBuf.Length;
                        valBuf = BitConverter.GetBytes(dataLen);
                        //k学生姓名字符串长度值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4;
                        //l学生姓名字符串字节序列
                        Buffer.BlockCopy(tempDataBuf, 0, dataPicBuf, curMsPos, dataLen);
                        curMsPos += dataLen;

                        //picViewIndex
                        valBuf = BitConverter.GetBytes(picV.picVImgIndex);
                        //m图片索引值
                        Buffer.BlockCopy(valBuf, 0, dataPicBuf, curMsPos, 4);
                        curMsPos += 4;

                    }
                    SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 92);
                    if (File.Exists("stupic.dat"))
                    {
                        File.Delete("stupic.dat");
                    }
                    FileStream fsData = new FileStream("stupic.dat", FileMode.CreateNew, FileAccess.Write);
                    //内存文件转存入磁盘
                    int WriteLen = 1000;
                    for (int writePos = 0; writePos < curMsPos; )
                    {
                        if ((writePos + WriteLen) < curMsPos)
                        {
                            fsData.Write(dataPicBuf, writePos, WriteLen);
                            writePos += WriteLen;
                        }
                        else
                        {
                            fsData.Write(dataPicBuf, writePos, curMsPos - writePos);
                            writePos = curMsPos;
                        }
                    }
                    SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 98);
                    //关闭文件对象
                    fsData.Flush();
                    fsData.Close();
                    fsData.Dispose();
                    if (File.Exists(DataClass.stuDataFileName))
                    {
                        File.Delete(DataClass.stuDataFileName);
                    }
                    if (File.Exists(DataClass.stuDataFileName))
                    {
                        File.Delete(DataClass.stuDataFileName);
                    }
                    File.Move("stupic.dat", DataClass.stuDataFileName);
                    SendMessage(mWinHandle, FRESHCOMBO, 0, 0);
                    SendMessage(mWinHandle, UPGRAD_PROGRESS, 0, 100);
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                }

            }
            catch (Exception e2)
            {
                MessageBox.Show(tmpLine + oneLine + e2.Message);
            }
            bFetchFile = false;
        }



        //刷新下拉框中的文件列表
        public void RefreshCombo()
        {
            comboBox1.Items.Clear();
            string[] stuFiles = Directory.GetFiles(Application.StartupPath, "*.dat");
            FileInfo fin;
            for (int i = 0; i < stuFiles.Length; i++)
            {
                fin = new FileInfo(stuFiles[i]);
                comboBox1.Items.Add(fin.Name);
            }
            if (comboBox1.Items.Count > 0)
            {
                comboBox1.Text = comboBox1.Items[0].ToString();
            }
        }
        private void ShowStuData()
        {
            dataGridView1.Columns.Add("ColStuName", "姓名");
            dataGridView1.Columns.Add("ColStuID", "学号");
            dataGridView1.Columns.Add("ColImgIndex", "图索引");
            dataGridView1.Columns["ColStuName"].Width = 70;
            dataGridView1.Columns["ColStuID"].Width = 100;
            dataGridView1.Columns["ColImgIndex"].Width = 80;
            dataGridView1.Rows.Clear();
            alSortedID.Sort();
            string str1, str2;
            for (int i = 0; i < alSortedID.Count; i++)
            {
                str1 = alSortedID[i].ToString();
                for (int j = 0; j < alPicView.Count; j++)
                {
                    str2 = ((PicView)alPicView[j]).strStuID;
                    if (string.Compare(str1, str2) == 0)
                    {
                        dataGridView1.Rows.Add(
                            ((PicView)alPicView[j]).strStuName,
                             ((PicView)alPicView[j]).strStuID,
                              ((PicView)alPicView[j]).picVImgIndex);
                    }
                }
            }
        }
        //重载窗体信息处理
        int progressValue;
        protected override void DefWndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case UPGRAD_PROGRESS:
                    progressValue = (int)m.LParam;
                    progressBar1.Value = progressValue;
                    label2.Text = string.Format("{0}", progressValue);
                    break;
                case FRESHCOMBO:
                    this.RefreshCombo();
                    break;
                case PROGRESSFULL:
                    //progressBar1.Visible=false;
                    break;
                default:
                    base.DefWndProc(ref m);
                    break;
            }
        }


        private void FrmMain_Load(object sender, EventArgs e)
        {
            mWinHandle = this.Handle;
            mComBoHandle = this.comboBox1.Handle;
            RefreshCombo();   //刷新文件列表
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LoadPicData();
            ShowStuData();
        }
        private void ChangeStu(int RowIndex)
        {
            dataGridView1.Rows[RowIndex].Selected = true;

            string strImgIndex;
            strImgIndex = dataGridView1.Rows[RowIndex].Cells["ColImgIndex"].Value.ToString();
            pictureBox1.Image = (Bitmap)alBmp[Int32.Parse(strImgIndex)];
            label1.Text = ((PicView)alPicView[Int32.Parse(strImgIndex)]).strStuName;
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            { ChangeStu(e.RowIndex); }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                ChangeStu(e.RowIndex);
            }
        }
        ArrayList alPicView = new ArrayList();
        ArrayList alSortedID = new ArrayList();
        ArrayList alBmp = new ArrayList();
        ArrayList alBmpIndex = new ArrayList();
        bool bFetchFile;
        //选择Word文件，读取文件内容包括学生照片、学号与姓名
        private void button4_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            openFileDialog1.InitialDirectory = Application.StartupPath;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                DataClass.wordFileName = openFileDialog1.FileName;
                FileInfo wdfInfo = new FileInfo(DataClass.wordFileName);
                DataClass.stuDataFileName = Application.StartupPath + "\\"
                    + wdfInfo.Name.Substring(0, wdfInfo.Name.Length - 4) + ".dat";
                if (string.Compare(wdfInfo.Extension, ".doc", true) == 0)
                {
                    if (!bFetchFile)
                    {
                        bFetchFile = true;
                        ThreadStart thStart = new ThreadStart(fetchStuInfoFromFile);
                        Thread thr = new Thread(thStart);
                        thr.Start();
                    }
                }
                else
                {
                    MessageBox.Show("请选择word 2003格式文档");
                }
            }
        }





        private void LoadPicData()
        {
            string stuFile = Application.StartupPath + "\\" + comboBox1.Text;
            if (File.Exists(stuFile))
            {
                //图片文件存在，读入学生图片和信息内容
                //首先清楚所有数据数组
                alPicView.Clear();
                alSortedID.Clear();
                alBmp.Clear();
                alBmpIndex.Clear();

                byte[] dataPicBuf;
                DataClass.msFile.Seek(0, SeekOrigin.Begin);
                dataPicBuf = DataClass.msFile.GetBuffer();   //总数据内存刘
                FileInfo finfo = new FileInfo(stuFile);
                int fileLen = (int)finfo.Length;
                FileStream fsData = new FileStream(stuFile, FileMode.Open);
                //内存文件转存入磁盘
                int readLen = 0, readPos = 0;
                for (; readPos < fileLen - 1; )
                {
                    readLen = fsData.Read(dataPicBuf, readPos, 1000);
                    readPos += readLen;
                }
                //关闭文件
                fsData.Close();
                fsData.Dispose();

                byte[] tempDataBuf;
                byte[] tmpPicBuf;
                tempDataBuf = new byte[2000];
                int dataLen = 0, curPos = 0;
                Bitmap bmpThumb, bmpBig;
                MemoryStream mspic;
                for (; curPos < fileLen - 1; )
                {
                    //strThumbpic,strBigpic,strStuID,strStuName,picViewIndex
                    PicView picv = new PicView();
                    //a小图片字符串字节长度
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    //b小图图片字符串
                    Buffer.BlockCopy(dataPicBuf, curPos, tempDataBuf, 0, dataLen);
                    curPos += dataLen;
                    picv.strThumbpic = Encoding.UTF8.GetString(tempDataBuf, 0, dataLen);

                    //c小图图片长度值
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    tmpPicBuf = DataClass.msTmpFile.GetBuffer();
                    for (int i = 0; i < dataLen; i++)
                    {
                        Buffer.SetByte(tmpPicBuf, i, 0);

                    }
                    //d小图片图像字节值
                    Buffer.BlockCopy(dataPicBuf, curPos, tmpPicBuf, 0, dataLen);
                    curPos += dataLen;
                    mspic = new MemoryStream(tmpPicBuf, 0, dataLen);
                    bmpThumb = new Bitmap(mspic);
                    //大图名字符串strBigpic

                    //e大图名字符串长度值
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    //f大图名字符串字节序列
                    Buffer.BlockCopy(dataPicBuf, curPos, tempDataBuf, 0, dataLen);
                    curPos += dataLen;
                    picv.strBigpic = Encoding.UTF8.GetString(tempDataBuf, 0, dataLen);

                    //g大图图片像素长度值
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    //h大图图片像素字节序列
                    Buffer.BlockCopy(dataPicBuf, curPos, tmpPicBuf, 0, dataLen);
                    curPos += dataLen;
                    mspic = new MemoryStream(tmpPicBuf, 0, dataLen);
                    bmpBig = new Bitmap(mspic);

                    //i学生ID字符串长度值
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    //j学生ID字符串字节序列
                    Buffer.BlockCopy(dataPicBuf, curPos, tempDataBuf, 0, dataLen);
                    curPos += dataLen;
                    picv.strStuID = Encoding.UTF8.GetString(tempDataBuf, 0, dataLen);

                    //k学生姓名字符串长度值
                    dataLen = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;
                    //l学生姓名字符串字节序列
                    Buffer.BlockCopy(dataPicBuf, curPos, tempDataBuf, 0, dataLen);
                    curPos += dataLen;
                    picv.strStuName = Encoding.UTF8.GetString(tempDataBuf, 0, dataLen);

                    //m图片索引值
                    picv.picVImgIndex = BitConverter.ToInt32(dataPicBuf, curPos);
                    curPos += 4;

                    alPicView.Add(picv);
                    alSortedID.Add(picv.strStuID);
                    alBmp.Add(bmpBig);
                    alBmpIndex.Add(picv.picVImgIndex);
                }
            }
        }
    }




}
