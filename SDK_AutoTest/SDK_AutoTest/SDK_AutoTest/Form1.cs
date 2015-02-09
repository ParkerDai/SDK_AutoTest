using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using SE190X;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        // Boolean flag used to determine when a character other than a number is entered.
        private bool nonNumberEntered = false;

        private bool STOP_WHEN_FAIL = false;
        private bool WAIT = false;
        private bool Run_Stop;

        TcpClient telnet = new TcpClient();
        NetworkStream telentStream; // 宣告網路資料流變數
        byte[] bytWrite_telnet;
        byte[] bytRead_telnet;

        public Thread rcvThread;

        // Ping
        System.Net.NetworkInformation.Ping objping = new System.Net.NetworkInformation.Ping();

        // C#取得主程式路徑(Application Path)
        string appPATH = Application.StartupPath;

        string fnameTmp;
        string MODEL_NAME, TARGET_IP;   // MODEL_NAME 程式測試&判斷使用，由文字檔檔名決定，強制大寫
        string model_name;              // model_name 出廠設定使用，由文字檔內文第一行決定，強制大寫
        static uint dev_num = 50;

        // new:建構ArrayList物件
        ArrayList TEST_STATUS = new ArrayList(50); // 0:未測試,1:PASS,2:fail,3:error

        ArrayList TEST_FunLog = new ArrayList(50);
        int idx_funlog;
        string[] TEST_RESULT;

        public Label[] lblFunction = new Label[dev_num];

        string COM_function, CAN_functiom, CAN_loopback, SD_function, WaitKey, USR, PWD;
        string data;
        int TestFun_MaxIdx;
        int row_num;
        int MOUSE_Idx, Test_Idx;
        DateTime time;
        Process proc;
        int[] COM_PID = new int[2];
        int NET_PID;
        int secretX;
        bool chooseStart = false;
        string tester_forExcel, productNum_forExcel, coreSN_forExcel, lanSN_forExcel, uartSN_forExcel, serial1SN_forExcel, serial2SN_forExcel, serial3SN_forExcel, serial4SN_forExcel;
        string startTime, endTime;

        string rxContents;
        string rxContents_EUT;

        public Form1()
        {
            InitializeComponent();

            // 表單中的焦點永遠在某個控制項上
            //this.Activated += new EventHandler(delegate(object o, EventArgs e)
            //{
            //    this.txt_Tx.Focus();
            //});
            //this.txt_Tx.Leave += new EventHandler(delegate(object o, EventArgs e)
            //{
            //    this.txt_Tx.Focus();
            //});
        }

        public delegate void myUICallBack(string myStr, TextBox txt); // delegate 委派；Invoke 調用

        /// <summary>
        /// 更新主線程的UI (txt_Rx.text) = Display
        /// </summary>
        /// <param name="myStr">字串</param>
        /// <param name="txt">指定的控制項，限定有Text屬性</param>
        public void myUI(string myStr, TextBox txt)
        {
            if (txt.InvokeRequired)    // if (this.InvokeRequired)
            {
                myUICallBack myUpdate = new myUICallBack(myUI);
                this.Invoke(myUpdate, myStr, txt);
            }
            else
            {
                int i;
                string[] line;
                int ptr = myStr.IndexOf("\r\n", 0); // vb6: ptr = InStr(1, keyword, vbCrLf, vbTextCompare)
                //Debug.Print(ptr.ToString());
                if (ptr == -1)  // Instr與IndexOf的起始位置不同，結果的表達也不同(參見MSDN)
                {
                    ptr = myStr.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), 0); // ←[J
                    if (ptr != -1)
                        ptr = ptr + 2;
                }
                // 判斷 txt_Rx.Text 中的字串是否超出最大長度
                if (txt.Text.Length + myStr.Length >= txt.MaxLength)
                {
                    if (myStr.Length >= txt.MaxLength)
                        //txt.Text = myStr.Substring(myStr.Length - 1 - txt.MaxLength, txt.MaxLength); // 右邊(S.Length-1-指定長度，指定長度)
                        txt.Text = myStr.Substring((myStr.Length - txt.MaxLength));
                    else if (txt.Text.Length >= myStr.Length)
                        //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.Text.Length - myStr.Length), (txt.Text.Length - myStr.Length));
                        txt.Text = txt.Text.Substring((txt.Text.Length - (txt.Text.Length - myStr.Length)));
                    else
                        //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.MaxLength - myStr.Length), (txt.MaxLength - myStr.Length));
                        txt.Text = txt.Text.Substring((txt.Text.Length - (txt.MaxLength - myStr.Length)));
                }
                txt.Text = txt.Text + myStr;

                // 處理((char)8)，例如開機倒數321訊息
                //int ptr1 = myStr.IndexOf(((char)8).ToString(), 0);
                //if (ptr1 != -1)
                //{
                //    while (((txt_Rx.Text.IndexOf(((char)8).ToString(), 0) + 1) > 0))
                //    {
                //        ptr1 = (txt_Rx.Text.IndexOf(((char)8).ToString(), 0) + 1);
                //        if ((ptr1 > 1))
                //        {
                //            txt_Rx.Text = (txt_Rx.Text.Substring(0, (ptr1 - 2)) + txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.Text.Length - ptr1))));
                //        }
                //        else
                //        {
                //            txt_Rx.Text = (txt_Rx.Text.Substring(0, (ptr1 - 1)) + txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.Text.Length - ptr1))));
                //        }
                //    }
                //}

                data = data + myStr;
                //Console.WriteLine(data);
                if (ptr == -1 || ptr == 0)
                {
                    return;
                }

                // 處理終端機上下鍵的動作(顯示上一個指令)
                if (myStr.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString()) != -1)
                {
                    line = txt_Rx.Text.Split('\r');
                    txt_Rx.Text = string.Empty;     // 文字會重複的問題
                    string Rx_tmp = string.Empty;   // Rx_tmp 解決卷軸滾動視覺效果
                    for (i = 1; i < line.GetUpperBound(0) - 1; i++)
                    {
                        Rx_tmp = Rx_tmp + "\r\n" + line[i];
                    }
                    txt_Rx.Text = Rx_tmp + "\r\n" + line[i + 1].Replace(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), "");
                }

                // 開機完，自動輸入USR、PWD
                if (WAIT)
                {
                    if (WaitKey == null)
                    {
                        WaitKey = string.Empty;
                    }
                    else if (WaitKey != string.Empty)
                    {
                        if (data.Contains(WaitKey))
                        {
                            if (WaitKey.Equals("login", StringComparison.OrdinalIgnoreCase))
                            {
                                if (serialPort1.IsOpen)
                                {
                                    //serialPort1.DiscardOutBuffer(); // 捨棄序列驅動程式傳輸緩衝區的資料
                                    if (!String.IsNullOrEmpty(USR)) //USR!=null || USR!=string.empty
                                    {
                                        serialPort1.Write(USR + ((char)13).ToString());
                                        System.Threading.Thread.Sleep(100);
                                        serialPort1.Write(PWD + ((char)13).ToString());
                                    }
                                    else
                                    {
                                        serialPort1.Write("root" + ((char)13).ToString());
                                        System.Threading.Thread.Sleep(100);
                                        serialPort1.Write("root" + ((char)13).ToString());
                                    }
                                }
                            }
                            WaitKey = string.Empty;
                            WAIT = false;
                        }
                    }
                }
                data = myStr.Substring((myStr.Length - (myStr.Length - ptr)));
                //Debug.Print(data);
            }
        }

        private void serialPort1_Display(object sender, EventArgs e)
        {
            int i;
            string[] line;
            int ptr = rxContents.IndexOf("\r\n", 0); // vb6: ptr = InStr(1, keyword, vbCrLf, vbTextCompare)
            //Debug.Print(ptr.ToString());
            if (ptr == -1)  // Instr與IndexOf的起始位置不同，結果的表達也不同(參見MSDN)
            {
                ptr = rxContents.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), 0); // ←[J
                if (ptr != -1)
                    ptr = ptr + 2;
            }
            // 判斷 txt_Rx.Text 中的字串是否超出最大長度
            if (txt_Rx.Text.Length + rxContents.Length >= txt_Rx.MaxLength)
            {
                if (rxContents.Length >= txt_Rx.MaxLength)
                    //txt.Text = myStr.Substring(myStr.Length - 1 - txt.MaxLength, txt.MaxLength); // 右邊(S.Length-1-指定長度，指定長度)
                    txt_Rx.Text = rxContents.Substring((rxContents.Length - txt_Rx.MaxLength));
                else if (txt_Rx.Text.Length >= rxContents.Length)
                    //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.Text.Length - myStr.Length), (txt.Text.Length - myStr.Length));
                    txt_Rx.Text = txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.Text.Length - rxContents.Length)));
                else
                    //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.MaxLength - myStr.Length), (txt.MaxLength - myStr.Length));
                    txt_Rx.Text = txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.MaxLength - rxContents.Length)));
            }
            txt_Rx.Text = txt_Rx.Text + rxContents;

            // 處理((char)8)，例如開機倒數321訊息
            // Bug!!!!! 快速接收到連續的重複訊息，會讓執行緒暴增...SE7416暫時不進入此段程式碼
            int ptr1 = rxContents.IndexOf(((char)8).ToString(), 0);
            if (ptr1 != -1)
            {
                while (((txt_Rx.Text.IndexOf(((char)8).ToString(), 0) + 1) > 0))
                {
                    ptr1 = (txt_Rx.Text.IndexOf(((char)8).ToString(), 0) + 1);
                    if ((ptr1 > 1))
                    {
                        txt_Rx.Text = (txt_Rx.Text.Substring(0, (ptr1 - 2)) + txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.Text.Length - ptr1))));
                        Debug.Print(txt_Rx.Text);
                    }
                    else
                    {
                        txt_Rx.Text = (txt_Rx.Text.Substring(0, (ptr1 - 1)) + txt_Rx.Text.Substring((txt_Rx.Text.Length - (txt_Rx.Text.Length - ptr1))));
                        Debug.Print(txt_Rx.Text);
                    }
                }
            }

            data = data + rxContents;
            //Console.WriteLine(data);
            if (ptr == -1 || ptr == 0)
            {
                return;
            }

            // 處理終端機上下鍵的動作(顯示上一個指令)
            if (rxContents.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString()) != -1)
            {
                line = txt_Rx.Text.Split('\r');
                txt_Rx.Text = string.Empty;     // 文字會重複的問題
                string Rx_tmp = string.Empty;   // Rx_tmp 解決卷軸滾動視覺效果
                for (i = 1; i < line.GetUpperBound(0) - 1; i++)
                {
                    Rx_tmp = Rx_tmp + "\r\n" + line[i];
                }
                txt_Rx.Text = Rx_tmp + "\r\n" + line[i + 1].Replace(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), "");
            }

            // 開機完，自動輸入USR、PWD
            if (WAIT)
            {
                if (WaitKey == null)
                {
                    WaitKey = string.Empty;
                }
                else if (WaitKey != string.Empty)
                {
                    if (data.Contains(WaitKey))
                    {
                        if (WaitKey.Equals("login", StringComparison.OrdinalIgnoreCase))
                        {
                            if (serialPort1.IsOpen)
                            {
                                //serialPort1.DiscardOutBuffer(); // 捨棄序列驅動程式發送的緩衝區的資料
                                if (!String.IsNullOrEmpty(USR)) //USR!=null || USR!=string.empty
                                {
                                    serialPort1.Write(USR + ((char)13).ToString());
                                    System.Threading.Thread.Sleep(100);
                                    serialPort1.Write(PWD + ((char)13).ToString());
                                }
                                else
                                {
                                    serialPort1.Write("root" + ((char)13).ToString());
                                    System.Threading.Thread.Sleep(100);
                                    serialPort1.Write("root" + ((char)13).ToString());
                                }
                            }
                        }
                        WaitKey = string.Empty;
                        WAIT = false;       // debug: check "data"
                    }
                }
            }
            data = rxContents.Substring((rxContents.Length - (rxContents.Length - ptr)));
            //Debug.Print(data);
        }

        private void serialPort2_Display(object sender, EventArgs e)
        {
            int i;
            string[] line;
            int ptr = rxContents_EUT.IndexOf("\r\n", 0); // vb6: ptr = InStr(1, keyword, vbCrLf, vbTextCompare)
            //Debug.Print(ptr.ToString());
            if (ptr == -1)  // Instr與IndexOf的起始位置不同，結果的表達也不同(參見MSDN)
            {
                ptr = rxContents_EUT.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), 0); // ←[J
                if (ptr != -1)
                    ptr = ptr + 2;
            }
            // 判斷 txt_Rx_EUT.Text 中的字串是否超出最大長度
            if (txt_Rx_EUT.Text.Length + rxContents_EUT.Length >= txt_Rx_EUT.MaxLength)
            {
                if (rxContents_EUT.Length >= txt_Rx_EUT.MaxLength)
                    //txt.Text = myStr.Substring(myStr.Length - 1 - txt.MaxLength, txt.MaxLength); // 右邊(S.Length-1-指定長度，指定長度)
                    txt_Rx_EUT.Text = rxContents_EUT.Substring((rxContents_EUT.Length - txt_Rx_EUT.MaxLength));
                else if (txt_Rx_EUT.Text.Length >= rxContents_EUT.Length)
                    //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.Text.Length - myStr.Length), (txt.Text.Length - myStr.Length));
                    txt_Rx_EUT.Text = txt_Rx_EUT.Text.Substring((txt_Rx_EUT.Text.Length - (txt_Rx_EUT.Text.Length - rxContents_EUT.Length)));
                else
                    //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.MaxLength - myStr.Length), (txt.MaxLength - myStr.Length));
                    txt_Rx_EUT.Text = txt_Rx_EUT.Text.Substring((txt_Rx_EUT.Text.Length - (txt_Rx_EUT.MaxLength - rxContents_EUT.Length)));
            }
            txt_Rx_EUT.Text = txt_Rx_EUT.Text + rxContents_EUT;

            // 處理((char)8)，例如開機倒數321訊息
            // Bug!!!!! 快速接收到連續的重複訊息，會讓執行緒暴增...SE7416暫時不進入此段程式碼
            int ptr1 = rxContents_EUT.IndexOf(((char)8).ToString(), 0);
            if (ptr1 != -1)
            {
                while (((txt_Rx_EUT.Text.IndexOf(((char)8).ToString(), 0) + 1) > 0))
                {
                    ptr1 = (txt_Rx_EUT.Text.IndexOf(((char)8).ToString(), 0) + 1);
                    if ((ptr1 > 1))
                    {
                        txt_Rx_EUT.Text = (txt_Rx_EUT.Text.Substring(0, (ptr1 - 2)) + txt_Rx_EUT.Text.Substring((txt_Rx_EUT.Text.Length - (txt_Rx_EUT.Text.Length - ptr1))));
                        Debug.Print(txt_Rx_EUT.Text);
                    }
                    else
                    {
                        txt_Rx_EUT.Text = (txt_Rx_EUT.Text.Substring(0, (ptr1 - 1)) + txt_Rx_EUT.Text.Substring((txt_Rx_EUT.Text.Length - (txt_Rx_EUT.Text.Length - ptr1))));
                        Debug.Print(txt_Rx_EUT.Text);
                    }
                }
            }

            data = data + rxContents_EUT;
            //Console.WriteLine(data);
            if (ptr == -1 || ptr == 0)
            {
                return;
            }

            // 處理終端機上下鍵的動作(顯示上一個指令)
            if (rxContents_EUT.IndexOf(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString()) != -1)
            {
                line = txt_Rx_EUT.Text.Split('\r');
                txt_Rx_EUT.Text = string.Empty;     // 文字會重複的問題
                string Rx_tmp = string.Empty;   // Rx_tmp 解決卷軸滾動視覺效果
                for (i = 1; i < line.GetUpperBound(0) - 1; i++)
                {
                    Rx_tmp = Rx_tmp + "\r\n" + line[i];
                }
                txt_Rx_EUT.Text = Rx_tmp + "\r\n" + line[i + 1].Replace(((char)27).ToString() + ((char)91).ToString() + ((char)74).ToString(), "");
            }

            // 開機完，自動輸入USR、PWD
            if (WAIT)
            {
                if (WaitKey == null)
                {
                    WaitKey = string.Empty;
                }
                else if (WaitKey != string.Empty)
                {
                    if (data.Contains(WaitKey))
                    {
                        if (WaitKey.Equals("login", StringComparison.OrdinalIgnoreCase))
                        {
                            if (serialPort1.IsOpen)
                            {
                                //serialPort1.DiscardOutBuffer(); // 捨棄序列驅動程式發送的緩衝區的資料
                                if (!String.IsNullOrEmpty(USR)) //USR!=null || USR!=string.empty
                                {
                                    serialPort1.Write(USR + ((char)13).ToString());
                                    System.Threading.Thread.Sleep(100);
                                    serialPort1.Write(PWD + ((char)13).ToString());
                                }
                                else
                                {
                                    serialPort1.Write("root" + ((char)13).ToString());
                                    System.Threading.Thread.Sleep(100);
                                    serialPort1.Write("root" + ((char)13).ToString());
                                }
                            }
                        }
                        WaitKey = string.Empty;
                        WAIT = false;       // debug: check "data"
                    }
                }
            }
            data = rxContents_EUT.Substring((rxContents_EUT.Length - (rxContents_EUT.Length - ptr)));
            //Debug.Print(data);
        }

        public void RecNote(int idx, string note)
        {
            string tmpNote = string.Empty;
            DateTime time = DateTime.Now;
            tmpNote = String.Format("{0:00}:{1:00}:{2:00}", time.Hour, time.Minute, time.Second) + " [" + lblFunction[idx].Tag + "]" + ": " + note + "\r\n";
            noteUI(tmpNote, txt_Note);
        }

        public delegate void noteUICallBack(string myStr, TextBox txt); // delegate 委派；Invoke 調用

        /// <summary>
        /// 更新主線程的UI (txt_Note.text)
        /// </summary>
        /// <param name="myStr">字串</param>
        /// <param name="txt">指定的控制項，限定有Text屬性</param>
        public void noteUI(string myStr, TextBox txt)
        {
            if (txt.InvokeRequired)    // if (this.InvokeRequired)
            {
                noteUICallBack myUpdate = new noteUICallBack(noteUI);
                this.Invoke(myUpdate, myStr, txt);
            }
            else
            {
                // 判斷 txt.Text 中的字串是否超出最大長度
                if (txt.Text.Length + myStr.Length >= txt.MaxLength)
                {
                    if (myStr.Length >= txt.MaxLength)
                        //txt.Text = myStr.Substring(myStr.Length - 1 - txt.MaxLength, txt.MaxLength); // 右邊(S.Length-1-指定長度，指定長度)
                        txt.Text = myStr.Substring((myStr.Length - txt.MaxLength));
                    else if (txt.Text.Length >= myStr.Length)
                        //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.Text.Length - myStr.Length), (txt.Text.Length - myStr.Length));
                        txt.Text = txt.Text.Substring((txt.Text.Length - (txt.Text.Length - myStr.Length)));
                    else
                        //txt.Text = txt.Text.Substring(txt.Text.Length - 1 - (txt.MaxLength - myStr.Length), (txt.MaxLength - myStr.Length));
                        txt.Text = txt.Text.Substring((txt.Text.Length - (txt.MaxLength - myStr.Length)));
                }
                txt.Text = txt.Text + myStr;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;   // 漏斗指標
            consoleToolStripMenuItem_CheckStateChanged(null, null);

            // 獲取電腦的有效串口
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                cmbDutCom.Items.Add(port);
                //cmbEutCom.Items.Add(port);
            }
            cmbDutCom.Sorted = true;
            cmbDutCom.SelectedIndex = 0;
            //cmbEutCom.Sorted = true;
            //cmbEutCom.SelectedIndex = 1;

            if (IsIP(txtDutIP.Text))
            {
                TARGET_IP = txtDutIP.Text;
            }
            //if (IsIP(txtEutIP.Text))
            //{
            //    TARGET_eutIP = txtEutIP.Text;
            //}

            this.Cursor = Cursors.Default;      // 還原預設指標
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            serialPort1_Close();
            //serialPort2_Close();
            if (telnet.Connected) { telnet.Close(); }
            Application.Exit();
        }

        #region Shell

        private int Shell(string FilePath, string FileName)
        {
            try
            {
                ////////////////////// like VB 【shell】 ///////////////////////
                //System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc = new Process();
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                proc.EnableRaisingEvents = false;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.FileName = FilePath + "\\" + FileName;
                proc.Start();
                return proc.Id;
                ////////////////////////////////////////////////////////////////
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " ' " + FileName + " ' ", "Shell error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }

        private void CloseShell(int pid)
        {
            //if (!Process.GetProcessById(pid).HasExited)
            //{
            //    // Close process by sending a close message to its main window.
            //    Process.GetProcessById(pid).CloseMainWindow();
            //    Process.GetProcessById(pid).WaitForExit(3000);
            //}
            if (!Process.GetProcessById(pid).HasExited)
            {
                Process.GetProcessById(pid).Kill();
                Process.GetProcessById(pid).WaitForExit(1000);
            }
        }

        #endregion Shell

        private void cmdOpeFile_Click(object sender, EventArgs e)
        {
            string[] cmd;
            int n = 0;
            String line;
            STOP_WHEN_FAIL = false;

            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Multiselect = false;
            openFileDialog1.InitialDirectory = appPATH;
            openFileDialog1.Filter = "純文字檔(*.txt)|*.txt|All(*.*)|*.*";
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fnameTmp = openFileDialog1.SafeFileName;
                    //fnameTmp = openFileDialog1.FileName.Replace(appPATH + "\\", string.Empty);
                    cmdOpeFile.Text = fnameTmp;

                    // Pass the file path and file name to the StreamReader constructor
                    using (StreamReader sr = new StreamReader(fnameTmp, Encoding.ASCII))
                    {
                        // 1. Read the first line of text
                        line = sr.ReadLine();
                        cmd = line.Split(' ');
                        if (cmd.GetUpperBound(0) < 1)
                        {
                            MessageBox.Show("檔案第一行錯誤，格式應該為 Model IP User Password ", "Error Message");
                            sr.Close();
                            return;
                        }
                        else
                            if (!IsIP(cmd[1]))
                            {
                                MessageBox.Show("檔案第一行錯誤，請檢查 IP 是否輸入正確 ", "Error Message");
                                sr.Close();
                                return;
                            }
                        Shell(appPATH, "arp-d.bat");
                        // model_name 出廠設定使用，由文字檔內文第一行決定，強制大寫
                        model_name = cmd[0].ToUpper();
                        // MODEL_NAME 程式測試&判斷使用，由文字檔檔名決定，強制大寫
                        MODEL_NAME = (fnameTmp.Replace(".txt", string.Empty)).ToUpper();
                        TARGET_IP = cmd[1];
                        USR = cmd[2];
                        if (cmd.GetUpperBound(0) > 2) { PWD = cmd[3]; }
                        else { PWD = string.Empty; }

                        this.Text = MODEL_NAME + "   SDK Auto-test";
                        chkLoop.Checked = false;
                        Test_Idx = 0;
                        Run_Stop = true;
                        //SYSTEM = 0;
                        serialPort1_Close();
                        //serialPort2_Close();
                        if (telnet.Connected) { telnet.Close(); }

                        MappingFunction();

                        RemoveControl(TestFun_MaxIdx);   // Initial Label
                        txt_Note.Text = string.Empty;
                        txt_Rx.Text = string.Empty;
                        TEST_STATUS.Clear();    // 將所有元素移除(Initial)
                        TEST_FunLog.Clear();

                        // 2. Continue to read until you reach end of file
                        line = sr.ReadLine();
                        while (line != null)
                        {
                            if (line != string.Empty)
                            {
                                cmd = line.Split(' ');
                                switch (cmd[0].ToUpper())
                                {
                                    case "APINFOR":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("APINFOR");
                                        break;
                                    case "BUZZER":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("BUZZER");
                                        break;
                                    case "CONSOLE-DUT":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0); // 0:將TEST_STATUS狀態設定為未測試
                                        break;
                                    case "CONSOLE-EUT":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0); // 0:將TEST_STATUS狀態設定為未測試
                                        break;
                                    case "CPU":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("CPU");
                                        break;
                                    case "COM":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("COM");
                                        break;
                                    case "COMTOCOM":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("COMTOCOM");
                                        break;
                                    case "CANTOCAN":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("CANTOCAN");
                                        break;
                                    case "DI":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("DI");
                                        break;
                                    case "DO":  // =Relay
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("DO");
                                        break;
                                    case "DOTODI":  // =DIO
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("DOTODI");
                                        break;
                                    case "DELETE":  // delete files in jffs2
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "EEPROM":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("EEPROM");
                                        break;
                                    case "FTP":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("FTP");
                                        break;
                                    case "FLASH":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("FLASH");
                                        break;
                                    case "FACTORYFILES":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "GWD":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("GWD");
                                        break;
                                    case "GPS":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("GPS");
                                        break;
                                    case "LOADTOOLS":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "LAN":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("LAN");
                                        break;
                                    case "MEMORY":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("MEMORY");
                                        break;
                                    case "NETWORK":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("NETWORK");
                                        break;
                                    case "NTP":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("NTP");
                                        break;
                                    case "NAT":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("NAT");
                                        break;
                                    case "POWER":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "POWERDETECT":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "RESTART":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("RESTART");
                                        break;
                                    case "RTC":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("RTC");
                                        break;
                                    case "RESTORE":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("RESTORE");
                                        break;
                                    case "SLEEP":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "SYSTEM":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "STOP":
                                        if (cmd[1].ToUpper() == "WHEN" && cmd[2].ToUpper() == "FAIL")
                                            STOP_WHEN_FAIL = true;
                                        break;
                                    case "TELNET":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("TELNET");
                                        break;
                                    case "USB":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("USB");
                                        break;
                                    case "USBWIFI":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("USBWIFI");
                                        break;
                                    case "UPGRADE":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        break;
                                    case "WATCHDOG":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("WATCHDOG");
                                        break;
                                    case "WIFI":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("WIFI");
                                        break;
                                    case "3G":
                                        AddFunction(line, cmd[0], n);
                                        n = n + 1;
                                        TEST_STATUS.Add(0);
                                        TEST_FunLog.Add("3G");
                                        break;
                                    default:
                                        break;
                                }
                            }
                            // 3. Read the next line
                            line = sr.ReadLine();
                        }
                        // 4. close the file
                        sr.Close();
                    }
                    if (n == 0)
                        return;
                    else
                        TestFun_MaxIdx = n;
                    composingTmr.Enabled = true;
                    TEST_STATUS.TrimToSize();   // TrimToSize():將容量設為實際元素數目
                    TEST_FunLog.TrimToSize();
                    TEST_RESULT = new string[TEST_FunLog.Count];

                    cmdOpeFile.Text = fnameTmp.Replace(".txt", string.Empty);

                    macEnabledStatus(MODEL_NAME);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace, "");
            }
            finally
            {
                //Debug.Print("STOP_WHEN_FAIL = " + STOP_WHEN_FAIL.ToString());
                //Debug.Print("測試陣列的大小 : " + lblFun_MaxIdx.ToString());
            }
        }

        private void macEnabledStatus(string modelname)
        {
            switch (modelname.Contains("SE19"))
            {
                case true:

                    break;
                default:

                    break;
            }
        }

        /// <summary>
        /// 新增控制項 lblFunction
        /// </summary>
        /// <param name="dat">檔案名稱的設定內容，給Tag屬性</param>
        /// <param name="item_name">測試項名稱，給Text屬性</param>
        /// <param name="n">控制項陣列的索引標籤，給TabIndex屬性</param>
        public void AddFunction(string dat, string item_name, int n)
        {
            lblFunction[n] = new Label();
            lblFunction[n].AutoSize = true;
            lblFunction[n].TextAlign = ContentAlignment.MiddleCenter;
            lblFunction[n].Font = new Font("Arial", 12, FontStyle.Bold); // new Font(字型, 大小, 樣式);
            lblFunction[n].BorderStyle = BorderStyle.FixedSingle;
            lblFunction[n].Enabled = true;
            lblFunction[n].Location = new Point(12, 48);
            lblFunction[n].Visible = false;
            lblFunction[n].Tag = dat;
            lblFunction[n].BackColor = Color.FromArgb(255, 255, 255);
            lblFunction[n].Text = item_name.Substring(0, 1).ToUpper() + item_name.Substring(1, item_name.Length - 1);
            // TabIndex => ((Label)sender).TabIndex
            lblFunction[n].TabIndex = n;
            //splitContainer1.Panel1.Controls.Add(lblFunction[n]);
            tabPage5.Controls.Add(lblFunction[n]);
            // 註冊事件
            lblFunction[n].MouseMove += new MouseEventHandler(lblFunction_MouseMove);
            lblFunction[n].MouseLeave += new EventHandler(lblFunction_MouseLeave);
            lblFunction[n].MouseDown += new MouseEventHandler(lblFunction_MouseDown);

            // 連結 contextMenuStrip (右鍵選單)
            lblFunction[n].ContextMenuStrip = contextMenuStrip1;
        }

        /// <summary>
        /// 移除控制項 lblFunction
        /// </summary>
        /// <param name="MaxIdx">控制項陣列的上限值</param>
        public void RemoveControl(int MaxIdx)
        {
            int idx;
            // NOTE: The code below uses the instance of the Label from the previous example.
            for (idx = 0; idx <= MaxIdx; idx++)
            {
                //if (splitContainer1.Panel1.Controls.Contains(lblFunction[idx]))
                if (tabPage5.Controls.Contains(lblFunction[idx]))
                {
                    // 移除事件
                    this.lblFunction[idx].MouseMove -= new MouseEventHandler(lblFunction_MouseMove);
                    lblFunction[idx].MouseLeave -= new EventHandler(lblFunction_MouseLeave);
                    lblFunction[idx].MouseDown -= new MouseEventHandler(lblFunction_MouseDown);
                    splitContainer1.Panel1.Controls.Remove(lblFunction[idx]);
                    lblFunction[idx].Dispose();
                }
            }
        }

        private void lblFunction_MouseMove(object sender, MouseEventArgs e)
        {
            string dat = System.Convert.ToString(((Label)sender).Tag);
            lbl_cmdTag.Text = dat;
        }

        private void lblFunction_MouseLeave(object sender, EventArgs e)
        {
            lbl_cmdTag.Text = string.Empty;
        }

        // 單擊測試 & 右鍵選單
        private void lblFunction_MouseDown(object sender, MouseEventArgs e)
        {
            string dat = System.Convert.ToString(((Label)sender).Text);
            int idx = ((Label)sender).TabIndex;
            if (cmdStart.Enabled == false) { return; }
            if (e.Button == MouseButtons.Left)
            {
                cmdOpeFile.Enabled = false;
                cmdStart.Enabled = false;
                cmdStop.Enabled = true;
                cmdNext.Enabled = false;
                TEST_STATUS[idx] = RunTest(idx);
                cmdOpeFile.Enabled = true;
                cmdStart.Enabled = true;
                cmdStop.Enabled = true;
                cmdNext.Enabled = true;
                Run_Stop = false;
            }
            else if (e.Button == MouseButtons.Right)
            {
                MOUSE_Idx = idx;
                if (dat == "Console-DUT" || dat == "Console-EUT" || dat == "Telnet" || dat == "Power")
                {
                    用Putty開啟ToolStripMenuItem.Visible = true;
                }
                else
                {
                    用Putty開啟ToolStripMenuItem.Visible = false;
                }
            }
        }

        public void MappingFunction()
        {
            switch ((MODEL_NAME.Substring(0, 4)).ToUpper())
            {
                default:
                    COM_function = "atop_tcp_server";
                    CAN_functiom = "dcan_tcpsvr";
                    CAN_loopback = "dcan_loopback";
                    break;
            }

            txtDutIP.Text = TARGET_IP;
            //string[] ip_split = new string[3];
            //ip_split = TARGET_IP.Split('.');
            //ip_split[3] = (Convert.ToInt32(ip_split[3]) + 2).ToString();
            //txtEutIP.Text = ip_split[0] + "." + ip_split[1] + "." + ip_split[2] + "." + ip_split[3];
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            //string rxContents;
            if (serialPort1.BytesToRead > 0)
            {
                //int bytes = serialPort1.BytesToRead;
                //byte[] comBuffer = new byte[bytes];
                byte[] comBuffer = new byte[serialPort1.BytesToRead];
                serialPort1.Read(comBuffer, 0, comBuffer.Length);
                rxContents = Encoding.ASCII.GetString(comBuffer);

                //myUI(rxContents, txt_Rx);
                this.Invoke(new EventHandler(serialPort1_Display));
            }
        }

        private void serialPort1_Close()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DataReceived -= new SerialDataReceivedEventHandler(serialPort1_DataReceived);
                Hold(100);
                serialPort1.Close();
            }
        }

        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //string rxContents_EUT;
            if (serialPort2.BytesToRead > 0)
            {
                byte[] comBuffer = new byte[serialPort2.BytesToRead];
                serialPort2.Read(comBuffer, 0, comBuffer.Length);
                rxContents_EUT = Encoding.ASCII.GetString(comBuffer);

                //myUI(rxContents_EUT, txt_Rx_EUT);
                this.Invoke(new EventHandler(serialPort2_Display));
            }
        }

        private void serialPort2_Close()
        {
            if (serialPort2.IsOpen)
            {
                serialPort2.DataReceived -= new SerialDataReceivedEventHandler(serialPort2_DataReceived);
                Hold(100);
                serialPort2.Close();
            }
        }

        private void telnet_Receive()
        {
            string rdData = string.Empty;
            while (true)
            {
                try
                {
                    Array.Resize(ref bytRead_telnet, telnet.ReceiveBufferSize); // Array.Resize等於vb的ReDim
                    telentStream.Read(bytRead_telnet, 0, telnet.ReceiveBufferSize);
                    rdData = (System.Text.Encoding.Default.GetString(bytRead_telnet));
                    myUI(rdData, txt_Rx);
                    Array.Clear(bytRead_telnet, 0, telnet.ReceiveBufferSize);
                    Thread.Sleep(100);
                }
                catch (Exception)
                {
                    //throw;
                }
            }
        }

        /// <summary>
        /// 發送指令
        /// </summary>
        /// <param name="cmd">Command</param>
        public void SendCmd(string cmd)
        {
            if (serialPort1.IsOpen)
            {
                //serialPort1.DiscardOutBuffer(); // 捨棄序列驅動程式傳輸緩衝區的資料
                if (cmd.StartsWith(((char)27).ToString()))
                {
                    serialPort1.Write(cmd);
                }
                else
                {
                    serialPort1.Write(cmd);
                    serialPort1.Write(((char)13).ToString());
                }
            }
            else if (telnet != null && telnet.Connected)
            {
                if (cmd.StartsWith(((char)27).ToString()))
                {
                    bytWrite_telnet = System.Text.Encoding.Default.GetBytes(cmd);
                    telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                }
                else
                {
                    bytWrite_telnet = System.Text.Encoding.Default.GetBytes(cmd + ((char)13).ToString());
                    telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                }
            }
        }

        private void cmdStart_Click(object sender, EventArgs e)
        {
            int idx;
            cmdOpeFile.Enabled = false;
            cmdStart.Enabled = false;
            cmdStop.Enabled = true;
            cmdNext.Enabled = false;
            Run_Stop = false;
            try
            {
                tabControl2.SelectedTab = tabPage5;
                Hold(10);
                time = DateTime.Now;
                startTime = String.Format("{0:00}/{1:00}" + ((char)10).ToString() + "{2:00}:{3:00}:{4:00}", time.Month, time.Day, time.Hour, time.Minute, time.Second);
                if (!chooseStart)
                {
                    for (idx = 0; idx < TestFun_MaxIdx; idx++)
                    {
                        if (!lblFunction[idx].Text.ToUpper().Contains("CONSOLE") || !lblFunction[idx].Text.ToUpper().Contains("TELNET"))
                        {
                            lblFunction[idx].BackColor = Color.FromArgb(255, 255, 255);
                        }
                    }
                    Hold(1);
                }
                retest:
                for (idx = Test_Idx; idx < TestFun_MaxIdx; idx++)
                {
                    TEST_STATUS[idx] = RunTest(idx);
                    if (Run_Stop)
                    {
                        return;
                    }
                    if (STOP_WHEN_FAIL && Convert.ToInt32(TEST_STATUS[idx]) == 2)
                    {
                        break;
                    }
                    Hold(1000);
                }
                if (chkLoop.CheckState == CheckState.Checked && Run_Stop == false)
                {
                    for (idx = 0; idx < TestFun_MaxIdx; idx++)
                    {
                        if (!lblFunction[idx].Text.ToUpper().Contains("CONSOLE") || !lblFunction[idx].Text.ToUpper().Contains("TELNET"))
                        {
                            lblFunction[idx].BackColor = Color.FromArgb(255, 255, 255);
                            Hold(1);
                        }
                    }
                    Test_Idx = 0;
                    goto retest;
                }
            }
            finally
            {
                cmdOpeFile.Enabled = true;
                cmdStart.Enabled = true;
                cmdStop.Enabled = true;
                cmdNext.Enabled = true;
                if (telnet.Connected) { telnet.Close(); }
                Test_Idx = 0;
                chooseStart = false;
            }
        }

        private void cmdStop_Click(object sender, EventArgs e)
        {
            try
            {
                Run_Stop = true;
                WAIT = false;
                SendCmd(((char)3).ToString()); // ((char)3):Ctrl+c
                Shell(appPATH, "arp-d.bat");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Stop error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                cmdOpeFile.Enabled = true;
                cmdStart.Enabled = true;
                cmdStop.Enabled = true;
                cmdNext.Enabled = true;
            }
        }

        /// <summary>
        /// 功能測試
        /// </summary>
        /// <param name="idx">控制項(lblFunction)陣列的索引標籤</param>
        /// <returns>回傳測試完的結果 0:未測試, 1:PASS, 2:fail, 3:error </returns>
        public int RunTest(int idx)
        {
            lblStatus.Text = string.Empty;
            int RunTest_result = 0; // 0:未測試,1:PASS,2:fail,3:error
            try
            {
                int i;
                string[] line;
                DialogResult dr;
                string[] cmd;
                int j;
                FileStream fs;
                StreamWriter sw;
                double duration;
                int secs;
                string fileDirectory;
                string filePath;
                long delay;
                string keyString = string.Empty;
                string ip_3G = string.Empty; // ip_3G used in "3G" and "NAT" test function.
                //telnet = new TcpClient();

                lblFunction[idx].BackColor = Color.FromArgb(0, 255, 255);   // 測試中

                cmd = Convert.ToString(lblFunction[idx].Tag).Split(' ');
                if (cmd[0].ToUpper() != "CONSOLE-DUT" & cmd[0].ToUpper() != "CONSOLE-EUT" & cmd[0].ToUpper() != "TELNET" & cmd[0].ToUpper() != "POWER")
                {
                    if (!serialPort1.IsOpen & !telnet.Connected)
                    {
                        lblStatus.Text = "Console-DUT 或 Telnet 未連接";
                        return RunTest_result = 3;
                    }
                    //if ((cmd[0].ToUpper() == "COMTOCOM" || cmd[0].ToUpper() == "CANTOCAN")
                    //    & !serialPort2.IsOpen)
                    //{
                    //    lblStatus.Text = "Console-EUT 未連接";
                    //    return RunTest_result = 3;
                    //}
                }
                // for excel log
                if (TEST_FunLog.Contains(cmd[0].ToUpper()))
                {
                    idx_funlog = TEST_FunLog.IndexOf(cmd[0].ToUpper());
                }
                //SendCmd(string.Empty);
                switch (cmd[0].ToUpper())
                {
                    case "CONSOLE-DUT":     // Console show
                        serialPort1_Close();
                        //if (serialPort1.IsOpen) { serialPort1.Close(); }
                        serialPort1.PortName = cmbDutCom.Text;
                        serialPort1.BaudRate = 115200;
                        serialPort1.Parity = Parity.None;
                        serialPort1.DataBits = 8;
                        serialPort1.StopBits = StopBits.One;
                        serialPort1.Handshake = Handshake.None; // 流量控制；交握協定
                        serialPort1.Open();
                        serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);
                        lblStatus.Text = "Console-DUT Connect OK !";
                        RunTest_result = 1;
                        if (cmd.GetUpperBound(0) >= 1)
                        {
                            if (cmd[1].ToUpper() == "SHOW")
                            {
                                consoleToolStripMenuItem.Checked = true;
                            }
                        }
                        else { consoleToolStripMenuItem.Checked = false; }

                        serialPort1.Write(((char)13).ToString());
                        Hold(500);
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("login"))
                            {
                                serialPort1.Write(USR + ((char)13).ToString());
                                Hold(200);
                                serialPort1.Write(PWD + ((char)13).ToString());
                                break;  // for
                            }
                            else if (line[i].Contains("Main Menu") || line[i].Contains("Manufactory Settings"))
                            {
                                break;  // for
                            }
                        }

                        SendCmd("");
                        break;
                    case "CONSOLE-EUT":     // Console
                        serialPort2_Close();
                        //if (serialPort2.IsOpen) { serialPort2.Close(); }
                        serialPort2.PortName = cmbEutCom.Text;
                        serialPort2.BaudRate = 115200;
                        serialPort2.Parity = Parity.None;
                        serialPort2.DataBits = 8;
                        serialPort2.StopBits = StopBits.One;
                        serialPort2.Handshake = Handshake.None; // 流量控制；交握協定
                        serialPort2.Open();
                        serialPort2.DataReceived += new SerialDataReceivedEventHandler(serialPort2_DataReceived);
                        lblStatus.Text = "Console-EUT Connect OK !";
                        RunTest_result = 1;
                        //serialPort2.Write("root" + ((char)13).ToString());
                        //Hold(200);
                        //serialPort2.Write("root" + ((char)13).ToString());
                        break;
                    case "TELNET":      // Telnet USR PWD
                        //Shell(appPATH, "arp-d.bat");
                        //Hold(1000);
                        //txt_Rx.Text = string.Empty;
                        RunTest_result = 1;
                        //if (telnet.Connected) { telnet.Close(); }
                        if (objping.Send(TARGET_IP, 1000).Status == System.Net.NetworkInformation.IPStatus.Success)
                        {
                            if (!telnet.Connected)
                            {
                                telnet = new TcpClient();
                                telnet.Connect(TARGET_IP, 23);   // 連接23端口 (Telnet的默認端口)
                                telentStream = telnet.GetStream();  // 建立網路資料流，將字串寫入串流

                                if (telnet.Connected)
                                {
                                    //lblStatus.Text = "連線成功，正在登錄...";
                                    lblStatus.Text = "正在登錄...";
                                    Hold(1000);
                                    // 背景telnet接收執行緒
                                    if (rcvThread == null || !rcvThread.IsAlive)
                                    {
                                        ThreadStart backgroundReceive = new ThreadStart(telnet_Receive);
                                        rcvThread = new Thread(backgroundReceive);
                                        rcvThread.IsBackground = true;
                                        rcvThread.Start();
                                    }
                                    bytWrite_telnet = System.Text.Encoding.Default.GetBytes(USR + ((char)13).ToString());
                                    telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                                    Hold(200);
                                    bytWrite_telnet = System.Text.Encoding.Default.GetBytes(PWD + ((char)13).ToString());
                                    telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                                    lblStatus.Text = "連線成功 ";
                                }
                            }
                        }
                        else
                        {
                            lblStatus.Text = "ping失敗，請確認你的IP設定或網路設定";
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            RunTest_result = 2;
                        }
                        break;
                    case "RESTART":
                        if (telnet.Connected) { telnet.Close(); }
                        RunTest_result = 3;
                        SendCmd("restart&");
                        SendCmd("atop_restart&");
                        RecNote(idx, "Restart");
                        if (cmd[1].ToUpper() == "LOGIN")
                        {
                            secs = Convert.ToInt32(cmd[2]);
                            RunTest_result = ReCntTelnet(secs);   // need to login
                        }
                        else if (cmd[1].ToUpper() == "NONE")
                        {
                            WaitKey = "U-Boot ";
                            if (Hold(5000))
                            {
                                secs = Convert.ToInt32(cmd[2]) * 1000;
                                WaitKey = "~#";     // doesn't need to login
                                if (Hold(secs))
                                {
                                    RunTest_result = 1;
                                }
                            }
                        }
                        break;
                    case "COM": // normal test: COM max_port mode
                        RunTest_result = 1;
                        // SE5901
                        if (MODEL_NAME.Equals("SE5901", StringComparison.CurrentCultureIgnoreCase))
                        {
                            SendCmd("/jffs2/" + cmd[2].ToLower() + "_loopback");
                            SendCmd("");// 讓WaitKey較能接收到資料做判斷
                            WaitKey = "PASS";
                            if (Hold(3000) == false)
                            {
                                RunTest_result = 2;
                                SendCmd(((char)3).ToString());
                                RecNote(idx, "RS232 loopback test Fail.");
                            }
                        }
                        // SE5901A
                        else if (MODEL_NAME.Contains("SE5901A"))
                        {
                            j = 0;
                            SendCmd("/jffs2/" + cmd[2].ToLower() + "_loopback " + j + " " + (j + 4));
                            SendCmd("");// 讓WaitKey較能接收到資料做判斷
                            WaitKey = "PASS";
                            if (Hold(3000) == false)
                            {
                                RunTest_result = 2;
                                SendCmd(((char)3).ToString());
                                string failmessage = "COM" + j + " -> COM" + (j + 4) + " Fail";
                                RecNote(idx, failmessage);
                            }
                            SendCmd("/jffs2/" + cmd[2].ToLower() + "_loopback " + (j + 4) + " " + j);
                            SendCmd("");
                            WaitKey = "PASS";
                            if (Hold(3000) == false)
                            {
                                RunTest_result = 2;
                                SendCmd(((char)3).ToString());
                                string failmessage = "COM" + (j + 4) + " -> COM" + j + " Fail";
                                RecNote(idx, failmessage);
                            }
                        }
                        // normal test
                        else
                        {
                            for (j = 1; j <= Convert.ToInt32(cmd[1]); j = j + 2)
                            {
                                SendCmd(cmd[2].ToLower() + "_loopback " + j + " " + (j + 1));
                                WaitKey = "test ok";
                                if (Hold(3000) == false)
                                {
                                    RunTest_result = 2;
                                    SendCmd(((char)3).ToString());
                                    string failmessage = "COM" + j + " -> COM" + (j + 1) + " Fail";
                                    RecNote(idx, failmessage);
                                }
                                SendCmd(cmd[2].ToLower() + "_loopback " + (j + 1) + " " + j);
                                WaitKey = "test ok";
                                if (Hold(3000) == false)
                                {
                                    RunTest_result = 2;
                                    SendCmd(((char)3).ToString());
                                    string failmessage = "COM" + (j + 1) + " -> COM" + j + " Fail";
                                    RecNote(idx, failmessage);
                                }
                            }
                        }

                        if (RunTest_result == 1)
                        {
                            lblStatus.Text = "COM loopback Test Pass.";
                        }
                        break;
                    case "COMTOCOM":    // COMtoCOM port(1-4 or 4 or 0-4) 陪測物IP mode BaudRate time unit
                        //serialPort2.Write("killall atop_tcp_server" + ((char)13).ToString());
                        //Hold(100);
                        //serialPort2.Write("atop_tcp_server " + cmd[3] + " " + cmd[4] + "&" + ((char)13).ToString());
                        //Hold(100);
                        SendCmd("killall tcp_server");
                        Hold(100);
                        SendCmd("/jffs2/tcp_server " + cmd[3] + "&");
                        Hold(100);
                        RunTest_result = 2;
                        // 建立檔案
                        fs = File.Open("Auto_Test", FileMode.OpenOrCreate, FileAccess.Write);
                        // 建構StreamWriter物件
                        sw = new StreamWriter(fs);
                        sw.Close();
                        fs.Close();
                        duration = Math.Round(TimeUnit(idx, 5) / 60, 2);
                        MultiPortTesting_settings(cmd[2], 1000, cmd[1], 4660, 1, duration.ToString());
                        COM_PID[0] = Shell(appPATH, "Multi-Port-Testingv1.6r.exe");
                        Hold(1000);
                        MultiPortTesting_settings(TARGET_IP, 1000, cmd[1], 4660, 0, duration.ToString());
                        COM_PID[1] = Shell(appPATH, "Multi-Port-Testingv1.6r.exe");
                        pause(duration);
                        if (File.Exists("Auto_Test"))
                        {
                            File.Delete("Auto_Test");
                        }
                        Hold(3000); // 因為Multi-Port-Testingv1.6p (2013/11/22)的行為，所以等待是必須的
                        CloseShell(COM_PID[0]);
                        CloseShell(COM_PID[1]);
                        COM_PID[0] = 0;
                        COM_PID[1] = 0;
                        if (!File.Exists("debug.txt"))
                        {
                            RunTest_result = 1;
                        }

                        //serialPort2.Write("killall atop_tcp_server" + ((char)13).ToString());
                        //Hold(100);
                        SendCmd("killall tcp_server");
                        Hold(100);

                        // 把關沒有產生debug.txt文件的其他error
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("Terminated")) { break; }
                            if (line[i].Contains("error") || line[i].Contains("No such file or directory") || line[i].Contains("not found"))
                            {
                                RunTest_result = 2;
                                break;  // for
                            }
                        }

                        if (RunTest_result == 1)
                        {
                            lblStatus.Text = cmd[0].ToUpper() + " Test Pass.";
                        }
                        else
                        {
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                        }
                        break;
                    case "CANTOCAN": // CAN  port(1-2 or 2 or 0-2) 陪測物IP 125 time unit
                        serialPort2.Write("killall " + CAN_functiom + ((char)13).ToString());
                        Hold(100);
                        serialPort2.Write(CAN_functiom + " " + cmd[3] + " &" + ((char)13).ToString());
                        Hold(100);
                        SendCmd("killall " + CAN_functiom);
                        Hold(100);
                        SendCmd(CAN_functiom + " " + cmd[3] + " &");
                        Hold(100);
                        RunTest_result = 2;
                        // 建立檔案
                        fs = File.Open("Auto_Test", FileMode.OpenOrCreate, FileAccess.Write);
                        // 建構StreamWriter物件
                        sw = new StreamWriter(fs);
                        sw.Close();
                        fs.Close();
                        duration = Math.Round(TimeUnit(idx, 4) / 60, 2);
                        MultiPortTesting_settings(cmd[2], 1000, cmd[1], 8000, 1, duration.ToString());
                        COM_PID[0] = Shell(appPATH, "Multi-Port-Testingv1.6r.exe");
                        Hold(500);
                        MultiPortTesting_settings(TARGET_IP, 1000, cmd[1], 8000, 0, duration.ToString());
                        COM_PID[1] = Shell(appPATH, "Multi-Port-Testingv1.6r.exe");
                        pause(duration);
                        if (File.Exists("Auto_Test"))
                        {
                            File.Delete("Auto_Test");
                        }
                        Hold(3000); // 因為Multi-Port-Testingv1.6p (2013/11/22)的行為，所以等待是必須的
                        CloseShell(COM_PID[0]);
                        CloseShell(COM_PID[1]);
                        COM_PID[0] = 0;
                        COM_PID[1] = 0;
                        if (!File.Exists("debug.txt"))
                        {
                            RunTest_result = 1;
                        }

                        serialPort2.Write("killall " + CAN_functiom + ((char)13).ToString());
                        Hold(100);
                        SendCmd("killall " + CAN_functiom);
                        Hold(100);

                        // 把關沒有產生debug.txt文件的其他error
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("Terminated")) { break; }
                            if (line[i].Contains("error") || line[i].Contains("No such file or directory") || line[i].Contains("not found"))
                            {
                                RunTest_result = 2;
                                break;  // for
                            }
                        }

                        if (RunTest_result == 1)
                        {
                            lblStatus.Text = cmd[0].ToUpper() + " Test Pass.";
                        }
                        else
                        {
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                        }
                        break;
                    case "LOADTOOLS":   // "MODEL_NAME"_Tools資料夾裡的所有檔案載入待測物
                        RunTest_result = 1;

                        if (MODEL_NAME.Contains("SE5901A"))
                        {
                            fileDirectory = "SE5901A_Tools";
                        }
                        else
                        {
                            fileDirectory = MODEL_NAME + "_Tools";
                        }

                        filePath = appPATH + "\\" + fileDirectory;
                        if (Directory.Exists(fileDirectory))
                        {
                            // Process the list of files found in the directory.
                            string[] fileEntries = Directory.GetFiles(filePath);
                            foreach (string fileName in fileEntries)
                            {
                                string sourceFile = fileName.Replace(filePath + "\\", "");
                                uploadFile(TARGET_IP, fileDirectory + "\\" + sourceFile, USR, PWD);
                                Hold(1);
                                bool check = checkFile(TARGET_IP, sourceFile, USR, PWD);
                                if (!check) // false代表沒有上載成功，檔案不存在
                                {
                                    RecNote(idx, sourceFile + " not exist!");
                                    RunTest_result = 2;
                                }
                            }
                        }
                        SendCmd("chmod 755 /jffs2/*");
                        Hold(100);

                        SendCmd("ls -al /jffs2/");
                        break;
                    case "FACTORYFILES":
                        RunTest_result = 1;

                        fileDirectory = MODEL_NAME + "_factoryfiles";

                        filePath = appPATH + "\\" + fileDirectory;
                        if (Directory.Exists(fileDirectory))
                        {
                            // Process the list of files found in the directory.
                            string[] fileEntries = Directory.GetFiles(filePath);
                            foreach (string fileName in fileEntries)
                            {
                                string sourceFile = fileName.Replace(filePath + "\\", "");
                                uploadFile(TARGET_IP, fileDirectory + "\\" + sourceFile, USR, PWD);
                                Hold(1);
                                bool check = checkFile(TARGET_IP, sourceFile, USR, PWD);
                                if (!check) // false代表沒有上載成功，檔案不存在
                                {
                                    RecNote(idx, sourceFile + " not exist!");
                                    RunTest_result = 2;
                                }
                            }
                        }
                        SendCmd("chmod 755 /jffs2/*");
                        Hold(100);
                        if (MODEL_NAME.Contains("SE19"))
                        {
                            SendCmd("mv /jffs2/atop_tcp_server485se190X131022 /jffs2/tcp_server485");
                            Hold(100);
                            SendCmd("mv /jffs2/atop_tcp_server232se190X131022 /jffs2/tcp_server232");
                            Hold(100);
                            //SendCmd("mv /jffs2/tcp_server422_powerpc /jffs2/tcp_server422");
                            //Hold(100);
                        }
                        SendCmd("ls -al /jffs2/");
                        break;
                    case "DELETE":
                        SendCmd("rm /jffs2/*");
                        RunTest_result = 1;

                        if (MODEL_NAME.Contains("SE5901A"))
                        {
                            fileDirectory = "SE5901A_Tools";
                        }
                        else
                        {
                            fileDirectory = MODEL_NAME + "_Tools";
                        }

                        filePath = appPATH + "\\" + fileDirectory;
                        if (Directory.Exists(fileDirectory))
                        {
                            // Process the list of files found in the directory.
                            string[] fileEntries = Directory.GetFiles(filePath);
                            foreach (string fileName in fileEntries)
                            {
                                string sourceFile = fileName.Replace(filePath + "\\", "");
                                bool check = checkFile(TARGET_IP, sourceFile, USR, PWD);
                                if (check)  // true代表沒有刪除成功
                                {
                                    RecNote(idx, sourceFile + " 沒有刪除成功!");
                                    RunTest_result = 2;
                                }
                            }
                        }
                        SendCmd("ls /jffs2/");
                        break;
                    case "NAT":
                        string[] nat = ip_3G.Split('.');
                        if (nat[0] == "111" || nat[0] == "42")
                        {
                            SendCmd("atop_nat");
                            // tcp connect
                        }
                        else
                        {
                            lblStatus.Text = "略過測試";
                            RecNote(idx, cmd[0].ToUpper() + " 略過測試");
                            RunTest_result = 1;
                        }
                        break;
                    case "LAN": // Lan dhcp IP NETMASK GATEWAY dns1 dns2 1 up/down
                        string[] lanInfo;
                        string lanIP, lanMask;
                        RunTest_result = 2;
                        // lan dhcp to eeprom, data - 1:off 0:on
                        SendCmd("/jffs2/set_lan 1 " + cmd[1]);
                        Hold(3000);
                        // lan IP to eeprom, data - IP address
                        SendCmd("/jffs2/set_lan 2 " + cmd[2]);
                        Hold(3000);
                        // lan NETMASK to eeprom, data - netmask address
                        SendCmd("/jffs2/set_lan 3 " + cmd[3]);
                        Hold(3000);
                        // lan GATEWAY to eeprom, data - gateway address
                        SendCmd("/jffs2/set_lan 4 " + cmd[4]);
                        Hold(3000);
                        // lan dns to eeprom, data - dns1 address, dns2 address
                        SendCmd("/jffs2/set_lan 5 " + cmd[5] + ", " + cmd[6]);
                        Hold(3000);
                        // lan is default GW to eeprom, data - 1(fixed)
                        SendCmd("/jffs2/set_lan 6 1");
                        Hold(3000);
                        // lan interface to system, data - up / down
                        SendCmd("/jffs2/set_lan 7 " + cmd[7]);
                        Hold(3000);
                        SendCmd("");
                        SendCmd("ifconfig eth0");
                        Hold(300);
                        keyString = GetLine("inet addr:", "ifconfig eth0").Trim();
                        lanInfo = keyString.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                        lanIP = lanInfo[0].Substring(10);
                        lanMask = lanInfo[2].Substring(5);
                        if (lanIP == cmd[2] && lanMask == cmd[3])
                        {
                            lblStatus.Text = "LAN IP:" + lanIP + "  Mask:" + lanMask;
                            RunTest_result = 1;
                        }
                        break;
                    case "USBWIFI": // Usbwifi SSID AuthenMode Encrytype WepDefaultKeyId 1,h,9999999999 2,h,8888888888 3,h,7777777777 4,h,6666666666 WPAKey
                        // SSID
                        SendCmd("/jffs2/atop_set_usbwifi 1 " + cmd[1]);
                        Hold(3000);
                        // Authen mode, data - 0:open,1:wep,2:wpa-psk,3:wpa2-psk
                        SendCmd("/jffs2/atop_set_usbwifi 2 " + cmd[2]);
                        Hold(3000);
                        // 由Authen mode:cmd[2]來決定cmd[3]的指令
                        switch (cmd[2])
                        {
                            case "0": // open(none)
                                // Encrytype, data - 0:none
                                SendCmd("/jffs2/atop_set_usbwifi 3 0");
                                Hold(3000);
                                break;
                            case "1": // wep
                                // Encrytype, data - 1:wep
                                SendCmd("/jffs2/atop_set_usbwifi 3 1");
                                Hold(3000);
                                // Wep default key id,data - 1-4
                                SendCmd("/jffs2/atop_set_usbwifi 4 " + cmd[4]);
                                Hold(3000);
                                // Wep key1 type & key1 string, data - <type> 0:hex,1:ascii <string> key string
                                SendCmd("/jffs2/atop_set_usbwifi 5 0 9999999999");
                                Hold(3000);
                                // Wep key2 type & key2 string, data - <type> 0:hex,1:ascii <string> key string
                                SendCmd("/jffs2/atop_set_usbwifi 6 0 8888888888");
                                Hold(3000);
                                // Wep key3 type & key3 string, data - <type> 0:hex,1:ascii <string> key string
                                SendCmd("/jffs2/atop_set_usbwifi 7 0 7777777777");
                                Hold(3000);
                                // Wep key4 type & key4 string, data - <type> 0:hex,1:ascii <string> key string
                                SendCmd("/jffs2/atop_set_usbwifi 8 0 6666666666");
                                Hold(3000);
                                break;
                            case "2": // wpa-psk
                                // Encrytype, data - 2:tkip,3:aes
                                SendCmd("/jffs2/atop_set_usbwifi 3 " + cmd[3]);
                                Hold(3000);
                                // Wep default key id,data - 1-4
                                SendCmd("/jffs2/atop_set_usbwifi 4 1");
                                Hold(3000);
                                // wpa key, data - wpa key string
                                SendCmd("/jffs2/atop_set_usbwifi 9 " + cmd[5]);
                                Hold(3000);
                                break;
                            case "3": // wpa2-psk
                                // Encrytype, data - 2:tkip,3:aes
                                SendCmd("/jffs2/atop_set_usbwifi 3 " + cmd[3]);
                                Hold(3000);
                                // Wep default key id,data - 1-4
                                SendCmd("/jffs2/atop_set_usbwifi 4 1");
                                Hold(3000);
                                // wpa key, data - wpa key string
                                SendCmd("/jffs2/atop_set_usbwifi 9 " + cmd[5]);
                                Hold(3000);
                                break;
                        }
                        // Create new wpa conf file
                        SendCmd("/jffs2/atop_set_usbwifi 10 1");
                        Hold(5000);
                        SendCmd("");
                        SendCmd("iwconfig wlan0");
                        Hold(1000);
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("SSID:\"" + cmd[1] + "\""))
                            {
                                RunTest_result = 1;
                                lblStatus.Text = cmd[0].ToUpper() + " have connect AP";
                                break;  // for
                            }
                            if (line[i].Contains("unassociated") || line[i].Contains("iwconfig wlan0"))
                            {
                                RunTest_result = 2;
                                lblStatus.Text = cmd[0].ToUpper() + " haven't connect AP";
                                RecNote(idx, cmd[0].ToUpper() + " haven't connect AP");
                                break;  // for
                            }
                        }
                        break;
                    case "WIFI":
                        // wifi dhcp to eeprom, data - 1:off 0:on
                        SendCmd("/jffs2/set_wifi 1 " + cmd[1]);
                        Hold(3000);
                        // 由dhcp:cmd[1]來決定
                        switch (cmd[1])
                        {
                            case "1": // dhcp off
                                // wifi IP to eeprom, data - IP address
                                SendCmd("/jffs2/set_wifi 2 " + cmd[2]);
                                Hold(3000);
                                // wifi NETMASK to eeprom, data - netmask address
                                SendCmd("/jffs2/set_wifi 3 " + cmd[3]);
                                Hold(3000);
                                // wifi GATEWAY to eeprom, data - gateway address
                                SendCmd("/jffs2/set_wifi 4 " + cmd[4]);
                                Hold(3000);
                                // wifi dns to eeprom, data - dns1 address, dns2 address
                                SendCmd("/jffs2/set_wifi 5 " + cmd[5] + ", " + cmd[6]);
                                Hold(3000);
                                // wifi is default GW to eeprom,data - 0(fixed)
                                SendCmd("/jffs2/set_wifi 6 0");
                                Hold(3000);
                                // wifi interface to system, data - up(set wifi configure) / down(close the wifi interface)
                                SendCmd("/jffs2/set_wifi 7 " + cmd[7]);
                                break;
                            case "0":
                                // wifi dns to eeprom, data - dns1 address, dns2 address
                                SendCmd("/jffs2/set_wifi 5 " + cmd[2] + ", " + cmd[3]);
                                Hold(3000);
                                // wifi is default GW to eeprom,data - 0(fixed)
                                SendCmd("/jffs2/set_wifi 6 0");
                                Hold(3000);
                                // wifi interface to system, data - up(set wifi configure) / down(close the wifi interface)
                                SendCmd("/jffs2/set_wifi 7 " + cmd[4]);
                                break;
                        }
                        Hold(25000);
                        SendCmd("");
                        SendCmd("ifconfig wlan0");
                        RunTest_result = 1;
                        break;
                    case "3G": // 3G bootup pincode pincodeEnable reconnect
                        RunTest_result = 2;

                        // Test
                        //    if (chkHumanSkip.CheckState == CheckState.Checked)
                        //    {
                        //        lblStatus.Text = "略過人工判斷";
                        //        Hold(2000); // wait for buzzer
                        //        RunTest_result = 1;
                        //    }
                        //    else
                        //    {
                        //        lblStatus.Text = "人工判斷";
                        //        Hold(1000);
                        //        dr = MessageBox.Show("運行燈是否一秒內快速連續閃爍 ? ", cmd[0] + " Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                        //MessageBoxDefaultButton.Button1);
                        //        if (dr == DialogResult.Yes) { RunTest_result = 1; }
                        //        else if (dr == DialogResult.No) { RecNote(idx, cmd[0].ToUpper() + " Test Fail."); RunTest_result = 2; }
                        //    }

                        // Setting
                        if (cmd[5] == "1")
                        {
                            // dial on bootup, data - 1:on 0:off
                            SendCmd("/jffs2/atop_set_threeg 1 " + cmd[1]);
                            Hold(3000);
                            //keyString = GetLine("dial on boot", "atop_set_threeg");
                            //if (keyString.Contains("[" + cmd[1] + "]"))
                            //{
                            //}
                            // pin code, data - pin code string
                            SendCmd("/jffs2/atop_set_threeg 2 " + cmd[2]);
                            Hold(3000);
                            // pin code enable, data - 1:on 0:off
                            SendCmd("/jffs2/atop_set_threeg 3 " + cmd[3]);
                            Hold(3000);
                            // reconnect enable, data - 1:on 0:off
                            SendCmd("/jffs2/atop_set_threeg 4 " + cmd[4]);
                            Hold(3000);
                            // 3G connect, data - 1
                            SendCmd("/jffs2/atop_set_threeg 5 1");
                            Hold(21000);
                            SendCmd("ifconfig ppp0");
                            Hold(300);
                            keyString = GetLine("inet addr:", "ifconfig ppp0").Trim();
                            ip_3G = keyString.Substring(10, 15).Trim();
                            lblStatus.Text = "3G IP: " + ip_3G;
                            RunTest_result = 1;
                        }
                        else if (cmd[5] == "0")
                        {
                            // 3G disconnect, data - 1
                            SendCmd("/jffs2/atop_set_threeg 6 1");
                            Hold(100);
                            WaitKey = "Connection terminated";
                            if (Hold(5000))
                            {
                                lblStatus.Text = "3G Connection terminated.";
                                RunTest_result = 1;
                            }
                        }
                        break;
                    case "DI":

                        break;
                    case "DO":
                        RunTest_result = 2;
                        if (cmd.GetUpperBound(0) >= 2)
                        {
                            if (cmd[2].ToUpper() == "SKIP")
                            {
                                lblStatus.Text = "略過人工判斷";
                                SendCmd("atop_do " + cmd[1] + " 1");
                                Hold(1000);
                                SendCmd("atop_do " + cmd[1] + " 0");
                                Hold(1000);
                                SendCmd("atop_do " + cmd[1] + " 1");
                                Hold(1000);
                                SendCmd("atop_do " + cmd[1] + " 0");
                                RunTest_result = 1;
                            }
                        }
                        else if (chkHumanSkip.CheckState == CheckState.Checked)
                        {
                            lblStatus.Text = "略過人工判斷";
                            SendCmd("atop_do " + cmd[1] + " 1");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 0");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 1");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 0");
                            RunTest_result = 1;
                        }
                        else
                        {
                            lblStatus.Text = "人工判斷";
                            SendCmd("atop_do " + cmd[1] + " 1");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 0");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 1");
                            Hold(1000);
                            SendCmd("atop_do " + cmd[1] + " 0");
                            dr = MessageBox.Show(" Relay 燈號是否正常 ? ", cmd[0] + " Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                            if (dr == DialogResult.Yes) { RunTest_result = 1; }
                            else if (dr == DialogResult.No) { RecNote(idx, cmd[0].ToUpper() + " Test Fail."); RunTest_result = 2; }
                        }
                        break;
                    case "DOTODI":  // dio number
                        string di1_1 = string.Empty;
                        string di1_2 = string.Empty;
                        string di2_1 = string.Empty;
                        string di2_2 = string.Empty;
                        // first status
                        SendCmd("/jffs2/atop_do 1 1 1 1");
                        Hold(300);
                        SendCmd("/jffs2/atop_di");
                        Hold(200);
                        di1_1 = GetLine("DI1", "atop_di");
                        di2_1 = GetLine("DI2", "atop_di");
                        // second status
                        SendCmd("/jffs2/atop_do 1 0 1 0");
                        Hold(300);
                        SendCmd("/jffs2/atop_di");
                        Hold(200);
                        di1_2 = GetLine("DI1", "atop_di");
                        di2_2 = GetLine("DI2", "atop_di");

                        di1_1 = di1_1.Replace("DI1 state is pull-", string.Empty);
                        di1_2 = di1_2.Replace("DI1 state is pull-", string.Empty);
                        di2_1 = di2_1.Replace("DI2 state is pull-", string.Empty);
                        di2_2 = di2_2.Replace("DI2 state is pull-", string.Empty);

                        // 判斷DI狀態有改變就pass
                        lblStatus.Text = "DI1:" + di1_1 + "->" + di1_2 + " , DI2:" + di2_1 + "->" + di2_2;
                        if (!di1_1.Equals(di1_2) && !di2_1.Equals(di2_2))
                        {
                            //Debug.Print((!di1_1.Equals(di1_2)).ToString() + "," + (!di2_1.Equals(di2_2)).ToString());
                            RunTest_result = 1;
                        }
                        else
                        {
                            RecNote(idx, "DI1:" + di1_1 + "->" + di1_2 + " , DI2:" + di2_1 + "->" + di2_2);
                            RunTest_result = 2;
                        }
                        break;
                    case "KEYPAD":

                        break;
                    case "EEPROM":
                        txt_Rx.Text = string.Empty;
                        SendCmd("/jffs2/atop_eeprom_util EEPROM 0 256");
                        Hold(1000);
                        if (txt_Rx.Text.Length >= 512 && txt_Rx.Text.Contains("ATOP"))
                        {
                            RunTest_result = 1;
                        }
                        else
                        {
                            RunTest_result = 2;
                        }
                        break;
                    case "RTC":
                        //time = DateTime.Now;
                        //time1 = String.Format("{0:00}/{1:00}/{2:00}-{3:00}:{4:00}:{5:00}", time.Year, time.Month, time.Day, time.Hour, time.Minute, time.Second);
                        //timeTmp = time1.Substring(0, time1.IndexOf("-"));
                        SendCmd("set_rtc " + "2000/09/09-03:15:00"); // 為了與NTP搭配測試，所以不設定正確時間!
                        Hold(100);
                        SendCmd("get_rtc");
                        Hold(300);
                        RunTest_result = 3;
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("get_rtc"))
                            {
                                if (line[i + 1].Contains("2000/09/09-03:15"))
                                {
                                    RunTest_result = 1;
                                    lblStatus.Text = cmd[0].ToUpper() + " Test Pass !";
                                }
                                else
                                {
                                    RunTest_result = 2;
                                    lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                                    RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                                }
                                break;  // for
                            }
                        }
                        break;
                    case "NTP":
                        SendCmd("/jffs2/atop_saving_time");
                        Hold(200);

                        /* NTP（Network Time Protocol) 伺服器
                         * tock.stdtime.gov.tw
                         * watch.stdtime.gov.tw
                         * time.stdtime.gov.tw
                         * clock.stdtime.gov.tw
                         * tick.stdtime.gov.tw
                        */

                        SendCmd("/jffs2/atop_ntpd 1 time.stdtime.gov.tw 48");
                        Hold(5000);
                        SendCmd("date");
                        Hold(200);
                        // 格式=Wed Aug 27 13:59:57 UTC 2014
                        string dateTime = GetLine("UTC", "date");
                        string[] NTPdate = dateTime.Split(' ');
                        NTPdate = Array.FindAll(NTPdate, isNotStringEmpty).ToArray();   // 檢查並刪除陣列中的空字串
                        string[] NTPtime = NTPdate[3].Split(':');
                        // DateTime顯示英文月份縮寫 ,System.Globalization.CultureInfo.InvariantCulture
                        time = DateTime.Now;
                        string[] localdate = (time.ToString("ddd MMM dd HH:mm:ss UTC yyyy", System.Globalization.CultureInfo.InvariantCulture)).Split(' ');
                        string[] localtime = localdate[3].Split(':');
                        // 比較時間(不比秒數)
                        DateTime t1 = Convert.ToDateTime(NTPdate[5] + "/" + NTPdate[1] + "/" + NTPdate[2] + " " + NTPtime[0] + ":" + "00:00");
                        DateTime t2 = Convert.ToDateTime(localdate[5] + "/" + localdate[1] + "/" + localdate[2] + " " + localtime[0] + ":" + "00:00");
                        if (DateTime.Compare(t1, t2) == 0) // 小於零 t1早於t2, Zero t1與t2相同, 大於零 t1晚於t2
                        {
                            if (Math.Abs(Convert.ToInt32(NTPtime[1]) - Convert.ToInt32(localtime[1])) <= 1) // 比較分鐘，一分以內之差
                            {
                                RunTest_result = 1;
                                lblStatus.Text = cmd[0].ToUpper() + " Test Pass !";
                            }
                            else
                            {
                                RunTest_result = 2;
                                lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                                RecNote(idx, cmd[0].ToUpper() + " Test Fail. (" + dateTime + ")");
                            }
                        }
                        else
                        {
                            RunTest_result = 2;
                            lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail. (" + dateTime + ")");
                        }
                        break;
                    case "GPS":
                        if (chkHumanSkip.CheckState == CheckState.Unchecked)
                        {
                        }

                        SendCmd("atop_get_gpstime");
                        Hold(300);
                        line = txt_Rx.Text.Split('\r');
                        RunTest_result = 2;
                        for (i = line.GetUpperBound(0); i >= 0; i--)
                        {
                            if (line[i].Contains("SYNC OK"))
                            {
                                RunTest_result = 1;
                                break;  // for
                            }
                            if (line[i].Contains("atop_get_gpstime"))
                            {
                                break;  // for
                            }
                        }
                        if (RunTest_result == 1)
                        {
                            lblStatus.Text = cmd[0].ToUpper() + " Test Pass !";
                        }
                        else
                        {
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                        }
                        break;
                    case "BUZZER":
                        SendCmd("/jffs2/atop_buzzer.sh");
                        if (cmd.GetUpperBound(0) >= 1)
                        {
                            if (cmd[1].ToUpper() == "SKIP")
                            {
                                lblStatus.Text = "略過人工判斷";
                                Hold(2000); // wait for buzzer
                                RunTest_result = 1;
                            }
                        }
                        else if (chkHumanSkip.CheckState == CheckState.Checked)
                        {
                            lblStatus.Text = "略過人工判斷";
                            Hold(2000); // wait for buzzer
                            RunTest_result = 1;
                        }
                        else
                        {
                            lblStatus.Text = "人工判斷";
                            Hold(1000);
                            dr = MessageBox.Show("是否有聽到蜂鳴器發出聲響 ? ", cmd[0] + " Test", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                            if (dr == DialogResult.Yes) { RunTest_result = 1; }
                            else if (dr == DialogResult.No) { RecNote(idx, cmd[0].ToUpper() + " Test Fail."); RunTest_result = 2; }
                        }
                        break;
                    case "FLASH":
                        Hold(3000);
                        txt_Rx.Text = string.Empty;
                        RunTest_result = 2;
                        SendCmd("rm /jffs2/Auto_Test");
                        Hold(500);
                        SendCmd("cp -rf /jffs2 /tmp/"); // flash copy to ram
                        Hold(1000);
                        SendCmd("ls /tmp");
                        WaitKey = "jffs2";
                        if (Hold(1000))
                        {
                            SendCmd("touch /jffs2/Auto_Test");  // 建立檔案
                            Hold(500);
                            SendCmd("ls /jffs2");
                            WaitKey = "Auto_Test";
                            if (Hold(1000))
                            {
                                SendCmd("umount /jffs2/");
                                Hold(500);
                                SendCmd("/tmp/jffs2/flash_eraseall /dev/mtd5");
                                WaitKey = "100 % complete";
                                if (Hold(90000))
                                {
                                    SendCmd("mount -t jffs2 /dev/mtdblock5 /jffs2");
                                    Hold(1000);
                                    SendCmd("ls /jffs2");
                                    WaitKey = "Auto_Test";
                                    if (!Hold(1000))
                                    {
                                        RunTest_result = 1;
                                    }
                                    SendCmd("cp -rf /tmp/jffs2/ ~");
                                    Hold(20000); // ram copy to flash
                                    SendCmd("df");
                                    WaitKey = "/dev/mtdblock5";
                                    if (Hold(30000))
                                    {
                                        // 等待複製jffs2
                                    }
                                }
                            }
                        }
                        if (RunTest_result == 1)
                        {
                            lblStatus.Text = cmd[0].ToUpper() + " Test Pass !";
                        }
                        else
                        {
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                        }
                        break;
                    case "SD":
                        //RunTest_result = 2;
                        //// 檢測SD有無mount
                        //SendCmd("df -h");
                        //Hold(300);
                        //line = txt_Rx.Text.Split('\r');
                        //for (i = line.GetUpperBound(0); i >= 0; i--)
                        //{
                        //    if (line[i].Contains("/mnt/sd"))
                        //    {
                        //        // 進行SD測試
                        //        SendCmd("cd /jffs2/");
                        //        Hold(100);
                        //        //SendCmd("chmod 755 *");
                        //        //Hold(100);
                        //        //SendCmd("sync");
                        //        //Hold(100);
                        //        SendCmd("cp " + SD_function + " /mnt/sd/");
                        //        Hold(100);
                        //        SendCmd("/mnt/sd/" + SD_function);
                        //        Hold(300);
                        //        line = txt_Rx.Text.Split('\r');
                        //        for (i = line.GetUpperBound(0); i >= 0; i--)
                        //        {
                        //            if (line[i].Contains("GET SD OK"))
                        //            {
                        //                RunTest_result = 1;
                        //                break;  // for
                        //            }
                        //            if (line[i].Contains("/mnt/sd/" + SD_function))
                        //            {
                        //                break;  // for
                        //            }
                        //        }
                        //        break;  // for
                        //    }
                        //    if (line[i].Contains("df -h"))
                        //    {
                        //        break;  // for
                        //    }
                        //}
                        //if (RunTest_result == 1)
                        //{
                        //    lblStatus.Text = cmd[0].ToUpper() + " Test Pass !";
                        //}
                        //else
                        //{
                        //    RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                        //    lblStatus.Text = cmd[0].ToUpper() + " Test Fail !";
                        //}
                        break;
                    case "USB":
                        txt_Rx.Text = string.Empty;
                        if (MODEL_NAME.Equals("SE5901", StringComparison.CurrentCultureIgnoreCase)) // not SE5901A
                        {
                            SendCmd("cat /proc/partitions");
                            Hold(2000);
                            if (!txt_Rx.Text.Contains("sda1"))
                            {
                                return RunTest_result = 2;
                            }
                            else
                            {
                                SendCmd("mkdir /mnt/sda1");
                                Hold(100);
                                SendCmd("mount /dev/sda1 /mnt/sda1");
                                Hold(300);
                            }
                        }
                        // 偵測到，開始測試
                        SendCmd("df -h");
                        Hold(2000);
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)
                        {
                            if (line[i].Contains("/dev/sda1"))
                            {
                                delay = 0;
                                if (MODEL_NAME.Contains("SE5901A"))
                                {
                                    SendCmd("/jffs2/iozone -r " + cmd[1] + "K -s 32M -f /usb/sda1/Auto_Test");
                                }
                                else if (MODEL_NAME.Equals("SE5901", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    SendCmd("/jffs2/iozone -r " + cmd[1] + "K -s 32M -f /mnt/sda1/Auto_Test");
                                }

                                while (WAIT)
                                {
                                    delay = delay + 1;
                                    lblStatus.Text = "等待 USB IO讀寫速度測試..." + delay / 10;
                                    Hold(100);
                                    if (txt_Rx.Text.Contains("not found") || txt_Rx.Text.Contains("Killed") || txt_Rx.Text.Contains("No such file or directory"))
                                    {
                                        lblStatus.Text = "未載入 Tool 或腳本參數設定錯誤";
                                        return RunTest_result = 3;
                                    }
                                    else if (txt_Rx.Text.ToLower().Contains("error"))
                                    {
                                        lblStatus.Text = cmd[0] + " Test Fail !";
                                        RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                                        return RunTest_result = 2;
                                    }
                                    else if (txt_Rx.Text.ToLower().Contains("iozone test complete"))
                                    {
                                        lblStatus.Text = cmd[0] + " Test Pass !";
                                        return RunTest_result = 1;
                                    }
                                    //if (Run_Stop) { break; }
                                }
                                break;
                            }
                            if (line[i].Contains("df -h"))
                            {
                                RunTest_result = 3;
                                break;
                            }
                        }
                        break;
                    case "WATCHDOG":
                        SendCmd("/jffs2/atop_hwd 30 &");
                        if (cmd.GetUpperBound(0) >= 1)
                        {
                            if (cmd[1].ToUpper() == "KILL")
                            {
                                Hold(2000);
                                SendCmd("killall atop_hwd");
                                WaitKey = "U-Boot ";
                                if (Hold(60000))
                                {
                                    RunTest_result = 1;
                                    ReCntTelnet(50);
                                }
                                else
                                {
                                    RunTest_result = 2;
                                }
                            }
                        }
                        else
                        {
                            RunTest_result = 1;
                            WaitKey = "Disable Hardware Watchdog";
                            if (Hold(65000) == false)
                            {
                                RunTest_result = 2;
                                RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            }
                        }
                        break;
                    case "POWER":

                        break;
                    case "FTP":

                        break;
                    case "GWD":

                        break;
                    case "CPU":
                        //txt_Rx.Text = string.Empty;
                        string mips = string.Empty;
                        SendCmd("/jffs2/dhry2");
                        Hold(1000);
                        WaitKey = "Press Enter";
                        if (Hold(30000))
                        {
                            mips = (GetLine("MIPS", "MIPS").Replace("VAX  MIPS rating = ", string.Empty)).Trim();
                            if (Convert.ToSingle(mips) >= 925)
                            {
                                RunTest_result = 1;
                                lblStatus.Text = "VAX  MIPS rating = " + mips.ToString();
                            }
                            else // mips < 925
                            {
                                RunTest_result = 2;
                                lblStatus.Text = "VAX  MIPS rating = " + mips.ToString() + "低於 925";
                                RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            }
                        }
                        else
                        {
                            RunTest_result = 3;
                            lblStatus.Text = "CPU benchmark 測試超時";
                            RecNote(idx, "CPU benchmark 測試超時");
                        }
                        SendCmd("");
                        break;
                    case "MEMORY":
                        txt_Rx.Text = string.Empty;
                        if (cmd.GetUpperBound(0) == 0)
                        {
                            SendCmd("/jffs2/memtester 10M 1");
                        }
                        else
                        {
                            SendCmd("/jffs2/memtester " + cmd[1] + "M 1");
                        }
                        RunTest_result = 1;
                        delay = 0;
                        // *
                        WaitKey = "Done";
                        WAIT = true;
                        while (WAIT)
                        {
                            delay = delay + 1;
                            lblStatus.Text = "等待記憶體測試..." + delay / 10;
                            Hold(10); // UI
                            if (txt_Rx.Text.Contains("not found") || txt_Rx.Text.Contains("Killed"))
                            {
                                lblStatus.Text = "未載入 Tool 或腳本參數設定錯誤";
                                return RunTest_result = 3;
                            }
                            else if (txt_Rx.Text.ToLower().Contains("error"))
                            {
                                lblStatus.Text = cmd[0] + " Test Fail !";
                                RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                                return RunTest_result = 2;
                            }
                            // *
                            //else if (txt_Rx.Text.ToLower().Contains("done"))
                            //{
                            //    lblStatus.Text = cmd[0] + " Test Pass !";
                            //    return RunTest_result = 1;
                            //}
                            //if (Run_Stop) { break; }
                        }
                        break;
                    case "NETWORK":
                        if (chkHumanSkip.CheckState == CheckState.Unchecked)
                        {
                            if (MODEL_NAME == "SE5901")
                            {
                                if (MessageBox.Show("請插拔切換 LAN 1、LAN 2 網路線", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                                {
                                    return RunTest_result = 0;
                                }
                            }
                        }
                        int speed = 75;
                        string serverIP = string.Empty;
                        if (!Directory.Exists(@"C:\temp"))
                        {
                            Directory.CreateDirectory(@"C:\temp");
                        }
                        if (!File.Exists(@"C:\temp\netserver.exe"))
                        {
                            File.Copy(appPATH + @"\netserver.exe", @"C:\temp\netserver.exe");
                        }
                        NET_PID = Shell(@"C:\temp", "netserver.exe");
                        Hold(200);
                        if (cmd.GetUpperBound(0) == 2)
                        {
                            //
                            if (cmd[1].ToLower() == "lan")
                            {
                                speed = 75;
                                //if (MODEL_NAME.Contains("SE5901A"))
                                //{
                                //    SendCmd("ifconfig wlan0 down");
                                //    Hold(500);
                                //}
                            }
                            else if (cmd[1].ToLower() == "wifi")
                            {
                                speed = 3;
                                //if (MODEL_NAME.Contains("SE5901A"))
                                //{
                                //    SendCmd("ifconfig eth0 down");
                                //    Hold(500);
                                //}
                            }
                            else
                            {
                                lblStatus.Text = "腳本參數設定錯誤";
                                RunTest_result = 3;
                                goto endtest;
                            }
                            //
                            if (IsIP(cmd[2]))
                            {
                                serverIP = cmd[2];
                            }
                            else
                            {
                                lblStatus.Text = "腳本參數設定錯誤";
                                RunTest_result = 3;
                                goto endtest;
                            }
                        }
                        else
                        {
                            lblStatus.Text = "腳本參數設定錯誤";
                            RunTest_result = 3;
                            goto endtest;
                        }

                        lblStatus.Text = "網路測速...";
                        SendCmd("/jffs2/netperf -H " + serverIP);
                        Hold(100);
                        WaitKey = "Throughput";
                        if (Hold(30000) == false)
                        {
                            lblStatus.Text = "Network throughput test fail !";
                            SendCmd(((char)3).ToString()); // ((char)3):Ctrl+c
                            RecNote(idx, "Network throughput test fail !");
                            RunTest_result = 2;
                            goto endtest;
                        }
                        Hold(500);
                        line = txt_Rx.Text.Split('\r');
                        for (i = line.GetUpperBound(0); i >= 0; i--)    // 從尾巴先搜尋
                        {
                            if (line[i].Contains("Throughput"))
                            {
                                lblStatus.Text = line[i + 3];
                                break;
                            }
                        }
                        //lblStatus.Text = line[line.GetUpperBound(0) - 1];
                        if (lblStatus.Text.Length > 10)
                        {
                            Single throughput;
                            lblStatus.Text = lblStatus.Text.Substring(lblStatus.Text.Length - 10).Trim();
                            Regex reg = new Regex(@"(?:\d*\.)?\d+");
                            if (reg.IsMatch(lblStatus.Text))
                            {
                                throughput = Convert.ToSingle(lblStatus.Text);
                                if (throughput < speed)
                                {
                                    RunTest_result = 2;
                                    lblStatus.Text = lblStatus.Text + "Mbps < " + speed;
                                    RecNote(idx, cmd[0].ToUpper() + " Test Fail." + " (" + lblStatus.Text + ")");
                                }
                                else
                                {
                                    RunTest_result = 1;
                                    lblStatus.Text = lblStatus.Text + "Mbps";
                                }
                            }
                            else
                            {
                                lblStatus.Text = "Network throughput test cancel";
                                goto endtest;
                            }
                        }
                        endtest:
                        CloseShell(NET_PID);
                        //if (MODEL_NAME.Contains("SE5901A"))
                        //{
                        //    if (cmd[1].ToLower() == "lan")
                        //    {
                        //        SendCmd("ifconfig wlan0 up");
                        //        Hold(9000);
                        //    }
                        //    else if (cmd[1].ToLower() == "wifi")
                        //    {
                        //        SendCmd("ifconfig eth0 up");
                        //        Hold(1000);
                        //    }
                        //}
                        SendCmd("");
                        break;
                    case "SLEEP":
                        duration = Math.Round(TimeUnit(idx, 1) / 60, 2);
                        pause(duration);
                        break;
                    case "SYSTEM":
                        duration = Math.Round(TimeUnit(idx, 4) / 60, 2);
                        break;
                    case "RESTORE":
                        SendCmd("cd /jffs2/");
                        Hold(100);
                        SendCmd("killall atop_restored");
                        Hold(100);
                        MessageBox.Show("按下確定鍵後，請在 10 秒內按壓 default鍵 2 秒以上，直到聽見 bee 一聲後放開 default鍵。  ", cmd[0] + " test", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        SendCmd("./restored &");
                        RunTest_result = 1;
                        WaitKey = "sh: restart: not found";
                        if (Hold(10000) == false)
                        {
                            RecNote(idx, cmd[0].ToUpper() + " Test Fail.");
                            RunTest_result = 2;
                        }
                        SendCmd("killall restored");
                        break;
                    case "UPGRADE":
                        if (cmd[1] == "-1" && cmd[2] == "-1" && cmd[3] == "-1") { return RunTest_result = 1; }
                        // 檢查目錄資料夾 C:\TFTP-Root
                        if (!Directory.Exists(@"C:\TFTP-Root"))
                        {
                            Directory.CreateDirectory(@"C:\TFTP-Root");
                        }
                        for (i = 1; i <= 3; i++)
                        {
                            if (cmd[i] != "-1")
                            {
                                if (!File.Exists(@"C:\TFTP-Root\" + cmd[i]))
                                {
                                    MessageBox.Show(@"C:\TFTP-Root\" + cmd[i] + "  檔案不存在   ", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return RunTest_result = 3;
                                }
                            }
                        }
                        serialPort1.Write("atop_restart" + ((char)13).ToString());
                        Hold(200);
                        do
                        {
                            serialPort1.Write(((char)27).ToString());
                            WaitKey = ":TFTP";
                            Hold(1000);
                        } while (WAIT);
                        serialPort1.Write("6");
                        WaitKey = "Download All Image";
                        if (!Hold(1000)) { return RunTest_result = 3; };
                        // Set New TFTP Server IP
                        serialPort1.Write("1");
                        WaitKey = "Address";
                        if (!Hold(1000)) { return RunTest_result = 3; };
                        serialPort1.Write("10.0.50.233" + ((char)13).ToString());
                        WaitKey = "OK";
                        if (!Hold(1000)) { return RunTest_result = 3; };
                        // Download Linux Kernel
                        if (cmd[2] != "-1" && cmd[2].ToUpper().Contains("K"))
                        {
                            serialPort1.Write("3");
                            WaitKey = "Linux Image";
                            if (!Hold(1000))
                            {
                                return RunTest_result = 3;
                            }
                            serialPort1.Write(cmd[2] + ((char)13).ToString());
                            WaitKey = "U-Boot ";
                            if (!Hold(60000))
                            {
                                SendCmd(((char)3).ToString());
                                Hold(100);
                                serialPort1.Write("0");
                                Hold(100);
                                serialPort1.Write("0");
                                RunTest_result = 3;
                            }
                            else
                            {
                                if (cmd[3] == "-1" && cmd[1] == "-1") { return RunTest_result = 1; }
                                else
                                {
                                    do
                                    {
                                        serialPort1.Write(((char)27).ToString());
                                        WaitKey = ":TFTP";
                                        Hold(1000);
                                    } while (WAIT);
                                    serialPort1.Write("6");
                                    WaitKey = "Download All Image";
                                    if (!Hold(1000)) { return RunTest_result = 3; };
                                }
                            }
                        }
                        // Download Linux RAMDisk Image
                        if (cmd[3] != "-1" && cmd[3].ToUpper().Contains("A"))
                        {
                            serialPort1.Write("4");
                            WaitKey = "Linux Image";
                            if (!Hold(1000))
                            {
                                return RunTest_result = 3;
                            }
                            serialPort1.Write(cmd[3] + ((char)13).ToString());
                            WaitKey = "U-Boot ";
                            if (!Hold(150000))
                            {
                                SendCmd(((char)3).ToString());
                                Hold(100);
                                serialPort1.Write("0");
                                Hold(100);
                                serialPort1.Write("0");
                                RunTest_result = 3;
                            }
                            else
                            {
                                if (cmd[1] == "-1") { return RunTest_result = 1; }
                                else
                                {
                                    do
                                    {
                                        serialPort1.Write(((char)27).ToString());
                                        WaitKey = ":TFTP";
                                        Hold(1000);
                                    } while (WAIT);
                                    serialPort1.Write("6");
                                    WaitKey = "Download All Image";
                                    if (!Hold(1000)) { return RunTest_result = 3; };
                                }
                            }
                        }
                        // Download Bootload
                        if (cmd[1] != "-1" && cmd[1].ToUpper().Contains("B"))
                        {
                            serialPort1.Write("2");
                            WaitKey = "input Bootloader";
                            if (!Hold(1000))
                            {
                                return RunTest_result = 3;
                            }
                            serialPort1.Write(cmd[1] + ((char)13).ToString());
                            WaitKey = "U-Boot ";
                            if (!Hold(15000))
                            {
                                SendCmd(((char)3).ToString());
                                Hold(100);
                                serialPort1.Write("0");
                                Hold(100);
                                serialPort1.Write("0");
                                RunTest_result = 3;
                            }
                            else
                            {
                                MessageBox.Show("網路線請更換為有支援 1000M   ", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); // for SE1908-4U message
                                RunTest_result = 1;
                            }
                        }
                        ReCntTelnet(26);
                        break;
                    case "APINFOR":
                        if (cmd.Length == 2)
                        {
                            if (cmd[1] != string.Empty)
                            {
                                SendCmd("/jffs2/atop_ap_ver " + cmd[1]);
                                Hold(200);
                            }
                            string APVer = GetLine("Get customer", "atop_ap_ver");
                            if (APVer.Contains(cmd[1]))
                            {
                                RunTest_result = 1;
                                lblStatus.Text = cmd[0] + " Test Pass.";
                            }
                            else
                            {
                                RunTest_result = 2;
                                lblStatus.Text = cmd[0] + " Test Fail.";
                            }
                        }
                        break;
                    case "POWERDETECT":
                        string detectResult = string.Empty;
                        SendCmd("/jffs2/atop_pow_detect");
                        Hold(6000);
                        detectResult = GetLine("Power from", "atop_pow_detect");
                        RunTest_result = 1;
                        lblStatus.Text = detectResult;
                        break;
                    default:
                        break;
                }
                // Excel log
                if (TEST_FunLog.Contains(cmd[0].ToUpper()))
                {
                    if (RunTest_result == 1)
                    {
                        TEST_RESULT[idx_funlog] = TEST_RESULT[idx_funlog] + "o";
                    }
                    else if (RunTest_result == 2)
                    {
                        TEST_RESULT[idx_funlog] = TEST_RESULT[idx_funlog] + "X";
                    }
                    else if (RunTest_result == 3)
                    {
                        TEST_RESULT[idx_funlog] = TEST_RESULT[idx_funlog] + "-";
                    }
                }
                return RunTest_result;  // switch use
            }
            catch (Exception ex)
            {
                RecNote(idx, ex.Message);
                SendCmd(((char)3).ToString()); // ((char)3):Ctrl+c
                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Run_Stop = true;
                return RunTest_result = 3;
            }
            finally
            {
                if (RunTest_result == 1)
                {
                    lblFunction[idx].BackColor = Color.FromArgb(0, 255, 0); /* 1:PASS Green */
                }
                else if (RunTest_result == 2)
                {
                    lblFunction[idx].BackColor = Color.FromArgb(255, 0, 0); /* 2:Fail Red */
                }
                else if (RunTest_result == 3)
                {
                    lblFunction[idx].BackColor = Color.FromArgb(255, 255, 0); /* 3:error Yellow */
                }
                else if (RunTest_result == 0) { lblFunction[idx].BackColor = Color.FromArgb(255, 255, 255); /* 0 */}
            }
        }

        /// <summary>
        /// 回傳關鍵字所在的整行文字
        /// </summary>
        /// <param name="Key">目標關鍵字</param>
        /// <param name="stopSearch">如果目標關鍵字不存在，則搜尋到stopSearch時停止搜尋</param>
        /// <returns>回傳整行文字，關鍵字不存在則回傳string.Empty</returns>
        private string GetLine(string Key, string stopSearch)
        {
            int i;
            string[] line;
            string get_line = string.Empty;
            line = txt_Rx.Text.Split('\r');
            for (i = line.GetUpperBound(0); i >= 0; i--)
            {
                if (line[i].Contains(Key))
                {
                    get_line = line[i].Replace("\n", "");
                    break;  // for
                }
                else if (line[i].Contains(stopSearch))
                {
                    break;
                }
            }
            return get_line;
        }

        // 判斷陣列中有無string.empty，並刪除此元素
        private static bool isNotStringEmpty(string element)
        {
            return element != "";
        }

        private void txt_Tx_KeyDown(object sender, KeyEventArgs e)
        {
            // Initialize the flag to false.
            nonNumberEntered = false;
            int key = e.KeyValue;
            //if (e.Control != true)//如果沒按Ctrl鍵
            //    return;
            switch (key)
            {
                case 13:
                    //按下Enter以後
                    SendCmd(txt_Tx.Text);
                    txt_Tx.Text = string.Empty;
                    nonNumberEntered = true;
                    break;
                case 38:
                    //按下向上鍵以後
                    SendCmd(((char)27).ToString() + ((char)91).ToString() + ((char)65).ToString()); // ←[A
                    nonNumberEntered = true;
                    break;
                case 40:
                    //按下向下鍵以後
                    SendCmd(((char)27).ToString() + ((char)91).ToString() + ((char)66).ToString()); // ←[B
                    nonNumberEntered = true;
                    break;
                default:
                    break;
            }
        }

        private void txt_Tx_KeyPress(object sender, KeyPressEventArgs e)
        {
            // KeyChar 無法抓取上下左右鍵
            // http://msdn.microsoft.com/zh-tw/library/system.windows.forms.keyeventargs.handled%28v=vs.110%29.aspx
            // Check for the flag being set in the KeyDown event.
            if (nonNumberEntered)
            {
                // Stop the character from being entered into the control since it is non-numerical.
                e.Handled = true;
            }
        }

        private void consoleToolStripMenuItem_CheckStateChanged(object sender, EventArgs e)
        {
            if (consoleToolStripMenuItem.Checked)
                tabControl1.SelectedTab = tabPage1;
            else if (tabControl1.SelectedTab == tabPage3)
                tabControl1.SelectedTab = tabPage3;
            else
                tabControl1.SelectedTab = tabPage2;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
                consoleToolStripMenuItem.Checked = true;
            else
                consoleToolStripMenuItem.Checked = false;
        }

        #region 自動保持 TextBox 垂直捲軸在最下方

        private void txt_Rx_TextChanged(object sender, EventArgs e)
        {
            // 自動保持捲軸在最下方
            txt_Rx.SelectionStart = txt_Rx.Text.Length;
            txt_Rx.ScrollToCaret();
        }

        private void txt_Note_TextChanged(object sender, EventArgs e)
        {
            txt_Note.SelectionStart = txt_Note.Text.Length;
            txt_Note.ScrollToCaret();
        }

        private void txt_Rx_EUT_TextChanged(object sender, EventArgs e)
        {
            // 自動保持捲軸在最下方
            txt_Rx_EUT.SelectionStart = txt_Rx_EUT.Text.Length;
            txt_Rx_EUT.ScrollToCaret();
        }

        #endregion 自動保持 TextBox 垂直捲軸在最下方

        private void composingTmr_Tick(object sender, EventArgs e)
        {
            int idx, X_StartPos, Y_StartPos;
            int X, Y;   // every position(location) of the panel
            X_StartPos = 52; Y_StartPos = 25;    // initial position(location) of the panel
            row_num = (this.Height - Y_StartPos) / (lblFunction[0].Height * 2) - 6;
            for (idx = 0; idx < TestFun_MaxIdx; idx++)    // composing Label
            {
                X = X_StartPos + (idx / row_num) * X_StartPos * 3;
                Y = Y_StartPos + (lblFunction[idx].Height * (idx % row_num) * 2);
                lblFunction[idx].Location = new Point(X, Y);
                lblFunction[idx].Visible = true;
            }
            composingTmr.Enabled = false;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (lblFunction[0] != null)
            {
                composingTmr.Enabled = true;
            }
        }

        private void 從這個測項開始測試ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Test_Idx = MOUSE_Idx;
            chooseStart = true;
            cmdStart_Click(null, null);
        }

        private void 無限次測試這個測項ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cmdOpeFile.Enabled = false;
            cmdStart.Enabled = false;
            cmdStop.Enabled = true;
            cmdNext.Enabled = false;
            Run_Stop = false;
            do
            {
                TEST_STATUS[MOUSE_Idx] = RunTest(MOUSE_Idx);
                if (STOP_WHEN_FAIL && Convert.ToInt32(TEST_STATUS[MOUSE_Idx]) == 2)
                {
                    return;
                }
                Hold(1000);
            } while (Run_Stop == false);
            cmdOpeFile.Enabled = true;
            cmdStart.Enabled = true;
            cmdStop.Enabled = true;
            cmdNext.Enabled = true;
        }

        private void 用Putty開啟ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string[] cmd;
            cmd = Convert.ToString(lblFunction[MOUSE_Idx].Tag).Split(' ');
            // Uses the ProcessStartInfo class to start new processes
            ProcessStartInfo startInfo = new ProcessStartInfo("putty.exe");
            startInfo.UseShellExecute = false;
            if (cmd[0].ToUpper() == "CONSOLE-DUT" || cmd[0].ToUpper() == "CONSOLE-EUT")
            {
                if (cmd[0].ToUpper() == "CONSOLE-DUT")
                {
                    serialPort1_Close();
                    //if (serialPort1.IsOpen) { serialPort1.Close(); }
                }
                else if (cmd[0].ToUpper() == "CONSOLE-EUT")
                {
                    //serialPort2_Close();
                    //if (serialPort2.IsOpen) { serialPort2.Close(); }
                }
                string info = "-serial COM" + cmd[1] + " -sercfg " + cmd[2] + ",8,n,1,n";
                startInfo.Arguments = info;
                Process.Start(startInfo);
            }
            else if (cmd[0].ToUpper() == "TELNET")
            {
                //USR = cmd[1];
                //if (cmd.GetUpperBound(0) > 1) { PWD = cmd[2]; }
                //else { PWD = string.Empty; }
                startInfo.Arguments = "-telnet -t " + TARGET_IP;
                Process.Start(startInfo);
                Hold(1000);
                SendKeys.SendWait(USR + "{ENTER}");
                Hold(1000);
                SendKeys.SendWait(PWD + "{ENTER}");
            }
            else if (cmd[0].ToUpper() == "POWER")
            {
            }
        }

        #region Hold / atop_timer

        public bool Hold(long timeout)
        {
            bool tmp_Hold = true;
            long delay = 0;
            WAIT = true;
            if (timeout > 0) { delay = timeout / 10; }
            while (WAIT)
            {
                Application.DoEvents();
                Thread.Sleep(10);
                if (timeout > 0)
                {
                    if (delay > 0)
                    {
                        delay -= 1;
                    }
                    else
                    {
                        tmp_Hold = false;   // 時間等到底
                        break;
                    }
                }
            }
            return tmp_Hold;
        }

        #endregion Hold / atop_timer

        #region lblStatus.ForeColor 隨著測試項目改變而變化Color

        // RGB to Hex
        // http://www.rapidtables.com/convert/color/rgb-to-hex.htm
        private void timer2_Tick(object sender, EventArgs e)
        {
            //Debug.Print(lblStatus.ForeColor.ToArgb().ToString());
            if (lblStatus.ForeColor.ToArgb() > 10 * 65536)
            {
                int hex_tmp = Convert.ToInt32(lblStatus.ForeColor.ToArgb());
                lblStatus.ForeColor = Color.FromArgb(hex_tmp - 50 * 65536);
            }
        }

        private void lblStatus_TextChanged(object sender, EventArgs e)
        {
            lblStatus.ForeColor = Color.FromArgb(255 * 65536);
        }

        #endregion lblStatus.ForeColor 隨著測試項目改變而變化Color

        public int ReCntTelnet(long timeout)
        {
            if (serialPort1.IsOpen)
            {
                WaitKey = "login";
                enterTmr.Enabled = true;    // 5秒按一次enter
                if (Hold(timeout * 1000) == false)
                {
                    enterTmr.Enabled = false;
                    return 2;
                }
                else
                {
                    enterTmr.Enabled = false;
                    return 1;
                }
            }
            else
            {
                int tm = 0;
                lblStatus.Text = "等待系統重開機...";
                do
                {
                    Hold(1000);
                    tm += 1;
                    if (tm > (timeout / 2))
                    {
                        lblStatus.Text = "連線失敗";
                        return 2; // 逾時
                    }
                } while (objping.Send(TARGET_IP, 1000).Status != IPStatus.Success);

                telnet = new TcpClient();
                if (!telnet.Connected)
                {
                    try
                    {
                        telnet.Connect(TARGET_IP, 23);   // 連接23端口 (Telnet的默認端口)
                        telentStream = telnet.GetStream();  // 建立網路資料流，將字串寫入串流

                        if (telnet.Connected)
                        {
                            //lblStatus.Text = "連線成功，正在登錄...";
                            lblStatus.Text = "正在登錄...";
                            Hold(1000);
                            // 背景telnet接收執行緒
                            ThreadStart backgroundReceive = new ThreadStart(telnet_Receive);
                            Thread rcvThread = new Thread(backgroundReceive);
                            rcvThread.IsBackground = true;
                            rcvThread.Start();

                            bytWrite_telnet = System.Text.Encoding.Default.GetBytes(USR + ((char)13).ToString());
                            telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                            Hold(200);
                            bytWrite_telnet = System.Text.Encoding.Default.GetBytes(PWD + ((char)13).ToString());
                            telentStream.Write(bytWrite_telnet, 0, bytWrite_telnet.Length);
                            lblStatus.Text = "連線成功";
                            return 1;
                        }
                    }
                    catch (Exception)
                    {
                        return 2;   // 目標主機連線沒反應
                    }
                }
            }
            return 2;
        }

        #region 驗證IP

        /// <summary>
        /// 驗證IP
        /// </summary>
        /// <param name="source"></param>
        /// <returns>規則運算式尋找到符合項目，則為 true，否則為 false</returns>
        public static bool IsIP(string source)
        {
            // Regex.IsMatch 方法 (String, String, RegexOptions)
            // 表示指定的規則運算式是否使用指定的比對選項，在指定的輸入字串中尋找相符項目
            return Regex.IsMatch(source, @"^(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])$", RegexOptions.IgnoreCase);
        }

        public static bool HasIP(string source)
        {
            return Regex.IsMatch(source, @"(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])", RegexOptions.IgnoreCase);
        }

        #endregion 驗證IP

        #region FTP

        /// <summary>
        /// FTP 上傳檔案至目標位置
        /// </summary>
        /// <param name="FTPAddress">目標位置</param>
        /// <param name="filePath">上傳的檔案</param>
        /// <param name="username">帳號</param>
        /// <param name="password">密碼</param>
        public void uploadFile(string IP, string filePath, string username, string password)
        {
            //if (!IP.StartsWith("ftp://")) { IP = "ftp://" + IP; }
            string FTPAddress = "ftp://" + IP;
            //Create FTP request
            FtpWebRequest request = (FtpWebRequest)FtpWebRequest.Create(FTPAddress + "/" + Path.GetFileName(filePath));
            request.Method = WebRequestMethods.Ftp.UploadFile;
            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(username, password);
            request.UsePassive = true;
            request.UseBinary = true;
            request.KeepAlive = false;
            request.ReadWriteTimeout = 7000;
            request.Timeout = 3000;

            //Load the file
            FileStream stream = File.OpenRead(filePath);
            byte[] buffer = new byte[stream.Length];

            stream.Read(buffer, 0, buffer.Length);
            stream.Close();

            //Upload file
            Stream reqStream = request.GetRequestStream();
            reqStream.Write(buffer, 0, buffer.Length);
            reqStream.Close();

            //Debug.Print("Uploaded Successfully !");
        }

        /// <summary>
        /// 列出 FTP 目錄的內容，並檢查檔案是否存在內容中。
        /// </summary>
        /// <param name="IP">目標 IP</param>
        /// <param name="fileName">欲檢查的檔案</param>
        /// <param name="username">帳號</param>
        /// <param name="password">密碼</param>
        /// <returns>true代表存在；false代表不存在</returns>
        public bool checkFile(string IP, string fileName, string username, string password)
        {
            string FTPAddress = "ftp://" + IP;
            //Create FTP request
            FtpWebRequest request = (FtpWebRequest)FtpWebRequest.Create(FTPAddress);
            request.Method = WebRequestMethods.Ftp.ListDirectory;
            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(username, password);
            request.UsePassive = true;
            request.UseBinary = true;
            request.KeepAlive = false;
            request.ReadWriteTimeout = 5000;
            request.Timeout = 3000;

            string responseTmp = string.Empty;
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            responseTmp = reader.ReadToEnd();
            if (responseTmp.Contains(fileName))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        ///* Rename File */
        //public void renameFile(string IP, string currentFileNameAndPath, string newFileName, string username, string password)
        //{
        //    string FTPAddress = "ftp://" + IP + "//jffs2";
        //    //Create FTP request
        //    FtpWebRequest request = (FtpWebRequest)FtpWebRequest.Create(FTPAddress);
        //    // This example assumes the FTP site uses anonymous logon.
        //    request.Credentials = new NetworkCredential(username, password);
        //    /* When in doubt, use these options */
        //    request.UseBinary = true;
        //    request.UsePassive = true;
        //    request.KeepAlive = false;
        //    /* Specify the Type of FTP Request */
        //    request.Method = WebRequestMethods.Ftp.Rename;
        //    /* Rename the File */
        //    request.RenameTo = newFileName;
        //    /* Establish Return Communication with the FTP Server */
        //    request = (FtpWebResponse)request.GetResponse();
        //    /* Resource Cleanup */
        //    request.Close();
        //    request = null;
        //}

        ///* Delete File */
        //public void deleteFile(string deleteFile)
        //{
        //    try
        //    {
        //        /* Create an FTP Request */
        //        ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + deleteFile);
        //        /* Log in to the FTP Server with the User Name and Password Provided */
        //        ftpRequest.Credentials = new NetworkCredential(user, pass);
        //        /* When in doubt, use these options */
        //        ftpRequest.UseBinary = true;
        //        ftpRequest.UsePassive = true;
        //        ftpRequest.KeepAlive = true;
        //        /* Specify the Type of FTP Request */
        //        ftpRequest.Method = WebRequestMethods.Ftp.DeleteFile;
        //        /* Establish Return Communication with the FTP Server */
        //        ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
        //        /* Resource Cleanup */
        //        ftpResponse.Close();
        //        ftpRequest = null;
        //    }
        //    catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        //    return;
        //}

        #endregion FTP

        #region 貓頭鷹 v1.6p (2013/11/22)

        private void MultiPortTesting_settings(string ip, int interval, string max_port, int server_port, int loopback, string duration)
        {
            int i;
            int min_port;

            if (max_port.Contains("-")) // 格式ex: 0-4、1-4
            {
                string[] port;
                port = max_port.Split(new char[] { '-' });      // 0-4 跳port數；1-4 全port數
                min_port = Convert.ToInt32(port[0]);
                max_port = port[1];
            }
            else
            {
                min_port = Convert.ToInt32(max_port);    // 格式ex: 4   指定單一port
            }

            if (File.Exists(appPATH + "\\setting.txt"))
            {
                File.Delete(appPATH + "\\setting.txt");
            }

            // 建立檔案
            FileStream fs = File.Open(appPATH + "\\setting.txt", FileMode.OpenOrCreate, FileAccess.Write);
            // 建構StreamWriter物件
            StreamWriter sw = new StreamWriter(fs);

            // 寫入
            sw.WriteLine(ip);           // IP
            sw.WriteLine("50");         // Send Lenth
            sw.WriteLine(interval);     // Send Interval
            sw.WriteLine(max_port);     // total port
            sw.WriteLine(server_port);
            sw.WriteLine(server_port);
            sw.WriteLine("5");          // timeout
            sw.WriteLine("0");          // pingpong
            sw.WriteLine("0");
            sw.WriteLine("0");
            sw.WriteLine("0");
            sw.WriteLine("0");
            sw.WriteLine("True");
            sw.WriteLine("False");
            sw.WriteLine("False");
            sw.WriteLine("0");
            sw.WriteLine(loopback);
            //sw.WriteLine(duration);
            sw.WriteLine("9999");
            for (i = 1; i <= 32; i++)
            {
                if (min_port <= i && i <= Convert.ToInt32(max_port))
                {
                    if (min_port == 0)
                    {
                        if (i % 2 == 1)
                        {
                            sw.WriteLine(Convert.ToString(Math.Abs(loopback - 1)));
                        }
                        else
                        {
                            sw.WriteLine(loopback);
                        }
                    }
                    else
                    {
                        sw.WriteLine("1");
                    }
                }
                else
                {
                    sw.WriteLine("0");
                }
            }

            // 清除目前寫入器(Writer)的所有緩衝區，並且造成任何緩衝資料都寫入基礎資料流
            sw.Flush();

            // 關閉目前的StreamWriter物件和基礎資料流
            sw.Close();
            fs.Close();
        }

        #endregion 貓頭鷹 v1.6p (2013/11/22)

        private float TimeUnit(int idx, int i)
        {
            string[] line;
            string tag = Convert.ToString(lblFunction[idx].Tag);
            line = tag.Split(' ');
            if (line.GetUpperBound(0) >= i + 1)
            {
                switch (line[i + 1].ToLower())
                {
                    case "s":
                        return Convert.ToInt64(line[i]) * 1;
                    case "m":
                        return Convert.ToInt64(line[i]) * 60;
                    case "h":
                        return Convert.ToInt64(line[i]) * 60 * 60;
                    case "d":
                        return Convert.ToInt64(line[i]) * 60 * 60 * 24;
                    default:
                        return Convert.ToInt64(line[i]) * 60;
                }
            }
            else { return Convert.ToInt64(line[i]) * 60; }
        }

        private void pause(double delay)
        {
            DateTime time_before = DateTime.Now;
            while (((TimeSpan)(DateTime.Now - time_before)).TotalMinutes < delay)
            {
                Application.DoEvents();
                Thread.Sleep(1000); // 至少打資料兩次
            }
        }

        private void lblSecret_Click(object sender, EventArgs e)
        {
            secretX += 1;
            if (secretX == 5)
            {
                debugMode.Visible = true;
                txt_Rx_EUT.Visible = true;
            }
        }

        private void disconnectALL_Click(object sender, EventArgs e)
        {
            int n;
            if (telnet.Connected) { telnet.Close(); }
            serialPort1_Close();
            //serialPort2_Close();
            //if (serialPort1.IsOpen) { serialPort1.Close(); }
            //if (serialPort2.IsOpen) { serialPort2.Close(); }

            Run_Stop = true;
            WAIT = false;

            for (n = 0; n < TestFun_MaxIdx; n++)
            {
                if (lblFunction[n].Text.ToUpper().Contains("CONSOLE-DUT") || lblFunction[n].Text.ToUpper().Contains("CONSOLE-EUT") || lblFunction[n].Text.ToUpper().Contains("TELNET"))
                {
                    lblFunction[n].BackColor = Color.FromArgb(255, 255, 255);
                }
            }
            lblStatus.Text = "所有的連線已經中斷";
        }

        private void lanEnvironmentSetting_Click(object sender, EventArgs e)
        {
            Shell(appPATH, "LAN_setting.bat");
        }

        private void 儲存Results分頁資訊_Click(object sender, EventArgs e)
        {
            //if (File.Exists(appPATH + "\\result.txt"))
            //{
            //    File.Delete(appPATH + "\\result.txt");
            //}
            // 建立檔案
            FileStream fs = File.Open(appPATH + "\\result.txt", FileMode.OpenOrCreate, FileAccess.Write);
            // 建構StreamWriter物件
            StreamWriter sw = new StreamWriter(fs);
            // 寫入
            sw.WriteLine(txt_Note.Text);
            // 清除目前寫入器(Writer)的所有緩衝區，並且造成任何緩衝資料都寫入基礎資料流
            sw.Flush();
            // 關閉目前的StreamWriter物件和基礎資料流
            sw.Close();
            fs.Close();
        }

        private void 開啟Monitor_Click(object sender, EventArgs e)
        {
            Shell(appPATH, "monitor2.5.exe");
        }

        private void 執行TFTPServer_Click(object sender, EventArgs e)
        {
            Shell(appPATH + "\\tftpd32.400", "tftpd32.exe");
        }

        // http://msdn.microsoft.com/zh-cn/library/aa168292(office.11).aspx
        // 設定必要的物件
        // 按照順序分別是Application > Workbook > Worksheet > Range > Cell
        // (1) Application ：代表一個 Excel 程序。
        // (2) WorkBook ：代表一個 Excel 工作簿。
        // (3) WorkSheet ：代表一個 Excel 工作表，一個 WorkBook 包含好幾個工作表。
        // (4) Range ：代表 WorkSheet 中的多個單元格區域。
        // (5) Cell ：代表 WorkSheet 中的一個單元格。
        private void writeExcelLog()
        {
            int j;

            // 檢查路徑的資料夾是否存在，沒有則建立
            if (!Directory.Exists(@"C:\Atop_Log\ATC\" + MODEL_NAME))
            {
                Directory.CreateDirectory(@"C:\Atop_Log\ATC\" + MODEL_NAME);
            }

            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名
            // C:\Atop_Log\ATC\產品名稱\年_月_工單號碼.xls
            time = DateTime.Now;
            string name = time.Year + "_" + time.Month + "_" + productNum_forExcel.ToUpper() + ".xls";
            string pathFile = @"C:\Atop_Log\ATC\" + MODEL_NAME + @"\" + name;

            Microsoft.Office.Interop.Excel.Application excelApp;
            Microsoft.Office.Interop.Excel._Workbook wBook;
            Microsoft.Office.Interop.Excel._Worksheet wSheet;
            Microsoft.Office.Interop.Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            // 讓Excel文件可見
            excelApp.Visible = false;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 開啟舊檔案
            if (GetFiles(pathFile))
            {
                wBook = excelApp.Workbooks.Open(pathFile, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            // 創建一個新的工作簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();
            // 設定開啟密碼
            //wBook.Password = "23242249";

            try
            {
                // 引用第一個工作表(轉型)
                wSheet = (Microsoft.Office.Interop.Excel._Worksheet)wBook.Worksheets[1];
                // 命名工作表的名稱
                wSheet.Name = "測試紀錄";
                // Worksheet.Protect 方法。保護工作表，使工作表無法修改
                wSheet.Protect("23242249", true, true, true, true, true, true, true, true, true, true, true, true, true, true, true);
                // 設定工作表焦點
                wSheet.Activate();
                // 所有儲存格 文字置中
                excelApp.Cells.HorizontalAlignment = 3;
                // 所有儲存格 自動換行
                excelApp.Cells.WrapText = true;
                // 所有儲存格格式強迫以文字來儲存
                excelApp.Cells.NumberFormat = "@";

                // 設定第1列資料
                excelApp.Cells[1, 1] = "測試人員";
                excelApp.Cells[1, 2] = "工單號碼";
                excelApp.Cells[1, 3] = "產品序號(SN)";
                excelApp.Cells[1, 4] = "產品名稱";
                excelApp.Cells[1, 5] = "MAC1";
                excelApp.Cells[1, 6] = "Kernel";
                excelApp.Cells[1, 7] = "AP";
                excelApp.Cells[1, 8] = "開始測試時間";
                excelApp.Cells[1, 9] = "結束測試時間";
                // 取得已經使用的Columns數(X軸)
                //int usedRangeColumns = wSheet.UsedRange.Columns.Count + 1;
                //for (j = usedRangeColumns; j < TEST_RESULT.Count + usedRangeColumns; j++)
                //{
                //    excelApp.Cells[1, j] = TEST_RESULT[j - usedRangeColumns];
                //}
                for (j = 10; j < TEST_FunLog.Count + 10; j++)
                {
                    excelApp.Cells[1, j] = TEST_FunLog[j - 10];
                    //Debug.Print(TEST_FunLog[j - 10].ToString());
                }
                // 設定第1列顏色
                wRange = wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, TEST_FunLog.Count + 9]);
                wRange.Select();
                wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
                //wRange.Columns.AutoFit();   // 自動調整欄寬
                wRange.ColumnWidth = 15; // 設置儲存格的寬度

                // 取得已經使用的Rows數(Y軸)
                int usedRangeRows = wSheet.UsedRange.Rows.Count + 1;
                // 設定第usedRange列資料
                excelApp.Cells[usedRangeRows, 1] = tester_forExcel.ToUpper();
                excelApp.Cells[usedRangeRows, 2] = productNum_forExcel.ToUpper();
                string snTemp = string.Empty;
                if (coreSN_forExcel != string.Empty && coreSN_forExcel != null)
                {
                    snTemp = "Core:" + coreSN_forExcel;
                }
                if (lanSN_forExcel != string.Empty && lanSN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Lan:" + lanSN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Lan:" + lanSN_forExcel;
                    }
                }
                if (uartSN_forExcel != string.Empty && uartSN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Uart:" + uartSN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Uart:" + uartSN_forExcel;
                    }
                }
                if (serial1SN_forExcel != string.Empty && serial1SN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Serial1:" + serial1SN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Serial1:" + serial1SN_forExcel;
                    }
                }
                if (serial2SN_forExcel != string.Empty && serial2SN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Serial2:" + serial2SN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Serial2:" + serial2SN_forExcel;
                    }
                }
                if (serial3SN_forExcel != string.Empty && serial3SN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Serial3:" + serial3SN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Serial3:" + serial3SN_forExcel;
                    }
                }
                if (serial4SN_forExcel != string.Empty && serial4SN_forExcel != null)
                {
                    if (snTemp == string.Empty)
                    {
                        snTemp = "Serial4:" + serial4SN_forExcel;
                    }
                    else
                    {
                        snTemp = snTemp + ((char)10).ToString() + "Serial4:" + serial4SN_forExcel;
                    }
                }
                excelApp.Cells[usedRangeRows, 3] = snTemp;
                excelApp.Cells[usedRangeRows, 4] = MODEL_NAME;
                excelApp.Cells[usedRangeRows, 5] = "";
                excelApp.Cells[usedRangeRows, 6] = "";
                excelApp.Cells[usedRangeRows, 7] = "";
                excelApp.Cells[usedRangeRows, 8] = startTime;
                excelApp.Cells[usedRangeRows, 9] = endTime;
                for (j = 10; j < TEST_FunLog.Count + 10; j++)
                {
                    excelApp.Cells[usedRangeRows, j] = TEST_RESULT[j - 10];
                }

                try
                {
                    // 另存活頁簿
                    wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MessageBox.Show("Excel log 儲存於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("產生 Excel log 時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            wBook.Close(false, Type.Missing, Type.Missing);

            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();
        }

        // 讀取目錄下所有檔案，並判斷指定檔案(不含副檔名)是否存在
        private bool GetFiles(string filename)
        {
            int i;
            string[] files;
            string keyword;

            files = Directory.GetFiles(@"C:\Atop_Log\ATC\" + MODEL_NAME);
            keyword = filename.Replace("C:\\Atop_Log\\ATC\\" + MODEL_NAME + "\\", string.Empty);
            for (i = 0; i < files.Length; i++)
            {
                files[i] = files[i].Replace(@"C:\Atop_Log\ATC\" + MODEL_NAME + "\\", string.Empty);
                if (files[i].Contains(keyword))
                {
                    return true;
                }
            }
            return false;
        }

        private void cmdNext_Click(object sender, EventArgs e)
        {
            int n;
            if (cmdOpeFile.Text != "檔案名稱")
            {
                InputBox inputbox = new InputBox();
                inputbox.ShowDialog();
                tester_forExcel = InputBox.tester;
                productNum_forExcel = InputBox.productNum;
                coreSN_forExcel = InputBox.coreSN;
                lanSN_forExcel = InputBox.lanSN;
                uartSN_forExcel = InputBox.uartSN;
                serial1SN_forExcel = InputBox.serial1SN;
                serial2SN_forExcel = InputBox.serial2SN;
                serial3SN_forExcel = InputBox.serial3SN;
                serial4SN_forExcel = InputBox.serial4SN;

                time = DateTime.Now;
                endTime = String.Format("{0:00}/{1:00}" + ((char)10).ToString() + "{2:00}:{3:00}:{4:00}", time.Month, time.Day, time.Hour, time.Minute, time.Second);

                writeExcelLog();

                Shell(appPATH, "arp-d.bat");
                if (telnet.Connected) { telnet.Close(); }
                serialPort1_Close();
                //serialPort2_Close();

                Test_Idx = 0;
                Run_Stop = true;
                WAIT = false;
                txt_Rx.Text = string.Empty;
                for (n = 0; n < TestFun_MaxIdx; n++)
                {
                    lblFunction[n].BackColor = Color.FromArgb(255, 255, 255);
                }
                for (n = 0; n < TEST_RESULT.Length; n++)
                {
                    TEST_RESULT[n] = string.Empty;
                }
            }
        }

        /// <summary>
        /// 為了設備Restart後，等待login字串而沒登入成功，導致循環測試無法順利，所以加此timer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void enterTmr_Tick(object sender, EventArgs e)
        {
            SendCmd("");
        }

        private void txtDutIP_TextChanged(object sender, EventArgs e)
        {
            if (IsIP(txtDutIP.Text))
            {
                TARGET_IP = txtDutIP.Text;
                txtDutIP.ForeColor = Color.Green;
            }
            else
            {
                txtDutIP.ForeColor = Color.Red;
            }
        }
    }
}

/*
 -----選取 columns-----
 xlWs.columns("H").select	'選取單行
 xlWs.columns("E:H").select	'選取連續行
 xlWs.columns("E:E,H:H")	'選取多行
 xlWs.range("E:E,G:G").select	'用range選取多行
 xlWs.columns.select	'選取全部行 = 全選
 -----用數字選取 columns-----
 xlWs.columns(3).select	'選取第3行
 xlWs.columns(i).select	'單選第i行

 xlWs.columns(i).columnwidth = 5	'第i行的欄寬=5
 xlWs.range("C:C,E:E,G:G").columnwidth = 5
 xlWs.columns(i).AutoFit	'第i行的欄寬=最適欄寛
 xlWs.columns("D:F").delete	'刪除行
 xlWs.range("C:C,E:E,G:G").delete	'刪除行

 -----選取 rows-----
 xlWs.rows(i).select	'選取單列
 xlWs.rows("2:6").select	'選取連續列
 xlWs.rows.select	'選取全部列 = 全選
 xlWs.range("3:3, 5:5, 8:8").select	'選取多列

 xlWs.rows(3).rowheight = 5	'列高
 xlWs.rows(3).insert	'插入列
 xlWs.rows(3).delete	'刪除列

 -----選取 cells-----
 xlWs.range("D4:D4").select	'選取單格
 xlWs.range("B2:H6").select	'選取範圍
 xlWs.range("D2:B5, F8:I9").select	'選取多個範圍

 xlWs.range("D4") = "TEST"	'儲存格內容
 xlWs.range("D4").font.name = "cambria"	'設定字型
 xlWs.range("D4").font.size = 24	'設定字體
 xlWs.range("D4").font.bold = true	'粗體
 xlWs.range("D4").font.color = vbblue	'設定文字顏色
 xlWs.range("D4").Interior.colorindex = 36	'設定背景顏色

 -----合併儲存格-----
 xlWs.range("E5:I6").mergecells = true	'合併儲存格
 tstring = "E" & i & ":" & "I" & j
 xlWs.range(tstring).mergecells = true	'合併儲存格

 -----儲存格對齊-----
 xlWs.range("D4").verticalalignment = 2	'上下對齊
 1=靠上 , 2=置中 , 3=靠下 , 4=垂直對齊??
 xlWs.range("D4").horizontalalignment = 1	'左右對齊
 1=一般 , 2=置左 , 3=置中 , 4=靠右 , 5=填滿 , 6=水平對齊? , 7=跨欄置中

 -----儲存格框線-----
 xlWs.range("D4").borders(n)	'框線方向
 n= 1:左, 2:右, 3:上, 4:下, 5:斜, 6:斜
 xlWs.range("D4").borders(4).color = 5
 xlWs.range("D4").borders(4).weight = 3	'框線粗細
 xlWs.range("D4").borders(4).linestyle = 1	'框線樣式
 線種類= 1,7:細實 2:細虛 4:一點虛 9:雙細實線
 xlWs.range("D4").borders(4).color = 6

 -----儲存格計算-----
 xlWs.range("I17").value = 20
 xlWs.range("I18").value = 30
 xlWs.Range("I19").Formula = xlWs.Range("I17") * xlWs.Range("I18") / 100
 xlWs.Range("I20").Formula = "=SUM(I17:I19)"

 -----加入註解-----
 xlWs.cells(n,1).AddComment
 xlWs.cells(n,1).Comment.visible = False
 xlWs.cells(n,1).Comment.text("有建BOM表,卻不計算BOM的成本")
 -----讀取註解,待測-----
 comment-text = xlWs.cells(n,1).Comment.text()
 comment-text = xlWs.cells(n,1).Comment.text

 -----列出 excel 字體顏色 color values-----
 for i = 1 to 56
 xlWs.cells(i + 3, 1).value = "value = " & i
 xlWs.cells(i + 3, 2).interior.colorindex = i
 next
*/