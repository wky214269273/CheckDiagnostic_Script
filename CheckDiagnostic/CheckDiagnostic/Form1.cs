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
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace CheckDiagnostic
{

    public partial class Form1 : Form
    {
        public struct ScriptMode
        {
            public byte ScriptType;  //命令类型
            /*
             * 1:OpenExcel()
             * 2:OpenWorkSheet(3,range)
             * 3:CloseWorkSheet()
             * 4:CloseExcel()
             * 5:Show(xxxxx)
             * 6:CheckItem(1,1,== or !=,Yes,---->Error:$10 Sevice must be supported)
             * 7:CheckNRC(1,1,"M1-M2-M3","U1-U2-U3")
            */
            public string CaculatorSymbol;  //运算符
            public string StringBuffer;  //字符缓存
            public string StringCheck;  //校核字符
            public string M_NRC_Des;
            public string U_NRC_Des;
            public Int32 Position_X;
            public Int32 Position_Y;
        }
        public struct ScriptStep
        {
            public string Define_Text;
            public int StepCounts;
            public List<ScriptMode> m_Step;
        }
        ScriptStep _Script_Step = new ScriptStep();  //宏步骤
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)  //脚本文件
        {
            m_Edit_Process.Text = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "ECU脚本文件(txt)|*.txt";
            ofd.FileName = "ECU";
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            m_Edit_File_Script.Text = ofd.FileName;
            ReadScript(ofd.FileName, ref _Script_Step);
            MessageBox.Show(string.Format("成功载入脚本文件!一共{0:D}条有效脚本指令", _Script_Step.StepCounts), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //Script解析：核心
        public bool ExcuteScript(ScriptStep _Step)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = null;
            object oMissiong = Missing.Value;
            object[,] data = null;
            try
            {
                for (int i = 0; i < _Script_Step.StepCounts; ++i)
                {
                    if (_Script_Step.m_Step[i].ScriptType == 1)
                    {
                        workBook = app.Workbooks.Open(m_Edit_File_Excel.Text, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                        if (workBook == null)
                        {
                            return false;
                        }
                    }
                    else if (_Script_Step.m_Step[i].ScriptType == 2)
                    {
                        Worksheet workSheet = (Worksheet)workBook.Worksheets.Item[_Script_Step.m_Step[i].Position_X];
                        data = workSheet.Range[_Script_Step.m_Step[i].StringBuffer].Value2;
                    }
                    else if (_Script_Step.m_Step[i].ScriptType == 5)
                    {
                        if (_Script_Step.m_Step[i].StringBuffer != "#NULL")
                        {
                            m_Edit_Process.AppendText(_Script_Step.m_Step[i].StringBuffer + "\r\n");
                        }
                        else
                        {
                            m_Edit_Process.AppendText("\r\n");
                        }
                    }
                    else if (_Script_Step.m_Step[i].ScriptType == 6)
                    {
                        string m_Check = (_Script_Step.m_Step[i].StringCheck != "#NULL") ? _Script_Step.m_Step[i].StringCheck : string.Empty;
                        string m_Buffer = (_Script_Step.m_Step[i].StringBuffer != "#NULL") ? _Script_Step.m_Step[i].StringBuffer : string.Empty;
                        if (_Script_Step.m_Step[i].CaculatorSymbol == "==")
                        {
                            if (Convert.ToString(data[_Script_Step.m_Step[i].Position_Y, _Script_Step.m_Step[i].Position_X]) != m_Check)
                            {
                                m_Edit_Process.AppendText(m_Buffer + "\r\n");
                            }
                        }
                        else if (_Script_Step.m_Step[i].CaculatorSymbol == "!=")
                        {
                            if (Convert.ToString(data[_Script_Step.m_Step[i].Position_Y, _Script_Step.m_Step[i].Position_X]) == m_Check)
                            {
                                m_Edit_Process.AppendText(m_Buffer + "\r\n");
                            }
                        }
                        else
                        {
                            m_Edit_Process.AppendText(string.Format("ScriptStep:{0:D} Error:Invalid CaculatorSymbol!\r\n", i));
                            return false;
                        }
                    }
                    else if (_Script_Step.m_Step[i].ScriptType == 7)
                    {
                        string m_Check = (_Script_Step.m_Step[i].StringCheck != "#NULL") ? _Script_Step.m_Step[i].StringCheck : string.Empty;
                        string m_Buffer = (_Script_Step.m_Step[i].StringBuffer != "#NULL") ? _Script_Step.m_Step[i].StringBuffer : string.Empty;
                        string m_NRCDes = (_Script_Step.m_Step[i].M_NRC_Des != "#NULL") ? _Script_Step.m_Step[i].StringBuffer : string.Empty;
                        string u_NRCDes = (_Script_Step.m_Step[i].U_NRC_Des != "#NULL") ? _Script_Step.m_Step[i].StringBuffer : string.Empty;
                        bool _Flag = false;
                        string _BufferText = _Script_Step.m_Step[i].M_NRC_Des;
                        string _NRC = Convert.ToString(data[_Script_Step.m_Step[i].Position_Y, _Script_Step.m_Step[i].Position_X]);
                        //首先分割必须支持的NRC
                        if (m_Buffer != string.Empty)
                        {
                            string[] M_NRC = m_Buffer.Split('-');
                            for (int j = 0; j < M_NRC.Length; ++j)
                            {
                                if (_NRC.Contains(M_NRC[j]) == false)
                                {
                                    _BufferText = _BufferText + " " + M_NRC[j];
                                    _Flag = true;
                                }
                            }
                            if (_Flag)
                            {
                                m_Edit_Process.AppendText(_BufferText + "\r\n");
                            }
                        }
                        //分割不需要支持的NRC
                        if (m_Check != string.Empty)
                        {
                            string[] U_NRC = m_Check.Split('-');
                            string _BufferTextU = _Script_Step.m_Step[i].U_NRC_Des;
                            _Flag = false;
                            for (int j = 0; j < U_NRC.Length; ++j)
                            {
                                if (_NRC.Contains(U_NRC[j]) == true)
                                {
                                    _BufferTextU = _BufferTextU + " " + U_NRC[j];
                                    _Flag = true;
                                }
                            }
                            if (_Flag)
                            {
                                m_Edit_Process.AppendText(_BufferTextU + "\r\n");
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                m_Edit_Process.AppendText("---->Error:" + err.Message);
                return false;
            }
            finally
            {
                //COM组件方式调用完记得释放资源
                if (workBook != null)
                {
                    workBook.Close(false, oMissiong, oMissiong);
                    Marshal.ReleaseComObject(workBook);
                    app.Workbooks.Close();
                    app.Quit();
                    PublicMethod.Kill(app);
                }
            }
        }
        public void ReadScript(string ScriptPath, ref ScriptStep _Step)
        {
            _Step.Define_Text = string.Empty;
            _Step.StepCounts = 0;
            _Step.m_Step = new List<ScriptMode>();
            string[] ScriptLine = File.ReadAllLines(ScriptPath);  //读取每行
            Match match;
            Regex ScriptDefine = new Regex("^Define\\(\"(.+)\"\\);");  //描述字符
            Regex ScriptShow = new Regex("^Show\\(\"(.+)\"\\);");  //显示字符
            Regex ScriptOpenExcel = new Regex("^OpenExcel\\(\\);");  //打开Excel
            Regex ScriptOpenWorkSheet = new Regex("^OpenWorkSheet\\((.+),\"(.+)\"\\);");  //打开工作页
            Regex ScriptCloseExcel = new Regex("^CloseExcel\\(\\);");  //关闭Excel
            Regex ScriptCloseWorkSheet = new Regex("^CloseWorkSheet\\(\\);");  //关闭工作页
            Regex ScriptCheckItem = new Regex("^CheckItem\\((.+),(.+),(.+),\"(.+)\",\"(.+)\"\\);");  //打开工作页
            Regex ScriptCheckNRC = new Regex("^CheckNRC\\((.+),(.+),\"(.+)\",\"(.+)\",\"(.+)\",\"(.+)\"\\);");  //打开工作页
            for (int i = 0; i < ScriptLine.Length; ++i)  //获取有效的行数，确定步骤数目
            {
                if (ScriptLine[i] != string.Empty)
                {
                    if (ScriptLine[i].Substring(0, 2) == "//")
                    {
                        //跳过注释行
                        continue;
                    }
                    else if (ScriptDefine.IsMatch(ScriptLine[i]))
                    {
                        match = ScriptDefine.Match(ScriptLine[i]);
                        if (match.Groups[1].Value != "#NULL")
                        {
                            _Step.Define_Text = _Step.Define_Text + match.Groups[1].Value + "\r\n";
                        }
                        else
                        {
                            _Step.Define_Text = _Step.Define_Text + "\r\n";
                        }
                        continue;
                    }
                    else if (ScriptCheckItem.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptCheckItem.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 6;
                        _Buffer.Position_Y = Convert.ToInt32(match.Groups[1].Value);
                        _Buffer.Position_X = Convert.ToInt32(match.Groups[2].Value);
                        _Buffer.CaculatorSymbol = match.Groups[3].Value;
                        _Buffer.StringCheck = match.Groups[4].Value;
                        _Buffer.StringBuffer = match.Groups[5].Value;
                        _Step.m_Step.Add(_Buffer);
                        _Step.StepCounts++;
                        continue;
                    }
                    else if (ScriptShow.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptShow.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 5;
                        _Buffer.StringBuffer = match.Groups[1].Value;
                        _Step.m_Step.Add(_Buffer);
                        _Step.StepCounts++;
                        continue;
                    }
                    else if (ScriptOpenExcel.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptOpenExcel.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 1;
                        _Step.m_Step.Add(_Buffer);
                        _Step.StepCounts++;
                        continue;
                    }
                    else if (ScriptOpenWorkSheet.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptOpenWorkSheet.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 2;
                        _Buffer.Position_X = Convert.ToInt32(match.Groups[1].Value);
                        _Buffer.StringBuffer = match.Groups[2].Value;
                        _Step.m_Step.Add(_Buffer);
                        _Step.StepCounts++;
                        continue;
                    }
                    else if (ScriptCheckNRC.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptCheckNRC.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 7;
                        _Buffer.Position_Y = Convert.ToInt32(match.Groups[1].Value);
                        _Buffer.Position_X = Convert.ToInt32(match.Groups[2].Value);
                        _Buffer.StringBuffer = match.Groups[3].Value;  //支持NRC
                        _Buffer.StringCheck = match.Groups[4].Value;  //不支持NRC
                        _Buffer.M_NRC_Des = match.Groups[5].Value;
                        _Buffer.U_NRC_Des = match.Groups[6].Value;
                        _Step.m_Step.Add(_Buffer);
                        _Step.StepCounts++;
                        continue;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            m_Edit_Process.AppendText(_Step.Define_Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Thread NewRecord = new Thread(ThreadCheckProcess);
            NewRecord.IsBackground = true;
            NewRecord.Start();
        }
        void ThreadCheckProcess()
        {
            if (ExcuteScript(_Script_Step))
            {
                MessageBox.Show("脚本检查结束，执行成功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("脚本检查结束，执行失败!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            m_Edit_Process.Text = string.Empty;
            ReadScript(m_Edit_File_Script.Text, ref _Script_Step);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "ECU诊断调查问卷(xlsx)|*.xlsx";
            ofd.FileName = "ECU";
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            m_Edit_File_Excel.Text = ofd.FileName;
        }
    }
    public class PublicMethod
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);//得到这个句柄，具体作用是得到这块内存入口 

            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }

    }
}
