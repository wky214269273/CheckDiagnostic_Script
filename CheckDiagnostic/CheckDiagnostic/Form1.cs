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
             * 8:CheckDTCCodeHex(1,2,46)  显示码起始列位置，HEX起始列位置，DTC数量
             * ---->Error: DTCCode: XXXXXX invalid length
             * ---->Error: DTCHex: XXXXX invalid length
             * ---->Error: DTCCode: XXXXXXX invalid group (注意：只有UCPB)
             * ---->Error: DTCCode: XXXXXXXX & DTCHex: XXXXXX is inconsistent
             * ---->Error: DTCCode: XXXXXXX invalid Failure type
             * ---->Error: DTCCode: XXXXXXXX Duplicate
            */
            public string CaculatorSymbol;  //运算符
            public string StringBuffer;  //字符缓存
            public string StringCheck;  //校核字符
            public string M_NRC_Des;
            public string U_NRC_Des;
            public Int32 Position_X;
            public Int32 Position_Y;
            public Int32 DTC_Count;
        }
        public struct ScriptStep
        {
            public string Define_Text;
            public int StepCounts;
            public List<ScriptMode> m_Step;
        }
        public struct DTCInfo
        {
            public string DTCCode;
            public string DTCHex;
            public int _NO;  //标记序号位置
        }
        ScriptStep _Script_Step = new ScriptStep();  //宏步骤
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
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
                        m_Edit_Process.AppendText("---->Info: Open worksheet " + workSheet.Name + "\r\n");
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
                    else if (_Script_Step.m_Step[i].ScriptType == 8)
                    {
                        List<DTCInfo> _DTC = new List<DTCInfo>();
                        DTCInfo _DTC_Buffer = new DTCInfo();
                        for (int j = 1; j <= _Script_Step.m_Step[i].DTC_Count; ++j)  //提取内容
                        {
                            _DTC_Buffer.DTCCode = Convert.ToString(data[j, _Script_Step.m_Step[i].Position_X]).Trim();
                            _DTC_Buffer.DTCHex = Convert.ToString(data[j, _Script_Step.m_Step[i].Position_Y]).Trim();
                            _DTC_Buffer._NO = j;
                            _DTC.Add(_DTC_Buffer);
                        }
                        //第一步：检测DTCCode重复
                        List<string> lisDupValues2 = _DTC.GroupBy(x => x.DTCCode).Where(x => x.Count() > 1).Select(x => x.Key).ToList();
                        for (int j = 0; j < lisDupValues2.Count; ++j)
                        {
                            m_Edit_Process.AppendText("---->Error: DTCCode: " + lisDupValues2[j] + " Duplicate\r\n");
                        }
                        //剔除重复的，然后再执行下列步骤
                        for (int j = 0; j < _DTC.Count; ++j)  //循环次数
                        {
                            for (int k = _DTC.Count - 1; k > j; --k)  //比较次数
                            {
                                if (_DTC[j].DTCCode == _DTC[k].DTCCode)
                                {
                                    _DTC.RemoveAt(k);
                                }
                            }
                        }
                        //执行后续检测
                        for (int j = 0; j < _DTC.Count; ++j)
                        {
                            bool _Flag_Code = true;
                            bool _Flag_HEX = true;
                            //检测长度
                            if (_DTC[j].DTCCode.Length != 7)
                            {
                                m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid Length\r\n", _DTC[j]._NO));
                                _Flag_Code = false;
                            }
                            else
                            {
                                if (!CheckDTCInvalid(0, _DTC[j]))
                                {
                                    m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid Value\r\n", _DTC[j]._NO));
                                    _Flag_Code = false;
                                }
                            }
                            if (_DTC[j].DTCHex.Length != 6)
                            {
                                m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCHex: " + _DTC[j].DTCHex + " Invalid Length\r\n", _DTC[j]._NO));
                                _Flag_HEX = false;
                            }
                            else
                            {
                                if (!CheckDTCInvalid(1, _DTC[j]))
                                {
                                    m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCHex: " + _DTC[j].DTCHex + " Invalid Value\r\n", _DTC[j]._NO));
                                    _Flag_HEX = false;
                                }
                            }
                            //检测失效类型(目前仅检测ISO预留范围)
                            if (_Flag_Code)
                            {
                                //检测最后一个字节的失效类型
                                string _FailureType = _DTC[j].DTCCode.Substring(5, 2);  //取失效类型
                                if (_FailureType[0] == '0')
                                {
                                    if (_FailureType[1] >= 'A')
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '1')
                                {
                                    if (_FailureType[1] == '0')
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '2')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= 'A') && (_FailureType[1] <= 'E')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '3')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= 'B') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '4')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= 'C') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '5')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= '6') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '6')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= '9') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '7')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= 'C') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '8')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= '9') && (_FailureType[1] <= 'E')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] == '9')
                                {
                                    if ((_FailureType[1] == '0') || ((_FailureType[1] >= '9') && (_FailureType[1] <= 'F')))
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_FailureType[0] >= 'A' && _FailureType[0] <= 'E')
                                {
                                    m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " Invalid FailureType(ISO/SAE Reserved)\r\n", _DTC[j]._NO));
                                    continue;
                                }
                                if (_FailureType[0] == 'F')
                                {
                                    m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Warning: DTCCode: " + _DTC[j].DTCCode + " Use OEM defined FailureType\r\n", _DTC[j]._NO));
                                    continue;
                                }
                            }
                            //检测HEX及Code一致性
                            if (_Flag_Code && _Flag_HEX)  //只有都是有效值，才会核对该内容
                            {
                                //B
                                if (_DTC[j].DTCCode.Substring(0,2) == "B0")
                                {
                                    if (_DTC[j].DTCCode.Replace("B0","8") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "B1")
                                {
                                    if (_DTC[j].DTCCode.Replace("B1", "9") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "B2")
                                {
                                    if (_DTC[j].DTCCode.Replace("B2", "A") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "B3")
                                {
                                    if (_DTC[j].DTCCode.Replace("B3", "B") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                //C
                                if (_DTC[j].DTCCode.Substring(0, 2) == "C0")
                                {
                                    if (_DTC[j].DTCCode.Replace("C0", "4") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "C1")
                                {
                                    if (_DTC[j].DTCCode.Replace("C1", "5") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "C2")
                                {
                                    if (_DTC[j].DTCCode.Replace("C2", "6") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "C3")
                                {
                                    if (_DTC[j].DTCCode.Replace("C3", "7") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                //P
                                if (_DTC[j].DTCCode.Substring(0, 2) == "P0")
                                {
                                    if (_DTC[j].DTCCode.Replace("P0", "0") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "P1")
                                {
                                    if (_DTC[j].DTCCode.Replace("P1", "1") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "P2")
                                {
                                    if (_DTC[j].DTCCode.Replace("P2", "2") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "P3")
                                {
                                    if (_DTC[j].DTCCode.Replace("P3", "3") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                //U
                                if (_DTC[j].DTCCode.Substring(0, 2) == "U0")
                                {
                                    if (_DTC[j].DTCCode.Replace("U0", "C") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "U1")
                                {
                                    if (_DTC[j].DTCCode.Replace("U1", "D") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "U2")
                                {
                                    if (_DTC[j].DTCCode.Replace("U2", "E") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                                if (_DTC[j].DTCCode.Substring(0, 2) == "U3")
                                {
                                    if (_DTC[j].DTCCode.Replace("U3", "F") != _DTC[j].DTCHex)
                                    {
                                        m_Edit_Process.AppendText(string.Format("---->[NO.{0:D}]Error: DTCCode: " + _DTC[j].DTCCode + " && DTCHex: " + _DTC[j].DTCHex + " is inconsistent\r\n", _DTC[j]._NO));
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                m_Edit_Process.AppendText("---->Error:" + err.Message + "\r\n");
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
        public bool CheckDTCInvalid(int _Flag, DTCInfo _DTC_Check)  //0:DTCCode,1:DTCHex
        {
            if (_Flag == 0)
            {
                if (_DTC_Check.DTCCode[0] != 'C' &&
                    _DTC_Check.DTCCode[0] != 'B' &&
                    _DTC_Check.DTCCode[0] != 'U' &&
                    _DTC_Check.DTCCode[0] != 'P')
                {
                    return false;
                }
                if (_DTC_Check.DTCCode[1] < '0' || _DTC_Check.DTCCode[1] > '3')
                {
                    return false;
                }
                for (int i = 0;i < 5;++i)
                {
                    if ((!(_DTC_Check.DTCCode[2 + i] >= '0' && _DTC_Check.DTCCode[2 + i] <= '9')) &&
                        (!(_DTC_Check.DTCCode[2 + i] >= 'A' && _DTC_Check.DTCCode[2 + i] <= 'F')))
                    {
                        return false;
                    }
                }
                return true;
            }
            else if (_Flag == 1)
            {
                for (int i = 0; i < 6; ++i)
                {
                    if ((!(_DTC_Check.DTCHex[i] >= '0' && _DTC_Check.DTCHex[i] <= '9')) &&
                        (!(_DTC_Check.DTCHex[i] >= 'A' && _DTC_Check.DTCHex[i] <= 'F')))
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                return false;
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
            Regex ScriptCheckDTC = new Regex("^CheckDTCCodeHex\\((.+),(.+),(.+)\\);");  //检查DTC
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
                    else if (ScriptCheckDTC.IsMatch(ScriptLine[i]))
                    {
                        ScriptMode _Buffer = new ScriptMode();
                        match = ScriptCheckDTC.Match(ScriptLine[i]);
                        _Buffer.ScriptType = 8;
                        _Buffer.Position_Y = Convert.ToInt32(match.Groups[2].Value);
                        _Buffer.Position_X = Convert.ToInt32(match.Groups[1].Value);
                        _Buffer.DTC_Count = Convert.ToInt32(match.Groups[3].Value);
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
