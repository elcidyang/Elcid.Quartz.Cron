using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Authentication.ExtendedProtection;
using System.Windows.Forms;
using Quartz;
using Quartz.Spi;

namespace Elcid.Quartz.Cron
{
    public partial class QuartzCronForm : Form
    {
        readonly int controlWidth = 48;
        readonly int pointY = 135;
        readonly int pointX = 35;
        public QuartzCronForm()
        {
            InitializeComponent();
            CreateControls();
            InitControlValue();
            BindControlEvent(Controls);
            //生成cron
            txtExpression.Text = GenerateCron();
            ReloadRunDetail(txtExpression.Text);
        }

        public void InitControlValue()
        {
            var currentYear = DateTime.Now.Year;
            nudYearFrequencyBaseNum.Minimum = currentYear;
            nudYearFrequencyBaseNum.Maximum = currentYear + 100;
            nudYearFrequencyBaseNum.Value = currentYear;
            nudYearCycleBegin.Minimum = currentYear;
            nudYearCycleBegin.Maximum = currentYear + 100;
            nudYearCycleBegin.Value = currentYear;
            nudYearCycleEnd.Minimum = currentYear;
            nudYearCycleEnd.Maximum = currentYear + 100;
            nudYearCycleEnd.Value = currentYear;
        }

        #region 生成控件
        private void CreateControls()
        {
            CreateSecondControls();
            CreateMinuteControls();
            CreateHourControls();
            CreateDayControls();
            CreateMonthControls();
            CreateWeekControls();
        }

        private void CreateSecondControls()
        {
            var xStep = 0;
            var yStep = 0;
            for (var i = 0; i < 60; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = i.ToString(),
                    Name = "chkSecond" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabSecond.Controls.Add(chkTime);
                xStep++;
            }
        }

        private void CreateMinuteControls()
        {
            var xStep = 0;
            var yStep = 0;
            for (var i = 0; i < 60; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = i.ToString(),
                    Name = "chkMinute" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabMinute.Controls.Add(chkTime);
                xStep++;
            }
        }

        private void CreateHourControls()
        {
            var xStep = 0;
            var yStep = 0;
            for (var i = 0; i < 24; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = i.ToString(),
                    Name = "chkHour" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabHour.Controls.Add(chkTime);
                xStep++;
            }
        }

        private void CreateDayControls()
        {
            var xStep = 0;
            var yStep = 60;
            for (var i = 0; i < 31; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = (i + 1).ToString(),
                    Name = "chkDay" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabDay.Controls.Add(chkTime);
                xStep++;
            }
        }

        private void CreateMonthControls()
        {
            var xStep = 0;
            var yStep = 0;
            for (var i = 0; i < 12; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = (i + 1).ToString(),
                    Name = "chkMonth" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabMonth.Controls.Add(chkTime);
                xStep++;
            }
        }

        private void CreateWeekControls()
        {
            var xStep = 0;
            var yStep = 70;
            for (var i = 0; i < 7; i++)
            {
                if ((i + 1) % 11 == 0)
                {
                    xStep = 0;
                    yStep += 25;
                }
                var chkTime = new CheckBox
                {
                    Text = (i + 1).ToString(),
                    Name = "chkWeek" + i,
                    Size = new Size(controlWidth, 20),
                    Location = new Point(controlWidth * xStep + pointX, pointY + yStep)
                };
                tabWeek.Controls.Add(chkTime);
                xStep++;
            }
        }

        #endregion

        /// <summary>
        /// 为控件绑定事件
        /// </summary>
        private void BindControlEvent(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                if (control.HasChildren)
                {
                    BindControlEvent(control.Controls);
                }
                else
                {
                    //checkbox事件
                    var chk = control as CheckBox;
                    if (chk != null)
                    {
                        chk.CheckedChanged += checkBox_CheckedChanged;
                        chk.CheckedChanged += GenerateExpressionAndShowDetail;
                    }
                    //numericUpDown事件
                    var nud = control.Parent as NumericUpDown;
                    if (nud != null)
                    {
                        nud.TextChanged += numericUpDown_ValueChanged;
                        nud.MouseDown += numericUpDown_ValueChanged;
                        nud.MouseUp += numericUpDown_ValueChanged;
                        nud.TextChanged += GenerateExpressionAndShowDetail;
                        nud.MouseDown += GenerateExpressionAndShowDetail;
                        nud.MouseUp += GenerateExpressionAndShowDetail;
                    }
                    //radio事件
                    var rdo = control as RadioButton;
                    if (rdo != null)
                    {
                        rdo.CheckedChanged += radioButton_CheckedChanged;
                        rdo.CheckedChanged += GenerateExpressionAndShowDetail;
                    }
                }
            }
        }

        /// <summary>
        ///通过txtCron获取完整表达式
        /// </summary>
        /// <returns></returns>
        private string GenerateCron()
        {
            var strCron = string.Empty;
            strCron += txtSecondCron.Text + " ";
            strCron += txtMinuteCron.Text + " ";
            strCron += txtHourCron.Text + " ";
            strCron += txtDayCron.Text + " ";
            strCron += txtMonthCron.Text + " ";
            strCron += txtWeekCron.Text + " ";
            strCron += txtYearCron.Text;
            return strCron;
        }

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            var chk = (CheckBox)sender;
            var currentTab = chk.Parent;
            //设置容器中相应radio选中
            if (chk.Checked)
            {
                foreach (Control control in currentTab.Controls)
                {
                    if (control.Name.Contains("rdoAppoint"))
                    {
                        ((RadioButton)control).Checked = true;
                    }
                }
            }
            var dic = GetCurrentCron(currentTab, GetCurrentCheckBoxValue);
            foreach (var kvp in dic)
            {
                LoadCronField(kvp.Key, kvp.Value);
            }
        }

        private void GenerateExpressionAndShowDetail(object sender, EventArgs e)
        {
            //生成cron
            txtExpression.Text = GenerateCron();
            ReloadRunDetail(txtExpression.Text);
        }

        private void numericUpDown_ValueChanged(object sender, EventArgs e)
        {
            var nud = (NumericUpDown)sender;
            var currentTab = nud.Parent;
            foreach (Control control in currentTab.Controls)
            {
                if (nud.Name.Contains("Cycle") && control.Name.Contains("rdoCycle"))
                {
                    ((RadioButton)control).Checked = true;
                    var dic = GetCurrentCron(currentTab, GetCurrentCycleRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                if (nud.Name.Contains("Frequency") && control.Name.Contains("rdoFrequency"))
                {
                    ((RadioButton)control).Checked = true;
                    var dic = GetCurrentCron(currentTab, GetCurrentFrequencyRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                if (nud.Name.Contains("Last") && control.Name.Contains("rdoLast"))
                {
                    ((RadioButton)control).Checked = true;
                    var dic = GetCurrentCron(currentTab, GetCurrentLastRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                if (nud.Name.Contains("Special") && control.Name.Contains("rdoSpecial"))
                {
                    ((RadioButton)control).Checked = true;
                    var dic = GetCurrentCron(currentTab, GetCurrentSpecialRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                if (nud.Name.Contains("Rencent") && control.Name.Contains("rdoRencent"))
                {
                    ((RadioButton)control).Checked = true;
                    var dic = GetCurrentCron(currentTab, s => nudDayRencent.Value + "W");
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }

            }
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            var rdo = (RadioButton)sender;
            var currentTab = rdo.Parent;
            if (rdo.Checked)
            {
                //判断指定rdo下的checkbox是否选中如果没有选中则默认选择第一个
                if (rdo.Name.Contains("rdoAppoint"))
                {
                    var isHaveChecked = false;
                    var firstCheckBox = new CheckBox();
                    var count = 0;
                    foreach (Control control in currentTab.Controls)
                    {
                        if (control.Name.Contains("chk"))
                        {
                            var chk = (CheckBox)control;
                            if (count == 0)
                            {
                                firstCheckBox = chk;
                            }
                            if (chk.Checked)
                            {
                                isHaveChecked = chk.Checked;
                            }
                            count++;
                        }
                    }
                    if (!isHaveChecked)
                    {
                        firstCheckBox.Checked = true;
                    }
                    var dic = GetCurrentCron(currentTab, GetCurrentCheckBoxValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoEvery值
                if (rdo.Name.Contains("Every"))
                {
                    var dic = GetCurrentCron(currentTab, s => "*");
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoNotSpecify值
                if (rdo.Name.Contains("NotSpecify"))
                {
                    var dic = rdo.Name.Contains("Year") ? GetCurrentCron(currentTab, s => " ") : GetCurrentCron(currentTab, s => "?");
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoCycle值
                if (rdo.Name.Contains("Cycle"))
                {
                    var dic = GetCurrentCron(currentTab, GetCurrentCycleRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoFrequency值
                if (rdo.Name.Contains("Frequency"))
                {
                    var dic = GetCurrentCron(currentTab, GetCurrentFrequencyRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoLast值
                if (rdo.Name.Contains("Last"))
                {
                    var dic = GetCurrentCron(currentTab, GetCurrentLastRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoSpecial值
                if (rdo.Name.Contains("Special"))
                {
                    var dic = GetCurrentCron(currentTab, GetCurrentSpecialRadioValue);
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //rdoRencent值
                if (rdo.Name.Contains("Rencent"))
                {
                    var dic = GetCurrentCron(currentTab, s => nudDayRencent.Value + "W");
                    foreach (var kvp in dic)
                    {
                        LoadCronField(kvp.Key, kvp.Value);
                    }
                }
                //星期与日互斥 检测是否点击星期radiobutton
                if (rdo.Name.Contains("Week") && !rdo.Name.Contains("NotSpecify"))
                {
                    //设置日期rdo不选中及表达式变为?
                    ClearCheckedTabDayRadioStatus();
                    LoadCronField("txtDayCron", "?");
                }
                //检测是否点击日期radiobutton
                if (rdo.Name.Contains("Day") && !rdo.Name.Contains("NotSpecify"))
                {
                    //设置星期rdo为不指定及表达式变为?
                    SetTabWeekRadioNotSpecifyChecked();
                    LoadCronField("txtWeekCron", "?");
                }
            }
        }

        /// <summary>
        /// 将day下面的radiobutton置为不选
        /// </summary>
        private void ClearCheckedTabDayRadioStatus()
        {
            foreach (Control control in tabDay.Controls)
            {
                if (!control.Name.Contains("rdo")) continue;
                var rdo = (RadioButton)control;
                if (rdo.Checked)
                {
                    rdo.Checked = false;
                }
            }
        }

        /// <summary>
        /// 设置星期为不指定
        /// </summary>
        private void SetTabWeekRadioNotSpecifyChecked()
        {
            foreach (Control control in tabWeek.Controls)
            {
                if (!control.Name.Contains("rdo")) continue;
                var rdo = (RadioButton)control;
                if (rdo.Checked)
                {
                    rdo.Checked = false;
                }
                if (rdo.Name.Contains("NotSpecify"))
                {
                    rdo.Checked = true;
                }
            }
        }

        private Dictionary<string, string> GetCurrentCron(Control currentTab, Func<Control, string> method)
        {
            var dic = new Dictionary<string, string>();
            var controlName = string.Empty;
            switch (currentTab.Name)
            {
                case "tabSecond":
                    controlName = "txtSecondCron";
                    break;
                case "tabMinute":
                    controlName = "txtMinuteCron";
                    break;
                case "tabHour":
                    controlName = "txtHourCron";
                    break;
                case "tabDay":
                    controlName = "txtDayCron";
                    break;
                case "tabMonth":
                    controlName = "txtMonthCron";
                    break;
                case "tabWeek":
                    controlName = "txtWeekCron";
                    break;
                case "tabYear":
                    controlName = "txtYearCron";
                    break;
            }
            dic.Add(controlName, method(currentTab));
            return dic;
        }

        private string GetCurrentCheckBoxValue(Control currentTab)
        {
            var strCron = string.Empty;
            foreach (Control control in currentTab.Controls)
            {
                if (!control.Name.Contains("chk")) continue;
                var chk = (CheckBox)control;
                if (chk.Checked)
                {
                    strCron += chk.Text + ",";
                }
            }
            strCron = string.IsNullOrEmpty(strCron) ? "?" : strCron.Substring(0, strCron.Length - 1);
            return strCron;
        }

        private string GetCurrentCycleRadioValue(Control currentTab)
        {
            var cycleBegin = string.Empty;
            var cycleEnd = string.Empty;
            foreach (Control control in currentTab.Controls)
            {
                if (control.Name.Contains("CycleBegin"))
                {
                    cycleBegin = control.Text;
                }
                if (control.Name.Contains("CycleEnd"))
                {
                    cycleEnd = control.Text;
                }
            }
            return cycleBegin + "-" + cycleEnd;
        }

        private string GetCurrentFrequencyRadioValue(Control currentTab)
        {
            var baseNum = string.Empty;
            var intervalNum = string.Empty;
            foreach (Control control in currentTab.Controls)
            {
                if (control.Name.Contains("FrequencyBaseNum"))
                {
                    baseNum = control.Text;
                }
                if (control.Name.Contains("IntervalNum"))
                {
                    intervalNum = control.Text;
                }
            }
            return baseNum + "/" + intervalNum;
        }

        private string GetCurrentLastRadioValue(Control currentTab)
        {
            var lastCron = "L";
            foreach (Control control in currentTab.Controls)
            {
                if (control.Name == "nudWeekLastDay")
                {
                    lastCron = control.Text + lastCron;
                }
            }
            return lastCron;
        }

        private string GetCurrentSpecialRadioValue(Control currentTab)
        {
            var baseNum = string.Empty;
            var day = string.Empty;
            foreach (Control control in currentTab.Controls)
            {
                if (control.Name.Contains("SpecialBaseNum"))
                {
                    baseNum = control.Text;
                }
                if (control.Name.Contains("SpecialDay"))
                {
                    day = control.Text;
                }
            }
            if (currentTab.Name.Contains("Week"))
            {
                return day + "#" + baseNum;
            }
            return baseNum + "#" + day;
        }

        /// <summary>
        ///反解析到ui
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReverse_Click(object sender, EventArgs e)
        {
            var cronStr = txtExpression.Text;
            if (string.IsNullOrEmpty(cronStr))
            {
                return;
            }
            ReloadRunDetail(cronStr);
            var crons = ReverseCron(cronStr);
            if (crons != null)
            {
                try
                {
                    ReverseUi(crons);
                }
                catch (Exception ex)
                {
                    txtRunDetail.Text = ex.Message;
                }
            }
        }

        /// <summary>
        /// 解析到表达式字段
        /// </summary>
        /// <param name="cronStr"></param>
        /// <returns></returns>
        private List<string> ReverseCron(string cronStr)
        {
            var crons = cronStr.Split(' ').ToList();
            if (crons.Count >= 6 && crons.Count <= 7)
            {
                txtSecondCron.Text = crons[0];
                txtMinuteCron.Text = crons[1];
                txtHourCron.Text = crons[2];
                txtDayCron.Text = crons[3];
                txtMonthCron.Text = crons[4];
                txtWeekCron.Text = crons[5];
                txtYearCron.Text = string.Empty;
                if (crons.Count == 7)
                {
                    txtYearCron.Text = crons[6];
                }
            }
            else
            {
                return null;
            }
            return crons;
        }

        private void ReloadRunDetail(string cronStr)
        {
            txtRunDetail.Text = String.Empty;
            try
            {
                var timeList = GetTaskeFireTime(cronStr);
                foreach (var time in timeList)
                {
                    var formatTime = Convert.ToDateTime(time).ToString("yyyy-MM-dd HH:mm:ss dddd");
                    txtRunDetail.Text += formatTime + "\r\n";
                }
            }
            catch (Exception ex)
            {
                txtRunDetail.Text = "解析错误：" + ex.Message;
            }
        }

        /// <summary>
        /// 反解析到ui
        /// </summary>
        /// <param name="crons"></param>
        private void ReverseUi(List<string> crons)
        {
            var tabPages = tab.TabPages;
            for (var i = 0; i < crons.Count; i++)
            {
                var tabControls = tabPages[i].Controls;
                //按表达式特征设置控件状态
                if (crons[i] == "*")
                {
                    foreach (Control control in tabControls)
                    {
                        if (!control.Name.Contains("rdo")) continue;
                        var rdo = (RadioButton)control;
                        rdo.Checked = rdo.Name.Contains("Every");
                    }
                }
                if (crons[i].Contains(",") || Common.IsInt(crons[i]))
                {
                    var nums = crons[i].Split(',');
                    foreach (Control control in tabControls)
                    {
                        if (!control.Name.Contains("chk")) continue;
                        var chk = (CheckBox)control;
                        chk.Checked = false;
                        foreach (var num in nums)
                        {
                            if (chk.Text == num)
                            {
                                chk.Checked = true;
                            }
                        }
                    }
                }
                if (crons[i].Contains("-"))
                {
                    var nums = crons[i].Split('-');
                    if (nums.Length != 2)
                    {
                        throw new Exception("表达式格式错误解析\"-\"时失败");
                    }
                    foreach (Control control in tabControls)
                    {
                        if (control.Name.Contains("nud") && control.Name.Contains("Cycle"))
                        {
                            var nud = (NumericUpDown)control;
                            if (nud.Name.Contains("Begin"))
                            {
                                nud.Value = Convert.ToDecimal(nums[0]);
                            }
                            if (nud.Name.Contains("End"))
                            {
                                nud.Value = Convert.ToDecimal(nums[1]);
                            }
                        }
                    }
                }
                if (crons[i].Contains("/"))
                {
                    var nums = crons[i].Split('/');
                    if (nums.Length != 2)
                    {
                        throw new Exception("表达式格式错误解析\"/\"时失败");
                    }
                    foreach (Control control in tabControls)
                    {
                        if (control.Name.Contains("nud") && control.Name.Contains("Frequency"))
                        {
                            var nud = (NumericUpDown)control;
                            if (nud.Name.Contains("BaseNum"))
                            {
                                nud.Value = Convert.ToDecimal(nums[0]);
                            }
                            if (nud.Name.Contains("IntervalNum"))
                            {
                                nud.Value = Convert.ToDecimal(nums[1]);
                            }
                        }
                    }
                }
                if (crons[i] == "?")
                {
                    foreach (Control control in tabControls)
                    {
                        if (!control.Name.Contains("rdo")) continue;
                        var rdo = (RadioButton)control;
                        rdo.Checked = rdo.Name.Contains("NotSpecify");
                    }
                }
                //特殊处理L C忽略
                if (crons[i].Contains("L"))
                {
                    var chars = crons[i].ToCharArray();
                    foreach (Control control in tabControls)
                    {
                        if (!control.Name.Contains("Last")) continue;
                        if (control.Name.Contains("nud"))
                        {
                            var nud = (NumericUpDown)control;
                            nud.Value = Common.IsInt(chars[0].ToString()) ? Convert.ToDecimal(chars[0].ToString()) : 1;
                        }
                        else
                        {
                            if (control.Name.Contains("rdo"))
                            {
                                var rdo = (RadioButton)control;
                                rdo.Checked = rdo.Name.Contains("Last");
                            }
                        }
                    }
                }
                //特殊处理W
                if (crons[i].Contains("W"))
                {
                    var rencentNum = crons[i].Substring(0, crons[i].Length - 1);
                    if (rencentNum.Length > 0)
                    {
                        foreach (Control control in tabControls)
                        {
                            if (!(control.Name.Contains("Rencent") && control.Name.Contains("nud"))) continue;
                            var nud = (NumericUpDown) control;
                            nud.Value =Convert.ToDecimal(rencentNum);
                        }
                    }
                }
                //特殊处理# 星期专用
                if (crons[i].Contains("#"))
                {
                    var nums = crons[i].Split('#');
                    if (nums.Length == 2)
                    {
                        foreach (Control control in tabControls)
                        {
                            if (!control.Name.Contains("undWeekSpecial")) continue;
                            var nud = (NumericUpDown)control;
                            if (nud.Name.Contains("BaseNum"))
                            {
                                nud.Value = Convert.ToDecimal(nums[1]);
                            }
                            if (nud.Name.Contains("Day"))
                            {
                                nud.Value = Convert.ToDecimal(nums[0]);
                            }
                        }
                    }
                }

            }
            //年可为空
            if (crons.Count < 7)
            {
                var tabControls = tabPages[6].Controls;
                foreach (Control control in tabControls)
                {
                    if (!control.Name.Contains("rdo")) continue;
                    var rdo = (RadioButton)control;
                    rdo.Checked = rdo.Name.Contains("NotSpecify");
                }
            }

        }

        /// <summary>
        /// 加载表达式字段
        /// </summary>
        private void LoadCronField(string controlName, string value)
        {
            var controls = this.Controls.Find(controlName, true);
            foreach (var control in controls)
            {
                control.Text = value;
            }
        }

        /// <summary>
        /// 获取任务在未来周期内哪些时间会运行
        /// </summary>
        /// <param name="cronExpression">Cron表达式</param>
        /// <param name="numTimes">运行次数</param>
        /// <returns>运行时间段</returns>
        public static List<string> GetTaskeFireTime(string cronExpression, int numTimes = 50)
        {
            if (numTimes < 0)
            {
                throw new Exception("参数numTimes值大于等于0");
            }
            //时间表达式
            ITrigger trigger = TriggerBuilder.Create().WithCronSchedule(cronExpression).Build();
            IList<DateTimeOffset> dates = TriggerUtils.ComputeFireTimes(trigger as IOperableTrigger, null, numTimes);
            List<string> list = new List<string>();
            foreach (DateTimeOffset dtf in dates)
            {
                list.Add(TimeZoneInfo.ConvertTimeFromUtc(dtf.DateTime, TimeZoneInfo.Local).ToString());
            }
            return list;
        }

        /// <summary>
        /// 复制表达式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtExpression.Text))
            {
                Clipboard.SetDataObject(txtExpression.Text);
            }
        }
    }
}
