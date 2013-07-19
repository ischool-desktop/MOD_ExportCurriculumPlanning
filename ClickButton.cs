using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using System.Xml;
using FISCA.Presentation;
using SHSchool.Data;

namespace ExportCurriculumPlanning
{
    public partial class ClickButton : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();
        BackgroundWorker BGW_Load = new BackgroundWorker();
        Workbook WBook { get; set; }

        public ClickButton()
        {
            InitializeComponent();

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);

            this.Text = "匯出課程規劃表(資料載入中...)";
            btnStartPrint.Enabled = false;
            BGW_Load.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_Load_RunWorkerCompleted);
            BGW_Load.DoWork += new DoWorkEventHandler(BGW_Load_DoWork);
            BGW_Load.RunWorkerAsync();
        }

        void BGW_Load_DoWork(object sender, DoWorkEventArgs e)
        {
            tool.Run();

            List<SHProgramPlanRecord> list = SHProgramPlan.SelectAll();
            e.Result = list;
        }

        void BGW_Load_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<SHProgramPlanRecord> list = (List<SHProgramPlanRecord>)e.Result;
            foreach (SHProgramPlanRecord LV1 in list)
            {
                ListViewItem item = new ListViewItem(LV1.Name);
                item.Tag = LV1;
                listView1.Items.Add(item);
            }

            this.Text = "匯出課程規劃表";
            btnStartPrint.Enabled = true;
        }

        private void btnStartPrint_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                btnStartPrint.Enabled = false;
                MotherForm.SetStatusBarMessage("開始列印課程規劃表...");

                List<SHProgramPlanRecord> list = new List<SHProgramPlanRecord>();
                foreach (ListViewItem item in listView1.Items)
                {
                    //未選擇,則不進行列印
                    if (item.Checked == false) continue;

                    SHProgramPlanRecord var = (SHProgramPlanRecord)item.Tag;
                    list.Add(var);
                }
                BGW.RunWorkerAsync(list);
            }
            else
            {
                MsgBox.Show("系統忙碌中,請稍後...");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //List<K12.Data.ClassRecord> Class = K12.Data.Class.SelectAll();

            WBook = new Workbook();
            WBook.Open(new MemoryStream(Properties.Resources.匯出課程規畫樣版));

            List<SHProgramPlanRecord> list = (List<SHProgramPlanRecord>)e.Argument;
            foreach (SHProgramPlanRecord var in list)
            {

                #region 處理科目級別名稱

                Dictionary<string, List<string>> SubjectNameDic = new Dictionary<string, List<string>>();
                foreach (K12.Data.ProgramSubject Xmlvar in var.Subjects)
                {
                    if (!SubjectNameDic.ContainsKey(Xmlvar.SubjectName))
                        SubjectNameDic.Add(Xmlvar.SubjectName, new List<string>());

                    if (Xmlvar.Level.HasValue)
                        SubjectNameDic[Xmlvar.SubjectName].Add(POPstartLevel(Xmlvar.Level.Value.ToString()));
                }

                foreach (List<string> each in SubjectNameDic.Values)
                {
                    each.Sort(tool.SortSubjectLevel);
                }

                #endregion

                //資料定位
                Dictionary<int, decimal> AddCredit = new Dictionary<int, decimal>();
                AddCredit.Add(6, 0); //一上
                AddCredit.Add(7, 0); //一下
                AddCredit.Add(8, 0);
                AddCredit.Add(9, 0);
                AddCredit.Add(10, 0);
                AddCredit.Add(11, 0);
                AddCredit.Add(12, 0); //四上
                AddCredit.Add(13, 0); //四下

                //
                string WorkName = var.Name;
                WBook.Worksheets[WBook.Worksheets.AddCopy("樣版範本")].Name = WorkName;
                WBook.Worksheets[WorkName].Cells[0, 0].PutValue("課程規劃名稱：" + var.Name);

                Range range_1 = WBook.Worksheets["樣版範本"].Cells.CreateRange(2, 1, false);
                Range range_2 = WBook.Worksheets["小計範本"].Cells.CreateRange(3, 6, false);

                int CountCells = 1;
                int TestC = 0;
                int Smtp = 0;

                int Count校訂必修 = 0;
                decimal Count校訂必修學分 = 0;
                int Count校訂必修不評分 = 0;

                int Count校訂選修 = 0;
                decimal Count校訂選修學分 = 0;
                int Count校訂選修不評分 = 0;

                int Count部訂必修 = 0;
                decimal Count部訂必修學分 = 0;
                int Count部訂必修不評分 = 0;

                //每一張報表內的每一行
                foreach (K12.Data.ProgramSubject Xmlvar in var.Subjects)
                {
                    //複製上方樣版範本
                    CountCells = CountCells + Xmlvar.RowIndex;

                    //當第一行出現 或 RowIndex不是0時 Copy行的樣式
                    if (TestC == 0 || TestC != Xmlvar.RowIndex)
                    {
                        WBook.Worksheets[WorkName].Cells.CreateRange(CountCells, 1, false).Copy(range_1);
                    }

                    WBook.Worksheets[WorkName].Cells[CountCells, 0].PutValue(Xmlvar.Domain); //領域
                    WBook.Worksheets[WorkName].Cells[CountCells, 1].PutValue(Xmlvar.Entry); //分項

                    if (SubjectNameDic[Xmlvar.SubjectName].Count == 0)
                        WBook.Worksheets[WorkName].Cells[CountCells, 2].PutValue(Xmlvar.SubjectName); //科目名稱
                    else
                        WBook.Worksheets[WorkName].Cells[CountCells, 2].PutValue(Xmlvar.SubjectName + "(" + string.Join("，", SubjectNameDic[Xmlvar.SubjectName]) + ")");

                    WBook.Worksheets[WorkName].Cells[CountCells, 3].PutValue(Xmlvar.RequiredBy); //校部定
                    WBook.Worksheets[WorkName].Cells[CountCells, 4].PutValue(Xmlvar.Required ? "必修" : "選修"); //必選修

                    string _StartLevel = Xmlvar.StartLevel.HasValue ? Xmlvar.StartLevel.Value.ToString() : "";
                    WBook.Worksheets[WorkName].Cells[CountCells, 5].PutValue(_StartLevel); //開始級別


                    //學分加總
                    Smtp = PopLev(Xmlvar.GradeYear.ToString(), Xmlvar.Semester.ToString());
                    if (Xmlvar.Credit.HasValue)
                        AddCredit[Smtp] += Xmlvar.Credit.Value; //學分加總
                    //學分數
                    WBook.Worksheets[WorkName].Cells[CountCells, Smtp].PutValue(Xmlvar.Credit.Value); //必選修

                    if (Xmlvar.NotIncludedInCredit)
                        WBook.Worksheets[WorkName].Cells[CountCells, 14].PutValue("★"); //不計學分

                    if (Xmlvar.NotIncludedInCalc)
                        WBook.Worksheets[WorkName].Cells[CountCells, 15].PutValue("★"); //不評分

                    CountCells = 1;
                    TestC = Xmlvar.RowIndex;
                    #region 小計

                    if (Xmlvar.RequiredBy == "校訂" && Xmlvar.Required)
                    {
                        #region 校訂,必修時
                        Count校訂必修++;
                        if (!Xmlvar.NotIncludedInCredit)
                        {
                            //進入此方法為要計入學分
                            if (Xmlvar.Credit.HasValue)
                                Count校訂必修學分 += Xmlvar.Credit.Value;
                        }
                        if (Xmlvar.NotIncludedInCalc)
                        {
                            Count校訂必修不評分++;
                        }
                        #endregion
                    }
                    else if (Xmlvar.RequiredBy == "校訂" && !Xmlvar.Required)
                    {
                        #region 校訂,選修時
                        Count校訂選修++;
                        if (!Xmlvar.NotIncludedInCredit)
                        {
                            if (Xmlvar.Credit.HasValue)
                                Count校訂選修學分 += Xmlvar.Credit.Value;
                        }
                        if (Xmlvar.NotIncludedInCalc)
                        {
                            Count校訂選修不評分++;
                        }
                        #endregion
                    }
                    else if (Xmlvar.RequiredBy == "部訂" && Xmlvar.Required)
                    {
                        #region 部訂,必修時
                        Count部訂必修++;
                        if (!Xmlvar.NotIncludedInCredit)
                        {
                            if (Xmlvar.Credit.HasValue)
                                Count部訂必修學分 += Xmlvar.Credit.Value;
                        }
                        if (Xmlvar.NotIncludedInCalc)
                        {
                            Count部訂必修不評分++;
                        }
                        #endregion
                    }

                    #endregion
                }


                #region 填入小計
                int MaxDataRowIndex = WBook.Worksheets[WorkName].Cells.MaxDataRow + 1;
                //複製範本2
                WBook.Worksheets[WorkName].Cells.CreateRange(MaxDataRowIndex, 6, false).Copy(range_2);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 6].PutValue(AddCredit[6] > 0 ? AddCredit[6] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 7].PutValue(AddCredit[7] > 0 ? AddCredit[7] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 8].PutValue(AddCredit[8] > 0 ? AddCredit[8] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 9].PutValue(AddCredit[9] > 0 ? AddCredit[9] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 10].PutValue(AddCredit[10] > 0 ? AddCredit[10] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 11].PutValue(AddCredit[11] > 0 ? AddCredit[11] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 12].PutValue(AddCredit[12] > 0 ? AddCredit[12] + "" : "");
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex, 13].PutValue(AddCredit[13] > 0 ? AddCredit[13] + "" : "");

                int SUMCredit = 0;

                foreach (int LV3 in AddCredit.Values)
                {
                    SUMCredit = SUMCredit + LV3;
                }
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 1, 6].PutValue(SUMCredit);

                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 3, 10].PutValue(Count部訂必修);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 3, 12].PutValue(Count部訂必修學分);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 3, 14].PutValue(Count部訂必修不評分);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 4, 10].PutValue(Count校訂必修);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 4, 12].PutValue(Count校訂必修學分);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 4, 14].PutValue(Count校訂必修不評分);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 5, 10].PutValue(Count校訂選修);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 5, 12].PutValue(Count校訂選修學分);
                WBook.Worksheets[WorkName].Cells[MaxDataRowIndex + 5, 14].PutValue(Count校訂選修不評分);
                #endregion
            }

            WBook.Worksheets.RemoveAt("樣版範本");
            WBook.Worksheets.RemoveAt("小計範本");
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStartPrint.Enabled = true;
            try
            {

                MotherForm.SetStatusBarMessage("請選擇儲存位置", 100);

                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Excel (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = "匯出課程規劃表";

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    WBook.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                }
                else
                {
                    MsgBox.Show("檔案未儲存");
                }
            }
            catch
            {
                MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                return;
            }
            MotherForm.SetStatusBarMessage("匯出課程規劃表已完成");
        }

        //回傳冊別
        private string POPstartLevel(string SL)
        {
            if (tool.PopDic.ContainsKey(SL))
            {
                return tool.PopDic[SL];
            }
            else
            {
                return "";
            }
        }

        //回傳定位
        private int PopLev(string GY, string EE)
        {
            int gy = int.Parse(GY);
            int ee = int.Parse(EE);

            int result = gy + gy + 3 + ee;
            if (result < 3 || result > 13)
            {
                MsgBox.Show("錯誤!");
                return 0;
            }
            return result;
        }

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem LV5 in listView1.Items)
            {
                LV5.Checked = checkBoxX1.Checked;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}