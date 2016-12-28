using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace yingying
{
    public partial class Form1 : Form
    {
        List<District> distList;
        List<BusinessSum> businessSumList;
        string reportYear, reportMonth;

        public Form1()
        {
            InitializeComponent();
            distList = new List<District>();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string year = cbSourceYear.Text;
            string month = cbSourceMonth.Text;

            if (string.IsNullOrEmpty(year) || string.IsNullOrEmpty(month))
            {
                return;
            }

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;

            if (dlg.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dlg.FileNames.Length;

            Excel.Application excel = new Excel.Application();
            object missing = System.Reflection.Missing.Value;
            Excel.Workbook workbook;
            string strConnection = "Data Source=;Initial Catalog=report;User Id=sa;Password=tony2684I8c;";
            using (SqlConnection conn = new SqlConnection(strConnection))
            {
                conn.Open();

                #region 数据库状态复位
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = string.Format("DELETE FROM BusinessHistory WHERE Y = {0} AND M = {1}", year, month);
                sqlcmd.Connection = conn;
                sqlcmd.ExecuteNonQuery();
                #endregion

                foreach (string file in dlg.FileNames)
                {
                    lblCurrentFile.Text = file;
                    progressBar1.Value++;
                    // 打开工作簿
                    workbook = excel.Workbooks.Open(file, missing, true, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    // 根据文件名得到区名
                    string distName = Path.GetFileNameWithoutExtension(file);
                    // 得到指定的工作表

                    FillDatabase(conn, workbook, distName, year, month);

                    workbook.Close(false);
                }
            }
            MessageBox.Show("All file imported.");
        }

        private void FillDatabase(SqlConnection objConnection, Excel.Workbook workbook, string distName, string year, string month)
        {
            int workSheetCount = workbook.Worksheets.Count;
            Excel.Worksheet worksheet1, worksheet2, worksheet3;
            if (workSheetCount < 2)
            {
                return;
            }

            worksheet1 = (Excel.Worksheet)workbook.Worksheets[1]; // 用人单位支付农民工工资情况统计表
            worksheet2 = (Excel.Worksheet)workbook.Worksheets[2]; // 建筑施工企业拖欠农民工工资案件分类情况统计表

            int baseRow1, baseCol1;
            GetFirstSheetBasePosition(worksheet1, out baseRow1, out baseCol1);
            int baseRow2, baseCol2;
            GetSecondSheetBasePosition(worksheet2, out baseRow2, out baseCol2);

            string distId;

            // 查询地区ID
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = string.Format("SELECT ID FROM DistInfo WHERE Title = '{0}'", distName);
            sqlcmd.Connection = objConnection;
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    distId = reader["ID"].ToString();
                }
                else
                {
                    return;
                }
            }

            ProcessWorksheet1(year, month, worksheet1, baseRow1, baseCol1, distId, "加工制造业", sqlcmd);
            baseRow1++;
            ProcessJianZhu(year, month, worksheet1, worksheet2, baseRow1, baseCol1, baseRow2, baseCol2, distId, sqlcmd);
            baseRow1++;
            ProcessWorksheet1(year, month, worksheet1, baseRow1, baseCol1, distId, "批发零售业", sqlcmd);
            baseRow1++;
            ProcessWorksheet1(year, month, worksheet1, baseRow1, baseCol1, distId, "餐饮住宿业", sqlcmd);
            baseRow1++;
            ProcessWorksheet1(year, month, worksheet1, baseRow1, baseCol1, distId, "居民服务业", sqlcmd);
            baseRow1++;
            ProcessWorksheet1(year, month, worksheet1, baseRow1, baseCol1, distId, "其它", sqlcmd);

            if (workSheetCount < 3)
            {
                return;
            }
            worksheet3 = (Excel.Worksheet)workbook.Worksheets[3]; // 农民工30人以上群体性讨要工资案件情况统计表

            //int baseRow3, baseCol3;
            //GetThirdSheetBasePosition(worksheet3, out baseRow3, out baseCol3);

            //FillTuFa(year, month, worksheet3, baseRow3, baseCol3, distId, sqlcmd);
        }

        private static void ProcessWorksheet1(string year, string month, Excel.Worksheet worksheet, int baseRow, int baseCol, string distId, string businessName, SqlCommand sqlcmd)
        {
            // 查询行业ID
            string businessId;
            sqlcmd.CommandText = string.Format("SELECT ID FROM Business WHERE Title = '{0}'", businessName);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    businessId = reader["ID"].ToString();
                }
                else
                {
                    return;
                }
            }

            #region 主动监察
            // 检查单位数
            int unitCount = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol]).Text.ToString(), out unitCount);
            // 结案数量
            int caseFinishedNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 1]).Text.ToString(), out caseFinishedNum);
            // 结案涉及人数
            int personNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 2]).Text.ToString(), out personNum);
            // 追发工资金额
            double amount = 0;
            double.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 3]).Text.ToString(), out amount);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM ZhuDong WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE ZhuDong SET UnitCount = {0}, CaseFinishedNum = {1}, PersonNum = {2}, Amount = {3} WHERE Business = {4} AND Y = {5} AND M = {6} AND Dist = {7}",
                        unitCount.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO ZhuDong(Business,Y,M,Dist,UnitCount,CaseFinishedNum,PersonNum,Amount) VALUES({0},{1},{2},{3},{4},{5},{6},{7})",
                        businessId, year, month, distId, unitCount.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString());
                }
            }
            sqlcmd.ExecuteNonQuery();

            // 更新分行业历史表
            sqlcmd.CommandText = string.Format("SELECT * FROM BusinessHistory WHERE BusinessId = {0} AND Y = {1} AND M = {2}", businessId, year, month);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE BusinessHistory SET CaseNum = CaseNum + {0}, PersonNum = PersonNum + {1}, Amount = Amount +{2} WHERE BusinessId = {3} AND Y = {4} AND M = {5}",
                        caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO BusinessHistory(BusinessId,Y,M,CaseNum,PersonNum,Amount) VALUES({0},{1},{2},{3},{4},{5})",
                        businessId, year, month, caseFinishedNum.ToString(), personNum.ToString(), amount.ToString());
                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion

            #region 投诉举报
            // 立案案件数
            int caseAllNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 4]).Text.ToString(), out caseAllNum);
            // 结案数量
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 5]).Text.ToString(), out caseFinishedNum);
            // 涉及人数
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 6]).Text.ToString(), out personNum);
            // 追发工资金额
            double.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 7]).Text.ToString(), out amount);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM TouSu WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE TouSu SET CaseAllNum = {0}, CaseFinishedNum = {1}, PersonNum = {2}, Amount = {3} WHERE Business = {4} AND Y = {5} AND M = {6} AND Dist = {7}",
                        caseAllNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO TouSu(Business,Y,M,Dist,CaseAllNum,CaseFinishedNum,PersonNum,Amount) VALUES({0},{1},{2},{3},{4},{5},{6},{7})",
                        businessId, year, month, distId, caseAllNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();

            // 更新分行业历史表
            sqlcmd.CommandText = string.Format("UPDATE BusinessHistory SET CaseNum = CaseNum + {0}, PersonNum = PersonNum + {1}, Amount = Amount +{2} WHERE BusinessId = {3} AND Y = {4} AND M = {5}",
                        caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month);
            sqlcmd.ExecuteNonQuery();
            #endregion

            #region 突发事件
            // 案件数
            int eventNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 8]).Text.ToString(), out eventNum);
            // 30人以上案件数
            int bigEventNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 9]).Text.ToString(), out bigEventNum);
            // 结案数量
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 10]).Text.ToString(), out caseFinishedNum);
            // 涉及人数
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 11]).Text.ToString(), out personNum);
            // 追发工资金额
            double.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 12]).Text.ToString(), out amount);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM TuFa WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE TuFa SET EventNum = {0}, BigEventNum = {1}, CaseFinishedNum = {2}, PersonNum = {3}, Amount = {4} WHERE Business = {5} AND Y = {6} AND M = {7} AND Dist = {8}",
                        eventNum.ToString(), bigEventNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO TuFa(Business,Y,M,Dist,EventNum,BigEventNum,CaseFinishedNum,PersonNum,Amount) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8})",
                        businessId, year, month, distId, eventNum.ToString(), bigEventNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion

            #region 案件处理情况
            // 责令改正
            int correntNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 13]).Text.ToString(), out correntNum);
            // 做出行政处理
            int dealNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 14]).Text.ToString(), out dealNum);
            // 处罚件数
            int penalizedNum = 0;
            int.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 15]).Text.ToString(), out penalizedNum);
            // 涉及人数
            double penalizedAmount = 0;
            double.TryParse(((Excel.Range)worksheet.Cells[baseRow, baseCol + 16]).Text.ToString(), out penalizedAmount);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM ChuLi WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE ChuLi SET CorrectNum = {0}, DealNum = {1}, PenalizedNum = {2}, PenalizedAmount = {3} WHERE Business = {4} AND Y = {5} AND M = {6} AND Dist = {7}",
                        correntNum.ToString(), dealNum.ToString(), penalizedNum.ToString(), penalizedAmount.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO ChuLi(Business,Y,M,Dist,CorrectNum,DealNum,PenalizedNum,PenalizedAmount) VALUES({0},{1},{2},{3},{4},{5},{6},{7})",
                        businessId, year, month, distId, correntNum.ToString(), dealNum.ToString(), penalizedNum.ToString(), penalizedAmount.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion
        }

        private static void ProcessJianZhu(string year, string month, Excel.Worksheet worksheet1, Excel.Worksheet worksheet2, int baseRow1, int baseCol1, int baseRow2, int baseCol2, string distId, SqlCommand sqlcmd)
        {
            // 查询行业ID
            string businessId;
            sqlcmd.CommandText = string.Format("SELECT ID FROM Business WHERE Title = '建筑施工业'");
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    businessId = reader["ID"].ToString();
                }
                else
                {
                    return;
                }
            }

            #region 主动监察
            // 检查单位数
            int unitCount = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1]).Text.ToString(), out unitCount);
            // 结案数量
            int caseFinishedNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 1]).Text.ToString(), out caseFinishedNum);
            // 结案涉及人数
            int personNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 2]).Text.ToString(), out personNum);
            // 追发工资金额
            double amount = 0;
            double.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 3]).Text.ToString(), out amount);
            // 三无工程
            int sanwu;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2]).Text.ToString(), out sanwu);
            // 拖欠工程款
            int gongchengkuan;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 1]).Text.ToString(), out gongchengkuan);
            // 结算纠纷
            int jieSuan;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 2]).Text.ToString(), out jieSuan);
            // 非法转包
            int zhuanBao;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 3]).Text.ToString(), out zhuanBao);
            // 使用零散工
            int sanGong;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 4]).Text.ToString(), out sanGong);
            // 无故拖欠工资
            int gongZi;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 5]).Text.ToString(), out gongZi);
            // 其他原因
            int other;
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 6]).Text.ToString(), out other);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM ZhuDong WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE ZhuDong SET UnitCount={0},CaseFinishedNum={1},PersonNum={2},Amount={3},SanWu={4},GongCheng={5},JieSuan={6},ZhuanBao={7},SanGong={8},GongZi={9},Other={10} WHERE Business = {11} AND Y = {12} AND M = {13} AND Dist = {14}",
                        unitCount.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(), jieSuan.ToString(), 
                        zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO ZhuDong(Business,Y,M,Dist,UnitCount,CaseFinishedNum,PersonNum,Amount,SanWu,GongCheng,JieSuan,ZhuanBao,SanGong,GongZi,Other) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14})",
                        businessId, year, month, distId, unitCount.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(),
                        jieSuan.ToString(), zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString());

                }

            }
            sqlcmd.ExecuteNonQuery();

            // 更新分行业历史表
            sqlcmd.CommandText = string.Format("SELECT * FROM BusinessHistory WHERE BusinessId = {0} AND Y = {1} AND M = {2}", businessId, year, month);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE BusinessHistory SET CaseNum = CaseNum + {0}, PersonNum = PersonNum + {1}, Amount = Amount +{2} WHERE BusinessId = {3} AND Y = {4} AND M = {5}",
                        caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO BusinessHistory(BusinessId,Y,M,CaseNum,PersonNum,Amount) VALUES({0},{1},{2},{3},{4},{5})",
                        businessId, year, month, caseFinishedNum.ToString(), personNum.ToString(), amount.ToString());
                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion

            baseRow2++;

            #region 投诉举报
            // 立案案件数
            int caseAllNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 4]).Text.ToString(), out caseAllNum);
            // 结案数量
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 5]).Text.ToString(), out caseFinishedNum);
            // 涉及人数
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 6]).Text.ToString(), out personNum);
            // 追发工资金额
            double.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 7]).Text.ToString(), out amount);
            // 三无工程
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2]).Text.ToString(), out sanwu);
            // 拖欠工程款
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 1]).Text.ToString(), out gongchengkuan);
            // 结算纠纷
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 2]).Text.ToString(), out jieSuan);
            // 非法转包
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 3]).Text.ToString(), out zhuanBao);
            // 使用零散工
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 4]).Text.ToString(), out sanGong);
            // 无故拖欠工资
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 5]).Text.ToString(), out gongZi);
            // 其他原因
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 6]).Text.ToString(), out other);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM TouSu WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE TouSu SET CaseAllNum = {0}, CaseFinishedNum = {1}, PersonNum = {2}, Amount = {3},SanWu={4},GongCheng={5},JieSuan={6},ZhuanBao={7},SanGong={8},GongZi={9},Other={10} WHERE Business = {11} AND Y = {12} AND M = {13} AND Dist = {14}",
                        caseAllNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(), jieSuan.ToString(),
                        zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO TouSu(Business,Y,M,Dist,CaseAllNum,CaseFinishedNum,PersonNum,Amount,SanWu,GongCheng,JieSuan,ZhuanBao,SanGong,GongZi,Other) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14})",
                        businessId, year, month, distId, caseAllNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(),
                        jieSuan.ToString(), zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();

            // 更新分行业历史表
            sqlcmd.CommandText = string.Format("UPDATE BusinessHistory SET CaseNum = CaseNum + {0}, PersonNum = PersonNum + {1}, Amount = Amount +{2} WHERE BusinessId = {3} AND Y = {4} AND M = {5}",
                        caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), businessId, year, month);
            sqlcmd.ExecuteNonQuery();
            #endregion

            baseRow2++;

            #region 突发事件
            // 案件数
            int eventNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 8]).Text.ToString(), out eventNum);
            // 30人以上案件数
            int bigEventNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 9]).Text.ToString(), out bigEventNum);
            // 结案数量
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 10]).Text.ToString(), out caseFinishedNum);
            // 涉及人数
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 11]).Text.ToString(), out personNum);
            // 追发工资金额
            double.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 12]).Text.ToString(), out amount);
            // 三无工程
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2]).Text.ToString(), out sanwu);
            // 拖欠工程款
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 1]).Text.ToString(), out gongchengkuan);
            // 结算纠纷
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 2]).Text.ToString(), out jieSuan);
            // 非法转包
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 3]).Text.ToString(), out zhuanBao);
            // 使用零散工
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 4]).Text.ToString(), out sanGong);
            // 无故拖欠工资
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 5]).Text.ToString(), out gongZi);
            // 其他原因
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow2, baseCol2 + 6]).Text.ToString(), out other);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM TuFa WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE TuFa SET EventNum = {0}, BigEventNum = {1}, CaseFinishedNum = {2}, PersonNum = {3}, Amount = {4}, SanWu={5},GongCheng={6},JieSuan={7},ZhuanBao={8},SanGong={9},GongZi={10},Other={11} WHERE Business = {12} AND Y = {13} AND M = {14} AND Dist = {15}",
                        eventNum.ToString(), bigEventNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(), jieSuan.ToString(),
                        zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO TuFa(Business,Y,M,Dist,EventNum,BigEventNum,CaseFinishedNum,PersonNum,Amount,SanWu,GongCheng,JieSuan,ZhuanBao,SanGong,GongZi,Other) VALUES({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15})",
                        businessId, year, month, distId, eventNum.ToString(), bigEventNum.ToString(), caseFinishedNum.ToString(), personNum.ToString(), amount.ToString(), sanwu.ToString(), gongchengkuan.ToString(), jieSuan.ToString(),
                        zhuanBao.ToString(), sanGong.ToString(), gongZi.ToString(), other.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion

            #region 案件处理情况
            // 责令改正
            int correntNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 13]).Text.ToString(), out correntNum);
            // 做出行政处理
            int dealNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 14]).Text.ToString(), out dealNum);
            // 处罚件数
            int penalizedNum = 0;
            int.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 15]).Text.ToString(), out penalizedNum);
            // 涉及人数
            double penalizedAmount = 0;
            double.TryParse(((Excel.Range)worksheet1.Cells[baseRow1, baseCol1 + 16]).Text.ToString(), out penalizedAmount);

            // 检查是否已存在对应项(Business, Year, Month, Dist)
            sqlcmd.CommandText = string.Format("SELECT * FROM ChuLi WHERE Business = {0} AND Y = {1} AND M = {2} AND Dist = {3}", businessId, year, month, distId);
            using (SqlDataReader reader = sqlcmd.ExecuteReader())
            {
                if (reader.Read())
                {
                    sqlcmd.CommandText = string.Format("UPDATE ChuLi SET CorrectNum = {0}, DealNum = {1}, PenalizedNum = {2}, PenalizedAmount = {3} WHERE Business = {4} AND Y = {5} AND M = {6} AND Dist = {7}",
                        correntNum.ToString(), dealNum.ToString(), penalizedNum.ToString(), penalizedAmount.ToString(), businessId, year, month, distId);
                }
                else
                {
                    sqlcmd.CommandText = string.Format("INSERT INTO ChuLi(Business,Y,M,Dist,CorrectNum,DealNum,PenalizedNum,PenalizedAmount) VALUES({0},{1},{2},{3},{4},{5},{6},{7})",
                        businessId, year, month, distId, correntNum.ToString(), dealNum.ToString(), penalizedNum.ToString(), penalizedAmount.ToString());

                }
            }
            sqlcmd.ExecuteNonQuery();
            #endregion
        }

        private static void FillTuFa(string year, string month, Excel.Worksheet worksheet3, int baseRow, int baseCol, string distId, SqlCommand sqlcmd)
        {
            // 清除原来的
            sqlcmd.CommandText = string.Format("DELETE FROM DistBigEvent WHERE Y = {0} AND M = {1} AND Dist = {2}", year, month, distId);
            sqlcmd.ExecuteNonQuery();

            int personNum = 0;
            string project;
            while (true)
            {
                project = ((Excel.Range)worksheet3.Cells[baseRow, baseCol]).Text.ToString();
                int.TryParse(((Excel.Range)worksheet3.Cells[baseRow, baseCol+4]).Text.ToString(), out personNum);

                if (personNum == 0)
                {
                    break;
                }

                sqlcmd.CommandText = string.Format("INSERT INTO DistBigEvent(Y,M,Dist,Project,PersonNum) VALUES({0},{1},{2},'{3}',{4})", year, month, distId, project, personNum);
                sqlcmd.ExecuteNonQuery();

                baseRow++;
            }
            
        }

        private static void FillReason(Excel.Worksheet worksheet2, int baseRow, int baseCol, ref Reason reason)
        {
            // 主动监察
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol]).Text.ToString(), out reason.SanWu);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 1]).Text.ToString(), out reason.GongChengKuan);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 2]).Text.ToString(), out reason.JieSuan);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 3]).Text.ToString(), out reason.ZhuanBao);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 4]).Text.ToString(), out reason.SanGong);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 5]).Text.ToString(), out reason.GongZi);
            int.TryParse(((Excel.Range)worksheet2.Cells[baseRow, baseCol + 6]).Text.ToString(), out reason.Other);
        }

        public void GetThirdSheetBasePosition(Excel.Worksheet worksheet, out int baseRow, out int baseCol)
        {
            // 在已使用单元格范围内搜索"涉及人数"，得到的单元格作为列基准
            Excel.Range findResult = worksheet.UsedRange.Find("项目名称", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            baseCol = findResult.Column;
            baseRow = findResult.Row + 1;
        }

        public void GetSecondSheetBasePosition(Excel.Worksheet worksheet, out int baseRow, out int baseCol)
        {
            // 在已使用单元格范围内搜索"三无工程"，得到的单元格作为列基准
            Excel.Range findResult = worksheet.UsedRange.Find("三无工程", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            baseCol = findResult.Column;

            // 在已使用单元格范围内搜索"主动监察"，得到的单元格作为行基准
            findResult = worksheet.UsedRange.Find("主动监察", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            baseRow = findResult.Row;
        }

        public void GetFirstSheetBasePosition(Excel.Worksheet worksheet, out int baseRow, out int baseCol)
        {
            // 在已使用单元格范围内搜索"主动监察"，得到的单元格作为列基准
            Excel.Range findResult = worksheet.UsedRange.Find("主动监察", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            if (findResult == null)
            {
                findResult = worksheet.UsedRange.Find("巡视检查", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            }
            baseCol = findResult.Column;

            // 在已使用单元格范围内搜索"加工制造业"，得到的单元格作为行基准
            findResult = worksheet.UsedRange.Find("加工制造业", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            baseRow = findResult.Row;
        }

        private void btnBusinessSum_Click(object sender, EventArgs e)
        {
            if (cbYear.SelectedItem == null)
            {
                MessageBox.Show("请选择报表月份");
                return;
            }
            else
            {
                reportYear = cbYear.Text;
            }

            if (cbMonth.SelectedItem == null)
            {
                MessageBox.Show("请选择报表月份");
                return;
            }
            else
            {
                reportMonth = cbMonth.Text;
            }

            int caseFinishedNum = 0;
            int personNum = 0;
            double amount = 0.0;
            string output;
            string strConnection = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = F:/report.accdb";
            using (OleDbConnection objConnection = new OleDbConnection(strConnection))
            {
                objConnection.Open();
                OleDbCommand sqlcmd = new OleDbCommand();

                #region 案件总数和同比
                // 计算报表月案件总数
                sqlcmd.Parameters.Add("year", OleDbType.Integer).Value = int.Parse(reportYear);
                sqlcmd.Parameters.Add("month", OleDbType.Integer).Value = int.Parse(reportMonth);
                sqlcmd.CommandText = "CaseFinishedSum";
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Connection = objConnection;

                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        caseFinishedNum = int.Parse(reader["casefinishednum"].ToString());
                    }
                    else
                    {
                        return;
                    }
                }

                // 更新历史表
                sqlcmd.Parameters.Clear();
                sqlcmd.CommandText = string.Format("SELECT * FROM History WHERE Y = {0} AND M = {1}", reportYear, reportMonth);
                sqlcmd.CommandType = CommandType.Text;

                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        sqlcmd.CommandText = string.Format("UPDATE History SET CaseSum = {0} WHERE Y = {1} AND M = {2}", caseFinishedNum, reportYear, reportMonth);
                    }
                    else
                    {
                        sqlcmd.CommandText = string.Format("INSERT INTO History(Y, M, CaseSum) VALUES({0}, {1}, {2})", reportYear, reportMonth, caseFinishedNum);
                    }
                }
                sqlcmd.ExecuteNonQuery();

                // 查询同比数据
                string lastCaseData;
                sqlcmd.CommandText = string.Format("SELECT CaseSum FROM History WHERE Y = {0} AND M = {1}", int.Parse(reportYear) - 1, reportMonth);
                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        lastCaseData = reader["CaseSum"].ToString();
                    }
                    else
                    {
                        return;
                    }
                }

                // 计算案件同比
                double caseYearRise = CalculateYearRise(caseFinishedNum, double.Parse(lastCaseData));
                string caseYearRiseString = GetYearRiseString(caseYearRise);
                #endregion

                #region 人数总数和同比
                // 计算报表月人数总数
                sqlcmd.Parameters.Add("year", OleDbType.Integer).Value = int.Parse(reportYear);
                sqlcmd.Parameters.Add("month", OleDbType.Integer).Value = int.Parse(reportMonth);
                sqlcmd.CommandText = "PersonSum";
                sqlcmd.CommandType = CommandType.StoredProcedure;

                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        personNum = int.Parse(reader["PersonNum"].ToString());
                    }
                    else
                    {
                        return;
                    }
                }

                // 更新历史表
                sqlcmd.Parameters.Clear();
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandText = string.Format("UPDATE History SET PersonSum = {0} WHERE Y = {1} AND M = {2}", personNum, reportYear, reportMonth);
                sqlcmd.ExecuteNonQuery();

                // 查询同比数据
                string lastPersonData;
                sqlcmd.CommandText = string.Format("SELECT PersonSum FROM History WHERE Y = {0} AND M = {1}", int.Parse(reportYear) - 1, reportMonth);
                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        lastPersonData = reader["PersonSum"].ToString();
                    }
                    else
                    {
                        return;
                    }
                }

                // 计算人数同比
                double personYearRise = CalculateYearRise(personNum, double.Parse(lastPersonData));
                #endregion

                #region 工资总数和同比
                // 计算报表月工资总数
                sqlcmd.Parameters.Add("year", OleDbType.Integer).Value = int.Parse(reportYear);
                sqlcmd.Parameters.Add("month", OleDbType.Integer).Value = int.Parse(reportMonth);
                sqlcmd.CommandText = "AmountSum";
                sqlcmd.CommandType = CommandType.StoredProcedure;

                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        amount = double.Parse(reader["Amount"].ToString());
                    }
                    else
                    {
                        return;
                    }
                }

                // 更新历史表
                sqlcmd.Parameters.Clear();
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandText = string.Format("UPDATE History SET AmountSum = {0} WHERE Y = {1} AND M = {2}", amount, reportYear, reportMonth);
                sqlcmd.ExecuteNonQuery();

                // 查询同比数据
                string lastAmountData;
                sqlcmd.CommandText = string.Format("SELECT AmountSum FROM History WHERE Y = {0} AND M = {1}", int.Parse(reportYear) - 1, reportMonth);
                using (OleDbDataReader reader = sqlcmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        lastAmountData = reader["AmountSum"].ToString();
                    }
                    else
                    {
                        return;
                    }
                }

                // 计算人数同比
                double amountYearRise = CalculateYearRise(amount, double.Parse(lastAmountData));
                #endregion

                string personAndMoneyRiseString = GetDoubleRiseString(personYearRise, amountYearRise);

                // 生成文本
                // 第一部分
                output = "二、拖欠农民工工资案件情况\r\n（一）查处拖欠农民工工资案件情况\r\n";
                output += reportMonth + "份共查处用人单位拖欠农民工工资案件" + caseFinishedNum + "件，" + caseYearRiseString + "；共为" + personNum + "名农民工追发工资"
                    + Math.Round(amount, 2, MidpointRounding.AwayFromZero) + "万元，" + personAndMoneyRiseString + "。\r\n\r\n";
                output += "图4：全市各月查处拖欠农民工工资案件数量图\r\n(见数据库History表)\r\n";
            }

            businessSumList = new List<BusinessSum>();

            UpdateJiaGongSum();
            UpdateJianZhuSum();
            UpdatePiFaSum();
            UpdateCanYinSum();
            UpdateFuWuSum();
            UpdateOtherSum();

            int jianZhuCaseFinishedNum = 0;
            int jianZhuPersonNum = 0;
            double jianZhuAmount = 0.0;

            foreach (BusinessSum b in businessSumList)
            {
                caseFinishedNum += b.caseFinishNum;
                personNum += b.personNum;
                amount += b.amount;
                if (b.Name == "建筑施工业")
                {
                    jianZhuCaseFinishedNum = b.caseFinishNum;
                    jianZhuPersonNum = b.personNum;
                    jianZhuAmount = b.amount;
                }
            }

            // 计算建筑行业数据
            List<Array> jianZhuCaseHistoryList;
            double lastJianZhuCaseData;
            ProcessHistoryData("case_history_jianzhu.csv", jianZhuCaseFinishedNum, out jianZhuCaseHistoryList, out lastJianZhuCaseData);
            UpdateNewData("case_history_jianzhu.csv", jianZhuCaseHistoryList);
            double jianZhuCaseYearRise = CalculateYearRise(jianZhuCaseFinishedNum, lastJianZhuCaseData);
            string jianZhuCaseYearRiseString = GetYearRiseString(jianZhuCaseYearRise);

            string lastYear = (DateTime.Now.Year - 1).ToString();
            //foreach (string[] arr in caseHistoryList)
            //{
            //    if (arr[0] == lastYear || arr[0] == DateTime.Now.Year.ToString()) 
            //    {
            //        output += string.Join(",", arr);
            //    }
            //    output += "\r\n";
            //}
            //output += "\r\n";

            output += "图5：" + reportMonth + "份全市查处拖欠农民工工资案件\r\n按行业分类数量、占总数比例、与同期对比表\r\n";

            List<Array> caseFenLeiLastYearHistoryList;
            ProcessFenLeiLastYear("case", out caseFenLeiLastYearHistoryList);
            List<Array> caseFenLeiThisYearHistoryList;
            ProcessFenLeiThisYear("case", out caseFenLeiThisYearHistoryList);

            for (int i = 0; i < businessSumList.Count; i++)
            {
                BusinessSum b = businessSumList[i];
                double ratio = Math.Round((double)b.caseFinishNum * 100 / caseFinishedNum, 2, MidpointRounding.AwayFromZero);
                string[] arr = (string[])caseFenLeiLastYearHistoryList[i + 1];
                //if (b.Name != arr[0])
                //{
                //    throw new Exception("行业顺序不匹配"); // 编码格式要处理一下
                //}
                int lastData = int.Parse(arr[cbMonth.SelectedIndex + 1]);
                double riseRate = Math.Round((double)(b.caseFinishNum - lastData) * 100 / lastData, 2, MidpointRounding.AwayFromZero);
                output += b.caseFinishNum + ",_," + ratio.ToString() + ",_," + riseRate.ToString() + ",_\r\n";
            }

            output += "图7：1月份全市查处拖欠农民工工资案件效果\r\n按行业分类数量与同期对比表\r\n";

            // 生成图7
            List<Array> personFenLeiHistoryList;
            ProcessFenLeiLastYear("person", out personFenLeiHistoryList);
            List<Array> moneyFenLeiHistoryList;
            ProcessFenLeiLastYear("money", out moneyFenLeiHistoryList);

            for (int i = 0; i < businessSumList.Count; i++)
            {
                BusinessSum b = businessSumList[i];
               
                string[] arrPerson = (string[])personFenLeiHistoryList[i + 1];
                string[] arrMoney = (string[])moneyFenLeiHistoryList[i + 1];
                //if (b.Name != arr[0])
                //{
                //    throw new Exception("行业顺序不匹配"); // 编码格式要处理一下
                //}
                int personLastData = int.Parse(arrPerson[cbMonth.SelectedIndex + 1]);
                double moneyLastData = double.Parse(arrMoney[cbMonth.SelectedIndex + 1]);
                double personRiseRate = Math.Round((double)(b.personNum - personLastData) * 100 / personLastData, 2, MidpointRounding.AwayFromZero);
                double moneyRiseRate = Math.Round((b.amount - moneyLastData) * 100 / moneyLastData, 2, MidpointRounding.AwayFromZero);
                double amountRound = Math.Round(b.amount, 2, MidpointRounding.AwayFromZero);
                output += b.personNum + ",_," + amountRound + ",_," + personRiseRate.ToString() + ",_," + moneyRiseRate.ToString() + ",_\r\n";
            }

            // 生成图8
            foreach (District d in distList)
            {
                int caseSum = d.JiaGong.zhudong.caseFinishNum + d.JiaGong.tousu.caseFinishNum + d.JiaGong.tufa.caseFinishNum;
                caseSum += d.JianZhu.zhudong.caseFinishNum + d.JianZhu.tousu.caseFinishNum + d.JianZhu.tufa.caseFinishNum;
                caseSum += d.PiFa.zhudong.caseFinishNum + d.PiFa.tousu.caseFinishNum + d.PiFa.tufa.caseFinishNum;
                caseSum += d.CanYin.zhudong.caseFinishNum + d.CanYin.tousu.caseFinishNum + d.CanYin.tufa.caseFinishNum;
                caseSum += d.FuWu.zhudong.caseFinishNum + d.FuWu.tousu.caseFinishNum + d.FuWu.tufa.caseFinishNum;
                caseSum += d.Other.zhudong.caseFinishNum + d.Other.tousu.caseFinishNum + d.Other.tufa.caseFinishNum;
                output += d.Name + ", " + caseSum + "\r\n";
            }


            // 第二部分
            output += "\r\n（二）查处建筑施工企业拖欠农民工工资案件情况";

            output += cbMonth.SelectedItem.ToString() + "份共查处建筑施工企业拖欠农民工工资案件" + jianZhuCaseFinishedNum + "件，" + jianZhuCaseYearRiseString + "。";

            // TODO: 前几月统计
            /*
            string toMonthString = "";
            if (DateTime.Now.Month > 1)
            {
                toMonthString = "-" + DateTime.Now.Month.ToString();
            }
            */

            output += "按原因分类：";

            BusinessSum jzbs = businessSumList[1]; // 建筑行业各区汇总数据
            if (jzbs.reason.SanWu != 0)
            {
                output += "因劳务企业原因的" + jzbs.reason.SanWu + "件，";
            }
            if (jzbs.reason.GongChengKuan != 0)
            {
                output += "因工程款不及时到位的" + jzbs.reason.GongChengKuan + "件，";
            }
            if (jzbs.reason.JieSuan != 0)
            {
                output += "因劳务费结算纠纷的" + jzbs.reason.JieSuan + "件，";
            }
            if (jzbs.reason.ZhuanBao != 0)
            {
                output += "因非法转包工程的" + jzbs.reason.ZhuanBao + "件，";
            }
            if (jzbs.reason.SanGong != 0)
            {
                output += "因随意使用零散工的" + jzbs.reason.SanGong + "件，";
            }
            if (jzbs.reason.GongZi != 0)
            {
                output += "因无故拖欠工资的" + jzbs.reason.GongZi + "件，";
            }
            if (jzbs.reason.Other != 0)
            {
                output += "因其它原因的" + jzbs.reason.Other + "件，";
            }

            output += "\r\n\r\n";

            // 第三部分
            int tuFaCountSum = 0;
            int tufaPersonSum = 0;

            
            foreach (District d in distList)
            {
                /* 这个是30人以上的统计
                tuFaCountSum += d.tuFaSum.count;
                tufaPersonSum += d.tuFaSum.person;
                */

                tuFaCountSum += d.JiaGong.tufa.caseFinishNum + d.JianZhu.tufa.caseFinishNum + d.PiFa.tufa.caseFinishNum + d.CanYin.tufa.caseFinishNum + d.FuWu.tufa.caseFinishNum + d.Other.tufa.caseFinishNum;
                tufaPersonSum += d.JiaGong.tufa.personNum + d.JianZhu.tufa.personNum + d.PiFa.tufa.personNum + d.CanYin.tufa.personNum + d.FuWu.tufa.personNum + d.Other.tufa.personNum;
            }

            output += cbMonth.SelectedItem.ToString() + "份共参与处理农民工群体性讨薪突发事件" + tuFaCountSum + "起，涉及农民工" + tufaPersonSum + "人，";

            // 计算突发事件件数
            List<Array> tufaCountHistoryList;
            double lastTuFaCountData;
            ProcessHistoryData("tufacount_history.csv", tuFaCountSum, out tufaCountHistoryList, out lastTuFaCountData);
            UpdateNewData("tufacount_history.csv", tufaCountHistoryList);
            double tuFaCountYearRise = CalculateYearRise(tuFaCountSum, lastTuFaCountData);

            // 计算突发事件人数
            List<Array> tufaPersonHistoryList;
            double lastTuFaPersonData;
            ProcessHistoryData("tufaperson_history.csv", tufaPersonSum, out tufaPersonHistoryList, out lastTuFaPersonData);
            UpdateNewData("tufaperson_history.csv", tufaPersonHistoryList);
            double tuFaPersonYearRise = CalculateYearRise(tufaPersonSum, lastTuFaPersonData);

            string tuFaCountAndPersonRiseString = GetDoubleRiseString(tuFaCountYearRise, tuFaPersonYearRise);

            output += tuFaCountAndPersonRiseString + "。\r\n\r\n";

            // 生成图10
            foreach (District d in distList)
            {
                output += d.Name + ", " + d.tuFaSum.count + "\r\n";
            }

            txtOutput.Text = output;
        }

        private static string GetDoubleRiseString(double yearRise1, double yearRise2)
        {
            bool isSameTrend = false;
            if (yearRise1 * yearRise2 >= 0)
            {
                isSameTrend = true;
            }

            string str = "同比分别";

            string personString = "";

            if (yearRise1 >= 0)
            {
                personString = "上升" + yearRise1.ToString() + "%";
            }
            else
            {
                personString = "下降" + (-yearRise1).ToString() + "%";
            }
            str += personString;

            string moneyString = "";

            if (isSameTrend)
            {
                moneyString = Math.Abs(yearRise2).ToString() + "%";
            }
            else
            {
                if (yearRise2 >= 0)
                {
                    moneyString = "上升" + yearRise2.ToString() + "%";
                }
                else
                {
                    moneyString = "下降" + (-yearRise2).ToString() + "%";
                }
            }
           

            str += "和" + moneyString;

            return str;
        }

        private static string GetYearRiseString(double caseYearRise)
        {
            string str1 = "";
            if (caseYearRise > 0)
            {
                str1 = "同比上升" + caseYearRise.ToString() + "%";
            }
            else if (caseYearRise < 0)
            {
                str1 = "同比下降" + (-caseYearRise).ToString() + "%";
            }
            else
            {
                str1 = "与去年持平";
            }

            return str1;
        }

        private static double CalculateYearRise(double thisMonthData, double lastData)
        {
            double yearRise = 0.0;

            if (lastData != 0)
            {
                yearRise = Math.Round((double)(thisMonthData - lastData) * 100 / lastData, 2, MidpointRounding.AwayFromZero);
            }

            return yearRise;
        }

        private static void UpdateNewData(string fileName, List<Array> caseHistoryList)
        {
            string newFileName = "new" + fileName;
            File.Delete(fileName);
            FileStream dstfs = new FileStream(newFileName, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(dstfs);
            foreach (string[] arr in caseHistoryList)
            {
                sw.WriteLine(string.Join(",", arr));
            }
            sw.Close();

            FileInfo fi = new FileInfo(newFileName);
            fi.MoveTo(fileName);
        }

        private void ProcessFenLeiLastYear(string type, out List<Array> fenLeiHistoryList)
        {
            ProcessFenLei(type, (DateTime.Now.Year - 1).ToString(), out fenLeiHistoryList);
        }

        private void ProcessFenLeiThisYear(string type, out List<Array> fenLeiHistoryList)
        {
            ProcessFenLei(type, DateTime.Now.Year.ToString(), out fenLeiHistoryList);
        }

        private void ProcessFenLei(string type, string year, out List<Array> fenLeiHistoryList)
        {
            string lastFileName = type + "_fenlei_" + year + ".csv";
            FileStream srcfs = new FileStream(lastFileName, FileMode.Open, FileAccess.Read);

            StreamReader sr = new StreamReader(srcfs);
            fenLeiHistoryList = new List<Array>();
            string line;
            string[] arrLine;
            while ((line = sr.ReadLine()) != null)
            {
                arrLine = line.Split(',');
                fenLeiHistoryList.Add(arrLine);
            }
            sr.Close();
        }

        private void ProcessHistoryData(string historyFile, double thisMonthSum, out List<Array> historyList, out double lastData)
        {
            FileStream srcfs = new FileStream(historyFile, FileMode.Open, FileAccess.Read);

            StreamReader sr = new StreamReader(srcfs);
            historyList = new List<Array>();
            string line;
            string[] arrLine;
            string lastYear = (DateTime.Now.Year - 1).ToString();
            lastData = 0;
            while ((line = sr.ReadLine()) != null)
            {
                arrLine = line.Split(',');
                if (arrLine[0] == lastYear)
                {
                    lastData = Double.Parse(arrLine[cbMonth.SelectedIndex + 1]);
                }
                if (arrLine[0] == DateTime.Now.Year.ToString())
                {
                    arrLine[cbMonth.SelectedIndex + 1] = thisMonthSum.ToString();
                }
                historyList.Add(arrLine);
            }
            sr.Close();
        }

        public void UpdateJiaGongSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "加工制造业";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.JiaGong.zhudong.caseFinishNum;
                b.caseFinishNum += d.JiaGong.tousu.caseFinishNum;

                b.personNum += d.JiaGong.zhudong.personNum;
                b.personNum += d.JiaGong.tousu.personNum;

                b.amount += d.JiaGong.zhudong.amount;
                b.amount += d.JiaGong.tousu.amount;
            }
            businessSumList.Add(b);
        }

        public void UpdateJianZhuSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "建筑施工业";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.JianZhu.zhudong.caseFinishNum;
                b.caseFinishNum += d.JianZhu.tousu.caseFinishNum;

                b.personNum += d.JianZhu.zhudong.personNum;
                b.personNum += d.JianZhu.tousu.personNum;

                b.amount += d.JianZhu.zhudong.amount;
                b.amount += d.JianZhu.tousu.amount;

                b.reason.SanWu += d.JianZhu.zhudong.reason.SanWu + d.JianZhu.tousu.reason.SanWu + d.JianZhu.tufa.reason.SanWu;
                b.reason.GongChengKuan += d.JianZhu.zhudong.reason.GongChengKuan + d.JianZhu.tousu.reason.GongChengKuan + d.JianZhu.tufa.reason.GongChengKuan;
                b.reason.JieSuan += d.JianZhu.zhudong.reason.JieSuan + d.JianZhu.tousu.reason.JieSuan + d.JianZhu.tufa.reason.JieSuan;
                b.reason.ZhuanBao += d.JianZhu.zhudong.reason.ZhuanBao + d.JianZhu.tousu.reason.ZhuanBao + d.JianZhu.tufa.reason.ZhuanBao;
                b.reason.SanGong += d.JianZhu.zhudong.reason.SanGong + d.JianZhu.tousu.reason.SanGong + d.JianZhu.tufa.reason.SanGong;
                b.reason.GongZi += d.JianZhu.zhudong.reason.GongZi + d.JianZhu.tousu.reason.GongZi + d.JianZhu.tufa.reason.GongZi;
                b.reason.Other += d.JianZhu.zhudong.reason.Other + d.JianZhu.tousu.reason.Other + d.JianZhu.tufa.reason.Other;
            }
            businessSumList.Add(b);
        }

        public void UpdatePiFaSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "批发零售业";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.PiFa.zhudong.caseFinishNum;
                b.caseFinishNum += d.PiFa.tousu.caseFinishNum;

                b.personNum += d.PiFa.zhudong.personNum;
                b.personNum += d.PiFa.tousu.personNum;

                b.amount += d.PiFa.zhudong.amount;
                b.amount += d.PiFa.tousu.amount;
            }
            businessSumList.Add(b);
        }

        public void UpdateCanYinSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "餐饮住宿业";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.CanYin.zhudong.caseFinishNum;
                b.caseFinishNum += d.CanYin.tousu.caseFinishNum;

                b.personNum += d.CanYin.zhudong.personNum;
                b.personNum += d.CanYin.tousu.personNum;

                b.amount += d.CanYin.zhudong.amount;
                b.amount += d.CanYin.tousu.amount;
            }
            businessSumList.Add(b);
        }

        public void UpdateFuWuSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "居民服务业";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.FuWu.zhudong.caseFinishNum;
                b.caseFinishNum += d.FuWu.tousu.caseFinishNum;

                b.personNum += d.FuWu.zhudong.personNum;
                b.personNum += d.FuWu.tousu.personNum;

                b.amount += d.FuWu.zhudong.amount;
                b.amount += d.FuWu.tousu.amount;
            }
            businessSumList.Add(b);
        }

        public void UpdateOtherSum()
        {
            BusinessSum b = new BusinessSum();
            b.Name = "其它";
            foreach (District d in distList)
            {
                b.caseFinishNum += d.Other.zhudong.caseFinishNum;
                b.caseFinishNum += d.Other.tousu.caseFinishNum;

                b.personNum += d.Other.zhudong.personNum;
                b.personNum += d.Other.tousu.personNum;

                b.amount += d.Other.zhudong.amount;
                b.amount += d.Other.tousu.amount;
            }
            businessSumList.Add(b);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            //cbMonth.SelectedIndex = now.Month - 1;
        }

        private void btnExportToWord_Click(object sender, EventArgs e)
        {
            Word.Application word = new Word.Application();
            object missing = System.Reflection.Missing.Value;
            Word.Document doc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Word.Range range = doc.Paragraphs[1].Range;
            range.Text = txtOutput.Text;
            object fileName = "report.docx";
            doc.SaveAs2(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            word.Quit(ref missing, ref missing, ref missing);
        }
    }
}
