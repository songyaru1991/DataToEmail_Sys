using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using System.Data;
using NPOI.HSSF.UserModel;
using System.IO;
using DAL;
using System.Net.Mail;
namespace BLL
{
    public class UserService
    {
        /// <summary>
        /// 生成Sheet页，运行一次这个方法生成一页
        /// </summary>
        /// <param name="workBook">文档对象</param>
        /// <param name="sheetName">文档名字</param>
        /// <param name="table">数据源</param>
        /// <returns></returns>
        public ISheet CreateSheetByCostId(IWorkbook workBook, string sheetName, DataTable table)
        {
            // int colIndex = -1;
            ISheet sheet = workBook.CreateSheet(sheetName);
            IRow RowHead = sheet.CreateRow(0);
            ICellStyle titleStyle = workBook.CreateCellStyle();
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;//居中
            IFont titleFont = workBook.CreateFont();
            titleFont.FontHeightInPoints = 9;//设置字体大小
            titleFont.FontName = "微软雅黑";
            titleStyle.SetFont(titleFont);//字体样式赋值
            //foreach (DataColumn col in table.Columns)
            //{
            //    colIndex++;
            //    RowHead.CreateCell(colIndex).SetCellValue(col.ColumnName);
            //}//头
            string curDate = GetDate(0);
            string[] strs = new string[] { "部门代码", "有刷卡總人數", "有上下刷總人數", "白班上下刷总人数", "夜班上下刷总人数", "生成加班單人數 " };
            for (int i = 0; i < strs.Length; i++)
            {
                RowHead.CreateCell(i).SetCellValue(strs[i]);
            }//execl头部赋值

            for (int iRowIndex = 0; iRowIndex < table.Rows.Count; iRowIndex++)
            {
                IRow RowBody = sheet.CreateRow(iRowIndex + 1);
                for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
                {

                    if (table.Rows[iRowIndex]["costid"].ToString().Equals("") || table.Rows[iRowIndex]["costid"] == null)
                    {
                        RowBody.CreateCell(0).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(0).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["costid"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["ac"].ToString().Equals("") || table.Rows[iRowIndex]["ac"] == null)
                    {
                        RowBody.CreateCell(1).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(1).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["ac"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["bc"].ToString().Equals("") || table.Rows[iRowIndex]["bc"] == null)
                    {
                        RowBody.CreateCell(2).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(2).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["bc"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["cc"].ToString().Equals("") || table.Rows[iRowIndex]["cc"] == null)
                    {
                        RowBody.CreateCell(3).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(3).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["cc"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["dc"].ToString().Equals("") || table.Rows[iRowIndex]["dc"] == null)
                    {
                        RowBody.CreateCell(4).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(4).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["dc"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["ec"].ToString().Equals("") || table.Rows[iRowIndex]["ec"] == null)
                    {
                        RowBody.CreateCell(5).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(5).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["ec"].ToString()));
                    }
                    
                    sheet.SetColumnWidth(iColumnIndex, 50 * 100);
                    sheet.GetRow(iRowIndex).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
                }
            }
            for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
            {
                sheet.GetRow(sheet.LastRowNum).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
            }
            return sheet;
        }

        public ISheet CreateSheetByThreeColumn(IWorkbook workBook, string sheetName, DataTable table, string[] strs, string[] strs_title)
        {
            // int colIndex = -1;
            ISheet sheet = workBook.CreateSheet(sheetName);
            IRow RowHead = sheet.CreateRow(0);
            ICellStyle titleStyle = workBook.CreateCellStyle();
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;//居中
            IFont titleFont = workBook.CreateFont();
            titleFont.FontHeightInPoints = 9;//设置字体大小
            titleFont.FontName = "微软雅黑";
            titleStyle.SetFont(titleFont);//字体样式赋值
          //  string[] strs = new string[] { "BU", "白夜班", "有上下刷總人數" };
            for (int i = 0; i < strs.Length; i++)
            {
                RowHead.CreateCell(i).SetCellValue(strs_title[i]);
            }//execl头部赋值

            for (int iRowIndex = 0; iRowIndex < table.Rows.Count; iRowIndex++)
            {
                IRow RowBody = sheet.CreateRow(iRowIndex + 1);
                for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
                {

                    if (table.Rows[iRowIndex]["aa"].ToString().Equals("") || table.Rows[iRowIndex]["aa"] == null)
                    {
                        RowBody.CreateCell(0).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(0).SetCellValue(table.Rows[iRowIndex]["aa"].ToString());
                    }

                    if (table.Rows[iRowIndex]["bb"].ToString().Equals("") || table.Rows[iRowIndex]["bb"] == null)
                    {
                        RowBody.CreateCell(1).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(1).SetCellValue(table.Rows[iRowIndex]["bb"].ToString());
                    }

                    if (table.Rows[iRowIndex]["cc"].ToString().Equals("") || table.Rows[iRowIndex]["cc"] == null)
                    {
                        RowBody.CreateCell(2).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(2).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["cc"].ToString()));
                    }

                    sheet.SetColumnWidth(iColumnIndex, 50 * 100);
                    sheet.GetRow(iRowIndex).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
                }              
            }
            for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
            {
                sheet.GetRow(sheet.LastRowNum).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
            }
            return sheet;
        }

        public ISheet CreateSheetByFourColumn(IWorkbook workBook, string sheetName, DataTable table, string[] strs, string[] strs_title)
        {
            // int colIndex = -1;
            ISheet sheet = workBook.CreateSheet(sheetName);
            IRow RowHead = sheet.CreateRow(0);
            ICellStyle titleStyle = workBook.CreateCellStyle();
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;//居中
            IFont titleFont = workBook.CreateFont();
            titleFont.FontHeightInPoints = 9;//设置字体大小
            titleFont.FontName = "微软雅黑";
            titleStyle.SetFont(titleFont);//字体样式赋值
            //  string[] strs = new string[] { "BU", "白夜班", "有上下刷總人數" };
            for (int i = 0; i < strs.Length; i++)
            {
                RowHead.CreateCell(i).SetCellValue(strs_title[i]);
            }//execl头部赋值

            for (int iRowIndex = 0; iRowIndex < table.Rows.Count; iRowIndex++)
            {
                IRow RowBody = sheet.CreateRow(iRowIndex + 1);
                for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
                {

                    if (table.Rows[iRowIndex]["aa"].ToString().Equals("") || table.Rows[iRowIndex]["aa"] == null)
                    {
                        RowBody.CreateCell(0).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(0).SetCellValue(table.Rows[iRowIndex]["aa"].ToString());
                    }

                    if (table.Rows[iRowIndex]["bb"].ToString().Equals("") || table.Rows[iRowIndex]["bb"] == null)
                    {
                        RowBody.CreateCell(1).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(1).SetCellValue(table.Rows[iRowIndex]["bb"].ToString());
                    }

                    if (table.Rows[iRowIndex]["cc"].ToString().Equals("") || table.Rows[iRowIndex]["cc"] == null)
                    {
                        RowBody.CreateCell(2).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(2).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["cc"].ToString()));
                    }

                    if (table.Rows[iRowIndex]["dd"].ToString().Equals("") || table.Rows[iRowIndex]["dd"] == null)
                    {
                        RowBody.CreateCell(3).SetCellValue(0);
                    }
                    else
                    {
                        RowBody.CreateCell(3).SetCellValue(Convert.ToInt32(table.Rows[iRowIndex]["dd"].ToString()));
                    }

                    sheet.SetColumnWidth(iColumnIndex, 50 * 100);
                    sheet.GetRow(iRowIndex).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
                }
            }
            for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
            {
                sheet.GetRow(sheet.LastRowNum).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
            }
            return sheet;
        }


        /// 生成Sheet页
        /// </summary>
        /// <param name="workBook">文档对象</param>
        /// <param name="sheetName">sheet名字</param>
        /// <param name="table">数据源</param>
        /// <param name="strs">区分内容的数组</param>
        /// <param name="strs_title">头部数组</param>
        /// <returns></returns>
        public ISheet CreateSheetbyEmp(IWorkbook workBook, string sheetName, DataTable table, string[] strs, string[] strs_title)
        {
            // int colIndex = -1;
            ISheet sheet = workBook.CreateSheet(sheetName);
            IRow RowHead = sheet.CreateRow(0);
            ICellStyle titleStyle = workBook.CreateCellStyle();
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.CENTER;//居中
            IFont titleFont = workBook.CreateFont();
            titleFont.FontHeightInPoints = 9;//设置字体大小
            titleFont.FontName = "微软雅黑";
            titleStyle.SetFont(titleFont);//字体样式赋值
            for (int i = 0; i < strs.Length; i++)
            {
                RowHead.CreateCell(i).SetCellValue(strs_title[i]);
            }//execl头部赋值

            for (int iRowIndex = 0; iRowIndex < table.Rows.Count; iRowIndex++)
            {
                IRow RowBody = sheet.CreateRow(iRowIndex+1);
                for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
                {
                    RowBody.CreateCell(iColumnIndex).SetCellValue((table.Rows[iRowIndex][iColumnIndex].ToString()));

                    sheet.SetColumnWidth(iColumnIndex, 30 * 100);

                    string ss = strs[iColumnIndex];
                    if (strs[iColumnIndex].Equals("csSw1") || strs[iColumnIndex].Equals("csSw2"))
                    {

                        sheet.SetColumnWidth(iColumnIndex, 50 * 100);
                    }
                    else if (strs[iColumnIndex].Equals("cedepname"))
                    {
                        sheet.SetColumnWidth(iColumnIndex, 70 * 100);
                    }
                    sheet.GetRow(iRowIndex).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
                }
            }
            for (int iColumnIndex = 0; iColumnIndex < strs.Length; iColumnIndex++)
            {
                sheet.GetRow(sheet.LastRowNum).GetCell(iColumnIndex).CellStyle = titleStyle;//居中赋值
            }
            return sheet;
        }



        /// <summary>
        /// 获取取得刷卡數據的sql By CostID
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        public string GetSql_GroupSUM(string selectDate)
        {
                //SELECT a.costid,a.ac,b.bc,c.cc FROM (SELECT costid,COUNT(*) ac FROM CSR_EMPLOYEE WHERE isOnWork = 0 GROUP BY costid) a LEFT OUTER JOIN (SELECT te.costid,COUNT(*) bc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.\"ID\"=ts.emp_id     AND ts.swipe_date='{0}' AND ts.SwipeCardTime IS NOT NULL  AND ts.SwipeCardTime2 IS NOT NULL  GROUP BY te.costid) b  ON a.costid = b.costid  LEFT OUTER JOIN  (SELECT costid, COUNT(*) cc  FROM notes_overtime_state  WHERE overtimedate='{1}' AND notesStates = 1 GROUP BY costid) c  ON a.costid = c.costid  Where  a.costID in(select depid from a2_dept) ORDER BY costid
            string sql = "SELECT a.costid,a.ac,b.bc,c.cc,d.dc,e.ec FROM (SELECT costid,COUNT(*) ac FROM CSR_EMPLOYEE WHERE isOnWork = 0 GROUP BY costid) a "
          /* +" LEFT OUTER JOIN (SELECT te.costid,COUNT(*) bc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.ID=ts.emp_id  "
           +" AND ts.swipe_date='" + selectDate + "' GROUP BY te.costid) b   ON a.costid = b.costid" */
           +" LEFT OUTER JOIN (SELECT te.costid,COUNT(*) bc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.ID=ts.emp_id "
           + " AND ts.swipe_date='" + selectDate + "' AND ts.SwipeCardTime IS NOT NULL  AND ts.SwipeCardTime2 IS NOT NULL  GROUP BY te.costid) b   ON a.costid = b.costid "
           +"   LEFT OUTER JOIN (SELECT te.costid,COUNT(*) cc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.ID=ts.emp_id" 
           +"  AND ts.swipe_date='" + selectDate + "'  AND ts.SwipeCardTime IS NOT NULL  AND ts.SwipeCardTime2 IS NOT NULL and shift='D' GROUP BY te.costid) c  ON a.costid = c.costid" 
           +"  LEFT OUTER JOIN (SELECT te.costid,COUNT(*) dc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.ID=ts.emp_id" 
           +"  AND ts.swipe_date='" + selectDate + "' AND ts.SwipeCardTime IS NOT NULL  AND ts.SwipeCardTime2 IS NOT NULL and shift='N' GROUP BY te.costid) d ON a.costid = d.costid"   
           +"  LEFT OUTER JOIN  (SELECT costid, COUNT(*) ec  FROM notes_overtime_state  WHERE overtimedate='" + selectDate + "'  AND notesStates = 1 GROUP BY costid) e ON a.costid = e.costid"  
           +"  Where  a.costID in(select depid from a2_dept) ORDER BY costid";           
            return sql;
        }

        /// <summary>
        ///获取取得刷卡數據的sql By  BU
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        public string GetSql_ComponentSOLO(string selectDate)
        {
            string sql = "SELECT '零件產品事業群' aa,'白班' bb,COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          +" FROM csr_employee te, csr_swipecardtime ts  WHERE te.id = ts.emp_ID "
                          + " AND swipe_date='" + selectDate + "'  and shift='D'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%零件%' group by A.DEPID)"
                          +" UNION"
                          + " SELECT '零件產品事業群' aa,'夜班' bb,COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          +"    FROM csr_employee te, csr_swipecardtime ts  WHERE te.id = ts.emp_ID "
                          + "  AND swipe_date='" + selectDate + "'  and shift='N'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%零件%' group by A.DEPID)"
                          +" UNION"
                          + " SELECT '組件產品事業群' aa,'白班' bb, COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          +" FROM csr_employee te,csr_swipecardtime ts  WHERE te.id = ts.emp_ID "
                          + "  AND swipe_date='" + selectDate + "'  and shift='D'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%組件%' group by A.DEPID)"
                          +" UNION"
                          + " SELECT '組件產品事業群' aa,'夜班' bb, COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          + " FROM csr_employee te, csr_swipecardtime ts  WHERE te.id = ts.emp_ID "
                          + " AND swipe_date='" + selectDate + "'  and shift='N'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%組件%' group by A.DEPID)"
                          +" UNION"
                           + " SELECT '通訊產品事業群' aa,'白班' bb,COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          +" FROM csr_employee te, csr_swipecardtime ts WHERE te.id = ts.emp_ID "
                          + " AND swipe_date='" + selectDate + "'  and shift='D'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%通訊%' group by A.DEPID)"
                          +" UNION"
                          + " SELECT '通訊產品事業群' aa,'夜班' bb,COUNT(*) cc,SUM(CASE WHEN ts.swipecardtime IS NOT NULL and ts.swipecardtime2 IS NOT NULL THEN 1 ELSE 0 END) AS dd"
                          +"  FROM csr_employee te,csr_swipecardtime ts  WHERE te.id = ts.emp_ID "
                           + " AND swipe_date='" + selectDate + "'  and shift='N'"
                          +" and costID in(select A.DEPID from A2_dept a,DEPT_RELATION b where A.DEPID=B.COSTID and B.DEPT_NAME like '%通訊%' group by A.DEPID)";
            return sql;
        }

        /// <summary>
        ///获取取得忘卡人員信息的sql
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        public string GetSql_FCSwipeCardEMP(string selectDate)
        {
            string sql = "select id ,name,costid,depid,'" + selectDate + "' fcDate from swipe.csr_employee "
				+" where isOnWork = 0 AND id NOT IN (SELECT distinct(emp_id) FROM swipe.csr_swipecardtime st WHERE" 
				+" st.emp_id IN (SELECT id FROM  swipe.csr_employee) and"
                + " st.swipe_date ='" + selectDate + "') and costID in(select depid from a2_dept) order by costid";
            return sql;
        }

       
        ///获取取得忘卡人員數據ByCostID的sql
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        public string GetSql_FCSwipeCardEMPByCostID(string selectDate)
        {
            string sql = "select  a.costid aa,a.fcDate bb,COUNT(*) cc from( "
                + " select id ,name,costid,depid,isOnWork,'" + selectDate + "' fcDate from swipe.csr_employee " 
				+ " where isOnWork = 0 AND id NOT IN (SELECT distinct(emp_id) FROM swipe.csr_swipecardtime st WHERE " 
				+ " st.emp_id IN (SELECT id FROM  swipe.csr_employee) and "
                + " st.swipe_date ='" + selectDate + "') and costID in(select depid from a2_dept)) a group by  a.costid,a.fcDate";
            return sql;
        }

        /// <summary>
        ///获取取得A2在职人員ByCOSTID
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        /// <summary>
        /// 
        public string GetSql_EMPByCostID()
        {
            //SELECT a.costid,a.ac,b.bc,c.cc FROM (SELECT costid,COUNT(*) ac FROM CSR_EMPLOYEE WHERE isOnWork = 0 GROUP BY costid) a LEFT OUTER JOIN (SELECT te.costid,COUNT(*) bc FROM CSR_EMPLOYEE te,CSR_SWIPECARDTIME ts     WHERE te.\"ID\"=ts.emp_id     AND ts.swipe_date='{0}' AND ts.SwipeCardTime IS NOT NULL  AND ts.SwipeCardTime2 IS NOT NULL  GROUP BY te.costid) b  ON a.costid = b.costid  LEFT OUTER JOIN  (SELECT costid, COUNT(*) cc  FROM notes_overtime_state  WHERE overtimedate='{1}' AND notesStates = 1 GROUP BY costid) c  ON a.costid = c.costid  Where  a.costID in(select depid from a2_dept) ORDER BY costid
            string sql = "SELECT costid,COUNT(*) empCount FROM CSR_EMPLOYEE WHERE isOnWork = 0 and costID in(select depid from a2_dept) GROUP BY costid order by costid";
            return sql;
        }

        /// <summary>
        ///获取取得A2在职人員详情的sql
        /// </summary>
        /// <param name="dayNum"></param>
        /// <returns></returns>
        /// <summary>
        /// 
        public string GetSql_EMP()
        {
            string sql = "SELECT id ,name,costid,depid,depname,cardid FROM csr_employee WHERE isOnWork = 0 and costID in(select depid from a2_dept) order by costid";
            return sql;
        }

        public string GetDate(int dayNum)
        {
            StringBuilder sb = new StringBuilder();
            DateTime sysTime = DateTime.Now.AddDays(dayNum);
            return sysTime.ToString("yyyyMMdd");
        }
        /// <summary>
        /// 格式化日期，oracleDB
        /// </summary>
        /// <param name="dayNum">当前时间上加几天</param>
        /// <returns></returns>
        public string GetTime_Oracle(int dayNum)
        {
            StringBuilder sb = new StringBuilder();
            DateTime sysTime = DateTime.Now.AddDays(dayNum);
            return sysTime.ToString("yyyy-MM-dd");
        }

        /// <summary>
        /// 格式化日期，MysqlDB
        /// </summary>
        /// <param name="dayNum">当前时间上加几天</param>
        /// <returns></returns>
        public string GetTime_Mysql(int dayNum)
        {
            StringBuilder sb = new StringBuilder();
            DateTime sysTime = DateTime.Now.AddDays(dayNum);
            string[] strs = sysTime.ToString("yyyy/MM/dd").Split('/');

            for (int i = 0; i < strs.Length; i++)
            {
                sb.Append(strs[i]);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 生成一个execl文件,存放刷卡数据
        /// 运行一次生成一个execl文件
        /// </summary>
        /// <param name="dayNum"></param>
        public void CreateSwipeCardExcel(int dayNum)
        {
            string curDate = GetTime_Oracle(0);
            string selectDate = GetTime_Oracle(dayNum);
            DataTable table = null;
            HSSFWorkbook workBook = new HSSFWorkbook();
            table = new DBHelper().QueryOracle(GetSql_GroupSUM(selectDate));
            CreateSheetByCostId(workBook, "部門刷卡數據統計", table);

            table = new DBHelper().QueryOracle(GetSql_ComponentSOLO(selectDate));
            CreateSheetByFourColumn(workBook, "事業群白夜班刷卡數據統計", table, new string[] { "aa", "bb", "cc","dd" }, new string[] { "事業群", "白夜班","有刷卡總人數", "有上下刷總人數" });

            table = new DBHelper().QueryOracle(GetSql_FCSwipeCardEMPByCostID(selectDate));
            CreateSheetByThreeColumn(workBook, "未打卡人員數據統計", table, new string[] { "aa", "bb", "cc" }, new string[] { "費用代碼", "未打卡日期", "未打卡人員總數" });

            table = new DBHelper().QueryOracle(GetSql_FCSwipeCardEMP(selectDate));
            CreateSheetbyEmp(workBook, "未打卡人員信息詳情", table, new string[] { "id", "name", "costid", "depid", "fcDate" }, new string[] { "員工號", "姓名", "費用代碼", "部門代碼", "未打卡日期" });

            string filePathName = string.Format(@"D:\DataToEmailSuper\{0}日A2刷卡人數統計數據.xls", GetDate(dayNum));//存储路径
            if (File.Exists(filePathName))
            {
                File.Delete(filePathName);//文件存在就给他删了，生成新的
            }
            using (FileStream file = new FileStream(filePathName, FileMode.Create))
            {
                workBook.Write(file);　　//创建Excel文件
                file.Close();
            }

        }
        /// <summary>
        /// 生成一个execl文件，存放当天A2在职人员信息
        /// 运行一次生成一个execl文件
        /// </summary>
        /// <param name="dayNum"></param>
        public void CreateOnWorkEmpExcel()
        {
            DataTable table = null;
            HSSFWorkbook workBook = new HSSFWorkbook();
            table = new DBHelper().QueryOracle(GetSql_EMPByCostID());
            CreateSheetbyEmp(workBook, "A2在職人員信息統計", table, new string[] { "costid", "empCount" }, new string[] { "費用代碼", "A2在職人員數量" });

            table = new DBHelper().QueryOracle(GetSql_EMP());
            CreateSheetbyEmp(workBook,  "A2在職人員信息詳情", table, new string[] { "id", "name", "costid", "depid", "depname", "cardid" }, new string[] { "員工號", "姓名", "費用代碼", "部門代碼", "部門名稱", "卡號" });

            string filePathName = string.Format(@"D:\DataToEmailSuper\{0}日A2在職人員信息統計.xls", GetDate(0));//存储路径
            if (File.Exists(filePathName))
            {
                File.Delete(filePathName);//文件存在就给他删了，生成新的
            }
            using (FileStream file = new FileStream(filePathName, FileMode.Create))
            {
                workBook.Write(file);　　//创建Excel文件
                file.Close();
            }

        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="addressee">收件人地址</param>
        public void SendMailUse(string sender, List<string> to_list)
        {
            string error = string.Empty;
            string host = "192.168.78.201";// 邮件服务器smtp.163.com表示网易邮箱服务器    
            string strfrom = sender;// 发送端账号  
            //string password = "doublechen520";// 发送端密码(这个客户端重置后的密码)
            SmtpClient client = new SmtpClient();
            client.DeliveryMethod = SmtpDeliveryMethod.Network;//指定电子邮件发送方式    
            client.Host = host;//邮件服务器
            client.UseDefaultCredentials = true;
            client.Credentials = new System.Net.NetworkCredential(strfrom, "");//用户名、密码                  
           // string strcc = "Yaru_Song@KUNSHAN";//抄送
            string subject = "實時工時系統數據統計";//邮件的主题             
            string body = "";//发送的邮件正文  
            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
            DirectoryInfo folder = new DirectoryInfo(@"D:\DataToEmailSuper");
            StringBuilder sb = new StringBuilder();
            string str = string.Empty;
            foreach (FileInfo file in folder.GetFiles("*.xls"))//获取这个目录下的所有xls格式文件！
            {
                sb.Append(file.FullName + "|");
            }
            str = sb.ToString();
            string newStr = str.Substring(0, str.Length - 1);
            if (newStr.Length != 0)//发送附件（多个附件的文件名称我用 | 隔开的，所以此处这样写）
            {
                string[] arrfile = newStr.Split('|');
                for (int i = 0; i < arrfile.Length; i++)
                {
                    if (arrfile[i].Length > 0)
                    {
                        msg.Attachments.Add(new Attachment(arrfile[i]));
                    }
                }
            }
            //  msg.Attachments.Add(a_file);
            string senderName = to_list[0].Replace("_", " ").Replace("@KUNSHAN", "");
            msg.From = new MailAddress(strfrom, senderName);//后面一参数是发件人名字
           // msg.From = new MailAddress(strfrom, "Yaru Song");//后面一参数是发件人名字
          //  msg.To.Add("Cheng_Qian@KUNSHAN");
            for (int i = 1; i < to_list.Count; i++)
            {
                msg.To.Add(to_list[i]);
            }
            msg.CC.Add(strfrom);
            msg.Subject = subject;//邮件标题   
            msg.Body = body;//邮件内容   
            msg.BodyEncoding = System.Text.Encoding.UTF8;//邮件内容编码   
            msg.IsBodyHtml = false;//是否是HTML邮件   
            msg.Priority = MailPriority.Normal;//邮件优先级  
           try
            {
                client.Send(msg);
                Console.WriteLine("发送成功");
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                Console.WriteLine(ex.Message, "发送邮件出错");
            }    
        }

        public List<string> ReadMail()
        {
            string myPath = @"D:\DataToEmailSuper\addresseeANDsenderInfo.txt";//获取当前程序所在的文件夹

            List<string> list_Str = new List<string>();
            try
            {
                StreamReader sr = new StreamReader(myPath, Encoding.Default);
                string txtStr = string.Empty;
                while ((txtStr = sr.ReadLine()) != null)
                {
                    if (txtStr.Equals(""))
                    {
                        continue;
                    }
                    list_Str.Add(txtStr);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return list_Str;
        }


    }
}
