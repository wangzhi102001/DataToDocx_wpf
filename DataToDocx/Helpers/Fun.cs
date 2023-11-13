using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Data.SQLite;
using System.Windows.Threading;
using Microsoft.Win32;
using MiniExcelLibs;
using System.Runtime.InteropServices;
using System.IO;
using CommunityToolkit.Mvvm.Messaging;
using Wpf.Ui.Controls;

namespace DataToDocx
{
   public class Fun
    {
        
        public static void UpdateState(string connstr, string tablename, string huid)
        {
            using (var connection = new SQLiteConnection(connstr))
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.Open();
                }

                //string query = $"UPDATE `{tablename}` SET `操作状况`=@state WHERE `户编号`=@huID;";
                string query = $"UPDATE `{tablename}` SET `操作状况`=@state WHERE `户主身份证号`=@huID;";
                //户主身份证号


                SQLiteCommand command = new SQLiteCommand(query, connection);

                command.Parameters.AddWithValue("@state", $"{DateTime.Now}已生成" ?? "");
                command.Parameters.AddWithValue("@huID", huid ?? "");
                command.ExecuteNonQuery();
            }
        }

        public static void Updatelogtext(string log)
        { // 检查标签是否从另一个线程访问
            WeakReferenceMessenger.Default.Send<string,string>(DateTime.Now.ToString() + " " + log,"log");


        }

        private static List<string> GetHeadKey(IEnumerable<IDictionary<string, object>> rows, int skip_num)
        {
            var ziduancheck = new List<string>();//字段集合
            ziduancheck.AddRange(rows.Skip(skip_num).First().Where(em => em.Value != null && em.Value.ToString() != "").Select(em => em.Value.ToString()));
            return ziduancheck;
        }

        private static string GetHeaderQuery(List<string> ziduancheck)
        {
            string sql_creat_part1 = "";
            foreach (var item in ziduancheck)
            {
                sql_creat_part1 += " `" + item + "` TEXT,";

            }
            string sql_creat1 = "CREATE TABLE IF NOT EXISTS TABLEA" + "(" + sql_creat_part1.Substring(0, sql_creat_part1.Length - 1) + ");";
            return sql_creat1;
        }

        /// <summary>
		/// 使用sqlite3.dll关闭文件同步、开启事务、对语句进行执行准备、开启wal模式，极速导入sqlite。
		/// </summary>
		/// <param name="filePath">excel文件目录</param>
		/// <param name="tablename">导入数据库的表名</param>
		/// <param name="connstr">数据库连接字符</param>
		/// <param name="label">用于跟踪导入计数的label</param>
		public int Import_Excel_plus(string filePath, string tablename, string connstr, Label label, bool ziduan_checkstate = true)
        {

            UTF8Encoding utf8 = new UTF8Encoding();
            int state = 0;

            //bool ziduan_checkstate = true;
            //ziduan_checkstate = Ziduan_Check(filePath, tablename);

            using (var stream = File.OpenRead(filePath))
            {
                var rows = stream.Query(excelType: ExcelType.XLSX).Cast<IDictionary<string, object>>();
                var columns = MiniExcel.GetColumns(filePath);
                int skip_num = GetHeadRow(rows);
                int ziduan_num = rows.Skip(skip_num).First().ToArray().Count();//字段数

                List<string> ziduancheck = GetHeadKey(rows, skip_num);
                string sql_creat1 = GetHeaderQuery(ziduancheck);//获取创建表的sql语句



                if (ziduan_checkstate)
                {

                    //Updatelogtext($"开始导入【{tablename}】表，导入文件路径{filePath}。", RTB_log);
                    //Updatelogtext($"【{tablename}】导入文件的表头为第{(skip_num + 1)}行。", RTB_log);
                    int cdt = ToSqlite(tablename, connstr, label, rows, skip_num, ziduancheck, sql_creat1);
                    if (cdt == 1)
                    {
                        state = 1;
                        return state;
                    }
                    else
                    {
                        state = 0;
                        return state;
                    }
                }
                else
                {
                    state = 2;
                    return state;
                }
            }
        }
        private static int GetHeadRow(IEnumerable<IDictionary<string, object>> rows)
        {
            var firstRow = rows.FirstOrDefault();
            var secondRow = rows.Skip(1).FirstOrDefault();
            var thirdRow = rows.Skip(2).FirstOrDefault();
            var fourR = rows.Skip(3).FirstOrDefault();
            int skip_num = 0;
            int firstcount = 0;
            int secondcount = 0;
            int thirdcount = 0;
            int fourthcount = 0;

            foreach (var item in firstRow.ToArray())
            {
                if (item.Value != null && item.Value.ToString() != "")
                {
                    firstcount++;
                }
            }

            foreach (var item in secondRow.ToArray())
            {
                if (item.Value != null && item.Value.ToString() != "")
                {
                    secondcount++;
                }
            }

            foreach (var item in thirdRow.ToArray())
            {
                if (item.Value != null && item.Value.ToString() != "")
                {
                    thirdcount++;
                }
            }

            foreach (var item in fourR.ToArray())
            {
                if (item.Value != null && item.Value.ToString() != "")
                {
                    fourthcount++;
                }
            }

            if (Math.Max(Math.Max(Math.Max(firstcount, secondcount), thirdcount), fourthcount) == firstcount)
            {
                skip_num = 0;
            }
            else if (Math.Max(Math.Max(Math.Max(firstcount, secondcount), thirdcount), fourthcount) == secondcount)
            {
                skip_num = 1;
            }
            else if (Math.Max(Math.Max(Math.Max(firstcount, secondcount), thirdcount), fourthcount) == thirdcount)
            {
                skip_num = 2;
            }
            else { skip_num = 3; }

            return skip_num;
        }
        private int ToSqlite(string tablename, string connstr, Label label, IEnumerable<IDictionary<string, object>> rows, int skip_num, List<string> ziduancheck, string sql_creat1)
        {
            int condition = 1;
            CreateNewTable(tablename, connstr, sql_creat1);
            //int count_begin = 1;
            int loadcount = Sqlite_plus(tablename, connstr, label, rows, ziduancheck, skip_num);
            if (loadcount != 0)
            {
                ChangeTableName(tablename, connstr);

                Updatelabel("√ 共导入" + loadcount + "条", label);

                //Updatelogtext("【" + tablename + "】导入完成，共导入" + loadcount + "条。", RTB_log);
                return condition;
            }
            else
            {


                Updatelabel("× 导入失败", label);

                //Updatelogtext("【" + tablename + "】导入完成，共导入" + loadcount + "条。", RTB_log);
                condition = 0;
                return condition;

            }
        }

        void Updatelabel(string text, Label label)
        { // 检查标签是否从另一个线程访问
            //if (label.InvokeRequired)
            //{
            //    // 在与标签相同的线程上调用方法
            //    label.Invoke(new Action<string, Label>(Updatelabel), text, label);
            //}
            //else
            //{
            //    // 设置标签文本
            //    label.Text = text;
            //}

            Dispatcher.CurrentDispatcher.Invoke(() =>
            {
                label.Content = text;

            });

        }

        

        private static void ChangeTableName(string tablename, string connstr)
        {
            using (var connection = new SQLiteConnection(connstr))//新建表
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.Open();
                }

                string sql_shift = $"ALTER TABLE TABLEA RENAME TO `{tablename}`;";
                SQLiteCommand command1 = new SQLiteCommand(sql_shift, connection);

                command1.ExecuteNonQuery();
                connection.Close();

            }
        }


        private static void CreateNewTable(string tablename, string connstr, string sql_creat1)
        {
            using (var connection = new SQLiteConnection(connstr))//新建表
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.Open();
                }
                new SQLiteCommand($"DROP TABLE IF EXISTS TABLEA;", connection).ExecuteNonQuery();
                new SQLiteCommand($"DROP TABLE IF EXISTS `{tablename}`;", connection).ExecuteNonQuery();
                SQLiteCommand command1 = new SQLiteCommand(sql_creat1, connection);
                command1.ExecuteNonQuery();
                connection.Close();
            }
        }

        private int Sqlite_plus(string tablename, string connstr, Label label, IEnumerable<IDictionary<string, object>> rows, List<string> ziduancheck, int skip_num)
        {
            int loadcount = 0;
            //声明两个指针
            IntPtr db;
            string connstr_utf8 = "";
            try
            {
                connstr_utf8 = $"{connstr.Replace("Data Source=", "").Replace(AppDomain.CurrentDomain.BaseDirectory, ".")}";
            }
            catch (Exception)
            {
                connstr_utf8 = $"{connstr.Replace("Data Source=", "")}";
            }

            try
            {
                sqlite3_open(connstr_utf8, out db);
                sqlite3_exec(db, "PRAGMA synchronous = OFF;", 0, IntPtr.Zero, "");
                sqlite3_exec(db, "begin;", 0, IntPtr.Zero, "");

                foreach (var row in rows.Skip(skip_num + 1))
                {
                    //开始事务
                    IntPtr P_prepare;
                    //sqlite3_exec(db, "begin;", 0, IntPtr.Zero, "");
                    IntPtr PzTail = new IntPtr(0);
                    string manyask = string.Join(",", Enumerable.Repeat("?", ziduancheck.Count));
                    string sql_plus = $"insert into TABLEA values({manyask});";
                    sqlite3_prepare_v2(db, sql_plus, sql_plus.Length, out P_prepare, PzTail);
                    //重置
                    sqlite3_reset(P_prepare);
                    int id_num = 1;
                    foreach (var em in row)
                    {
                        if (em.Value != null)
                        {
                            sqlite3_bind_text(P_prepare, id_num, Encoding.UTF8.GetBytes(em.Value.ToString()), Encoding.UTF8.GetBytes(em.Value.ToString()).Length, IntPtr.Zero);
                        }
                        else
                        {
                            sqlite3_bind_text(P_prepare, id_num, Encoding.UTF8.GetBytes(" "), Encoding.UTF8.GetBytes(" ").Length, IntPtr.Zero);
                        }
                        id_num++;
                    }
                    sqlite3_step(P_prepare);
                    sqlite3_finalize(P_prepare);

                    loadcount++;
                    if (loadcount % 1000 == 0)
                    {

                        sqlite3_exec(db, "commit;", 0, IntPtr.Zero, "");
                        //sqlite3_finalize(P_prepare);
                        sqlite3_free(PzTail);
                        sqlite3_exec(db, "begin;", 0, IntPtr.Zero, "");

                        //Updatelabel(tablename + "\r\n已导入" + loadcount + "条", label);


                    }

                }

                sqlite3_exec(db, "commit;", 0, IntPtr.Zero, "");
                sqlite3_close(db);
            }
            catch (System.AccessViolationException)
            {

                loadcount = 0;
            }
            return loadcount;
        }


        ///绑定sqlite3.dll数据
        ///
        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_open", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_open(string filename, out IntPtr db);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_close", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_close(IntPtr db);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_prepare_v2", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_prepare_v2(IntPtr db, string zSql,
        int nByte, out IntPtr ppStmpt, IntPtr pzTail);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_step", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_step(IntPtr stmHandle);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_reset", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_reset(IntPtr stmHandle);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_finalize", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_finalize(IntPtr stmHandle);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_free", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_free(IntPtr stmHandle);


        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_exec", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_exec(IntPtr db, string zSql, int funcptr, IntPtr funcparm, string msg);

        //这个绑定是一系列的，有好多，按照不同的类型用不同的绑定
        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_bind_int", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_bind_int(IntPtr stmHandle, int clmindex, int value);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_bind_double", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_bind_double(IntPtr stmHandle, int clmindex, double value);

        [DllImport("sqlite3.dll", EntryPoint = "sqlite3_bind_text", CallingConvention = CallingConvention.Cdecl)]
        static extern int sqlite3_bind_text(IntPtr stmHandle, int clmindex, byte[] value, int valuelen, IntPtr funcparm);
        private void Create_Click(object sender, EventArgs e)
        {



        }
        void EnCtrl()
        {
            //textBox_zhen.Enabled = true;
            //textBox_cun.Enabled = true;
            //btn_upload_bf.Enabled = true;
            //btn_upload_qzd.Enabled = true;
            //Create.Enabled = true;
        }

        void DisCtrl()
        {
            //textBox_zhen.Enabled = false;
            //textBox_cun.Enabled = false;
            //btn_upload_bf.Enabled = false;
            //btn_upload_qzd.Enabled = false;
            //Create.Enabled = false;
        }

        private void InputTask(string path, string tablename, string connstr, Label label)
        {
            if (path != "")
            {

                bool ziduancheckstate = true;



                try
                {
                    if (ziduancheckstate)
                    {
                        //Updatelogtext($"【{tablename}】导入任务加入队列。", RTB_log);
                    }
                    Import_Excel_plus(path, tablename, connstr, label, ziduancheckstate);
                }
                catch (Exception e)
                {
                    //MessageBox.Show($"【{tablename}】导入失败！\r" + ex.Message);
                    //Updatelogtext($"【{tablename}】导入失败。失败原因{e.Message}", RTB_log);
                }


            }
            else
            {
               System.Windows. MessageBox.Show("文件路径为空，请选择。");
            }
        }


        public class TaskQueue
        {
            /// <summary>
            /// 任务队列
            /// </summary>
            private Queue<TaskData> QueuesTask = new Queue<TaskData>();
            /// <summary>
            /// 任务队列是否在执行中
            /// </summary>
            private bool isExecuteing = false;

            private static TaskQueue _instance = null;
            public static TaskQueue Instance
            {
                get
                {
                    if (_instance == null)
                        _instance = new TaskQueue();
                    return _instance;
                }
            }

            /// <summary>
            /// 任务是否进行中
            /// </summary>
            /// <returns></returns>
            public bool IsTasking()
            {
                return isExecuteing;
            }

            /// <summary>
            /// 添加任务，任务会按照队列自动执行
            /// </summary>
            /// <param name="task"></param>
            public void AddTaskAndRuning(TaskData taskData)
            {
                if (taskData == null) return;

                QueuesTask.Enqueue(taskData);

                StartPerformTask();
            }

            /// <summary>
            /// 执行任务
            /// </summary>
            private async void StartPerformTask()
            {
                if (isExecuteing) { return; }

                while (QueuesTask.Count > 0)
                {
                    isExecuteing = true;
                    await Task.Run(() =>
                    {
                        TaskData taskDatas = QueuesTask.Dequeue();
                        Task task = taskDatas.Tasks;
                        task.Start();
                        task.Wait();
                        if (taskDatas.CallBack != null) taskDatas.CallBack(null);
                    });
                }

                isExecuteing = false;
            }

            private TaskQueue()
            {
            }
        }

        public class TaskData
        {
            /// <summary>
            /// 任务名
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 任务
            /// </summary>
            public Task Tasks { get; set; }
            /// <summary>
            /// 任务完成后的回调
            /// </summary>
            public Action<string> CallBack { get; set; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Title"></param>
        /// <param name="Message"></param>
        /// <param name="Type">1=Error,2=Info,3=Caution,4=Success</param>
        public static void ShowSnackbar(string Title,string Message, int Type)
        {
            switch (Type)
            {
                case 1:
                    WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message
            }, "snackbarError");
                    break;
                case 2:
                    WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message
            }, "snackbarInfo");
                    break;
                case 3:
                    WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message
            }, "snackbarCaution");
                    break;
                case 4:
                    WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message
            }, "snackbarSuccess");
                    break;
                default:
                    break;
            }

            
        }

        public static void ShowDialog(string Title,string Message,string CloseButtonText)
        {
            WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message,CloseButtonText
            }, "dialogAlart");
        }
        //public static void ShowContentDialog(string Title, string Message, string CloseButtonText,string PrimaryButtonText="",string SecondaryButtonText = "")
        //{
        //    WeakReferenceMessenger.Default.Send<SimpleContentDialogCreateOptions, string>(new SimpleContentDialogCreateOptions()
        //    {
        //        Title=Title,
        //        Content=Message,
        //        CloseButtonText = CloseButtonText,
        //        PrimaryButtonText = PrimaryButtonText,
        //        SecondaryButtonText = SecondaryButtonText

        //    }, "dialogContent");
        //}
        public static DataSet SqliteToDataset(SQLiteCommand sqliteCommand)
        {
            var dt = new DataTable();
            using (sqliteCommand)
            {
                using (var reader = sqliteCommand.ExecuteReader())
                {
                    dt.Load(reader);
                }
            }
            var ds2 = new DataSet();
            ds2.Tables.Add(dt);
            return ds2;
        }

    }



}

