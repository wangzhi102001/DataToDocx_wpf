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
using Microsoft.Data.Sqlite;
using System.Diagnostics.Metrics;

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
            StringBuilder sql_creat = new StringBuilder();

            sql_creat.Append($"CREATE TABLE IF NOT EXISTS TABLEA (");
            for (int i = 0; i < ziduancheck.Count; i++)
            {
                if (i!= ziduancheck.Count-1)
                {
                    sql_creat.Append($" `{ziduancheck[i]}` TEXT,");
                }
                else
                {
                    sql_creat.Append($" `{ziduancheck[i]}` TEXT");
                }
            }


            //foreach (var item in ziduancheck)
            //{
            //    sql_creat_part1.Append($" `{item}` TEXT,");

            //}
            sql_creat.Append(");");
            
            return sql_creat.ToString();
        }

        public static bool ExcelToSqlite(string excelFilePath,string tableName,string connstr,out int progress)
        {

            int countA = 0;
            //TODO 读取excel获取rows
            try
            {

                using (var stream = File.OpenRead(excelFilePath))
                {
                    var rows = stream.Query(excelType: ExcelType.XLSX).Cast<IDictionary<string, object>>();
                    if (rows == null)
                    {
                        progress = 0;
                        return false;
                    }
                    int skip_num = GetHeadRow(rows);
                    int ziduan_num = rows.Skip(skip_num).First().ToArray().Count();//字段数
                    List<string> attrs = GetHeadKey(rows, skip_num);//字段列表
                    StringBuilder sql_creat = new StringBuilder();
                    sql_creat.Append($"CREATE TABLE IF NOT EXISTS `{tableName}` (");
                    for (int i = 0; i < attrs.Count; i++)
                    {
                        if (i != attrs.Count - 1)
                        {
                            sql_creat.Append($" `{attrs[i]}` TEXT,");
                        }
                        else
                        {
                            sql_creat.Append($" `{attrs[i]}` TEXT");
                        }
                    }
                    sql_creat.Append(");");

                    using (var connection = new SQLiteConnection(connstr))//新建表
                    {
                        if (connection.State == ConnectionState.Closed)
                        {
                            connection.Open();
                        }

                        new SQLiteCommand($"DROP TABLE IF EXISTS `{tableName}`;", connection).ExecuteNonQuery();
                        SQLiteCommand command1 = new SQLiteCommand(sql_creat.ToString(), connection);
                        command1.ExecuteNonQuery();//建立空表
                        connection.Close();
                    }

                    
                    progress = 0;
                    using (var connection = new SQLiteConnection(connstr))//新建表
                    {
                        if (connection.State == ConnectionState.Closed)
                        {
                            connection.Open();
                        }
                        StringBuilder insertsql = new StringBuilder();
                        insertsql.Append($"insert into `{tableName}` values(");
                        for (int i = 0; i < attrs.Count; i++)
                        {
                            if (i != attrs.Count - 1)
                            {
                                insertsql.Append($"$value{i},");
                            }
                            else
                            {
                                insertsql.Append($"$value{i})");
                            }
                        }
                        var transaction = connection.BeginTransaction();

                        var command = connection.CreateCommand();
                        command.CommandText = insertsql.ToString();
                        //List<SqliteParameter> sqliteParameters = new List<SqliteParameter>();
                        for (int i = 0; i < attrs.Count; i++)
                        {
                            var parameter = command.CreateParameter();
                            parameter.ParameterName = $"$value{i}";
                            command.Parameters.Add(parameter);
                        }


                        //var transaction = connection.BeginTransaction();
                        foreach (var row in rows.Skip(skip_num + 1))
                        {
                            int idx = 0;
                            foreach (var em in row)
                            {
                                if (idx == attrs.Count)
                                {
                                    break;
                                }
                                if (em.Value != null)
                                {
                                    command.Parameters[idx].Value = em.Value.ToString();
                                }
                                else
                                {
                                    command.Parameters[idx].Value = "";
                                }
                                idx++;
                            }
                            command.Prepare();
                            command.ExecuteNonQuery();
                            countA++;
                            
                            if (countA%500==0)
                            {
                                progress = countA;
                                
                            }


                        }
                        progress = countA;
                        transaction.Commit();
                        transaction.Dispose();
                        connection.Close();

                    }
                }
                Dispatcher.CurrentDispatcher.Invoke(() =>
                {
                ShowSnackbar("导入完成", $"已导入{countA}条数据。", 1);
                }); 
                return true;
            }
            catch (Exception ex)
            {
                Dispatcher.CurrentDispatcher.Invoke(() =>
                {
                    ShowSnackbar("导入错误", $"错误提示：{ex.Message}", 1);
                });
                
                progress = 0;
                return false;
                
            }
        }



       
        /// <summary>
        /// 根据excel的流式rows获取字段行
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
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
            int loadcount = 0; //Sqlite_plus(tablename, connstr, label, rows, ziduancheck, skip_num);
            if (loadcount != 0)
            {
                ChangeTableName(tablename, connstr);

                Updatelabel("√ 共导入" + loadcount + "条", label);

                //Updatelogtext("【" + tablename + "】导入完成，共导入" + loadcount + "条。", RTB_log);
                return condition;
            }
            else
            {                Updatelabel("× 导入失败", label);

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
                    //Import_Excel_plus(path, tablename, connstr, label, ziduancheckstate);
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
            Application.Current.Dispatcher.Invoke(()=>{
WeakReferenceMessenger.Default.Send<List<string>, string>(new List<string>()
            {
                Title, Message,CloseButtonText
            }, "dialogAlart");
        }});
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

