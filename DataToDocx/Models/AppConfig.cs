// This Source Code Form is subject to the terms of the MIT License.
// If a copy of the MIT was not distributed with this file, You can obtain one at https://opensource.org/licenses/MIT.
// Copyright (C) Leszek Pomianowski and WPF UI Contributors.
// All Rights Reserved.


using System.Data.SQLite;
using System.Data;
using System;
using System.Security.Permissions;
using System.Collections.ObjectModel;
using System.IO;

namespace DataToDocx.Models
{
    public class AppConfig
    {
        public string ConfigurationsFolder { get; set; }

        public string AppPropertiesFileName { get; set; }

        private static string? sqliteConnstr;

       public static string SqliteConnstr
        {
            get { sqliteConnstr ??= $"Data Source={AppDomain.CurrentDomain.BaseDirectory}{DateTime.Now.ToString("yyyyMMddHHmm")}.db";
                return sqliteConnstr; }
            
        }

        public static ObservableCollection<Database> Dbload(bool isNeedTabContent = true)
        {

            var files = new List<string>(Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.db"));

            files.Sort();
            files.Reverse();
            ObservableCollection<Database> dataBases = new ObservableCollection<Database>();
            dataBases.Clear();
            foreach (string file in files)
            {
                dataBases.Add(new Database(Path.GetFileName(file), file));
            }


            for (int i = 0; i < dataBases.Count; i++)
            {
                dataBases[i] = Tabload(dataBases[i].Name, dataBases[i].Path, isNeedTabContent, true);
            }
            //foreach (var item in dataBases)
            //    {
            //    item= Tabload(item.Name,item.Path, isNeedTabContent,true);
            //    } 


            return dataBases;


        }

        public static Database Tabload(string dbname, string dbpath, bool isNeedTabContent = true, bool isNeedTabCount = true)
        {
            Database database = new(dbname, dbpath);
            string connstr = $"Data Source={database.Path}";
            using (var connection = new SQLiteConnection(connstr))
            {
                try
                {
                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }
                    using (SQLiteCommand command = new SQLiteCommand($"SELECT COUNT(*) FROM sqlite_master WHERE type='table' ", connection))
                    {
                        int a = 0;
                        try
                        {
                            a = (int)(long)command.ExecuteScalar();

                        }
                        catch (Exception)
                        {
                            connection.Close();
                            
                            return database;
                        }


                        if (a != 0)
                        {
                            var sqlcommand = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' ", connection);
                            var ds2 = Fun.SqliteToDataset(sqlcommand);
                            //adapter.Fill(ds2);

                            foreach (DataRow row in ds2.Tables[0].Rows)
                            {

                                if (database.Tabs == null)
                                {
                                    database.Tabs = new List<DataTab>
                                    {
                                    new DataTab(row["name"].ToString())
                                    };
                                }
                                else
                                {
                                    database.Tabs.Add(new DataTab(row["name"].ToString()));
                                }

                            }

                            if (isNeedTabCount)


                            {

                                if (database.Tabs == null)
                                {
                                    
                                    return database;

                                }
                                foreach (DataTab dataTab in database.Tabs)
                                {
                                    using (SQLiteCommand command2 = new SQLiteCommand($"SELECT COUNT(*) FROM '{dataTab.TabName}' ", connection))
                                    {
                                        dataTab.Count = (int)(long)command2.ExecuteScalar();
                                    }
                                    //using (SQLiteCommand command3 = new SQLiteCommand($"SELECT COUNT(*) FROM '{dataTab.TabName}' WHERE `操作状况`='' and `错误状况`=''", connection))
                                    //{
                                    //    dataTab.LeftTaskCount = (int)(long)command3.ExecuteScalar();
                                    //}
                                    //using (SQLiteCommand command3 = new SQLiteCommand($"SELECT COUNT(*) FROM '{dataTab.TabName}' WHERE `错误状况`='是'", connection))
                                    //{
                                    //    dataTab.ErrorTaskCount = (int)(long)command3.ExecuteScalar();
                                    //}
                                    //using (SQLiteCommand command3 = new SQLiteCommand($"SELECT COUNT(*) FROM '{dataTab.TabName}' WHERE `操作状况`='成功'", connection))
                                    //{
                                    //    dataTab.TypedTasksCount = (int)(long)command3.ExecuteScalar();
                                    //}

                                }
                            }

                            if (isNeedTabContent)
                            {
                                if (database.Tabs == null)
                                {
                                    
                                    return database;
                                }
                                foreach (DataTab dataTab in database.Tabs)
                                {

                                    using (SQLiteCommand command2 = new SQLiteCommand($"SELECT COUNT(*) FROM '{dataTab.TabName}'", connection))
                                    {
                                        dataTab.Count = (int)(long)command2.ExecuteScalar();
                                    }

                                    var cmd = new SQLiteCommand($"SELECT * FROM pragma_table_info('{dataTab.TabName}')", connection);
                                    var ds3 = Fun.SqliteToDataset(cmd);
                                    //adp.Fill(ds3);
                                    foreach (DataRow row in ds3.Tables[0].Rows)
                                    {
                                        if (dataTab.Atts == null)
                                        {
                                            dataTab.Atts = new List<string>
                                    {
                                        row["name"].ToString()
                                    };
                                        }
                                        else
                                        {
                                            dataTab.Atts.Add(row["name"].ToString());
                                        }
                                    }
                                    //var adapter3 = new SQLiteDataAdapter($"SELECT * FROM `{dataTab.TabName}` LIMIT 500  ", connstr);
                                    var cmd2 = new SQLiteCommand($"SELECT * FROM `{dataTab.TabName}`  ", connection);
                                    var ds4 = Fun.SqliteToDataset(cmd2);
                                    //adapter3.Fill(ds4);
                                    dataTab.TabContent = ds4.Tables[0];
                                }
                            }



                        }
                    }
                    connection.Close();
                    return database;
                }
                catch (Exception e)
                {
                    
                    return database;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
    }
}
