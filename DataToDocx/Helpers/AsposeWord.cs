using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Data.Entity;
using System.Net.NetworkInformation;


namespace DataToDocx
{
    internal class AsposeWord
    {
        /// <summary>
        /// 通过WORD模板导出文档
        /// </summary>
        public class WordModelEvent
        {
            /// <summary>
            /// 根据Word模板进行数据的导出Word操作
            /// </summary>
            /// <param name="modelDoc">模板路径</param>
            /// <param name="exptDoc">导出文件名</param>
            /// <param name="fieldNames">字段名字符串数组</param>
            /// <param name="fieldValues">字段值数组</param>
            /// <param name="dt">需要循环的数据DataTable</param>
            /// <returns>返回的文档路径</returns>
            public static string ExpertWordToModel(string modelDoc, string exptDoc, DataTable maindt, DataTable dt = null, DataTable dt2 = null, DataTable dt3 = null, DataTable dt4 = null, DataTable dt5 = null)
            {
                //try
                //{
                removeWatermark();//加载监听,用于去除水印
                                  //string tempPath = HttpContext.Current.Server.MapPath("~/Demos/" + modelDoc + ".docx");//word模板路径
                string tempPath = "" + modelDoc + ".docx";
                //导出的WORD存放的位置
                const string saveFold = "word/";
                //string outputPath = HttpContext.Current.Server.MapPath("~/" + saveFold);
                string outputPath = "" + saveFold;
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }
                //生成的WORD文件名
                string fileName = exptDoc + ".docx";
                outputPath += fileName;
                //载入模板
                Document doc = new Document(tempPath);
                doc.MailMerge.Execute(maindt);
                
                //doc.MailMerge.Execute()
                //将我们获得的DataTable类型的数据：EduDataTable放入doc方法中做处理
                if (dt != null && dt.Rows.Count > 0)
                {
                    doc.MailMerge.ExecuteWithRegions(dt);
                }
                if (dt2 != null && dt2.Rows.Count > 0)
                {
                    doc.MailMerge.ExecuteWithRegions(dt2);
                }
                if (dt3 != null && dt3.Rows.Count > 0)
                {
                    doc.MailMerge.ExecuteWithRegions(dt3);
                }
                if (dt4 != null && dt4.Rows.Count > 0)
                {
                    doc.MailMerge.ExecuteWithRegions(dt4);
                }
                if (dt5 != null && dt5.Rows.Count > 0)
                {
                    doc.MailMerge.ExecuteWithRegions(dt5);
                }


                //获取下载地址
                String StrVisitURL = saveFold + fileName;
                //合并模版，相当于页面的渲染
                doc.MailMerge.Execute(new[] { "PageCount" }, new object[] { doc.PageCount });
                //保存合并后的文档
                doc.Save(outputPath);
                //MessageBox.Show($"生成完成{outputPath}");
                return StrVisitURL;
                //}
                //catch (Exception er)
                //{
                //    MessageBox.Show($"{er.Message}");
                //    return er.Message;
                //}
            }

            /// <summary>
            /// 加载监听,用于去除水印
            /// </summary>
            public static void removeWatermark()
            {
                new Aspose.Words.License().SetLicense(new MemoryStream(Convert.FromBase64String("PExpY2Vuc2U+CiAgPERhdGE+CiAgICA8TGljZW5zZWRUbz5TdXpob3UgQXVuYm94IFNvZnR3YXJlIENvLiwgTHRkLjwvTGljZW5zZWRUbz4KICAgIDxFbWFpbFRvPnNhbGVzQGF1bnRlYy5jb208L0VtYWlsVG8+CiAgICA8TGljZW5zZVR5cGU+RGV2ZWxvcGVyIE9FTTwvTGljZW5zZVR5cGU+CiAgICA8TGljZW5zZU5vdGU+TGltaXRlZCB0byAxIGRldmVsb3BlciwgdW5saW1pdGVkIHBoeXNpY2FsIGxvY2F0aW9uczwvTGljZW5zZU5vdGU+CiAgICA8T3JkZXJJRD4yMDA2MDIwMTI2MzM8L09yZGVySUQ+CiAgICA8VXNlcklEPjEzNDk3NjAwNjwvVXNlcklEPgogICAgPE9FTT5UaGlzIGlzIGEgcmVkaXN0cmlidXRhYmxlIGxpY2Vuc2U8L09FTT4KICAgIDxQcm9kdWN0cz4KICAgICAgPFByb2R1Y3Q+QXNwb3NlLlRvdGFsIGZvciAuTkVUPC9Qcm9kdWN0PgogICAgPC9Qcm9kdWN0cz4KICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl0aW9uVHlwZT4KICAgIDxTZXJpYWxOdW1iZXI+OTM2ZTVmZDEtODY2Mi00YWJmLTk1YmQtYzhkYzBmNTNhZmE2PC9TZXJpYWxOdW1iZXI+CiAgICA8U3Vic2NyaXB0aW9uRXhwaXJ5PjIwMjEwODI3PC9TdWJzY3JpcHRpb25FeHBpcnk+CiAgICA8TGljZW5zZVZlcnNpb24+My4wPC9MaWNlbnNlVmVyc2lvbj4KICAgIDxMaWNlbnNlSW5zdHJ1Y3Rpb25zPmh0dHBzOi8vcHVyY2hhc2UuYXNwb3NlLmNvbS9wb2xpY2llcy91c2UtbGljZW5zZTwvTGljZW5zZUluc3RydWN0aW9ucz4KICA8L0RhdGE+CiAgPFNpZ25hdHVyZT5wSkpjQndRdnYxV1NxZ1kyOHFJYUFKSysvTFFVWWRrQ2x5THE2RUNLU0xDQ3dMNkEwMkJFTnh5L3JzQ1V3UExXbjV2bTl0TDRQRXE1aFAzY2s0WnhEejFiK1JIWTBuQkh1SEhBY01TL1BSeEJES0NGbWg1QVFZRTlrT0FxSzM5NVBSWmJRSGowOUNGTElVUzBMdnRmVkp5cUhjblJvU3dPQnVqT1oyeDc4WFE9PC9TaWduYXR1cmU+CjwvTGljZW5zZT4=")));
            }



            /// <summary>
            /// 文件名
            /// </summary>
            public static string fileName = "";
            /// <summary>
            /// 根据WORD模板基本信息
            /// </summary>
            /// <param name="HuID">需要循环的子数据的所属上级ID</param>
            /// <returns>返回完成的文件下载路径</returns>
            public static string GetWordEvent(RichTextBox richTextBox,string zhen_filter,string cun_filter,string HuID = "")
            {
                try
                {

                    

                    //"Data Source=" + Application.StartupPath + "\\" + DateTime.Now.ToString("yyyyMMddHH") + ".db";
                    string connstr = $@"Data Source={AppDomain.CurrentDomain.BaseDirectory}{"database"}.db";
                    //这里是实例化Model，如果项目已有封装的查询数据库信息的方法，直接使用参数“id”进行动态数据查询即可
                    //MemberUnit member = new MemberUnit();
                    try
                    {
                        string sqlquery1 = "alter table `全县全字段` add column `操作状况` TEXT";

                        getDataTable(sqlquery1, connstr);
                    }
                    catch (Exception)
                    {


                    }
                    string sqlquery;
                    if (zhen_filter != "" || cun_filter != "")
                    {
                        if (zhen_filter != "" && cun_filter == "")
                        {
                            //sqlquery = $"SELECT `县`,`乡`,`村`,`户联系电话`,`户类型`,`监测对象类别`,`是否军烈属`,`工资性收入`,`转移性收入`,`养老金或离退休金`,`生产经营性收入`,`计划生育金`,`生态补偿金`,`财产性收入`,`最低生活保障金`,`其他转移性收入`,`资产收益扶贫分红收入`,`特困人员救助供养金`,`其他财产性收入（旧指标）`,`生产经营性支出`,`纯收入（元）`,`当前家庭人口数`,`人均纯收入（元）`,`年度家庭人口数`,`耕地面积（亩）`,`牧草地面积`,`水面面积`,`林地面积（亩）`,`退耕还林面积（亩）`,`林果面积（亩）`,`入户路类型`,`与村主干路距离`,`是否加入农民专业合作组织`,`是否危房户`,`住房面积`,`主要燃料类型`,`是否有龙头企业带动`,`是否有创业致富带头人带动`,`是否通生活用电`,`是否通广播电视`,`到户产业帮扶类型`,`是否解决安全饮用水`,`是否通生产用电`,`户编号`,`人口编号`,`姓名`,`证件号码`,`与户主关系`,`公益性岗位收入`,`其他工资性收入`,`产业奖励`,`就业奖励`,`生产经营性支出（合计）`,`专项用于减少生产经营性支出的补贴`,`年收入（元）`,`是否有卫生厕所`,`户主姓名`,`户主证件号码` FROM `全县全字段`  WHERE `与户主关系`='户主' and `乡`='{zhen_filter}' and `操作状况`is null  LIMIT 1 ;";
                            sqlquery = $"SELECT * FROM `全县全字段`  WHERE `与户主关系`='户主' and `镇`='{zhen_filter}' and `操作状况`is null  LIMIT 1 ;";
                        }
                        else if (zhen_filter =="" && cun_filter != "")
                        {
                            sqlquery = $"SELECT * FROM `全县全字段`  WHERE `与户主关系`='户主' and `村`='{cun_filter}' and `操作状况`is null LIMIT 1 ;";

                        }
                        else
                        {
                            sqlquery = $"SELECT * FROM `全县全字段`  WHERE `与户主关系`='户主'and `村`='{cun_filter}'and `镇`='{zhen_filter}' and `操作状况`is null  LIMIT 1 ;";
                        }

                    }
                    else
                    {
                        sqlquery = $"SELECT * FROM `全县全字段`  WHERE `与户主关系`='户主' and `操作状况`is null  LIMIT 1 ;";
                    }
                    Fun.Updatelogtext($"正在获取信息......");
                    DataTable dtmain = getDataTable(sqlquery, connstr);

                    string xiang = dtmain.Rows[0]["镇"].ToString();
                    string cun = dtmain.Rows[0]["村"].ToString();
                    string huzhuName = dtmain.Rows[0]["户主"].ToString();
                    string huzhuID = dtmain.Rows[0]["户主身份证号"].ToString();
                    //MemberUnit member = Fun.FetchOneHHFromDB(connstr, "全县全字段");
                    Fun.Updatelogtext($"正在获取{xiang}-{cun}-{huzhuName}信息......");

                    //if (member.huID == null || member.huID == "")
                    //{

                    //    return "已生成完成";
                    //}

                    if (dtmain.Rows.Count==0)
                    {

                        return "已生成完成";
                    }

                    //HuID = dtmain.Rows[0]["户主身份证号"].ToString();
                    #region //主字段
                    //Object[] fieldValues = new object[fieldNames.Length];
                    //fieldValues[0] = member.xian;
                    //fieldValues[1] = member.xiang;
                    //fieldValues[2] = member.cun;
                    //fieldValues[3] = member.huTEL;
                    //fieldValues[4] = member.huleixing;
                    //fieldValues[5] = member.jianceleix;
                    //fieldValues[6] = member.isjunlieshi;
                    //fieldValues[7] = member.gongzixingshouruheji;
                    //fieldValues[8] = member.zhuanyixingshouruheji;
                    //fieldValues[9] = member.yanglaojinhuotuixiujin;
                    //fieldValues[10] = member.shengchanjingyingxingshouru;
                    //fieldValues[11] = member.jihuashengyujin;
                    //fieldValues[12] = member.shengtaibuchangjin;
                    //fieldValues[13] = member.caichanxingshouru;
                    //fieldValues[14] = member.zuidishenghuobaozhangjin;
                    //fieldValues[15] = member.qitazhuanyixingshouru;
                    //fieldValues[16] = member.zichanshouyifenhong;
                    //fieldValues[17] = member.tekunrenyuanjiuzhugongyangjin;
                    //fieldValues[18] = member.qitacaichanxing;
                    //fieldValues[19] = member.shengchanxzhichu;
                    //fieldValues[20] = member.jiatingchunshouru;
                    //fieldValues[21] = member.jiatingrenkoushu;
                    //fieldValues[22] = member.renjunchunshouru;
                    //fieldValues[23] = member.niandujiatingrenkoushu;
                    //fieldValues[24] = member.gengdimianji;
                    //fieldValues[25] = member.mucaomianji;
                    //fieldValues[26] = member.shuimianmianji;
                    //fieldValues[27] = member.lindimianji;
                    //fieldValues[28] = member.tuigenghuancaomianji;
                    //fieldValues[29] = member.linguomianji;
                    //fieldValues[30] = member.ruhudaolu;
                    //fieldValues[31] = member.cunzhuganjuli;
                    //fieldValues[32] = member.nongminzhuanyehezuoshe;
                    //fieldValues[33] = member.weifang;
                    //fieldValues[34] = member.zhufangmianji;
                    //fieldValues[35] = member.ranliaoleix;
                    //fieldValues[36] = member.longtouqiyedaidong;
                    //fieldValues[37] = member.chuangyezhifudaitouren;
                    //fieldValues[38] = member.tongshenghyongdian;
                    //fieldValues[39] = member.tongguangbodianshi;
                    //fieldValues[40] = member.chanyebangfuleixing;
                    //fieldValues[41] = member.yinshuianquan;
                    //HuID = member.huID;
                    //                    "xian","xiang","cun","huTEL","huleixing","jianceleix","isjunlieshi","gongzixingshouruheji","zhuanyixingshouruheji","yanglaojinhuotuixiujin","shengchanjingyingxingshouru","jihuashengyujin","shengtaibuchangjin","caichanxingshouru","zuidishenghuobaozhangjin","qitazhuanyixingshouru","zichanshouyifenhong","tekunrenyuanjiuzhugongyangjin","qitacaichanxing","shengchanxzhichu","jiatingchunshouru","jiatingrenkoushu","renjunchunshouru","niandujiatingrenkoushu","gengdimianji","mucaomianji","shuimianmianji","lindimianji","tuigenghuancaomianji","linguomianji","ruhudaolu","cunzhuganjuli","nongminzhuanyehezuoshe","weifang","zhufangmianji","ranliaoleix","longtouqiyedaidong","chuangyezhifudaitouren","tongshenghyongdian","tongguangbodianshi","chanyebangfuleixing","yinshuianquan" 
#endregion
                    DataTable dt = new DataTable();
                    DataTable dtadd = new DataTable();
                    DataTable dtBF = new DataTable();
                    DataTable dt_2 = new DataTable();
                    DataTable dt_3 = new DataTable();
                    #region /// 注意：如果在Word中没有类似表格或者列表的数据，此处则不需要，删除即可

                    if (!string.IsNullOrEmpty(huzhuID))
                    {
                        //string sql = $" SELECT `县`,`乡`,`村`,`户编号`,`人口编号`,`姓名`,`证件号码`,`与户主关系`,  \r\n\r\nCASE WHEN `与户主关系`='户主' then '1' WHEN `与户主关系`='配偶' then '2' ELSE NULL END as `户内序号`,\r\n\r\n  `性别`,`证件类型`,`出生日期`,`民族`,`家庭成员联系电话`,`文化程度`,`在校生状况`,`劳动技能`,`务工时间（月）`,`健康状况`,`政治面貌`,`是否参加城乡居民基本医疗保险`,`是否参加城乡居民基本养老保险`,`是否会讲普通话`,`义务教育阶段未上学原因`,`是否享受人身意外保险补贴`,`是否参加商业补充医疗保险`,`是否国外务工`,`是否接受医疗救助`,`是否接受其他健康扶贫`,`是否参加城镇职工基本医疗保险`,`大专或本科毕业生未就业原因`,`户主证件号码`,`户主姓名` FROM `全县全字段` WHERE `户编号`='{HuID}' ORDER BY CASE WHEN `户内序号`='1' then '1' WHEN `户内序号`='2' then '2' ELSE '3' END ASC";


                        string sql = $" SELECT * FROM `全县全字段` WHERE `户主身份证号`='{huzhuID}' ORDER BY CASE WHEN `户内序号`='1' then '1' WHEN `户内序号`='2' then '2' ELSE '3' END ASC";
                        dt = getDataTable(sql, connstr);
                        //这里的UserList很关键，要与WORD模板的域设置对应，不是随便写的，后面还会用到
                        dt.TableName = "PeopleList";
                    }
                    //if (!string.IsNullOrEmpty(HuID))
                    //{
                    //    string sql = $" SELECT `县`,`乡`,`村`,`户编号`,`人口编号`,`姓名`,`证件号码`,`与户主关系`,CASE WHEN `与户主关系`='户主' then '1' WHEN `与户主关系`='配偶' then '2' ELSE NULL END as `户内序号`,`务工所在地`,`务工企业名称`,`是否享受农村居民最低生活保障`,`是否参加城镇职工基本养老保险`,`是否参加大病保险`,`是否参加城乡居民基本医疗保险`,`是否参加城乡居民基本养老保险`,`义务教育阶段未上学原因`,`是否会讲普通话`,`务工时间（月）`,`是否享受人身意外保险补贴`,`是否参加商业补充医疗保险`,`是否国外务工`,`是否接受医疗救助`,`是否接受其他健康扶贫`,`是否参加城镇职工基本医疗保险`,`是否特困供养人员`,`公益性岗位`,`公益性岗位(月数)`,`残疾证办证年度`,`大专或本科毕业生未就业原因`,`是否事实无人抚养儿童` FROM `全县全字段` WHERE `户编号`='{HuID}' ORDER BY CASE WHEN `户内序号`='1' then '1' WHEN `户内序号`='2' then '2' ELSE '3' END ASC";
                    //    dtadd = getDataTable(sql, connstr);
                    //    //这里的UserList很关键，要与WORD模板的域设置对应，不是随便写的，后面还会用到
                    //    dtadd.TableName = "PeopleAddList";
                    //}
                    //if (!string.IsNullOrEmpty(HuID))
                    //{
                    //    string sql = $" SELECT `县`,`乡`,`村`,`户编号`,`人口编号`,`姓名`,`证件号码`,`与户主关系`,CASE WHEN `与户主关系`='户主' then '1' WHEN `与户主关系`='配偶' then '2' ELSE NULL END as `户内序号`, \r\n\r\nCASE WHEN `性别`='1' then '男' WHEN `性别`='2' then '女' ELSE NULL END as `性别`, `证件类型`,`出生日期` FROM `全县全字段` WHERE `户编号`='{HuID}' ORDER BY CASE WHEN `户内序号`='1' then '1' WHEN `户内序号`='2' then '2' ELSE '3' END ASC";
                    //    dt_2 = getDataTable(sql, connstr);
                    //    //这里的UserList很关键，要与WORD模板的域设置对应，不是随便写的，后面还会用到
                    //    dt_2.TableName = "People_2List";
                    //}
                    //if (!string.IsNullOrEmpty(HuID))
                    //{
                    //    string sql = $" SELECT `县`,`乡`,`村`,`户编号`,`人口编号`,`姓名`,`证件号码`,`与户主关系`,CASE WHEN `与户主关系`='户主' then '1' WHEN `与户主关系`='配偶' then '2' ELSE NULL END as `户内序号`,CASE WHEN `性别`='1' then '男' WHEN `性别`='2' then '女' ELSE NULL END as `性别`,`证件类型`,`出生日期` FROM `全县全字段` WHERE `户编号`='{HuID}' ORDER BY CASE WHEN `户内序号`='1' then '1' WHEN `户内序号`='2' then '2' ELSE '3' END ASC ";
                    //    dt_3 = getDataTable(sql, connstr);
                    //    //这里的UserList很关键，要与WORD模板的域设置对应，不是随便写的，后面还会用到
                    //    dt_3.TableName = "People_3List";
                    //}

                    //if (!string.IsNullOrEmpty(HuID))
                    //{
                    //    string sql = $" SELECT `市`,`县`,`乡`,`户编号`,`单位名称`,`隶属关系`,`姓名`,`性别`,`联系人电话`,`户主个人编号`,`户主姓名`,`户主证件号码`,`帮扶开始时间`,`帮扶结束时间`,`帮扶责任人编号`,`帮扶单位编号`,`政治面貌` FROM ((SELECT * FROM `帮扶责任人结对查询` WHERE LENGTH(`户主个人编号`)>1 ) as bf JOIN (SELECT `户编号`,`人口编号` FROM `全县全字段`) as qu  ON bf.`户主个人编号`=qu.`人口编号`)  WHERE `户编号`='{HuID}' ORDER BY `帮扶开始时间` ASC ";
                    //    dtBF = getDataTable(sql, connstr);
                    //    //这里的UserList很关键，要与WORD模板的域设置对应，不是随便写的，后面还会用到
                    //    dtBF.TableName = "BFList";
                    //}

                    #endregion
                    //string xiang = dtmain.Rows[0]["乡"].ToString();
                    //string cun = dtmain.Rows[0]["村"].ToString();
                    //string huzhuName = dtmain.Rows[0]["户主姓名"].ToString();
                    //string huzhuID = dtmain.Rows[0]["户主证件号码"].ToString();
                    Fun.Updatelogtext($"获取{xiang}-{cun}-{huzhuName}信息完成。");
                    //string StrVisitURL = WordModelEvent.ExpertWordToModel("模版", $"{xiang}-{cun}-{huzhuName}-{huzhuID}脱贫户和监测户信息采集表", dtmain, dt, dtadd, dtBF, dt_2, dt_3);
                    string StrVisitURL = WordModelEvent.ExpertWordToModel("模版", $"{xiang}-{cun}-{huzhuName}-{huzhuID}民情信息采集表", dtmain, dt);
                    //dt.Dispose();
                    //MessageBox.Show($"{member.xiang}-{member.cun}-{member.huzhuName}-{member.huzhuID}生成完成！");

                    Fun.UpdateState(connstr, "全县全字段", huzhuID);

                    return $"{xiang}-{cun}-{huzhuName}-{huzhuID}生成完成！";

                }
                catch (Exception er)
                {
                    //MessageBox.Show(er.Message);
                    return $"生成失败！失败原因{er.Message}";
                }
            }
            /// <summary>
            /// 通过SQL获得 DataTable
            /// </summary>
            /// <param name="sql">sql查询语句</param>
            /// <param name="ConnectionString">你的数据链接字符串</param>
            /// <returns></returns>
            private static DataTable getDataTable(string sql, string ConnectionString = "")
            {
                var dt = new DataTable();
                using (var connection = new SQLiteConnection(ConnectionString))
                {
                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }

                    using (var command = new SQLiteCommand(sql, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            dt.Load(reader);
                        }
                    }

                }
                return dt;



            }
        

        }
    }
}
