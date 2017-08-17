#region Using directives

using System.Collections.Generic;
using System.IO;
using Excel;
using System.Text;
using System;
using System.Windows.Forms;

#endregion

namespace ExcelApplication1
{
    public class Worker
    {

        public static bool merge(string selectedPath, List<ExcelFile> excelFileList,Settings settings ,System.ComponentModel.BackgroundWorker backgroundWorker)
        {
            // Start the search for primes and wait.
            UTF8Encoding utf8 = new UTF8Encoding(false);
            var writer = new StreamWriter(settings.TargetPath + "result.csv", false, utf8);
            writer.WriteLine("虚拟卡编号,虚拟卡密码,开始有效期,截止有效期,活动名称,奖品名称,是否使用,活动编号");
            try
            {
                int i = 0;
                foreach (var f in excelFileList)
                {
                    i++;
                    if (backgroundWorker.CancellationPending)
                    {
                        // Return without doing any more work.
                        throw new Exception("用户取消了操作");
                    }

                    if (backgroundWorker.WorkerReportsProgress)
                    {
                        backgroundWorker.ReportProgress(i);
                    }

                    if (!f.IsSelected) continue;
                    f.Status = "处理中";
                    var stream = new FileStream(f.Path, FileMode.Open);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    try
                    {
                        excelReader.Read();
                        while (excelReader.Read())
                        {
                            //虚拟卡编号,虚拟卡密码,开始有效期,截止有效期,活动名称,奖品名称,是否使用,活动编号
                            var sec = (excelReader.GetString(0).Trim());// 序号
                            var code = (excelReader.GetString(1).Trim());
                            var pwd = (excelReader.GetString(2).Trim());
                            var desc = (excelReader.GetString(3).Trim()); // 优惠券方案
                            var start_time = excelReader.GetString(4).Trim();
                            DateTime dt;
                            if (DateTime.TryParse(start_time, out dt))
                            {
                                start_time = dt.ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            else
                            {
                                throw new ArgumentException("开始有效期日期格式不正确"); // throw
                            }

                            var end_time = excelReader.GetString(5).Trim();

                            if (DateTime.TryParse(end_time, out dt))
                            {
                                end_time = dt.AddDays(1).AddSeconds(-1).ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            else
                            {
                                throw new ArgumentException("截止有效期日期格式不正确"); // throw
                            }

                            var activity_id = settings.ActivityId;
                            var prize_id = settings.PrizeId;
                            var is_used = settings.IsUsed ? 1 : 0;
                            var prjcode = settings.Prjcode;

                            if (!string.IsNullOrEmpty(code))
                            {
                                writer.WriteLine("{0},{1},{2},{3},{4},{5},{6},{7}", code, pwd, start_time, end_time, activity_id, prize_id, is_used, prjcode);
                            }
                        }

                        f.Status = "处理完成";
                    }
                    catch (Exception e1)
                    {
                        f.Status = "处理失败";
                        throw e1;
                    }
                    finally
                    {
                        excelReader.Close();
                    }
                }
                
                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                throw e;
                return false;
            }
            finally
            {
                writer.Flush();
                writer.Close();
            }
        }

        
    }
}
