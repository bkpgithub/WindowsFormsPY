using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using WindowsFormsPY.Properties;

namespace WindowsFormsPY
{
    public partial class Form1 : Form
    {
        public Dictionary<string, List<string>> loopholeInfoDic = new Dictionary<string, List<string>>();
        public int id = 0;
        public Form1()
        {
            InitializeComponent();
            
        }
        public static bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }
        public void CrawlData(string filePath)
        {
            int portId = 0;
            string ip = "";
            try
            {
                WebClient c = new WebClient();
                c.Encoding = Encoding.GetEncoding("UTF-8");
                string html = c.DownloadString(filePath);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);
                #region ip
                HtmlNode nodeinfoIp = doc.GetElementbyId("content");
                if(nodeinfoIp == null)
                {
                    return;
                }
                foreach(HtmlNode row in nodeinfoIp.SelectNodes(".//tr[contains(@class,'even')]"))
                {
                    if (row.SelectNodes("th|td")[0].InnerText.Equals("IP地址"))
                    {
                        ip = row.SelectNodes("th|td")[1].InnerText;
                        break;
                    }                    
                }
                #endregion
                #region 表1
                HtmlNode nodeinfo = doc.GetElementbyId("vuln_list");
                if (nodeinfo == null)
                {
                    return;
                }
                foreach (HtmlNode row in nodeinfo.SelectNodes(".//tr"))
                {
                    string port = row.SelectNodes("th|td")[0].InnerText;
                    if (IsNumeric(port) == true)
                    {
                        portId = Convert.ToInt32(port);
                    }
                    try
                    {
                        foreach (HtmlNode li in row.SelectNodes(".//li"))
                        {
                            HtmlNode tdNode_danger_high = li.SelectSingleNode(".//span[contains(@class,'level_danger_high')]");
                            AddListInfo(ip, portId, tdNode_danger_high,1);
                            HtmlNode tdNode_danger_middle = li.SelectSingleNode(".//span[contains(@class,'level_danger_middle')]");
                            AddListInfo(ip, portId, tdNode_danger_middle,2);
                        }
                    }
                    catch (Exception)
                    {
                    }
                    
                }
                #endregion
                bool startAdd = false;
                string name = "";
                #region 表2
                HtmlNode nodeinfo2 = doc.GetElementbyId("vul_detail");
                foreach (HtmlNode row in nodeinfo2.SelectNodes(".//tr"))
                {
                    string txt = "";
                    if(startAdd == true)
                    {
                        txt = row.InnerText;
                        txt = txt.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "");
                        string solution = "解决办法";
                        string threatScore = "威胁分值";
                        int indexSolution = txt.IndexOf(solution);
                        int indexThreatScore = txt.IndexOf(threatScore);
                        foreach (var item in loopholeInfoDic)
                        {
                            if (item.Value[1].Equals(name) && item.Value.Count == 4)
                            {
                                item.Value.Add(txt.Substring(4, indexSolution - 4));
                                item.Value.Add(txt.Substring(indexSolution + 4, indexThreatScore - indexSolution - 4));
                                break;
                            }
                        }
                        //loopholeInfoDic[id.ToString()].Add(txt.Substring(4,indexSolution - 4));
                        //loopholeInfoDic[id.ToString()].Add(txt.Substring(indexSolution + 4 , indexThreatScore - indexSolution - 4));
                        startAdd = false;
                    }
                    HtmlNode tdNode_danger_high = row.SelectSingleNode(".//span[contains(@class,'level_danger_high')]");
                    if (tdNode_danger_high != null)
                    {
                        foreach (var item in loopholeInfoDic)
                        {
                            if (item.Value[1].Equals(tdNode_danger_high.InnerText)&& item.Value.Count == 4)
                            {
                                startAdd = true;
                                name = tdNode_danger_high.InnerText;
                                break;
                            }
                        }
                    }
                    HtmlNode tdNode_danger_middle = row.SelectSingleNode(".//span[contains(@class,'level_danger_middle')]");
                    if(tdNode_danger_middle != null)
                    {
                        foreach (var item in loopholeInfoDic)
                        {
                            if (item.Value[1].Equals(tdNode_danger_middle.InnerText) && item.Value.Count == 4)
                            {
                                startAdd = true;
                                name = tdNode_danger_middle.InnerText;
                                break;
                            }
                        }
                    }
                }
                #endregion
            }
            catch (WebException webEx)
            {
                Console.WriteLine(webEx.Message.ToString());
            }
        }
        private void AddListInfo(string ip, int portId, HtmlNode tdNode_danger_high,int lel)
        {
            if (tdNode_danger_high != null)
            {
                string loopholeName = tdNode_danger_high.InnerText;
                //if (loopholeInfoDic.Count == 0)
                //{
                loopholeInfoDic.Add(id.ToString(), new List<string>());
                loopholeInfoDic[id.ToString()].Add(ip);
                loopholeInfoDic[id.ToString()].Add(loopholeName);
                loopholeInfoDic[id.ToString()].Add(portId.ToString());
                if(lel == 1)
                {
                    loopholeInfoDic[id.ToString()].Add("高");
                }
                else if(lel == 2)
                {
                    loopholeInfoDic[id.ToString()].Add("中");
                }
                id++;
            }
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if(textBoxPath.Text !="")
            {
                this.label1.Visible = true;
                this.pictureBox1.Visible = true;
                this.label1.Text = "正在整理中......";
                this.pictureBox1.Image = Resources._123;
                List<String> list = Director(textBoxPath.Text);
                foreach(var path in list)
                {
                    CrawlData(path);
                }
            }
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                //新增加一个工作簿，Workbook是直接保存，不会弹出保存对话框，加上Application会弹出保存对话框，值为false会报错    
                Microsoft.Office.Interop.Excel.Workbook xBook = excel.Workbooks.Add(Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];

                xSheet.Cells[1, 1] = "ip地址";
                xSheet.Cells[1, 2] = "漏洞名称";
                xSheet.Cells[1, 3] = "详细描述";
                xSheet.Cells[1, 4] = "加固建议";
                xSheet.Cells[1, 5] = "风险端口";
                xSheet.Cells[1, 6] = "危险程度";
                List<string> loopholeInfoDicKey = new List<string>(loopholeInfoDic.Keys);
                for(int i = 0; i < loopholeInfoDicKey.Count; i++)
                {
                    for(int j = 0; j < 6; j++)
                    {
                        if(j == 2)
                        {
                            if(loopholeInfoDic[loopholeInfoDicKey[i]].Count > 4)
                            {
                                xSheet.Cells[i + 2, j + 1] = loopholeInfoDic[loopholeInfoDicKey[i]][4];
                            }
                            
                        }
                        else if(j==3)
                        {
                            if (loopholeInfoDic[loopholeInfoDicKey[i]].Count > 4)
                            {
                                xSheet.Cells[i + 2, j + 1] = loopholeInfoDic[loopholeInfoDicKey[i]][5];
                            }                                
                        }
                        else if (j == 4)
                        {
                            xSheet.Cells[i + 2, j + 1] = loopholeInfoDic[loopholeInfoDicKey[i]][2];
                        }
                        else if (j == 5)
                        {
                            xSheet.Cells[i + 2, j + 1] = loopholeInfoDic[loopholeInfoDicKey[i]][3];
                        }
                        else
                        {
                            xSheet.Cells[i + 2, j + 1] = loopholeInfoDic[loopholeInfoDicKey[i]][j];
                        }
                    }
                }
                excel.DisplayAlerts = false;
                string saveName = GetSystemSecond().ToString();
                string savePath = Application.StartupPath + "\\" + saveName + ".xlsx";
                xBook.SaveAs(savePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);       
                xSheet = null;
                xBook = null;
                excel.Quit();
                excel = null;
                GC.Collect();//如果不使用这条语句会导致excel进程无法正常退出，使用后正常退出  
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存excel出错: "+ex.Message);
                Close();
            }
            this.label1.Text = "整理完成！";
            this.pictureBox1.Image = Resources._234;
        }
        

        public List<String> Director(string dirs)
        {
            List<String> list = new List<String>();
            //绑定到指定的文件夹目录
            DirectoryInfo dir = new DirectoryInfo(dirs);
            //检索表示当前目录的文件和子目录
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            //遍历检索的文件和子目录
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                //判断是否为空文件夹　　
                if (fsinfo is DirectoryInfo)
                {
                    //递归调用
                    Director(fsinfo.FullName);
                }
                else
                {
                    Console.WriteLine(fsinfo.FullName);
                    //将得到的文件全路径放入到集合中
                    list.Add(fsinfo.FullName);
                }
            }
            return list;
        }
        public static int GetSystemSecond()
        {
            return (DateTime.Now.Hour * 3600 + DateTime.Now.Minute * 60 + DateTime.Now.Second);
        }
    }
}
