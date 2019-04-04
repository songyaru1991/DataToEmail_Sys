using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BLL;
using System.Net.Mail;
using System.IO;
namespace Main
{
    public partial class Frm_Main : Form
    {
        public Frm_Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void Frm_Main_Load(object sender, EventArgs e)
        {
            try
            {
                DirectoryInfo folder = new DirectoryInfo(@"D:\DataToEmailSuper");
                foreach (FileInfo file in folder.GetFiles("*.xls"))
                {
                    if (File.Exists(file.FullName))
                    {
                        File.Delete(file.FullName);
                    }
                }            //把里面所有文件都删了
                string weekstr = DateTime.Now.DayOfWeek.ToString();
                switch (weekstr)
                {
                    case "Monday":                        //星期一，抓上周五，也就是前三天的数据
                        new UserService().CreateSwipeCardExcel(-3);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Tuesday":  //星期二，抓周六和周日，前三天和前两天
                        new UserService().CreateSwipeCardExcel(-3);
                        new UserService().CreateSwipeCardExcel(-2);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Wednesday":
                        new UserService().CreateSwipeCardExcel(-2);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Thursday":
                        new UserService().CreateSwipeCardExcel(-2);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Friday":
                        new UserService().CreateSwipeCardExcel(-2);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Saturday":
                        new UserService().CreateSwipeCardExcel(-2);
                        new UserService().CreateOnWorkEmpExcel();
                        break;
                    case "Sunday":
                        System.Environment.Exit(0);//星期日不用运行程序  运行直接退出
                        break;
                }
                List<string> list_str = new UserService().ReadMail();
                new UserService().SendMailUse(list_str[0], list_str);
            }
            catch (Exception ex)
            {
                string userName = "Cheng_Qian@KUNSHAN";
                SmtpClient client = new SmtpClient();
                client.DeliveryMethod = SmtpDeliveryMethod.Network;//指定电子邮件发送方式    
                client.Host = "192.168.78.201";//邮件服务器
                client.UseDefaultCredentials = true;
                client.Credentials = new System.Net.NetworkCredential(userName, "");//用户名、密码
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                msg.From = new MailAddress(userName, "Cheng_Qian");//后面一参数是发件人名字
                msg.To.Add("Cheng_Qian@KUNSHAN");//收件人地址              
                msg.CC.Add("Cheng_Qian@KUNSHAN");//抄送
                msg.Subject = "實時工時系統數據統計";//邮件标题  
                msg.Body = "<style type=" + "text/css" + "> h3{ color:red; font-family:" + " 微软雅黑" + "}</style><h3>&nbsp;&nbsp;&nbsp;&nbsp;" + ex.Message + "</h3>";
                msg.BodyEncoding = System.Text.Encoding.UTF8;//邮件内容编码   
                msg.IsBodyHtml = true;//是否是HTML邮件   
                msg.Priority = MailPriority.High;//邮件优先级    
                client.Send(msg);
            }
            finally
            {
                System.Environment.Exit(0);    //程序退出
            }
        }  //报加班人数对比数据
    }
}
