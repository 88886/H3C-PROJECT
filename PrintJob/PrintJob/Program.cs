//************************************************
//2010-9-14 by David.Xu
//Version 1.0
//Description:������ǩ���ߴ�ӡ��
//�����׼��Ч��/������Ч�� ������ʾ
//************************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Collections;


namespace PrintJob
{
    class Program
    {
        static void Main(string[] args)
        {

            string cn = "Data Source=10.11.1.27;Initial Catalog=Print;Persist Security Info=True;User ID=sa;Password=sa;Pooling=true";
            SqlConnection conn = new SqlConnection(cn);

            conn.Open();


            string str = "SELECT * FROM tblHuaWei where datediff(day,convert(datetime,MSValidFrom),dateadd(day,5,getdate()))>=0 or datediff(day,convert(datetime,ValidFrom),dateadd(day,5,getdate()))>=0 ";


            SqlDataAdapter sda = new SqlDataAdapter(str, conn);


            DataSet dsHW = new DataSet();

            sda.Fill(dsHW);

            string sql = "SELECT * FROM tblh3c where datediff(day,convert(datetime,MSValidFrom),dateadd(day,5,getdate()))>=0 or datediff(day,convert(datetime,ValidFrom),dateadd(day,5,getdate()))>=0 ";

            SqlDataAdapter da = new SqlDataAdapter(sql, conn);

            DataSet dsH3C = new DataSet();

            da.Fill(dsH3C);


            string strmail = "select * from tblmailList ";
            SqlDataAdapter daMail = new SqlDataAdapter(strmail, conn);
            DataSet dsMail = new DataSet();
            daMail.Fill(dsMail);

            if (conn.State == ConnectionState.Open)
            {
                conn.Dispose();
                conn.Close();
            }


            MailAddress from = new MailAddress("WebMaster@cn1.flashelec.com", "WebMaster@cn1.flashelec.com");
            MailMessage mail = new MailMessage();
            mail.Subject = "�����׼�ͽ�����ɵ�����Ѷ!";
            mail.From = from;

            for (int i = 0; i < dsMail.Tables[0].Rows.Count; i++)
            {
                mail.To.Add(dsMail.Tables[0].Rows[i]["Mail"].ToString());
            }

            mail.Bcc.Add("David.Xu@cn.asteelflash.com");
            
            System.Text.StringBuilder sbmail = new System.Text.StringBuilder();
            sbmail.Append("<p style='color:#FF0000;'>���±�׼��Ч���ѵ���5���ڵ��ڣ���ϸ���£�</p>");
            sbmail.Append("<table cellpadding='0' cellspacing='0' border='1'><tr align='center' bgcolor='#FFCC33'>");
            sbmail.Append("<td width='150' height='30'>��Ʒ����</td><td width='100'>Ӳ���汾</td><td width='200'>��Ʒ����</td><td width='150'>�����׼</td>");
            sbmail.Append("<td width='150'>�����׼��Ч��</td><td width='150'>������ɺ�</td><td width='150'>���������Ч��</td></tr>");
            for (int m = 0; m < dsHW.Tables[0].Rows.Count; m++)
            {
                sbmail.Append("<tr align='left'><td height='23'>" + (string)dsHW.Tables[0].Rows[m]["SN"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["HV"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["Des"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["MS"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["MSValidFrom"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["NAL"] + "</td><td>" + (string)dsHW.Tables[0].Rows[m]["ValidFrom"] + "</td></tr>");
            }
            for (int n = 0; n < dsH3C.Tables[0].Rows.Count; n++)
            {
                sbmail.Append("<tr align='left'><td height='23'>" + (string)dsH3C.Tables[0].Rows[n]["SN"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["HV"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["Des"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["MS"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["MSValidFrom"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["NAL"] + "</td><td>" + (string)dsH3C.Tables[0].Rows[n]["ValidFrom"] + "</td></tr>");
            }
                
            sbmail.Append("</table>");


            mail.Body = sbmail.ToString();
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;

            SmtpClient client = new SmtpClient();
            client.Host = "sz-sql01.cn1.flashelec.com";
            client.Port = 25;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("adminbackup", "mib*fla12");
            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            client.Send(mail);

        }

    }
}
