using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FSTIMAIN.DAL                                                   //��Ŀ����.DAL
{
    class SqlHelper
    {
        private static readonly string connstr = "Data Source=192.168.8.67;database=UltimusBusiness;uid=ERP;pwd=bpm+123";
            //ConfigurationManager.ConnectionStrings["connstr"].ConnectionString;    //�����ļ�App.config���������ݿ��ַ���
        public static int ExecuteNonQuery(string cmdText, params SqlParameter[] para)
        {
            SqlConnection conn = new SqlConnection(connstr);
            conn.Open();
            SqlCommand cmd = new SqlCommand(cmdText, conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteNonQuery();
        }
        public static object ExecuteScalar(string cmdText, params SqlParameter[] para)
        {
            SqlConnection conn = new SqlConnection(connstr);
            conn.Open();
            SqlCommand cmd = new SqlCommand(cmdText, conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteScalar();
        }
        public static DataTable ExecuteDataTable(string cmdText, params SqlParameter[] para)
        {
            SqlConnection conn = new SqlConnection(connstr);
            SqlCommand cmd = new SqlCommand(cmdText, conn);
            cmd.Parameters.AddRange(para);
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            return dt;
        }
    }
}
//�������  System.Configuration
//App.config �������Ĵ���
 // <connectionStrings>
  //  <add name="connstr" connectionString="Data Source=DESKTOP-4TK3U1D;Initial Catalog=sanceng;Integrated Security=True"/>
  // "Data Source=(local);database=BOOKS;uid=sa;pwd=123456"
  //"Data Source=192.168.36.90;database=BOOKS;uid=sa;pwd=123456"
  //  </connectionStrings>    