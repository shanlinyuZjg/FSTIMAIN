using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using FSTIMAIN.Model;

namespace FSTIMAIN.DAL
{
class YW_JXSYSQ_EXDAL
{
public int AddNew(YW_JXSYSQ_EX model)
{
object obj= SqlHelper.ExecuteScalar("insert into YW_JXSYSQ_EX(ParentGuid,ZX,MS,DW,YY,SL,SH,ZL,LL,SHX,SX,STATUS,CREATE_DATE,XH) values(@ParentGuid,@ZX,@MS,@DW,@YY,@SL,@SH,@ZL,@LL,@SHX,@SX,@STATUS,@CREATE_DATE,@XH); select @@identity"
,new SqlParameter("ParentGuid",model.ParentGuid??(object)DBNull.Value)
,new SqlParameter("ZX",model.ZX??(object)DBNull.Value)
,new SqlParameter("MS",model.MS??(object)DBNull.Value)
,new SqlParameter("DW",model.DW??(object)DBNull.Value)
,new SqlParameter("YY",model.YY??(object)DBNull.Value)
,new SqlParameter("SL",model.SL??(object)DBNull.Value)
,new SqlParameter("SH",model.SH??(object)DBNull.Value)
,new SqlParameter("ZL",model.ZL??(object)DBNull.Value)
,new SqlParameter("LL",model.LL??(object)DBNull.Value)
,new SqlParameter("SHX",model.SHX??(object)DBNull.Value)
,new SqlParameter("SX",model.SX??(object)DBNull.Value)
,new SqlParameter("STATUS",model.STATUS??(object)DBNull.Value)
,new SqlParameter("CREATE_DATE",model.CREATE_DATE??(object)DBNull.Value)
,new SqlParameter("XH",model.XH??(object)DBNull.Value)
);
return Convert.ToInt32(obj);
}
public int Delete(int id)
{
return SqlHelper.ExecuteNonQuery("delete from YW_JXSYSQ_EX where Guid = @id", new SqlParameter("id", id));
}
public int Update(YW_JXSYSQ_EX model)
{
return SqlHelper.ExecuteNonQuery("update YW_JXSYSQ_EX set ParentGuid=@ParentGuid,ZX=@ZX,MS=@MS,DW=@DW,YY=@YY,SL=@SL,SH=@SH,ZL=@ZL,LL=@LL,SHX=@SHX,SX=@SX,STATUS=@STATUS,CREATE_DATE=@CREATE_DATE,XH=@XH where Guid = @id"
,new SqlParameter("id", model.Guid)
,new SqlParameter("ParentGuid",model.ParentGuid??(object)DBNull.Value)
,new SqlParameter("ZX",model.ZX??(object)DBNull.Value)
,new SqlParameter("MS",model.MS??(object)DBNull.Value)
,new SqlParameter("DW",model.DW??(object)DBNull.Value)
,new SqlParameter("YY",model.YY??(object)DBNull.Value)
,new SqlParameter("SL",model.SL??(object)DBNull.Value)
,new SqlParameter("SH",model.SH??(object)DBNull.Value)
,new SqlParameter("ZL",model.ZL??(object)DBNull.Value)
,new SqlParameter("LL",model.LL??(object)DBNull.Value)
,new SqlParameter("SHX",model.SHX??(object)DBNull.Value)
,new SqlParameter("SX",model.SX??(object)DBNull.Value)
,new SqlParameter("STATUS",model.STATUS??(object)DBNull.Value)
,new SqlParameter("CREATE_DATE",model.CREATE_DATE??(object)DBNull.Value)
,new SqlParameter("XH",model.XH??(object)DBNull.Value)
);
}
public YW_JXSYSQ_EX Get(string id)
{
DataTable dt = SqlHelper.ExecuteDataTable("select * from YW_JXSYSQ_EX where Guid = @id", new SqlParameter("id", id));
if (dt.Rows.Count != 1)
{
return null;
}
else
{
return tomodel(dt.Rows[0]);
}
}
public IEnumerable<YW_JXSYSQ_EX> GetAll(string conStr)
{
DataTable dt = SqlHelper.ExecuteDataTable(conStr);
List<YW_JXSYSQ_EX> list = new List<YW_JXSYSQ_EX>();
foreach (DataRow dr in dt.Rows)
{
list.Add(tomodel(dr));
}
return list;
}
public YW_JXSYSQ_EX tomodel(DataRow dr)
{
YW_JXSYSQ_EX model = new YW_JXSYSQ_EX();
 model.Guid = (string)dr["Guid"];
model.ParentGuid = dr["ParentGuid"]==DBNull.Value?null:(string)dr["ParentGuid"];
model.ZX = dr["ZX"]==DBNull.Value?null:(string)dr["ZX"];
model.MS = dr["MS"]==DBNull.Value?null:(string)dr["MS"];
model.DW = dr["DW"]==DBNull.Value?null:(string)dr["DW"];
model.YY = dr["YY"]==DBNull.Value?null:(string)dr["YY"];
model.SL = dr["SL"]==DBNull.Value?null:(string)dr["SL"];
model.SH = dr["SH"]==DBNull.Value?null:(string)dr["SH"];
model.ZL = dr["ZL"]==DBNull.Value?null:(string)dr["ZL"];
model.LL = dr["LL"]==DBNull.Value?null:(string)dr["LL"];
model.SHX = dr["SHX"]==DBNull.Value?null:(string)dr["SHX"];
model.SX = dr["SX"]==DBNull.Value?null:(string)dr["SX"];
model.STATUS = dr["STATUS"]==DBNull.Value?null:(string)dr["STATUS"];
model.CREATE_DATE = dr["CREATE_DATE"]==DBNull.Value?null:(string)dr["CREATE_DATE"];
model.XH = dr["XH"]==DBNull.Value?null:(string)dr["XH"];
return model;
}
}
}
