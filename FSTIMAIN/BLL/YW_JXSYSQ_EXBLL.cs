using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FSTIMAIN.DAL;
using FSTIMAIN.Model;

namespace FSTIMAIN.BLL
{
class YW_JXSYSQ_EXBLL
{
public int AddNew(YW_JXSYSQ_EX model)
{
return new YW_JXSYSQ_EXDAL().AddNew(model);
}
public int Delete(int id)
{
return new YW_JXSYSQ_EXDAL().Delete(id);
}
public int Update(YW_JXSYSQ_EX model)
{
return new YW_JXSYSQ_EXDAL().Update(model);
}
public YW_JXSYSQ_EX Get(string  id)
{
return new YW_JXSYSQ_EXDAL().Get(id);
}
public IEnumerable<YW_JXSYSQ_EX> GetAll(string ParentGuid)
{
return new YW_JXSYSQ_EXDAL().GetAll(ParentGuid);
}
}
}
