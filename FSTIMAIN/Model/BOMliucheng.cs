using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FSTIMAIN.Model
{
    class BOMliucheng
    {
        public Int32 流水号 { get; set; }
        public string 申请方式 { get; set; }
        public string 摘要 { get; set; }
        public string 父物料号 { get; set; }
        public string 物料描述 { get; set; }
        public string 申请人 { get; set; }
        public DateTime 申请时间 { get; set; }
        public string 发起部门 { get; set; }
        public string 申请类型 { get; set; }
        public string ParentGuid { get; set; }
    }
}
