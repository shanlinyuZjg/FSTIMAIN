using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FSTIMAIN.Model
{
    class ITMBliucheng
    {
        public Int32 流水号 { get; set; }
        public string 申请方式 { get; set; }
        public string 摘要 { get; set; }
        public string 接收单位 { get; set; }
        public string 制购类型 { get; set; }
       
        public string 申请人 { get; set; }
        public DateTime 申请时间 { get; set; }
        public string 发起部门 { get; set; }
        public string 联系电话 { get; set; }
        public string ParentGuid { get; set; }
    }
}
