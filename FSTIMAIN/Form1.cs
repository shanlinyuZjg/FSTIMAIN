using FSTIMAIN.DAL;
using SoftBrands.FourthShift.Transaction;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FSTIMAIN.Model;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using FSTIMAIN.Properties;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Threading;
using NPOI.SS.UserModel;

namespace FSTIMAIN
{
    
    public partial class Form1 : Form
    {
        public Form1()//初始化组件
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;//多线程访问控件 不安全方式
        }
        private FSTIClient _fstiClient = null;//声明FSTIClient类的对象
        private void Form1_Load(object sender, EventArgs e)//初始化登录四班账号
        {
            tabControl1.SelectedIndex = 5;//默认打开登录界面
            ItemClass.SelectedIndex = 5;
            ItemClasssh.SelectedIndex = 5;
            this.Show();
            btnInitialize_Click(null, null);
            this.textUserId.Focus();

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)//关闭程序时，退出登录
        {
            if (_fstiClient != null)
            {
                _fstiClient.Terminate();
                _fstiClient = null;
            }
        }
        private void btnInitialize_Click(object sender, EventArgs e)//登录界面配置文件的初始化按钮
        {
            try
            {
                if (_fstiClient != null)
                {
                    _fstiClient.Terminate();
                    _fstiClient = null;
                }
                toolStripStatusLabel1.Text = "未登录";
                toolStripStatusLabel2.Text = "";
                toolStripStatusLabel3.Text = "";
                _fstiClient = new FSTIClient();

                // call InitializeByConfigFile
                // second parameter == true is to participate in unified logon
                // third parameter == false, no support for impersonation is needed

                _fstiClient.InitializeByConfigFile(textConfig.Text, true, false);
                //MessageBox.Show(_fstiClient.UserId);
                // Since this program is participating in unified logon, need to
                // check if a logon is required.

                if (_fstiClient.IsLogonRequired)
                {
                    // Logon is required, enable the logon button
                    btnLogon.Enabled = true;
                    btnLogon.Focus();
                }
                else
                {
                    toolStripStatusLabel1.Text = "ID:" + _fstiClient.UserId;
                    toolStripStatusLabel2.Text = "配置文件：" + textConfig.Text.Trim();
                    toolStripStatusLabel3.Text = "服务器:" + _fstiClient.ServerName;
                }
                // Disable the Initialize button
                btnInitialize.Enabled = false;
            }

            catch (FSTIApplicationException exception)
            {
                MessageBox.Show(exception.Message, "FSTIApplication Exception");
                _fstiClient.Terminate();
                _fstiClient = null;
            }
        }
        private void btnLogon_Click(object sender, EventArgs e)//登录界面配置文件的登录按钮
        {
            if (btnLogon.Enabled == false)
            { MessageBox.Show("请先初始化！"); return; }
            string message = null;     // used to hold a return message, from the logon
            int status;         // receives the return value from the logon call
            textUserId.Text = textUserId.Text.Trim();
            textPassword.Text = textPassword.Text.Trim();
            status = _fstiClient.Logon(textUserId.Text, textPassword.Text, ref message);
            if (status > 0)
            {
                MessageBox.Show("Invalid user id or password");
            }
            else
            {
                btnLogon.Enabled = false;
                toolStripStatusLabel1.Text = "ID:" + _fstiClient.UserId;
                toolStripStatusLabel2.Text = "配置文件：" + textConfig.Text.Trim();
                toolStripStatusLabel3.Text = "服务器:" + _fstiClient.ServerName;
            }
        }
        private void tabControl1_Click(object sender, EventArgs e)//切换tabControl1界面的事件
        {
            if (tabControl1.SelectedIndex == 5)
            {
                btnInitialize.Enabled = true;
                btnLogon.Enabled = false;
                //textUserId.Text = "";
                //textPassword.Text = "";
            }
        }

        private void GetBOm_Click(object sender, EventArgs e)//获得BOM流程
        {
            //测试获得BOM流程
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and PROCESSNAME='RY增加BOM申请流程' and TASKUSER='BPM/zuojinguo' and ENDTIME >'2020/9/17' and STEPLABEL='ERP系统录入'");
            //实际获得BOM流程 and TASKUSER='BPM/zuojinguo'
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY增加BOM申请流程' and TASKUSER='BPM/zuojinguo'  and STEPLABEL='ERP系统录入'");
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY增加BOM申请流程'   and STEPLABEL='ERP系统录入' ");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY增加BOM申请流程'");
            List<BOMliucheng> list1 = new List<BOMliucheng>();
            foreach (DataRow dr in Incidents.Rows)
            {
                BOMliucheng bom1 = Tolist(SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_RYZY_ZJBOM] where REV_INCIDENT=" + dr[0]));
                list1.Add(bom1);

            }

            BOM.DataSource = list1;
            BOM.Columns["摘要"].Width = 260;
            BOM.Columns["申请时间"].Width = 110;
            BOMResult.Items.Clear();
            dgvBOMDetail.DataSource = null;
        }
        private BOMliucheng Tolist(DataTable dt)//dt转化为BOMliucheng类对象的方法
        {

            BOMliucheng bom1 = new BOMliucheng();
            bom1.ParentGuid = (string)(dt.Rows[0]["REV_CID"]);
            bom1.申请人 = (string)(dt.Rows[0]["REV_CREATER_NAME"]);
            bom1.发起部门 = (string)(dt.Rows[0]["REV_CREATER_DPT"]);
            bom1.申请时间 = (DateTime)(dt.Rows[0]["REV_CREATER_DATE"]);

            bom1.流水号 = (Int32)(dt.Rows[0]["REV_INCIDENT"]);
            bom1.申请方式 = (string)(dt.Rows[0]["SQFS"]);
            bom1.申请类型 = (string)(dt.Rows[0]["SQLX"]);
            bom1.摘要 = (string)(dt.Rows[0]["ZY"]);
            bom1.父物料号 = (string)(dt.Rows[0]["FWLH"]);
            bom1.物料描述 = (string)(dt.Rows[0]["WLMS"]);
            if (dt.Rows.Count > 1)
                MessageBox.Show(bom1.流水号 + "有多个相同的流水号");

            return bom1;
        }


        private void BOM_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//BOM流程列表的cell双击事件
        {
            BOMResult.Items.Clear();
            int rowindex = e.RowIndex;

            if (rowindex != -1)
            {
                if (BOM.Rows[rowindex].DefaultCellStyle.ForeColor != Color.Red)
                {
                    BOM.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                    for (int a = 0; a < BOM.Rows.Count; a++)
                    {
                        if (a != rowindex && BOM.Rows[a].DefaultCellStyle.ForeColor != Color.Red)
                            BOM.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
                ParentItem.Text = BOM.Rows[rowindex].Cells["父物料号"].Value.ToString();
                BOMliushuihao.Text = BOM.Rows[rowindex].Cells["流水号"].Value.ToString();
                BOMhanghao.Text = (rowindex + 1).ToString();
                string ParentGuid = BOM.Rows[rowindex].Cells["ParentGuid"].Value.ToString();

                dgvBOMDetail.DataSource = SqlHelper.ExecuteDataTable("SELECT ZT as 状态, upper(ltrim(rtrim(YY))) as 用于,upper(ltrim(rtrim(XH))) as 序号,upper(ltrim(rtrim(ZX))) as 子项,upper(ltrim(rtrim(ZL))) as 子类,upper(ltrim(rtrim(SL))) as 数量,upper(ltrim(rtrim(LL))) as 量类,upper(ltrim(rtrim(DW))) as 单位,upper(ltrim(rtrim(SHX))) as 生效,upper(ltrim(rtrim(SX))) as 失效,upper(ltrim(rtrim(SH))) as 损耗 FROM YW_JXSYSQ_EX where ParentGuid = '" + ParentGuid + "'");
                if (BOM.Rows[rowindex].Cells["申请方式"].Value.ToString() == "增加")
                { AddBOM.Enabled = true; UpdateBOM.Enabled = false; }
                if (BOM.Rows[rowindex].Cells["申请方式"].Value.ToString() == "修改")
                { UpdateBOM.Enabled = true; AddBOM.Enabled = false; }
            }
        }
        private void DumpErrorObject(ITransaction transaction, FSTIError fstiErrorObject, ListBox listResult)//通用FSTIClient错误信息方法
        {
            listResult.Items.Add("Transaction Error:");
            listResult.Items.Add("");
            listResult.Items.Add(String.Format("Transaction: {0}", transaction.Name));
            listResult.Items.Add(String.Format("Description: {0}", fstiErrorObject.Description));
            tbISO.Text = fstiErrorObject.Description;
            listResult.Items.Add(String.Format("MessageFound: {0} ", fstiErrorObject.MessageFound));
            listResult.Items.Add(String.Format("MessageID: {0} ", fstiErrorObject.MessageID));
            listResult.Items.Add(String.Format("MessageSource: {0} ", fstiErrorObject.MessageSource));
            listResult.Items.Add(String.Format("Number: {0} ", fstiErrorObject.Number));
            listResult.Items.Add(String.Format("Fields in Error: {0} ", fstiErrorObject.NumberOfFieldsInError));
            for (int i = 0; i < fstiErrorObject.NumberOfFieldsInError; i++)
            {
                int field = fstiErrorObject.GetFieldNumber(i);
                listResult.Items.Add(String.Format("Field[{0}]: {1}", i, field));
                ITransactionField myField = transaction.get_Field(field);
                listResult.Items.Add(String.Format("Field name: {0}", myField.Name));
            }
        }
        private void AddBOM_Click(object sender, EventArgs e)//增加BOM
        {
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            BOMResult.Items.Clear();
            if (dgvBOMDetail.Rows.Count == 0)
            {
                MessageBox.Show("BOM信息为空！");
                return;
            }
            //MessageBox.Show("当前共有" + dgvBOMDetail.Rows.Count + "条待添加数据！");
            //string requiredquantity = string.Empty;
            int chenggong = 0;
            int shibai = 0;
            for (int i = 0; i < dgvBOMDetail.Rows.Count; i++)
            {
                BILL00 myBill = new BILL00();
                if (dgvBOMDetail["序号", i].Value.ToString() == "")
                    continue;
                myBill.Parent.Value = ParentItem.Text.ToString();
                myBill.PointOfUseID.Value = dgvBOMDetail["用于", i].Value.ToString();
                myBill.OperationSequenceNumber.Value = dgvBOMDetail["序号", i].Value.ToString();
                myBill.ComponentItemNumber.Value = dgvBOMDetail["子项", i].Value.ToString();
                string requiredquantity = dgvBOMDetail["数量", i].Value.ToString();

                if (requiredquantity.Length > 10)
                {
                    //MessageBox.Show("需求数量已经超过系统允许的最大位数10，只使用前10位的数据！");
                    if (requiredquantity.Substring(0, 1) == "0")
                        myBill.RequiredQuantity.Value = requiredquantity.Substring(1, 10).ToString();//从第2个开始，截取10个字符。
                    else
                        myBill.RequiredQuantity.Value = requiredquantity.Substring(0, 10).ToString();//从第1个开始，截取10个字符。
                }
                else
                {
                    myBill.RequiredQuantity.Value = requiredquantity;
                }
                if (dgvBOMDetail["损耗", i].Value.ToString().Length > 4)
                {
                    myBill.ScrapPercent.Value = dgvBOMDetail["损耗", i].Value.ToString().Substring(0, 4);
                }
                else
                {
                    myBill.ScrapPercent.Value = dgvBOMDetail["损耗", i].Value.ToString();
                }
                myBill.QuantityType.Value = dgvBOMDetail["量类", i].Value.ToString();
                myBill.ComponentType.Value = dgvBOMDetail["子类", i].Value.ToString();
                //DateTime src = DateTime.Now;
                //string result = src.AddMonths(-1).ToString("MMddyy");//为了避免车间人员操作出现错误，BOM中子项的生效日期提前一个月
                myBill.InEffectivityDate.Value = "100103";
                myBill.OutEffectivityDate.Value = "123179";
                if (_fstiClient.ProcessId(myBill, null))
                {
                    chenggong++;
                    // success, get the response and display it using a list box
                    BOMResult.Items.Add("Success:");

                    BOMResult.Items.Add(_fstiClient.CDFResponse);
                }
                else
                {
                    shibai++;
                    // failure, retrieve the error object 
                    // and then dump the information in the list box
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myBill, itemError, BOMResult);
                    dgvBOMDetail.Rows[i].Cells[0].Style.BackColor = Color.Red;
                }

            }
            MessageBox.Show(string.Format("共{0}条数据，成功{1}条，失败{2}条。", dgvBOMDetail.Rows.Count, chenggong, shibai));
            BOM.Rows[Convert.ToInt32(BOMhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
            #region  ITMB状态O
            ITMB01 myItmb = new ITMB01();
            myItmb.ItemNumber.Value = ParentItem.Text.ToString();
            myItmb.ItemStatus.Value = "O";      //状态0    
            if (_fstiClient.ProcessId(myItmb, null))
            { BOMResult.Items.Add("物料启用状态O设置成功"); }
            else
            {
                BOMResult.Items.Add("物料启用状态O设置失败");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmb, itemError, BOMResult);
            }
            #endregion
            AddBOM.Enabled = false;
            SubmitBOM.Enabled = true;
        }
        private void UpdateBOM_Click(object sender, EventArgs e)//修改BOM
        {
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            BOMResult.Items.Clear();
            if (dgvBOMDetail.Rows.Count == 0)
            {
                MessageBox.Show("BOM信息为空！");
                return;
            }
            #region 删除旧子项
            int shanchu = 0;
            int shanchushibai = 0;
            #region 获得父项ItemKey
            string fuxiang = ParentItem.Text.ToString();//父项物料编码
            string parentKey;
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("select ItemKey from _NoLock_FS_Item where ItemNumber = '" + fuxiang + "'", conn);
                parentKey = cmd.ExecuteScalar().ToString();
            }
            #endregion
            for (int i = 0; i < dgvBOMDetail.Rows.Count; i++)
            {
                BILL03 myBill = new BILL03();
                if (dgvBOMDetail["序号", i].Value.ToString() == "")
                    continue;
                myBill.Parent.Value = fuxiang;
                int xuhao = Convert.ToInt32(dgvBOMDetail["序号", i].Value);//序号
                #region 通过父项ItemKey、序号 查找 子项、用于。
                string zixiang = "", yongyu = "";
                using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                {
                    SqlCommand cmd = new SqlCommand("select PointOfUseID,ComponentItemNumber from _NoLock_FS_BillOfMaterial where ParentItemKey = '" + parentKey + "' and OperationSequenceNumber ='" + xuhao + "'", conn);
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    { continue; }
                    else if (dt.Rows.Count == 1)
                    {
                        zixiang = dt.Rows[0]["ComponentItemNumber"].ToString();
                        yongyu = dt.Rows[0]["PointOfUseID"].ToString();
                    }
                    else
                    { throw new Exception("有多条相同序号的记录"); }
                }
                #endregion
                myBill.OperationSequenceNumber.Value = xuhao.ToString().PadLeft(3, '0');
                myBill.ComponentItemNumber.Value = zixiang;
                myBill.PointOfUseID.Value = yongyu;
                if (_fstiClient.ProcessId(myBill, null))
                {
                    shanchu++;
                    // success, get the response and display it using a list box
                    BOMResult.Items.Add("0" + xuhao.ToString() + "删除成功:");
                    BOMResult.Items.Add(_fstiClient.CDFResponse);
                }
                else
                {
                    shanchushibai++;
                    // failure, retrieve the error object 
                    // and then dump the information in the list box
                    FSTIError itemError = _fstiClient.TransactionError;
                    BOMResult.Items.Add("0" + xuhao.ToString() + "删除失败:");
                    DumpErrorObject(myBill, itemError, BOMResult);
                }
            }
            #endregion
            #region 增加新子项
            int chenggong = 0;
            int shibai = 0;
            for (int i = 0; i < dgvBOMDetail.Rows.Count; i++)
            {
                if (dgvBOMDetail["状态", i].Value.ToString() == "删除") continue;
                BILL00 myBill = new BILL00();
                if (dgvBOMDetail["序号", i].Value.ToString() == "")
                    continue;
                myBill.Parent.Value = ParentItem.Text.ToString();
                myBill.PointOfUseID.Value = dgvBOMDetail["用于", i].Value.ToString();
                myBill.OperationSequenceNumber.Value = dgvBOMDetail["序号", i].Value.ToString();
                myBill.ComponentItemNumber.Value = dgvBOMDetail["子项", i].Value.ToString();
                string requiredquantity = dgvBOMDetail["数量", i].Value.ToString();

                if (requiredquantity.Length > 10)
                {
                    //MessageBox.Show("需求数量已经超过系统允许的最大位数10，只使用前10位的数据！");
                    if (requiredquantity.Substring(0, 1) == "0")
                        myBill.RequiredQuantity.Value = requiredquantity.Substring(1, 10).ToString();//从第2个开始，截取10个字符。
                    else
                        myBill.RequiredQuantity.Value = requiredquantity.Substring(0, 10).ToString();//从第1个开始，截取10个字符。
                }
                else
                {
                    myBill.RequiredQuantity.Value = requiredquantity;
                }
                if (dgvBOMDetail["损耗", i].Value.ToString().Length > 4)
                {
                    myBill.ScrapPercent.Value = dgvBOMDetail["损耗", i].Value.ToString().Substring(0, 4);
                }
                else
                {
                    myBill.ScrapPercent.Value = dgvBOMDetail["损耗", i].Value.ToString();
                }
                myBill.QuantityType.Value = dgvBOMDetail["量类", i].Value.ToString();
                myBill.ComponentType.Value = dgvBOMDetail["子类", i].Value.ToString();
                //DateTime src = DateTime.Now;
                //string result = src.AddMonths(-1).ToString("MMddyy");//为了避免车间人员操作出现错误，BOM中子项的生效日期提前一个月
                myBill.InEffectivityDate.Value = "100103";
                myBill.OutEffectivityDate.Value = "123179";
                if (_fstiClient.ProcessId(myBill, null))
                {
                    chenggong++;
                    // success, get the response and display it using a list box
                    BOMResult.Items.Add("Success:");

                    BOMResult.Items.Add(_fstiClient.CDFResponse);
                }
                else
                {
                    shibai++;
                    // failure, retrieve the error object 
                    // and then dump the information in the list box
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myBill, itemError, BOMResult);
                    dgvBOMDetail.Rows[i].HeaderCell.Style.BackColor = Color.Red;
                }

            }
            MessageBox.Show(string.Format("删除成功{0}条旧子项，删除失败{4}条旧子项；增加{1}条新子项，成功{2}条，失败{3}条。", shanchu, dgvBOMDetail.Rows.Count, chenggong, shibai, shanchushibai));
            BOM.Rows[Convert.ToInt32(BOMhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
            #endregion
            #region  ITMB状态O
            ITMB01 myItmb = new ITMB01();
            myItmb.ItemNumber.Value = ParentItem.Text.ToString();
            myItmb.ItemStatus.Value = "O";      //状态0    
            if (_fstiClient.ProcessId(myItmb, null))
            { BOMResult.Items.Add("物料启用状态O设置成功"); }
            else
            {
                MessageBox.Show("物料启用状态O设置失败");
                BOMResult.Items.Add("物料启用状态O设置失败");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmb, itemError, BOMResult);
            }
            #endregion
            UpdateBOM.Enabled = false;
            SubmitBOM.Enabled = true;
        }
        private void jihuo_Click(object sender, EventArgs e)//激活增加BOM，更新BOM按钮
        {
            AddBOM.Enabled = true;
            UpdateBOM.Enabled = true;
            SubmitBOM.Enabled = true;
        }
        private void jihuoVendor_Click(object sender, EventArgs e)//激活增加供应商，更新供应商按钮
        {
            AddVendor.Enabled = true;
            UpdateVendor.Enabled = true;
            SubmitVendor.Enabled = true;
        }
        private void GetVendor_Click(object sender, EventArgs e)//获得增加供应商流程
        {
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and PROCESSNAME='RY增加供应商流程' and TASKUSER='BPM/zuojinguo' and ENDTIME >'2019/5/18' and STEPLABEL='系统添加'");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY增加供应商流程' and TASKUSER='BPM/zuojinguo' and STEPLABEL='系统添加'");
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY增加供应商流程'  and STEPLABEL='系统添加'");
            List<Vendorliucheng> list1 = new List<Vendorliucheng>();
            foreach (DataRow dr in Incidents.Rows)
            {
                Vendorliucheng VENDOR1 = TolistVendor(SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_ZJGYS] where REV_INCIDENT=" + dr[0]));
                list1.Add(VENDOR1);

            }

            dgvVendor.DataSource = list1;
            for (int i = 0; i < this.dgvVendor.Columns.Count; i++)
            {
                this.dgvVendor.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvVendor.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            VendorResult.Items.Clear();
        }
        private Vendorliucheng TolistVendor(DataTable dt)//dt转化为Vendorliucheng的方法
        {
            if (dt.Rows.Count == 1)
            {
                Vendorliucheng bom1 = new Vendorliucheng();
                bom1.申请人 = (string)(dt.Rows[0]["REV_CREATER_NAME"]);
                bom1.发起部门 = (string)(dt.Rows[0]["REV_CREATER_DPT"]);
                bom1.申请时间 = (DateTime)(dt.Rows[0]["REV_CREATER_DATE"]);
                bom1.流水号 = (Int32)(dt.Rows[0]["REV_INCIDENT"]);
                if ((string)(dt.Rows[0]["ZJXG"]) == "zj")
                    bom1.申请方式 = "增加";
                if ((string)(dt.Rows[0]["ZJXG"]) == "xg")
                    bom1.申请方式 = "修改";
                bom1.供应商代码 = (string)(dt.Rows[0]["GYSDM"]);
                bom1.供应商名称 = (string)(dt.Rows[0]["GYSMC"]);
                bom1.摘要 = (string)(dt.Rows[0]["ZY"]);
                bom1.银行名称 = (string)(dt.Rows[0]["YHMC"]);
                bom1.银行账号 = (string)(dt.Rows[0]["YHZH"]);
                return bom1;
            }
            else
            {
                throw new Exception("有0个或多个的流水号");
            }
        }
        private void dgvVendor_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//供应商流程列表的cell双击事件
        {
            VendorResult.Items.Clear();
            int rowindex = e.RowIndex;

            if (rowindex != -1)
            {
                if (dgvVendor.Rows[rowindex].DefaultCellStyle.ForeColor != Color.Red)
                {
                    dgvVendor.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                    for (int a = 0; a < dgvVendor.Rows.Count; a++)
                    {
                        if (a != rowindex && dgvVendor.Rows[a].DefaultCellStyle.ForeColor != Color.Red)
                            dgvVendor.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
                VendorItem.Text = dgvVendor.Rows[rowindex].Cells["供应商代码"].Value.ToString().Trim().ToUpper();
                Vendorliushuihao.Text = dgvVendor.Rows[rowindex].Cells["流水号"].Value.ToString().Trim().ToUpper();
                Vendorhanghao.Text = (rowindex + 1).ToString();
                DataTable dt = SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_ZJGYS] where REV_INCIDENT=" + Vendorliushuihao.Text);
                if (dt.Rows.Count == 1)
                {
                    tbVendorCode.Text = dt.Rows[0]["GYSDM"].ToString().Trim().ToUpper();
                    tbVendorName.Text = dt.Rows[0]["GYSMC"].ToString().Trim().ToUpper();
                    tbVendorContact.Text = dt.Rows[0]["LXR"].ToString().Trim().ToUpper();
                    tbVendorContactPhone.Text = dt.Rows[0]["LXDH"].ToString().Trim().ToUpper();

                    tbVendorAddress.Text = dt.Rows[0]["GYSDZ"].ToString().Trim().ToUpper();
                    cbVendorState.Text = dt.Rows[0]["ZT"].ToString().Trim().Substring(0, 1).ToUpper();
                    tbVendorPayeeName.Text = dt.Rows[0]["SKRMC"].ToString().Trim().ToUpper();
                    GYSBZ.Text = dt.Rows[0]["GYSBZ"].ToString().Trim().ToUpper();

                    cbMoneyType.Text = dt.Rows[0]["ZKHB"].ToString().ToUpper().Trim();
                    tbVendorCurrencyCode.Text = dt.Rows[0]["HBDM"].ToString().Trim().ToUpper();
                    cbPaymentForm.Text = dt.Rows[0]["FKFS"].ToString().Trim().ToUpper();
                    cbStandardTerm.Text = dt.Rows[0]["BZTK"].ToString().Trim().ToUpper();
                    cbVendorClass.Text = dt.Rows[0]["GYSFL"].ToString().Trim().ToUpper();

                    tbVendorDepositBank.Text = dt.Rows[0]["YHMC"].ToString().Trim().ToUpper();
                    tbVendorBankAccount.Text = dt.Rows[0]["YHZH"].ToString().Trim().ToUpper();
                    tbVendorTaxCode.Text = dt.Rows[0]["VATDM"].ToString().Trim().ToUpper();

                    tbVendorAccountant.Text = dt.Rows[0]["KJLXR"].ToString().Trim().ToUpper();
                    tbVendorAccountantPhone.Text = dt.Rows[0]["KJLXDH"].ToString().Trim().ToUpper();
                    tbVendorzhaiyao.Text = dt.Rows[0]["ZY"].ToString().Trim().ToUpper();

                    tbPayableAccountWithInvoice.Text = dt.Rows[0]["YPYFZK"].ToString().Trim().ToUpper();
                    tbPayableAccountWithoutInvoice.Text = dt.Rows[0]["WPYFZK"].ToString().Trim().ToUpper();
                    Vendormiaoshu.Text = dt.Rows[0]["MS"].ToString().Trim().ToUpper();
                    Vendorzuzhihao.Text = dt.Rows[0]["ZZH"].ToString().Trim().ToUpper();
                    #region
                    if (tbVendorName.Text != tbVendorPayeeName.Text)
                    {
                        MessageBox.Show("供应商名称跟收款人名称不一致");
                    }
                    if (StrLength(tbVendorName.Text) > 60)
                    { MessageBox.Show("供应商名称字符数大于60"); }
                    else if (StrLength(tbVendorName.Text) > 35)
                    {
                        { MessageBox.Show("供应商名称字符数大于35小于60,GLOS供应商名称不全"); }
                    }
                    if (StrLength(tbVendorAddress.Text) > 60)
                    { MessageBox.Show("供应商地址字符数大于60"); }
                    #endregion
                }
                else
                { throw new Exception("有0条或多条该流水号的数据。"); }
                if (dgvVendor.Rows[rowindex].Cells["申请方式"].Value.ToString() == "增加")
                { AddVendor.Enabled = true; UpdateVendor.Enabled = false; }
                if (dgvVendor.Rows[rowindex].Cells["申请方式"].Value.ToString() == "修改")
                { UpdateVendor.Enabled = true; AddVendor.Enabled = false; }
            }
        }
        private void AddVendor_Click(object sender, EventArgs e)//增加供应商
        {
            VendorResult.Items.Clear();
            if (tbVendorCode.Text == "")
            {
                MessageBox.Show("供应商信息为空！");
                return;
            }

            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            if (StrLength(tbVendorTaxCode.Text.Trim()) > 20)
            {
                MessageBox.Show("税号大于20位!");
                return;
            }

            #region 检查供应商名称是否重复
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                string CustomerName = EncodingLD.GetString(EncodingCH.GetBytes(tbVendorName.Text.Trim()));
                SqlCommand cmd = new SqlCommand("select VendorID from _NoLock_FS_Vendor where VendorName = '" + CustomerName + "'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count > 0)
                {
                    MessageBox.Show("有相同供应商名称的记录，请检查" + dtcust.Rows[0][0].ToString());
                    return;
                }

                cmd = new SqlCommand("SELECT VendorID FROM _NoLock_FS_Vendor  WHERE  TaxID = '" + tbVendorTaxCode.Text.Trim() + "'", conn);
                sda = new SqlDataAdapter(cmd);
                dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count > 0)
                {
                    MessageBox.Show("有相同供应商税号的记录，请检查" + dtcust.Rows[0][0].ToString() + "!");
                    return;
                }
            }
            #endregion

            int i = 0;//判断是否全部成功标识
            if (ADDGLOS(Vendorzuzhihao.Text.Trim(), Vendormiaoshu.Text.Trim(), VendorResult) == false)
                i = 1;
            if (ADDGLAV(Vendorzuzhihao.Text.Trim(), "212100", VendorResult) == false)
                i = 1;
            if (ADDGLAV(Vendorzuzhihao.Text.Trim(), "212101", VendorResult) == false)
                i = 1;
            #region  添加供应商VEID
            VEID00 myVendor = new VEID00();
            myVendor.VendorID.Value = tbVendorCode.Text.ToString();
            string vendorname = tbVendorName.Text.ToString().Trim();
            myVendor.VendorName.Value = vendorname;//供应商名称最多60个字符
            myVendor.VendorContact.Value = tbVendorContact.Text.ToString();
            myVendor.VendorContactPhone.Value = tbVendorContactPhone.Text.ToString();

            //供应商的状态都是通过一个字符A-可用的，P-淘汰的，I-已停用来进行标示，此处进行判断
            if (cbVendorState.Text.ToString().Contains("A") || cbVendorState.Text.ToString().Contains("C"))
            {
                myVendor.VendorStatus.Value = "A";
            }
            else if (cbVendorState.Text.ToString().Contains("P"))
            {
                myVendor.VendorStatus.Value = "P";
            }
            else
            {
                myVendor.VendorStatus.Value = "I";
            }

            if (_fstiClient.ProcessId(myVendor, null))
            {
                VendorResult.Items.Add("VEID供应商添加成功:");
                VendorResult.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                i = 1;
                VendorResult.Items.Add("VEID供应商添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myVendor, itemError, VendorResult);
                MessageBox.Show("VEID供应商添加失败!");
                return;
            }
            #endregion
            #region  添加供应商明细VEIDF8
            VEID03 myVendorDetail = new VEID03();
            myVendorDetail.VendorID.Value = tbVendorCode.Text.ToString();
            string vendoraddress = tbVendorAddress.Text.ToString();
            if (vendoraddress.Length > 60)
            {
                myVendorDetail.VendorAddress1.Value = vendoraddress.Substring(0, 60);
                myVendorDetail.PayeeAddress1.Value = vendoraddress.Substring(0, 60);
            }
            else
            {
                myVendorDetail.VendorAddress1.Value = vendoraddress;
                myVendorDetail.PayeeAddress1.Value = vendoraddress;
            }

            myVendorDetail.PayeeName1.Value = tbVendorPayeeName.Text.ToString();
            myVendorDetail.AccountingContact.Value = tbVendorAccountant.Text.ToString();
            myVendorDetail.AccountingContactPhone.Value = tbVendorAccountantPhone.Text.ToString();
            //供应商的类型分几个大类，根据选择的不同，写入的数据也不同
            if (cbVendorClass.Text.ToString().Contains("M"))
            {
                myVendorDetail.VendorClass6.Value = "M";
                myVendorDetail.VendorClass7.Value = "原料供";
            }
            else if (cbVendorClass.Text.ToString().Contains("P"))
            {
                myVendorDetail.VendorClass6.Value = "P";
                myVendorDetail.VendorClass7.Value = "包材供";
            }
            else if (cbVendorClass.Text.ToString().Contains("E"))
            {
                myVendorDetail.VendorClass6.Value = "E";
                myVendorDetail.VendorClass7.Value = "设备供";
            }
            else if (cbVendorClass.Text.ToString().Contains("A"))
            {
                myVendorDetail.VendorClass6.Value = "A";
                myVendorDetail.VendorClass7.Value = "五金供";
            }
            else if (cbVendorClass.Text.ToString().Contains("S"))
            {
                myVendorDetail.VendorClass6.Value = "S";
                myVendorDetail.VendorClass7.Value = "服务供";
            }
            else
            { throw new Exception("供应商分类不在MPEAS中"); }
            //货币类型
            if (cbMoneyType.Text.ToString().Contains("L") || cbMoneyType.Text.ToString() == "")
            {
                //默认L 00000
                //myVendorDetail.VendorControllingCode.Value = "L";
                //myVendorDetail.VendorCurrencyCode.Value = "00000";
            }
            else if (cbMoneyType.Text.ToString().ToUpper().Contains("F") || cbMoneyType.Text.ToString().ToUpper().Contains("f"))
            {
                MessageBox.Show("货币类型是F，已默认添加为L，添加完毕后请检查流程！");
                //myVendorDetail.VendorControllingCode.Value = "F";
                //myVendorDetail.VendorCurrencyCode.Value = tbVendorCurrencyCode.Text.ToString().Substring(0, 5);
            }

            //标准条款
            if (cbStandardTerm.Text.ToString().Contains("S"))
            {
                myVendorDetail.TermsCode.Value = "S";
            }
            else if (cbStandardTerm.Text.ToString().Contains("M"))
            {
                myVendorDetail.TermsCode.Value = "M";

            }
            else if (cbStandardTerm.Text.ToString().Contains("D"))
            {
                myVendorDetail.TermsCode.Value = "D";
            }
            else { throw new Exception("标准条款不在SMD中"); }

            //付款方式
            if (cbPaymentForm.Text.ToString().Contains("B"))
            {
                myVendorDetail.PaymentForm.Value = "B";
            }
            else if (cbPaymentForm.Text.ToString().Contains("C"))
            {
                myVendorDetail.PaymentForm.Value = "C";
            }
            else if (cbPaymentForm.Text.ToString().Contains("D"))
            {
                myVendorDetail.PaymentForm.Value = "D";
            }
            else if (cbPaymentForm.Text.ToString().Contains("G"))
            {
                myVendorDetail.PaymentForm.Value = "D";//G、T需要银行路线 所以录入D
            }
            else if (cbPaymentForm.Text.ToString().Contains("T"))
            {
                myVendorDetail.PaymentForm.Value = "D";
            }
            else { throw new Exception("付款方式不在BCDGT中"); }
            myVendorDetail.BankName.Value = tbVendorDepositBank.Text.ToString();
            myVendorDetail.BankAccountNumber.Value = tbVendorBankAccount.Text.ToString().Replace(" ", "");
            //                MessageBox.Show(tbVendorTaxCode.Text.ToString().Trim());
            //                myVendorDetail.TaxID.Value = tbVendorTaxCode.Text.ToString().Trim();
            myVendorDetail.VATID.Value = tbVendorTaxCode.Text.ToString().Replace(" ", "");
            DateTime src = DateTime.Now;
            string result = src.ToString("MMddyy");
            myVendorDetail.VendorStartDate.Value = result;

            if (_fstiClient.ProcessId(myVendorDetail, null))
            {
                VendorResult.Items.Add("供应商明细添加成功:");
                VendorResult.Items.Add(_fstiClient.CDFResponse);

            }
            else
            {
                i = 1;
                VendorResult.Items.Add("供应商明细添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myVendorDetail, itemError, VendorResult);
            }
            #endregion
            #region VEIDF-F8-供应商分类账账户
            VEID08 myVendorGL = new VEID08();
            myVendorGL.VendorID.Value = tbVendorCode.Text.ToString();
            myVendorGL.VoucheredAPAccount.Value = tbPayableAccountWithInvoice.Text.ToString().Trim();
            myVendorGL.UnvoucheredAPAccount.Value = tbPayableAccountWithoutInvoice.Text.ToString().Trim();

            if (_fstiClient.ProcessId(myVendorGL, null))
            {
                VendorResult.Items.Add("供应商分类账账户添加成功:");
                VendorResult.Items.Add(_fstiClient.CDFResponse);

            }
            else
            {
                i = 1;
                VendorResult.Items.Add("供应商分类账账户添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myVendorGL, itemError, VendorResult);
            }
            #endregion
            #region 检查是否有重复的供应商名称
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("select VendorName from _NoLock_FS_Vendor where VendorID = '" + tbVendorCode.Text.Trim() + "'", conn);
                object vendornamelading = cmd.ExecuteScalar();
                if (vendornamelading == null)
                { MessageBox.Show("没有该供应商代码的记录!"); }
                else
                {
                    string VendorName = vendornamelading.ToString();
                    cmd = new SqlCommand("select count(*) from _NoLock_FS_Vendor where VendorName = '" + VendorName + "'", conn);
                    if (Convert.ToInt32(cmd.ExecuteScalar()) != 1)
                    {
                        MessageBox.Show("有多个相同供应商名称的记录，请检查！");
                        MessageBox.Show("有多个相同供应商名称的记录，请检查！");
                    }
                }
            }
            #endregion
            if (i == 0)
            {
                MessageBox.Show("供应商全部添加成功");
                AddVendor.Enabled = false;
                SubmitVendor.Enabled = true;
                #region 供应商信息groupBox4内控件清空
                foreach (Control control in groupBox4.Controls)
                {
                    if (!(control is Label))
                    {
                        control.Text = "";
                    }
                }
                #endregion
                try
                {
                    dgvVendor.Rows[Convert.ToInt32(Vendorhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                }
                catch (Exception)
                {

                }
            }
            else
            { MessageBox.Show("供应商部分添加失败，请检查原因"); }

        }
        private bool ADDGLOS(string strCode, string strName, ListBox listResult)//增加三级组织号GLOS
        {
            //strPayableAccountWithInvoice格式：1ABC-DE-FGH-XXXXXX
            string GLOC_1 = strCode.Substring(0, 4);
            string GLOC_2 = strCode.Substring(0, 7);
            string GLOC_3 = strCode.Substring(0, 11);


            GLOS03 FirstGLOC = new GLOS03();
            GLOS03 SecondGLOC = new GLOS03();
            GLOS03 ThirdGLOC = new GLOS03();

            //添加第一级的机构号的信息
            FirstGLOC.ParentGLOrganization.Value = "1";
            FirstGLOC.ChildGLOrganization.Value = GLOC_1;
            FirstGLOC.ConsolidationGLOrganization.Value = GLOC_1;
            FirstGLOC.IsGLOrganizationActiveOrInactive.Value = "A";


            //添加第二级的机构号信息
            SecondGLOC.ParentGLOrganization.Value = GLOC_1;
            SecondGLOC.ChildGLOrganization.Value = GLOC_2;
            SecondGLOC.ConsolidationGLOrganization.Value = GLOC_2;
            SecondGLOC.IsGLOrganizationActiveOrInactive.Value = "A";

            //添加第三极的机构号信息

            ThirdGLOC.ParentGLOrganization.Value = GLOC_2;
            ThirdGLOC.ChildGLOrganization.Value = GLOC_3;
            if (StrLength(strName) > 35)
            {
                int Len1 = strName.Length;
                for (int i = strName.Length; i > 0; i--)
                {
                    if (StrLength(strName.Substring(0, i)) <= 35)
                    { Len1 = i; break; }
                }
                ThirdGLOC.GLOrganizationGroupDescription.Value = strName.Substring(0, Len1);
            }
            else
            {
                ThirdGLOC.GLOrganizationGroupDescription.Value = strName;
            }
            ThirdGLOC.ConsolidationGLOrganization.Value = GLOC_3;
            ThirdGLOC.IsGLOrganizationActiveOrInactive.Value = "A";

            _fstiClient.ProcessId(FirstGLOC, null);
            _fstiClient.ProcessId(SecondGLOC, null);

            if (_fstiClient.ProcessId(ThirdGLOC, null))
            {
                listResult.Items.Add("添加三级组织机构号成功:");
                // success, get the response and display it using a list box

                listResult.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listResult.Items.Add("添加三级组织机构号失败:");
                // failure, retrieve the error object 
                // and then dump the information in the list box
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(ThirdGLOC, itemError, listResult);
                
            }
            return true;
        }
        public bool ADDGLAV(string OrganizationCode, string strCode, ListBox listResult)//增加账号GLAV
        {
            GLAV00 myGLAV = new GLAV00();
            //GLAV屏幕，通过113100添加组织机构号
            myGLAV.GLAccountGroup.Value = strCode;
            myGLAV.GLAccountValidationCode.Value = "1";
            myGLAV.GLOrganization.Value = OrganizationCode;
            if (_fstiClient.ProcessId(myGLAV, null))
            {
                listResult.Items.Add(strCode + "添加组织帐号成功:");
                listResult.Items.Add(_fstiClient.CDFResponse);
                return true;
            }
            else
            {
                listResult.Items.Add(strCode + "添加组织帐号失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myGLAV, itemError, listResult);
                return false;
            }

        }
        public bool ADDGLAV(string OrganizationCode, string strCode)//增加账号GLAV不在ListResult中显示结果
        {
            
            #region FSTI增加GLAV
            GLAV00 myGLAV = new GLAV00();
            myGLAV.GLAccountGroup.Value = strCode;
            myGLAV.GLAccountValidationCode.Value = "1";
            myGLAV.GLOrganization.Value = OrganizationCode;
            if (_fstiClient.ProcessId(myGLAV, null))
            {
                return true;
            }
            else
            {
                return false;
            }
            #endregion
        }
        private void UpdateVendor_Click(object sender, EventArgs e)//修改供应商
        {
            VendorResult.Items.Clear();
            if (tbVendorCode.Text == "")
            {
                MessageBox.Show("供应商信息为空！");
                return;
            }

            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text.ToString())
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            #region 检查供应商名称是否重复
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                string CustomerName = EncodingLD.GetString(EncodingCH.GetBytes(tbVendorName.Text.Trim()));
                SqlCommand cmd = new SqlCommand("select VendorID from _NoLock_FS_Vendor where VendorName = '" + CustomerName + "' and VendorID !='" + tbVendorCode.Text.Trim() + "'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count > 0)
                {
                    MessageBox.Show("有相同供应商名称的记录，请检查" + dtcust.Rows[0][0].ToString());
                    return;
                }

            }

            int i = 0;//判断是否成功标识
            #endregion
            #region  修改供应商VEID
            VEID01 myVendor = new VEID01();
            myVendor.VendorID.Value = tbVendorCode.Text.ToString();
            string vendorname = tbVendorName.Text.ToString().Trim();
            myVendor.VendorName.Value = vendorname;//供应商名称最多60个字符
            myVendor.VendorContact.Value = tbVendorContact.Text.ToString();
            myVendor.VendorContactPhone.Value = tbVendorContactPhone.Text.ToString();
            //供应商的状态都是通过一个字符A-可用的，P-淘汰的，I-已停用来进行标示，此处进行判断
            if (cbVendorState.Text.ToString().Contains("A") || cbVendorState.Text.ToString().Contains("C"))
            {
                myVendor.VendorStatus.Value = "A";
            }
            else if (cbVendorState.Text.ToString().Contains("P"))
            {
                myVendor.VendorStatus.Value = "P";
            }
            else
            {
                myVendor.VendorStatus.Value = "I";
            }

            if (_fstiClient.ProcessId(myVendor, null))
            {
                VendorResult.Items.Add("VEID供应商修改成功:");
                VendorResult.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                i = 1;
                VendorResult.Items.Add("VEID供应商修改失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myVendor, itemError, VendorResult);
            }
            #endregion
            #region  修改供应商明细VEIDF8
            VEID03 myVendorDetail = new VEID03();
            myVendorDetail.VendorID.Value = tbVendorCode.Text.ToString();
            string vendoraddress = tbVendorAddress.Text.ToString();
            if (vendoraddress.Length > 60)
            {
                myVendorDetail.VendorAddress1.Value = vendoraddress.Substring(0, 60);
                myVendorDetail.PayeeAddress1.Value = vendoraddress.Substring(0, 60);
            }
            else
            {
                myVendorDetail.VendorAddress1.Value = vendoraddress;
                myVendorDetail.PayeeAddress1.Value = vendoraddress;
            }

            myVendorDetail.PayeeName1.Value = tbVendorPayeeName.Text.ToString();
            myVendorDetail.AccountingContact.Value = tbVendorAccountant.Text.ToString();
            myVendorDetail.AccountingContactPhone.Value = tbVendorAccountantPhone.Text.ToString();
            //供应商的类型分几个大类，根据选择的不同，写入的数据也不同
            if (cbVendorClass.Text.ToString().Contains("M"))
            {
                myVendorDetail.VendorClass6.Value = "M";
                myVendorDetail.VendorClass7.Value = "原料供";
            }
            else if (cbVendorClass.Text.ToString().Contains("P"))
            {
                myVendorDetail.VendorClass6.Value = "P";
                myVendorDetail.VendorClass7.Value = "包材供";
            }
            else if (cbVendorClass.Text.ToString().Contains("E"))
            {
                myVendorDetail.VendorClass6.Value = "E";
                myVendorDetail.VendorClass7.Value = "设备供";
            }
            else if (cbVendorClass.Text.ToString().Contains("A"))
            {
                myVendorDetail.VendorClass6.Value = "A";
                myVendorDetail.VendorClass7.Value = "五金供";
            }
            else if (cbVendorClass.Text.ToString().Contains("S"))
            {
                myVendorDetail.VendorClass6.Value = "S";
                myVendorDetail.VendorClass7.Value = "服务供";
            }
            else
            { throw new Exception("供应商分类不在MPEAS中"); }
            //货币类型
            if (cbMoneyType.Text.ToString().Contains("L") || cbMoneyType.Text.ToString() == "")
            {
                //默认L 00000
                //myVendorDetail.VendorControllingCode.Value = "L";
                //myVendorDetail.VendorCurrencyCode.Value = "00000";
            }
            else if (cbMoneyType.Text.ToString().ToUpper().Contains("F") || cbMoneyType.Text.ToString().ToUpper().Contains("f"))
            {
                MessageBox.Show("货币类型是F，已默认添加为L，添加完毕后请检查流程！");
                //myVendorDetail.VendorControllingCode.Value = "F";
                //myVendorDetail.VendorCurrencyCode.Value = tbVendorCurrencyCode.Text.ToString().Substring(0, 5);
            }
            //标准条款
            if (cbStandardTerm.Text.ToString().Contains("S"))
            {
                myVendorDetail.TermsCode.Value = "S";
            }
            else if (cbStandardTerm.Text.ToString().Contains("M"))
            {
                myVendorDetail.TermsCode.Value = "M";

            }
            else if (cbStandardTerm.Text.ToString().Contains("D"))
            {
                myVendorDetail.TermsCode.Value = "D";
            }
            else { throw new Exception("标准条款不在SMD中"); }

            //付款方式
            if (cbPaymentForm.Text.ToString().Contains("B"))
            {
                myVendorDetail.PaymentForm.Value = "B";
            }
            else if (cbPaymentForm.Text.ToString().Contains("C"))
            {
                myVendorDetail.PaymentForm.Value = "C";
            }
            else if (cbPaymentForm.Text.ToString().Contains("D"))
            {
                myVendorDetail.PaymentForm.Value = "D";
            }
            else if (cbPaymentForm.Text.ToString().Contains("G"))
            {
                myVendorDetail.PaymentForm.Value = "D";//GT需要银行路线所以录入D
            }
            else if (cbPaymentForm.Text.ToString().Contains("T"))
            {
                myVendorDetail.PaymentForm.Value = "D";
            }
            else { throw new Exception("付款方式不在BCDGT中"); }
            myVendorDetail.BankName.Value = tbVendorDepositBank.Text.ToString();
            myVendorDetail.BankAccountNumber.Value = tbVendorBankAccount.Text.ToString().Replace(" ", "");
            //                MessageBox.Show(tbVendorTaxCode.Text.ToString().Trim());
            //                myVendorDetail.TaxID.Value = tbVendorTaxCode.Text.ToString().Trim();
            if (tbVendorTaxCode.Text.ToString().Replace(" ", "") == "")
            { myVendorDetail.VATID.Value = "无"; }
            else
            {
                myVendorDetail.VATID.Value = tbVendorTaxCode.Text.ToString().Replace(" ", "");
            }
            //DateTime src = DateTime.Now;
            //string result = src.ToString("MMddyy");
            //myVendorDetail.VendorStartDate.Value = result;

            if (_fstiClient.ProcessId(myVendorDetail, null))
            {
                VendorResult.Items.Add("供应商明细修改成功:");
                VendorResult.Items.Add(_fstiClient.CDFResponse);

            }
            else
            {
                i = 1;
                VendorResult.Items.Add("供应商明细修改失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myVendorDetail, itemError, VendorResult);
            }
            #endregion
            #region VEIDF-F8-供应商分类账账户修改
            //VEID08 myVendorGL = new VEID08();
            //myVendorGL.VendorID.Value = tbVendorCode.Text.ToString();
            //myVendorGL.VoucheredAPAccount.Value = tbPayableAccountWithInvoice.Text.ToString().Trim();
            //myVendorGL.UnvoucheredAPAccount.Value = tbPayableAccountWithoutInvoice.Text.ToString().Trim();

            //if (_fstiClient.ProcessId(myVendorGL, null))
            //{
            //    VendorResult.Items.Add("供应商分类账账户修改成功:");
            //    VendorResult.Items.Add(_fstiClient.CDFResponse);
            //    MessageBox.Show("供应商修改成功");
            //}
            //else
            //{
            //    VendorResult.Items.Add("供应商分类账账户修改失败:");
            //    FSTIError itemError = _fstiClient.TransactionError;
            //    DumpErrorObject(myVendorGL, itemError, VendorResult);
            //}
            #endregion
            #region 修改GLOS组织描述
            GLOS04 ThirdGLOC = new GLOS04();
            ThirdGLOC.ParentGLOrganization.Value = Vendorzuzhihao.Text.Substring(0, 7);
            ThirdGLOC.ChildGLOrganization.Value = Vendorzuzhihao.Text;
            ThirdGLOC.GLOrganizationGroupDescription.Value = Vendormiaoshu.Text;
            ThirdGLOC.ConsolidationGLOrganization.Value = Vendorzuzhihao.Text;
            ThirdGLOC.IsGLOrganizationActiveOrInactive.Value = "A";

            if (_fstiClient.ProcessId(ThirdGLOC, null))
            {
                VendorResult.Items.Add("修改三级组织机构号成功:");
                // success, get the response and display it using a list box

                VendorResult.Items.Add(_fstiClient.CDFResponse);
                //MessageBox.
            }
            else
            {
                i = 1;
                VendorResult.Items.Add("修改三级组织机构号失败:");
                // failure, retrieve the error object 
                // and then dump the information in the list box
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(ThirdGLOC, itemError, VendorResult);

            }
            #endregion
            #region 检查是否有重复的供应商名称
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("select VendorName from _NoLock_FS_Vendor where VendorID = '" + tbVendorCode.Text.Trim() + "'", conn);
                object vendornamelading = cmd.ExecuteScalar();
                if (vendornamelading == null)
                { MessageBox.Show("没有该供应商代码的记录!"); }
                else
                {
                    string VendorName = vendornamelading.ToString();
                    cmd = new SqlCommand("select count(*) from _NoLock_FS_Vendor where VendorName = '" + VendorName + "'", conn);
                    if (Convert.ToInt32(cmd.ExecuteScalar()) != 1)
                    {
                        MessageBox.Show("有多个相同供应商名称的记录，请检查！");
                        MessageBox.Show("有多个相同供应商名称的记录，请检查！");
                    }
                }
            }
            #endregion
            if (i == 0)
            {
                MessageBox.Show("供应商全部修改成功"); UpdateVendor.Enabled = false;
                SubmitVendor.Enabled = true;
                #region 供应商信息groupBox4内控件清空
                foreach (Control control in groupBox4.Controls)
                {
                    if (!(control is Label))
                    {
                        control.Text = "";
                    }
                }
                #endregion
                try
                {
                    dgvVendor.Rows[Convert.ToInt32(Vendorhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                }
                catch (Exception)
                {

                    throw;
                }
            }
            else
            { MessageBox.Show("供应商部分修改失败，请检查原因"); }

        }
        private void AddITMB_Click(object sender, EventArgs e)//增加物料成本按钮
        {
            ITMBResult.Items.Clear();
            if (dgvItmbDetail.Rows.Count == 0)
            {
                MessageBox.Show("物料成本信息为空！");
                return;
            }
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            int AllSuccess = 1;
            for (int i = 0; i < dgvItmbDetail.Rows.Count; i++)
            {
                ITMBResult.Items.Add("*********************第" + (i + 1) + "行*********************");
                if (AddItemMasterFile(i))//添加ITMB
                {
                    int chenggong = 1;
                    if (AddItemMasterFileDetail(i) == false)//添加ITMB物料明细（F8）
                        chenggong = 0;
                    if (AddItemLotTraceAndSerializeDetail(i) == false)//添加ITMB批号明细（alt+F8）
                        chenggong = 0;
                    if (AddItemMasterFilePlanDetail(i) == false)//添加ITMB计划明细（F9）
                        chenggong = 0;
                    #region A类F/S类ITMB及计划明细调整
                    string materialcode = dgvItmbDetail["物料代码", i].Value.ToString().Trim();
                    string fenlei = materialcode.Substring(0, 1).ToUpper();
                    if (fenlei == "A")
                    {
                        ITMB01 myItmb = new ITMB01();
                        myItmb.ItemNumber.Value = materialcode;
                        myItmb.OrderPolicy.Value = "0";      //订货策略0    
                        if (_fstiClient.ProcessId(myItmb, null))
                        { ITMBResult.Items.Add("A类订货策略0修正成功"); }
                        else
                        {
                            chenggong = 0;
                            ITMBResult.Items.Add("A类订货策略0修正失败");
                        }
                    }
                    if (fenlei == "F" || fenlei == "S")
                    {
                        ITMB01 myItmb = new ITMB01();
                        myItmb.ItemNumber.Value = materialcode;
                        myItmb.OrderPolicy.Value = "3";      //订货策略3   
                        if (_fstiClient.ProcessId(myItmb, null))
                        { ITMBResult.Items.Add("FS类订货策略3修正成功"); }
                        else
                        { ITMBResult.Items.Add("FS类订货策略3修正失败"); chenggong = 0; }

                        ITMB03 myItmb03 = new ITMB03();//ITMB计划明细
                        myItmb03.ItemNumber.Value = materialcode;
                        myItmb03.LotSizeDays.Value = dgvItmbDetail["批量订货天数", i].Value.ToString().Trim();      //批量订货天数 
                        if (_fstiClient.ProcessId(myItmb03, null))
                        { ITMBResult.Items.Add("FS类批量订货天数修正成功"); }
                        else
                        { ITMBResult.Items.Add("FS类批量订货天数修正失败"); chenggong = 0; }
                    }
                    #endregion
                    if (AddItemCostData(i) == false)//添加ITMC成本
                        chenggong = 0;
                    if (AddItemProductLineAndInventoryAccount(i) == false)//添加ITMC产品线及账号
                        chenggong = 0;

                    if (chenggong == 1)
                    {
                        //MessageBox.Show(string.Format("第{0}行物料({1})全部添加成功!", i + 1, materialcode));
                        //AddITMB.Enabled = false;
                        //SubmitITMB.Enabled = true;
                    }
                    else
                    {
                        dgvItmbDetail.Rows[i].HeaderCell.Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells[0].Style.BackColor = Color.Red;
                        MessageBox.Show(string.Format("第{0}行物料({1})部分添加失败!请检查！！！", i + 1, materialcode));
                        AllSuccess = 0;
                    }
                }
                else
                {
                    dgvItmbDetail.Rows[i].HeaderCell.Style.ForeColor = Color.Red;
                    dgvItmbDetail.Rows[i].Cells[0].Style.BackColor = Color.Red;
                    MessageBox.Show("第" + (i + 1) + "条记录的ITMB添加失败！");

                    AllSuccess = 0;
                }
            }
            dgvItmb.Rows[Convert.ToInt32(ITMBhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
            if (AllSuccess == 1)
            {
                MessageBox.Show("添加成功");
            }
            else
            {
                MessageBox.Show("部分添加失败，请检查！");
            }
        }
        private bool AddItemProductLineAndInventoryAccount(int rowIndex)//添加ITMC产品线及账号
        {
            ITMC03 myItmc = new ITMC03();
            string materialcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            string fenlei = materialcode.Substring(0, 1).ToUpper();
            if (fenlei == "M")
            {
                ADDGLAV(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), "121100");
            }
            if (fenlei == "A")
            {
                ADDGLAV(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), "123100");
            }
            if (fenlei == "P")
            {
                ADDGLAV(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), "122100");
            }
            if (fenlei == "F")
            {
                ADDGLAV(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), "124300");
                myItmc.SalesAccount.Value = dgvItmbDetail["销售账号", rowIndex].Value.ToString().Trim();
                myItmc.CostOfGoodsSoldAccount.Value = dgvItmbDetail["成本账号", rowIndex].Value.ToString().Trim();
            }
            if (fenlei == "S")
            {
                ADDGLAV(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), "124100");
                myItmc.SalesAccount.Value = dgvItmbDetail["销售账号", rowIndex].Value.ToString().Trim();
                myItmc.CostOfGoodsSoldAccount.Value = dgvItmbDetail["成本账号", rowIndex].Value.ToString().Trim();
            }
            myItmc.ItemNumber.Value = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            myItmc.ProductLine.Value = dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim();
            myItmc.InventoryAccount.Value = dgvItmbDetail["库存账号", rowIndex].Value.ToString().Trim();

            if (_fstiClient.ProcessId(myItmc, null))
            {
                ITMBResult.Items.Add("ITMC产品线和库存账号添加成功:");
                ITMBResult.Items.Add(_fstiClient.CDFResponse);
                return true;
            }
            else
            {
                if (fenlei == "F" || fenlei == "S")
                {
                    if ("15" + materialcode == dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim().Replace("-", "") || "150" + materialcode == dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim().Replace("-", ""))
                    {
                        ADDGLOS(dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), dgvItmbDetail["物料描述", rowIndex].Value.ToString().Trim(), ITMBResult);
                        AddGlAVALL(rowIndex, dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim(), fenlei);
                        ITMC03 myItmc03 = new ITMC03();
                        myItmc03.ItemNumber.Value = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
                        myItmc03.ProductLine.Value = dgvItmbDetail["产品线", rowIndex].Value.ToString().Trim();
                        myItmc03.InventoryAccount.Value = dgvItmbDetail["库存账号", rowIndex].Value.ToString().Trim();
                        myItmc03.SalesAccount.Value = dgvItmbDetail["销售账号", rowIndex].Value.ToString().Trim();
                        myItmc03.CostOfGoodsSoldAccount.Value = dgvItmbDetail["成本账号", rowIndex].Value.ToString().Trim();
                        if (_fstiClient.ProcessId(myItmc03, null))
                        {
                            ITMBResult.Items.Add("ITMC产品线和库存账号添加成功:");
                            ITMBResult.Items.Add(_fstiClient.CDFResponse);
                            return true;
                        }
                        else
                        {
                            MessageBox.Show("产品线和库存账号添加失败！");
                            FSTIError itemError = _fstiClient.TransactionError;
                            DumpErrorObject(myItmc03, itemError, ITMBResult);
                        }
                    }
                    else
                    {
                        MessageBox.Show("FS类需要新增的产品线与物料编码不匹配，请检查！");
                        FSTIError itemError = _fstiClient.TransactionError;
                        DumpErrorObject(myItmc, itemError, ITMBResult);
                    }
                }
                else
                {
                    MessageBox.Show("增加产品线和库存账号失败！");
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myItmc, itemError, ITMBResult);
                }

            }
            return false;
        }
        private bool AddItemCostData(int rowIndex)//添加ITMC成本
        {
            string materialcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            string fenlei = materialcode.Substring(0, 1).ToUpper();

            ITMC00 myItmc = new ITMC00();
            myItmc.ItemNumber.Value = materialcode;
            myItmc.CostType.Value = "0";
            myItmc.CostCode.Value = "1";
            if (fenlei == "M" || fenlei == "P" || fenlei == "A")
            {
                string cailiaofei = Math.Round(Convert.ToDouble(dgvItmbDetail["材料费", rowIndex].Value.ToString().Trim()), 9).ToString();
                myItmc.AtThisLevelMaterialCost.Value = cailiaofei;
                myItmc.RolledMaterialCost.Value = cailiaofei;
            }
            if (fenlei == "F" || fenlei == "S")
            {
                myItmc.AtThisLevelMaterialCost.Value = "0";

            }
            myItmc.AtThisLevelLaborCost.Value = "0";
            myItmc.AtThisLevelVariableOverheadCost.Value = "0";
            myItmc.AtThisLevelFixedOverheadCost.Value = "0";
            if (_fstiClient.ProcessId(myItmc, null))
            {
                ITMBResult.Items.Add("ITMC成本添加成功!");
                ITMBResult.Items.Add(_fstiClient.CDFResponse);

                return true;
            }
            else
            {
                MessageBox.Show("ITMC成本添加失败！");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmc, itemError, ITMBResult);
            }

            return false;
        }
        private bool AddItemMasterFilePlanDetail(int rowIndex)//添加ITMB计划明细
        {
            ITMB03 myItmb = new ITMB03();
            string materialcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            myItmb.ItemNumber.Value = materialcode;
            string fenlei = materialcode.Substring(0, 1).ToUpper();

            if (fenlei == "F" || fenlei == "S")
            {
                myItmb.Planner.Value = dgvItmbDetail["计划采购", rowIndex].Value.ToString().Trim();
                //myItmb.Buyer.Value = "BYR";
            }
            else if (fenlei == "P" || fenlei == "M" || fenlei == "A")
            {
                myItmb.Buyer.Value = dgvItmbDetail["计划采购", rowIndex].Value.ToString().Trim();
                //myItmb.Planner.Value = "PLR";
            }
            //myItmb.PlanningPolicy.Value = "N";
            myItmb.RunLeadTimeDays.Value = dgvItmbDetail["运行", rowIndex].Value.ToString();
            myItmb.FixedLeadTimeDays.Value = dgvItmbDetail["FIX", rowIndex].Value.ToString();
            myItmb.InspectionLeadTimeDays.Value = dgvItmbDetail["检验", rowIndex].Value.ToString();
            myItmb.PreferredStockroom.Value = dgvItmbDetail["优先库", rowIndex].Value.ToString();
            myItmb.PreferredBin.Value = dgvItmbDetail["位", rowIndex].Value.ToString();
            string dgvItemUM = dgvItmbDetail["单位", rowIndex].Value.ToString().Trim().ToUpper();
            if (dgvItemUM == "BI")
            { myItmb.DecimalPrecision.Value = "3"; }
            if (dgvItemUM == "KG"|| dgvItemUM == "L" || dgvItemUM == "M" || dgvItemUM == "C" || dgvItemUM == "S" || dgvItemUM == "T")
            { myItmb.DecimalPrecision.Value = "2"; }
            if (dgvItemUM == "WA"|| dgvItemUM == "WG")
            { myItmb.DecimalPrecision.Value = "4"; }
            myItmb.LotSizeMinimum.Value = dgvItmbDetail["最小批量订货", rowIndex].Value.ToString();
            myItmb.LotSizeMultiplier.Value = dgvItmbDetail["批量订货倍数", rowIndex].Value.ToString();
            if (fenlei != "A")
            {
                myItmb.ForecastConsumptionCode.Value = dgvItmbDetail["预测码", rowIndex].Value.ToString();
                myItmb.ForecastPeriod.Value = dgvItmbDetail["预测阶段", rowIndex].Value.ToString().Substring(0, 1);
            }
            if (fenlei == "F" || fenlei == "S")
            {
                myItmb.LotSizeQuantity.Value = dgvItmbDetail["批量订货数目", rowIndex].Value.ToString();
                myItmb.ATPCode.Value = "Y";
                myItmb.GatewayWorkCenter.Value = dgvItmbDetail["起始工作中心", rowIndex].Value.ToString();
            }
            else
            {
                myItmb.LotSizeDays.Value = dgvItmbDetail["批量订货天数", rowIndex].Value.ToString();
            }
            if (_fstiClient.ProcessId(myItmb, null))
            {
                ITMBResult.Items.Add("ITMB计划明细添加成功:");
                ITMBResult.Items.Add(_fstiClient.CDFResponse);
                ITMBResult.Items.Add("*************************************");
                return true;
            }
            else
            {
                ITMBResult.Items.Add("ITMB计划明细添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmb, itemError, ITMBResult);

            }
            return false;
        }
        private bool AddItemLotTraceAndSerializeDetail(int rowIndex)//添加ITMB批号明细
        {

            string materialcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            string fenlei = materialcode.Substring(0, 1).ToUpper();

            if (fenlei == "F" || fenlei == "S" || fenlei == "M")
            {
                ITMB07 itmb07 = new ITMB07();
                itmb07.ItemNumber.Value = materialcode;
                itmb07.LotNumberAssignmentPolicy.Value = "C";
                itmb07.LotDefaultPolicy.Value = "N";
                itmb07.IsLotTraceItemFIFO.Value = "Y";
                itmb07.BackflushPolicy.Value = "N";
                //itmb07.LastUsedLotCounter.Value = "000000000";
                //if (fenlei == "M")
                //{
                    itmb07.LotNumberMask.Value = "XXXXXXXXXXXXXXXXXXXX";
                //}
                //else
                //{
                //    itmb07.LotNumberMask.Value = "YYMMDDNN";
                //}
                if (_fstiClient.ProcessId(itmb07, null))
                {
                    ITMBResult.Items.Add("ITMB批号明细添加成功:");
                    ITMBResult.Items.Add(_fstiClient.CDFResponse);
                    ITMBResult.Items.Add("*************************************");

                }
                else
                {
                    ITMBResult.Items.Add("ITMB批号明细添加失败:");
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(itmb07, itemError, ITMBResult);
                    return false;
                }
            }
            else if (fenlei == "P")
            {
                ITMB07 itmb07 = new ITMB07();
                itmb07.ItemNumber.Value = materialcode;
                //itmb07.LastUsedLotCounter.Value = "000000000";
                itmb07.LotNumberMask.Value = "XXXXXXXXXXXXXXXXXXXX";

                if (_fstiClient.ProcessId(itmb07, null))
                {
                    ITMBResult.Items.Add("P类ITMB批号明细添加成功:");
                    ITMBResult.Items.Add(_fstiClient.CDFResponse);
                    ITMBResult.Items.Add("*************************************");

                }
                else
                {
                    ITMBResult.Items.Add("P类ITMB批号明细添加失败:");
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(itmb07, itemError, ITMBResult);
                    return false;
                }
            }
            else if (fenlei == "A")
                ITMBResult.Items.Add("A类物料ITMB批号明细无需添加！");
            return true;
        }
        private bool AddItemMasterFileDetail(int rowIndex)//增加ITMB物料明细
        {
            ITMB02 myItmb = new ITMB02();
            string itemcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            string fenlei = itemcode.Substring(0, 1).ToUpper();
            myItmb.ItemNumber.Value = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
            if (fenlei == "F" || fenlei == "S")
            {

                myItmb.ItemReference4.Value = "0";
            }
            myItmb.ItemReference3.Value = dgvItmbDetail["库管员代码", rowIndex].Value.ToString().Trim();
            if (ITMBfaqibumen.Text == "原料事业部综合办公室" || ITMBfaqibumen.Text == "综合办公室" || ITMBfaqibumen.Text == "固体制剂及水针事业部综合办公室")
            {
                if (ITMBfaqibumen.Text == "综合办公室")
                {
                    myItmb.ItemReference2.Value = "A";
                }
                if (ITMBfaqibumen.Text == "固体制剂及水针事业部综合办公室")
                {
                    myItmb.ItemReference2.Value = "B";
                }
                if (ITMBfaqibumen.Text == "原料事业部综合办公室")
                {
                    myItmb.ItemReference2.Value = "C";
                }
            }
            else if (ITMBjieshoudanwei.Text == "粉针事业部")//综合办公室
            {
                myItmb.ItemReference2.Value = "A";
            }
            else if (ITMBjieshoudanwei.Text == "固体制剂水针事业部")//固体制剂及水针事业部综合办公室
            {
                myItmb.ItemReference2.Value = "B";
            }
            else if (ITMBjieshoudanwei.Text == "原料事业部")//原料事业部综合办公室
            {
                myItmb.ItemReference2.Value = "C";
            }
            else
            { myItmb.ItemReference2.Value = "D"; }
            if (_fstiClient.ProcessId(myItmb, null))
            {
                ITMBResult.Items.Add("ITMB物料明细添加成功:");
                ITMBResult.Items.Add(_fstiClient.CDFResponse);
                ITMBResult.Items.Add("*************************************");
                return true;
            }
            else
            {
                ITMBResult.Items.Add("ITMB物料明细添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmb, itemError, ITMBResult);
            }
            return false;
        }
        private bool AddItemMasterFile(int rowIndex)//添加ITMB物料主文件的基本信息
        {
            try
            {
                string materialcode = dgvItmbDetail["物料代码", rowIndex].Value.ToString().Trim();
                string fenlei = materialcode.Substring(0, 1).ToUpper();
                if (fenlei == "M" || fenlei == "P")
                {
                    ITMB00 myItmb = new ITMB00();
                    myItmb.ItemNumber.Value = materialcode;
                    myItmb.ItemDescription.Value = dgvItmbDetail["物料描述", rowIndex].Value.ToString().Trim().Replace("’", "'");
                    string materialtype = dgvItmbDetail["单位", rowIndex].Value.ToString().Trim();
                    if (materialtype == "")
                    {
                        MessageBox.Show("物料单位为空，ITMB无法添加！");
                        return false;
                    }
                    else if (materialtype.Length > 2)
                    {
                        MessageBox.Show("物料单位长度不正确，ITMB无法添加！");
                        return false;
                    }
                    else
                    {
                        myItmb.ItemUM.Value = materialtype;
                    }
                    myItmb.ItemStatus.Value = "A";
                    myItmb.MakeBuyCode.Value = "B";      //制购B      
                    if (_fstiClient.ProcessId(myItmb, null))
                    {
                        ITMBResult.Items.Add("ITMB添加成功:");
                        ITMBResult.Items.Add(_fstiClient.CDFResponse);
                        return true;
                    }
                    else
                    {
                        FSTIError itemError = _fstiClient.TransactionError;
                        DumpErrorObject(myItmb, itemError, ITMBResult);
                    }
                }
                if (fenlei == "F" || fenlei == "S")
                {
                    ITMB00 myItmb = new ITMB00();
                    myItmb.ItemNumber.Value = materialcode;
                    myItmb.ItemDescription.Value = dgvItmbDetail["物料描述", rowIndex].Value.ToString().Trim();
                    string materialtype = dgvItmbDetail["单位", rowIndex].Value.ToString().Trim();
                    if (materialtype == "")
                    {
                        MessageBox.Show("物料单位为空，ITMB无法添加！");
                        return false;
                    }
                    else if (materialtype.Length > 2)
                    {
                        MessageBox.Show("物料单位长度不正确，ITMB无法添加！");
                        return false;
                    }
                    else
                    {
                        myItmb.ItemUM.Value = materialtype;
                    }
                    myItmb.ItemStatus.Value = "O";//状态O
                    myItmb.OrderPolicy.Value = "4";//订货4，添加完明细后改成3
                    if (_fstiClient.ProcessId(myItmb, null))
                    {
                        ITMBResult.Items.Add("ITMB添加成功:");
                        ITMBResult.Items.Add(_fstiClient.CDFResponse);
                        return true;
                    }
                    else
                    {
                        FSTIError itemError = _fstiClient.TransactionError;
                        DumpErrorObject(myItmb, itemError, ITMBResult);
                    }
                }
                if (fenlei == "A")
                {
                    ITMB00 myItmb = new ITMB00();
                    myItmb.ItemNumber.Value = materialcode;
                    myItmb.ItemDescription.Value = dgvItmbDetail["物料描述", rowIndex].Value.ToString().Trim();
                    string materialtype = dgvItmbDetail["单位", rowIndex].Value.ToString().Trim();
                    if (materialtype == "")
                    {
                        MessageBox.Show("物料单位为空，ITMB无法添加！");
                        return false;
                    }
                    else if (materialtype.Length > 2)
                    {
                        MessageBox.Show("物料单位长度不正确，ITMB无法添加！");
                        return false;
                    }
                    else
                    {
                        myItmb.ItemUM.Value = materialtype;
                    }
                    myItmb.MakeBuyCode.Value = "B";      //制购B  
                    myItmb.IsLotTraced.Value = "N";      //批号N
                    myItmb.IsInspectionRequired.Value = "N";//要求检验N
                    myItmb.ItemStatus.Value = "A";
                    if (_fstiClient.ProcessId(myItmb, null))
                    {
                        ITMBResult.Items.Add("ITMB添加成功:");
                        ITMBResult.Items.Add(_fstiClient.CDFResponse);
                        return true;
                    }
                    else
                    {
                        FSTIError itemError = _fstiClient.TransactionError;
                        DumpErrorObject(myItmb, itemError, ITMBResult);
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("增加ITMB主文件失败：" + ex.Message);
                return false;

            }
        }
        private void UpdateITMB_Click(object sender, EventArgs e)//修改物料成本按钮
        {
            MessageBox.Show("请打开流程，手动修改，谢谢！");
        }
        private bool AddGlAVALL(int rowIndex, string productline, string stritemnumber)//添加所有GLAV账号及CNFA
        {
            //库存账号添加
            if (stritemnumber == "S")
            {
                ADDGLAV(productline, "124100");//S类自制半成品挂接科目为124100
            }
            else if (stritemnumber == "F")
            {
                ADDGLAV(productline, "124300");//F类成品挂接科目为124300
            }

            //CNFA屏幕下，所有相关的科目下添加组织机构号
            ADDGLAV(productline, "410110");
            ADDGLAV(productline, "125100");
            ADDGLAV(productline, "999999");
            ADDGLAV(productline, "510100");
            ADDGLAV(productline, "410210");
            ADDGLAV(productline, "410220");
            ADDGLAV(productline, "123700");
            ADDGLAV(productline, "123800");
            ADDGLAV(productline, "540100");
            ADDGLAV(productline, "123600");
            ADDGLAV(productline, "123500");
            ADDGLAV(productline, "123400");
            ADDGLAV(productline, "123200");
            ADDGLAV(productline, "123900");
            ADDGLAV(productline, "410580");


            CNFA00 cnfa00 = new CNFA00();//添加产品线
            cnfa00.ProductLine.Value = productline;
            cnfa00.ProductLineDescription.Value = dgvItmbDetail["物料描述", rowIndex].Value.ToString().Trim();
            cnfa00.ProductLineStatus.Value = "A";
            _fstiClient.ProcessId(cnfa00, null);
            CNFA03 cnfa = new CNFA03();//添加产品线明细
            cnfa.ProductLine.Value = productline;
            cnfa.InternalWIPAccount.Value = productline + "-410110";
            cnfa.StandardCostVarianceAccount.Value = productline + "-123800";
            cnfa.ExternalWIPAccount.Value = productline + "-125100";
            cnfa.OverheadVarianceAccount.Value = productline + "-123600";
            cnfa.CustomProductWIPAccount.Value = productline + "-999999";
            cnfa.LaborVarianceAccount.Value = productline + "-123500";
            cnfa.CustomProductVarianceAccount.Value = productline + "-999999";
            cnfa.MaterialUsageVarianceAccount.Value = productline + "-123400";
            cnfa.SalesAccount.Value = productline + "-510100";
            cnfa.MiscellaneousOrderVarianceAccount.Value = productline + "-123700";//
            cnfa.CostOfGoodsSoldAccount.Value = productline + "-540100";
            cnfa.PPVAccount.Value = productline + "-123200";
            cnfa.ResourceDirectLaborAccount.Value = productline + "-410210";
            cnfa.POIssueVarianceAccount.Value = productline + "-123900";
            cnfa.ResourceVariableOverheadAccount.Value = productline + "-410220";
            cnfa.ManufacturingExpenseAccount.Value = productline + "-410580";
            cnfa.ResourceFixedOverheadAccount.Value = productline + "-410220";
            cnfa.AppliedLaborAccount.Value = productline + "-410210";
            cnfa.MOCostChangeAccount.Value = productline + "-123700";
            cnfa.AppliedOverheadAccount.Value = productline + "-410220";
            cnfa.ItemCostChangeAccount.Value = productline + "-123800";
            cnfa.AACPOVarianceAccount.Value = productline + "-123200";
            cnfa.POCostChangeAccount.Value = productline + "-123700";

            if (_fstiClient.ProcessId(cnfa, null))
            {
                ITMBResult.Items.Add("CNFA主明细添加成功:");
                ITMBResult.Items.Add("");
                ITMBResult.Items.Add(_fstiClient.CDFResponse);
                ITMBResult.Items.Add("*************************************");
                return true;
            }
            else
            {
                ITMBResult.Items.Add("CNFA主明细添加失败！");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(cnfa, itemError, ITMBResult);
            }
            return false;
        }
        private void dgvItmb_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//物料成本流程列表的cell双击事件
        {
            ITMBResult.Items.Clear();

            int rowindex = e.RowIndex;
            if (rowindex != -1)
            {
                ItemName.Text = "";
                ItemNamedgv.DataSource = null;
                if (dgvItmb.Rows[rowindex].DefaultCellStyle.ForeColor != Color.Red)
                {
                    dgvItmb.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                    for (int a = 0; a < dgvItmb.Rows.Count; a++)
                    {
                        if (a != rowindex && dgvItmb.Rows[a].DefaultCellStyle.ForeColor != Color.Red)
                            dgvItmb.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
                ITMBjieshoudanwei.Text = dgvItmb.Rows[rowindex].Cells["接收单位"].Value.ToString().Trim();
                ITMBliushuihao.Text = dgvItmb.Rows[rowindex].Cells["流水号"].Value.ToString().Trim();
                ITMBfaqibumen.Text = dgvItmb.Rows[rowindex].Cells["发起部门"].Value.ToString().Trim();
                ITMBhanghao.Text = (rowindex + 1).ToString();
                string ParentGuid = dgvItmb.Rows[rowindex].Cells["ParentGuid"].Value.ToString();

                dgvItmbDetail.DataSource = toITMBITMC(SqlHelper.ExecuteDataTable("SELECT ltrim(rtrim(WLBM1))  as 物料代码,WLMS as 物料描述,DW as 单位,ltrim(rtrim(KGYDM)) as 库管员代码,upper(ltrim(rtrim(JHY))) as 计划采购,YX as 运行,FIX,JY as 检验,PLDHTS as 批量订货天数,ZXPLDH as 最小批量订货,PLDHBS as 批量订货倍数,PLDHSM as 批量订货数目,QSGZZX as 起始工作中心,YXK as 优先库,W as 位," +
                    "CLF as 材料费,HJ as 合计,CPX as 产品线,KCZH as 库存账号,XSZH as 销售账号,CBZH as 成本账号,YCM as 预测码,YCJD as 预测阶段 FROM YW_ZJWLCB_EX where ParentGuid = '" + ParentGuid + "'"));
                for (int i = 0; i < this.dgvItmbDetail.Columns.Count; i++)
                {
                    this.dgvItmbDetail.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    this.dgvItmbDetail.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
                if (dgvItmb.Rows[rowindex].Cells["申请方式"].Value.ToString() == "增加")
                { AddITMB.Enabled = true; UpdateITMB.Enabled = false; }
                if (dgvItmb.Rows[rowindex].Cells["申请方式"].Value.ToString() == "修改")
                { UpdateITMB.Enabled = true; AddITMB.Enabled = false; }
                #region 若物料名称超出四班字符数上限，弹出提示

                for (int i = 0; i < this.dgvItmbDetail.Rows.Count; i++)
                {
                    String Item = dgvItmbDetail["物料描述", i].Value.ToString().Trim();
                    String ItemNum = dgvItmbDetail["物料代码", i].Value.ToString().Trim();
                    String fenlei = "1";
                    try
                    {
                        fenlei = ItemNum.Substring(0, 1);
                    }
                    catch
                    {


                    }
                    if (Item == "" || dgvItmbDetail["单位", i].Value.ToString().Trim() == "" || dgvItmbDetail["计划采购", i].Value.ToString().Trim() == "" || dgvItmbDetail["运行", i].Value.ToString().Trim() == "" || dgvItmbDetail["FIX", i].Value.ToString().Trim() == "" || dgvItmbDetail["批量订货天数", i].Value.ToString().Trim() == "" || dgvItmbDetail["最小批量订货", i].Value.ToString().Trim() == "")
                    {
                        //dgvItmbDetail.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["单位"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["计划采购"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["运行"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["FIX"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["批量订货天数"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["最小批量订货"].Style.BackColor = Color.Red;

                    }
                    if (ItemNum == "" || dgvItmbDetail["库管员代码", i].Value.ToString().Trim() == "" || dgvItmbDetail["优先库", i].Value.ToString().Trim() == "" || dgvItmbDetail["位", i].Value.ToString().Trim() == "" || dgvItmbDetail["产品线", i].Value.ToString().Trim() == "" || dgvItmbDetail["库存账号", i].Value.ToString().Trim() == "")
                    {
                        //dgvItmbDetail.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["物料代码"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["库管员代码"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["优先库"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["位"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["产品线"].Style.BackColor = Color.Red;
                        dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;

                    }
                    if (dgvItmb.Rows[rowindex].Cells["制购类型"].Value.ToString().Trim().Substring(0, 1) == "M")
                    {
                        if (dgvItmbDetail["起始工作中心", i].Value.ToString().Trim() == "" || dgvItmbDetail["销售账号", i].Value.ToString().Trim() == "" || dgvItmbDetail["成本账号", i].Value.ToString().Trim() == "")
                        {
                            dgvItmbDetail.Rows[i].Cells["起始工作中心"].Style.BackColor = Color.Red;
                            dgvItmbDetail.Rows[i].Cells["销售账号"].Style.BackColor = Color.Red;
                            dgvItmbDetail.Rows[i].Cells["成本账号"].Style.BackColor = Color.Red;
                        }
                        if (dgvItmbDetail["批量订货数目", i].Value.ToString().Trim() == "")
                        {
                            MessageBox.Show("第" + (i + 1) + "行[批量订货数目]有误，请检查！");
                            dgvItmbDetail.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            try
                            {
                                if (Convert.ToDecimal(dgvItmbDetail["批量订货数目", i].Value.ToString().Trim()) <= 1)
                                {
                                    MessageBox.Show("第" + (i + 1) + "行[批量订货数目]<= 1，请检查！");
                                    dgvItmbDetail.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("第" + (i + 1) + "行[批量订货数目]有误，请检查！");
                                dgvItmbDetail.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                            }
                        }
                    }
                    else
                    {
                        if (dgvItmbDetail["材料费", i].Value == null || dgvItmbDetail["材料费", i].Value.ToString().Trim() == "")
                        {
                            dgvItmbDetail.Rows[i].Cells["材料费"].Style.BackColor = Color.Red;
                        }
                    }
                    if (fenlei == "M")
                    {
                        if (dgvItmbDetail["产品线", i].Value.ToString().Trim() + "-121100" != dgvItmbDetail["库存账号", i].Value.ToString().Trim())
                        {
                            MessageBox.Show("第" + (i + 1) + "行:M类物料库存账号不是121100");
                            dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        }
                    }
                    if (fenlei == "A")
                    {
                        if (dgvItmbDetail["产品线", i].Value.ToString().Trim() + "-123100" != dgvItmbDetail["库存账号", i].Value.ToString().Trim())
                        {
                            MessageBox.Show("第" + (i + 1) + "行:A类物料库存账号不是123100");
                            dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        }
                    }
                    if (fenlei == "P")
                    {
                        if (dgvItmbDetail["产品线", i].Value.ToString().Trim() + "-122100" != dgvItmbDetail["库存账号", i].Value.ToString().Trim())
                        {
                            MessageBox.Show("第" + (i + 1) + "行:P类物料库存账号不是122100");
                            dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        }
                    }
                    if (fenlei == "F")
                    {
                        if (dgvItmbDetail["产品线", i].Value.ToString().Trim() + "-124300" != dgvItmbDetail["库存账号", i].Value.ToString().Trim())
                        {
                            MessageBox.Show("第" + (i + 1) + "行:F类物料库存账号不是124300");
                            dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        }
                    }
                    if (fenlei == "S")
                    {
                        if (dgvItmbDetail["产品线", i].Value.ToString().Trim() + "-124100" != dgvItmbDetail["库存账号", i].Value.ToString().Trim())
                        {
                            MessageBox.Show("第" + (i + 1) + "行:S类物料库存账号不是124100");
                            dgvItmbDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        }
                    }
                    if (dgvItmbDetail["位", i].Value.ToString().Trim().Length % 2 == 1)
                    {
                        dgvItmbDetail.Rows[i].Cells["位"].Style.BackColor = Color.Red;
                        MessageBox.Show("第" + (i + 1) + "行:库位的位数不是偶数，请录入四班后修改(前面补0)");
                    }
                    if (StrLength(Item) > 70)
                    {
                        MessageBox.Show("第" + (i + 1) + "行[物料描述]超出字符数限制");
                        dgvItmbDetail.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                    }
                    if (Item.Contains(' ') || Item.Contains('（') || Item.Contains('）') || Item.ToUpper() != Item)
                    {
                        MessageBox.Show("第" + (i + 1) + "行[物料描述]不符合ERP物料描述规范");
                        dgvItmbDetail.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                    }
                }
                #endregion
                for (int y = 0; y < this.dgvItmbDetail.Rows.Count; y++)
                {
                    for (int x = 0; x < this.dgvItmbDetail.Columns.Count; x++)
                    {
                        if (dgvItmbDetail.Rows[y].Cells[x].Style.BackColor == Color.Red)
                        {
                            MessageBox.Show("信息有误已红色标示，请检查！"); return;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 获得字符串的区分中英文的字符长度
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private int StrLength(string str)// 获得字符串的区分中英文的字符长度
        {
            if (string.IsNullOrEmpty(str)) return 0;
            int len = 0;
            byte[] b;

            for (int i = 0; i < str.Length; i++)
            {
                b = Encoding.Default.GetBytes(str.Substring(i, 1));
                len += b.Length;

            }

            return len;
        }
        private IEnumerable<ITMBITMC> toITMBITMC(DataTable dt)//DataRow转 流程明细信息
        {
            List<ITMBITMC> ITMBitmc = new List<ITMBITMC>();
            foreach (DataRow dr in dt.Rows)
            {
                ITMBITMC iTMBITMC = new ITMBITMC();
                iTMBITMC.物料代码 = dr["物料代码"].ToString().Trim().ToUpper();
                iTMBITMC.物料描述 = dr["物料描述"].ToString().Trim().ToUpper();
                iTMBITMC.单位 = dr["单位"].ToString().Trim().ToUpper();
                iTMBITMC.库管员代码 = dr["库管员代码"].ToString().Trim().ToUpper();
                iTMBITMC.计划采购 = dr["计划采购"].ToString().Trim().ToUpper();
                iTMBITMC.运行 = dr["运行"].ToString().Trim().ToUpper();
                iTMBITMC.FIX = dr["FIX"].ToString().Trim().ToUpper();
                iTMBITMC.检验 = dr["检验"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货天数 = dr["批量订货天数"].ToString().Trim().ToUpper();
                iTMBITMC.最小批量订货 = dr["最小批量订货"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货倍数 = dr["批量订货倍数"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货数目 = dr["批量订货数目"].ToString().Trim().ToUpper();
                iTMBITMC.起始工作中心 = dr["起始工作中心"].ToString().Trim().ToUpper();
                iTMBITMC.优先库 = dr["优先库"].ToString().Trim().ToUpper();
                iTMBITMC.位 = dr["位"].ToString().Trim().ToUpper();
                if (dr["材料费"].ToString().Trim() != "")
                    iTMBITMC.材料费 = Math.Round(Convert.ToDouble(dr["材料费"].ToString().Trim()), 9).ToString();
                if (dr["合计"].ToString().Trim() != "")
                    iTMBITMC.合计 = Math.Round(Convert.ToDouble(dr["合计"].ToString().Trim()), 9).ToString();
                iTMBITMC.产品线 = dr["产品线"].ToString().Trim().ToUpper();
                iTMBITMC.库存账号 = dr["库存账号"].ToString().Trim().ToUpper();
                iTMBITMC.销售账号 = dr["销售账号"].ToString().Trim().ToUpper();
                iTMBITMC.成本账号 = dr["成本账号"].ToString().Trim().ToUpper();
                iTMBITMC.预测码 = dr["预测码"].ToString().Trim().ToUpper();
                iTMBITMC.预测阶段 = dr["预测阶段"].ToString().Trim().ToUpper();
                ITMBitmc.Add(iTMBITMC);

            }
            return ITMBitmc;
        }

        private void GetITMB_Click(object sender, EventArgs e)//获得物料成本流程
        {
            ItemName.Text = "";
            ItemNamedgv.DataSource = null;
            dgvItmbDetail.DataSource = null;
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and   PROCESSNAME='RY增加物料申请流程' and (STEPLABEL = '系统管理员维护')");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and   PROCESSNAME='RY增加物料申请流程' and (STEPLABEL = '系统管理员维护' or STEPLABEL = 'ERP管理员审核') and STARTTIME >'2019-5-10'");
            List<ITMBliucheng> list1 = new List<ITMBliucheng>();
            foreach (DataRow dr in Incidents.Rows)
            {
                ITMBliucheng Vendor1 = TolistITMB(SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_ZJWLCB] where REV_INCIDENT=" + dr[0]));
                list1.Add(Vendor1);

            }

            dgvItmb.DataSource = list1;
            for (int i = 0; i < this.dgvItmb.Columns.Count; i++)
            {
                this.dgvItmb.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvItmb.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            ITMBResult.Items.Clear();
        }
        private ITMBliucheng TolistITMB(DataTable dt)//dt转化为ITMBliucheng的方法
        {

            ITMBliucheng bom1 = new ITMBliucheng();
            bom1.申请人 = (string)(dt.Rows[0]["REV_CREATER_NAME"]);
            bom1.发起部门 = (string)(dt.Rows[0]["REV_CREATER_DPT"]);
            bom1.申请时间 = (DateTime)(dt.Rows[0]["REV_CREATER_DATE"]);
            bom1.流水号 = (Int32)(dt.Rows[0]["REV_INCIDENT"]);
            if ((string)(dt.Rows[0]["XGWL"]) == "")
                bom1.申请方式 = "增加";
            if ((string)(dt.Rows[0]["XGWL"]) == "修改物料")
                bom1.申请方式 = "修改";
            bom1.联系电话 = (string)(dt.Rows[0]["REV_CREATER_TEL"]);
            bom1.接收单位 = (string)(dt.Rows[0]["JSDW"]);
            bom1.摘要 = (string)(dt.Rows[0]["ZY"]);
            bom1.制购类型 = (string)(dt.Rows[0]["ZGLX"]);
            bom1.ParentGuid = (string)(dt.Rows[0]["REV_CID"]);

            if (dt.Rows.Count != 1)
            {
                MessageBox.Show(bom1.流水号 + "有0个或多个的流水号");
            }
            return bom1;

        }

        private void jihuoITMB_Click(object sender, EventArgs e)//激活增加修改ITMB按钮
        {
            SubmitITMB.Enabled = true;
            AddITMB.Enabled = true;
            UpdateITMB.Enabled = true;
        }

        private void Daily_Click(object sender, EventArgs e)//每日检查
        {
            xunikucun.DataSource = SqlHelper1.ExecuteDataTable(SqlHelper.FSDBSQL,"SELECT * FROM [dbo].[VirtualInventory]");//虚拟库存表
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                SqlCommand cmd = new SqlCommand("SELECT *  FROM (SELECT dbo.FS_Item.ItemNumber, CAST(LTRIM(str( " +
                                                "(dbo.FS_ItemData.OnHandQuantity + dbo.FS_ItemData.OnHoldQuantity + dbo.FS_ItemData.InShippingQuantity + dbo.FS_ItemData.InInspectionQuantity), 10, dbo.FS_Item.DecimalPrecision)) AS DECIMAL(20, 10)) TotalQty, " +
                                                "CAST(LTRIM(str((SELECT SUM(dbo.FS_ItemInventory.InventoryQuantity) FROM dbo.FS_ItemInventory WHERE dbo.FS_ItemData.ItemKey = dbo.FS_ItemInventory.ItemKey), 10, dbo.FS_Item.DecimalPrecision)) AS DECIMAL(20, 10)) LinesQty " +
                                                 "FROM dbo.FS_Item " +
                                                "INNER JOIN dbo.FS_ItemData ON dbo.FS_Item.ItemKey = dbo.FS_ItemData.ItemKey) as IISSCHECK Where IISSCHECK.TotalQty <> IISSCHECK.LinesQty or (IISSCHECK.LinesQty is NULL and IISSCHECK.TotalQty != 0)", conn);
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                SSIICKECK.DataSource = dt;                     //SSII屏幕显示总数量和实际数量不一致的物料
            }
            using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
            {
                //SqlCommand cmd = new SqlCommand("SELECT VendorID AS 供应商编码,VendorName AS 供应商名称,PayeeName1 AS 收款人名称1,PayeeName2 AS 收款人名称2,"+
                //                "VendorClass6 AS 供应商分类代码,VendorClass7 AS 供应商分类,BankName AS 银行名称,BankAccountNumber AS 银行账号,UnvoucheredAccount AS 无票应付,VoucheredAccount AS 有票应付"+
                //            "FROM _NoLock_FS_Vendor WHERE VendorStatus = 'A' AND( VendorName<>(PayeeName1 + PayeeName2) OR UnvoucheredAccount NOT LIKE '%212101' OR VoucheredAccount NOT LIKE '%212100'" +
                //             "OR BankName IS NULL OR BankName = '' OR BankAccountNumber IS NULL OR BankAccountNumber = '')", conn);
                OleDbCommand cmd = new OleDbCommand("SELECT VendorID AS 供应商编码, VendorName AS 供应商名称,PayeeName1 AS 收款人名称1,PayeeName2 AS 收款人名称2,VendorClass6 AS 供应商分类代码,VendorClass7 AS 供应商分类," +
                                "BankName AS 银行名称,BankAccountNumber AS 银行账号,UnvoucheredAccount AS 无票应付,VoucheredAccount AS 有票应付 FROM [dbo].[_NoLock_FS_Vendor] WHERE VendorStatus = 'A' AND ( VendorName <> (PayeeName1 + PayeeName2) OR UnvoucheredAccount NOT LIKE '%212101' " +
 "OR VoucheredAccount NOT LIKE '%212100' OR BankName IS NULL OR BankName = '' OR BankAccountNumber IS NULL OR BankAccountNumber = '')", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                VendorCheck.DataSource = dt;                     //供应商检查 检查银行账户为空、供应商名称与开票名称不一致的供应商
            }
            using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
            {
                //SqlCommand cmd = new SqlCommand("SELECT VendorID AS 供应商编码,VendorName AS 供应商名称,PayeeName1 AS 收款人名称1,PayeeName2 AS 收款人名称2,"+
                //                "VendorClass6 AS 供应商分类代码,VendorClass7 AS 供应商分类,BankName AS 银行名称,BankAccountNumber AS 银行账号,UnvoucheredAccount AS 无票应付,VoucheredAccount AS 有票应付"+
                //            "FROM _NoLock_FS_Vendor WHERE VendorStatus = 'A' AND( VendorName<>(PayeeName1 + PayeeName2) OR UnvoucheredAccount NOT LIKE '%212101' OR VoucheredAccount NOT LIKE '%212100'" +
                //             "OR BankName IS NULL OR BankName = '' OR BankAccountNumber IS NULL OR BankAccountNumber = '')", conn);
                OleDbCommand cmd = new OleDbCommand("SELECT  ItemNumber AS 物料编码,ItemDescription AS 物料描述,InventoryAccount as 库存,SalesAccount AS 销售,CostOfGoodsSoldAccount as COGS FROM FSDBMR.dbo._NoLock_FS_Item  WHERE ItemStatus='A' AND ((LEFT(ItemNumber,1)='M' AND RIGHT(InventoryAccount,6)<>'121100') OR (LEFT(ItemNumber,1)='A' AND RIGHT(InventoryAccount,6)<>'123100') OR (LEFT(ItemNumber,1)='P' AND RIGHT(InventoryAccount,6)<>'122100') OR ((LEFT(ItemNumber,1)='F' AND (RIGHT(InventoryAccount,6)<>'124300' OR  RIGHT(SalesAccount,6)<>'510100' OR  RIGHT(CostOfGoodsSoldAccount,6)<>'540100'))) OR ((LEFT(ItemNumber,1)='S' AND (RIGHT(InventoryAccount,6)<>'124100' OR  RIGHT(SalesAccount,6)<>'510100' OR  RIGHT(CostOfGoodsSoldAccount,6)<>'540100'))))", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                zhanghaoCheck.DataSource = dt;                     //物料账号检查
            }
            using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
            {
                //SqlCommand cmd = new SqlCommand("SELECT VendorID AS 供应商编码,VendorName AS 供应商名称,PayeeName1 AS 收款人名称1,PayeeName2 AS 收款人名称2,"+
                //                "VendorClass6 AS 供应商分类代码,VendorClass7 AS 供应商分类,BankName AS 银行名称,BankAccountNumber AS 银行账号,UnvoucheredAccount AS 无票应付,VoucheredAccount AS 有票应付"+
                //            "FROM _NoLock_FS_Vendor WHERE VendorStatus = 'A' AND( VendorName<>(PayeeName1 + PayeeName2) OR UnvoucheredAccount NOT LIKE '%212101' OR VoucheredAccount NOT LIKE '%212100'" +
                //             "OR BankName IS NULL OR BankName = '' OR BankAccountNumber IS NULL OR BankAccountNumber = '')", conn);
                OleDbCommand cmd = new OleDbCommand("SELECT A.ItemNumber AS 物料编码, A.ItemDescription AS 物料描述,B.[AtThisLevelMaterialCost] as 材料费,B.[AtThisLevelLaborCost] as 人工费,B.[AtThisLevelVariableOverheadCost]  as  可变间接费,B.[AtThisLevelFixedOverheadCost] as 固定间接费 FROM FSDBMR.dbo._NoLock_FS_Item  AS  A  LEFT JOIN FSDBMR.dbo._NoLock_FS_ItemCost  AS  B ON A.ItemKey = B.ItemKey  WHERE  A.ItemStatus = 'A' and B.CostType='0' and CostCode='1' AND B.[AtThisLevelMaterialCost]=0", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                chengbenCheck.DataSource = dt;                     //已启用物料成本检查
            }
            using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
            {
                //SqlCommand cmd = new SqlCommand("SELECT VendorID AS 供应商编码,VendorName AS 供应商名称,PayeeName1 AS 收款人名称1,PayeeName2 AS 收款人名称2,"+
                //                "VendorClass6 AS 供应商分类代码,VendorClass7 AS 供应商分类,BankName AS 银行名称,BankAccountNumber AS 银行账号,UnvoucheredAccount AS 无票应付,VoucheredAccount AS 有票应付"+
                //            "FROM _NoLock_FS_Vendor WHERE VendorStatus = 'A' AND( VendorName<>(PayeeName1 + PayeeName2) OR UnvoucheredAccount NOT LIKE '%212101' OR VoucheredAccount NOT LIKE '%212100'" +
                //             "OR BankName IS NULL OR BankName = '' OR BankAccountNumber IS NULL OR BankAccountNumber = '')", conn);
                OleDbCommand cmd = new OleDbCommand("SELECT ItemNumber AS 物料编码,ItemDescription AS 物料描述,ItemUM as 单位,LotSizeQuantity as 批量订货数目,LotSizeMinimum AS 最小批量订货 FROM FSDBMR.dbo._NoLock_FS_Item WHERE LotSizeQuantity<5 and (ItemNumber LIKE 'F%' or ItemNumber LIKE 'S%') AND ItemStatus='A' ", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                PLDHSMCheck.DataSource = dt;                     //FS类批量订货数目检查
            }
        }


        private void qiyong_Click(object sender, EventArgs e)//启用物料
        {
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            wuliaoqiyong.Items.Clear();
            ITMB01 myItmb = new ITMB01();
            string ItemCode = ITMB["物料", 0].Value.ToString().Trim().ToUpper();
            myItmb.ItemNumber.Value = ItemCode;
            myItmb.ItemStatus.Value = "A";      //状态A  
            if (_fstiClient.ProcessId(myItmb, null))
            {
                wuliaoqiyong.Items.Add(ItemCode + "物料启用成功");
                qiyong.Enabled = false;

                OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB);
                OleDbCommand cmd = new OleDbCommand("SELECT ItemNumber 物料,ItemDescription 描述,ItemUM 单位,ItemRevision 版,MakeBuyCode 制购,ItemType 物类,ItemStatus 状态,IsLotTraced 批号,IsSerialized 系号,OrderPolicy 订货,IsInspectionRequired 要求检验 FROM[dbo].[_NoLock_FS_Item] where ItemNumber = '" + ITMB["物料", 0].Value.ToString().Trim().ToUpper() + "'", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                ITMB.DataSource = dt;
                ITMB.Columns["状态"].DefaultCellStyle.ForeColor = Color.Blue;

                MessageBox.Show(ItemCode + ":启用成功");
            }
            else
            {

                MessageBox.Show(ItemCode + ":启用失败");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myItmb, itemError, wuliaoqiyong);
                wuliaoqiyong.Items.Add(ItemCode + "物料启用失败");
            }
        }

        private void SubmitBOM_Click(object sender, EventArgs e)
        {

            SubmitBOM.Enabled = false;
            MessageBox.Show("请在浏览器中提交！");
            //GetBOm_Click(sender,e);
        }

        private void SubmitVendor_Click(object sender, EventArgs e)
        {
            SubmitVendor.Enabled = false;
            MessageBox.Show("请在浏览器中提交！");
            //GetVendor_Click(sender,e);
        }

        private void SubmitITMB_Click(object sender, EventArgs e)
        {
            SubmitITMB.Enabled = false;
            MessageBox.Show("请在浏览器中提交！");
            //GetITMB_Click(sender, e);
        }

        private void daochuExcel_Click(object sender, EventArgs e)
        {
            #region  选择datagridview
            DataGridView dgv;
            if (comboBox1.Text == "xunikucun")
                dgv = xunikucun;
            else if (comboBox1.Text == "SSIICKECK")
                dgv = SSIICKECK;
            else if (comboBox1.Text == "VendorCheck")
                dgv = VendorCheck;
            else if (comboBox1.Text == "chengbenCheck")
                dgv = chengbenCheck;
            else if (comboBox1.Text == "zhanghaoCheck")
                dgv = zhanghaoCheck;
            else if (comboBox1.Text == "PLDHSMCheck")
                dgv = PLDHSMCheck;
            else
            {
                MessageBox.Show("请在下拉框中选择需要导出的表格");
                return;

            }
            #endregion
            string fileName = "11";
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xlsx";
            saveDialog.Filter = "Excel文件|*.xlsx";
            saveDialog.FileName = fileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0) return; //被点了取消
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 
                                                                                                                                  //写入标题   

            for (int i = 0; i < dgv.ColumnCount; i++)
            { worksheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText; }
            //写入数值
            for (int r = 0; r < dgv.Rows.Count; r++)
            {
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = dgv.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
            MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);
            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);  //fileSaved = true;                 
                }
                catch (Exception ex)
                {//fileSaved = false;                      
                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                }
            }
            xlApp.Quit();
            GC.Collect();//强行销毁     
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            picture picture = new picture();
            picture.Show();
        }

        private void textPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnLogon_Click(sender, e);
            }
        }

        private void ItemName_KeyDown(object sender, KeyEventArgs e)//相似物料查询
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(ItemName.Text.Trim()))
                {
                    #region
                    Encoding EncodingEN = Encoding.GetEncoding("ISO-8859-1");
                    Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                    string nametolading = "%" + EncodingEN.GetString(EncodingCH.GetBytes(ItemName.Text.Trim())) + "%";
                    string fenLei = ItemClass.Text == "全部" ? "%" : ItemClass.Text;
                    #endregion
                    using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                    {
                        OleDbCommand cmd = new OleDbCommand(" SELECT dbo._NoLock_FS_Item.ItemNumber AS 物料代码, dbo._NoLock_FS_Item.ItemDescription AS 物料描述, dbo._NoLock_FS_Item.ItemUM AS 物料单位, dbo._NoLock_FS_ItemCost.CostType AS 类别, dbo._NoLock_FS_ItemCost.CostCode AS 方法, dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计, dbo._NoLock_FS_Item.GatewayWorkCenter AS 工作中心 FROM dbo._NoLock_FS_Item INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey WHERE (dbo._NoLock_FS_ItemCost.CostType = '0') AND (dbo._NoLock_FS_Item.ItemDescription LIKE '" + nametolading + "') and (dbo._NoLock_FS_Item.ItemNumber like '" + fenLei + "%')", conn);
                        DataTable dt = new DataTable();
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dt);
                        ItemNamedgv.DataSource = dt;
                        ItemNamedgv.Columns[1].Width = 360;
                        ItemNamedgv.Columns[5].Width = 160;
                    }
                }
                else
                {
                    ItemNamedgv.DataSource = null;
                }
            }
        }

        private void dgvItmbDetail_CellClick(object sender, DataGridViewCellEventArgs e)//相似物料查询
        {
            int rowIndex = e.RowIndex;
            if (rowIndex > -1)
            {
                ItemName.Text = dgvItmbDetail.Rows[rowIndex].Cells[1].Value.ToString();
                if (!string.IsNullOrEmpty(ItemName.Text.Trim()))
                {
                    #region
                    Encoding EncodingEN = Encoding.GetEncoding("ISO-8859-1");
                    Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                    string nametolading = "%" + EncodingEN.GetString(EncodingCH.GetBytes(ItemName.Text.Trim())) + "%";
                    string fenLei = ItemClass.Text == "全部" ? "%" : ItemClass.Text;
                    #endregion
                    using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                    {
                        OleDbCommand cmd = new OleDbCommand(" SELECT dbo._NoLock_FS_Item.ItemNumber AS 物料代码, dbo._NoLock_FS_Item.ItemDescription AS 物料描述, dbo._NoLock_FS_Item.ItemUM AS 物料单位, dbo._NoLock_FS_ItemCost.CostType AS 类别, dbo._NoLock_FS_ItemCost.CostCode AS 方法, dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计, dbo._NoLock_FS_Item.GatewayWorkCenter AS 工作中心 FROM dbo._NoLock_FS_Item INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey WHERE (dbo._NoLock_FS_ItemCost.CostType = '0') AND (dbo._NoLock_FS_Item.ItemDescription LIKE '" + nametolading + "') and (dbo._NoLock_FS_Item.ItemNumber like '" + fenLei + "%')", conn);
                        DataTable dt = new DataTable();
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dt);
                        ItemNamedgv.DataSource = dt;
                        ItemNamedgv.Columns[1].Width = 360;
                        ItemNamedgv.Columns[5].Width = 160;
                    }
                }
                else
                {
                    ItemNamedgv.DataSource = null;
                }
            }
        }

        private void GetCustomerprocess_Click(object sender, EventArgs e)//获得客户流程
        {
            #region 客户信息groupBox7信息清空
            foreach (Control control in groupBox7.Controls)
            {
                if (!(control is Label))
                {
                    control.Text = null;
                }
            }
            #endregion
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and PROCESSNAME='RY开户审批流程' and TASKUSER='BPM/cuiqingjuan' and ENDTIME >'2019/12/13' and STEPLABEL='ERP管理员新增客户'");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY开户审批流程' and TASKUSER='BPM/cuiqingjuan' and STEPLABEL='ERP管理员新增客户'");
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY开户审批流程'  and STEPLABEL='ERP管理员新增客户'");//当前环节流程
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY开户审批流程' ");//所有未完成流程
            string cmdstr = @"SELECT REV_INCIDENT AS 流水号, REV_CREATER_NAME	AS 发起人, REV_CREATER_DPT	AS 发起部门, REV_CREATER_TEL  AS 联系电话, (CASE   WHEN XGKHMC <> '' THEN '客户信息修改'  ELSE '新开户'  END ) AS 类型, KHBM	AS 客户代码, YKHMC AS 原客户名称, KHMC	AS 客户名称, KHDZ	AS 客户地址,(CASE   WHEN KHLB = 'yy' THEN '医院'  ELSE (CASE   WHEN KHLB = 'yd' THEN '药店'  ELSE (CASE   WHEN KHLB = 'mz' THEN '门诊'  ELSE  (CASE   WHEN KHLB = 'gs' or  KHLB = 'yc' THEN '公司'  ELSE   '其他'   END )    END )     END )     END )	AS 公司类型, MS	AS 合并账户, ZKHB	AS 货币类型, YZBM	AS 邮编, DH	AS 电话, CZ	AS 传真,  KHYH	AS 开户银行, ZH	AS 银行账户, SH	AS 税号, ZJL	AS 总经理, YXJL	AS 销售经理, ZB	AS 主办, CWJL	AS 财务经理, YWY	AS 业务员, YWYH	AS 业务代码, KJ	AS 会计, KJDH	AS 会计电话,BZ as 备注,KHSZSF as 客户所在省份 FROM YW_KHSP where REV_INCIDENT=-12345";
            //string cmdstr = "SELECT * FROM YW_KHSP where REV_INCIDENT=" + Incidents.Rows[0][0];

            for (int i = 0; i < Incidents.Rows.Count; i++)
            {
                cmdstr += " or REV_INCIDENT=" + Incidents.Rows[i][0];
            }

            dgvCustomer.DataSource = SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, cmdstr);
            for (int i = 0; i < this.dgvCustomer.Columns.Count; i++)
            {
                this.dgvCustomer.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvCustomer.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void dgvCustomer_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//双击客户流程
        {
            listBoxCustomer.Items.Clear();
            int rowindex = e.RowIndex;
            #region 客户信息groupBox7信息清空
            foreach (Control control in groupBox7.Controls)
            {
                if (!(control is Label))
                {
                    control.Text = null;
                }
            }
            #endregion
            if (rowindex == -1) return;
            CbChildCustomer.Checked = false;
            if (dgvCustomer.Rows[rowindex].Cells["类型"].Value.ToString() == "客户信息修改")
            { MessageBox.Show("请手工修改客户信息"); return; }

            if (dgvCustomer.Rows[rowindex].DefaultCellStyle.ForeColor != Color.Red)
            {
                dgvCustomer.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                for (int a = 0; a < dgvCustomer.Rows.Count; a++)
                {
                    if (a != rowindex && dgvCustomer.Rows[a].DefaultCellStyle.ForeColor != Color.Red)
                        dgvCustomer.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                }
            }
            customercode.Text = dgvCustomer.Rows[rowindex].Cells["客户代码"].Value.ToString().Trim().ToUpper();
            customerliushuihao.Text = dgvCustomer.Rows[rowindex].Cells["流水号"].Value.ToString().Trim().ToUpper();
            customerhanghao.Text = (rowindex + 1).ToString();
            #region 
            if (dgvCustomer["类型", rowindex].Value.ToString() == "新开户")
            {

                tbCustomerCode.Text = dgvCustomer["客户代码", rowindex].Value.ToString().Trim();
                tbCustomerName.Text = dgvCustomer["客户名称", rowindex].Value.ToString().Trim();
                tbCustAddress.Text = dgvCustomer["客户地址", rowindex].Value.ToString().Trim();
                tbUniteAccount.Text = dgvCustomer["合并账户", rowindex].Value.ToString().Trim();
                tbPostcode.Text = dgvCustomer["邮编", rowindex].Value.ToString().Trim();
                tbProvince.Text = dgvCustomer["客户所在省份", rowindex].Value.ToString().Trim().Trim();


                if (string.IsNullOrWhiteSpace(dgvCustomer["总经理", rowindex].Value.ToString()))
                {
                    tbContactPerson.Text = dgvCustomer["财务经理", rowindex].Value.ToString().Trim();
                }
                else
                {
                    tbContactPerson.Text = dgvCustomer["总经理", rowindex].Value.ToString().Trim();
                }
                int al = dgvCustomer["公司类型", rowindex].Value.ToString().Trim().Length;
                if (al > 1)
                    cbIndustry.Text = dgvCustomer["公司类型", rowindex].Value.ToString().Trim().Substring(al - 2, 2);
                else
                    cbIndustry.Text = dgvCustomer["公司类型", rowindex].Value.ToString().Trim();

                if (dgvCustomer["货币类型", rowindex].Value.ToString().Trim() == "00000")
                {
                    tbCustomerCurrencyCode.Text = "00000";
                    cbMoney.Text = "本币";

                }
                else
                {
                    //美元（USD）、欧元(EURO)、瑞士法郎（CHF）  本币（00000）
                    tbCustomerCurrencyCode.Text = dgvCustomer["货币类型", rowindex].Value.ToString().Trim();
                    cbMoney.Text = "外币";
                    cbIndustry.Text = "外贸公司";
                }

                tbContactTelephone.Text = dgvCustomer["电话", rowindex].Value.ToString().Trim();
                tbContactFax.Text = dgvCustomer["传真", rowindex].Value.ToString().Trim();
                tbSalesmanName.Text = dgvCustomer["业务员", rowindex].Value.ToString().Trim();
                tbSalesmanCode.Text = dgvCustomer["业务代码", rowindex].Value.ToString().Trim();
                tbBankOfDeposit.Text = dgvCustomer["开户银行", rowindex].Value.ToString().Trim();
                tbBankAccount.Text = dgvCustomer["银行账户", rowindex].Value.ToString().Replace(" ", "").Trim();
                tbTaxCode.Text = dgvCustomer["税号", rowindex].Value.ToString().Trim();
                tbAccountantName.Text = dgvCustomer["会计", rowindex].Value.ToString().Trim();
                tbAccountantPhone.Text = dgvCustomer["会计电话", rowindex].Value.ToString().Trim();

                tBCustBZ.Text = dgvCustomer["备注", rowindex].Value.ToString().Trim();
                if (StrLength(tbProvince.Text) > 10)
                { MessageBox.Show("客户所在省份大于10个字符(5个汉字)，请编辑！"); }
            }
            #endregion
            if (dgvCustomer.Rows[rowindex].Cells["类型"].Value.ToString() == "新开户")
            { AddCustomer.Enabled = true; UpdateCustomer.Enabled = false; }
            if (dgvCustomer.Rows[rowindex].Cells["类型"].Value.ToString() == "客户信息修改")
            { UpdateCustomer.Enabled = true; AddCustomer.Enabled = false; }
            string Customercode1 = tbCustomerCode.Text.Trim();
            tbUniteAccount.Text = "1A" + Customercode1.Substring(6, 1) + Customercode1.Substring(0, 1) + "-" + Customercode1.Substring(1, 2) + "-" + Customercode1.Substring(3, 3);

        }

        private void AddCustomer_Click(object sender, EventArgs e)//增加客户button
        {

            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            if (tbCustomerCode.Text.ToString() == "" || tbCustomerName.Text.ToString() == "" || tbCustAddress.Text.ToString() == "" || tbSalesmanCode.Text.ToString() == "" || tbSalesmanName.Text.ToString() == "")
            {
                MessageBox.Show("客户信息不完整，无法添加！ ");
                return;
            }
            tbCustomerCurrencyCode.Text = tbCustomerCurrencyCode.Text.Trim();
            string CCode = tbCustomerCurrencyCode.Text.Trim();
            if (cbIndustry.Text == "" || cbMoney.Text == "")
            { MessageBox.Show("请检查行业类别|货币类型！"); return; }
            if ((cbMoney.Text == "本币" && tbCustomerCurrencyCode.Text != "00000") || (cbMoney.Text == "外币" && tbCustomerCurrencyCode.Text != "USD" && tbCustomerCurrencyCode.Text != "EURO" && tbCustomerCurrencyCode.Text != "CHF")||string.IsNullOrWhiteSpace(cbMoney.Text)||string.IsNullOrWhiteSpace(tbCustomerCurrencyCode.Text))
            { MessageBox.Show("请检查货币类型|货币代码是否对应！"); return; }
            if (StrLength(tbBankOfDeposit.Text.Trim()) > 30)
            {
                MessageBox.Show("开户银行超出30个字符，请调整！ ");
                return;
            }
            if (StrLength(tbBankAccount.Text.Trim()) > 30)
            {
                MessageBox.Show("开户银行账号超出30个字符，请调整！ ");
                return;
            }
            if (StrLength(tbCustAddress.Text.Trim()) > 60 && cbMoney.Text == "本币")
            {
                MessageBox.Show("内销客户地址超出60个字符，请调整！ ");
                return;
            }
            if (StrLength(tbCustAddress.Text.Trim()) > 120 && cbMoney.Text == "外币")
            {
                MessageBox.Show("外贸客户地址超出120个字符，请调整！ ");
                return;
            }

            if (StrLength(tbTaxCode.Text.Trim()) > 20)
            {
                MessageBox.Show("税号超出20个字符，请检查！ ");
                return;
            }
            if (StrLength(tbProvince.Text.Trim()) > 10)
            {
                MessageBox.Show("客户所在省份超出10个字符，请调整！ ");
                return;
            }
            if (StrLength(tbUniteAccount.Text.Trim()) != 11)
            {
                MessageBox.Show("合并账号不是11位，请检查！ ");
                return;
            }
            #region 检查客户名称是否重复
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                string CustomerName = EncodingLD.GetString(EncodingCH.GetBytes(tbCustomerName.Text.Trim()));
                SqlCommand cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "' and CustomerID not like'" + tbCustomerCode.Text.Trim().Substring(0, 6) + "%'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count > 0)
                {
                    MessageBox.Show("有相同客户名称的记录，请检查" + dtcust.Rows[0][0].ToString());
                    return;
                }
            }
            #endregion
            #region 检查税号是否重复
            if (!string.IsNullOrEmpty(tbTaxCode.Text.Trim()))
            {
                using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                {
                    string CustomerName = tbTaxCode.Text.Trim().Replace(" ", "");
                    SqlCommand cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where FederalPrimaryTaxExemptCertificateNumber = '" + CustomerName + "' and CustomerID not like'" + tbCustomerCode.Text.Trim().Substring(0, 6) + "%'", conn);
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dtcust = new DataTable();
                    sda.Fill(dtcust);
                    if (dtcust.Rows.Count > 0)
                    {
                        MessageBox.Show("有相同税号的记录，请检查" + dtcust.Rows[0][0].ToString());
                        return;
                    }

                }
            }
            #endregion
            listBoxCustomer.Items.Clear();
            if (CbChildCustomer.Checked)
            {
                //int i = 0;
                //using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                //{
                //    using (OleDbCommand cmd = new OleDbCommand(@"select CustomerName,CustomerAddress1,CustomerContact,CustomerContactPhone,CustomerContactFax,AccountingContact,              AccountingContactPhone,AccountingContactFax,CustomerControllingCode,CustomerCurrencyCode, BankReference1,BankReference2,FederalPrimaryTaxExemptCertificateNumber  from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn))
                //    {
                //        conn.Open();
                //        DataTable dt = new DataTable();
                //        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                //        da.Fill(dt);
                //        i = dt.Rows.Count;
                //        conn.Close();
                //    }

                //}
                //i==0  表示没有主客户
                listBoxCustomer.Items.Add("开始增加子客户------->-------->------->-------->");
                if (AddChildCustomer() == true)
                {
                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户成功<<<");
                    MessageBox.Show("子客户添加成功！");
                    #region 检查是否有重复的客户名称
                    using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                    {
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("select CustomerName from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn);
                        object vendornamelading = cmd.ExecuteScalar();
                        if (vendornamelading == null)
                        { MessageBox.Show("没有该客户的记录!"); }
                        else
                        {
                            string CustomerName = vendornamelading.ToString();
                            cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "'", conn);
                            SqlDataAdapter sda = new SqlDataAdapter(cmd);
                            DataTable dtcust = new DataTable();
                            sda.Fill(dtcust);
                            foreach (DataRow dr in dtcust.Rows)
                            {
                                if (dr["CustomerID"].ToString().Trim().Substring(0, 6) != tbCustomerCode.Text.Trim().Substring(0, 6))
                                {
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                }
                            }
                        }
                    }
                    #endregion
                    #region 客户信息groupBox7信息清空
                    foreach (Control control in groupBox7.Controls)
                    {
                        if (!(control is Label))
                        {
                            control.Text = null;
                        }
                    }
                    #endregion
                    try
                    {
                        dgvCustomer.Rows[Convert.ToInt32(customerhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    catch (Exception)
                    {

                    }

                }
                else
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户失败<<<");
                    MessageBox.Show(" 子客户添加失败！");
                }
            }
            else  //BPM流程需要添加客户及子客户
            {
                    listBoxCustomer.Items.Add("开始增加主客户-->");
                    if (AddCustomerCompany())
                    {
                        listBoxCustomer.Items.Add("-->主客户录入成功");

                    }
                    else
                    {
                        listBoxCustomer.Items.Add("-->主客户录入失败");
                        MessageBox.Show("主客户录入失败！请检查！！！");
                        return;
                    }
                
                listBoxCustomer.Items.Add("开始增加子客户------->-------->------->-------->");

                if (AddChildCustomer() == true)
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户成功<<<");
                    MessageBox.Show("主客户子客户添加成功！");
                    #region 检查是否有重复的客户名称
                    using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                    {
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("select CustomerName from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn);
                        object vendornamelading = cmd.ExecuteScalar();
                        if (vendornamelading == null)
                        { MessageBox.Show("没有该客户的记录!"); }
                        else
                        {
                            string CustomerName = vendornamelading.ToString();
                            cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "'", conn);
                            SqlDataAdapter sda = new SqlDataAdapter(cmd);
                            DataTable dtcust = new DataTable();
                            sda.Fill(dtcust);
                            foreach (DataRow dr in dtcust.Rows)
                            {
                                if (dr["CustomerID"].ToString().Trim().Substring(0, 6) != tbCustomerCode.Text.Trim().Substring(0, 6))
                                {
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                }
                            }
                        }
                    }
                    #endregion
                    #region 客户信息groupBox7信息清空
                    foreach (Control control in groupBox7.Controls)
                    {
                        if (!(control is Label))
                        {
                            control.Text = null;
                        }
                    }
                    #endregion
                    try
                    {
                        dgvCustomer.Rows[Convert.ToInt32(customerhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    catch (Exception)
                    {

                    }

                }
                else
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户失败<<<");
                    MessageBox.Show(" 子客户添加失败！");
                }

            }
           
        }
        int linshiindex = -100;
        private bool AddChildCustomer()//程序增加子客户
        {
            #region 增加GLOS GLAV
            ADDGLOS(tbUniteAccount.Text.Trim(), tbCustomerName.Text.Trim(), listBoxCustomer);
            ADDGLAV(tbUniteAccount.Text.Trim(), "113100");
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                SqlCommand cmd = new SqlCommand("SELECT * FROM [dbo].[_NoLock_FS_GLAccountOrganizationValidation] where GLAccountNumber='113100' and GLOrganization='" + tbUniteAccount.Text.Trim() + "'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count == 0)
                {
                    if (linshiindex == -100)
                    {
                        MessageBox.Show(tbUniteAccount.Text.Trim() + "-113100 GLOS GLAV添加失败，子客户未添加，请检查");
                    }
                    return false;
                }
            }
            #endregion
            #region 录入子客户主题信息
            string CustomerCode = tbCustomerCode.Text.Trim().Substring(0, 7);//子客户代码
            //添加客户基础信息
            SOPC00 myCustomerBasic = new SOPC00();
            myCustomerBasic.CustomerID.Value = CustomerCode;//子客户代码
            myCustomerBasic.CustomerName.Value = tbCustomerName.Text.Trim();//子客户名称
            myCustomerBasic.CustomerLevel.Value = "C";//客户是否为主公司，p为主公司,c为子客户
            myCustomerBasic.TradeClassName.Value = cbIndustry.Text.Trim();//行业类别
            myCustomerBasic.ParentCustomer.Value = CustomerCode.Substring(0, 6);//主客户
            if (_fstiClient.ProcessId(myCustomerBasic, null))
            {
                listBoxCustomer.Items.Add("子客户新建成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                if (linshiindex != -100)
                {
                    dgvUpdateStockNum["状态", linshiindex].Value = "子客户新建失败";
                }
                listBoxCustomer.Items.Add("子客户新建失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerBasic, itemError, listBoxCustomer);
                return false;
            }
            
            //添加客户财务应收账款及税金
            SOPC04 myCustomerAccount = new SOPC04();
            myCustomerAccount.CustomerID.Value = CustomerCode;//客户代码 
            string result = DateTime.Now.ToString("MMddyy");
            myCustomerAccount.CustomerStartDate.Value = result;//客户启用日期,格式为：032517

            if (cbMoney.Text.Trim() == "本币")
            {
                myCustomerAccount.AROpenItemBalanceForwardCode.Value = "B";
            }
            else
            {
                myCustomerAccount.AROpenItemBalanceForwardCode.Value = "O";
            }
            myCustomerAccount.FederalPrimaryTaxExemptCertificateNumber.Value = tbTaxCode.Text.Trim();//国家免税证书 VAT税金代码
            if (_fstiClient.ProcessId(myCustomerAccount, null))
            {
                listBoxCustomer.Items.Add("子客户财务应收账款及税金添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("子客户财务应收账款及税金添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerAccount, itemError, listBoxCustomer);
                return false;
            }
            //添加通信地址
            SOPC01 myCustomerAddress = new SOPC01();
            myCustomerAddress.CustomerID.Value = CustomerCode;//子客户代码
            #region 客户地址
            string strCustAddress = tbCustAddress.Text.Trim();
            if (StrLength(strCustAddress) > 60)
            {
                int Len1 = strCustAddress.Length;
                for (int i = strCustAddress.Length; i > 0; i--)
                {
                    if (StrLength(strCustAddress.Substring(0, i)) <= 60)
                    { Len1 = i; break; }
                }
                myCustomerAddress.CustomerAddress1.Value = strCustAddress.Substring(0, Len1);
                myCustomerAddress.CustomerAddress2.Value = strCustAddress.Substring(Len1, strCustAddress.Length - Len1);
            }
            else
            {
                myCustomerAddress.CustomerAddress1.Value = strCustAddress;//客户地址 
            }
            #endregion 
            myCustomerAddress.CustomerContact.Value = tbContactPerson.Text.Trim();//客户联系人姓名
            myCustomerAddress.CustomerContactPhone.Value = tbContactTelephone.Text.Trim();//客户联系人电话
            myCustomerAddress.CustomerContactFax.Value = tbContactFax.Text.Trim();//客户联系人传真
            myCustomerAddress.CustomerZip.Value = tbPostcode.Text.Trim();//客户联系人邮编
            myCustomerAddress.CustomerState.Value = tbProvince.Text.Trim();//客户所在省份
            if (_fstiClient.ProcessId(myCustomerAddress, null))
            {
                listBoxCustomer.Items.Add("子客户地址添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("子客户地址添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerAddress, itemError, listBoxCustomer);
                return false;
            }

            //概要及信用
            SOPC03 myCustomerSales = new SOPC03();
            myCustomerSales.CustomerID.Value = CustomerCode;//子客户代码
            myCustomerSales.SalesRegion.Value = tbSalesmanName.Text.ToString().Trim();//销售地区
            myCustomerSales.CSR.Value = tbSalesmanCode.Text.ToString().Trim();//客户服务代表字段
            myCustomerSales.CustomerState.Value = "A";
            if (cbMoney.Text.Trim() == "本币")
            {
                myCustomerSales.CreditLimitControllingAmount.Value = "1.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "N";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "L";
                myCustomerSales.CustomerCurrencyCode.Value = "00000";
                myCustomerSales.ShipmentCreditHoldCode.Value = "R";//客户订单信用强制--发货
                myCustomerSales.OrderEntryCreditHoldCode.Value = "R";//客户订单信用强制--订单录入
            }
            else
            {
                myCustomerSales.CreditLimitControllingAmount.Value = "0.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "N";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "F";
                myCustomerSales.CustomerCurrencyCode.Value = tbCustomerCurrencyCode.Text.Trim();
                myCustomerSales.ShipmentCreditHoldCode.Value = "H";//客户订单信用强制--发货
                myCustomerSales.OrderEntryCreditHoldCode.Value = "H";//客户订单信用强制--订单录入
            }
            //myCustomerSales.ShipmentCreditHoldCode.Value = "H";//客户订单信用强制--发货
            //myCustomerSales.OrderEntryCreditHoldCode.Value = "H";//客户订单信用强制--订单录入
            if (_fstiClient.ProcessId(myCustomerSales, null))
            {
                listBoxCustomer.Items.Add("子客户概要及信用添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("子客户概要及信用添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerSales, itemError, listBoxCustomer);
                return false;
            }

            //添加会计信息、银行账户和应收账款账号
            SOPC05 myCustomerBank = new SOPC05();
            myCustomerBank.CustomerID.Value = CustomerCode;//客户代码
            myCustomerBank.AccountingContact.Value = tbAccountantName.Text.Trim();//会计
            myCustomerBank.AccountingContactPhone.Value = tbAccountantPhone.Text.Trim();//会计电话
            myCustomerBank.BankReference1.Value = tbBankOfDeposit.Text.Trim();//开户银行
            myCustomerBank.BankReference2.Value = tbBankAccount.Text.Trim();//开户银行账号
            myCustomerBank.AccountsReceivableAccount.Value = tbUniteAccount.Text.Trim() + "-113100";
            if (_fstiClient.ProcessId(myCustomerBank, null))
            {
                listBoxCustomer.Items.Add("子客户财务会计信息添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("子客户财务会计信息添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerBank, itemError, listBoxCustomer);
                return false;
            }

            #endregion

            return true;
        }

        private bool AddCustomerCompany()//程序增加主客户
        {
            string CustomerCode = tbCustomerCode.Text.Trim().Substring(0, 6);//客户代码
            //添加客户基础信息
            SOPC00 myCustomerBasic = new SOPC00();
            myCustomerBasic.CustomerID.Value = CustomerCode;//客户代码
            myCustomerBasic.CustomerName.Value = tbCustomerName.Text.Trim();//客户名称
            myCustomerBasic.CustomerLevel.Value = "P";//客户是否为主公司，p为主公司
            myCustomerBasic.TradeClassName.Value = cbIndustry.Text.Trim();//行业类别
            if (_fstiClient.ProcessId(myCustomerBasic, null))
            {
                listBoxCustomer.Items.Add("主客户新建成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("主客户新建失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerBasic, itemError, listBoxCustomer);
                return false;
            }
            //添加通信地址
            SOPC01 myCustomerAddress = new SOPC01();
            myCustomerAddress.CustomerID.Value = CustomerCode;//客户代码
            #region 客户地址
            string strCustAddress = tbCustAddress.Text.Trim();
            if (StrLength(strCustAddress) > 60)
            {
                int Len1 = strCustAddress.Length;
                for (int i = strCustAddress.Length; i > 0; i--)
                {
                    if (StrLength(strCustAddress.Substring(0, i)) <= 60)
                    { Len1 = i; break; }
                }
                myCustomerAddress.CustomerAddress1.Value = strCustAddress.Substring(0, Len1);
                myCustomerAddress.CustomerAddress2.Value = strCustAddress.Substring(Len1, strCustAddress.Length - Len1);
            }
            else
            {
                myCustomerAddress.CustomerAddress1.Value = strCustAddress;//客户地址 
            }
            #endregion
            myCustomerAddress.CustomerContact.Value = tbContactPerson.Text.Trim();//客户联系人姓名
            myCustomerAddress.CustomerContactPhone.Value = tbContactTelephone.Text.Trim();//客户联系人电话
            myCustomerAddress.CustomerContactFax.Value = tbContactFax.Text.Trim();//客户联系人传真
            myCustomerAddress.CustomerZip.Value = tbPostcode.Text.Trim();//客户联系人邮编
            myCustomerAddress.CustomerState.Value = tbProvince.Text.Trim();//客户所在省份
            if (_fstiClient.ProcessId(myCustomerAddress, null))
            {
                listBoxCustomer.Items.Add("主客户地址添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("主客户地址添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerAddress, itemError, listBoxCustomer);
                return false;
            }
            //添加客户财务应收账款及税金
            SOPC04 myCustomerAccount = new SOPC04();
            myCustomerAccount.CustomerID.Value = CustomerCode;//客户代码 
            string result = DateTime.Now.ToString("MMddyy");
            myCustomerAccount.CustomerStartDate.Value = result;//客户启用日期,格式为：032517

            if (cbMoney.Text.Trim() == "本币")
            {
                myCustomerAccount.AROpenItemBalanceForwardCode.Value = "B";
            }
            else
            {
                myCustomerAccount.AROpenItemBalanceForwardCode.Value = "O";
            }
            myCustomerAccount.FederalPrimaryTaxExemptCertificateNumber.Value = tbTaxCode.Text.Trim();//国家免税证书 VAT税金代码
            if (_fstiClient.ProcessId(myCustomerAccount, null))
            {
                listBoxCustomer.Items.Add("主客户财务应收账款及税金添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("主客户财务应收账款及税金添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerAccount, itemError, listBoxCustomer);
                return false;
            }
            //概要及信用
            SOPC03 myCustomerSales = new SOPC03();
            myCustomerSales.CustomerID.Value = CustomerCode;//客户代码
            myCustomerSales.SalesRegion.Value = tbSalesmanName.Text.ToString().Trim();//销售地区
            myCustomerSales.CSR.Value = tbSalesmanCode.Text.ToString().Trim();//客户服务代表字段
            myCustomerSales.CustomerState.Value = "A";
            if (cbMoney.Text.Trim() == "本币")
            {
                myCustomerSales.CreditLimitControllingAmount.Value = "1.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "N";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "L";
                myCustomerSales.CustomerCurrencyCode.Value = "00000";
                myCustomerSales.ShipmentCreditHoldCode.Value = "R";//客户订单信用强制--发货
                myCustomerSales.OrderEntryCreditHoldCode.Value = "R";//客户订单信用强制--订单录入
            }
            else
            {
                myCustomerSales.CreditLimitControllingAmount.Value = "0.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "N";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "F";
                myCustomerSales.CustomerCurrencyCode.Value = tbCustomerCurrencyCode.Text.Trim();
                myCustomerSales.ShipmentCreditHoldCode.Value = "H";//客户订单信用强制--发货
                myCustomerSales.OrderEntryCreditHoldCode.Value = "H";//客户订单信用强制--订单录入
            }
            //myCustomerSales.ShipmentCreditHoldCode.Value = "H";//客户订单信用强制--发货
            //myCustomerSales.OrderEntryCreditHoldCode.Value = "H";//客户订单信用强制--订单录入
            if (_fstiClient.ProcessId(myCustomerSales, null))
            {
                listBoxCustomer.Items.Add("主客户概要及信用添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("主客户概要及信用添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerSales, itemError, listBoxCustomer);
                return false;
            }

            //添加会计信息、银行账户和应收账款账号
            SOPC05 myCustomerBank = new SOPC05();
            myCustomerBank.CustomerID.Value = CustomerCode;//客户代码
            myCustomerBank.AccountingContact.Value = tbAccountantName.Text.Trim();//会计
            myCustomerBank.AccountingContactPhone.Value = tbAccountantPhone.Text.Trim();//会计电话
            myCustomerBank.BankReference1.Value = tbBankOfDeposit.Text.Trim();//开户银行
            myCustomerBank.BankReference2.Value = tbBankAccount.Text.Trim();//开户银行账号
            if (_fstiClient.ProcessId(myCustomerBank, null))
            {
                listBoxCustomer.Items.Add("主客户财务会计信息添加成功:");
                listBoxCustomer.Items.Add(_fstiClient.CDFResponse);
            }
            else
            {
                listBoxCustomer.Items.Add("主客户财务会计信息添加失败:");
                FSTIError itemError = _fstiClient.TransactionError;
                DumpErrorObject(myCustomerBank, itemError, listBoxCustomer);
                return false;
            }


            return true;
        }

        private void tbCustomerCode_KeyDown(object sender, KeyEventArgs e)//通过客户代码获得客户信息
        {
            if (e.KeyCode != Keys.Enter)
                return;
            //定义简体中文和西欧文编码字符集
            //Encoding GB2312 = Encoding.GetEncoding("gb2312");
            //Encoding ISO88591 = Encoding.GetEncoding("iso-8859-1");
            AddCustomer.Enabled = false;
            CbChildCustomer.Checked = true;
            string CustomerCode = tbCustomerCode.Text.ToString().Trim().ToUpper();
            tbCustomerCode.Text = CustomerCode;
            int i = CustomerCode.Length;//获得客户代码的长度 
            if (i == 6)
            {
                foreach (Control ct in groupBox7.Controls)
                {
                    if (ct is TextBox)
                        ct.Text = "";
                }
                tbCustomerCode.Text = CustomerCode;
            }
            else if (i == 7)//代码长度为7位，表示添加的为子公司，需要查询主公司的信息
            {
                string ParentCustomerCode = CustomerCode.Substring(0, 6);
                string strsql = "select CustomerName,CustomerAddress1,CustomerAddress2,CustomerZip,CustomerContact,CustomerContactPhone,CustomerContactFax,AccountingContact, AccountingContactPhone,AccountingContactFax,CSR,SalesRegion,TradeClassName,CustomerControllingCode,CustomerCurrencyCode, BankReference1,BankReference2,FederalPrimaryTaxExemptCertificateNumber  from _NoLock_FS_Customer where CustomerID = '" + ParentCustomerCode + "'";
                using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                {
                    using (OleDbCommand cmd = new OleDbCommand(strsql, conn))
                    {
                        conn.Open();
                        OleDbDataAdapter oledbDA = new OleDbDataAdapter(cmd);
                        DataTable dt1 = new DataTable();
                        oledbDA.Fill(dt1);
                        conn.Close();
                        if (dt1.Rows.Count == 1)
                        {
                            DataRow myDR = dt1.Rows[0];
                            tbCustomerName.Text = myDR["CustomerName"].ToString();
                            tbCustAddress.Text = myDR["CustomerAddress1"].ToString() + myDR["CustomerAddress2"].ToString().Trim();
                            tbContactPerson.Text = myDR["CustomerContact"].ToString();
                            tbContactTelephone.Text = myDR["CustomerContactPhone"].ToString();
                            tbContactFax.Text = myDR["CustomerContactFax"].ToString();
                            tbPostcode.Text = myDR["CustomerZip"].ToString();
                            cbIndustry.Text = myDR["TradeClassName"].ToString();
                            tbSalesmanName.Text = myDR["SalesRegion"].ToString();
                            tbSalesmanCode.Text = myDR["CSR"].ToString();
                            tbAccountantName.Text = myDR["AccountingContact"].ToString();
                            tbAccountantPhone.Text = myDR["AccountingContactPhone"].ToString();
                            cbMoney.Text = myDR["CustomerControllingCode"].ToString() == "L" ? "本币" : "外币";
                            tbCustomerCurrencyCode.Text = myDR["CustomerCurrencyCode"].ToString();
                            tbBankOfDeposit.Text = myDR["BankReference1"].ToString();
                            tbBankAccount.Text = myDR["BankReference2"].ToString();
                            tbTaxCode.Text = myDR["FederalPrimaryTaxExemptCertificateNumber"].ToString();

                            AddCustomer.Enabled = true;
                            string Customercode1 = tbCustomerCode.Text.Trim();
                            tbUniteAccount.Text = "1A" + Customercode1.Substring(6, 1) + Customercode1.Substring(0, 1) + "-" + Customercode1.Substring(1, 2) + "-" + Customercode1.Substring(3, 3);
                        }
                        if (dt1.Rows.Count == 0)
                        {
                            foreach (Control ct in groupBox7.Controls)
                            {
                                if (ct is TextBox)
                                    ct.Text = "";
                            }
                            tbCustomerCode.Text = CustomerCode;
                            MessageBox.Show("没有该客户！");
                        }
                    }

                }
                tBCustBZ.Text = "";
                customercode.Text = "customercode";
                customerliushuihao.Text = " liushuihao";
                customerhanghao.Text = " hanghao";
            }
            else
            {
                foreach (Control ct in groupBox7.Controls)
                {
                    if (ct is TextBox)
                        ct.Text = "";
                }
                tbCustomerCode.Text = CustomerCode;
                MessageBox.Show("客户代码不准确，请进行确认！");
            }
        }

        private void UpdateCustomer_Click(object sender, EventArgs e)
        {
            MessageBox.Show("请手动修改！");
            UpdateCustomer.Enabled = false;
            #region 客户信息groupBox7信息清空
            foreach (Control control in groupBox7.Controls)
            {
                if (!(control is Label))
                {
                    control.Text = null;
                }
            }
            #endregion
        }

        private void ActiveCustomer_Click(object sender, EventArgs e)
        {
            AddCustomer.Enabled = true;
            UpdateCustomer.Enabled = true;
        }


        private void textUserId_TextChanged(object sender, EventArgs e)
        {
            textUserId.Text = textUserId.Text.ToUpper();
            textUserId.SelectionStart = textUserId.Text.Length;
        }

        private void tbGB2312_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                tbISO.Text = EncodingLD.GetString(EncodingCH.GetBytes(tbGB2312.Text.Trim()));
            }
        }

        private void tbISO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                tbGB2312.Text = EncodingCH.GetString(EncodingLD.GetBytes(tbISO.Text.Trim()));
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedTab == tabPage4)
            {
                this.qiyongItemCode.Focus();
            }
            if (this.tabControl1.SelectedTab == tabPage6)
            {
                this.textPassword.Focus();
            }
        }
        private void qiyongItemCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || (e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == '\b' || sign)
            {
                e.KeyChar = Convert.ToChar(e.KeyChar.ToString().ToUpper());
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }
        private bool sign = false;

        private void qiyongItemCode_KeyDown(object sender, KeyEventArgs e)//物料启用
        {
            if (e.KeyData == (Keys.C | Keys.Control) || e.KeyData == (Keys.A | Keys.Control) || e.KeyData == (Keys.V | Keys.Control) || e.KeyData == (Keys.X | Keys.Control))
                sign = true;
            else
                sign = false;
            if (e.KeyCode == Keys.Enter)
            {
                if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
                {
                    MessageBox.Show("请登录四班账号！");
                    return;
                }
                if (qiyong.Enabled == false) return;
                wuliaoqiyong.Items.Clear();
                ITMB01 myItmb = new ITMB01();
                string ItemCode = ITMB["物料", 0].Value.ToString().Trim().ToUpper();
                myItmb.ItemNumber.Value = ItemCode;
                myItmb.ItemStatus.Value = "A";      //状态A  
                if (_fstiClient.ProcessId(myItmb, null))
                {
                    wuliaoqiyong.Items.Add(ItemCode + "物料启用成功");
                    qiyong.Enabled = false;
                    MessageBox.Show(ItemCode + ":启用成功");
                    #region 刷新
                    qiyongItemCode_KeyUp(sender, e);
                    #endregion
                }
                else
                {

                    MessageBox.Show(ItemCode + ":启用失败");
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myItmb, itemError, wuliaoqiyong);
                    wuliaoqiyong.Items.Add(ItemCode + "物料启用失败");
                }
                #region 注释的
                //wuliaoqiyong.Items.Clear();
                //qiyong.Enabled = false;
                //OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB);
                //OleDbCommand cmd = new OleDbCommand("SELECT ItemNumber 物料,ItemDescription 描述,ItemUM 单位,dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计,MakeBuyCode 制购,ItemType 物类,ItemStatus 状态,IsLotTraced 批号,IsSerialized 系号,OrderPolicy 订货,IsInspectionRequired 要求检验 FROM[dbo].[_NoLock_FS_Item] INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey where ItemNumber = '" + qiyongItemCode.Text.Trim().ToUpper() + "' and (dbo._NoLock_FS_ItemCost.CostType = '0')", conn);
                //DataTable dt = new DataTable();
                //OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                //da.Fill(dt);
                //ITMB.DataSource = dt;
                //ITMB.Columns[1].Width = 330;
                //ITMB.Columns["累计成本合计"].Width = 140;
                //ITMB.Columns["制购"].Width = 60;
                //ITMB.Columns["物类"].Width = 60;
                //ITMB.Columns["状态"].DefaultCellStyle.ForeColor = Color.Red;
                //if (ITMB.Rows.Count == 1)
                //{
                //    if (ITMB["状态", 0].Value.ToString().Trim() == "O")
                //        qiyong.Enabled = true;
                //}
                #endregion
            }
        }
        private void qiyongItemCode_KeyUp(object sender, KeyEventArgs e)
        {
            if (qiyongItemCode.Text.Length > 4)
            {

                wuliaoqiyong.Items.Clear();
                qiyong.Enabled = false;
                OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB);
                OleDbCommand cmd = new OleDbCommand("SELECT ItemNumber 物料,ItemDescription 描述,ItemUM 单位,dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计,MakeBuyCode 制购,ItemType 物类,ItemStatus 状态,IsLotTraced 批号,IsSerialized 系号,OrderPolicy 订货,IsInspectionRequired 要求检验 FROM[dbo].[_NoLock_FS_Item] INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey where ItemNumber = '" + qiyongItemCode.Text.Trim().ToUpper() + "' and (dbo._NoLock_FS_ItemCost.CostType = '0')", conn);
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                ITMB.DataSource = dt;
                ITMB.Columns[1].Width = 330;
                ITMB.Columns["累计成本合计"].Width = 140;
                ITMB.Columns["制购"].Width = 60;
                ITMB.Columns["物类"].Width = 60;
                ITMB.Columns["状态"].DefaultCellStyle.ForeColor = Color.Red;
                if (ITMB.Rows.Count == 1 && ITMB["状态", 0].Value.ToString().Trim() == "O")
                {
                    qiyong.Enabled = true;
                }
                else
                {
                    qiyong.Enabled = false;
                }


            }

        }

        private void dgvItmbDetail_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, ((DataGridView)sender).RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), ((DataGridView)sender).RowHeadersDefaultCellStyle.Font, rectangle, ((DataGridView)sender).RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void GetITMBsh_Click(object sender, EventArgs e)
        {
            ItemNamesh.Text = "";
            ItemNamedgvsh.DataSource = null;
            dgvItmbDetailsh.DataSource = null;
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and   PROCESSNAME='RY增加物料申请流程' and  STEPLABEL = 'ERP管理员审核'");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and   PROCESSNAME='RY增加物料申请流程' and (STEPLABEL = '系统管理员维护' or STEPLABEL = 'ERP管理员审核') and STARTTIME >'2019-5-10'");
            List<ITMBliucheng> list1 = new List<ITMBliucheng>();
            foreach (DataRow dr in Incidents.Rows)
            {
                ITMBliucheng Vendor1 = TolistITMB(SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_ZJWLCB] where REV_INCIDENT=" + dr[0]));
                list1.Add(Vendor1);

            }

            dgvItmbsh.DataSource = list1;
            for (int i = 0; i < this.dgvItmbsh.Columns.Count; i++)
            {
                this.dgvItmbsh.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvItmbsh.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void dgvItmbsh_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowindex = e.RowIndex;
            if (rowindex != -1)
            {
                ItemNamesh.Text = "";
                ItemNamedgvsh.DataSource = null;
                dgvItmbsh.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                for (int a = 0; a < dgvItmbsh.Rows.Count; a++)
                {
                    if (a != rowindex)
                        dgvItmbsh.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                }

                ITMBjieshoudanweish.Text = dgvItmbsh.Rows[rowindex].Cells["接收单位"].Value.ToString().Trim();
                ITMBliushuihaosh.Text = dgvItmbsh.Rows[rowindex].Cells["流水号"].Value.ToString().Trim();
                ITMBhanghaosh.Text = (rowindex + 1).ToString();
                string ParentGuid = dgvItmbsh.Rows[rowindex].Cells["ParentGuid"].Value.ToString();

                dgvItmbDetailsh.DataSource = toITMBITMCsh(SqlHelper.ExecuteDataTable("SELECT ltrim(rtrim(WLBM1))  as 物料代码,WLMS as 物料描述,DW as 单位,ltrim(rtrim(KGYDM)) as 库管员代码,upper(ltrim(rtrim(JHY))) as 计划采购,YX as 运行,FIX,JY as 检验,PLDHTS as 批量订货天数,ZXPLDH as 最小批量订货,PLDHBS as 批量订货倍数,PLDHSM as 批量订货数目,QSGZZX as 起始工作中心,YXK as 优先库,W as 位," +
                    "CLF as 材料费,HJ as 合计,CPX as 产品线,KCZH as 库存账号,XSZH as 销售账号,CBZH as 成本账号,YCM as 预测码,YCJD as 预测阶段 FROM YW_ZJWLCB_EX where ParentGuid = '" + ParentGuid + "'"));
                for (int i = 0; i < this.dgvItmbDetailsh.Columns.Count; i++)
                {
                    this.dgvItmbDetailsh.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    this.dgvItmbDetailsh.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
                for (int i = 0; i < this.dgvItmbDetailsh.Rows.Count; i++)
                {
                    String Item = dgvItmbDetailsh["物料描述", i].Value.ToString().Trim();

                    if (Item == "" || dgvItmbDetailsh["单位", i].Value.ToString().Trim() == "" || dgvItmbDetailsh["计划采购", i].Value.ToString().Trim() == "" || dgvItmbDetailsh["运行", i].Value.ToString().Trim() == "" || dgvItmbDetailsh["FIX", i].Value.ToString().Trim() == "" || dgvItmbDetailsh["批量订货天数", i].Value.ToString().Trim() == "" || dgvItmbDetailsh["最小批量订货", i].Value.ToString().Trim() == "")
                    {
                        //dgvItmbDetailsh.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["单位"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["计划采购"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["运行"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["FIX"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["批量订货天数"].Style.BackColor = Color.Red;
                        dgvItmbDetailsh.Rows[i].Cells["最小批量订货"].Style.BackColor = Color.Red;

                    }
                    if (dgvItmbsh.Rows[rowindex].Cells["制购类型"].Value.ToString().Trim().Substring(0, 1) == "M")
                    {
                        if (dgvItmbDetailsh["起始工作中心", i].Value.ToString().Trim() == "")
                        { dgvItmbDetailsh.Rows[i].Cells["起始工作中心"].Style.BackColor = Color.Red; }
                        if (dgvItmbDetailsh["批量订货数目", i].Value.ToString().Trim() == "")
                        {
                            MessageBox.Show("第" + (i + 1) + "行[批量订货数目]有误，请检查！");
                            dgvItmbDetailsh.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            try
                            {
                                if (Convert.ToDecimal(dgvItmbDetailsh["批量订货数目", i].Value.ToString().Trim()) <= 1)
                                {
                                    MessageBox.Show("第" + (i + 1) + "行[批量订货数目]<= 1，请检查！");
                                    dgvItmbDetailsh.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("第" + (i + 1) + "行[批量订货数目]有误，请检查！");
                                dgvItmbDetailsh.Rows[i].Cells["批量订货数目"].Style.BackColor = Color.Red;
                            }
                        }
                    }

                    if (StrLength(Item) > 70)
                    {
                        MessageBox.Show("第" + (i + 1) + "行[物料描述]超出字符数限制");
                        dgvItmbDetailsh.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                    }
                    if (Item.Contains(' ') || Item.Contains('（') || Item.Contains('）') || Item.ToUpper() != Item)
                    {
                        MessageBox.Show("第" + (i + 1) + "行[物料描述]不符合ERP物料描述规范");
                        dgvItmbDetailsh.Rows[i].Cells["物料描述"].Style.BackColor = Color.Red;
                    }
                }
                for (int y = 0; y < this.dgvItmbDetailsh.Rows.Count; y++)
                {
                    for (int x = 0; x < this.dgvItmbDetailsh.Columns.Count; x++)
                    {
                        if (dgvItmbDetailsh.Rows[y].Cells[x].Style.BackColor == Color.Red)
                        {
                            MessageBox.Show("信息有误已红色标示，请检查！"); return;
                        }
                    }
                }
            }
        }

        private IEnumerable<ITMBITMC> toITMBITMCsh(DataTable dt)//DataRow转 流程明细信息
        {
            List<ITMBITMC> ITMBitmc = new List<ITMBITMC>();
            foreach (DataRow dr in dt.Rows)
            {
                ITMBITMC iTMBITMC = new ITMBITMC();
                iTMBITMC.物料代码 = dr["物料代码"].ToString().Trim().ToUpper();
                iTMBITMC.物料描述 = dr["物料描述"].ToString().Trim().ToUpper();
                iTMBITMC.单位 = dr["单位"].ToString().Trim().ToUpper();
                iTMBITMC.库管员代码 = dr["库管员代码"].ToString().Trim().ToUpper();
                iTMBITMC.计划采购 = dr["计划采购"].ToString().Trim().ToUpper();
                iTMBITMC.运行 = dr["运行"].ToString().Trim().ToUpper();
                iTMBITMC.FIX = dr["FIX"].ToString().Trim().ToUpper();
                iTMBITMC.检验 = dr["检验"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货天数 = dr["批量订货天数"].ToString().Trim().ToUpper();
                iTMBITMC.最小批量订货 = dr["最小批量订货"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货倍数 = dr["批量订货倍数"].ToString().Trim().ToUpper();
                iTMBITMC.批量订货数目 = dr["批量订货数目"].ToString().Trim().ToUpper();
                iTMBITMC.起始工作中心 = dr["起始工作中心"].ToString().Trim().ToUpper();
                iTMBITMC.优先库 = dr["优先库"].ToString().Trim().ToUpper();
                iTMBITMC.位 = dr["位"].ToString().Trim().ToUpper();
                if (dr["材料费"].ToString().Trim() != "")
                    iTMBITMC.材料费 = Math.Round(Convert.ToDouble(dr["材料费"].ToString().Trim()), 9).ToString();
                if (dr["合计"].ToString().Trim() != "")
                    iTMBITMC.合计 = Math.Round(Convert.ToDouble(dr["合计"].ToString().Trim()), 9).ToString();
                iTMBITMC.产品线 = dr["产品线"].ToString().Trim().ToUpper();
                iTMBITMC.库存账号 = dr["库存账号"].ToString().Trim().ToUpper();
                iTMBITMC.销售账号 = dr["销售账号"].ToString().Trim().ToUpper();
                iTMBITMC.成本账号 = dr["成本账号"].ToString().Trim().ToUpper();
                iTMBITMC.预测码 = dr["预测码"].ToString().Trim().ToUpper();
                iTMBITMC.预测阶段 = dr["预测阶段"].ToString().Trim().ToUpper();
                ITMBitmc.Add(iTMBITMC);

            }
            return ITMBitmc;
        }

        private void dgvItmbDetailsh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            if (rowIndex > -1)
            {
                ItemNamesh.Text = dgvItmbDetailsh.Rows[rowIndex].Cells[1].Value.ToString();
                if (!string.IsNullOrEmpty(ItemNamesh.Text.Trim()))
                {
                    #region
                    Encoding EncodingEN = Encoding.GetEncoding("ISO-8859-1");
                    Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                    string nametolading = "%" + EncodingEN.GetString(EncodingCH.GetBytes(ItemNamesh.Text.Trim())) + "%";
                    string fenLei = ItemClasssh.Text == "全部" ? "%" : ItemClasssh.Text;
                    #endregion
                    using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                    {
                        OleDbCommand cmd = new OleDbCommand(" SELECT dbo._NoLock_FS_Item.ItemNumber AS 物料代码, dbo._NoLock_FS_Item.ItemDescription AS 物料描述, dbo._NoLock_FS_Item.ItemUM AS 物料单位, dbo._NoLock_FS_ItemCost.CostType AS 类别, dbo._NoLock_FS_ItemCost.CostCode AS 方法, dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计, dbo._NoLock_FS_Item.GatewayWorkCenter AS 工作中心 FROM dbo._NoLock_FS_Item INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey WHERE (dbo._NoLock_FS_ItemCost.CostType = '0') AND (dbo._NoLock_FS_Item.ItemDescription LIKE '" + nametolading + "') and (dbo._NoLock_FS_Item.ItemNumber like '" + fenLei + "%')", conn);
                        DataTable dt = new DataTable();
                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dt);
                        ItemNamedgvsh.DataSource = dt;
                        for (int i = 0; i < this.ItemNamedgvsh.Columns.Count; i++)
                        {
                            this.ItemNamedgvsh.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                            this.ItemNamedgvsh.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }
                    }
                }
                else
                {
                    ItemNamedgvsh.DataSource = null;
                }
            }
        }

        private void ItemNamesh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;

            if (!string.IsNullOrEmpty(ItemNamesh.Text.Trim()))
            {
                #region
                Encoding EncodingEN = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                string nametolading = "%" + EncodingEN.GetString(EncodingCH.GetBytes(ItemNamesh.Text.Trim())) + "%";
                string fenLei = ItemClasssh.Text == "全部" ? "%" : ItemClasssh.Text;
                #endregion
                using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                {
                    OleDbCommand cmd = new OleDbCommand(" SELECT dbo._NoLock_FS_Item.ItemNumber AS 物料代码, dbo._NoLock_FS_Item.ItemDescription AS 物料描述, dbo._NoLock_FS_Item.ItemUM AS 物料单位, dbo._NoLock_FS_ItemCost.CostType AS 类别, dbo._NoLock_FS_ItemCost.CostCode AS 方法, dbo._NoLock_FS_ItemCost.TotalRolledCost AS 累计成本合计, dbo._NoLock_FS_Item.GatewayWorkCenter AS 工作中心 FROM dbo._NoLock_FS_Item INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey WHERE (dbo._NoLock_FS_ItemCost.CostType = '0') AND (dbo._NoLock_FS_Item.ItemDescription LIKE '" + nametolading + "') and (dbo._NoLock_FS_Item.ItemNumber like '" + fenLei + "%')", conn);
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    ItemNamedgvsh.DataSource = dt;
                    for (int i = 0; i < this.ItemNamedgvsh.Columns.Count; i++)
                    {
                        this.ItemNamedgvsh.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        this.ItemNamedgvsh.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    }
                }
            }
            else
            {
                ItemNamedgvsh.DataSource = null;
            }
        }

        private void dgvBOMDetail_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, ((DataGridView)sender).RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), ((DataGridView)sender).RowHeadersDefaultCellStyle.Font, rectangle, ((DataGridView)sender).RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void GetItmbUpdate_Click(object sender, EventArgs e)
        {
            dgvItemNumber.DataSource = null;
            dgvItmbUpdateDetail.DataSource = null;
            ItmbUpdateResult.Items.Clear();
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and   PROCESSNAME='RY标准成本修改流程' and (STEPLABEL = '系统管理员维护')");

            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and   PROCESSNAME='RY增加物料申请流程' and (STEPLABEL = '系统管理员维护' or STEPLABEL = 'ERP管理员审核') and STARTTIME >'2019-5-10'");
            List<ITMBliucheng> list1 = new List<ITMBliucheng>();
            string Sqlstr = "SELECT REV_INCIDENT 流水号,ZY 摘要,REV_CREATER_NAME 申请人,REV_CREATER_DPT 申请部门,REV_CREATER_DATE 申请时间,REV_CID  FROM [dbo].[YW_RY_BZCBXG] where REV_INCIDENT=-123";
            foreach (DataRow dr in Incidents.Rows)
            {
                Sqlstr += " or REV_INCIDENT = " + dr[0].ToString();

            }

            dgvItmbUpdate.DataSource = SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, Sqlstr);
            for (int i = 0; i < this.dgvItmbUpdate.Columns.Count; i++)
            {
                this.dgvItmbUpdate.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvItmbUpdate.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void ItmbUpdate_Click(object sender, EventArgs e)//修改成本
        {
            ItmbUpdateResult.Items.Clear();
            if (dgvItmbUpdateDetail.Rows.Count == 0)
            {
                MessageBox.Show("物料成本信息为空！");
                return;
            }
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            int a = 1;
            for (int i = 0; i < dgvItmbUpdateDetail.Rows.Count; i++)
            {
                string itemcode = dgvItmbUpdateDetail["物料代码", i].Value.ToString().Trim();
                string fenlei = itemcode.Substring(0, 1).ToUpper();
                #region 修改成本
                ITMC01 myItmb = new ITMC01();
                myItmb.ItemNumber.Value = itemcode;
                myItmb.CostType.Value = "0";
                myItmb.CostCode.Value = "1";
                try
                {
                    myItmb.AtThisLevelMaterialCost.Value = Convert.ToDecimal(dgvItmbUpdateDetail["修改后成本", i].Value.ToString().Trim()).ToString("0.000000000");
                }
                catch (Exception ex)
                {
                    a = 0;
                    dgvItmbUpdateDetail["修改后成本", i].Style.BackColor = Color.Red;
                    dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                    dgvItmbUpdateDetail["库存账号", i].Style.BackColor = Color.Red;
                    MessageBox.Show(string.Format("第{0}行(修改后成本)不是数值请检查:" + ex.Message, (i + 1)));
                    continue;
                }
                //myItmb.AtThisLevelLaborCost.Value = Convert.ToDecimal(dgvBOM["人工费", i].Value.ToString().Trim()).ToString("0.000000000");
                //myItmb.AtThisLevelVariableOverheadCost.Value = Convert.ToDecimal(dgvBOM["可变间接费", i].Value.ToString().Trim()).ToString("0.000000000");
                //myItmb.AtThisLevelFixedOverheadCost.Value = Convert.ToDecimal(dgvBOM["固定间接费", i].Value.ToString().Trim()).ToString("0.000000000");

                myItmb.RolledMaterialCost.Value = Convert.ToDecimal(dgvItmbUpdateDetail["修改后成本", i].Value.ToString().Trim()).ToString("0.000000000");
                //myItmb.RolledLaborCost.Value = Convert.ToDecimal(dgvBOM["人工费", i].Value.ToString().Trim()).ToString("0.000000000");
                //myItmb.RolledVariableOverheadCost.Value = Convert.ToDecimal(dgvBOM["可变间接费", i].Value.ToString().Trim()).ToString("0.000000000");
                //myItmb.RolledFixedOverheadCost.Value = Convert.ToDecimal(dgvBOM["固定间接费", i].Value.ToString().Trim()).ToString("0.000000000");
                if (_fstiClient.ProcessId(myItmb, null))
                {
                    ItmbUpdateResult.Items.Add(string.Format("第{0}行ITMC成本修改成功!", (i + 1)));
                    ItmbUpdateResult.Items.Add(_fstiClient.CDFResponse);
                }
                else
                {
                    a = 0;
                    dgvItmbUpdateDetail["修改后成本", i].Style.BackColor = Color.Red;
                    ItmbUpdateResult.Items.Add(string.Format("第{0}行ITMC成本修改失败!", (i + 1)));
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myItmb, itemError, ItmbUpdateResult);
                }
                #endregion 修改成本
                #region 修改账号
                if (fenlei == "M")
                {
                    if (dgvItmbUpdateDetail["产品线", i].Value.ToString().Trim() + "-121100" != dgvItmbUpdateDetail["库存账号", i].Value.ToString().Trim())
                    {
                        MessageBox.Show("第" + (i + 1) + "行:M类物料库存账号不是121100");
                        dgvItmbUpdateDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                        a = 0;
                        continue;
                    }
                }
                if (fenlei == "A")
                {
                    if (dgvItmbUpdateDetail["产品线", i].Value.ToString().Trim() + "-123100" != dgvItmbUpdateDetail["库存账号", i].Value.ToString().Trim())
                    {
                        MessageBox.Show("第" + (i + 1) + "行:A类物料库存账号不是123100");
                        dgvItmbUpdateDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                        a = 0;
                        continue;
                    }
                }
                if (fenlei == "P")
                {
                    if (dgvItmbUpdateDetail["产品线", i].Value.ToString().Trim() + "-122100" != dgvItmbUpdateDetail["库存账号", i].Value.ToString().Trim())
                    {
                        MessageBox.Show("第" + (i + 1) + "行:P类物料库存账号不是122100");
                        dgvItmbUpdateDetail.Rows[i].Cells["库存账号"].Style.BackColor = Color.Red;
                        dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                        a = 0;
                        continue;
                    }
                }

                ITMC03 myItmc = new ITMC03();
                myItmc.ItemNumber.Value = dgvItmbUpdateDetail["物料代码", i].Value.ToString().Trim();
                myItmc.ProductLine.Value = dgvItmbUpdateDetail["产品线", i].Value.ToString().Trim();
                myItmc.InventoryAccount.Value = dgvItmbUpdateDetail["库存账号", i].Value.ToString().Trim();

                if (_fstiClient.ProcessId(myItmc, null))
                {
                    ItmbUpdateResult.Items.Add(string.Format("第{0}行ITMC产品线和库存账号修改成功!", (i + 1)));
                    ItmbUpdateResult.Items.Add(_fstiClient.CDFResponse);
                }
                else
                {
                    a = 0;
                    dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                    dgvItmbUpdateDetail["库存账号", i].Style.BackColor = Color.Red;
                    ItmbUpdateResult.Items.Add(string.Format("第{0}ITMC产品线和库存账号修改失败!", (i + 1)));
                    FSTIError itemError = _fstiClient.TransactionError;
                    DumpErrorObject(myItmc, itemError, ItmbUpdateResult);
                }
                #endregion 修改账号
            }
            if (a == 1)
            {
                dgvItmbUpdate.Rows[Convert.ToInt32(ItmbUpdateRowNumber.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                MessageBox.Show("全部修改成功！");
            }
            if (a == 0)
                MessageBox.Show("部分删除失败！已红色标注请检查！！！");
        }

        private void dgvItmbUpdate_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ItmbUpdateResult.Items.Clear();
            dgvItemNumber.DataSource = null;
            int rowindex = e.RowIndex;
            if (rowindex != -1)
            {
                if (dgvItmbUpdate.Rows[rowindex].DefaultCellStyle.ForeColor != Color.Red)
                {
                    dgvItmbUpdate.Rows[rowindex].DefaultCellStyle.ForeColor = Color.Blue;
                    for (int a = 0; a < dgvItmbUpdate.Rows.Count; a++)
                    {
                        if (a != rowindex && dgvItmbUpdate.Rows[a].DefaultCellStyle.ForeColor != Color.Red)
                            dgvItmbUpdate.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
                ItmbUpdateNumber.Text = dgvItmbUpdate.Rows[rowindex].Cells["流水号"].Value.ToString().Trim();
                ItmbUpdateRowNumber.Text = (rowindex + 1).ToString();
                string ParentGuid = dgvItmbUpdate.Rows[rowindex].Cells["REV_CID"].Value.ToString();

                dgvItmbUpdateDetail.DataSource = SqlHelper.ExecuteDataTable("SELECT ltrim(rtrim(WLBM))  as 物料代码,WLMS as 物料描述,DW as 单位,CB as 原成本,CPX as 产品线,KCZH as 库存账号,XGCB as 修改后成本 FROM YW_RY_BZCBXG_DH where ParentGuid = '" + ParentGuid + "'");
                for (int i = 0; i < this.dgvItmbUpdateDetail.Columns.Count; i++)
                {
                    this.dgvItmbUpdateDetail.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    this.dgvItmbUpdateDetail.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
                bool bl = true;
                for (int i = 0; i < this.dgvItmbUpdateDetail.Rows.Count; i++)
                {
                    if (dgvItmbUpdateDetail["产品线", i].Value.ToString().Trim() == "" || dgvItmbUpdateDetail["库存账号", i].Value.ToString().Trim() == "" || dgvItmbUpdateDetail["修改后成本", i].Value.ToString().Trim() == "")
                    {
                        dgvItmbUpdateDetail["产品线", i].Style.BackColor = Color.Red;
                        dgvItmbUpdateDetail["库存账号", i].Style.BackColor = Color.Red;
                        dgvItmbUpdateDetail["修改后成本", i].Style.BackColor = Color.Red;
                        bl = false;
                    }
                }
                if (bl == false)
                {
                    MessageBox.Show("有未填信息，已红色标识！");
                }

            }
        }

        private void dgvItmbUpdateDetail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            if (rowIndex > -1)
            {
                string Item = dgvItmbUpdateDetail.Rows[rowIndex].Cells[0].Value.ToString().Trim();
                //#region
                //Encoding EncodingEN = Encoding.GetEncoding("ISO-8859-1");
                //Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                //string nametolading = "%" + EncodingEN.GetString(EncodingCH.GetBytes(ItemName.Text.Trim())) + "%";
                //string fenLei = ItemClass.Text == "全部" ? "%" : ItemClass.Text;
                //#endregion
                using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                {
                    OleDbCommand cmd = new OleDbCommand(" SELECT dbo._NoLock_FS_Item.ItemNumber AS 物料代码, dbo._NoLock_FS_Item.ItemDescription AS 物料描述, dbo._NoLock_FS_Item.ItemUM AS 物料单位, dbo._NoLock_FS_ItemCost.CostType AS 类别, dbo._NoLock_FS_ItemCost.CostCode AS 方法, dbo._NoLock_FS_ProductLine.ProductLine AS 产品线,dbo._NoLock_FS_Item.InventoryAccount AS 库存账号,dbo._NoLock_FS_ItemCost.AtThisLevelMaterialCost AS 材料费,dbo._NoLock_FS_ItemCost.RolledMaterialCost AS 累计材料费, dbo._NoLock_FS_Item.GatewayWorkCenter AS 工作中心 FROM dbo._NoLock_FS_Item INNER JOIN dbo._NoLock_FS_ItemCost ON dbo._NoLock_FS_Item.ItemKey = dbo._NoLock_FS_ItemCost.ItemKey INNER JOIN dbo._NoLock_FS_ProductLine ON dbo._NoLock_FS_Item.ProductLineKey = dbo._NoLock_FS_ProductLine.ProductLineKey WHERE (dbo._NoLock_FS_ItemCost.CostType = '0') AND (dbo._NoLock_FS_Item.ItemNumber ='" + Item + "')", conn);
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    dgvItemNumber.DataSource = dt;
                    dgvItemNumber.Columns[1].Width = 360;
                    dgvItemNumber.Columns[5].Width = 160;
                    dgvItemNumber.Columns[6].Width = 160;
                }
            }
        }
        private void TemplateDownload_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();

            saveDialog.DefaultExt = "";

            saveDialog.Filter = "Excel文件|*.xlsx";

            saveDialog.FileName = "修改库管员代码模板";

            if (saveDialog.ShowDialog() != DialogResult.OK)

            {

                return;

            }


            FileStream fs = new FileStream(saveDialog.FileName, FileMode.OpenOrCreate);

            BinaryWriter bw = new BinaryWriter(fs);
            byte[] data = Resources.修改库管员代码;
            bw.Write(data, 0, data.Length);
            bw.Close();
            fs.Close();

            if (File.Exists(saveDialog.FileName))

            {

                System.Diagnostics.Process.Start(saveDialog.FileName); //打开文件

            }

        }

        private void ExcelImport_Click(object sender, EventArgs e)
        {
            try
            {
                string file = "";
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = true;//该值确定是否可以选择多个文件
                dialog.Title = "请选择文件夹";
                dialog.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    file = dialog.FileName;
                }
                else return;
                dgvUpdateStockNum.DataSource = ExcelToTable(file);
                #region 设置列宽
                for (int i = 0; i < this.dgvUpdateStockNum.Columns.Count; i++)
                {
                    this.dgvUpdateStockNum.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                    this.dgvUpdateStockNum.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public DataTable ExcelToTable(string file)
        {
            // AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            if (file == "") return null;

            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            string fileoutExt = Path.GetFileNameWithoutExtension(file).ToLower();

            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx")//新版本excel2007
                { workbook = new XSSFWorkbook(fs); }
                else if (fileExt == ".xls")//早期版本excel2003
                { workbook = new HSSFWorkbook(fs); }
                else { workbook = null; }
                if (workbook == null) { return null; }
                ISheet sheet = workbook.GetSheetAt(0);//下标为零的工作簿
                //创建表头      FirstRowNum:获取第一个有数据的行好(默认0)
                IRow header = sheet.GetRow(sheet.FirstRowNum);//第一行是头部信息
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)//LastCellNum 获取列的条数
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));//如果excel没有列头就自定义
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));//获取excel列头
                    columns.Add(i);
                }
                //构建数据   sheet.FirstRowNum + 1 表示去掉列头信息
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)//LastRowNum最后一条数据的行号
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;//判断是否有值
                    foreach (int j in columns)
                    {
                        //if (sheet.GetRow(i) == null) continue;//如果没数据
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)//判断至少一列有值
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }
        private static object GetValueType(ICell cell)
        {

            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Formula: //BOOLEAN: 
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                case CellType.Numeric: //NUMERIC:  
                    if (DateUtil.IsCellDateFormatted(cell))//判断是否日期
                        return cell.DateCellValue.ToString("yyyy/MM/dd");
                    else
                        return cell.NumericCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.String: //STRING:  
                default:
                    return cell.StringCellValue;

            }
        }

        private void ExportStock_Click(object sender, EventArgs e)
        {
            if (dgvUpdateStockNum.Rows.Count == 0)
            { MessageBox.Show("无数据！"); return; }

            string filePath = getExcelpath();
            if (filePath.IndexOf(":") < 0)
            { return; }
            TableToExcel(dgvUpdateStockNum, filePath);
            MessageBox.Show("导出完成");
        }
        private static string getExcelpath()
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xlsx";
            saveDialog.Filter = "EXCEL表格|*.xlsx";
            //saveDialog.FileName = "条形码";
            saveDialog.ShowDialog();
            return saveDialog.FileName;
        }
        public static void TableToExcel(DataGridView dt, string file)
        {
            //AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx")
            { workbook = new XSSFWorkbook(); }
            else if (fileExt == ".xls")
            { workbook = new HSSFWorkbook(); }
            else { workbook = null; }
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.Name) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.Name);

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].HeaderText);
            }

            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i].Cells[j].Value.ToString());
                }
            }

            //转为字节数组  
            MemoryStream stream = new MemoryStream();//读写内存的对象
            workbook.Write(stream);
            var buf = stream.ToArray();//字节数组
            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();//缓冲区在内存中有个临时区域  盆 两个水缸 //缓冲区装满才会自动提交
            }
        }

        private void BtUpdateStockNum_Click(object sender, EventArgs e)
        {
            if (dgvUpdateStockNum.Rows.Count == 0)
            {
                MessageBox.Show("无内容！");
                return;
            }
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            Thread tread = new Thread(UpdateStockNum);
            tread.IsBackground = true;//变为后台程序，随主窗体结束线程
            tread.Start(dgvUpdateStockNum);
        }

        private void UpdateStockNum(object dgv0)
        {
            DataGridView dgv = (DataGridView)dgv0;
            int a = 1;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string itemcode = dgv["物料编码", i].Value.ToString().Trim().ToUpper();
                this.Invoke((EventHandler)delegate
                {
                    RowIndexIntime.Text = (i + 1).ToString() + "  " + itemcode;
                });
                ITMB02 myItmb = new ITMB02();
                myItmb.ItemNumber.Value = itemcode;
                myItmb.ItemReference3.Value = dgv["库管员代码", i].Value.ToString().Trim();

                if (_fstiClient.ProcessId(myItmb, null))
                {
                    //referenceResult.Items.Add(itemcode + ":参考字段1修改成功");
                    dgv["状态", i].Value = "1";
                }
                else
                {
                    a = 0;

                    FSTIError itemError = _fstiClient.TransactionError;
                    dgv["状态", i].Value = "0:" + itemError.Description;
                }
                //break;
            }
            if (a == 1)
                MessageBox.Show("全部修改成功！");
            if (a == 0)
                MessageBox.Show("部分删除失败！请检查！！！！！！");
        }

        private void dgvUpdateStockNum_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Rectangle rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, ((DataGridView)sender).RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), ((DataGridView)sender).RowHeadersDefaultCellStyle.Font, rectangle, ((DataGridView)sender).RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void BtUpdateZikehu_Click(object sender, EventArgs e)
        {
            if (dgvUpdateStockNum.Rows.Count == 0)
            {
                MessageBox.Show("无内容！");
                return;
            }
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            CbChildCustomer.Checked = true;
            Thread tread = new Thread(UpdateZikehu);
            tread.IsBackground = true;//变为后台程序，随主窗体结束线程
            tread.Start(dgvUpdateStockNum);
        }

        private void UpdateZikehu(object dgv0)
        {


            DataGridView dgv = (DataGridView)dgv0;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                linshiindex = i;
                AddCustomer.Enabled = false;
                string CustomerCode = dgv["单位编码", i].Value.ToString().Trim().ToUpper();
                tbCustomerCode.Text = CustomerCode;
                this.Invoke((EventHandler)delegate
                {
                    RowIndexIntime.Text = (i + 1).ToString() + "  " + CustomerCode;
                });
                if (CustomerCode.Length == 7)//代码长度为7位，表示添加的为子公司，需要查询主公司的信息
                {
                    

                    string ParentCustomerCode = CustomerCode.Substring(0, 6);
                    string strsql = "select CustomerName,CustomerAddress1,CustomerAddress2,CustomerZip,CustomerContact,CustomerContactPhone,CustomerContactFax,AccountingContact, AccountingContactPhone,AccountingContactFax,CSR,SalesRegion,TradeClassName,CustomerControllingCode,CustomerCurrencyCode, BankReference1,BankReference2,FederalPrimaryTaxExemptCertificateNumber  from _NoLock_FS_Customer where CustomerID = '" + ParentCustomerCode + "'";
                    using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                    {
                        using (OleDbCommand cmd = new OleDbCommand(strsql, conn))
                        {
                            conn.Open();
                            OleDbDataAdapter oledbDA = new OleDbDataAdapter(cmd);
                            DataTable dt1 = new DataTable();
                            oledbDA.Fill(dt1);
                            conn.Close();
                            if (dt1.Rows.Count == 1)
                            {
                                DataRow myDR = dt1.Rows[0];
                                tbCustomerName.Text = myDR["CustomerName"].ToString();
                                tbCustAddress.Text = myDR["CustomerAddress1"].ToString() + myDR["CustomerAddress2"].ToString().Trim();
                                tbContactPerson.Text = myDR["CustomerContact"].ToString();
                                tbContactTelephone.Text = myDR["CustomerContactPhone"].ToString();
                                tbContactFax.Text = myDR["CustomerContactFax"].ToString();
                                tbPostcode.Text = myDR["CustomerZip"].ToString();
                                cbIndustry.Text = myDR["TradeClassName"].ToString();
                                tbSalesmanName.Text = myDR["SalesRegion"].ToString();
                                tbSalesmanCode.Text = myDR["CSR"].ToString();
                                tbAccountantName.Text = myDR["AccountingContact"].ToString();
                                tbAccountantPhone.Text = myDR["AccountingContactPhone"].ToString();
                                cbMoney.Text = myDR["CustomerControllingCode"].ToString() == "L" ? "本币" : "外币";
                                tbCustomerCurrencyCode.Text = myDR["CustomerCurrencyCode"].ToString();
                                tbBankOfDeposit.Text = myDR["BankReference1"].ToString();
                                tbBankAccount.Text = myDR["BankReference2"].ToString();
                                tbTaxCode.Text = myDR["FederalPrimaryTaxExemptCertificateNumber"].ToString();

                                AddCustomer.Enabled = true;
                                string Customercode1 = tbCustomerCode.Text.Trim();
                                tbUniteAccount.Text = "1A" + Customercode1.Substring(6, 1) + Customercode1.Substring(0, 1) + "-" + Customercode1.Substring(1, 2) + "-" + Customercode1.Substring(3, 3);
                            }
                            if (dt1.Rows.Count == 0)
                            {
                                foreach (Control ct in groupBox7.Controls)
                                {
                                    if (ct is TextBox)
                                        ct.Text = "";
                                }
                                tbCustomerCode.Text = CustomerCode;
                                dgv["状态", i].Value = "没有该客户！";
                                continue;
                            }
                        }

                    }
                    tBCustBZ.Text = "";
                    customercode.Text = "customercode";
                    customerliushuihao.Text = " liushuihao";
                    customerhanghao.Text = " hanghao";
                }
                else
                {

                    dgv["状态", i].Value = "子客户编码不是7位";
                    continue;
                }
                #region  修改业务员姓名及编码
                tbSalesmanName.Text = dgv["业务员", i].Value.ToString().Trim().ToUpper();
                tbSalesmanCode.Text = dgv["业务员编码", i].Value.ToString().Trim().ToUpper();
                #endregion


                addZikehu(i);



            }
            linshiindex = -100;
        }

        private void addZikehu(int index)
        {
            if (tbCustomerCode.Text.ToString() == "" || tbCustomerName.Text.ToString() == "" || tbCustAddress.Text.ToString() == "" || tbSalesmanCode.Text.ToString() == "" || tbSalesmanName.Text.ToString() == "")
            {
                MessageBox.Show("客户信息不完整，无法添加！ ");
                return;
            }
            tbCustomerCurrencyCode.Text = tbCustomerCurrencyCode.Text.Trim();
            string CCode = tbCustomerCurrencyCode.Text.Trim();
            if (cbIndustry.Text == "" || cbMoney.Text == "")
            { MessageBox.Show("请检查行业类别|货币类型！"); return; }
            if ((cbMoney.Text == "本币" && tbCustomerCurrencyCode.Text != "00000") || (cbMoney.Text == "外币" && tbCustomerCurrencyCode.Text != "USD" && tbCustomerCurrencyCode.Text != "EURO"))
            { MessageBox.Show("请检查货币类型|货币代码是否对应！"); return; }
            if (StrLength(tbBankOfDeposit.Text.Trim()) > 30)
            {
                MessageBox.Show("开户银行超出30个字符，请调整！ ");
                return;
            }
            if (StrLength(tbBankAccount.Text.Trim()) > 30)
            {
                MessageBox.Show("开户银行超出30个字符，请调整！ ");
                return;
            }
            if (StrLength(tbCustAddress.Text.Trim()) > 60 && cbMoney.Text == "本币")
            {
                MessageBox.Show("内销客户地址超出60个字符，请调整！ ");
                return;
            }
            if (StrLength(tbCustAddress.Text.Trim()) > 120 && cbMoney.Text == "外币")
            {
                MessageBox.Show("外贸客户地址超出120个字符，请调整！ ");
                return;
            }

            if (StrLength(tbTaxCode.Text.Trim()) > 20)
            {
                MessageBox.Show("税号超出20个字符，请检查！ ");
                return;
            }
            if (StrLength(tbProvince.Text.Trim()) > 10)
            {
                MessageBox.Show("客户所在省份超出10个字符，请调整！ ");
                return;
            }
            if (StrLength(tbUniteAccount.Text.Trim()) != 11)
            {
                MessageBox.Show("合并账号不是11位，请检查！ ");
                return;
            }
            #region 检查客户名称是否重复
            using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
            {
                Encoding EncodingLD = Encoding.GetEncoding("ISO-8859-1");
                Encoding EncodingCH = Encoding.GetEncoding("GB2312");
                string CustomerName = EncodingLD.GetString(EncodingCH.GetBytes(tbCustomerName.Text.Trim()));
                SqlCommand cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "' and CustomerID not like'" + tbCustomerCode.Text.Trim().Substring(0, 6) + "%'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count > 0)
                {
                    MessageBox.Show("有相同客户名称的记录，请检查" + dtcust.Rows[0][0].ToString());
                    return;
                }

            }
            #endregion
            #region 检查税号是否重复
            if (!string.IsNullOrEmpty(tbTaxCode.Text.Trim()))
            {
                using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                {
                    string CustomerName = tbTaxCode.Text.Trim().Replace(" ", "");
                    SqlCommand cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where FederalPrimaryTaxExemptCertificateNumber = '" + CustomerName + "' and CustomerID not like'" + tbCustomerCode.Text.Trim().Substring(0, 6) + "%'", conn);
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dtcust = new DataTable();
                    sda.Fill(dtcust);
                    if (dtcust.Rows.Count > 0)
                    {
                        MessageBox.Show("有相同税号的记录，请检查" + dtcust.Rows[0][0].ToString());
                        dgvUpdateStockNum["状态", index].Value = "相同税号:"+dtcust.Rows[0][0].ToString();
                        return;
                    }

                }
            }
            #endregion

            listBoxCustomer.Items.Clear();
            AddCustomer.Enabled = false;
            if (CbChildCustomer.Checked)
            {
                //int i = 0;
                //using (OleDbConnection conn = new OleDbConnection(SqlHelper.FSDBMRSQLOLEDB))
                //{
                //    using (OleDbCommand cmd = new OleDbCommand(@"select CustomerName,CustomerAddress1,CustomerContact,CustomerContactPhone,CustomerContactFax,AccountingContact,              AccountingContactPhone,AccountingContactFax,CustomerControllingCode,CustomerCurrencyCode, BankReference1,BankReference2,FederalPrimaryTaxExemptCertificateNumber  from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn))
                //    {
                //        conn.Open();
                //        DataTable dt = new DataTable();
                //        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                //        da.Fill(dt);
                //        i = dt.Rows.Count;
                //        conn.Close();
                //    }

                //}
                //if (i == 0)
                listBoxCustomer.Items.Add("开始增加子客户------->-------->------->-------->");

                if (AddChildCustomer() == true)
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户成功<<<");
                    //MessageBox.Show("子客户添加成功！");
                    dgvUpdateStockNum["状态", index].Value = "1";
                    #region 检查是否有重复的客户名称
                    using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                    {
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("select CustomerName from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn);
                        object vendornamelading = cmd.ExecuteScalar();
                        if (vendornamelading == null)
                        { MessageBox.Show("没有该客户的记录!"); }
                        else
                        {
                            string CustomerName = vendornamelading.ToString();
                            cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "'", conn);
                            SqlDataAdapter sda = new SqlDataAdapter(cmd);
                            DataTable dtcust = new DataTable();
                            sda.Fill(dtcust);
                            foreach (DataRow dr in dtcust.Rows)
                            {
                                if (dr["CustomerID"].ToString().Trim().Substring(0, 6) != tbCustomerCode.Text.Trim().Substring(0, 6))
                                {
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                }
                            }
                        }
                    }
                    #endregion
                    #region 客户信息groupBox7信息清空
                    foreach (Control control in groupBox7.Controls)
                    {
                        if (!(control is Label))
                        {
                            control.Text = null;
                        }
                    }
                    #endregion
                    try
                    {
                        dgvCustomer.Rows[Convert.ToInt32(customerhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    catch (Exception)
                    {

                    }

                }
                else
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户失败<<<");
                    MessageBox.Show(" 子客户添加失败！");
                }
            }
            else
            {
                listBoxCustomer.Items.Add("开始增加主客户-->");
                if (AddCustomerCompany())
                {
                    listBoxCustomer.Items.Add("-->主客户录入成功");

                }
                else
                {
                    listBoxCustomer.Items.Add("-->主客户录入失败");
                    MessageBox.Show("主客户录入失败！请检查！！！");
                    return;
                }
            
                listBoxCustomer.Items.Add("开始增加子客户------->-------->------->-------->");

                if (AddChildCustomer() == true)
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户成功<<<");
                    //MessageBox.Show("子客户添加成功！");
                    dgvUpdateStockNum["状态", index].Value = "1";
                    #region 检查是否有重复的客户名称
                    using (SqlConnection conn = new SqlConnection(SqlHelper.FSDBMRSQL))
                    {
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("select CustomerName from _NoLock_FS_Customer where CustomerID = '" + tbCustomerCode.Text.Trim().Substring(0, 6) + "'", conn);
                        object vendornamelading = cmd.ExecuteScalar();
                        if (vendornamelading == null)
                        { MessageBox.Show("没有该客户的记录!"); }
                        else
                        {
                            string CustomerName = vendornamelading.ToString();
                            cmd = new SqlCommand("select CustomerID from _NoLock_FS_Customer where CustomerName = '" + CustomerName + "'", conn);
                            SqlDataAdapter sda = new SqlDataAdapter(cmd);
                            DataTable dtcust = new DataTable();
                            sda.Fill(dtcust);
                            foreach (DataRow dr in dtcust.Rows)
                            {
                                if (dr["CustomerID"].ToString().Trim().Substring(0, 6) != tbCustomerCode.Text.Trim().Substring(0, 6))
                                {
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                    MessageBox.Show("有多个相同客户名称的记录，请检查！");
                                }
                            }
                        }
                    }
                    #endregion
                    #region 客户信息groupBox7信息清空
                    foreach (Control control in groupBox7.Controls)
                    {
                        if (!(control is Label))
                        {
                            control.Text = null;
                        }
                    }
                    #endregion
                    try
                    {
                        dgvCustomer.Rows[Convert.ToInt32(customerhanghao.Text) - 1].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    catch (Exception)
                    {

                    }

                }
                else
                {

                    listBoxCustomer.Items.Add("----------->---------->----------->增加子客户失败<<<");
                    MessageBox.Show(" 子客户添加失败！");
                }

            }
            
        }

        private void ZikehuTemplateDownload_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();

            saveDialog.DefaultExt = "";

            saveDialog.Filter = "Excel文件|*.xlsx";

            saveDialog.FileName = "增加子客户模板";

            if (saveDialog.ShowDialog() != DialogResult.OK)

            {

                return;

            }


            FileStream fs = new FileStream(saveDialog.FileName, FileMode.OpenOrCreate);

            BinaryWriter bw = new BinaryWriter(fs);
            byte[] data = Resources.增加子客户模板;
            bw.Write(data, 0, data.Length);
            bw.Close();
            fs.Close();

            if (File.Exists(saveDialog.FileName))

            {

                System.Diagnostics.Process.Start(saveDialog.FileName); //打开文件

            }
        }

        private void btnGetPONumber_Click(object sender, EventArgs e)
        {
            if (dgvUpdateStockNum.Rows.Count == 0)
            {
                MessageBox.Show("无内容！");
                return;
            }
            Thread tread = new Thread(GetPONumber);
            tread.IsBackground = true;//变为后台程序，随主窗体结束线程
            tread.Start(dgvUpdateStockNum);
        }

        private void GetPONumber(object dgv0)
        {
            DataGridView dgv = (DataGridView)dgv0;
            SqlConnection conn = new SqlConnection(SqlHelper.FSDBSQL);
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {

                cmd.CommandText = "SELECT PONumber from (SELECT PONumber,sum(case when ReceiptQuantity is null then 0-ReversedQuantity else ReceiptQuantity end) Quantity FROM [dbo].[PORV] where ItemNumber='"+dgv["物料代码",i].Value.ToString().Trim()+ "' and LotNumber = '" + dgv["批号", i].Value.ToString().Trim() + "' Group BY PONumber) as t1 where t1.Quantity>0";
                dgv["采购单号", i].Value=cmd.ExecuteScalar();
            }
            conn.Close();
            MessageBox.Show("匹配完成");
        }


        private void GetCustomerprocessNotYet_Click(object sender, EventArgs e)
        {
            #region 客户信息groupBox7信息清空
            foreach (Control control in groupBox7.Controls)
            {
                if (!(control is Label))
                {
                    control.Text = null;
                }
            }
            #endregion
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 3 and PROCESSNAME='RY开户审批流程' and TASKUSER='BPM/cuiqingjuan' and ENDTIME >'2019/12/13' and STEPLABEL='ERP管理员新增客户'");
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY开户审批流程' and TASKUSER='BPM/cuiqingjuan' and STEPLABEL='ERP管理员新增客户'");
            DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME ='RY开户审批流程'  and STEPLABEL !='ERP管理员新增客户'");//未完成且不属于当前环节的流程
            //DataTable Incidents = SqlHelper1.ExecuteDataTable(SqlHelper.ultimusSQL, "SELECT INCIDENT FROM [dbo].[TASKS] where STATUS = 1 and PROCESSNAME='RY开户审批流程' ");//所有未完成流程
            string cmdstr = @"SELECT REV_INCIDENT AS 流水号, REV_CREATER_NAME	AS 发起人, REV_CREATER_DPT	AS 发起部门, REV_CREATER_TEL  AS 联系电话, (CASE   WHEN XGKHMC <> '' THEN '客户信息修改'  ELSE '新开户'  END ) AS 类型, KHBM	AS 客户代码, YKHMC AS 原客户名称, KHMC	AS 客户名称, KHDZ	AS 客户地址,(CASE   WHEN KHLB = 'yy' THEN '医院'  ELSE (CASE   WHEN KHLB = 'yd' THEN '药店'  ELSE (CASE   WHEN KHLB = 'mz' THEN '门诊'  ELSE  (CASE   WHEN KHLB = 'gs' or  KHLB = 'yc' THEN '公司'  ELSE   '其他'   END )    END )     END )     END )	AS 公司类型, MS	AS 合并账户, ZKHB	AS 货币类型, YZBM	AS 邮编, DH	AS 电话, CZ	AS 传真,  KHYH	AS 开户银行, ZH	AS 银行账户, SH	AS 税号, ZJL	AS 总经理, YXJL	AS 销售经理, ZB	AS 主办, CWJL	AS 财务经理, YWY	AS 业务员, YWYH	AS 业务代码, KJ	AS 会计, KJDH	AS 会计电话,BZ as 备注,KHSZSF as 客户所在省份 FROM YW_KHSP where REV_INCIDENT=-12345";
            //string cmdstr = "SELECT * FROM YW_KHSP where REV_INCIDENT=" + Incidents.Rows[0][0];

            for (int i = 0; i < Incidents.Rows.Count; i++)
            {
                cmdstr += " or REV_INCIDENT=" + Incidents.Rows[i][0];
            }

            dgvCustomer.DataSource = SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, cmdstr);
            for (int i = 0; i < this.dgvCustomer.Columns.Count; i++)
            {
                this.dgvCustomer.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvCustomer.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        private void NumForSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
         
            List<ITMBliucheng> list1 = new List<ITMBliucheng>();
                ITMBliucheng Vendor1 = TolistITMB(SqlHelper1.ExecuteDataTable(SqlHelper.UltimusBusinessSQL, "SELECT * FROM [dbo].[YW_ZJWLCB] where REV_INCIDENT=" + NumForSearch.Text.Trim()));
                list1.Add(Vendor1);

            dgvItmb.DataSource = list1;
            for (int i = 0; i < this.dgvItmb.Columns.Count; i++)
            {
                this.dgvItmb.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                this.dgvItmb.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            ITMBResult.Items.Clear();
        }

        private void BtnUpdateLotNumberMask_Click(object sender, EventArgs e)
        {
            if (dgvUpdateStockNum.Rows.Count == 0)
            {
                MessageBox.Show("无内容！");
                return;
            }
            if (toolStripStatusLabel1.Text == "未登录" || "ID:" + _fstiClient.UserId != toolStripStatusLabel1.Text)
            {
                MessageBox.Show("请登录四班账号！");
                return;
            }
            Thread tread = new Thread(UpdateLotNumberMask);
            tread.IsBackground = true;//变为后台程序，随主窗体结束线程
            tread.Start(dgvUpdateStockNum);
        }
        private void UpdateLotNumberMask(object dgv0)
        {
            DataGridView dgv = (DataGridView)dgv0;
            int a = 1;
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string itemcode = dgv["ItemNumber", i].Value.ToString().Trim().ToUpper();
                this.Invoke((EventHandler)delegate
                {
                    RowIndexIntime.Text = (i + 1).ToString() + "  " + itemcode;
                });
                ITMB07 itmb07 = new ITMB07();
                itmb07.ItemNumber.Value = itemcode;
                itmb07.LotNumberMask.Value = "XXXXXXXXXXXXXXXXXXXX";

                if (_fstiClient.ProcessId(itmb07, null))
                {
                    dgv["状态", i].Value = "1";

                }
                else
                {
                    a = 0;

                    FSTIError itemError = _fstiClient.TransactionError;
                    dgv["状态", i].Value = "0:" + itemError.Description;
                }
            }
            if (a == 1)
                MessageBox.Show("全部修改成功！");
            if (a == 0)
                MessageBox.Show("部分删除失败！请检查！！！！！！");
        }
    }
}
