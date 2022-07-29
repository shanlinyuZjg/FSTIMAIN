using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SoftBrands.FourthShift.Transaction;

namespace FSTIMAIN
{
    public partial class AddSubCustomer : Form
    {
        private readonly string FSDBMRSQLOLEDB = "Provider=SQLOLEDB.1; Data Source=192.168.8.11;Initial Catalog=FSDBMR;User ID=program;PassWord=program;Auto Translate=False;";
        private readonly string FSDBMRSQL = "Data Source=192.168.8.11;database=FSDBMR;uid=program;pwd=program";
        private readonly string FSDBSQL = "Data Source=192.168.8.11;database=FSDB;uid=xym;pwd=xym-123";
        private string UserId;
        private string UserName;
        private string UserDept;
        private FSTIClient _fstiClient = null;//声明FSTIClient类的对象
        private string ConfigFile = @"m:\mfgsys\fs.cfg";
        //private string ConfigFile = @"T:\mfgsys\fs.cfg";
        private string Log;
        private string LogStatus;
        private string SubCustomerCode;
        private string SalesmanName;
        private string SalesmanCode;
        //public AddSubCustomer()
        //{
        //    InitializeComponent();
        //}
        /// <summary>
        /// 增加子客户3参数构造函数
        /// </summary>
        /// <param name="userId">用户账号</param>
        /// <param name="userName">用户姓名</param>
        /// <param name="userDept">用户所属部门</param>
        public AddSubCustomer(string userId,string userName,string userDept)
        {
            InitializeComponent();
            UserId = userId;
            UserName = userName;
            UserDept = userDept;
        }
        private void AddSubCustomer_Load(object sender, EventArgs e)
        {

        }
        private void tbCustomerCode_KeyDown(object sender, KeyEventArgs e)//通过客户代码获得客户信息
        {
            if (e.KeyCode != Keys.Enter)
                return;
            btnAddSubCustomer.Tag = null;
            //定义简体中文和西欧文编码字符集
            //Encoding GB2312 = Encoding.GetEncoding("gb2312");
            //Encoding ISO88591 = Encoding.GetEncoding("iso-8859-1");
            string CustomerCode = tbCustomerCode.Text.ToString().Trim().ToUpper();
            tbCustomerCode.Text = CustomerCode;
            int i = CustomerCode.Length;//获得客户代码的长度 
            if (i == 7)//代码长度为7位，表示添加的为子公司，需要查询主公司的信息
            { 
                string ParentCustomerCode = CustomerCode.Substring(0, 6);
                string strsql = "select CustomerName,CustomerAddress1,CustomerAddress2,CustomerZip,CustomerContact,CustomerContactPhone,CustomerContactFax,AccountingContact, AccountingContactPhone,AccountingContactFax,CSR,SalesRegion,TradeClassName,CustomerControllingCode,CustomerCurrencyCode, BankReference1,BankReference2,FederalPrimaryTaxExemptCertificateNumber  from _NoLock_FS_Customer where CustomerID = '" + ParentCustomerCode + "'";
                using (OleDbConnection conn = new OleDbConnection(FSDBMRSQLOLEDB))
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
                            if (tbCustomerName.Text.Contains("#"))
                            { MessageBox.Show("现款户不能添加子客户");return; }
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

                            string Customercode1 = tbCustomerCode.Text.Trim();
                            tbUniteAccount.Text = "1A" + Customercode1.Substring(6, 1) + Customercode1.Substring(0, 1) + "-" + Customercode1.Substring(1, 2) + "-" + Customercode1.Substring(3, 3);

                            btnAddSubCustomer.Tag = CustomerCode;
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
            }
            else
            {
                foreach (Control ct in groupBox7.Controls)
                {
                    if (ct is TextBox)
                        ct.Text = "";
                }
                tbCustomerCode.Text = CustomerCode;
                MessageBox.Show("子客户编码不是7位！");
            }
        }

        private void btnAddSubCustomer_Click(object sender, EventArgs e)
        {
            
            #region  数据检查
            tbCustomerCode.Text = tbCustomerCode.Text.ToString().Trim().ToUpper();
            if (btnAddSubCustomer.Tag != null)
            {
                if (tbCustomerCode.Text != btnAddSubCustomer.Tag.ToString())
                {
                    MessageBox.Show("子客户不匹配！");
                    return;
                }
            }
            else
            {
                MessageBox.Show("子客户不匹配！");
                return;
            }
            //日志初始化
            Log = string.Empty;
            LogStatus = string.Empty;

            if (tbCustomerCode.Text.ToString().Trim() == "" || tbCustomerName.Text.ToString().Trim() == "" || tbCustAddress.Text.ToString().Trim() == "" || tbSalesmanCode.Text.ToString().Trim() == "" || tbSalesmanName.Text.ToString().Trim() == "")
            {
                MessageBox.Show("客户信息不完整，无法添加！ ");
                return;
            }
            tbCustomerCurrencyCode.Text = tbCustomerCurrencyCode.Text.Trim();
            string CCode = tbCustomerCurrencyCode.Text;
            if (cbIndustry.Text == "" || cbMoney.Text == "")
            { MessageBox.Show("请检查行业类别|货币类型！"); return; }
            if ((cbMoney.Text == "本币" && tbCustomerCurrencyCode.Text != "00000") || (cbMoney.Text == "外币" && tbCustomerCurrencyCode.Text != "USD" && tbCustomerCurrencyCode.Text != "EURO" && tbCustomerCurrencyCode.Text != "CHF") || string.IsNullOrWhiteSpace(cbMoney.Text) || string.IsNullOrWhiteSpace(tbCustomerCurrencyCode.Text))
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
            using (SqlConnection conn = new SqlConnection(FSDBMRSQL))
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
                using (SqlConnection conn = new SqlConnection(FSDBMRSQL))
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
            #endregion 数据检查
            #region  FSTI四班账号登录退出及增加子客户
            try
            {
                
                try
                {
                    if (_fstiClient != null)
                    {
                        _fstiClient.Terminate();
                        _fstiClient = null;
                    }
                    _fstiClient = new FSTIClient();

                    // call InitializeByConfigFile
                    // second parameter == true is to participate in unified logon
                    // third parameter == false, no support for impersonation is needed

                    _fstiClient.InitializeByConfigFile(ConfigFile, true, false);
                    //MessageBox.Show(_fstiClient.UserId);
                    // Since this program is participating in unified logon, need to
                    // check if a logon is required.

                    if (_fstiClient.IsLogonRequired)
                    {
                        // Logon is required, enable the logon button
                        string message = null;     // used to hold a return message, from the logon
                        int status;         // receives the return value from the logon call
                        status = _fstiClient.Logon("FSTI", "fst1123", ref message);
                        if (status > 0)
                        {
                            MessageBox.Show("Invalid user id or password;无效的四班用户名或密码");
                            return;
                        }
                    }
                    else
                    {
                        if (_fstiClient.UserId != "FSTI")
                        {
                            MessageBox.Show("当前登录用户不是FSTI，请退出四班客户端，再次尝试。");
                            if (_fstiClient != null)
                            {
                                _fstiClient.Terminate();
                                _fstiClient = null;
                            }
                            return;
                        }
                    }
                }
                catch (FSTIApplicationException exception)
                {
                    MessageBox.Show("四班初始化或登录失败：" + exception.Message, "FSTIApplication Exception");
                    if (_fstiClient != null)
                    {
                        _fstiClient.Terminate();
                        _fstiClient = null;
                    }
                    return;
                }
                Log += "基础信息:" + tbCustomerCode.Text + "  " + tbSalesmanName.Text.Trim() + "  " + tbSalesmanCode.Text.Trim() + ";";
                SubCustomerCode = tbCustomerCode.Text;
                SalesmanName = tbSalesmanName.Text.Trim();
                SalesmanCode = tbSalesmanCode.Text.Trim().ToUpper();
                if (AddChildCustomer())
                {
                    MessageBox.Show(tbCustomerCode.Text+"增加子客户成功");
                    LogStatus = "成功";
                    label2.Text = tbCustomerCode.Text + "增加子客户成功";
                    #region 客户信息groupBox7信息清空
                    foreach (Control control in groupBox7.Controls)
                    {
                        if (!(control is Label))
                        {
                            control.Text = null;
                        }
                    }
                    btnAddSubCustomer.Tag = null;
                    #endregion
                }
                else
                {
                    MessageBox.Show(tbCustomerCode.Text + "增加子客户失败");
                    LogStatus = "失败";
                    label2.Text = tbCustomerCode.Text + "增加子客户失败";
                }
            }
            catch
            {

            }
            finally
            {
                if (_fstiClient != null)
                {
                    _fstiClient.Terminate();
                    _fstiClient = null;
                }
            }
            #endregion FSTI四班账号登录退出及增加子客户

            #region  日志
            if (LogStatus == "成功" || LogStatus == "失败")
            {
                string SQLstr = "INSERT INTO [Log_AddSubCustomer]( [UserID], [UserName], [UserDept], [Log], [LogStatus],  [SubCustomerCode], [SalesmanName], [SalesmanCode]) VALUES ( @UserID, @UserName, @UserDept, @Log, @LogStatus,  @SubCustomerCode, @SalesmanName, @SalesmanCode)";
                SqlParameter[] para = { new SqlParameter("@UserID",UserId),
                new SqlParameter("@UserName",UserName),
                new SqlParameter("@UserDept",UserDept),
                new SqlParameter("@Log",Log),
                new SqlParameter("@LogStatus",LogStatus),
                new SqlParameter("@SubCustomerCode",SubCustomerCode),
                new SqlParameter("@SalesmanName",SalesmanName),
                new SqlParameter("@SalesmanCode",SalesmanCode)
                };
                if (ExecuteNonQuery(SQLstr, para) == 1)
                {
                    //MessageBox.Show("添加日志成功"); 
                }
                else
                { MessageBox.Show("添加日志失败，请联系软件服务处"); }
            }
            #endregion 日志
        }
        private int ExecuteNonQuery(string cmdText, params SqlParameter[] para)
        {
            SqlConnection conn = new SqlConnection(FSDBSQL);
            conn.Open();
            SqlCommand cmd = new SqlCommand(cmdText, conn);
            cmd.Parameters.AddRange(para);
            return cmd.ExecuteNonQuery();
        }
        private bool AddChildCustomer()//程序增加子客户
        {
            #region 增加GLOS GLAV
            ADDGLOS(tbUniteAccount.Text.Trim(), tbCustomerName.Text.Trim());
            ADDGLAV(tbUniteAccount.Text.Trim(), "113100");
            using (SqlConnection conn = new SqlConnection(FSDBMRSQL))
            {
                SqlCommand cmd = new SqlCommand("SELECT * FROM [dbo].[_NoLock_FS_GLAccountOrganizationValidation] where GLAccountNumber='113100' and GLOrganization='" + tbUniteAccount.Text.Trim() + "'", conn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dtcust = new DataTable();
                sda.Fill(dtcust);
                if (dtcust.Rows.Count == 0)
                {
                    Log += tbUniteAccount.Text.Trim() + "-113100组织对应账号不存在;";
                    MessageBox.Show(tbUniteAccount.Text.Trim() + "-113100 GLOS GLAV添加失败，子客户未添加，请检查");
                    return false;
                }
            }
            #endregion
            #region 录入子客户主题信息
            string CustomerCode = tbCustomerCode.Text.Trim().ToUpper().Substring(0, 7);//子客户代码
            //添加客户基础信息
            SOPC00 myCustomerBasic = new SOPC00();
            myCustomerBasic.CustomerID.Value = CustomerCode;//子客户代码
            myCustomerBasic.CustomerName.Value = tbCustomerName.Text.Trim();//子客户名称
            myCustomerBasic.CustomerLevel.Value = "C";//客户是否为主公司，p为主公司,c为子客户
            myCustomerBasic.TradeClassName.Value = cbIndustry.Text.Trim();//行业类别
            myCustomerBasic.ParentCustomer.Value = CustomerCode.Substring(0, 6);//主客户
            if (_fstiClient.ProcessId(myCustomerBasic, null))
            {
                Log+="子客户新建成功:"+_fstiClient.CDFResponse+";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "子客户新建失败:" + itemError.Description + ";";
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
                Log += "子客户财务应收账款及税金添加成功:" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "子客户财务应收账款及税金添加失败:" + itemError.Description + ";";
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
                Log += "子客户地址添加成功:" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "子客户地址添加失败:" + itemError.Description + ";";
                return false;
            }

            //概要及信用
            SOPC03 myCustomerSales = new SOPC03();
            myCustomerSales.CustomerID.Value = CustomerCode;//子客户代码
            myCustomerSales.SalesRegion.Value = tbSalesmanName.Text.ToString().Trim();//销售地区
            myCustomerSales.CSR.Value = tbSalesmanCode.Text.ToString().Trim().ToUpper();//客户服务代表字段
            myCustomerSales.CustomerState.Value = "A";
            if (cbMoney.Text.Trim() == "本币")
            {
                myCustomerSales.CreditLimitControllingAmount.Value = "1.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "Y";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "L";
                myCustomerSales.CustomerCurrencyCode.Value = "00000";

            }
            else
            {
                //美元（USD）、欧元(EURO)、瑞士法郎（CHF） 本币（00000）
                myCustomerSales.CreditLimitControllingAmount.Value = "0.00";//信用额度总值,此处不用带RMB即可，系统根据采用的货币区域自动添加，实际存储时没有货币代码
                myCustomerSales.IsCustomerOnCreditHold.Value = "N";//客户信用策略冻结
                myCustomerSales.CustomerControllingCode.Value = "F";
                myCustomerSales.CustomerCurrencyCode.Value = tbCustomerCurrencyCode.Text.Trim();

            }
            myCustomerSales.ShipmentCreditHoldCode.Value = "H";//客户订单信用强制--发货
            myCustomerSales.OrderEntryCreditHoldCode.Value = "H";//客户订单信用强制--订单录入
            if (_fstiClient.ProcessId(myCustomerSales, null))
            {
                Log += "子客户概要及信用添加成功:" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "子客户概要及信用添加失败:" + itemError.Description + ";";
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
                Log += "子客户财务会计信息添加成功:" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "子客户财务会计信息添加失败:" + itemError.Description + ";";
                return false;
            }
            #endregion

            return true;
        }
        private bool ADDGLOS(string strCode, string strName)//增加三级组织号GLOS
        {
            Log += "增加三级组织号GLOS:";
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

            if (_fstiClient.ProcessId(FirstGLOC, null))
            {
                Log += "一级增加成功-" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "一级增加失败-" + itemError.Description + ";";
            }
            if (_fstiClient.ProcessId(SecondGLOC, null))
            {
                Log += "二级增加成功-" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "二级增加失败-" + itemError.Description + ";";
            }
            if (_fstiClient.ProcessId(ThirdGLOC, null))
            {
                Log += "三级增加成功-" + _fstiClient.CDFResponse + ";";
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "三级增加失败-" + itemError.Description + ";";
            }
            return true;
        }
        public bool ADDGLAV(string OrganizationCode, string strCode)//增加账号GLAV不在ListResult中显示结果
        {
            Log += "增加账号GLAV:";
            #region FSTI增加GLAV
            GLAV00 myGLAV = new GLAV00();
            myGLAV.GLAccountGroup.Value = strCode;
            myGLAV.GLAccountValidationCode.Value = "1";
            myGLAV.GLOrganization.Value = OrganizationCode;
            if (_fstiClient.ProcessId(myGLAV, null))
            {
                Log += "GLAV增加成功-" + _fstiClient.CDFResponse + ";";
                return true;
            }
            else
            {
                FSTIError itemError = _fstiClient.TransactionError;
                Log += "GLAV增加失败-" + itemError.Description + ";";
                return false;
            }
            #endregion
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
    }
}
