using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3.Patterns;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using static BurnInTestValidate.Program;



namespace BurnInTestValidate
{

    public class DataManagent
    {
        //public readonly DbConnectionFactory _frmBurnIntest;

        //public DataManagent(DbConnectionFactory frmBurnIntest)
        //{
        //    _frmBurnIntest = frmBurnIntest;
        //}

        string errordesc = "";
        public string[] nextidinfo = { "", "" };
        public string[] reworkidinfo = { "24", "Rework" };
        public string[] infosfromprint = { "", "", "" }; // FG, Sno, WO
        public string[] infosfromboard = { "", "" };// WO,RW
        public string[] infologindetails = { "", "" };

        public string lbl_app_id = string.Empty;
        public string lblstagename = string.Empty;
        public string lblemp_id = string.Empty;
        public string lblappver = string.Empty;
        public string lbl_pcba_id = string.Empty;
        public int boardfailcount = 0;
         SqlConnection SFCS_db;

        // Timers
        private System.Windows.Forms.Timer monitorTimer;      // FlaUI UI monitor timer
        private System.Windows.Forms.Timer fileMonitorTimer;  // SSDMP.txt file monitor timer

        // FlaUI Window reference
        private FlaUI.Core.AutomationElements.Window mainWindow;


        private readonly DbConnectionFactory _ConnectionString;
        public int result = 0;
        public DataManagent(DbConnectionFactory dbConnectionFactory)
        {

            _ConnectionString = dbConnectionFactory;
        }

        public int inserthistory(PassmarkHistory objHistory)
        {
            using (SqlConnection sqlConnection = _ConnectionString.CreateConnection(Program.DatabaseType.Master))
            {
                using (SqlCommand sqlCommand = new SqlCommand("pro_passmarkhistory", sqlConnection))
                {
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.Parameters.AddWithValue("@FgNumber", objHistory.FgNumber);
                    sqlCommand.Parameters.AddWithValue("@CustomerSerialNumber", objHistory.CustomerSerialNumber);
                    sqlCommand.Parameters.AddWithValue("@PCBAID", objHistory.PCBAID);
                    sqlCommand.Parameters.AddWithValue("@DiskPartition", objHistory.DiskPartition);
                    sqlCommand.Parameters.AddWithValue("@CrystalReport", objHistory.CrystalReport);
                    sqlCommand.Parameters.AddWithValue("@read_one", string.IsNullOrEmpty(objHistory.read_one) ? "0" : objHistory.read_one);
                    sqlCommand.Parameters.AddWithValue("@read_two", string.IsNullOrEmpty(objHistory.read_two) ? "0" : objHistory.read_two);
                    sqlCommand.Parameters.AddWithValue("@read_three", string.IsNullOrEmpty(objHistory.read_three) ? "0" : objHistory.read_three);
                    sqlCommand.Parameters.AddWithValue("@read_four", string.IsNullOrEmpty(objHistory.read_four) ? "0" : objHistory.read_four);
                    sqlCommand.Parameters.AddWithValue("@write_one", string.IsNullOrEmpty(objHistory.write_one) ? "0" : objHistory.write_one);
                    sqlCommand.Parameters.AddWithValue("@write_two", string.IsNullOrEmpty(objHistory.write_two) ? "0" : objHistory.write_two);
                    sqlCommand.Parameters.AddWithValue("@write_three", string.IsNullOrEmpty(objHistory.write_three) ? "0" : objHistory.write_three);
                    sqlCommand.Parameters.AddWithValue("@write_four", string.IsNullOrEmpty(objHistory.write_four) ? "0" : objHistory.write_four);
                    sqlCommand.Parameters.AddWithValue("@burnintest", string.IsNullOrEmpty(objHistory.burnintest) ? "0" : objHistory.burnintest);
                    sqlCommand.Parameters.AddWithValue("@overall_result", string.IsNullOrEmpty(objHistory.overall_result) ? "0" : objHistory.overall_result);
                    sqlCommand.Parameters.AddWithValue("@createid", objHistory.CreatedBy);
                    sqlConnection.Open();
                    result = sqlCommand.ExecuteNonQuery();
                    sqlConnection.Close();
                }
            }
            return result;
        }
         
        public async Task<Dictionary<bool,int>> Check_Curr_Stage(string serialno, string app_id, string stage, bool boardonline = true)
        {
            bool checkCurrStageResult = false;

            Dictionary<bool, int> result = new Dictionary<bool, int>();
            try
            {
                var con = _ConnectionString.CreateConnection(DatabaseType.BurnIn);
                if (boardonline)
                {
                    con.Close();
                    using (SqlCommand cmd = new SqlCommand(
                        "SELECT * FROM PCBA_NextStage(NOLOCK) WHERE PCBA_Id = '" + serialno + "'", con))
                    {
                        if (con.State == ConnectionState.Closed)
                            con.Open();

                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.Read())
                            {
                                infosfromboard[0] = sdr["Work_order_no"].ToString();
                                infosfromboard[1] = sdr["Rework_count"].ToString();

                                //Fill_Response_Data("Board Waiting ID : " + sdr["Next_Stage_id"].ToString());
                                //Fill_Response_Data("Board Workorder : " + infosfromboard[0]);
                                //Fill_Response_Data("Board RW : " + infosfromboard[1]);

                                if (sdr["Next_Stage_id"].ToString() == app_id)
                                {
                                    con.Close();
                                    checkCurrStageResult = true;
                                   FailCountCheck(serialno, infosfromboard[0], infosfromboard[1],  stage);

                                }
                                else
                                {
                                    errordesc = "Stage Mismatch for this PCB : " + serialno + ".\n" +
                                                "Expected is : " + sdr["Next_Stage_id"] + "|" + sdr["Next_Stage_name"] + ".\n" +
                                                "Actual is : " + app_id + "|" + stage + ".";
                                    con.Close();
                                    checkCurrStageResult = false;
                                }
                            }
                            else
                            {
                                errordesc = "No Data for this PCB : " + serialno + " in SFCS Master Table.";
                                con.Close();
                                checkCurrStageResult = false;
                            }
                        }
                    }
                }
                else
                {
                    checkCurrStageResult = true;
                }
                result.Add(checkCurrStageResult, boardfailcount);

                return result;
            }
            catch (Exception ex)
            {
                checkCurrStageResult = false;
                return result;
            }
        }

        public void FailCountCheck(string pcbaid,string workorderno, string reworkcount,string stage)
        {
            SFCS_db= _ConnectionString.CreateConnection(DatabaseType.BurnIn);
            try {
                boardfailcount = 0;

                using (SqlCommand countCmd = new SqlCommand(
                    "SELECT COUNT(TEST) FROM TESTINGFAILCOUNTCHECK_ESSENCORE (NOLOCK) " +
                    "WHERE PCBAID = @pcbaid AND WORKORDER = @wo AND RW = @rw AND TEST = @test",
                    SFCS_db))
                {
                    countCmd.Parameters.Add("@pcbaid", SqlDbType.VarChar).Value = pcbaid;
                    countCmd.Parameters.Add("@wo", SqlDbType.VarChar).Value = workorderno;
                    countCmd.Parameters.Add("@rw", SqlDbType.Int).Value = 
                        string.IsNullOrEmpty(reworkcount.ToString()) ? 0 : Convert.ToInt32(reworkcount.ToString());
                    countCmd.Parameters.Add("@test", SqlDbType.VarChar).Value = stage;

                    SFCS_db.Open();

                    object result = countCmd.ExecuteScalar();

                    if (result != null && result != DBNull.Value)
                        boardfailcount = (int)Convert.ToInt64(result);

                    SFCS_db.Close();
                }

                //lbl_try.Text = (boardfailcount + 1).ToString();
               

               
            }
            catch(Exception ex)
            {
                throw ex;
               
            }
        }

        public int getReworkCount(string PcbSno)
        {
            int reworkresult = 0;
            var getcon = _ConnectionString.CreateConnection(DatabaseType.BurnIn);

            try
            {
                using (SqlCommand countCmd = new SqlCommand("pro_getReworkCount", getcon))
                {
                    countCmd.CommandType = CommandType.StoredProcedure;
                    countCmd.Parameters.AddWithValue("@PcbSno", PcbSno);
                    getcon.Open();
                    object result = countCmd.ExecuteScalar();
                    reworkresult = (result != null) ? Convert.ToInt32(result) : 0;
                    getcon.Close();
                }
            }
            catch (Exception ex)
            {
                reworkresult = 0;
            }
            return reworkresult;
        }
        public string SQL_Upload(string PcbSno, string CusSno, bool boardfail, string Result_Remarks , string[] stage)
        {
            var con = _ConnectionString.CreateConnection(DatabaseType.BurnIn);
            string sqluploadresult = string.Empty;
            if (stage.Length > 1)
            {
                lbl_app_id = "262";
                lblstagename = "Performance Test";
            }
            //var stagenext = Nextstartchecksfcs(CusSno);
            //if (stagenext != null)
            //{
            //    nextidinfo[0] = stagenext[0];
            //    nextidinfo[1]= stagenext[1];
            //}
            int failcount = boardfail == true ? 1 : 0;
            var checkReworkcount = getReworkCount(PcbSno);
            int boardfailcountnew = checkReworkcount + failcount;
            try
            {
                con.Close();

                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO DASHBOARD_ESSENCOREDATAS VALUES (" +
                    "'" + lbl_app_id + "'," +
                    "'" + lblstagename + "'," +
                    "'" + infosfromboard[0] + "'," +
                    "'" + PcbSno + "'," +
                    "'" + CusSno + "'," +

                    "'" + (boardfail ? "FAIL" : "PASS") + "'," +
                    "'" + errordesc + "'," +
                    "'" + infosfromboard[1] + "'," +
                    "CONCAT(FORMAT(CURRENT_TIMESTAMP,'HH'),'-',FORMAT(DATEADD(HOUR,1,CURRENT_TIMESTAMP),'HH'))," +
                    "CASE " +
                    "WHEN (FORMAT(CURRENT_TIMESTAMP,'HH:mm:ss') >= CONVERT(datetime,'08:00:00',103)) AND (FORMAT(CURRENT_TIMESTAMP,'HH:mm:ss') < CONVERT(datetime,'16:00:00',103)) THEN 'SHIFT-A' " +
                    "WHEN (FORMAT(CURRENT_TIMESTAMP,'HH:mm:ss') >= CONVERT(datetime,'16:00:00',103)) AND (FORMAT(CURRENT_TIMESTAMP,'HH:mm:ss') < CONVERT(datetime,'23:59:59',103)) THEN 'SHIFT-B' " +
                    "ELSE 'SHIFT-C' END," +
                    "FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy')," +
                    "'" + lblemp_id + "'," +
                    "FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy HH:mm:ss')," +
                    "HOST_NAME(),'','','')",
                    con);

                if (con.State == ConnectionState.Closed)
                    con.Open();

                cmd.ExecuteNonQuery();
                con.Close();
                // Fill_Response_Data("SQL Report Success. - SFCS Dashboard");
                sqluploadresult = "Success";
            }
            catch (Exception ex)
            {
                Update_Error_in_Server("Exception", "ERR-SQL-02", ex.Message.ToString(),
                    "SFCS Dashboard", "PCBA:" + PcbSno + ",Workorder:" + lblemp_id + ",CustomerNo:" + CusSno + ".");
                //lbl_result.Text += "SFCS Dashboard Failed.";
                //lbl_result.BackColor = Color.Red;
                //lbl_result.ForeColor = Color.Yellow;
                //Fill_Response_Data("SQL Report Failed. - SFCS Dashboard");
                sqluploadresult = "Error " + "Exception" + "ERR-SQL-02" + ex.Message.ToString() +
                    "SFCS Dashboard" + "PCBA:" + PcbSno + ",Workorder:" + lblemp_id + ",CustomerNo:" + CusSno + ".";
            }

            for (int tryupdate = 1; tryupdate <= 3; tryupdate++)
            {
                try
                {
                    con.Close();

                    SqlCommand cmd = new SqlCommand(
                        "UPDATE PCBA_NextStage SET " +
                        "Next_Stage_Id = '" + (boardfail ? (boardfailcountnew > 2 ? reworkidinfo[0] : lbl_app_id) : nextidinfo[0]) + "', " +
                        "Next_Stage_Name = '" + (boardfail ? (boardfailcountnew > 2 ? reworkidinfo[1] : lblstagename) : nextidinfo[1]) + "', " +
                        "Rework_Count = '" + boardfailcountnew + "'," +
                        "Previous_Stage = '" + lblstagename + "', " +
                        "Update_timestamp = FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy HH:mm:ss.ffff'), " +
                        "Update_Machine_id = HOST_NAME(), " +
                        "Update_Emp_id = '" + lblemp_id + "' " +
                        "WHERE PCBA_Id = '" + PcbSno + "'",
                        con);

                    if (con.State == ConnectionState.Closed)
                        con.Open();

                    cmd.ExecuteNonQuery();
                    con.Close();
                    //Fill_Response_Data("Next Stage : " + (boardfail ? reworkidinfo[1] : nextidinfo[1]));
                    //Fill_Response_Data("SFCS Next Stage Update Success.");
                    sqluploadresult = "SFCS Next Stage Update Success." + "Next Stage : " + (boardfail ? reworkidinfo[1] : stage[1]);
                    break;
                }
                catch (Exception ex)
                {
                    Update_Error_in_Server("Exception", "ERR-SQL-03", ex.Message.ToString(),
                        "SFCS Nextstage Failed", "PCBA:" + PcbSno + ",Workorder:" + infosfromboard[0] + ",CustomerNo:" + CusSno + ".");
                    //lbl_result.Text += "SFCS Nextstage Failed.";
                    //lbl_result.BackColor = Color.Red;
                    //lbl_result.ForeColor = Color.Yellow;
                    //Fill_Response_Data("SFCS Next Stage Update Failed.");
                    sqluploadresult = "Error " + "Exception" + "ERR-SQL-03" + ex.Message.ToString() + "SFCS Nextstage Failed" + "PCBA:" + PcbSno + ",Workorder:" + infosfromboard[0] + ",CustomerNo:" + CusSno + ".";
                }
            }

            try
            {
                con.Close();

                SqlCommand cmd = new SqlCommand("INSERT INTO FCT VALUES('CHN1','" + lblstagename + "','" + PcbSno + "','" + (boardfail ? "FAIL" : "PASS") + "','" + errordesc + "',FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy HH:mm:ss.ffff'),'" + lblemp_id + "',HOST_NAME(),'" + infosfromboard[1] + "','" + infosfromboard[0] + "','')", con);
                if (con.State == ConnectionState.Closed)
                    con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                // Fill_Response_Data("SFCS FCT Success.");
                sqluploadresult = "SFCS FCT Success.";
            }
            catch (Exception ex)
            {
                Update_Error_in_Server("Exception", "ERR-SQL-04", ex.Message.ToString(),
                    "SFCS FCT Failed", "PCBA:" + PcbSno + ",Workorder:" + infosfromboard[0] + ",CustomerNo:" + CusSno + ".");
                //lbl_result.Text += "SFCS FCT Failed.";
                //lbl_result.BackColor = Color.Red;
                //lbl_result.ForeColor = Color.Yellow;
                //Fill_Response_Data("SFCS FCT Update Failed.");
                sqluploadresult = "SFCS FCT Update Failed.";
            }
            try
            {

            }
            catch (Exception ex)
            {
            }
            if (boardfail)
            {
                var machinename = Environment.MachineName;
               var failcountinsert = insertFailCountCheck(PcbSno, CusSno, boardfailcountnew.ToString(), lblstagename, machinename);
            }
            //try
            //{
            //    con.Close();

            //    string datatoenter = string.Join(",", xmlinfos.Skip(1).Take(8));

            //    SqlCommand cmd = new SqlCommand(
            //        "INSERT INTO PROD_PERFORMACETESTDATA VALUES (" +
            //        "'Essencore'," +
            //        "'" + expectedvalues[3] + "'," +
            //        "'" + infosfromboard[0] + "'," +
            //        "'" + lbl_pcba_id + "'," +
            //        "'" + lbl_SN + "'," +
            //        "'" + Path.GetFileName(filenametocheck) + "'," +
            //        "'" + xmlinfos[0] + "'," +
            //        "'" + datatoenter + "'," +
            //        "'" + (boardfail ? "FAIL" : "PASS") + "'," +
            //        "'" + xmlinfos[9] + "--" + errordesc + "'," +
            //        "SYSDATETIME()," +
            //        "'" + Login.infologindetails[0] + "'," +
            //        "HOST_NAME()," +
            //        "'','','','','')",
            //        con);

            //    if (con.State == ConnectionState.Closed)
            //        con.Open();

            //    cmd.ExecuteNonQuery();
            //    con.Close();

            //    Fill_Response_Data("SQL Report Success - Essencore DB");
            //}
            //catch (Exception ex)
            //{
            //    Update_Error_in_Server(
            //        "Exception",
            //        "ERR-SQL-02",
            //        ex.Message.ToString(),
            //        "Essencore DB",
            //        "PCBA:" + lbl_pcba_id + ",SN:" + lbl_SN
            //    );

            //    //lbl_result.Text += "Essencore DB Failed.";
            //    //lbl_result.BackColor = Color.Red;
            //    //lbl_result.ForeColor = Color.Yellow;

            //    Fill_Response_Data("SQL Report Failed - Essencore DB");
            //}
            return sqluploadresult;
        }

        public int insertFailCountCheck(string pcbid, string workorder, string rowcount, string stagename, string machinename)
        {
            int insertresult = 0;
            var insertcon= _ConnectionString.CreateConnection(DatabaseType.BurnIn);
            try
            {
                using(SqlCommand cmd=new SqlCommand ("pro_insertTESTINGFAILCOUNTCHECK_ESSENCORE",insertcon))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@WORKORDER", workorder);
                    cmd.Parameters.AddWithValue("@Test", stagename);
                    cmd.Parameters.AddWithValue("@PCBAID", pcbid);
                    cmd.Parameters.AddWithValue("@Reworkcount", rowcount);
                    cmd.Parameters.AddWithValue("@Result", "Fail");
                    cmd.Parameters.AddWithValue("@ErrorCode", "Fail");
                    cmd.Parameters.AddWithValue("@ERRORDESC", "performance Test Fail Crystal Report/ BurnIn Test");
                    cmd.Parameters.AddWithValue("@UPDATEMACHINE", machinename);
                    insertcon.Open();
                    insertresult=cmd.ExecuteNonQuery();
                    insertcon.Close();
                   return insertresult;

                }
            }
            catch (Exception ex) { 
                insertcon.Close();
                return insertresult;
            }
        }

        public async Task<string[]> startchecksfcs(string FgNumber)
        {
            try
            {
                var SFCS_db = _ConnectionString.CreateConnection(DatabaseType.BurnIn);
                if (SFCS_db.State == ConnectionState.Open)
                    SFCS_db.Close();

                SqlCommand cmd = new SqlCommand(
                    "SELECT * FROM RoutingStages WHERE FG = '" + FgNumber + "'",
                    SFCS_db);

                if (SFCS_db.State == ConnectionState.Closed)
                    SFCS_db.Open();

                SqlDataReader sdr = cmd.ExecuteReader();

                if (sdr.Read())
                {
                    try
                    {
                        string[] stages = sdr["Stages"].ToString().Split(',');
                        int index = Array.IndexOf(stages,"262");

                        if (index >= 0 && index < stages.Length - 1)
                            nextidinfo[0] = stages[index +1];
                      
                    }
                    catch (Exception)
                    {
                        // ignored (same as VB empty catch)
                    }
                }

                sdr.Close();
                SFCS_db.Close();

                string cmd1 = "SELECT App_ID, Application_Name FROM App_ver WHERE App_ID = '" + nextidinfo[0] + "'";
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1, SFCS_db);
                DataSet ds1 = new DataSet();
                da1.Fill(ds1, "app_name");

                if (ds1.Tables[0].Rows.Count > 0)
                {
                    nextidinfo[1] = ds1.Tables[0].Rows[0][1].ToString();
                }

                SFCS_db.Close();

                //Fill_Response_Data("SFCS Next Stage : " + nextidinfo[0] + "|" + nextidinfo[1]);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(),
                //                "Exception in Fetching Next Stage",
                //                MessageBoxButtons.OK,
                //                MessageBoxIcon.Information);

                return nextidinfo;   // equivalent to Exit Sub
            }
            return nextidinfo;
        }

        public string[] Nextstartchecksfcs(string FgNumber)
        {
            try
            {
                var SFCS_db = _ConnectionString.CreateConnection(DatabaseType.BurnIn);
                if (SFCS_db.State == ConnectionState.Open)
                    SFCS_db.Close();

                SqlCommand cmd = new SqlCommand(
                    "SELECT * FROM RoutingStages WHERE FG = '" + FgNumber + "'",
                    SFCS_db);

                if (SFCS_db.State == ConnectionState.Closed)
                    SFCS_db.Open();

                SqlDataReader sdr = cmd.ExecuteReader();

                if (sdr.Read())
                {
                    try
                    {
                        string[] stages = sdr["Stages"].ToString().Split(',');
                        int index = Array.IndexOf(stages, "262");

                        if (index >= 0 && index < stages.Length - 1)
                        {
                            nextidinfo[0] = stages[index + 1];
                        }

                    }
                    catch (Exception)
                    {
                        // ignored (same as VB empty catch)
                    }
                }

                sdr.Close();
                SFCS_db.Close();

                string cmd1 = "SELECT App_ID, Application_Name FROM App_ver WHERE App_ID = '" + nextidinfo[0] + "'";
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1, SFCS_db);
                DataSet ds1 = new DataSet();
                da1.Fill(ds1, "app_name");

                if (ds1.Tables[0].Rows.Count > 0)
                {
                    nextidinfo[1] = ds1.Tables[0].Rows[0][1].ToString();
                }

                SFCS_db.Close();

                //Fill_Response_Data("SFCS Next Stage : " + nextidinfo[0] + "|" + nextidinfo[1]);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString(),
                //                "Exception in Fetching Next Stage",
                //                MessageBoxButtons.OK,
                //                MessageBoxIcon.Information);

                return nextidinfo;   // equivalent to Exit Sub
            }
            return nextidinfo;
        }
        private void Update_Error_in_Server(string errortype, string errorcode, string errordesc, string errorloc, string errorremarks)
        {
            string inqry = string.Empty;
            var con = _ConnectionString.CreateConnection(DatabaseType.BurnIn);
            try
            {
                con.Close();

                inqry = "INSERT INTO EXCEPTIONLOGS_MEMORY VALUES ('" +
                        errortype + "','" +
                        lbl_app_id + "','" +
                        lblstagename + "','" +
                        lblappver + "','" +
                        errorcode + "','" +
                        errordesc.Replace("'", "@") + "','" +
                        errorloc.Replace("'", "@") + "','" +
                        errorremarks.Replace("'", "@") + "'," +
                        "FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy')," +
                        "FORMAT(CURRENT_TIMESTAMP,'dd-MM-yyyy HH:mm:ss.fff')," +
                        "HOST_NAME(),'" + lblemp_id + "','','')";

                using (SqlCommand cmd = new SqlCommand(inqry, con))
                {
                    if (con.State == ConnectionState.Closed)
                        con.Open();

                    cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
                // Fill_Response_Data("Exception in Updating Error in Server.");
                //Fill_Response_Data(ex.Message.ToString());
                //updateerrorresult = "Exception in Updating Error in Server. " + ex.Message.ToString();
            }

        }

        public DataTable GetProductTypes()
        {
            DataTable dt = new DataTable();
            using (SqlConnection sqlConnection = _ConnectionString.CreateConnection(Program.DatabaseType.Master))
            {
                using (SqlCommand sqlCommand = new SqlCommand("pro_getProductType", sqlConnection))
                {
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlConnection.Open();
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
                    {
                        sqlDataAdapter.Fill(dt);
                    }
                    sqlConnection.Close();
                }
            }
            return dt;
        }
        public DataTable GetFGNames(int productTypeId)
        {
            DataTable dt = new DataTable();
            using (SqlConnection sqlConnection = _ConnectionString.CreateConnection(Program.DatabaseType.Master))
            {
                using (SqlCommand sqlCommand = new SqlCommand("pro_getFgName", sqlConnection))
                {
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.Parameters.AddWithValue("@producttype", productTypeId);
                    sqlConnection.Open();
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
                    {
                        sqlDataAdapter.Fill(dt);
                    }
                    sqlConnection.Close();
                }
            }
            return dt;
        }

        public bool ValidateUser(string username, string password)
        {
            bool isValid = false;
            using (SqlConnection sqlConnection = _ConnectionString.CreateConnection(Program.DatabaseType.Master))
            {
                using (SqlCommand sqlCommand = new SqlCommand("pro_validateUser", sqlConnection))
                {
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.Parameters.AddWithValue("@username", username);
                    sqlCommand.Parameters.AddWithValue("@password", password);
                    sqlConnection.Open();
                    var result = sqlCommand.ExecuteScalar();
                    if (result != null && Convert.ToInt32(result) > 0)
                    {
                        isValid = true;
                    }
                    sqlConnection.Close();
                }
            }
            return isValid;
        }

        //private void Fill_Response_Data(string datarep)
        //{
        //    txt_live_stat.AppendText(
        //        DateTime.Now.ToString("HH:mm:ss.fff") + " : " + datarep + Environment.NewLine
        //    );

        //    txt_live_stat.SelectionStart = txt_live_stat.TextLength;
        //    txt_live_stat.ScrollToCaret();
        //    base.Update();
        //    Application.DoEvents();
        //}

        public async Task<FgDetails> GetFgDetails(int productId, string cusNumber)
        {
            FgDetails details = null;
            using (SqlConnection sqlConnection = _ConnectionString.CreateConnection(Program.DatabaseType.Reporting))
            {
                using (SqlCommand sqlCommand = new SqlCommand("Pro_getFgDetails", sqlConnection))
                {
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.Parameters.AddWithValue("@producttype", productId);
                    sqlCommand.Parameters.AddWithValue("@customerserialnumber", cusNumber);
                    sqlConnection.Open();
                    using (SqlDataReader reader = sqlCommand.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            details = new FgDetails
                            {
                                FgName = reader["FgNumber"].ToString(),
                                ProductType = reader["PCBSerialNo"].ToString(),
                            };
                        }
                    }
                    sqlConnection.Close();
                }
            }
            return details;

        }
    }
}
