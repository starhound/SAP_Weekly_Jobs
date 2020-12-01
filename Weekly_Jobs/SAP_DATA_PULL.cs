using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace Weekly_Jobs
{
    class SAP_DATA_PULL
    {
        public ArrayList results = new ArrayList();
        const int STARTDATE = 0;
        const int SUBJOBID = 1;
        const int JOBTITLE = 2;
        const int CARDNAME = 3;
        const int LOTBLOCK = 4;
        const int MODEL = 5;
        const int TYPE = 6;
        const int CATEGORY = 7;
        const int INVOICENUM = 8;
        const int PARENT_JOB_ID = 9;
        const int LOCATION = 10;

        static Dictionary<int, string> DATA_COLUMN_NAMES = new Dictionary<int, string>()
        {
            { STARTDATE, "STARTDATE" },
            { SUBJOBID, "SUB_JOB_ID" },
            { JOBTITLE, "ADDRESS" },
            { CARDNAME, "BUILDER" },
            { LOTBLOCK, "LOT" },
            { MODEL, "MODEL" },
            { TYPE, "TYPE" },
            { CATEGORY, "SUBDIVSION" },
            { INVOICENUM, "INVOICE_DATE" },
            { LOCATION, "LOCATION" },
            { PARENT_JOB_ID, "PARENT_JOB_ID" },
        };

        const string ConnectionString = @"YOUR_CONNECTION_STRING";

        public ArrayList Get_Results()
        {
            Pull_SAP_Data();
            return results;
        }

        public static string Pull_SAP_JobID(string subJobID)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand jobIDQuery = new SqlCommand("select JOBID from [dbo].[ENPRISE_JOBCOST_SUBJOB] where SUBJOBID = @SUBJOBID", conn);
            string jobID = null;
            using (conn)
            {
                jobIDQuery.Parameters.AddWithValue("@SUBJOBID", subJobID);
                conn.Open();
                SqlDataReader reader = jobIDQuery.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        jobID = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            return jobID;
        }

        public static string Pull_SAP_Parent_Job_Location(string jobID)
        {
            string location_code = null;
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand locationQuery = new SqlCommand("select U_Region from [dbo].[ENPRISE_JOBCOST_JOB] where JOBID = @PARENT_JOB_ID", conn);
            using (conn)
            {
                locationQuery.Parameters.AddWithValue("@PARENT_JOB_ID", jobID);
                conn.Open();
                SqlDataReader reader = locationQuery.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        location_code = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            string location_name;
            switch (location_code)
            {
                case "01":
                    location_name = "Hudson";
                    break;
                case "02":
                    location_name = "Orlando";
                    break;
                case "03":
                    location_name = "South";
                    break;
                default:
                    location_name = "NONE";
                    break;
            }
            return location_name;
        }

        public static string Job_Plumbing_Permit(string job)
        {
            string permit = null;
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand locationQuery = new SqlCommand("select U_PLPermit from [dbo].[ENPRISE_JOBCOST_JOB] where JOBID = @PARENT_JOB_ID", conn);
            using (conn)
            {
                locationQuery.Parameters.AddWithValue("@PARENT_JOB_ID", job);
                conn.Open();
                SqlDataReader reader = locationQuery.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        permit = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            return permit;
        }

        public static string Job_HVAC_Permit(string job)
        {
            string permit = null;
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand locationQuery = new SqlCommand("select U_HVACPermit from [dbo].[ENPRISE_JOBCOST_JOB] where JOBID = @PARENT_JOB_ID", conn);
            using (conn)
            {
                locationQuery.Parameters.AddWithValue("@PARENT_JOB_ID", job);
                conn.Open();
                SqlDataReader reader = locationQuery.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        permit = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            return permit;
        }

        public static string Job_Get_Altcode(string job)
        {
            string code = null;
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand alt_code_query = new SqlCommand("select ALTCODE from [dbo].[ENPRISE_JOBCOST_SUBJOB] where SUBJOBID = @SUB_JOB_ID", conn);
            using (conn)
            {
                alt_code_query.Parameters.AddWithValue("@SUB_JOB_ID", job);
                conn.Open();
                SqlDataReader reader = alt_code_query.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        code = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            return code;
        }

        public static string Job_Gas_Permit(string job)
        {
            string permit = null;
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand locationQuery = new SqlCommand("select U_GASPermit from [dbo].[ENPRISE_JOBCOST_JOB] where JOBID = @PARENT_JOB_ID", conn);
            using (conn)
            {
                locationQuery.Parameters.AddWithValue("@PARENT_JOB_ID", job);
                conn.Open();
                SqlDataReader reader = locationQuery.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        permit = reader[0].ToString();
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            return permit;
        }

        void Pull_SAP_Data()
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            SqlCommand cmd = new SqlCommand("Job_Schedule_ReportPhase", conn);
            DataSet ds = new DataSet();
            cmd.CommandType = CommandType.StoredProcedure;
            string today = DateTime.Today.ToString("yyyy-MM-dd");

            cmd.Parameters.Add(new SqlParameter("@FDate", today + " 00:00:00"));
            cmd.Parameters.Add(new SqlParameter("@TDate", today + " 00:00:00"));
            conn.Open();
            cmd.ExecuteNonQuery();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            adapter.Fill(ds);

            foreach (DataTable table in ds.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    //ArrayList rowData = new ArrayList();
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    int count = 0;
                    bool flag = false;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (count == TYPE)
                        {
                            string type = row[col].ToString();
                            if (type.Equals("Water/Sewer Service") || type.Equals("PL Camera Lines"))
                                flag = true;
                        }
                        object item = row[col];
                        if (count == INVOICENUM)
                        {
                            item = "          ";
                        }
                        string ID = DATA_COLUMN_NAMES[count];

                        rowData.Add(ID, item);
                        count++;
                    }
                    if (!flag)
                    {
                        Console.WriteLine("Pulling Job #" + rowData["SUB_JOB_ID"].ToString());
                        rowData.Add("INSTALLER", " ");
                        rowData.Add("STATUS", " ");
                        rowData.Add("INSPECTION DATE", " ");
                        rowData.Add("PAYMENT DATE", " ");
                        results.Add(rowData);
                    }
                }
            }
            Console.WriteLine("All Jobs Pulled\n\n");
            conn.Close();
        }
    }
}
