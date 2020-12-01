using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Weekly_Jobs
{
    class Program
    {
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

        static void Main(string[] args)
        {

            SAP_DATA_PULL data_pull = new SAP_DATA_PULL();

            //SQL function results
            ArrayList results = data_pull.Get_Results();

            foreach (Dictionary<string, object> job in results)
            {
                Console.WriteLine("Adding additional data for Job #" + job["SUB_JOB_ID"].ToString());
                string jobID = SAP_DATA_PULL.Pull_SAP_JobID(job["SUB_JOB_ID"].ToString());
                string location = SAP_DATA_PULL.Pull_SAP_Parent_Job_Location(jobID);
                string plum = SAP_DATA_PULL.Job_Plumbing_Permit(jobID);
                string ac = SAP_DATA_PULL.Job_HVAC_Permit(jobID);
                string altcode = SAP_DATA_PULL.Job_Get_Altcode(job["SUB_JOB_ID"].ToString());
                string gas = SAP_DATA_PULL.Job_Gas_Permit(jobID);
                job.Add("PARENT_JOB_ID", jobID);
                job.Add("LOCATION", location);
                job.Add("PLUM PERMIT", plum);
                job.Add("AC PERMIT", ac);
                job.Add("GAS PERMIT", gas);
                job.Add("ALTCODE", altcode);
            }

            Application oXL = new Application();

            Workbook wb = oXL.Workbooks.Add(XlSheetType.xlWorksheet);

            //worksheets
            Worksheet hudson_ac = (Worksheet)oXL.ActiveSheet;
            hudson_ac.Name = "HUDSON AC";

            Worksheet orlando_ac = wb.Worksheets.Add();
            orlando_ac.Name = "ORLANDO AC";

            Worksheet south_ac = wb.Worksheets.Add();
            south_ac.Name = "SOUTH AC";

            Worksheet hudson_plum = wb.Worksheets.Add();
            hudson_plum.Name = "HUDSON PLUMBING";

            Worksheet orlando_plum = wb.Worksheets.Add();
            orlando_plum.Name = "ORLANDO PLUMBING";

            Worksheet south_plum = wb.Worksheets.Add();
            south_plum.Name = "SOUTH PLUMBING";

            ArrayList list = new ArrayList()
            {
                "SUBDIVSION",
                "PERMIT",
                "LOT",
                "ADDRESS",
                "PHASE",
                "INSTALLER",
                "START DATE",
                "INSPECTION DATE",
                "STATUS",
                "ALTCODE",
                "PAYMENT DATE",
            };

            for (int i = 1; i < 12; i++)
            {
                hudson_ac.Cells[1, i] = list[i - 1];
                hudson_plum.Cells[1, i] = list[i - 1];
                orlando_ac.Cells[1, i] = list[i - 1];
                orlando_plum.Cells[1, i] = list[i - 1];
                south_ac.Cells[1, i] = list[i - 1];
                south_plum.Cells[1, i] = list[i - 1];
            }

            //row iterators
            int hudson_ac_row = 2;
            int orlando_ac_row = 2;
            int south_ac_row = 2;
            int hudson_plum_row = 2;
            int orlando_plum_row = 2;
            int south_plum_row = 2;

            //iteration over each job
            foreach (Dictionary<string, object> dict in results)
            {
                Console.WriteLine("Writing Job #" + dict["SUB_JOB_ID"].ToString());

                string location = dict["LOCATION"].ToString();
                string type = dict["TYPE"].ToString();
                string subdivsion = dict["SUBDIVSION"].ToString();
                string permit = dict["PLUM PERMIT"].ToString();

                if (permit.Length == 0)
                    permit = dict["AC PERMIT"].ToString();
                if (permit.Length == 0)
                    permit = dict["GAS PERMIT"].ToString();

                string lot = dict["LOT"].ToString();
                string address = dict["ADDRESS"].ToString();
                string phase = dict["TYPE"].ToString();
                string start_date = dict["STARTDATE"].ToString();
                string altcode = dict["ALTCODE"].ToString();

                if (location == "Hudson")
                {
                    if (phase.Contains("PL"))
                    {
                        hudson_plum.Cells[hudson_plum_row, 1] = subdivsion;
                        hudson_plum.Cells[hudson_plum_row, 2] = permit;
                        hudson_plum.Cells[hudson_plum_row, 3] = lot;
                        hudson_plum.Cells[hudson_plum_row, 4] = address;
                        hudson_plum.Cells[hudson_plum_row, 5] = phase;
                        hudson_plum.Cells[hudson_plum_row, 6] = "";
                        hudson_plum.Cells[hudson_plum_row, 7] = start_date;
                        hudson_plum.Cells[hudson_plum_row, 8] = "";
                        hudson_plum.Cells[hudson_plum_row, 9] = "";
                        hudson_plum.Cells[hudson_plum_row, 10] = altcode;
                        hudson_plum.Cells[hudson_plum_row, 11] = "";
                        hudson_plum_row++;
                    }
                    else
                    {
                        hudson_ac.Cells[hudson_ac_row, 1] = subdivsion;
                        hudson_ac.Cells[hudson_ac_row, 2] = permit;
                        hudson_ac.Cells[hudson_ac_row, 3] = lot;
                        hudson_ac.Cells[hudson_ac_row, 4] = address;
                        hudson_ac.Cells[hudson_ac_row, 5] = phase;
                        hudson_ac.Cells[hudson_ac_row, 6] = "";
                        hudson_ac.Cells[hudson_ac_row, 7] = start_date;
                        hudson_ac.Cells[hudson_ac_row, 8] = "";
                        hudson_ac.Cells[hudson_ac_row, 9] = "";
                        hudson_ac.Cells[hudson_ac_row, 10] = altcode;
                        hudson_ac.Cells[hudson_ac_row, 11] = "";
                        hudson_ac_row++;
                    }
                }
                if (location == "Orlando")
                {
                    if (phase.Contains("PL"))
                    {
                        orlando_plum.Cells[orlando_plum_row, 1] = subdivsion;
                        orlando_plum.Cells[orlando_plum_row, 2] = permit;
                        orlando_plum.Cells[orlando_plum_row, 3] = lot;
                        orlando_plum.Cells[orlando_plum_row, 4] = address;
                        orlando_plum.Cells[orlando_plum_row, 5] = phase;
                        orlando_plum.Cells[orlando_plum_row, 6] = "";
                        orlando_plum.Cells[orlando_plum_row, 7] = start_date;
                        orlando_plum.Cells[orlando_plum_row, 8] = "";
                        orlando_plum.Cells[orlando_plum_row, 9] = "";
                        orlando_plum.Cells[orlando_plum_row, 10] = altcode;
                        orlando_plum.Cells[orlando_plum_row, 11] = "";
                        orlando_plum_row++;
                    }
                    else
                    {
                        orlando_ac.Cells[orlando_ac_row, 1] = subdivsion;
                        orlando_ac.Cells[orlando_ac_row, 2] = permit;
                        orlando_ac.Cells[orlando_ac_row, 3] = lot;
                        orlando_ac.Cells[orlando_ac_row, 4] = address;
                        orlando_ac.Cells[orlando_ac_row, 5] = phase;
                        orlando_ac.Cells[orlando_ac_row, 6] = "";
                        orlando_ac.Cells[orlando_ac_row, 7] = start_date;
                        orlando_ac.Cells[orlando_ac_row, 8] = "";
                        orlando_ac.Cells[orlando_ac_row, 9] = "";
                        orlando_ac.Cells[orlando_ac_row, 10] = altcode;
                        orlando_ac.Cells[orlando_ac_row, 11] = "";
                        orlando_ac_row++;
                    }
                }

                if (location == "South")
                {
                    if (phase.Contains("PL"))
                    {
                        south_plum.Cells[south_plum_row, 1] = subdivsion;
                        south_plum.Cells[south_plum_row, 2] = permit;
                        south_plum.Cells[south_plum_row, 3] = lot;
                        south_plum.Cells[south_plum_row, 4] = address;
                        south_plum.Cells[south_plum_row, 5] = phase;
                        south_plum.Cells[south_plum_row, 6] = "";
                        south_plum.Cells[south_plum_row, 7] = start_date;
                        south_plum.Cells[south_plum_row, 8] = "";
                        south_plum.Cells[south_plum_row, 9] = "";
                        south_plum.Cells[south_plum_row, 10] = altcode;
                        south_plum.Cells[south_plum_row, 11] = "";
                        south_plum_row++;
                    }
                    else
                    {
                        south_ac.Cells[south_ac_row, 1] = subdivsion;
                        south_ac.Cells[south_ac_row, 2] = permit;
                        south_ac.Cells[south_ac_row, 3] = lot;
                        south_ac.Cells[south_ac_row, 4] = address;
                        south_ac.Cells[south_ac_row, 5] = phase;
                        south_ac.Cells[south_ac_row, 6] = "";
                        south_ac.Cells[south_ac_row, 7] = start_date;
                        south_ac.Cells[south_ac_row, 8] = "";
                        south_ac.Cells[south_ac_row, 9] = "";
                        south_ac.Cells[south_ac_row, 10] = altcode;
                        south_ac.Cells[south_ac_row, 11] = "";
                        south_ac_row++;
                    }
                }
            }

            //autofitting sheets
            hudson_ac.UsedRange.Columns.AutoFit();
            hudson_ac.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            hudson_plum.UsedRange.Columns.AutoFit();
            hudson_plum.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            orlando_ac.UsedRange.Columns.AutoFit();
            orlando_ac.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            orlando_plum.UsedRange.Columns.AutoFit();
            orlando_plum.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            south_plum.UsedRange.Columns.AutoFit();
            south_plum.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            south_ac.UsedRange.Columns.AutoFit();
            south_ac.UsedRange.Columns.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            Console.WriteLine("Form Complete");
            oXL.Visible = true;

        }
    }
}
