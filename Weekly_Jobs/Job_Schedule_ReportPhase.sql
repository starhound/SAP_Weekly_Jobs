SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE Job_Schedule_ReportPhase 
	@FDate as datetime,
	@TDate as datetime
AS
BEGIN
	SET NOCOUNT ON;

	 select 
		t0.DUEDATE [STARTDATE], 
		t0.SUBJOBID as JOBID, 
		t1.JOBTITLE, 
		t1.CARDNAME, 
		t1.U_Lot_Block, 
		t1.U_Model, 
		t2.TYPE, 
		t3.CATEGORY,
	    INVOICENUM =  ISNULL(( Select top 1 DocNum from INV1 S0, OINV S1 WHERE S1.DocEntry = S0.DocEntry AND S0.U_SubJob = T0.SUBJOBID ),-1)

	 from 
		ENPRISE_JOBCOST_SUBJOB T0, 
		ENPRISE_JOBCOST_JOB T1, 
		ENPRISE_JOBCOST_JOBTYPE T2, 
		ENPRISE_JOBCOST_CATEGORY T3
	 where 
		t1.JOBID = t0.JOBID 
		and 
		t2.SEQNO = t0.JOBTYPE 
		and 
		t3.SEQNO = t0.CATEGORY 
		and 
		t0.STATUS IN (1, 20)  
		and 
		t0.ACTIVE = 'Y'
		and  
		t0.DUEDATE >= @FDate 
		and 
		t0.DUEDATE <= @TDate
	order by 
		t0.DUEDATE, 
		t2.TYPE
END
GO
