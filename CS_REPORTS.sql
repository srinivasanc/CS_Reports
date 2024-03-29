/****** Object:  StoredProcedure [dbo].[RPT_CSFB_MAIN_TICKET_INFO]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_MAIN_TICKET_INFO]
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_MAIN_TICKET_INFO
		Written On	:   4th Feb 2014 By Ramkumar
		Description	:   Produces data for Main Ticket Info
	*/
	
	------ Test Data Start
	--DECLARE @SITES		VARCHAR(1024)	= 'NY'
	--DECLARE @TYPE		VARCHAR(20)		= '%'
	--DECLARE @FROMDATE	DATETIME		= '2014-01-01'
	--DECLARE @TODATE		DATETIME		= '2014-02-28'
	--DECLARE @USER		VARCHAR(50)		= 'ADMIN'
	--DECLARE @URL		VARCHAR(255)	= ''
	------ Test Data End
		
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	
	SELECT	
		[Job #]					=	J.job_no,
		[Pool]					=	J.[site],	
		[Submission Date]		=	SUBSTRING(CONVERT(VARCHAR, J.[REQUESTED_DATE], 101), 0, 13),			
		[Submission Time]		=	LEFT(CONVERT(VARCHAR(8), CAST(J.[REQUESTED_DATE] AS DATETIME), 108), 5),
		[Report Title]			=	JOB_TITLE,	
		[Due Date]				=	SUBSTRING(CONVERT(VARCHAR, J.[DUE_DATE], 101), 0, 13),	
		[Due Time]				=	LEFT(CONVERT(VARCHAR(8), CAST(J.[DUE_DATE] AS DATETIME), 108), 5), 
		[Business Purpose]		=	[ADDRESS],	
		[Completed Date]		=   SUBSTRING(CONVERT(VARCHAR, J.COMPLETED, 101), 0, 13),	
        [Completed Time]		=   SUBSTRING(CONVERT(VARCHAR, DATEADD(DAY, DATEDIFF(DAY, J.COMPLETED, 0), J.COMPLETED)), 12, 20),
		[Project Code]			=	ALT01,
		[Client Code]			=	EML06,
		[Recruitment Code]		=	ATTENTION,
		[MyTE Report Started]	=	BATES_NO1,
		[Attachments]			=	FILEPATH,
		[Banker Name]			=	J.ACCT_EXEC,
		[Banker ID]     		=   CUSTOMER,	
		[Division]				=	DEPARTMENT,	
		[Group]					=	JOB_GROUP,	
		[Location]				=	LOCATION,	
		[Cost Center]			=	CLIENT,	
		[Submitted By]			=	REQUESTED_BY,
		[Submittors  email]		=	EMAIL,
		[Time Estimation]		=   COMPANY,
		[Scale]					=	ALT02,
		[Receipt]				=	ALT03,
		[Amounts]				=	ALT04
	FROM	
		VWJOBTICKETS J WITH(NOLOCK) 
		
	WHERE	
		J.[REQUESTED_DATE] BETWEEN @FROMDATE AND @TODATE
		AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 
		AND JOB_TYPE LIKE @TYPE
    
	ORDER BY J.JOB_NO
END
GO
/****** Object:  StoredProcedure [dbo].[RPT_CSFB_WORKFLOW]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_WORKFLOW] --'pool1','Expense','01/01/2014','01/28/2014','admin',''
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_WORKFLOW
		Written On	:   01/22/2014 By Muthamil Selvan
		Description	:   Produces data for Workflow	
	*/

	
			
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	
	SELECT	[Job #]				=	j.job_no,
			[Pool]				=	j.[site],
			[Submitted By]		=	REQUESTED_BY,	
			[Submission Date]	=	SUBSTRING(CONVERT(VARCHAR, J.[REQUESTED_DATE], 101), 0, 13),			
			[Submission Time]	=	LEFT(CONVERT(VARCHAR(8), CAST(J.[REQUESTED_DATE] AS DATETIME), 108), 5),
			[Started Time]		=	W.[Started],
			--[Days To Start]     =   DateDiff(Day, JT.ACCEPTED,  Min(W.STARTED)),
			[Days To Start]     =   CASE when JT.ACCEPTED IS NULL THEN J.[Requested_Date] ELSE DateDiff(Day, JT.ACCEPTED,  Min(W.STARTED)) END ,
			[Report Title]		=	JOB_TITLE,	
			[Due Date]			=	SUBSTRING(CONVERT(VARCHAR, J.[DUE_DATE], 101), 0, 13),	
			[Due Time]			=	LEFT(CONVERT(VARCHAR(8), CAST(J.[DUE_DATE] AS DATETIME), 108), 5), 
			[Business Purpose]	=	[ADDRESS],
			[Completed Date]    =   SUBSTRING(CONVERT(VARCHAR, J.COMPLETED, 101), 0, 13),	
            [Completed Time]    =   SUBSTRING(CONVERT(VARCHAR, DATEADD(DAY, DATEDIFF(DAY, J.COMPLETED, 0), J.COMPLETED)), 12, 20),
			[Days To Complete]  =   DateDiff(Day,J.[REQUESTED_DATE],J.Completed),	
			[Group]				=	JOB_GROUP,	
			[Total Hours]		=	W.[Hours]
	FROM	VWJOBTICKETS J WITH(NOLOCK) JOIN JOB_TICKET_WORKFLOW W WITH(NOLOCK) ON J.JOB_NO = W.JOB_NO AND J.[SITE] = W.[SITE]  
	        JOIN  job_ticket_routes JT ON J.JOB_NO = JT.JOB_NO AND J.[SITE] = JT.[SITE]
	WHERE	J.[REQUESTED_DATE] BETWEEN @FROMDATE AND @TODATE 
	        AND J.[status] IN ('ACCEPTED','PENDING','COMPLETED')
			AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 AND JOB_TYPE LIKE @TYPE
			and w.[started] is not null
	GROUP BY j.job_no, j.[site], REQUESTED_BY, J.[REQUESTED_DATE], W.[Started], JT.ACCEPTED, W.STARTED, JOB_TITLE, J.[DUE_DATE], [ADDRESS], J.COMPLETED,  JOB_GROUP, W.[Hours]
	ORDER BY j.job_no
	

END
GO
/****** Object:  StoredProcedure [dbo].[RPT_CSFB_TIME_ESTIMATION]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_TIME_ESTIMATION]
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_TIME_ESTIMATION
		Written On	:   01/14/2014 By Muthamil Selvan
		Description	:   Produces data for Time Estimation		
	*/
	
		
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	
	SELECT	[Job #]					=	J.job_no,
			[Pool]					=	J.[site],	
			[Submission Date]		=	SUBSTRING(CONVERT(VARCHAR, J.[REQUESTED_DATE], 101), 0, 13),			
			[Submission Time]		=	LEFT(CONVERT(VARCHAR(8), CAST(J.[REQUESTED_DATE] AS DATETIME), 108), 5),
			[Report Title]			=	JOB_TITLE,	
			[Due Date]				=	SUBSTRING(CONVERT(VARCHAR, J.[DUE_DATE], 101), 0, 13),	
			[Due Time]				=	LEFT(CONVERT(VARCHAR(8), CAST(J.[DUE_DATE] AS DATETIME), 108), 5), 
			[Business Purpose]		=	[ADDRESS],	
			[Completed Date]		=   SUBSTRING(CONVERT(VARCHAR, J.COMPLETED, 101), 0, 13),	
            [Completed Time]		=   SUBSTRING(CONVERT(VARCHAR, DATEADD(DAY, DATEDIFF(DAY, J.COMPLETED, 0), J.COMPLETED)), 12, 20),
			[Project Code]			=	ALT01,
			[Client Code]			=	EML06,	
			[Recruitment Code]		=	ATTENTION,
			[Banker Name]			=	J.ACCT_EXEC,
			[Banker ID]     		=   CUSTOMER,	
			[Division]				=	DEPARTMENT,	
			[Group]					=	JOB_GROUP,	
			[Location]				=	LOCATION,	
			[Cost Center]			=	CLIENT,	
			[Submitted By]			=	REQUESTED_BY,
			[Time Estimation]		=   COMPANY,
			[Total Hours]			=	SUM(W.Hours),
			[Hours Minus Estimation] =  case when isnumeric(company)=1 then company - sum(w.hours) else 0 end
			
	FROM	VWJOBTICKETS J WITH(NOLOCK) JOIN JOB_TICKET_WORKFLOW W WITH(NOLOCK) ON J.JOB_NO = W.JOB_NO AND J.[SITE] = W.[SITE]       
	WHERE	J.[REQUESTED_DATE] BETWEEN @FROMDATE AND @TODATE 
	        AND J.status = 'COMPLETED'
			AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 AND JOB_TYPE LIKE @TYPE
    GROUP BY j.job_no, j.[site],J.[REQUESTED_DATE], J.[DUE_DATE],JOB_TITLE, [ADDRESS], J.COMPLETED,ALT01,EML06,ATTENTION,J.ACCT_EXEC,CUSTOMER,DEPARTMENT,JOB_GROUP,LOCATION, CLIENT,REQUESTED_BY,COMPANY
	ORDER BY j.job_no
	

END
GO
/****** Object:  StoredProcedure [dbo].[RPT_CSFB_PERSONNEL]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_PERSONNEL]
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_PERSONNEL
		Written On	:   01/17/2014 By Muthamil Selvan
		Description	:   Produces data for Personnel		
	*/
		
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	
	SELECT	[Job #]				=	j.job_no,
			[Site]				=	j.[site],	
			[Banker Name]		=	j.ACCT_EXEC,
			[Personnel]			=	DBO.FN_GETPERSONNELNAME(W.PERSONNEL),	
			[Started]			=	W.[Started],	
			[Completed]			=	W.Completed,	
			[Hours Logged]		=	W.[Hours]
	FROM	VWJOBTICKETS J WITH(NOLOCK) JOIN JOB_TICKET_WORKFLOW W WITH(NOLOCK) ON J.JOB_NO = W.JOB_NO AND J.[SITE] = W.[SITE]       
	WHERE	W.[Started] BETWEEN @FROMDATE AND @TODATE 
			AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 AND JOB_TYPE LIKE @TYPE
	ORDER BY j.job_no
	

END
GO
/****** Object:  StoredProcedure [dbo].[RPT_CSFB_DAILY_AUDIT]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_DAILY_AUDIT]
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_DAILY_AUDIT
		Written On	:   01/14/2014 By Sudalai Mani
		Description	:   Produces data for daily audit		
	*/
		
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	
	SELECT	[Job #]				=	j.job_no,
			[Pool]				=	j.[site],	
			[Banker Name]		=	j.ACCT_EXEC,	
			[Submission Date]	=	SUBSTRING(CONVERT(VARCHAR, J.[REQUESTED_DATE], 101), 0, 13),			
			[Submission Time]	=	LEFT(CONVERT(VARCHAR(8), CAST(J.[REQUESTED_DATE] AS DATETIME), 108), 5),
			[Report Title]		=	JOB_TITLE,	
			[Due Date]			=	SUBSTRING(CONVERT(VARCHAR, J.[DUE_DATE], 101), 0, 13),	
			[Due Time]			=	LEFT(CONVERT(VARCHAR(8), CAST(J.[DUE_DATE] AS DATETIME), 108), 5), 
			[Business Purpose]	=	[ADDRESS],	
			[Project Code]		=	ALT01,
			[Client Code]		=	EML06,	
			[Recruitment Code]	=	ATTENTION,	
			[Division]			=	DEPARTMENT,	
			[Group]				=	JOB_GROUP,	
			[Location]			=	LOCATION,	
			[Cost Center]		=	CLIENT,	
			[Submitted By]		=	REQUESTED_BY,	
			[Status]			=	J.[status],		
			[Personnel]			=	DBO.FN_GETPERSONNELNAME(W.PERSONNEL),	
			[Started]			=	W.[Started],	
			[Completed]			=	W.Completed,	
			[Total Hours]		=	W.[Hours]
	FROM	VWJOBTICKETS J WITH(NOLOCK) JOIN JOB_TICKET_WORKFLOW W WITH(NOLOCK) ON J.JOB_NO = W.JOB_NO AND J.[SITE] = W.[SITE]       
	WHERE	W.[Started] BETWEEN @FROMDATE AND @TODATE AND J.[status] IN ('ON HOLD', 'Processing', 'CANCELLED', 'COMPLETED')
			AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 AND JOB_TYPE LIKE @TYPE
	ORDER BY j.job_no
	

END
GO
/****** Object:  StoredProcedure [dbo].[RPT_CSFB_COMPLETED_JOBS]    Script Date: 02/10/2014 05:52:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[RPT_CSFB_COMPLETED_JOBS]
	@SITES		VARCHAR(1024),
	@TYPE		VARCHAR(20),
	@FROMDATE	DATETIME,
	@TODATE		DATETIME,
	@USER		VARCHAR(50),
	@URL		VARCHAR(255)
AS
BEGIN
	/*
		Name		:	RPT_CSFB_COMPLETED_JOBS
		Written On	:   01/16/2014 By Muthamil Selvan
		Description	:   Produces data for Completed Jobs		
	*/
		
	--Do we have a site
    IF DATALENGTH(RTRIM(@SITES)) = 0 
    BEGIN
		RAISERROR ('SITE IS REQUIRED', 16, 1)
		RETURN
    END 
    
    --Do we have a type
    IF DATALENGTH(RTRIM(@TYPE)) = 0 
    BEGIN
		RAISERROR ('TYPE IS REQUIRED', 16, 1)
		RETURN
    END

	--Do we have a start date
    IF @FROMDATE IS NULL
    BEGIN
		RAISERROR ('STARTING DATE IS REQUIRED', 16, 1)
		RETURN
    END

    --Do we have an end date
    IF @TODATE IS NULL
    BEGIN
		RAISERROR ('ENDING DATE IS REQUIRED', 16, 1)
		RETURN
    END
	
	SET @SITES = REPLACE(@SITES, '''', '')
	--select * from events where site='pool1'
	--insert into events (site,comments) values('pool1',@SITES)
	SELECT	[Job #]				=	j.job_no,
			[Pool]				=	j.[site],	
			[Submission Date]	=	SUBSTRING(CONVERT(VARCHAR, J.[REQUESTED_DATE], 101), 0, 13),			
			[Submission Time]	=	LEFT(CONVERT(VARCHAR(8), CAST(J.[REQUESTED_DATE] AS DATETIME), 108), 5),
			[Report Title]		=	JOB_TITLE,	
			[Due Date]			=	SUBSTRING(CONVERT(VARCHAR, J.[DUE_DATE], 101), 0, 13),	
			[Due Time]			=	LEFT(CONVERT(VARCHAR(8), CAST(J.[DUE_DATE] AS DATETIME), 108), 5), 
			[Business Purpose]	=	[ADDRESS],
			[Completed Date]    =   SUBSTRING(CONVERT(VARCHAR, J.COMPLETED, 101), 0, 13),	
            [Completed Time]    =   SUBSTRING(CONVERT(VARCHAR, DATEADD(DAY, DATEDIFF(DAY, J.COMPLETED, 0), J.COMPLETED)), 12, 20),	
			[Project Code]		=	ALT01,
			[Client Code]		=	EML06,	
			[Recruitment Code]	=	ATTENTION,	
			[Banker Name]		=	j.ACCT_EXEC,
			[Banker ID]     	=   CUSTOMER,
			[Division]			=	DEPARTMENT,	
			[Group]				=	JOB_GROUP,	
			[Location]			=	LOCATION,	
			[Cost Center]		=	CLIENT,	
			[Submitted By]		=	REQUESTED_BY,	
			[Time Estimation]   =   COMPANY,
			[Total Hours]		=	SUM(W.Hours)
	FROM	VWJOBTICKETS J WITH(NOLOCK) JOIN JOB_TICKET_WORKFLOW W WITH(NOLOCK) ON J.JOB_NO = W.JOB_NO AND J.[SITE] = W.[SITE]       
	WHERE	J.[Completed] BETWEEN @FROMDATE AND @TODATE 
	        AND J.status = 'COMPLETED'
            AND DBO.FN_INSTRING(J.[SITE], @SITES, ',') = 1 AND JOB_TYPE LIKE @TYPE
   GROUP BY 
            J.job_no, J.site,J.[REQUESTED_DATE],JOB_TITLE,J.[DUE_DATE],ADDRESS,J.COMPLETED,
            ALT01, EML06, ATTENTION, J.ACCT_EXEC, CUSTOMER, DEPARTMENT, JOB_GROUP, LOCATION, CLIENT, REQUESTED_BY, COMPANY
	        
	ORDER BY J.job_no
	

END
GO
