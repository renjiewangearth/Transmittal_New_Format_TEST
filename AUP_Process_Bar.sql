USE [OPERATIONS]
GO

/****** Object:  StoredProcedure [dbo].[AUP_Process_Bar]    Script Date: 9/3/2020 2:25:23 PM ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO















/*
ALTER TABLE [dbo].[AUP_Report_Log_Bar] ADD  [Data_Source] VARCHAR(50) default(NULL);
ALTER TABLE [dbo].[AUP_Report_Log_Bar] ADD  [first_run] int default(0);
ALTER TABLE [dbo].[AUP_Report_Log_Bar] ADD  [extended_aff_pk] int default(NULL);
--
-------------------------------------------------------------------------------------
-- ALTER TABLE [dbo].[AUP_Report_Log_Bar] ADD  [Full_Potentials] int default(NULL);
-------------------------------------------------------------------------------------
sp_help [AUP_Report_Log_Bar]

*/
CREATE PROCEDURE [dbo].[AUP_Process_Bar] 
	@@row_count int = 0 OUTPUT,
	@@output_directory [varchar] (255) = '' OUTPUT,
	@load_id int = 0,
	@file_path varchar(200) 	= NULL,
	@file_name varchar(55)		= NULL,
	@file_delimiter [char](1) 	= NULL,
	@fmt_file_path varchar(200) 	= NULL,
	@fmt_file_name varchar(55) 	= NULL,
	@reporting_aff_pk int 		= NULL,
	@as_of_date datetime 		= NULL,	-- Default to GETDATE(),
	@is_partial int 		= 0,			-- Default to FULL
	@trackit_id  varchar(10) 	= '',	-- Default to Empty String
	@data_source varchar(50) 	= '',		-- Default to Empry string -- 'UWare'
	@first_run	 int 			= NULL,	-- Default to NULL
	@full_potentials int 		= NULL	-- Default to NULL
	
AS
BEGIN
	/***************************************************************
	History:
		-- 20070306 -- GP -- Append Load_ID to the Load_Tag 
		-- 20070307 -- GP -- Add a mini-Gap of 50 person_pk(s)/Load_ID
		-- 20070328 -- GP -- Save @is_partial information to AUP_Reports_Log
		-- 20070410 -- GP -- Add "AUP_rpt_Address_90_Days_Bypass"
		-- 20070502 -- GP -- Add parameter @trackit_id 
		-- 20070516 -- GP -- Add [edits_done] = 1
		-- 20070521 -- GP -- Add AUP_rpt_Inactive_Officers_Detail'	 -- invalid
		-- 20071022 -- GP -- Add AUP_rpt_Change_Phone_Summary
		-- 20071203 -- GP -- Add AUP_Fill_AUP_Aff_Tree_BAR
		-- 20080110 -- GP -- Add AUP_Rpt_Duplicate_ADDs
		-- 20080117 -- GP -- Add AUP_Rpt_Potential_Duplicate_Records
		-- 20080211 -- GP -- Add AUP_rpt_Change_Name_SSN_Detail
		-- 20090721--va addred @@output_directory  OUTPUT,
		-- 20160310 -- GP -- Use [dbo].[AUP_Add_Person_Member_Info_BAR] -- for testing 
		-- 20190806 -- GP/RW -- AUP_rpt_Inactive_Officers_Detail_w_IVP_BAR 
		-- 20190829 -- GP/RW -- AUP_UW_ID_Conflicting_with_PPK
	****************************************************************/
	--
	SET NOCOUNT ON
	--
	DECLARE @num_rows [int]
	DECLARE @return_value [int]
	DECLARE @aff_date_tag [varchar](50)
	DECLARE @output_path [varchar] (255)
	DECLARE	@sql_str [varchar](2000)
	DECLARE @reporting_aff_type [char](1);
	-------------------------------------------------------------------------------
	-- Required parameters should not be NULL -- They may be invalid, though!!!
	-------------------------------------------------------------------------------	
	IF (ISNULL(@file_delimiter, '') NOT IN ('|', ';', ',', 'T'))
	   BEGIN
		PRINT 'Error: Acceptable File delimiters are (''|'', '';'', '','', ''T'') i.e. (Vertical Bar, SemiColon, Comma, Tab). Please use one of them!'
		RETURN (-10)
	   END
	--
	IF (   @file_path IS NULL 
		OR @file_name IS NULL 
		OR @fmt_file_path IS NULL 
		OR @fmt_file_name IS NULL 
		OR @reporting_aff_pk IS NULL
		)
	   BEGIN
		PRINT 'Error: One of the input file parameters is null.'
		RETURN (-1)
	   END
	-- 
	IF (@reporting_aff_pk IS NULL)
	   BEGIN
		PRINT 'ERROR: Variable @reporting_aff_pk IS NULL'
		RETURN (-2)
	   END
	IF NOT EXISTS (SELECT 'True' FROM afscme_oltp6..aff_organizations WHERE aff_pk = @reporting_aff_pk)
	   BEGIN
		PRINT 'ERROR: Variable @reporting_aff_pk = '''+CAST(ISNULL(@reporting_aff_pk, 0) AS varchar)+''' is NOT present in aff_organizations.'
		RETURN (-3)
	   END
	--
	------------------------------------
	-- GET value of @reporting_aff_type 
	------------------------------------
	SET @reporting_aff_type = ISNULL((SELECT [aff_type] FROM [afscme_oltp6].[dbo].[Aff_Organizations] WHERE [aff_pk] = @reporting_aff_pk), '')
	IF  (@reporting_aff_type) NOT IN ('C','L', 'R','S')
	BEGIN
		PRINT 'ERROR: Variable @reporting_aff_pk = '''+CAST(ISNULL(@reporting_aff_pk, 0) AS varchar)+''' is NOT one of the types (''C'', ''L'',''R'', ''S'').  We do not support [aff_type] = '''+@reporting_aff_type+'''.'
		RETURN (-4)
	   END
	-- Set as_of_date field
	IF (@as_of_date IS NULL)
		SELECT @as_of_date = CONVERT(varchar, GETDATE(), 101)
		
	-- Get a "Load_Tag"
	SELECT  @aff_date_tag 	= '_' 
			+ CASE
				WHEN @reporting_aff_pk = 2626 THEN 'PA_C_13_NWHS'	-- This is C=13/NWHS PA -- Actually L = 1419
				ELSE aff_stateNat_type
					+ '_' + aff_type
					+ '_' + CASE aff_type
						WHEN 'C' THEN aff_councilRetiree_chap
						WHEN 'R' THEN aff_councilRetiree_chap
						WHEN 'L' THEN aff_localSubChapter
						WHEN 'S' THEN aff_localSubChapter
						ELSE aff_councilRetiree_chap+'_'+aff_localSubChapter+'_'+aff_subUnit
						END
				END
				+ CASE
					WHEN (@trackit_id IS NULL) THEN ''
					WHEN (@trackit_id = '') THEN ''
					ELSE '_'+LTRIM(@trackit_id)
					END
				+ '_' + REPLACE(CONVERT(varchar, @as_of_date, 102), '.', '')
	FROM 	afscme_oltp6.dbo.aff_organizations
	WHERE	aff_pk = @reporting_aff_pk
	--
	--------------------------------------------------------------------
	-- We do NOT want to load new data before we finished processing
	--------------------------------------------------------------------
	SELECT @num_rows = count(*) FROM [dbo].[AUP_RAW_BAR]
	IF (@num_rows > 0)
	   BEGIN
		PRINT 'There is data in table [AUP_RAW_BAR], cannot load new transmittal until previous is cleared.'
		RETURN (-666)
	   END
	-- 
	SELECT @num_rows = count(*) FROM [dbo].[AUP_Input_BAR]
	IF (@num_rows > 0)
	   BEGIN
		PRINT 'There is data in table [AUP_Input_BAR], cannot load new transmittal until previous is cleared.'
		RETURN (-666)
	   END
	-- 
	SELECT @num_rows = count(*) FROM [dbo].[AUP_Code_BAR]
	IF (@num_rows > 0)
	   BEGIN
		PRINT 'There is data in table [AUP_Code_BAR], cannot load new transmittal until previous is cleared.'
		RETURN (-666)
	   END
	-------------------------------------------------------------------------------
	-- Start Processing
	-------------------------------------------------------------------------------
	-- Make a Log Entry and get Load_ID
	----------------------------------------------
	-- 20170203 -- GP -- We want to use the same Root source for muliple processes
	INSERT INTO [dbo].[AUP_Report_Log_Root] ([ProcessType], [as_of_date], [Load_Tag],  [TrackIT_ID],             [reporting_aff_pk], [is_partial], [Server_Name])
	VALUES                                  ('BAR',          @as_of_date, @aff_date_tag, ISNULL(@trackit_id, ''), @reporting_aff_pk,  @is_partial, 'AFSSQL1604')
	SELECT @load_id = SCOPE_IDENTITY();
	-- 20070328 -- GP
	INSERT INTO [dbo].[AUP_Report_Log_Bar] ([Load_id],[reporting_aff_pk], [as_of_date], [Load_Tag],   [file_path], [file_name], [fmt_file_path], [fmt_file_name], [is_partial], [TrackIT_ID],          [File_Delimiter], [Data_Source], [first_run], [Full_Potentials]) 
	VALUES (                                 @load_id, @reporting_aff_pk, @as_of_date,  @aff_date_tag, @file_path, @file_name,  @fmt_file_path,  @fmt_file_name,  @is_partial,  ISNULL(@trackit_id, ''), @file_delimiter, @data_source, @first_run,   @full_potentials)
	-- Get value for Load_ID
	-- SELECT @load_id = MAX([Load_ID]) FROM [dbo].[AUP_Report_Log_Bar] WHERE reporting_aff_pk = @reporting_aff_pk AND as_of_date = @as_of_date
	--
PRINT '1ST POINT'
	IF (ISNULL(@load_id, 0) = 0)
	   BEGIN
		PRINT 'Unable to get value for parameter @load_id;  Aborting'
		RETURN (-666)
	   END
	--
	-- 20070306 -- GP-- Append Load_ID to the Load_Tag 
	UPDATE 	[dbo].[AUP_Report_Log_Bar]
	SET	[Load_Tag] = [Load_Tag]+'_'+CAST(@load_id as varchar)
	WHERE	[Load_ID] = @load_id
	----------------------------------------------
	-- Load RAW transmittal data
	----------------------------------------------
PRINT '2ND POINT'
	SELECT @@row_count = 0
	-- PRINT 'Calling EXEC @return_value = [dbo].[AUP_BulkInsertTo_RAW_BAR]'
	EXEC @return_value = [dbo].[AUP_BulkInsertTo_RAW_BAR_2016_1]
					@@row_count OUTPUT, 
					@load_id,
					@file_path 		= @file_path,
					@file_name 		= @file_name,
					@fmt_file_path 		= @fmt_file_path,
					@fmt_file_name 		= @fmt_file_name,
					@reporting_aff_pk 	= @reporting_aff_pk,
					@as_of_date 		= @as_of_date,
					@file_delimiter		= @file_delimiter
	--
PRINT '3RD POINT'
	IF (ISNULL(@@row_count, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure  ; 0 records loaded. ' ,
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1001)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[loaded_to_raw] = 1,
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
	   END
	-------------------------------------------------------------------------------
	-- At this time we can create a dedicated directory for this transmittal
	--    The problem will be how to deal w. it on a Re-Start
	-------------------------------------------------------------------------------
PRINT '4TH POINT'
	SELECT  @output_path = output_path FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id
	-- 
	SELECT  @sql_str = 'MKDIR "'+ @output_path+'"'
	EXEC master..xp_cmdshell @sql_str
	SELECT  @sql_str = 'MOVE /Y "'+ @file_path+'\'+@file_name+'" "'+@output_path+'"'
--	SELECT  @sql_str = 'COPY /Y "'+ @file_path+'\'+@file_name+'" "'+@output_path+'"'  -- TESTING
	EXEC master..xp_cmdshell @sql_str
PRINT '5TH POINT'	
	SELECT @@output_directory = @output_path	-- va 20090721
	-------------------------------------------------------------------------
	-- Fill table AUP_Aff_Tree -- 20071203 -- GP --
	-------------------------------------------------------------------------
	EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR] @load_id, @reporting_aff_pk
	-- EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_NUHHCE_Subunit2] @load_id, @reporting_aff_pk	-- RW 2018/06/13: For full transmittal of subunit 2 of NUHHCE only
	-- EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_NUHHCE_Subunit3] @load_id, @reporting_aff_pk	-- RW 2018/06/13: For full transmittal of subunit 3 of NUHHCE only
	-- EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_NUHHCE_Subunit7] @load_id, @reporting_aff_pk	-- RW 2020/07/27: For full transmittal of subunit 7 of NUHHCE only
	-- EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_MI_C25_L1659_Sub11] @load_id, @reporting_aff_pk	-- RW 2019/03/26: For full transmittal of MI C25 L1659 Subunit 11 only
	-- EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_C65] @load_id, @reporting_aff_pk				-- RW: 2018/07/26: For transmittal of C65 with R65 excluded
	 --EXEC @return_value = [dbo].[AUP_Fill_AUP_Aff_Tree_BAR_C37_6Locals_Only] @load_id, @reporting_aff_pk	-- RW 2019/12/11: Temp for TRANS-825 - a Full transmittal for only 6 locals under NY C37

	--
PRINT '6TH POINT'
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Fill_AUP_Aff_Tree_BAR; 0 records loaded in Table AUP_Aff_Tree_BAR, or Else. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1002)
	   END
PRINT '7TH POINT'	
---------------------
--
	----------------------------------------------
	-- Move Data to Input and Input_WIP
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_MoveDataFromRawToWIP_BAR] @load_id, @reporting_aff_type
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_MoveDataFromRawToWIP_BAR; 0 records loaded or Else. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1002)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[moved_to_input] =1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
	-- 
	----------------------------------------------
	-- Edit Data in Input
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_Edit_WIP_Data_BAR] @load_id, @reporting_aff_pk, @reporting_aff_type
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Edit_WIP_Data_BAR. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1003)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[edits_done] = 1,		-- 20070516 -- GP 
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
	--
	----------------------------------------------
	-- Add Desired Affiliate information
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_Add_Affiliate_Info_BAR] @load_id
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Add_Affiliate_Info_BAR. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1003)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[aff_data_done] = 1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
	-- 
/*	
	----------------------------------------------
	-- Add Desired Extended Affiliate information
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_Add_Extended_Affiliate_Info_BAR] @load_id
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Add_Extended_Affiliate_Info_BAR. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1013)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[aff_data_done] = 1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
*/
	-- 
	----------------------------------------------
	-- Add Desired Person/Member/Transaction information
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_Add_Person_Member_Info_BAR] @load_id, @is_partial
	-- EXEC @return_value = [dbo].[AUP_Add_Person_Member_Info_BAR_CA_L3930] @load_id, @is_partial			-- RW 4/23/2018: for L3930
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Add_Person_Member_Info_BAR. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1004)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[person_data_done] = 1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
	-- 
	------------------------------------------------------------------
	-- Dump the Match and Transactions reports -- And hope it works!!!
	------------------------------------------------------------------
DECLARE @ret_val 		[int];
DECLARE @var0	[sql_variant];
-- declare @output_path [VARCHAR](100);
-- declare @aff_date_tag [VARCHAR](100);
--
SET @var0 = CAST(CAST((@output_path + '\Rpt_Change_Name_SSN_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Rpt_Change_Name_SSN_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Change_SSN_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Change_SSN_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_ADDs_Detail_Bar' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_ADDs_Detail_Bar'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Duplicate_Invalid_Records_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Rpt_Duplicate_Invalid_Records_BAR'
,@package_name		= 'Std_Rpt_Package_NEW_GP1.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Match_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Match_Detail_BAR'
,@package_name		= 'Std_Rpt_Package_NEW_GP1.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Match_Summary_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Match_Summary_BAR'
,@package_name		= 'Std_Rpt_Package_NEW_GP1.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Member_Counts_Affiliate_File_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Member_Counts_Affiliate_File_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Member_Counts_Enterprise_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Member_Counts_Enterprise_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Potential_Duplicate_Records_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Rpt_Potential_Duplicate_Records_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Transactions_Summary_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Transactions_Summary_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Multiple_PPKs_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Multiple_PPKs_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Failed_Match_on_PPK_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Failed_Match_on_PPK_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Update_Person_Info_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Update_Person_Info_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Update_Person_Addr_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Update_Person_Addr_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Inactive_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Inactive_Detail_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Inactive_Officers_Detail_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_rpt_Inactive_Officers_Detail_w_IVP_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\Rpt_Duplicate_ADDs_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Rpt_Duplicate_ADDs_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
SET @var0 = CAST(CAST((@output_path + '\Rpt_InValid_Transactions_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Rpt_Invalid_Transactions_Bar'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
--
-- (19 row(s) affected)
--
SET @var0 = CAST(CAST((@output_path + '\Rpt_Match_to_Deceased_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name        = 'BAR_TAB_SMC_Transmittal'
,@project_name       = 'AUP_Rpt_Match_to_Deceased_BAR'
,@package_name       = 'Package.dtsx'
,@parameter_name     = 'FinalDestination'
,@parameter_value    = @var0
,@object_type        = 30
,@reference_id       = NULL
,@use32bitruntime    = False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
--
SET @var0 = CAST(CAST((@output_path + '\Rpt_UW_ID_Conflicting_with_PPK_BAR' + @aff_date_tag +'.CSV') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name        = 'BAR_TAB_SMC_Transmittal'
,@project_name       = 'AUP_UW_ID_Conflicting_with_PPK'
,@package_name       = 'Package.dtsx'
,@parameter_name     = 'FinalDestination'
,@parameter_value    = @var0
,@object_type        = 30
,@reference_id       = NULL
,@use32bitruntime    = False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
--
--  (17/19 row(s) affected)

/*  
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_ADDs_Detail_Bar'				, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_Change_Name_SSN_Detail_BAR'	, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Change_SSN_Detail_BAR'		, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Change_Suffix_Detail_BAR'		, NO_OUTPUT
--	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Change_SSN_Summary_BAR'		, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_Duplicate_ADDs_BAR'			, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_Duplicate_Invalid_Records_BAR', NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Inactive_Detail_BAR'			, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Inactive_Officers_Detail_BAR'	, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_Invalid_Transactions_BAR'		, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Match_Detail_BAR'				, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Match_Summary_BAR'				, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Member_Counts_Affiliate_File_BAR' , NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Member_Counts_Enterprise_BAR'	, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_Potential_Duplicate_Records_BAR', NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Transactions_Summary_BAR'		, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Update_Person_Addr_Detail_BAR', NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Update_Person_Info_Detail_BAR', NO_OUTPUT
--	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Change_Phone_Summary_BAR'		, NO_OUTPUT
--	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Change_Suffix_Summary_BAR'	, NO_OUTPUT
--	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Rpt_ADD_DEL_Match'				, NO_OUTPUT
--	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Address_90_Days_Bypass'	, NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Multiple_PPKs_Detail_BAR', NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_rpt_Failed_Match_on_PPK_BAR', NO_OUTPUT
*/
	--------------------------
	-- Add other reports Here
	--------------------------
/* */
	UPDATE 	[dbo].[AUP_Report_Log_Bar]
	SET	[invalid_reports_done] = 1,
		[edit_reports_done] = 1,
		[aff_reports_done] = 1,
		[match_reports_done] = 1,
		[lst_mod_dt] = GETDATE()
	WHERE	[Load_ID] = @load_id
	--	
	----------------------------------------------
	-- Generate Code
	----------------------------------------------
	-- 20070307 -- GP -- Add a mini-Gap of 50 person_pk(s)/Load_ID
	EXEC @return_value = [dbo].[AUP_Generate_Code_BAR] @load_id
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_Generate_Code_BAR. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 0, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1005)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[code_generated] =1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id
	   END
	--
	
	----------------------------------------------
	-- Dump the Code -- And hope it works!!!
	----------------------------------------------

--
SET @var0 = CAST(CAST((@output_path + '\AUP_Generated_Code_BAR' + @aff_date_tag +'.SQL') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Generated_Code_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--

SET @var0 = CAST(CAST((@output_path + '\AUP_Generated_Code_Summary_BAR' + @aff_date_tag +'.TXT') AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'AUP_Generated_Code_Summary_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'FinalDestination'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
,@use32bitruntime	= False;
--
SELECT @ret_val = CAST(@ret_val AS [VARCHAR])
--
/*
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Generated_Code_BAR'	 , NO_OUTPUT
	EXEC master..xp_cmdshell 'DTSRUN /SSQL05Reports /E /NAUP_Generated_Code_Summary_BAR', NO_OUTPUT
*/
	--
	UPDATE 	[dbo].[AUP_Report_Log_Bar]
	SET	[code_downloaded] = 1,
		[lst_mod_dt] = GETDATE()
	WHERE	[Load_ID] = @load_id
	--
	----------------------------------------------
	-- Move Data to HIST
	----------------------------------------------
	EXEC @return_value = [dbo].[AUP_MoveDataTo_Hist_BAR] @load_id
	--
	IF (ISNULL(@return_value, 0) <= 0)
	   BEGIN
		-- Flag Log entry as invalid and EXIT
		UPDATE 	[dbo].[AUP_Report_Log_Bar] 
		SET 	[is_valid] = 0,
			[comment] = [comment]+'-- Error in procedure AUP_MoveDataTo_Hist. ',
			[lst_mod_dt] = GETDATE()
		WHERE 	[Load_ID] = @load_id
		--
		EXEC [dbo].[AUP_Clean_Up_BAR] @load_id = @load_id, @clean_up_hist = 1, @clean_up_log = 0, @clean_up_curr = 1
		--
		RETURN (-1006)
	   END
	ELSE
	   BEGIN
		UPDATE 	[dbo].[AUP_Report_Log_Bar]
		SET	[moved_to_hist] = 1,
			[lst_mod_dt] = GETDATE()
		WHERE	[Load_ID] = @load_id	   
	   END
	--
	-------------------------------------------------------------------------------
	-- End Processing
	-------------------------------------------------------------------------------
	IF (@load_id IS NULL)
		SELECT @load_id = 0
	--
	RETURN @load_id
	--
	SET NOCOUNT OFF
END
/*
USAGE:
USAGE:
select * from  Aup_raw_bar
truncate table Aup_raw_bar
select * from aup_input_bar
---
	-- USE [OPERATIONS]
	-- 
	SET NOCOUNT ON
	DECLARE @curr_date datetime
    SELECT  @curr_date = CONVERT(varchar, GETDATE(), 101)
	-- select getdate()
	-- go
	DECLARE @@row_count int
    declare @@output_directory  varchar(255)
	DECLARE @return_value int
	EXEC @return_value = [dbo].[AUP_Process_Bar]
		@@row_count OUTPUT, 
        @@output_directory  OUTPUT,
		@load_id 			= 0,
		@file_path 			= 'D:\OLTP4\Transmittals\Files',
		@file_name 			= 'tix6869_R32_full.txt',			-- 'BAR_Layout_File_new.BAR',		-- 'BAR_Layout_File_Sample1.txt',
		@file_delimiter		= 'T',
		@fmt_file_path 		= 'NotUsed',
		@fmt_file_name 		= 'NotUsed',
		@reporting_aff_pk 	= 6848,
		@as_of_date 		= @curr_date,
		@is_partial			= 0,
		@trackit_id			= 'GP_R32TT'	-- Default to Empty String
		@data_source		= 'UWare'	-- Default to Empry string -- ws 'UWare'
		@first_run			= 0			-- Default to NULL		   -- vs @first_run	= 0	-- first run for 'UWare'
	SELECT '@return_value' = @return_value, '@@row_count' = @@row_count
	IF (@return_value <= 0) 
		PRINT 'Procedure [dbo].[AUP_Process_Bar] == ERROR'
	ELSE
		PRINT 'Procedure [dbo].[AUP_Process_Bar] == OK'
	PRINT '----------------------------------------'
	SELECT '@@output_directory = '+@@output_directory
	SET NOCOUNT OFF
	-- go
	-- select getdate()
	-- go
----------------------------------------
@return_value @@row_count
------------- -----------
5776          449

Procedure [dbo].[AUP_Process_Bar] == OK
----------------------------------------

---
USE [OPERATIONS]
GO

DECLARE @RC int
DECLARE @load_id int
DECLARE @clean_up_hist int
DECLARE @clean_up_log int
DECLARE @clean_up_curr int

-- TODO: Set parameter values here.

SELECT * FROM AUP_REPORT_LOG_BAR

EXECUTE @RC = [dbo].[AUP_Clean_Up_BAR] 
   @load_id = 5190
  ,@clean_up_hist =1
  ,@clean_up_log  =0
  ,@clean_up_curr =1
GO

DECLARE @ret_val 		[int];
DECLARE @var0	[sql_variant];
-- declare @output_path [VARCHAR](100);
-- declare @aff_date_tag [VARCHAR](100);
--
SET @var0 = CAST(CAST('D:\OLTP4\Transmittals\Files\Bug6850_WARetiree10-6849_TAB.txt' AS [nvarchar](1000)) AS sql_variant)
EXEC @ret_val = [OPERATIONS].[dbo].[Run_STD_Rpt_Package_NEW_GP]
 @folder_name		= 'BAR_TAB_SMC_Transmittal'
,@project_name		= 'Load_TAB_File_to_RAW_BAR'
,@package_name		= 'Package.dtsx'
,@parameter_name	= 'OriginalFile'
,@parameter_value	= @var0
,@object_type		= 30
,@reference_id		= NULL
*/
GO

