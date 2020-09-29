USE [OPERATIONS]
GO

/****** Object:  StoredProcedure [dbo].[AUP_Generate_Code_BAR]    Script Date: 9/21/2020 2:52:55 PM ******/
SET ANSI_NULLS OFF
GO

SET QUOTED_IDENTIFIER ON
GO














ALTER PROCEDURE [dbo].[AUP_Generate_Code_BAR] @load_id int = 0
AS
BEGIN
	/*************************************************************************************
	Description:
		Generates code to be executed later at ACC/7Space.
	------------------------------------------------------------------------------------	
	History:
		-- 20070108 GP  -- Set Person.valid_ssn_fg = 0 if LEFT(SSN, 4) = '0000'
				-- Add code to update Person.SSN
		-- 20070118 GP  -- Improve "do_not_mail" handling
				-- Add political_do_not_call_fg for Agency Fee Payers
		-- 20070223 GP  -- Do not Request Cards for anyone who is not on v_Member_Cards
		-- 20070326 GP  -- Replaced T200.middle_name w. "REPLACE(T200.Middle_Name, '''', '''''')"
		-- 20070412 GP  -- If we have no adds, set @min_person_pk and=@max_person_pk = 0 -- ZERO
		-- 20070416 GP  -- Log Execution of code to AUP_Job_Log table
		-- 20070427 GP  -- We do not preallocate person_pk(s) -- Replace all w. EXEC AUP_Insert_Person_Member
		-- 20070507 GP  -- Date_Joined set to @as_of_date WHERE not available
		-- 20070516 GP  -- Made changes to exclude from Card Requests for "Name Changes" the ones requested for "Not Having Membership Record in desired Affiliate"
		-- 20070524 GP  -- Report Card Requests by category
		-- 20070827 GP  -- Add code for [insert_PA] and [update_SMA]
		-- 20070925 GP  -- Set Person_Address.end_dt = NULL when we update an existing address
		-- 20071025 GP  -- Add MLBP_Persons Maintenance for OH_L_11 (@reporting_aff_pk = 711)
		-- 20071203 GP  -- Replace v_aff_tree_AdminC w. [AUP_Aff_Tree_BAR]
						-- Removed params. @gap and @min_person_pk 
		-- 20071210	GP  -- Deal w. ADDs having multiple memberships (i.e. One Person record, multiple Aff_Member records)
		-- 20080110 GP  -- Adjust phone_prmry_fg for existence of non_home primary_phone
		-- 20080208 GP  -- Add fields [duplicate_ADDs] and [potential_Dups] to AUP_Input_WIN
		-- 20080214 GP  -- Comment out "-- T200.[duplicate_ID] IS NULL" for Simple ADDs 
		-- 20080227 GP  -- Switch Change of Status w. Change of Type for Cards generation
		-- 20080501 GP  -- Add phone_bad_fg/phone_marked_bad_dt to Person_Phone
		-- 20080509 GP  -- Add code for [update_officer_addr] = 1
		-- 20090727 GP  -- Inactivate all Active Memberships which are NOT in the file for Existing People AND NOT "Should Be A Member"
		-- 20090903 GP  -- Add Person_pk = 12826192 (PAC Migration) to the list of 90 Days Address Update NON-Exclusion
		-- 20110209 GP  -- Keep in Sync Officer_History.[pos_address_from_person_pk] w. Current_SMA
						-- But do NOT update [Officer_History].[lst_mod_dt] since it will trigger a Card
		-- 20110503 GP  -- Add Update Person_Email (Home and Work)
		-- 20110824 GP  -- Fix Single Quotes in Email_ Home and Work
		-- 20121012 GP  -- Inactivate all Active Memberships which are NOT in the file for Existing People AND NOT "Staff" Either
		--              -- AND NOT "Should Be A Member"
		--              -- AND NOT "Staff"
		-- 20140323 GP  -- ADD "Potential Member'
		-- 20140522 GP  -- Make sure all 'P's match the 'N's.
		-- 20140624 GP	-- Make sure we have valid person_pkS when inserting addresses
		---------------------------------------------------------------------------------
		-- 20160524 -- GP -- ADD MULTIPLE PHONES -- but NO OTHER Phone
		-- 20160526 -- GP -- ADD MISC DATA       -- but NOT EmployerCode
		-- 20161025 -- GP -- ADD DUBLE SINGLE QUOTES IN MISC DATA
		-- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
		-- 20170116 -- GP -- ADD DateOfBirth / dob
		-- 20170131 -- GP -- Duplicate ADDs -- We want the same Person, but different Misc_Data_BAR 
		-- 20170215 -- GP -- sp_update_Misc_Data_Bar uses @lst_mod_user_pk and @lst_mod_dt
		-- 20170328 -- GP -- For Membership Change of an Existing Person -- Only request cards for (Regular, Retiree, Union Shop, Retiree Spouse)
		-- 20170328 -- GP -- [Status_Code] in  IN ('A', 'N', 'X', 'P', 'O', 'Y')
		-- 20170814 -- GP -- We want to keep the full [varchar](50) Addresses
		-- 20180731 -- RW -- Request Cards based on NewCard = 1 in Custom field (UnionWare)
		-- 20190227 -- RW -- Added code to exit automatically if this script is being executed in AFSCME_OLTP4 on AFSSQL_1604
		-- 20190227 -- GP -- Supports @data_source = 'UWare'
		--				  -- If @data_source  = 'UWare'
		--				  --	If @first_run  = 0  -- We blank out all existing [mbr_no_local], so we only use NEW UWare IDs
		--				  --   						-- We update [mr_no_local] = [Affiliate_Member_ID]
		--				  --    If @first_run <> 0  -- We DO NOT UPDATE [mr_no_local] = [Affiliate_Member_ID] (unless [mr_no_local] is Blank or Null)
		--				  --   
		-- 20190521 -- GP --    If @first_run  = 0  -- We blank out all existing [mbr_no_local] 
		--                --       we Cover ALL affiliates even is not active, excluded, etc.  WE WANT A CLEAN SLATE
		-- 20190826 -- RW -- Changed 'CONVERT(varchar, @as_of_date, 101)' to 'CAST(getdate() AS date)' whenever updating table COM_Weekly_Mbr_Card_Run
							 so that the date when transmittal is moved to production is avaliable when processing weekly member cards
		-- 20190826 -- RW -- Changed 'CONVERT(varchar, @as_of_date, 101)' to 'CAST(getdate() AS date)' whenever inserting table Aff_Members
							 so that the date when transmittal is moved to production is avaliable when processing weekly member cards
							 (requred by sp_Member_Type_Conversion_by_Transmittal of Impact Analysis)
							 
		-- 20191021 -- GP  -- Inactivate and Activate Potential Members, if the @full_potentials Flag = 1
		-- 20200810 -- RW  -- Propagate UW_ID/AffiliateMemberID through entire affiliate tree for UWare transmittal
		-- 20200825 -- RW  -- If we first-run full transmittal for a subunit only (such as subunit 7 of NUHHCE), we wipe mbr_no_local under the subunit only, 
						   -- rathern than under the entire parent local
	*************************************************************************************/
	SET NOCOUNT ON
	--
	TRUNCATE TABLE [dbo].[AUP_Code_BAR]
	--
	DECLARE @reporting_aff_pk int 
	DECLARE @as_of_date datetime
	DECLARE @NSQLString NVARCHAR (2000)
	DECLARE @load_tag varchar(50)			-- 20070416 GP
	DECLARE @trackit_id [varchar](10)		-- 20070416 GP
	DECLARE @data_source varchar(50);		-- 20190218 GP
	DECLARE @first_run int;				-- 20190218 GP
	DECLARE @full_potentials int;	-- 20191021 -- GP -- Inactivate and Activate Potential Members, if the @full_potentials Flag = 1
	-- 
	-- Required parameters should be NOT NULL
	IF (@load_id IS NULL)
	   BEGIN
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'ERROR: Variable @load_id should be NOT NULL'
		RETURN -1
	   END
	IF NOT EXISTS (SELECT 'True' FROM AUP_Report_Log_Bar WHERE [Load_ID] = @load_id)
	   BEGIN
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'ERROR: [Load_ID] =  = '''+CAST(@load_id AS varchar)+ ''' is NOT present in AUP_Report_Log_Bar.'
		RETURN -2
	   END
	IF NOT EXISTS (SELECT 'True' FROM AUP_RAW_BAR WHERE [Load_ID] = @load_id)
	   BEGIN
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'ERROR: [Load_ID] =  = '''+CAST(@load_id AS varchar)+ ''' is NOT present in AUP_RAW_BAR.'
		RETURN -3
	   END
	IF NOT EXISTS (SELECT 'True' FROM AUP_Input_BAR WHERE [Load_ID] = @load_id)
	   BEGIN
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'ERROR: [Load_ID] =  = '''+CAST(@load_id AS varchar)+ ''' is NOT present in AUP_Input_BAR.'
		RETURN -4
	   END
	--
	SELECT @reporting_aff_pk = reporting_aff_pk FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1
	SELECT @as_of_date       = as_of_date       FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1
	SELECT @load_tag         = Load_Tag         FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1	-- 20070416 GP
	SELECT @trackit_id       = TrackIT_ID       FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1	-- 20070416 GP
	SELECT @full_potentials  = COALESCE(Full_Potentials, 0)  FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1	-- 20191021 GP
	-- SELECT @full_potentials = 1;
	--
	SELECT @data_source	 = Data_Source      FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1
	SELECT @first_run	 = first_run        FROM [dbo].[AUP_Report_Log_Bar] WHERE [Load_ID] = @load_id AND [is_valid] = 1
	--
	----------------------------------------------------------------------
	----------------------------------------------------------------------
	----------------------------------------------------------------------
	
	SET NOCOUNT ON
	-- PK into Time_Dim for membership activity statistics
	DECLARE @time_pk_as_of_date int
	SELECT  @time_pk_as_of_date = time_pk 
	FROM 	afscme_oltp6.dbo.Time_Dim 
	WHERE	calendar_year = YEAR(@as_of_date) 
	AND	calendar_month = MONTH(@as_of_date)
	
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'USE afscme_oltp4'

	-- Exit immediately if running on afscme_oltp4 on afssql1604 since the DB is in sync
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'IF @@SERVERNAME = ''AFSSQL1604'' AND DB_NAME() = ''afscme_oltp4'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, ' BEGIN '
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, ' PRINT ''ERROR: Please use DB AFSCME_OLTP5 since ASFCME_OLTP4 is in sync with other DBs.  This script is not executed.'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, ' RAISERROR(''Please use DB AFSCME_OLTP5 since ASFCME_OLTP4 is in sync with other DBs.'', 20, -1) with log'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, ' GOTO NOEXECUTION'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, ' END'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'IF NOT EXISTS (SELECT * FROM AUP_Job_Log WHERE [Load_ID] = '+CAST(@load_id AS varchar)+')'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'BEGIN'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
	--								 -- 20070416 GP
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT ON
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO AUP_Job_Log (Job_Type, Load_ID, reporting_aff_pk, as_of_date, Load_Tag, TrackIT_ID, posted) VALUES ('
			 + '''Affiliate Transmittal'''
		+ ', '   + CAST(@load_id AS varchar)
		+ ', '   + CAST(@reporting_aff_pk AS varchar)
		+ ', ''' + CAST(@as_of_date AS varchar(20)) + ''''
		+ ', ''' + @load_tag + ''''
		+ ', ''' + @trackit_id + ''''
		+ ', 0)'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	---------------------------------------------------------------------------------------

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Counts numbers of active members by type before the transmittal and generates report later (only on AFSSQL1604 and AFSSQL1604\AFSCME_TEST)'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'IF @@SERVERNAME IN (''AFSSQL1604'', ''AFSSQL1604\AFSCME_TEST'')'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    EXEC [OPERATIONS].[dbo].[sp_Member_Count_Before_&_After_Transmittal] @load_id = ' + CAST(@load_id AS varchar) + ', @reporting_aff_pk = ' + CAST(@reporting_aff_pk AS varchar) + ', @before_after = ' + '''before''' + ', @view_only = 0, @suppress_output = 1'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	---------------------------------------------------------------------------------------
	-- Inserts and Updates as needed
	---------------------------------------------------------------------------------------
	---------------------------------------------------------------------------------------
	--- Person
	--  Users
	--- Person_address
	--- Person_SMA
	--- Person_demographics
	--- Person_Phone
	--- Person_Email
	--- Person_Political_Legislative
	--- Aff_Members
	----------------------------------------------------
	-- We need a talble to store the Stats
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'TRUNCATE TABLE [OPERATIONS].[dbo].[AUP_temp_stats_BAR]'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'TRUNCATE TABLE [OPERATIONS].[dbo].[AUP_temp_stats_unique_BAR]'
	--
	-- Person_Member Insert
	DECLARE @membershipDept int,
		@PersonAddressType int,
		@PhoneTypeHome int,
		@Active int,
		@Temporary int,
		@Regular int,
		@Retiree int,
		@AgencyFeePayer int,
		@UnionShop int,
		@RetireeSpouse int,
		@AssociateMember int,
		@PotentialMember [int],		-- 20140323 -- GP
		@ActivityType_Add varchar(10),
		@OptOutMember 	[int],		-- 20170328 -- GP
		@Pending 	[int]		-- 20170328 -- GP
		
	SET @membershipDept 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd   = 'AR'   AND com_cd_type_key = 'Department')	
	SET @PersonAddressType 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_desc = 'Home' AND com_cd_type_key = 'PersonAddressType')
	SET @PhoneTypeHome 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_desc = 'Home' AND com_cd_type_key = 'PhoneType')
	
	SET @Active 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'A'  AND com_cd_type_key = 'MemberStatus')
	SET @Temporary 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'T'  AND com_cd_type_key = 'MemberStatus')
	SET @Regular 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'R'  AND com_cd_type_key = 'MemberType')
	SET @Retiree 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'T'  AND com_cd_type_key = 'MemberType')
	SET @AgencyFeePayer 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'AF' AND com_cd_type_key = 'MemberType') 
	SET @UnionShop 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'U'  AND com_cd_type_key = 'MemberType')	
	SET @RetireeSpouse 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'S'  AND com_cd_type_key = 'MemberType') 
	SET @AssociateMember	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'AS' AND com_cd_type_key = 'MemberType')
	SET @PotentialMember 	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'PM' AND com_cd_type_key = 'MemberType')	-- 20140323 -- GP
	SET @OptOutMember	= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'OO' AND com_cd_type_key = 'MemberType')	-- 20170328 -- GP
	SET @Pending 		= (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd = 'P'  AND com_cd_type_key = 'MemberStatus')	-- 20170328 -- GP
	SET @ActivityType_Add 	= CAST((SELECT com_cd_pk FROM afscme_oltp6.dbo.common_codes WHERE com_cd_desc = 'Add' AND com_cd_type_key = 'ActivityType') AS varchar)

	
	-- We want to print the counts
	-- -- SET NOCOUNT OFF
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT ON
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'DECLARE @last_person_pk [int]'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT  @first_person_pk = 1+MAX(person_pk) FROM Person'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'  
-- 
	-------------------------------------------------------------------------------------------------
	-- If @first_run  = 0  -- We blank out all existing [mbr_no_local], so we only use NEW UWare ID
	-- 20190521 -- GP --    If @first_run  = 0  -- We blank out all existing [mbr_no_local] 
	--                --       we Cover ALL affiliates even is not active, excluded, etc.  WE WANT A CLEAN SLATE
	-------------------------------------------------------------------------------------------------
	IF (ISNULL(@data_source, '') = 'UWare') AND @first_run  = 0
	BEGIN
		-- print 'GCode UWare';
		-- INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE [dbo].[AFF_MEMBERS] SET [MBR_NO_LOCAL] = '''' , lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE aff_pk = '+ CAST([aff_pk] AS VARCHAR
		-- FROM  [dbo].[AUP_Aff_Tree_BAR]
		-- INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of [AFF_MEMBERS] SET [MBR_NO_LOCAL] = '''' = '+ CAST(@@ROWCOUNT AS varchar)
		----------------------------------------------------------------------------------------------
		-- 20190521 -- GP --    If @first_run  = 0  -- We blank out all existing [mbr_no_local] 
		--                --       we Cover ALL affiliates even is not active, excluded, etc.  WE WANT A CLEAN SLATE
		----------------------------------------------------------------------------------------------
		--
		DROP TABLE IF EXISTS #AUP_Aff_Tree_extended_BAR
		--
	   	SELECT  [aff_pk], [Parent_aff_pk], [GParent_aff_pk],[Root_aff_pk]
		INTO	#AUP_Aff_Tree_extended_BAR
	   	FROM 	afscme_oltp6.dbo.[V_aff_tree] 
	   	WHERE 	GParent_aff_pk   = @reporting_aff_pk 
		OR 		Parent_aff_pk    = @reporting_aff_pk 
		OR 		aff_pk           = @reporting_aff_pk
		--
		-- Remove RetChapter if  MN_65
		IF @reporting_aff_pk = 4598
			DELETE FROM #AUP_Aff_Tree_extended_BAR	WHERE [aff_pk] = 6863
			--
			-- SELECT * FROM #AUP_Aff_Tree_extended_BAR
			--

		-- We may do First_Run for a single Subunit of NUHHCE (aff_pk = 207 ) where we have to restrict the affiliate tree to a single subunit -- RW 8/25/2020
		IF @reporting_aff_pk = 207				-- If we run full transmittal for only a single subunit of NUHHCE, the aff_tree should have been restrited to the subunit only -- RW 8/25/2020
			DELETE FROM #AUP_Aff_Tree_extended_BAR WHERE [aff_pk] NOT IN (SELECT aff_pk FROM [OPERATIONS].[dbo].[AUP_Aff_Tree_BAR] WHERE Load_ID = @load_id) -- RW 8/25/2020
												-- If we run full transmittal for entire NUHHCE, the above block of code does nothing so it does not hurt 
		

		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code])  
		SELECT @load_id, 'UPDATE [dbo].[AFF_MEMBERS] SET [MBR_NO_LOCAL] = '''' , lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE aff_pk = '+ CAST([aff_pk] AS VARCHAR)
		FROM	#AUP_Aff_Tree_extended_BAR
		--
		DROP TABLE IF EXISTS #AUP_Aff_Tree_extended_BAR
		--
	END    
	-- print '[AUP_Insert_Person_Member_BAR]';                                      
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id,  'EXEC [dbo].[AUP_Insert_Person_Member_BAR] '
	-- Person
	+ ' @last_person_pk OUTPUT'
	+ ', @prefix_nm='	+ CAST(ISNULL((SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Prefix' AND Legacy_Code = T200.Title), 0) AS varchar)
	+ ', @first_nm='	+''''+RTRIM(CAST(REPLACE(T200.First_Name, '''', '''''') AS varchar)) +''''
	+ ', @middle_nm='	+''''+ RTRIM(CAST(REPLACE(T200.Middle_Name, '''', '''''') AS varchar)) +''''
	+ ', @last_nm='		+''''+ RTRIM(CAST(REPLACE(T200.Last_Name, '''', '''''') AS varchar)) +''''
	+ ', @suffix_nm='	+ CAST(ISNULL((SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Suffix' AND Legacy_Code = T200.Suffix), 0) AS varchar)
	+ ', @alternate_mailing_nm='''''
	+ ', @ssn='	+ CASE
				WHEN T200.SSN IS NULL THEN 'NULL'
				WHEN T200.SSN = '000000000' THEN 'NULL'
				WHEN T200.SSN = '' THEN 'NULL'
				ELSE ''''+T200.SSN+''''
				END
	+ ', @valid_ssn_fg='	+ CASE
					WHEN T200.SSN IS NULL THEN '0'
					WHEN T200.SSN = '000000000' THEN '0'
					WHEN LEFT(T200.SSN, 4) = '0000' THEN '0'
					WHEN T200.SSN = '' THEN '0'
					ELSE '1'
					END
	+ ', @duplicate_ssn_fg='	+ CAST(ISNULL(T200.duplicate_ssn_fg, 0) AS varchar)
	+ ', @mbr_barred_fg = 0'	
	+ ', @marked_for_deletion_fg = 0'
	+ ', @member_fg = 1'	
	+ ', @created_user_pk='	+'11949451'
	+ ', @created_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
	+ ', @lst_mod_user_pk='	+'11949451'
	+ ', @lst_mod_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
	+ ', @person_mst_lst_mod_user_pk='	+'11949451'
	+ ', @person_mst_lst_mod_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
	-- Person_Address
	+ ', @addr1='	+''''+ CAST(REPLACE(LTRIM(RTRIM(LEFT(ISNULL(Addr1, ''), 50))), '''', '''''') AS varchar(50)) +''''
	+ ', @addr2='	+''''+ CAST(REPLACE(LTRIM(RTRIM(LEFT(ISNULL(Addr2, ''), 50))), '''', '''''') AS varchar(50)) +''''
	+ ', @city='	+''''+ CAST(REPLACE(LTRIM(RTRIM(ISNULL(City, ''))), '''', '''''') AS varchar) +''''
	+ ', @state='	+''''+ CASE
				WHEN	State IS NULL THEN ''
				WHEN 	State IN ('ZZ', 'XX') THEN ''
				WHEN    (LEN(LTRIM(RTRIM(State)))=2) THEN State
				ELSE	''
				END +''''
	+ ', @zipcode='+''''+ CASE 
				WHEN Zip IS NULL THEN ''
				WHEN Zip = '00000' THEN ''
				ELSE LTRIM(RTRIM(Zip ))
				END + ''''
	+ ', @zip_plus='+''''+ CASE
				WHEN Zip_4 IS NULL THEN ''
				WHEN Zip_4 = '0000' THEN ''
				ELSE LTRIM(RTRIM(Zip_4))
				END + ''''
	+ ', @eff_dt='	+''''+CONVERT(varchar, @as_of_date, 101) +''''
	+ ', @dept='	+ CAST(@membershipDept AS varchar) 
	+ ', @addr_type='	+ CAST(@PersonAddressType AS varchar)
	+ ', @addr_bad_fg='+ CASE
				WHEN 	Addr_Mailable_fg IS NULL THEN '1'
				WHEN 	Addr_Mailable_fg = 'Y'   THEN '0'
				WHEN 	Addr_Mailable_fg = 'N'   THEN '1'
				-- WHEN 	(Addr1 = '' AND Addr2 = '')  THEN '1'
				ELSE	'1'
				END 
	+ ', @addr_marked_bad_dt=' + CASE 
					WHEN 	Addr_Mailable_fg IS NULL THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
					WHEN 	Addr_Mailable_fg <> 'Y'  THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
					-- WHEN 	(Addr1 = '' AND Addr2 = '')  THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
					ELSE	'NULL'
					END
	+ ', @addr_source=''U'''
	+ ', @addr_source_if_aff_apply_upd='+''''+CAST(@reporting_aff_pk AS varchar) +''''
	+ ', @addr_prmry_fg = 1'
	-- Phone
	-- 20160524 -- GP -- ADD MULTIPLE PHONES
	+ ', @home_phone='''+T200.Phone+''''
	/*
	+ ', @country_cd=''1'''
	+ ', @area_code='+''''+ LEFT(LTRIM(T200.Phone), 3)+''''
	+ ', @phone_no=' +''''+ SUBSTRING(LTRIM(T200.Phone), 4, 7)+''''
	+ ', @phone_prmry_fg=1'
	+ ', @phone_type='+ CAST(@PhoneTypeHome AS varchar)
	*/
	+ ', @work_phone='''+T200.[WorkPhone]+''''
	+ ', @cell_phone='''+T200.[CellPhone]+''''
	+ ', @fax_number='''+T200.[FaxNumber]+''''
--	+ ', @other_phone='''+ T200.[OtherPhone]+''''
	-- Person_Email
	+ ', @email_type_home=71001'
	+ ', @email_type_work=71002'
	-- Person_Demographics
	+ ', @gender='+ISNULL((SELECT CAST(com_cd_pk AS varchar) FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Gender' AND Legacy_Code = T200.Sex_cd), 'NULL') 
	+ ', @dob =' +  (CASE WHEN T200.[DateOfBirth] = '' THEN 'NULL' 
						 ELSE ''''+T200.[DateOfBirth]+''''
						 END)
	-- Person_Political_Legislative
	+ ', @political_do_not_call_fg='	+ CASE
						WHEN T200.Status_Code = 'N' THEN '1'
						WHEN T200.Status_Code = 'P' THEN '1'
						WHEN T200.Status_Code = 'O' THEN '1'	-- 20170328 GP
						WHEN T200.Status_Code = 'Y' THEN '1'	-- 20170328 GP
						ELSE '0'
						END  
	+ ', @political_party='	+CAST(CASE
					WHEN T200.Political_Party IS NULL THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
					WHEN T200.Political_Party = 'U' THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
					ELSE (SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'PoliticalParty' AND Legacy_Code = T200.Political_Party)
					END AS varchar)
	+ ', @political_registered_voter='	+ ISNULL((SELECT CAST(com_cd_pk AS varchar) FROM afscme_oltp6.dbo.common_codes WHERE com_cd_type_key = 'RegisteredVoter' AND com_cd_cd = T200.Registered_Voter_fg), 'NULL')
	-- Aff_Members
	+ ', @aff_pk='	+ CAST(T200.desired_aff_pk AS varchar)
	+ ', @mbr_status='	+ CAST(CASE
				WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('A','O', 'P') THEN @Active
				WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Temporary
				WHEN T200.Status_Code = 'N' THEN @Active   	
				WHEN T200.Status_Code = 'P' THEN @Active							-- 20130323 -- Add 'Potential Members'
				WHEN T200.Affiliate_Identifier = 'C' THEN @Active
				WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code = 'X' THEN @Active
				WHEN T200.Status_Code = 'Y' THEN @Pending							-- 20170328 -- GP
				ELSE 0 
				END AS varchar)
	+ ', @mbr_type='	+ CAST(CASE
				WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A', 'Y') THEN @Regular	-- 20170328 -- GP
				WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A', 'O') THEN @Retiree	-- 20170328 -- GP
				WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('T', 'R') THEN @Regular 	
				WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Retiree 	
				WHEN T200.Status_Code = 'N' THEN @AgencyFeePayer
				WHEN T200.Status_Code = 'C' THEN @UnionShop
				WHEN T200.Status_Code = 'P' THEN @PotentialMember						-- 20130323 -- Add 'Potential Members'
				WHEN T200.Status_Code = 'O' THEN @OptOutMember							-- 20130323 -- Add 'Potential Members'
				WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN @RetireeSpouse	
				-- WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code = 'X' THEN @AssociateMember
				ELSE 0
				END AS varchar)
	+ ', @no_mail_fg='	+ CAST(CASE T200.No_Mail_fg WHEN '9' THEN 1 ELSE 0 END AS varchar)
	+ ', @no_cards_fg='	+ CASE
				WHEN ISNULL((SELECT unit_wide_no_mbr_cards_fg FROM afscme_oltp6.dbo.Aff_Mbr_Rpt_Info WHERE aff_pk = T200.desired_aff_pk), 0) = 1 THEN '1'
				WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') THEN  '1'						-- 20130323 -- Add 'Potential Members'
				WHEN T200.No_Mail_fg = '1' THEN '1'
				WHEN T200.No_Mail_fg = '3' THEN '1'
				ELSE '0'
				END
	+ ', @no_public_emp_fg='	+ CASE										-- no_public_emp_fg
					WHEN ISNULL((SELECT unit_wide_no_pe_mail_fg FROM afscme_oltp6.dbo.Aff_Mbr_Rpt_Info WHERE aff_pk = T200.desired_aff_pk), 0) = 1 THEN '1'
			 		WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN '1'
				 	WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')  THEN '1'				-- 20170328 -- GP
					WHEN T200.No_Mail_fg = '2' THEN '1'
					WHEN T200.No_Mail_fg = '3' THEN '1'
					ELSE '0'
					END
	+ ', @no_legislative_mail_fg='	+ CASE							-- no_legislative_mail_fg
						WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')  THEN '1'		-- 20170328 -- GP	
						ELSE '0'
						END
	+ ', @mbr_join_dt='''	+ CASE
					WHEN T200.Date_Joined = '' THEN CONVERT(varchar, @as_of_date, 112)	-- 20070507 GP
					ELSE T200.Date_Joined	-- 20170113 GP+VA
					END+''''
	+ ', @mbr_no_local='	+ ''''+ CASE
						WHEN (T200.Affiliate_Member_ID IS NULL) THEN ''
						WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = '' THEN ''
						ELSE LTRIM(RTRIM(T200.Affiliate_Member_ID))
						END + ''''
	+ ', @primary_information_source='	+ CASE
							WHEN (T200.[Information_Source] IS NULL) THEN 'NULL'
							WHEN  T200.[Information_Source] = '' THEN 'NULL'
							WHEN  T200.[Information_Source] = '4' THEN '47004'
							ELSE  '47008'
							END
	+ ', @activityType='+@ActivityType_Add
	-- 20110503 GP  -- Add Update Person_Email (Home and Work)
	+ ', @home_email='''+REPLACE(LTRIM(RTRIM(T200.[Home_Email])), '''', '''''')+''''
	+ ', @work_email='''+REPLACE(LTRIM(RTRIM(T200.[Work_Email])), '''', '''''')+''''
	-- 20160526 -- GP -- ADD MISC DATA
	+ ', @JobTitle='''			+  REPLACE(T200.[JobTitle], '''', '''''')+''''
	+ ', @JobSector='  			+  (CASE WHEN T200.[JobSector] = '' THEN 'NULL' ELSE CAST(T200.[JobSector] AS [varchar])  END)
	+ ', @EmplSector='  			+  (CASE WHEN T200.[EmplSector]= '' THEN 'NULL' ELSE CAST(T200.[EmplSector] AS [varchar]) END)
	+ ', @JobHireDate='			+  (CASE WHEN T200.[JobHireDate] = '' THEN 'NULL' 
											 ELSE ''''+T200.[JobHireDate]+''''	-- 20170113 GP+VA -- PROPER CONVERSION FOR Dates
											 END)
	+ ', @WorkSiteName='''			+  REPLACE(T200.[WorkSiteName] , '''', '''''')+''''
	+ ', @WorkSiteAddr1='''			+  REPLACE(T200.[WorkSiteAddr1], '''', '''''')+''''
	+ ', @WorkSiteAddr2='''			+  REPLACE(T200.[WorkSiteAddr2], '''', '''''')+''''
	+ ', @WorkSiteCity='''			+  REPLACE(T200.[WorkSiteCity] , '''', '''''')+''''
	+ ', @WorkSiteState='''			+  T200.[WorkSiteState]+''''
	+ ', @WorkSiteZip5='''			+  T200.[WorkSiteZip5]+''''
	+ ', @WorkSiteZip4='''			+  T200.[WorkSiteZip4]+''''
	+ ', @EmployerName='''			+  REPLACE(T200.[EmployerName] , '''', '''''')+''''
	+ ', @EmployerAddr1='''			+  REPLACE(T200.[EmployerAddr1], '''', '''''')+''''
	+ ', @EmployerAddr2='''			+  REPLACE(T200.[EmployerAddr2], '''', '''''')+''''
	+ ', @EmployerCity='''			+  REPLACE(T200.[EmployerCity] , '''', '''''')+''''
	+ ', @EmployerState='''			+  T200.[EmployerState]+''''
	+ ', @EmployerZip5='''			+  T200.[EmployerZip5]+''''
	+ ', @EmployerZip4='''			+  T200.[EmployerZip4]+''''
	+ ', @SalaryRange='				+  (CASE WHEN T200.[SalaryRange]= '' THEN 'NULL' ELSE CAST(T200.[SalaryRange]  AS [varchar]) END)
	+ ', @SalaryType='''			+  REPLACE(T200.[SalaryType], '''', '''''')+''''
	+ ', @SalaryAmount='  			+  (CASE	WHEN T200.[SalaryAmount] = '' THEN 'NULL' ELSE T200.[SalaryAmount] END)
	+ ', @ElectedLeadershipTitle='''		+  REPLACE(T200.[ElectedLeadershipTitle], '''', '''''')+''''
	+ ', @ElectedLeadershipElectionDate='	+  (CASE WHEN T200.[ElectedLeadershipElectionDate] = '' THEN 'NULL' 
													 ELSE ''''+T200.[ElectedLeadershipElectionDate]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
													 END)
	+ ', @Steward='''				+  T200.[Steward]+''''
	+ ', @Activist='''				+  T200.[Activist]+''''
	+ ', @EnterpriseID='''			+  T200.[EnterpriseID]+''''
	+ ', @PEOPLEContributionAmount='		+  (CASE 	WHEN T200.[PEOPLEContributionAmount] = '' THEN 'NULL' ELSE T200.[PEOPLEContributionAmount] END)
	+ ', @PEOPLEContributionPayPeriod=' 	+  (CASE	WHEN T200.[PEOPLEContributionPayPeriod] = '' THEN 'NULL' 
														ELSE ''''+T200.[PEOPLEContributionPayPeriod]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
														END)	-- NEW
	+ ', @PEOPLEContributionFrequency='''	+  REPLACE(T200.[PEOPLEContributionFrequency], '''', '''''')+''''
	+ ', @PEOPLECheckIssuer='''		+  REPLACE(T200.[PEOPLECheckIssuer], '''', '''''')+''''
	+ ', @CustomName1='''			+  REPLACE(T200.[CustomName1] , '''', '''''')+''''
	+ ', @CustomValue1='''			+  REPLACE(T200.[CustomValue1], '''', '''''')+''''
	+ ', @CustomName2='''			+  REPLACE(T200.[CustomName2] , '''', '''''')+''''
	+ ', @CustomValue2='''			+  REPLACE(T200.[CustomValue2], '''', '''''')+''''
	+ ', @CustomName3='''			+  REPLACE(T200.[CustomName3] , '''', '''''')+''''
	+ ', @CustomValue3='''			+  REPLACE(T200.[CustomValue3], '''', '''''')+''''
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.[Tran_Type] = 'A' 
	AND 	T200.[Insert_P] = 1
	-- AND		T200.[duplicate_ID] IS NULL		-- 20080214 GP -- We could have Potential Duplicates	-- There are no duplicate ADDs -- 20071210 -- GP
	AND		T200.[duplicate_ADDs] IS NULL		-- 20080208 GP
	AND 	T200.desired_aff_pk IS NOT NULL
	ORDER BY T200.desired_aff_pk, T200.[ID]
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person Inserts Generated AS Simple ADDs = '+ CAST(@@ROWCOUNT AS varchar)
	--
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--	
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT ON
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--' 
			-- print '[dbo].[AUP_Insert_Person_Member_BAR]';                                           
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id,  CASE
			WHEN T200.[duplicate_ID] = T200.[ID] 	--------------------------------------------> WE want Person_Member for the first Record (MIN([ID])
				THEN 	'EXEC [dbo].[AUP_Insert_Person_Member_BAR] '
					-- Person
					+ ' @last_person_pk OUTPUT'
					+ ', @prefix_nm='	+ CAST(ISNULL((SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Prefix' AND Legacy_Code = T200.Title), 0) AS varchar)
					+ ', @first_nm='	+''''+RTRIM(CAST(REPLACE(T200.First_Name, '''', '''''') AS varchar)) +''''
					+ ', @middle_nm='	+''''+ RTRIM(CAST(REPLACE(T200.Middle_Name, '''', '''''') AS varchar)) +''''
					+ ', @last_nm='		+''''+ RTRIM(CAST(REPLACE(T200.Last_Name, '''', '''''') AS varchar)) +''''
					+ ', @suffix_nm='	+ CAST(ISNULL((SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Suffix' AND Legacy_Code = T200.Suffix), 0) AS varchar)
					+ ', @alternate_mailing_nm='''''
					+ ', @ssn='	+ CASE
								WHEN T200.SSN IS NULL THEN 'NULL'
								WHEN T200.SSN = '000000000' THEN 'NULL'
								WHEN T200.SSN = '' THEN 'NULL'
								ELSE ''''+T200.SSN+''''
								END
					+ ', @valid_ssn_fg='	+ CASE
									WHEN T200.SSN IS NULL THEN '0'
									WHEN T200.SSN = '000000000' THEN '0'
									WHEN LEFT(T200.SSN, 4) = '0000' THEN '0'
									WHEN T200.SSN = '' THEN '0'
									ELSE '1'
									END
					+ ', @duplicate_ssn_fg='	+ CAST(ISNULL(T200.duplicate_ssn_fg, 0) AS varchar)
					+ ', @mbr_barred_fg = 0'	
					+ ', @marked_for_deletion_fg = 0'
					+ ', @member_fg = 1'	
					+ ', @created_user_pk='	+'11949451'
					+ ', @created_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
					+ ', @lst_mod_user_pk='	+'11949451'
					+ ', @lst_mod_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
					+ ', @person_mst_lst_mod_user_pk='	+'11949451'
					+ ', @person_mst_lst_mod_dt='	+''''+CONVERT(varchar, @as_of_date, 101)+''''
					-- Person_Address
					+ ', @addr1='	+''''+ CAST(REPLACE(LTRIM(RTRIM(LEFT(ISNULL(Addr1, ''), 50))), '''', '''''') AS varchar(50)) +''''
					+ ', @addr2='	+''''+ CAST(REPLACE(LTRIM(RTRIM(LEFT(ISNULL(Addr2, ''), 50))), '''', '''''') AS varchar(50)) +''''
					+ ', @city='	+''''+ CAST(REPLACE(LTRIM(RTRIM(ISNULL(City, ''))), '''', '''''') AS varchar) +''''
					+ ', @state='	+''''+ CASE
								WHEN	State IS NULL THEN ''
								WHEN 	State IN ('ZZ', 'XX') THEN ''
								WHEN    (LEN(LTRIM(RTRIM(State)))=2) THEN State
								ELSE	''
								END +''''
					+ ', @zipcode='+''''+ CASE 
								WHEN Zip IS NULL THEN ''
								WHEN Zip = '00000' THEN ''
								ELSE LTRIM(RTRIM(Zip ))
								END + ''''
					+ ', @zip_plus='+''''+ CASE
								WHEN Zip_4 IS NULL THEN ''
								WHEN Zip_4 = '0000' THEN ''
								ELSE LTRIM(RTRIM(Zip_4))
								END + ''''
					+ ', @eff_dt='	+''''+CONVERT(varchar, @as_of_date, 101) +''''
					+ ', @dept='	+ CAST(@membershipDept AS varchar) 
					+ ', @addr_type='	+ CAST(@PersonAddressType AS varchar)
					+ ', @addr_bad_fg='+ CASE
								WHEN 	Addr_Mailable_fg IS NULL THEN '1'
								WHEN 	Addr_Mailable_fg = 'Y'   THEN '0'
								WHEN 	Addr_Mailable_fg = 'N'   THEN '1'
								-- WHEN 	(Addr1 = '' AND Addr2 = '')  THEN '1'
								ELSE	'1'
								END 
					+ ', @addr_marked_bad_dt=' + CASE 
									WHEN 	Addr_Mailable_fg IS NULL THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
									WHEN 	Addr_Mailable_fg <> 'Y'  THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
									-- WHEN 	(Addr1 = '' AND Addr2 = '')  THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
									ELSE	'NULL'
									END
					+ ', @addr_source=''U'''
					+ ', @addr_source_if_aff_apply_upd='+''''+CAST(@reporting_aff_pk AS varchar) +''''
					+ ', @addr_prmry_fg = 1'
					-- Phone
					-- 20160524 -- GP -- ADD MULTIPLE PHONES
					+ ', @home_phone='''+T200.Phone+''''
					/*
					+ ', @country_cd=''1'''
					+ ', @area_code='+''''+ LEFT(LTRIM(T200.Phone), 3)+''''
					+ ', @phone_no=' +''''+ SUBSTRING(LTRIM(T200.Phone), 4, 7)+''''
					+ ', @phone_prmry_fg=1'
					+ ', @phone_type='+ CAST(@PhoneTypeHome AS varchar)
					*/
					+ ', @work_phone='''+T200.[WorkPhone]+''''
					+ ', @cell_phone='''+T200.[CellPhone]+''''
					+ ', @fax_number='''+T200.[FaxNumber]+''''
--					+ ', @other_phone='''+ T200.[OtherPhone]+''''
					-- Person_Email
					+ ', @email_type_home=71001'
					+ ', @email_type_work=71002'
					-- Person_Demographics
					+ ', @gender='+ISNULL((SELECT CAST(com_cd_pk AS varchar) FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'Gender' AND Legacy_Code = T200.Sex_cd), 'NULL') 
					+ ', @dob =' +  (CASE WHEN T200.[DateOfBirth] = '' THEN 'NULL' 
										  ELSE ''''+T200.[DateOfBirth]+''''
										  END)
					-- Person_Political_Legislative
					+ ', @political_do_not_call_fg='	+ CASE
										WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')  THEN '1'		-- 20170328 -- GP
										ELSE '0'
										END  
					+ ', @political_party='	+CAST(CASE
									WHEN T200.Political_Party IS NULL THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
									WHEN T200.Political_Party = 'U' THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
									ELSE (SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'PoliticalParty' AND Legacy_Code = T200.Political_Party)
									END AS varchar)
					+ ', @political_registered_voter='	+ ISNULL((SELECT CAST(com_cd_pk AS varchar) FROM afscme_oltp6.dbo.common_codes WHERE com_cd_type_key = 'RegisteredVoter' AND com_cd_cd = T200.Registered_Voter_fg), 'NULL')
					-- Aff_Members
					+ ', @aff_pk='	+ CAST(T200.desired_aff_pk AS varchar)
					+ ', @mbr_status='	+ CAST(CASE
								WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('A', 'O', 'P') THEN @Active	-- 20170328 -- GP
								WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Temporary 	
								WHEN T200.Status_Code = 'N' THEN @Active  
								WHEN T200.Status_Code = 'P' THEN @Active									-- 20130323 -- Add 'Potential Members'
								WHEN T200.Affiliate_Identifier = 'C' THEN @Active
								WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code = 'X' THEN @Active
								ELSE 0 
								END AS varchar)
					+ ', @mbr_type='	+ CAST(CASE
								WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A','O') THEN @Regular
								WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A','O') THEN @Retiree
								WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('T', 'R') THEN @Regular 	
								WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Retiree 		
								WHEN T200.Status_Code = 'N' THEN @AgencyFeePayer
								WHEN T200.Status_Code = 'C' THEN @UnionShop
								WHEN T200.Status_Code = 'P' THEN @PotentialMember					-- 20130323 -- Add 'Potential Members'
								WHEN T200.Status_Code = 'O' THEN @OptOutMember						-- 20130323 -- Add 'Potential Members'
								WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN @RetireeSpouse	
								-- WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code = 'X' THEN @AssociateMember
								ELSE 0
								END AS varchar)
					+ ', @no_mail_fg='	+ CAST(CASE T200.No_Mail_fg WHEN '9' THEN 1 ELSE 0 END AS varchar)
					+ ', @no_cards_fg='	+ CASE
								WHEN ISNULL((SELECT unit_wide_no_mbr_cards_fg FROM afscme_oltp6.dbo.Aff_Mbr_Rpt_Info WHERE aff_pk = T200.desired_aff_pk), 0) = 1 THEN '1'
								WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') THEN  '1'								-- 20161031  GP
								WHEN T200.No_Mail_fg = '1' THEN '1'
								WHEN T200.No_Mail_fg = '3' THEN '1'
								ELSE '0'
								END
					+ ', @no_public_emp_fg='	+ CASE										-- no_public_emp_fg
									WHEN ISNULL((SELECT unit_wide_no_pe_mail_fg FROM afscme_oltp6.dbo.Aff_Mbr_Rpt_Info WHERE aff_pk = T200.desired_aff_pk), 0) = 1 THEN '1'
									WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN '1'
									WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') THEN '1'				-- 20170328 -- GP		
									WHEN T200.No_Mail_fg = '2' THEN '1'
									WHEN T200.No_Mail_fg = '3' THEN '1'
									ELSE '0'
									END
					+ ', @no_legislative_mail_fg='	+ CASE										-- no_legislative_mail_fg
									WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') THEN '1'			-- 20170328 -- GP	
									ELSE '0'
									END
					+ ', @mbr_join_dt='''	+ CASE
									WHEN T200.Date_Joined = '' THEN CONVERT(varchar, @as_of_date, 101)	-- 20070507 GP
									ELSE T200.Date_Joined	-- 20170113 GP+VA
									END+''''
					+ ', @mbr_no_local='	+ ''''+ CASE
										WHEN (T200.Affiliate_Member_ID IS NULL) THEN ''
										WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = '' THEN ''
										ELSE LTRIM(RTRIM(T200.Affiliate_Member_ID))
										END + ''''
					+ ', @primary_information_source='	+ CASE
											WHEN (T200.[Information_Source] IS NULL) THEN 'NULL'
											WHEN  T200.[Information_Source] = '' THEN 'NULL'
											WHEN  T200.[Information_Source] = '4' THEN '47004'
											ELSE  '47008'
											END
					+ ', @activityType='+@ActivityType_Add
					-- 20110503 GP  -- Add Update Person_Email (Home and Work)
					+ ', @home_email='''+REPLACE(LTRIM(RTRIM(T200.[Home_Email])), '''', '''''')+''''
					+ ', @work_email='''+REPLACE(LTRIM(RTRIM(T200.[Work_Email])), '''', '''''')+''''
					-- 20160526 -- GP -- ADD MISC DATA
					+ ', @JobTitle='''			+  REPLACE(T200.[JobTitle], '''', '''''')+''''
					+ ', @JobSector='  			+  (CASE WHEN T200.[JobSector] = '' THEN 'NULL' ELSE CAST(T200.[JobSector] AS [varchar])  END)
					+ ', @EmplSector='  			+  (CASE WHEN T200.[EmplSector]= '' THEN 'NULL' ELSE CAST(T200.[EmplSector] AS [varchar]) END)
					+ ', @JobHireDate='			+  (CASE WHEN T200.[JobHireDate] = '' THEN 'NULL' 
														      ELSE ''''+T200.[JobHireDate]+''''	-- 20170113 GP+VA -- PROPER CONVERSION FOR Dates
														      END)
					+ ', @WorkSiteName='''			+  REPLACE(T200.[WorkSiteName] , '''', '''''')+''''
					+ ', @WorkSiteAddr1='''			+  REPLACE(T200.[WorkSiteAddr1], '''', '''''')+''''
					+ ', @WorkSiteAddr2='''			+  REPLACE(T200.[WorkSiteAddr2], '''', '''''')+''''
					+ ', @WorkSiteCity='''			+  REPLACE(T200.[WorkSiteCity] , '''', '''''')+''''
					+ ', @WorkSiteState='''			+  T200.[WorkSiteState]+''''
					+ ', @WorkSiteZip5='''			+  T200.[WorkSiteZip5]+''''
					+ ', @WorkSiteZip4='''			+  T200.[WorkSiteZip4]+''''
					+ ', @EmployerName='''			+  REPLACE(T200.[EmployerName] , '''', '''''')+''''
--					+ ', @EmployerCode='  			+  (CASE WHEN T200.[EmployerCode]= '' THEN 'NULL' ELSE CAST(T200.[EmployerCode] AS [varchar]) END)
					+ ', @EmployerAddr1='''			+  REPLACE(T200.[EmployerAddr1], '''', '''''')+''''
					+ ', @EmployerAddr2='''			+  REPLACE(T200.[EmployerAddr2], '''', '''''')+''''
					+ ', @EmployerCity='''			+  REPLACE(T200.[EmployerCity], '''', '''''')+''''
					+ ', @EmployerState='''			+  T200.[EmployerState]+''''
					+ ', @EmployerZip5='''			+  T200.[EmployerZip5]+''''
					+ ', @EmployerZip4='''			+  T200.[EmployerZip4]+''''
					+ ', @SalaryRange='				+  (CASE WHEN T200.[SalaryRange]= '' THEN 'NULL' ELSE CAST(T200.[SalaryRange]  AS [varchar]) END)
					+ ', @SalaryType='''			+  REPLACE(T200.[SalaryType], '''', '''''')+''''
					+ ', @SalaryAmount='  			+  (CASE	WHEN T200.[SalaryAmount] = '' THEN 'NULL' ELSE T200.[SalaryAmount] END)
					+ ', @ElectedLeadershipTitle='''		+  REPLACE(T200.[ElectedLeadershipTitle], '''', '''''')+''''
					+ ', @ElectedLeadershipElectionDate='	+  (CASE WHEN T200.[ElectedLeadershipElectionDate] = '' THEN 'NULL' 
																	 ELSE ''''+T200.[ElectedLeadershipElectionDate]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
																	 END)
					+ ', @Steward='''				+  T200.[Steward]+''''
					+ ', @Activist='''				+  T200.[Activist]+''''
					+ ', @EnterpriseID='''			+  T200.[EnterpriseID]+''''
					+ ', @PEOPLEContributionAmount='		+  (CASE	WHEN T200.[PEOPLEContributionAmount] = '' THEN 'NULL' ELSE T200.[PEOPLEContributionAmount] END)
					+ ', @PEOPLEContributionPayPeriod=' 	+  (CASE	WHEN T200.[PEOPLEContributionPayPeriod] = '' THEN 'NULL' 
																		ELSE ''''+T200.[PEOPLEContributionPayPeriod]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
																		END)	-- NEW
					+ ', @PEOPLEContributionFrequency='''	+  REPLACE(T200.[PEOPLEContributionFrequency], '''', '''''')+''''
					+ ', @PEOPLECheckIssuer='''		+  REPLACE(T200.[PEOPLECheckIssuer], '''', '''''')+''''
					+ ', @CustomName1='''			+  REPLACE(T200.[CustomName1] , '''', '''''')+''''
					+ ', @CustomValue1='''			+  REPLACE(T200.[CustomValue1], '''', '''''')+''''
					+ ', @CustomName2='''			+  REPLACE(T200.[CustomName2] , '''', '''''')+''''
					+ ', @CustomValue2='''			+  REPLACE(T200.[CustomValue2], '''', '''''')+''''
					+ ', @CustomName3='''			+  REPLACE(T200.[CustomName3] , '''', '''''')+''''
					+ ', @CustomValue3='''			+  REPLACE(T200.[CustomValue3], '''', '''''')+''''
			ELSE	--------------------------------------------> WE want ONLY the Membership Record for ALL but the first	
					--------------------------------------------> -- 20170131 -- GP -- Duplicate ADDs -- We want the same Person, but different Misc_Data_BAR 
					'  INSERT INTO Aff_Members (person_pk, aff_pk, mbr_status, mbr_type,  no_mail_fg, no_cards_fg, no_public_emp_fg, no_legislative_mail_fg, mbr_join_dt, mbr_no_local, primary_information_source, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
						+ '@last_person_pk, '
						+ CAST(T200.desired_aff_pk AS varchar) + ', '
						+ CAST(CASE
							WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('A', 'O', 'P') THEN @Active
							WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Temporary 	
							WHEN T200.Status_Code = 'N' THEN @Active 														-- 20130323 -- Add 'Potential Members'
							WHEN T200.Affiliate_Identifier = 'C' THEN @Active
							WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code = 'X' THEN @Active
							WHEN T200.Status_Code = 'Y' THEN @Pending					-- 20170328 -- GP
							ELSE 0 
							END AS varchar) + ', '	
						+ CAST(CASE
							WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A', 'Y') THEN @Regular	-- 20170328 -- GP
							WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A'     ) THEN @Retiree
							WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('T', 'R') THEN @Regular 	
							WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Retiree 	
							WHEN T200.Status_Code = 'N' THEN @AgencyFeePayer
							WHEN T200.Status_Code = 'C' THEN @UnionShop
							WHEN T200.Status_Code = 'P' THEN @PotentialMember						-- 20130323 -- Add 'Potential Members'
							WHEN T200.Status_Code = 'O' THEN @OptOutMember							-- 20170328 -- GP					-- 20130323 -- Add 'Potential Members'
							WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN @RetireeSpouse	
							-- WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code = 'X' THEN @AssociateMember	
							ELSE 0
							END AS varchar) + ', '
						+ CAST(CASE T200.No_Mail_fg WHEN '9' THEN 1 ELSE 0 END AS varchar) + ', '	-- no_mail_fg
						+ CASE										-- no_cards_fg
							WHEN ISNULL(MRI.unit_wide_no_mbr_cards_fg, 0) = 1 THEN '1'
							WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'		-- 20170328 -- GP	-- 20130323 -- Add 'Potential Members'
							WHEN T200.No_Mail_fg  = '1'                       THEN '1'
							WHEN T200.No_Mail_fg  = '3'                       THEN '1'
							ELSE '0'
							END+', '
						+ CASE										-- no_public_emp_fg
							WHEN ISNULL(MRI.unit_wide_no_pe_mail_fg, 0) = 1   THEN '1'
							WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'		-- 20170328 -- GP
							WHEN T200.No_Mail_fg  = '2'                       THEN '1'
							WHEN T200.No_Mail_fg  = '3'                       THEN '1'
							ELSE '0'
							END+', '
						+ CASE										-- no_legislative_mail_fg
							WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'		-- 20170328 -- GP
							ELSE '0'
							END+', '
						+ CASE
							WHEN T200.Date_Joined = '' THEN ''''+ CONVERT(varchar, @as_of_date, 101) +''''		-- 20070507 GP
							ELSE ''''+T200.Date_Joined+''''	-- 20170113 GP+VA
							END + ', ' 
						+ ''''+ CASE
								WHEN (ISNULL(T200.Affiliate_Member_ID, '') = '') THEN ''
								WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = '' THEN ''
								ELSE LTRIM(RTRIM(T200.Affiliate_Member_ID))
								END + ''', '
						+ CASE
							WHEN (T200.[Information_Source] IS NULL) THEN 'NULL'
							WHEN  T200.[Information_Source] = '' THEN 'NULL'
							WHEN  T200.[Information_Source] = '4' THEN '47004'
							ELSE  '47008'
							END + ', '
						+ '11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''');'+'
						EXEC [dbo].[sp_Insert_Misc_Data_BAR] '
						--+   '@person_pk= CAST(@last_person_pk AS [varchar])'		-- R.W. 5/10/2017: modified
						+   '@person_pk= @last_person_pk'
						+ ', @aff_pk='+CAST(T200.desired_aff_pk AS varchar)
						+ ', @JobTitle='''				+  REPLACE(T200.[JobTitle], '''', '''''')+''''
						+ ', @JobSector='  				+  (CASE WHEN T200.[JobSector] = '' THEN 'NULL' ELSE CAST(T200.[JobSector] AS [varchar])  END)
						+ ', @EmplSector='  			+  (CASE WHEN T200.[EmplSector]= '' THEN 'NULL' ELSE CAST(T200.[EmplSector] AS [varchar]) END)
						+ ', @JobHireDate='				+  (CASE WHEN T200.[JobHireDate] = '' THEN 'NULL' 
																 ELSE ''''+T200.[JobHireDate]+''''	-- 20170113 GP+VA -- PROPER CONVERSION FOR Dates
																 END)
						+ ', @WorkSiteName='''			+  REPLACE(T200.[WorkSiteName] , '''', '''''')+''''
						+ ', @WorkSiteAddr1='''			+  REPLACE(T200.[WorkSiteAddr1], '''', '''''')+''''
						+ ', @WorkSiteAddr2='''			+  REPLACE(T200.[WorkSiteAddr2], '''', '''''')+''''
						+ ', @WorkSiteCity='''			+  REPLACE(T200.[WorkSiteCity] , '''', '''''')+''''
						+ ', @WorkSiteState='''			+  T200.[WorkSiteState]+''''
						+ ', @WorkSiteZip5='''			+  T200.[WorkSiteZip5]+''''
						+ ', @WorkSiteZip4='''			+  T200.[WorkSiteZip4]+''''
						+ ', @EmployerName='''			+  REPLACE(T200.[EmployerName], '''', '''''')+''''
						+ ', @EmployerAddr1='''			+  REPLACE(T200.[EmployerAddr1], '''', '''''')+''''
						+ ', @EmployerAddr2='''			+  REPLACE(T200.[EmployerAddr2], '''', '''''')+''''
						+ ', @EmployerCity='''			+  REPLACE(T200.[EmployerCity] , '''', '''''')+''''
						+ ', @EmployerState='''			+  T200.[EmployerState]+''''
						+ ', @EmployerZip5='''			+  T200.[EmployerZip5]+''''
						+ ', @EmployerZip4='''			+  T200.[EmployerZip4]+''''
						+ ', @SalaryRange='				+  (CASE WHEN T200.[SalaryRange]= '' THEN 'NULL' ELSE CAST(T200.[SalaryRange]  AS [varchar]) END)
						+ ', @SalaryType='''			+  REPLACE(T200.[SalaryType], '''', '''''')+''''
						+ ', @SalaryAmount='  			+  (CASE	WHEN T200.[SalaryAmount] = '' THEN 'NULL' ELSE T200.[SalaryAmount] END)
						+ ', @ElectedLeadershipTitle='''		+  REPLACE(T200.[ElectedLeadershipTitle], '''', '''''')+''''
						+ ', @ElectedLeadershipElectionDate='	+(CASE WHEN T200.[ElectedLeadershipElectionDate] = '' THEN 'NULL' 
																	  ELSE ''''+T200.[ElectedLeadershipElectionDate]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
																	  END)
						+ ', @Steward='''				+  T200.[Steward]+''''
						+ ', @Activist='''				+  T200.[Activist]+''''
						+ ', @EnterpriseID='''			+  T200.[EnterpriseID]+''''
						+ ', @PEOPLEContributionAmount='		+  (CASE	WHEN T200.[PEOPLEContributionAmount] = '' THEN 'NULL' ELSE T200.[PEOPLEContributionAmount] END)
						+ ', @PEOPLEContributionPayPeriod=' 	+  (CASE	WHEN T200.[PEOPLEContributionPayPeriod] = '' THEN 'NULL' 
																			ELSE ''''+T200.[PEOPLEContributionPayPeriod]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
																			END)	-- NEW
						+ ', @PEOPLEContributionFrequency='''	+  REPLACE(T200.[PEOPLEContributionFrequency], '''', '''''')+''''
						+ ', @PEOPLECheckIssuer='''		+  REPLACE(T200.[PEOPLECheckIssuer], '''', '''''')+''''
						+ ', @CustomName1='''			+  REPLACE(T200.[CustomName1] , '''', '''''')+''''
						+ ', @CustomValue1='''			+  REPLACE(T200.[CustomValue1], '''', '''''')+''''
						+ ', @CustomName2='''			+  REPLACE(T200.[CustomName2] , '''', '''''')+''''
						+ ', @CustomValue2='''			+  REPLACE(T200.[CustomValue2], '''', '''''')+''''
						+ ', @CustomName3='''			+  REPLACE(T200.[CustomName3] , '''', '''''')+''''
						+ ', @CustomValue3='''			+  REPLACE(T200.[CustomValue3], '''', '''''')+''''
						+ ', @creator_pk = 11949451, @created_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'+'
						INSERT INTO COM_Weekly_Mbr_Card_Run (person_pk, aff_pk, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
						+ '@last_person_pk, '	+ CAST(T200.desired_aff_pk AS varchar) 
						-- +', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''');'
						+', 11949451, CAST(getdate() AS date), 11949451, CAST(getdate() AS date));'			-- RW 8/26/2019
				END	
				FROM	[dbo].[AUP_Input_BAR] T200
						LEFT OUTER JOIN afscme_oltp6.dbo.Aff_Mbr_Rpt_Info MRI
						ON  MRI.aff_pk = T200.desired_aff_pk
				WHERE 	T200.is_valid_record = 1
				AND		T200.tran_type = 'A'
				AND		T200.[duplicate_ID] IS NOT NULL			-- Only for duplicate ADDs -- 20071210 -- GP
				AND		T200.[duplicate_ADDs] = 1				-- 20080208 GP
				AND 	T200.desired_aff_pk IS NOT NULL
				ORDER BY T200.[duplicate_ID], T200.[ID]			-- Make sure we keep the MIN([ID]) As the "duplicate_ID" of choice
				INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person/Member(S) Inserts Generated AS "duplicate ADDs" = '+ CAST(@@ROWCOUNT AS varchar)
			--
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
			--
	------------------------------------------------------------------------------------------
	------------------------------------------------------------------------------------------
	------------------------------------------------------------------------------------------
	--
	----------------------------------------------------------------
	-- Deletes -- Set Member Records as Inactive 
	----------------------------------------------------------------
	DECLARE @status_inactive int
	SELECT  @status_inactive = (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'MemberStatus' and Com_cd_desc = 'Inactive')
	DECLARE @status_active int
	SELECT  @status_active = (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'MemberStatus' and Com_cd_desc = 'Active')
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT ON'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--
	-- print 'Set Member Records as Inactive';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE  Aff_Members SET mbr_status = '+CAST(@status_inactive AS varchar)+', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE person_pk = '+ CAST(T200.person_pk As varchar) + ' AND aff_pk = '+ CAST(T200.existing_aff_pk AS varchar) + ' and mbr_status = '+CAST(@status_active AS varchar)
	FROM	[dbo].[AUP_Input_BAR] T200
	WHERE T200.[Tran_Type] = 'D' AND T200.[deactivate_AM] = 1
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Deletes Generated in Affiliate Tree = '+ CAST(@@ROWCOUNT AS varchar)
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--
	----------------------------------------
	-- Collect Stats for the process
	---------------------------------------------------------------------------
	-- Deletes Second -- Should duplicate Deletes: 
	-- UPDATE  Aff_Members SET mbr_status = '+CAST(@status_inactive AS varchar)
	--------------------------------------------------------------------------------
	DECLARE @ActivityType_Deactivate varchar(30)
	SELECT  @ActivityType_Deactivate = CAST((SELECT com_cd_pk FROM afscme_oltp6.dbo.common_codes WHERE com_cd_desc = 'Deactivate' AND com_cd_type_key = 'ActivityType') AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO [OPERATIONS].[dbo].[AUP_temp_stats_BAR] (aff_pk, time_pk, membership_activity_type, membership_activity_count) VALUES ('
		+CAST(T200.existing_aff_pk AS varchar)+', '
		+CAST(@time_pk_as_of_date AS varchar)+', '
		+@ActivityType_Deactivate+', '
		+CAST(count(*) AS varchar)+')'
	FROM	[dbo].[AUP_Input_BAR] T200
	WHERE T200.[Tran_Type] = 'D' AND T200.[deactivate_AM] = 1
	GROUP BY T200.existing_aff_pk
	ORDER BY T200.existing_aff_pk
	---------------------------------------------------
	--End For Generate Script for ADD/DELETE
	---------------------------------------------------
	
	
	---------------------------------------------------
	--Start of Generate Script for UPDATE
	---------------------------------------------------
	----------------------------------------------------
	-- Person Name Update
	----------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'Person Name Update';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person SET prefix_nm = ' + CAST(ISNULL(DMV1.com_cd_pk, 0) AS varchar) 
			+ ', first_nm = '''+ LTRIM(RTRIM(CAST(REPLACE(T200.First_Name, '''', '''''') AS varchar))) 
			+ ''', middle_nm = ''' + LTRIM(RTRIM(CAST(REPLACE(T200.Middle_Name, '''', '''''') AS varchar))) 	-- 20070326 GP
			+ ''', last_nm = ''' + LTRIM(RTRIM(CAST(REPLACE(T200.Last_Name, '''', '''''') AS varchar))) 
			+ ''', suffix_nm = ' + CAST(ISNULL(DMV2.com_cd_pk, 0) AS varchar) 
			+', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE person_pk = '
			+ CAST(T200.person_pk AS varchar)
			+' AND lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''	
	FROM 	[dbo].[AUP_Input_BAR] T200
		LEFT OUTER JOIN afscme_oltp6.dbo.DM_Code_Mapping_view DMV1 ON DMV1.com_cd_type_key = 'Prefix' AND DMV1.Legacy_Code = T200.Title
		LEFT OUTER JOIN afscme_oltp6.dbo.DM_Code_Mapping_view DMV2 ON DMV2.com_cd_type_key = 'Suffix' AND DMV2.Legacy_Code = T200.Suffix
	WHERE 	T200.update_P_Name = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_name_updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- 
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	-------------------------------------------------------
	-- Person SSN Update -- 20070108 -- GP -- update_P_SSN
	-------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'Person SSN Update';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 
			'UPDATE Person SET ssn = ''' + T200.SSN + ''''
			+ ', valid_ssn_fg = '+CASE
						WHEN T200.SSN IS NULL THEN '0'
						WHEN T200.SSN = '000000000' THEN '0'
						WHEN LEFT(T200.SSN, 4) = '0000' THEN '0'	-- GP 20070108
						WHEN T200.SSN = '' THEN '0'
						ELSE '1'
						END
			+ ', duplicate_ssn_fg = '+CAST(ISNULL(T200.duplicate_ssn_fg, 0) AS varchar)
			+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101)+ '''' 
			+ ' WHERE person_pk = '+ CAST(T200.person_pk AS varchar)
			+ ' AND lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''	
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.update_P_SSN = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_SSN_updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- 
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	----------------------------------------------------
	-- Person Address Update
	----------------------------------------------------
	DECLARE @membershipDept2 int,
		@PersonAddressType1 int
	 
	SET @membershipDept2 =   (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_cd   = 'AR'   AND com_cd_type_key = 'Department')   
	SET @PersonAddressType1 =(SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_desc = 'Home' AND com_cd_type_key = 'PersonAddressType')
	 
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'Person Address Update';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Address SET addr1 = '''
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr1, '') AS varbinary) AS varchar(50)))), '''', '''''')	-- 20170814 -- GP -- We want to keep the full [varchar](50) Addresses
	            +''', addr2 = ''' 
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr2, '') AS varbinary) AS varchar(50)))), '''', '''''')
	            +''', city = ''' 
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.City, '') AS varbinary) AS varchar(25)))), '''', '''''') 
		    +''', state = ''' 
	            + CASE
	                        WHEN  T200.State IS NULL THEN ''
	                        WHEN T200.State IN ('ZZ', 'XX') THEN ''
	                        WHEN    (LEN(LTRIM(RTRIM(T200.State)))=2) THEN T200.State
	                        ELSE    ''
	                        END +''', zipcode = '''
	            + CASE 
	                        WHEN T200.Zip IS NULL THEN ''
	                        WHEN T200.Zip = '00000' THEN ''
	                        ELSE LTRIM(RTRIM(T200.Zip ))
	                        END + ''', zip_plus = ''' 
	            + CASE
	                        WHEN T200.Zip_4 IS NULL THEN ''
	                        WHEN T200.Zip_4 = '0000' THEN ''
	                        ELSE LTRIM(RTRIM(T200.Zip_4))
	                        END + ''', dept =  '
	            + CAST(@membershipDept2 AS varchar) 
	            + ', addr_bad_fg = ' + CASE
	                                                WHEN T200.Addr_Mailable_fg IS NULL THEN '1'
	                                                WHEN T200.Addr_Mailable_fg = 'Y'   THEN '0'
	                                                WHEN T200.Addr_Mailable_fg = 'N'   THEN '1'
	                                                ELSE '1'
	                                                END
	            + ', addr_marked_bad_dt = '+ CASE
	                                                WHEN T200.Addr_Mailable_fg IS NULL THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                WHEN T200.Addr_Mailable_fg = 'Y'   THEN 'NULL'
	                                                WHEN T200.Addr_Mailable_fg = 'N'   THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                ELSE ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                END
	            + ', addr_source = ''U'', eff_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
	            + ', end_dt = NULL'		-- 20070925 GP  -- Set Person_Address.end_dt = NULL when we update an existing address
	            + ', addr_source_if_aff_apply_upd = '''+CAST(@reporting_aff_pk AS varchar)+''''
	            + ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
	            + ' WHERE address_pk = ' + CAST(T200.old_address_pk as varchar) 
	            +' AND lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''
	-- select  T200.*
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.update_PA = 1
	AND	T200.is_valid_record = 1
	AND 	T200.old_address_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_address_updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- 
	-------------------------------------------------------------
	-- 20070827 GP  -- Add code for [insert_PA] and [update_SMA]
	-------------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print '[update_officer_addr]';
	-- 20080509 GP  -- Add code for [update_officer_addr] = 1
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[AUP_Insert_Person_Address]'
		+ ' @person_pk = ' + CAST(T200.person_pk AS [varchar])
		-- Person_Address
		+ ', @addr_type = 12001' 							-- Home
		+ ', @dept = ' + CAST(@membershipDept2 AS varchar) 
		+ ', @addr_source = ''U'''
		+ ', @addr_source_if_aff_apply_upd = '''+CAST(@reporting_aff_pk AS varchar)+''''
		+ ', @addr_prmry_fg = ' + CAST(ISNULL(T200.[update_SMA], 0) AS [varchar])  -- Primary if SMA
		+ ', @addr_bad_fg = ' + CASE
					WHEN T200.Addr_Mailable_fg IS NULL THEN '1'
					WHEN T200.Addr_Mailable_fg = 'Y'   THEN '0'
					WHEN T200.Addr_Mailable_fg = 'N'   THEN '1'
					ELSE '1'
					END
		+ ', @addr_marked_bad_dt = '+ CASE
					WHEN T200.Addr_Mailable_fg IS NULL THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
					WHEN T200.Addr_Mailable_fg = 'Y'   THEN 'NULL'
					WHEN T200.Addr_Mailable_fg = 'N'   THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
					ELSE ''''+CONVERT(varchar, @as_of_date, 101) +''''
					END
		+ ', @addr_private_fg = 0'
		+ ', @addr1 = '''+REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr1, '') AS varbinary) AS varchar(50)))), '''', '''''') + ''''	-- 20170814 -- GP -- We want to keep the full [varchar](50) Addresses
		+ ', @addr2 = '''+REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr2, '') AS varbinary) AS varchar(50)))), '''', '''''') + ''''
		+ ', @city = ''' +REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.City, '') AS varbinary) AS varchar(25)))), '''', '''''')  + ''''
		+ ', @state = '''+ CASE
					WHEN  T200.State IS NULL THEN ''
					WHEN T200.State IN ('ZZ', 'XX') THEN ''
					WHEN    (LEN(LTRIM(RTRIM(T200.State)))=2) THEN T200.State
					ELSE    ''
					END + ''''
		+ ', @zipcode = '''+ CASE 
					WHEN T200.Zip IS NULL THEN ''
					WHEN T200.Zip = '00000' THEN ''
					ELSE LTRIM(RTRIM(T200.Zip ))
					END + '''' 
		+ ', @zip_plus = '''+ CASE
					WHEN T200.Zip_4 IS NULL THEN ''
					WHEN T200.Zip_4 = '0000' THEN ''
					ELSE LTRIM(RTRIM(T200.Zip_4))
					END + ''''
		+ ', @province = NULL'
		+ ', @carrier_route_info = NULL'
		+ ', @country  = NULL' 								-- 9001 = USA
		+ ', @county  = NULL'
		+ ', @eff_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ', @end_dt  = NULL'
		+ ', @created_user_pk = 11949451'
		+ ', @created_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ', @lst_mod_user_pk = 11949451'
		+ ', @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		-- SMA/Officer
		+ ', @set_SMA = ' + CAST(ISNULL(T200.[update_SMA], 0) AS [varchar])
		+ ', @update_officer_addr = ' + CAST(ISNULL(T200.[update_officer_addr], 0) AS [varchar])
	-- select  T200.*
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.[insert_PA] = 1
	AND		T200.is_valid_record = 1
	AND		T200.[person_pk] IS NOT NULL								-- 20140624 GP	-- Make sure we have valid person_pkS when inserting addresses
	ORDER BY T200.[update_SMA], T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Existing Person_address_inserts Generated (SMA or not) = '+ CAST(@@ROWCOUNT AS varchar)
	-- 	
	-------------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE person_address SET zip_plus = NULL WHERE LEN(RTRIM(zip_plus)) = 0'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	----------------------------------------------------
	-- Person_Demographics -- gender / DOB update -- 20170116 -- GP -- ADD DateOfBirth / dob
	----------------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'UPDATE Person_Demographics';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Demographics SET '
		+(CASE WHEN ISNULL(T200.[update_Gender], 0) = 1 THEN ' gender = '+ ISNULL(CAST(CMV.com_cd_pk AS varchar), 'NULL') ELSE '' END)
		+(CASE WHEN ISNULL(T200.[update_Gender], 0) = 1 AND ISNULL(T200.[update_DOB], 0) = 1 THEN ',' ELSE '' END)
		+(CASE WHEN ISNULL(T200.[update_DOB], 0) = 1 THEN ' [dob] = '''+ T200.[DateOfBirth] +'''' ELSE '' END)
		+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ' WHERE lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''' AND person_pk = '	+ CAST(T200.person_pk AS varchar)
	FROM 	[dbo].[AUP_Input_BAR] T200
		LEFT OUTER JOIN afscme_oltp6.dbo.DM_Code_Mapping_view CMV
			ON  CMV.com_cd_type_key = 'Gender' 
			AND CMV.Legacy_Code = T200.Sex_cd
	WHERE 	T200.update_PDG = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_gender Updates+DOB Generated = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- (2289/2312 row(s) affected) -- Missing 23 records  

	----------------------------------------------------
	--- Person_Political_Legislative -- Update
	----------------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'Update PPL';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Political_Legislative SET '
		+ ' political_do_not_call_fg = '+ CASE	
							WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'		-- 20170328 -- GP
							ELSE '0'
							END
		+ ', political_party = ' + CAST(CASE
			WHEN T200.Political_Party IS NULL THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
			WHEN T200.Political_Party = 'U' THEN (SELECT com_cd_pk FROM afscme_oltp6.dbo.Common_Codes WHERE com_cd_type_key = 'PoliticalParty' AND com_cd_cd = 'U')
			ELSE (SELECT com_cd_pk FROM afscme_oltp6.dbo.DM_Code_Mapping_view WHERE com_cd_type_key = 'PoliticalParty' AND Legacy_Code = T200.Political_Party)
			END AS varchar)
		+ ', political_registered_voter = ' + CASE
			WHEN CC.com_cd_pk IS NULL THEN 'NULL'
			ELSE CAST(CC.com_cd_pk AS varchar)
			END
		+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ' WHERE (lst_mod_dt IS NULL OR lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''') AND person_pk = '+ CAST(T200.person_pk AS varchar)
	FROM 	[dbo].[AUP_Input_BAR] T200
		LEFT OUTER JOIN afscme_oltp6.dbo.common_codes CC
			ON  CC.com_cd_type_key = 'RegisteredVoter'
			AND CC.com_cd_cd = T200.Registered_Voter_fg
	WHERE 	T200.update_PPL = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_Political_Legislative Updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	------------------------------------------
	-- Phones -- Insert first
	------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 20160524 -- GP -- ADD MULTIPLE PHONES
	-- print '[sp_Insert_Person_Phone_BAR]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Insert_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_num = '''+T200.[phone]+''', @phone_type = 3001, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[insert_PPh] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Home Phone Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-- print '[sp_Insert_Person_Phone_BAR_work]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Insert_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_num = '''+T200.[WorkPhone]+''', @phone_type = 3002, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[WorkPhone] > ''
	AND		T200.[edit_WPh] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Work Phone Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 	
	-- print '[sp_Insert_Person_Phone_BAR_cell]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Insert_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_num = '''+T200.[CellPhone]+''', @phone_type = 3003, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[CellPhone] > ''
	AND		T200.[edit_CPh] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Cell Phone Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-- print '[sp_Insert_Person_Phone_BAR_Fax]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Insert_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_num = '''+T200.[FaxNumber]+''', @phone_type = 3004, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[FaxNumber] > ''
	AND		T200.[edit_Fax] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Fax Number Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	------------------------------------------
	-- Phones -- Updates Second
	------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--
	-- print '[sp_Update_Person_Phone_BAR]';
	-- 20160524 -- GP -- ADD MULTIPLE PHONES
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Update_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_pk = '+CAST(T200.[old_phone_pk] AS varchar)+', @phone_num = '''+T200.[phone]+''', @phone_type = 3001, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[update_PPh] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Home Phone Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-- print '[sp_Update_Person_Phone_BAR_work]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Update_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_pk = '+CAST(T200.[old_WPh_pk] AS varchar)+', @phone_num = '''+T200.[WorkPhone]+''', @phone_type = 3002, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[WorkPhone] > ''
	AND		T200.[edit_WPh] = 2
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Work Phone Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-- print '[sp_Update_Person_Phone_BAR_cell]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Update_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_pk = '+CAST(T200.[old_CPh_pk] AS varchar)+', @phone_num = '''+T200.[CellPhone]+''', @phone_type = 3003, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[CellPhone] > ''
	AND		T200.[edit_CPh] = 2
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Cell Phone Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-- print '[sp_Update_Person_Phone_BAR_Fax]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'EXEC [dbo].[sp_Update_Person_Phone_BAR] @person_pk = '+CAST(T200.person_pk AS varchar)+', @phone_pk = '+CAST(T200.[old_Fax_pk] AS varchar)+', @phone_num = '''+T200.[FaxNumber]+''', @phone_type = 3004, @phone_prmry_fg = 0, @dept = 4001, @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[FaxNumber] > ''
	AND		T200.[edit_Fax] = 2
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Personal Fax Number Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 

/*	
	------------------------------------------
	-- Person_Phone -- Insert first
	------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT DISTINCT @load_id, 'INSERT INTO Person_Phone (person_pk, country_cd, area_code, phone_no, phone_prmry_fg, phone_bad_fg, phone_marked_bad_dt, phone_type, dept, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
		+ CAST(T200.person_pk as varchar)
		+ ', ''1'''
		+ ', ''' + LEFT(LTRIM(T200.Phone), 3)
		+ ''', '''+ SUBSTRING(LTRIM(T200.Phone), 4, 7)
		+ ''', ' + (CASE WHEN (PPH.[phone_prmry_fg] IS NULL) THEN '1' ELSE '0' END)		-- 20080110 GP
		-- 20080501 GP  -- Add phone_bad_fg/phone_marked_bad_dt
		+ ', '+ (CASE WHEN REPLACE(REPLACE(SUBSTRING(LTRIM(T200.Phone), 4, 7), '0', ''), '9', '') > '' THEN '0' ELSE '1' END)
		+ ', '+ (CASE WHEN REPLACE(REPLACE(SUBSTRING(LTRIM(T200.Phone), 4, 7), '0', ''), '9', '') > '' THEN 'NULL' ELSE ''''+CONVERT(varchar, @as_of_date, 101) +'''' END)
		+ ', ' + CAST(@PhoneTypeHome AS varchar)
		+ ', ' + CAST(@membershipDept As varchar)	-- Was @membershipDept1
		+ ', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''')'
	FROM 	[dbo].[AUP_Input_BAR] T200
			LEFT OUTER JOIN [afscme_oltp6].[dbo].[Person_Phone] PPH
				ON  PPH.person_pk = T200.person_pk
				AND ISNULL(PPH.[phone_bad_fg], 0) = 0
				AND PPH.[phone_prmry_fg] = 1
	WHERE 	T200.[insert_PPh] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY 2
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_Phone Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'


	------------------------------------------
	-- Person_Phone -- Update second
	------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Phone SET '
		+ ' country_cd = ''1'''
		+ ', area_code = ''' + LEFT(LTRIM(T200.Phone), 3)+''''
		+ ', phone_no = '''+ SUBSTRING(LTRIM(T200.Phone), 4, 7)+''''
		--	+ ', phone_prmry_fg = 1'													-- 20080110 GP
		-- 20080501 GP  -- Add phone_bad_fg/phone_marked_bad_dt
		+ ', phone_bad_fg = '       + (CASE WHEN REPLACE(REPLACE(SUBSTRING(LTRIM(T200.Phone), 4, 7), '0', ''), '9', '') > '' THEN '0' ELSE '1' END)
		+ ', phone_marked_bad_dt = '+ (CASE WHEN REPLACE(REPLACE(SUBSTRING(LTRIM(T200.Phone), 4, 7), '0', ''), '9', '') > '' THEN 'NULL' ELSE ''''+CONVERT(varchar, @as_of_date, 101) +'''' END)
		+ ', phone_type = ' + CAST(@PhoneTypeHome AS varchar)
		+ ', dept = ' + CAST(@membershipDept As varchar)	-- Was @membershipDept1
		+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ' WHERE phone_pk = '+ CAST(T200.[old_phone_pk] AS varchar)
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.[update_PPh] = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_Phone Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
*/
	-- print '[sp_Insert_Misc_Data_BAR]';
	-- 20160526 -- GP -- ADD MISC DATA
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 
	'EXEC [dbo].[sp_Insert_Misc_Data_BAR] '
	+   '@person_pk='+CAST(T200.person_pk AS varchar)
	+ ', @aff_pk='+CAST(T200.desired_aff_pk AS varchar)
	+ ', @JobTitle='''				+  REPLACE(T200.[JobTitle], '''', '''''')+''''
	+ ', @JobSector='  				+  (CASE WHEN T200.[JobSector] = '' THEN 'NULL' ELSE CAST(T200.[JobSector] AS [varchar])  END)
	+ ', @EmplSector='  			+  (CASE WHEN T200.[EmplSector]= '' THEN 'NULL' ELSE CAST(T200.[EmplSector] AS [varchar]) END)
	+ ', @JobHireDate='				+  (CASE WHEN T200.[JobHireDate] = '' THEN 'NULL' 
											 ELSE ''''+T200.[JobHireDate]+''''	-- 20170113 GP+VA -- PROPER CONVERSION FOR Dates
											 END)
	+ ', @WorkSiteName='''			+  REPLACE(T200.[WorkSiteName] , '''', '''''')+''''
	+ ', @WorkSiteAddr1='''			+  REPLACE(T200.[WorkSiteAddr1], '''', '''''')+''''
	+ ', @WorkSiteAddr2='''			+  REPLACE(T200.[WorkSiteAddr2], '''', '''''')+''''
	+ ', @WorkSiteCity='''			+  REPLACE(T200.[WorkSiteCity] , '''', '''''')+''''
	+ ', @WorkSiteState='''			+  T200.[WorkSiteState]+''''
	+ ', @WorkSiteZip5='''			+  T200.[WorkSiteZip5]+''''
	+ ', @WorkSiteZip4='''			+  T200.[WorkSiteZip4]+''''
	+ ', @EmployerName='''			+  REPLACE(T200.[EmployerName], '''', '''''')+''''
--	+ ', @EmployerCode='  			+  (CASE WHEN T200.[EmployerCode]= '' THEN 'NULL' ELSE CAST(T200.[EmployerCode] AS [varchar]) END)
	+ ', @EmployerAddr1='''			+  REPLACE(T200.[EmployerAddr1], '''', '''''')+''''
	+ ', @EmployerAddr2='''			+  REPLACE(T200.[EmployerAddr2], '''', '''''')+''''
	+ ', @EmployerCity='''			+  REPLACE(T200.[EmployerCity] , '''', '''''')+''''
	+ ', @EmployerState='''			+  T200.[EmployerState]+''''
	+ ', @EmployerZip5='''			+  T200.[EmployerZip5]+''''
	+ ', @EmployerZip4='''			+  T200.[EmployerZip4]+''''
	+ ', @SalaryRange='				+  (CASE WHEN T200.[SalaryRange]= '' THEN 'NULL' ELSE CAST(T200.[SalaryRange]  AS [varchar]) END)
	+ ', @SalaryType='''			+  REPLACE(T200.[SalaryType], '''', '''''')+''''
	+ ', @SalaryAmount='  			+  (CASE	WHEN T200.[SalaryAmount] = '' THEN 'NULL' ELSE T200.[SalaryAmount] END)
	+ ', @ElectedLeadershipTitle='''		+  REPLACE(T200.[ElectedLeadershipTitle], '''', '''''')+''''
	+ ', @ElectedLeadershipElectionDate='	+(CASE WHEN T200.[ElectedLeadershipElectionDate] = '' THEN 'NULL' 
												  ELSE ''''+T200.[ElectedLeadershipElectionDate]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
												  END)
	+ ', @Steward='''				+  T200.[Steward]+''''
	+ ', @Activist='''				+  T200.[Activist]+''''
	+ ', @EnterpriseID='''			+  T200.[EnterpriseID]+''''
	+ ', @PEOPLEContributionAmount='		+  (CASE	WHEN T200.[PEOPLEContributionAmount] = '' THEN 'NULL' ELSE T200.[PEOPLEContributionAmount] END)
	+ ', @PEOPLEContributionPayPeriod=' 	+  (CASE	WHEN T200.[PEOPLEContributionPayPeriod] = '' THEN 'NULL' 
														ELSE ''''+T200.[PEOPLEContributionPayPeriod]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
														END)	-- NEW
	+ ', @PEOPLEContributionFrequency='''	+  REPLACE(T200.[PEOPLEContributionFrequency], '''', '''''')+''''
	+ ', @PEOPLECheckIssuer='''		+  REPLACE(T200.[PEOPLECheckIssuer], '''', '''''')+''''
	+ ', @CustomName1='''			+  REPLACE(T200.[CustomName1] , '''', '''''')+''''
	+ ', @CustomValue1='''			+  REPLACE(T200.[CustomValue1], '''', '''''')+''''
	+ ', @CustomName2='''			+  REPLACE(T200.[CustomName2] , '''', '''''')+''''
	+ ', @CustomValue2='''			+  REPLACE(T200.[CustomValue2], '''', '''''')+''''
	+ ', @CustomName3='''			+  REPLACE(T200.[CustomName3] , '''', '''''')+''''
	+ ', @CustomValue3='''			+  REPLACE(T200.[CustomValue3], '''', '''''')+''''
	+ ', @creator_pk = 11949451, @created_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	-- select *
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[edit_MiscData] = 1 -- INSERT
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk		  IS NOT NULL
	AND		T200.[desired_aff_pk] IS NOT NULL	
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Member Misc Data Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'	
	-- 
	-- print '[sp_Update_Misc_Data_BAR]';
	-- 20170118 -- GP -- ADD MISC DATA
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 
	'EXEC [dbo].[sp_Update_Misc_Data_BAR] '
	+   '@person_pk='+CAST(T200.person_pk AS varchar)
	+ ', @aff_pk='+CAST(T200.desired_aff_pk AS varchar)
	+ ', @JobTitle='''				+  REPLACE(T200.[JobTitle], '''', '''''')+''''
	+ ', @JobSector='  				+  (CASE WHEN T200.[JobSector] = '' THEN 'NULL' ELSE CAST(T200.[JobSector] AS [varchar])  END)
	+ ', @EmplSector='  			+  (CASE WHEN T200.[EmplSector]= '' THEN 'NULL' ELSE CAST(T200.[EmplSector] AS [varchar]) END)
	+ ', @JobHireDate='				+  (CASE WHEN T200.[JobHireDate] = '' THEN 'NULL' 
											 ELSE ''''+T200.[JobHireDate]+''''	-- 20170113 GP+VA -- PROPER CONVERSION FOR Dates
											 END)
	+ ', @WorkSiteName='''			+  REPLACE(T200.[WorkSiteName] , '''', '''''')+''''
	+ ', @WorkSiteAddr1='''			+  REPLACE(T200.[WorkSiteAddr1], '''', '''''')+''''
	+ ', @WorkSiteAddr2='''			+  REPLACE(T200.[WorkSiteAddr2], '''', '''''')+''''
	+ ', @WorkSiteCity='''			+  REPLACE(T200.[WorkSiteCity] , '''', '''''')+''''
	+ ', @WorkSiteState='''			+  T200.[WorkSiteState]+''''
	+ ', @WorkSiteZip5='''			+  T200.[WorkSiteZip5]+''''
	+ ', @WorkSiteZip4='''			+  T200.[WorkSiteZip4]+''''
	+ ', @EmployerName='''			+  REPLACE(T200.[EmployerName], '''', '''''')+''''
--	+ ', @EmployerCode='  			+  (CASE WHEN T200.[EmployerCode]= '' THEN 'NULL' ELSE CAST(T200.[EmployerCode] AS [varchar]) END)
	+ ', @EmployerAddr1='''			+  REPLACE(T200.[EmployerAddr1], '''', '''''')+''''
	+ ', @EmployerAddr2='''			+  REPLACE(T200.[EmployerAddr2], '''', '''''')+''''
	+ ', @EmployerCity='''			+  REPLACE(T200.[EmployerCity] , '''', '''''')+''''
	+ ', @EmployerState='''			+  T200.[EmployerState]+''''
	+ ', @EmployerZip5='''			+  T200.[EmployerZip5]+''''
	+ ', @EmployerZip4='''			+  T200.[EmployerZip4]+''''
	+ ', @SalaryRange='				+  (CASE WHEN T200.[SalaryRange]= '' THEN 'NULL' ELSE CAST(T200.[SalaryRange]  AS [varchar]) END)
	+ ', @SalaryType='''			+  REPLACE(T200.[SalaryType], '''', '''''')+''''
	+ ', @SalaryAmount='  			+  (CASE	WHEN T200.[SalaryAmount] = '' THEN 'NULL' ELSE T200.[SalaryAmount] END)
	+ ', @ElectedLeadershipTitle='''		+  REPLACE(T200.[ElectedLeadershipTitle], '''', '''''')+''''
	+ ', @ElectedLeadershipElectionDate='	+(CASE WHEN T200.[ElectedLeadershipElectionDate] = '' THEN 'NULL' 
												  ELSE ''''+T200.[ElectedLeadershipElectionDate]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
												  END)
	+ ', @Steward='''				+  T200.[Steward]+''''
	+ ', @Activist='''				+  T200.[Activist]+''''
	+ ', @EnterpriseID='''			+  T200.[EnterpriseID]+''''
	+ ', @PEOPLEContributionAmount='		+  (CASE	WHEN T200.[PEOPLEContributionAmount] = '' THEN 'NULL' ELSE T200.[PEOPLEContributionAmount] END)
	+ ', @PEOPLEContributionPayPeriod=' 	+  (CASE	WHEN T200.[PEOPLEContributionPayPeriod] = '' THEN 'NULL' 
														ELSE ''''+T200.[PEOPLEContributionPayPeriod]+'''' -- 20170113 GP+VA -- PROPER CONVERSION FOR FOR Dates
														END)	-- NEW
	+ ', @PEOPLEContributionFrequency='''	+  REPLACE(T200.[PEOPLEContributionFrequency], '''', '''''')+''''
	+ ', @PEOPLECheckIssuer='''		+  REPLACE(T200.[PEOPLECheckIssuer], '''', '''''')+''''
	+ ', @CustomName1='''			+  REPLACE(T200.[CustomName1] , '''', '''''')+''''
	+ ', @CustomValue1='''			+  REPLACE(T200.[CustomValue1], '''', '''''')+''''
	+ ', @CustomName2='''			+  REPLACE(T200.[CustomName2] , '''', '''''')+''''
	+ ', @CustomValue2='''			+  REPLACE(T200.[CustomValue2], '''', '''''')+''''
	+ ', @CustomName3='''			+  REPLACE(T200.[CustomName3] , '''', '''''')+''''
	+ ', @CustomValue3='''			+  REPLACE(T200.[CustomValue3], '''', '''''')+''''
	+ ', @lst_mod_user_pk = 11949451, @lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''';'
	-- select *
	FROM 	[dbo].[AUP_Input_BAR] T200	
	WHERE 	T200.[edit_MiscData] = 2 -- Update
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk		  IS NOT NULL
	AND		T200.[desired_aff_pk] IS NOT NULL	
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Member Misc Data Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'	
	-- 
	-- 20110503 GP  -- Add Update Person_Email (Home and Work)
	------------------------------------------
	-- Person_Email -- Update Home First
	------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print '[Person_Email -- Update Home First]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Email SET '
		+ ' [person_email_addr] = ''' + REPLACE(LTRIM(RTRIM(T200.[Home_Email])), '''', '''''')+ ''''
		+ ', email_bad_fg = 0'
		+ ', email_marked_bad_dt = NULL'
		+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ' WHERE [person_pk] = '+ CAST(T200.[person_pk] AS varchar)
		+ ' AND [email_type] = 71001'
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.[update_HMail] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.[person_pk] IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Home Person_email Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--
	------------------------------------------
	-- Person_Email -- Update Work Second
	------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--  print '[Person_Email -- Update Work Second]';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Person_Email SET '
		+ ' [person_email_addr] = ''' + REPLACE(LTRIM(RTRIM(T200.[Work_Email])), '''', '''''')+ ''''
		+ ', email_bad_fg = 0'
		+ ', email_marked_bad_dt = NULL'
		+ ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
		+ ' WHERE [person_pk] = '+ CAST(T200.[person_pk] AS varchar)
		+ ' AND [email_type] = 71002'
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.[update_WMail] = 1
	AND		T200.is_valid_record = 1
	AND 	T200.[person_pk] IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Work Person_email Updates Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	-----------------------------------------------------------------------------
	-- Change Member Type -- This may become UPDATE MEMBER (many or ALL fields)
	-----------------------------------------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'UPDATE Aff_Members';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Aff_Members SET '
		+ ' mbr_type = ' + CAST(CASE
			WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A', 'Y') THEN @Regular	-- 20170328 -- GP
			WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A'     ) THEN @Retiree	-- 20170328 -- GP
			WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('T', 'R') THEN @Regular
			WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Retiree
			WHEN T200.Status_Code = 'N' THEN @AgencyFeePayer
			WHEN T200.Status_Code = 'C' THEN @UnionShop
			WHEN T200.Status_Code = 'P' THEN @PotentialMember					-- 20130323 -- Add 'Potential Members'
			WHEN T200.Status_Code = 'O' THEN @OptOutMember			-- 20170328 -- GP	-- 20130323 -- Add 'Potential Members'
	 		WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN @RetireeSpouse
	 		-- WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code = 'X' THEN @AssociateMember
	 		ELSE 0
			END AS varchar) 
		+ ', mbr_status = ' + CAST(CASE
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('A', 'O', 'P') THEN @Active
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Temporary 	
			WHEN T200.Status_Code = 'N' THEN @Active  
			WHEN T200.Status_Code = 'P' THEN @Active						-- 20130323 -- Add 'Potential Members'
			WHEN T200.Affiliate_Identifier = 'C' THEN @Active
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code = 'X' THEN @Active
			WHEN T200.Status_Code = 'Y' THEN @Pending			-- 20170328 -- GP
			ELSE 0 
			END AS varchar)
		+ ', no_mail_fg = ' + CASE T200.No_Mail_fg 
					WHEN '9' THEN '1' 
					ELSE '0'
					END
		+ ', no_legislative_mail_fg = ' + CASE 
			WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') THEN '1'		-- 20170328 -- GP
			ELSE '0'
			END
		+ ', no_cards_fg = ' + CASE
			WHEN ISNULL(MRI.unit_wide_no_mbr_cards_fg, 0) = 1 THEN '1'
			WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') 	  THEN '1'	-- 20170328 -- GP										-- 20130323 -- Add 'Potential Members'
			WHEN T200.No_Mail_fg  = '1'                       THEN '1'
			WHEN T200.No_Mail_fg  = '3'                       THEN '1'
			ELSE '0'
			END
		+ ', no_public_emp_fg = ' + CASE
			WHEN ISNULL(MRI.unit_wide_no_pe_mail_fg, 0) = 1   THEN '1'
	 		WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN '1'
	 		WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y') 	  THEN '1'	-- 20170328 -- GP
			WHEN T200.No_Mail_fg  = '2'                       THEN '1'
			WHEN T200.No_Mail_fg  = '3'                       THEN '1'
	 		ELSE '0'
	 		END
		+ CASE
			WHEN T200.Date_Joined = '' THEN (CASE 	-- 20070507 GP
								WHEN AM.mbr_join_dt IS NOT NULL THEN ''
								ELSE ', mbr_join_dt = '''+ CONVERT(varchar, @as_of_date, 101) + ''''
								END)
			ELSE ', mbr_join_dt = '''+T200.Date_Joined+''''	-- 20170113 GP+VA
			END
		+ CASE
			WHEN ISNULL(@data_source, '') <> 'UWare' THEN
				(CASE 
					WHEN (T200.Affiliate_Member_ID IS NULL)				THEN ''
					WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = ''		THEN ''
					ELSE ', mbr_no_local = '''+LTRIM(RTRIM(T200.Affiliate_Member_ID))+''''
					END)
			WHEN ISNULL(@data_source, '') = 'UWare' AND ISNULL(@first_run, -1) = 0 THEN
				(CASE 
					WHEN (T200.Affiliate_Member_ID IS NULL)				THEN ''
					WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = ''		THEN ''
					ELSE ', mbr_no_local = '''+LTRIM(RTRIM(T200.Affiliate_Member_ID))+''''
					END)
			WHEN @data_source = 'UWare'  AND ISNULL(@first_run, -1) <> 0 THEN
				(CASE 
					WHEN (T200.Affiliate_Member_ID IS NULL)				THEN ''
					WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = ''		THEN ''
					WHEN LTRIM(RTRIM(AM.mbr_no_local)) = LTRIM(RTRIM(T200.Affiliate_Member_ID)) THEN ''
					WHEN ISNULL(AM.mbr_no_local, '') = '' THEN ', mbr_no_local = '''+LTRIM(RTRIM(T200.Affiliate_Member_ID))+''''
					ELSE ''  --VA/GP -- 20190320
					END)
			END
		+ ', primary_information_source = '+ CASE
			WHEN (T200.[Information_Source] IS NULL) THEN 'NULL'
			WHEN  T200.[Information_Source] = '' THEN 'NULL'
			WHEN  T200.[Information_Source] = '4' THEN '47004'
			ELSE  '47008'
			END 
		+', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE person_pk = '
		+ CAST(T200.person_pk As varchar) + ' AND aff_pk = '
		+ CAST(T200.existing_aff_pk AS varchar) 
		+ ' and lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''
	FROM 	[dbo].[AUP_Input_BAR] T200
		INNER JOIN afscme_oltp6.dbo.Aff_Members AM		-- 20070507 GP
			ON  AM.person_pk = T200.person_pk
			AND AM.aff_pk = T200.existing_aff_pk
		LEFT OUTER JOIN afscme_oltp6.dbo.Aff_Mbr_Rpt_Info MRI
			ON  MRI.aff_pk = T200.desired_aff_pk
	WHERE 	T200.update_AM = 1
	AND		T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL 
	AND		T200.existing_aff_pk IS NOT NULL
	ORDER BY T200.existing_aff_pk, T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Member_type_updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- 
	
	----------------------------------------
	-- Collect Stats for the process
	---------------------------------------------------------------------------
	-- Updates Third -- Should duplicate "Change mbr type and status":
	---------------------------------------------------------------------------
	-- print 'UPDATE Aff_Members, Change mbr type and status';
	DECLARE @ActivityType_Update varchar(30)
	SELECT  @ActivityType_Update = CAST((SELECT com_cd_pk FROM afscme_oltp6.dbo.common_codes WHERE com_cd_desc = 'Update' AND com_cd_type_key = 'ActivityType') AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO [OPERATIONS].[dbo].[AUP_temp_stats_BAR] (aff_pk, time_pk, membership_activity_type, membership_activity_count) VALUES ('
		+CAST(T200.existing_aff_pk AS varchar)+', '
		+CAST(@time_pk_as_of_date AS varchar)+', '
		+@ActivityType_Update+', '
		+CAST(count(*) AS varchar)+')'
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.update_AM = 1
	AND	T200.is_valid_record = 1
	AND 	T200.person_pk IS NOT NULL 
	AND	T200.existing_aff_pk IS NOT NULL
	GROUP BY T200.existing_aff_pk
	ORDER BY T200.existing_aff_pk
	-- 
	-- print 'Collect Stats for the process';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-------------------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- We need to make sure the PK of aff_mbr_activity is respected'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-------------------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- TRUNCATE TABLE [OPERATIONS].[dbo].[AUP_temp_stats_unique_BAR]'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO [OPERATIONS].[dbo].[AUP_temp_stats_unique_BAR]([aff_pk], [time_pk], [membership_activity_type], [membership_activity_count]) SELECT aff_pk, time_pk, membership_activity_type, SUM(membership_activity_count) FROM [OPERATIONS].[dbo].[AUP_temp_stats_BAR] GROUP BY aff_pk, time_pk, membership_activity_type'
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Existing Records in aff_mbr_activity should be incremented'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE AMA SET AMA.membership_activity_count = (AMA.membership_activity_count + TS.membership_activity_count) FROM [OPERATIONS].[dbo].[AUP_temp_stats_unique_BAR] TS INNER JOIN  dbo.aff_mbr_activity AMA ON AMA.aff_pk = TS.aff_pk AND AMA.time_pk = TS.time_pk AND AMA.membership_activity_type = TS.membership_activity_type'
	
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- New Records in aff_mbr_activity should be inserted'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------------------------------------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO dbo.aff_mbr_activity (aff_pk, time_pk, membership_activity_type, membership_activity_count) SELECT TS.aff_pk, TS.time_pk, TS.membership_activity_type, TS.membership_activity_count FROM [OPERATIONS].[dbo].[AUP_temp_stats_unique_BAR] TS LEFT OUTER JOIN dbo.aff_mbr_activity AMA ON AMA.aff_pk = TS.aff_pk AND AMA.time_pk = TS.time_pk AND AMA.membership_activity_type = TS.membership_activity_type WHERE AMA.aff_pk IS NULL'
	
	----------------------------------------------------
	-- COM_Weekly_Mbr_Card_Run
	-- name changes member need cards
	----------------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'name changes member need cards';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO COM_Weekly_Mbr_Card_Run (person_pk, aff_pk, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
	            + CAST(T200.person_pk AS varchar) +', '
	            -- + CAST(T200.desired_aff_pk AS varchar) +', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''')'
	            + CAST(T200.desired_aff_pk AS varchar) +', 11949451, CAST(getdate() AS date), 11949451, CAST(getdate() AS date))'			-- RW 8/26/2019
	FROM   [dbo].[AUP_Input_BAR] T200
	WHERE 	T200.request_Card = 1
	AND     T200.is_valid_record = 1
	AND     ISNULL(T200.tran_type, '') NOT IN ('A', 'I')
	AND     T200.desired_aff_pk = T200.existing_aff_pk	-- 20070516 GP -- Requested for other reasons than Not Having Membership in desired Affiliate
	AND     T200.person_pk IS NOT NULL
	AND     T200.desired_aff_pk IS NOT NULL
	ORDER BY T200.person_pk
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requested for Change of Name, Status or Type Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- 
	-- 20070524 -- Report Card Requests by category
	-- Request cards for change of Name
	-- print 'cards for change of name';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requested for Change of Name   = '
		+ CAST(	(SELECT count(*) 
			FROM 	[dbo].[AUP_Input_BAR] T200
			WHERE 	T200.[request_Card] = 1
			AND 	T200.[is_valid_record] = 1
			AND 	ISNULL(T200.tran_type, '') NOT IN ('A', 'I', 'D')
			AND 	T200.person_pk IS NOT NULL
			AND 	T200.desired_aff_pk IS NOT NULL
			AND 	T200.desired_aff_pk = T200.existing_aff_pk
			AND 	ISNULL(T200.[update_P_Name], 0) = 1
			) AS varchar)
	--
	-- Requested cards for Agency Fee Payers becoming Regular or Retiree -- 20080227 GP
	-- print 'AEFs -. Regular/Rests';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requested for Change of Type = '
		+ CAST(	(SELECT count(*) 
			FROM	[dbo].[AUP_Input_BAR] T200
				INNER JOIN afscme_oltp6.dbo.aff_members AM
					ON  AM.person_pk     = T200.person_pk
					AND AM.aff_pk        = T200.existing_aff_pk
					AND AM.mbr_type      = 29003	-- Is Agency Fee Payer on Enterprise
					AND T200.Status_Code = 'A' 	-- And is Active (Regular/Retiree) in the Transmittal file
			WHERE 	T200.[request_Card] = 1
			AND	T200.[is_valid_record] = 1
			AND     ISNULL(T200.tran_type, '') NOT IN ('A', 'I', 'D')
			AND     T200.person_pk IS NOT NULL
			AND     T200.desired_aff_pk IS NOT NULL
			AND     T200.desired_aff_pk = T200.existing_aff_pk
			) AS varchar)
	-- 
	-- Request Cards for "Inactive" becoming "Active"						-- 20080227 GP
	-- print 'Inactive_to_Actives';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requested for Change of Status   = '
		+ CAST(	(SELECT count(*) 
			FROM	[dbo].[AUP_Input_BAR] T200
				INNER JOIN afscme_oltp6.dbo.aff_members AM
					ON  AM.person_pk     = T200.person_pk
					AND AM.aff_pk        = T200.existing_aff_pk
					AND AM.mbr_status    = 31002	-- Is "Inactive" on Enterprise
					AND T200.Status_Code = 'A' 	-- And is "Active" (Regular/Retiree) in the Transmittal file
			WHERE 	T200.[request_Card] = 1
			AND	T200.[is_valid_record] = 1
			AND     ISNULL(T200.tran_type, '') NOT IN ('A', 'I', 'D')
			AND     T200.person_pk IS NOT NULL
			AND     T200.desired_aff_pk IS NOT NULL
			AND     T200.desired_aff_pk = T200.existing_aff_pk
			) AS varchar)

	-- Request Cards based on NewCard = 1 in Custom field for UW transmittals			-- 20080731 RW (not a perfect implementation. may be refined later)
	-- print 'Custom cards NewCard = 1';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requested based on NewCard = 1 in Custom field   = '
		+ CAST(	(SELECT count(*) 
			FROM	[dbo].[AUP_Input_BAR] T200
			WHERE 	CASE WHEN ISNULL(CustomName1, '') = 'NEWCARD' THEN ISNULL(CustomValue1, 0) 
						 WHEN ISNULL(CustomName2, '') = 'NEWCARD' THEN ISNULL(CustomValue2, 0) 
						 WHEN ISNULL(CustomName1, '') = 'NEWCARD' THEN ISNULL(CustomValue3, 0) ELSE 0 END = 1
			AND	T200.[request_Card] = 1
			AND	T200.[is_valid_record] = 1
			AND     ISNULL(T200.tran_type, '') NOT IN ('A', 'I', 'D')
			AND		T200.Status_Code = 'A'
			AND     T200.person_pk IS NOT NULL
			AND     T200.desired_aff_pk IS NOT NULL
			AND     T200.desired_aff_pk = T200.existing_aff_pk
			) AS varchar)
	-- End -- 20070524 -- Report Card Requests by category
	-- 
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	------------------------------
	--end of generate script on update
	------------------------------
/************************************************* NEED TO DO ALL THIS ************************	
	-----------------------------------------------------------------------------------------
	-- 20060915 -- GP -- Un_marked_for_deletion_fg if that's the only match we find.
	--			We also update Name and Address
	-----------------------------------------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	SELECT 'UPDATE Person SET marked_for_deletion_fg = 0, member_fg = 1'
			+ ', prefix_nm = ' + CAST(ISNULL(DMV1.com_cd_pk, 0) AS varchar) 
			+ ', first_nm = '''+ LTRIM(RTRIM(CAST(REPLACE(T200.First_Name, '''', '''''') AS varchar))) 
			+ ''', middle_nm = ''' + LTRIM(RTRIM(CAST(REPLACE(T200.Middle_Name, '''', '''''') AS varchar))) 	-- 20070326 GP
			+ ''', last_nm = ''' + LTRIM(RTRIM(CAST(REPLACE(T200.Last_Name, '''', '''''') AS varchar))) 
			+ ''', suffix_nm = ' + CAST(ISNULL(DMV2.com_cd_pk, 0) AS varchar) 
			+', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE person_pk = '
			+ CAST(T200.tppk AS varchar)
			+' AND lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''	
	-- select count(*)	
	FROM 	[dbo].[AUP_Input_BAR] T200
		--INNER JOIN tNEW200_Mbrmas_BL T200		ON T200.mbr_key = TRANS.tmbr_key
		--LEFT OUTER JOIN afscme_oltp6.dbo.Person P	ON P.person_pk = T200.person_pk
		LEFT OUTER JOIN afscme_oltp6.dbo.DM_Code_Mapping_view DMV1 ON DMV1.com_cd_type_key = 'Prefix' AND DMV1.Legacy_Code = T200.Title
		LEFT OUTER JOIN afscme_oltp6.dbo.DM_Code_Mapping_view DMV2 ON DMV2.com_cd_type_key = 'Suffix' AND DMV2.Legacy_Code = T200.Suffix
	WHERE 	T200.ttran_code = 'A'				-- Should be an ADD
	AND 	ISNULL(T200.add_final, '0') IN ('D', 'E', 'F')	-- But we match on Person as marked_for_deletion_fg = 0
	AND	T200.tppk IS NOT NULL				-- And we have a person_pk
	AND	T200.desired_aff_pk IS NOT NULL				-- And we have an aff_pk
	ORDER BY T200.tppk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person UN-marked_for_deletion_fg Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- (/row(s) affected)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	--------------------------------------------------------------
	-- Person Address Update for the Un_marked_for_deletion_fg
	--------------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	SELECT 'UPDATE Person_Address SET addr1 = '''
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr1, '') AS varbinary) AS varchar))), '''', '''''')
	            +''', addr2 = ''' 
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.Addr2, '') AS varbinary) AS varchar))), '''', '''''')
	            +''', city = ''' 
		    + REPLACE(LTRIM(RTRIM(CAST(CAST(ISNULL(T200.City, '') AS varbinary) AS varchar))), '''', '''''') 
		    +''', state = ''' 
	            + CASE
	                        WHEN T200.State IS NULL THEN ''
	                        WHEN T200.State IN ('ZZ', 'XX') THEN ''
	                        WHEN    (LEN(LTRIM(RTRIM(T200.State)))=2) THEN T200.State
	                        ELSE    ''
	                        END +''', zipcode = '''
	            + CASE 
	                        WHEN T200.Zip IS NULL THEN ''
	                        WHEN T200.Zip = '00000' THEN ''
	                        ELSE LTRIM(RTRIM(T200.Zip ))
	                        END + ''', zip_plus = ''' 
	            + CASE
	                        WHEN T200.Zip_4 IS NULL THEN ''
	                        WHEN T200.Zip_4 = '0000' THEN ''
	                        ELSE LTRIM(RTRIM(T200.Zip_4))
	                        END + ''', dept =  '
	            + CAST(@membershipDept2 AS varchar)
	            + ', addr_bad_fg = ' + CASE
	                                                WHEN T200.Addr_Mailable_fg IS NULL THEN '1'
	                                                WHEN T200.Addr_Mailable_fg = 'Y'   THEN '0'
	                                                WHEN T200.Addr_Mailable_fg = 'N'   THEN '1'
	                                                ELSE '1'
	                                                END
	            + ', addr_marked_bad_dt = '+ CASE
	                                                WHEN T200.Addr_Mailable_fg IS NULL THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                WHEN T200.Addr_Mailable_fg = 'Y'   THEN 'NULL'
	                                                WHEN T200.Addr_Mailable_fg = 'N'   THEN ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                ELSE ''''+CONVERT(varchar, @as_of_date, 101) +''''
	                                                END
	            + ', addr_source = ''U'', eff_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
	            + ', end_dt = NULL'		-- 20070925 GP  -- Set Person_Address.end_dt = NULL when we update an existing address
	            + ', addr_source_if_aff_apply_upd = '''+CAST(@reporting_aff_pk AS varchar)+''''
	            + ', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''''
	            + ' WHERE person_pk = ' + CAST(T200.tppk as varchar) 
	            +' AND addr_type = '+ CAST(@PersonAddressType1 AS varchar) + ' AND addr_prmry_fg = 1'
	            +' AND lst_mod_dt <= '''+CONVERT(varchar, @as_of_date, 101) +''''
	-- select  T200.*
	FROM 	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.ttran_code = 'A'				-- Should be an ADD
	AND 	ISNULL(T200.add_final, '0') IN ('D', 'E', 'F')	-- But we match on Person as marked_for_deletion_fg = 0
	AND	T200.tppk IS NOT NULL				-- And we have a person_pk
	AND	T200.desired_aff_pk IS NOT NULL				-- And we have an aff_pk
	ORDER BY T200.tppk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Person_address_updates Generated = '+ CAST(@@ROWCOUNT AS varchar)
	-- (3286 row(s) affected)
	
	
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE person_address SET zip_plus = NULL WHERE LEN(RTRIM(zip_plus)) = 0'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
***********************************************************************/
	-----------------------------------------------------------------------------------------
	-- 20060911 -- GP -- For ADDs we find in Person
				-- Deactivate memberships we do not have in current file for Aff_Tree
	-- 
	-----------------------------------------------------------------------------------------
	--------------------------------------------------------------------------------------------------------------
	-- 20060908 -- GP -- Add aff_member records IF MISSING for ADDs where (SSN+FName OR LName_FName+Addr5) match 
	--------------------------------------------------------------------------------------------------------------
	--  Second Additional New Members 
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'		--mbr_no_old_afscme,
	-- print 'Add aff_member records IF MISSING for ADDs';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO Aff_Members (person_pk, aff_pk, mbr_status, mbr_type,  no_mail_fg, no_cards_fg, no_public_emp_fg, no_legislative_mail_fg, mbr_join_dt, mbr_no_local, primary_information_source, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
		+ CAST(T200.person_pk AS varchar) + ', '
		+ CAST(T200.desired_aff_pk AS varchar) + ', '
		+ CAST(CASE
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('A', 'O', 'P') THEN @Active
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Temporary 	
			WHEN T200.Status_Code = 'N' THEN @Active  
			WHEN T200.Affiliate_Identifier = 'C' THEN @Active
			WHEN T200.Affiliate_Identifier IN ('L', 'U', 'R', 'S') AND T200.Status_Code = 'X' THEN @Active
			WHEN T200.Status_Code = 'Y' THEN @Pending		-- 20170328 -- GP
			ELSE 0 
			END AS varchar) + ', '	
		+ CAST(CASE
			WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A', 'Y') THEN @Regular
			WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A'     ) THEN @Retiree
			WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('T', 'R') THEN @Regular	
			WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('T', 'R') THEN @Retiree	
			WHEN T200.Status_Code = 'N' THEN @AgencyFeePayer	
			WHEN T200.Status_Code = 'C' THEN @UnionShop
			WHEN T200.Status_Code = 'P' THEN @PotentialMember  
			WHEN T200.Status_Code = 'O' THEN @OptOutMember		-- 20170328 -- GP	
	 		WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN @RetireeSpouse	
	 		-- WHEN T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code = 'X' THEN @AssociateMember	
	 		ELSE 0
			END AS varchar) + ', '
		+ CAST(CASE T200.No_Mail_fg WHEN '9' THEN 1 ELSE 0 END AS varchar) + ', '	-- no_mail_fg
		+ CASE										-- no_cards_fg
			WHEN ISNULL(MRI.unit_wide_no_mbr_cards_fg, 0) = 1 THEN '1'
			WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'	-- 20179328 -- GP
			WHEN T200.No_Mail_fg  = '1'                       THEN '1'
			WHEN T200.No_Mail_fg  = '3'                       THEN '1'
	 		ELSE '0'
	 		END+', '
		+ CASE										-- no_public_emp_fg
			WHEN ISNULL(MRI.unit_wide_no_pe_mail_fg, 0) = 1   THEN '1'
	 		WHEN T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code = 'X' THEN '1'
	 		WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'	-- 20179328 -- GP
			WHEN T200.No_Mail_fg  = '2'                       THEN '1'
			WHEN T200.No_Mail_fg  = '3'                       THEN '1'
	 		ELSE '0'
	 		END+', '
		+ CASE										-- no_legislative_mail_fg
			WHEN T200.Status_Code IN ('N', 'P', 'O', 'Y')     THEN '1'	-- 20179328 -- GP
			ELSE '0'
			END+', '
		+ CASE
			WHEN T200.Date_Joined = '' THEN ''''+ CONVERT(varchar, @as_of_date, 101) +''''		-- 20070507 GP
			ELSE ''''+ T200.Date_Joined+''''	-- 20170113 GP+VA
			END + ', ' 
		+ ''''+ CASE
				WHEN (ISNULL(T200.Affiliate_Member_ID, '') = '') THEN ''
				WHEN REPLACE(T200.Affiliate_Member_ID, '0', '') = '' THEN ''
				ELSE LTRIM(RTRIM(T200.Affiliate_Member_ID))
				END + ''', '
		+ CASE
			WHEN (T200.[Information_Source] IS NULL) THEN 'NULL'
			WHEN  T200.[Information_Source] = '' THEN 'NULL'
			WHEN  T200.[Information_Source] = '4' THEN '47004'
			ELSE  '47008'
			END + ', '
		-- + '11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''')'
		+ '11949451, CAST(getdate() AS date), 11949451, CAST(getdate() AS date))'
	FROM	[dbo].[AUP_Input_BAR] T200
		LEFT OUTER JOIN afscme_oltp6.dbo.Aff_Mbr_Rpt_Info MRI
			ON  MRI.aff_pk = T200.desired_aff_pk
	WHERE 	T200.is_valid_record = 1
	AND	ISNULL(T200.tran_type, '') NOT IN ('A', 'I')
	AND	T200.desired_aff_pk <> ISNULL(T200.existing_aff_pk, 0)	-- No Membership Record in desired Affiliate
	AND	T200.person_pk IS NOT NULL
	AND	T200.desired_aff_pk IS NOT NULL
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Aff_Members Inserts Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	------------------------------------------------------------------------------------------------------------------------------------------------
	-- 20060912 -- GP -- Add COM_Weekly_Mbr_Card_Run requests for Aff_Members Inserts Generated for All Person Records we matched.
	------------------------------------------------------------------------------------------------------------------------------------------------
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'INSERT INTO COM_Weekly_Mbr_Card_Run';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO COM_Weekly_Mbr_Card_Run (person_pk, aff_pk, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt) VALUES ('
	            + CAST(T200.person_pk AS varchar) +', '
	            -- + CAST(T200.desired_aff_pk AS varchar) +', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''', 11949451, '''+CONVERT(varchar, @as_of_date, 101) +''')'
	            + CAST(T200.desired_aff_pk AS varchar) +', 11949451, CAST(getdate() AS date), 11949451, CAST(getdate() AS date))'			-- RW 8/26/2019
	FROM	[dbo].[AUP_Input_BAR] T200
	WHERE 	T200.is_valid_record = 1
	AND	ISNULL(T200.tran_type, '') NOT IN ('A', 'I')
	AND	T200.desired_aff_pk <> ISNULL(T200.existing_aff_pk, 0)	-- No Membership Record in desired Affiliate
	AND	T200.person_pk IS NOT NULL
	AND	T200.desired_aff_pk IS NOT NULL
	AND	(   (T200.Affiliate_Identifier IN ('L', 'U') AND T200.Status_Code IN ('A','O'))		-- 20170328 -- GP -- Only request cards for (Regular, Retiree, Union Shop, Retiree Spouse)
		 OR (T200.Affiliate_Identifier IN ('R', 'S') AND T200.Status_Code IN ('A','O','X')))
	ORDER BY T200.person_pk
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Mbr. Card Requests Generated for Affiliate Membership Change of an Existing Person = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	-- Deactivate Inexisting memberships within the current Affiliate Tree (Only the Active Ones!)
	-- 20090727 GP  -- Inactivate all Active Memberships which are NOT in the file for Existing People AND NOT "Should Be A Member"
	-- 20121012 GP  -- Inactivate all Active Memberships which are NOT in the file for Existing People AND NOT "Staff" Either
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'deactivate missing records';									
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE Aff_Members SET mbr_status = '+CAST(@status_inactive AS varchar)+', lst_mod_user_pk = 11949451, lst_mod_dt = '''+CONVERT(varchar, @as_of_date, 101) +''' WHERE person_pk = '+ CAST(T200.person_pk As varchar) + ' AND aff_pk = '+ CAST(AM.aff_pk AS varchar) + ' and mbr_status = '+CAST(@status_active AS varchar) 
	FROM	[dbo].[AUP_Input_BAR] T200
		INNER JOIN [afscme_oltp6].[dbo].[Aff_Members] AM
			ON  AM.[person_pk] = T200.[person_pk]
			AND AM.[aff_pk]   <> T200.[desired_aff_pk]						-- Not the desired_aff_pk
			AND AM.aff_pk IN (SELECT aff_pk FROM OPERATIONS.dbo.[AUP_Aff_Tree_BAR])			-- In Affiliate Tree
			AND AM.mbr_status = @status_active							-- Actives Only
			AND AM.mbr_type =  (CASE WHEN @full_potentials =  1 THEN 82780 END)
			AND AM.mbr_type <> (CASE WHEN @full_potentials <> 1 THEN 82780 END)
			AND AM.mbr_type <> 82354								-- NOT "Should Be A Member"
			AND AM.mbr_type <> 82679								-- NOT "Staff"
			AND AM.aff_pk NOT IN (SELECT 	T.desired_aff_pk					-- Which is NOT an alternate membership
					      FROM	[dbo].[AUP_Input_BAR] T					-- in current file
					      WHERE	T.person_pk = T200.person_pk				-- for this person_pk
					      AND 	T.is_valid_record = 1
					      )
	WHERE 	T200.is_valid_record = 1										-- Valid Records Only
	AND	ISNULL(T200.tran_type, '') NOT IN ('A', 'I', 'D')							-- Not ADDs, Invalids or DELETEs
	-- AND	T200.desired_aff_pk <> ISNULL(T200.existing_aff_pk, 0)							-- No Membership Record in desired Affiliate
	AND	T200.person_pk IS NOT NULL
	AND	T200.desired_aff_pk IS NOT NULL										-- Where we have a Desired_aff_pk
	AND	T200.existing_aff_pk IS NOT NULL									-- That's Overkill!?
	ORDER BY T200.person_pk, AM.[aff_pk]
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Deletes Generated for Existing Persons = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	
	-------------------------------------------------------------------------
	-- 20071025 GP 	-- Add MLBP_Persons Maintenance for OH_L_11 (@reporting_aff_pk = 711)
	-------------------------------------------------------------------------
	IF (@reporting_aff_pk = 711)
	BEGIN
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'DECLARE @com_cd_pk [int]'
--		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT com_cd_pk FROM dbo.Common_Codes WHERE com_cd_type_key = 'TagsOther' AND com_cd_desc = ''Correctional Officers Affiliate'''
		--
		----------------------------------------------------------------
		-- Delete Missing Members of Correctional Officers Affiliates --
		----------------------------------------------------------------
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'DELETE	MLBP
		-- SELECT MLBP.*
		FROM	MLBP_persons MLBP
				INNER JOIN 
				(	Person_Tag_Values PTV_E WITH (NOLOCK)
					INNER JOIN Aff_Organizations AO_E WITH (NOLOCK)
						ON  PTV_E.person_pk = 11949451				-- Affiliate Transmittals
						AND	AO_E.aff_pk = PTV_E.aff_pk
					    -- AND	PTV_E.com_cd_pk = @com_cd_pk			-- TagsOther-Correctional Officers Affiliate
						AND AO_E.parent_aff_fk = 711				-- OH_L_11 -- We actually use the aff_pk
					INNER JOIN Aff_Members AM_E  WITH (NOLOCK)
						ON  AM_E.aff_pk = AO_E.aff_pk
				)	ON MLBP.person_pk = AM_E.person_pk
					AND MLBP.lst_mod_user_pk IN (11949451, 12826192, 10000001, 10000002)	-- We only delete if not Manually Edited
				LEFT OUTER JOIN 
				(
					Person_Tag_Values PTV WITH (NOLOCK)
					INNER JOIN Aff_Organizations AO WITH (NOLOCK)
						ON  PTV.person_pk = 11949451				-- Affiliate Transmittals
						AND	AO.aff_pk = PTV.aff_pk				-- TagsOther-Correctional Officers Affiliate
						-- AND	PTV.com_cd_pk = @com_cd_pk			-- 
						AND AO.parent_aff_fk = 711				-- OH_L_11 -- We actually use the aff_pk
					INNER JOIN Aff_Members AM  WITH (NOLOCK)
						ON  AM.aff_pk = AO.aff_pk
						AND	AM.mbr_status = 31001				-- Active
						AND	AM.mbr_type NOT IN (	29003,			-- Not Agency Fee Payer
										29005,			-- Not Union Shop Objector
										82075,			-- Not Associate Member
										82780			-- NOT Potential Member				-- 20130323 -- 
										)
						AND ISNULL(AM.no_mail_fg, 0) = 0			-- NO no_mail_fg = 1
						AND ISNULL(AM.no_legislative_mail_fg, 0) = 0		-- NO no_legislative_mail_fg = 1
					INNER JOIN Person p WITH (NOLOCK)
						ON  P.person_pk = AM.person_pk 
						AND ISNULL(P.marked_for_deletion_fg, 0) = 0		-- Not marked for deletion
					INNER JOIN Person_Demographics PD WITH (NOLOCK)			-- NOT Deceased
						ON  PD.person_pk = AM.person_pk
						AND ISNULL(PD.deceased_fg, 0) = 0
						AND PD.deceased_dt IS NULL
				)
				ON MLBP.person_pk = AM.person_pk
		WHERE	MLBP.MLBP_mailing_list_pk = 31							-- Correctional Officers
		AND	AM.person_pk IS NULL								-- Not Eligible anymore'
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
		-- INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Deletes from MLBP_Persons for COA = '+ CAST(@@ROWCOUNT AS varchar)
		-- INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
		-------------------------------------------------------------
		-- Add Missing Members of Correctional Officers Affiliates --
		-------------------------------------------------------------
		-- print 'Add Missing Members of Correctional Officers Affiliates';
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'INSERT INTO MLBP_persons (person_pk, MLBP_mailing_list_pk, address_pk, created_user_pk, created_dt, lst_mod_user_pk, lst_mod_dt)
		SELECT DISTINCT AM.person_pk, 31, PA.address_pk, 11949451, GETDATE(), 11949451, GETDATE()
		-- SELECT AM.*
		FROM	Person_Tag_Values PTV WITH (NOLOCK)
				INNER JOIN Aff_Organizations AO WITH (NOLOCK)
					ON  PTV.person_pk = 11949451			-- Affiliate Transmittals
					AND	AO.aff_pk = PTV.aff_pk
					AND AO.parent_aff_fk = 711			-- OH_L_11 -- We actually use the aff_pk
					-- AND	PTV.com_cd_pk = @com_cd_pk		-- TagsOther-Correctional Officers Affiliate
				INNER JOIN Aff_Members AM  WITH (NOLOCK)
					ON  AM.aff_pk = AO.aff_pk
					AND	AM.mbr_status = 31001			-- Active
					AND	AM.mbr_type NOT IN (	29003,		-- Not Agency Fee Payer
									29005,		-- Not Union Shop Objector
									82075,		-- Not Associate Member
									82780		-- NOT Potential Member				-- 20130323 -- 
									)
					AND ISNULL(AM.no_mail_fg, 0) = 0		-- NO no_mail_fg = 1
					AND ISNULL(AM.no_legislative_mail_fg, 0) = 0	-- NO no_legislative_mail_fg = 1
				INNER JOIN Person p WITH (NOLOCK)
					ON  P.person_pk = AM.person_pk 
					AND ISNULL(P.marked_for_deletion_fg, 0) = 0	-- Not marked for deletion
				INNER JOIN Person_Demographics PD WITH (NOLOCK)		-- NOT Deceased
					ON  PD.person_pk = AM.person_pk
					AND ISNULL(PD.deceased_fg, 0) = 0
					AND PD.deceased_dt IS NULL
				INNER JOIN Person_SMA SMA WITH (NOLOCK)
					ON  SMA.person_pk = AM.Person_pk
					AND SMA.current_fg = 1
				INNER JOIN Person_Address PA WITH (NOLOCK)
					ON  PA.address_pk = SMA.address_pk
					AND ISNULL(PA.addr_bad_fg, 0) = 0		-- NO bad addresses		
				LEFT OUTER JOIN MLBP_persons MLBP WITH (NOLOCK)
					ON  AM.person_pk = MLBP.person_pk
					AND	MLBP.MLBP_mailing_list_pk = 31		-- Correctional Officers
		WHERE	MLBP.person_pk IS NULL						-- Not already on List=31'
		-- INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of ADDs to MLBP_Persons for COA = '+ CAST(@@ROWCOUNT AS varchar)
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
		INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	END

	----------------------------------------------------------------------------------------------------------
	-- 20200810 RW  -- Propagate UW_ID/AffiliateMemberID through entire affiliate tree for UWare transmittal
	----------------------------------------------------------------------------------------------------------
	-- RW: Have to use restricted aff_tree so applicable for C37 6-locals fulll transmittal and NUHHCE subunit only transmittal
	-- RW 9/3/2020: Found not working for PA C13, an admin council that has 470 aff_pks. So exclude PA C13 until we know to increate @aff_pks to varchar(MAX) in package of report
 
	IF (ISNULL(@data_source, '') = 'UWare') AND @reporting_aff_pk NOT IN (143, 4571)			-- excluding CASEA and PA C13 (admin C) for now
	BEGIN

		DECLARE @aff_pks varchar(4000)
		SET @aff_pks = '' 

		SELECT @aff_pks = @aff_pks + CAST(aff_pk AS VARCHAR) + ','
		FROM [OPERATIONS].[dbo].[AUP_Aff_Tree_BAR]

		-- SELECT @aff_pks

		IF RIGHT(@aff_pks,1) = ','
		SELECT @aff_pks = LEFT(@aff_pks, LEN(@aff_pks)-1)

		-- SELECT @aff_pks

		IF LEN(@aff_pks) > 0
		BEGIN

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Propagate UW_ID/AffiliateMemberID through entire affiliate tree for UWare transmittal'

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'UPDATE am SET mbr_no_local = am2.mbr_no_local FROM dbo.Aff_Members am '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'INNER JOIN (SELECT person_pk, mbr_no_local FROM dbo.Aff_Members WHERE aff_pk IN ' 
			
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, '(' + @aff_pks + ') '
			
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'AND LEN(ISNULL(mbr_no_local,'''')) > 0 AND mbr_status = 31001) am2 ON am2.person_pk = am.person_pk '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'WHERE am.aff_pk IN '
			
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, '(' + @aff_pks + ') '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'AND LEN(ISNULL(am.mbr_no_local,'''')) = 0'

			/*
			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Propagate UW_ID/AffiliateMemberID through entire affiliate tree for UWare transmittal'

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'UPDATE am SET mbr_no_local = am2.mbr_no_local FROM dbo.Aff_Members am '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'INNER JOIN (SELECT person_pk, mbr_no_local FROM dbo.Aff_Members WHERE aff_pk IN (' + @aff_pks + ') AND LEN(ISNULL(mbr_no_local,'''')) > 0 AND mbr_status = 31001) am2 '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'ON am2.person_pk = am.person_pk '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'WHERE am.aff_pk IN (' + @aff_pks + ') '

			INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) 
			SELECT @load_id, 'AND LEN(ISNULL(am.mbr_no_local,'''')) = 0'
			*/

		END

	END    

	--------------------------------------------------------------------------------------------------
	-- 20110209 GP  -- Keep in Sync Officer_History.[pos_address_from_person_pk] w. Current_SMA
	--------------------------------------------------------------------------------------------------
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- '-- SET NOCOUNT ON'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE OH SET OH.[pos_addr_from_person_pk] = SMA.[address_pk] FROM [dbo].[Officer_History] OH INNER JOIN [dbo].[Person] P WITH(NOLOCK) ON P.[person_pk] = OH.[person_pk] AND ISNULL(P.[marked_for_deletion_fg], 0) = 0 INNER JOIN [dbo].[Person_SMA] SMA WITH(NOLOCK) ON SMA.[person_pk] = P.[person_pk] AND SMA.[current_fg] = 1 WHERE	(OH.[pos_end_dt] IS NULL OR OH.[pos_end_dt] > GETDATE()) AND OH.[pos_addr_from_person_pk] IS NOT NULL AND OH.[pos_addr_from_person_pk] <> SMA.[address_pk]'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Number of Officer_History.[pos_address_from_person_pk] Updated to SMA = '+ CAST(@@ROWCOUNT AS varchar)
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- SET NOCOUNT OFF'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	----------------------
	----------------------
	----------------------
/* -- GP -- 20170306
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- ADDs --'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_cnt 		= count(*) FROM Person WHERE person_pk BETWEEN @first_person_pk AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_Address_cnt	= count(*) FROM Person_Address WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_SMA_cnt 		= count(*) FROM Person_SMA WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_Phone_cnt 		= count(*) FROM Person_Phone WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_Email_cnt 		= count(*) FROM Person_Email WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_Demographics_cnt 	= count(*) FROM Person_Demographics WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Person_PolLegisl_cnt = count(*) FROM Person_Political_Legislative WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Aff_Members_cnt           = count(*) FROM Aff_Members WHERE person_pk  BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT Weekly_Mbr_Card_Run_cnt   = count(*) FROM COM_Weekly_Mbr_Card_Run WHERE person_pk BETWEEN @first_person_pk  AND @last_person_pk'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT ''Num_Mbrs'' = count(*), AO.aff_pk, AO.parent_aff_fk, AO.aff_type, AO.aff_localSubChapter, AO.aff_stateNat_type, AO.aff_subUnit, AO.aff_councilRetiree_chap, AO.old_aff_unit_cd_legacy
	FROM 	Aff_Members AM
		INNER JOIN Aff_Organizations AO
			ON AO.aff_pk = AM.aff_pk
	WHERE person_pk  BETWEEN @first_person_pk AND @last_person_pk 
	GROUP BY AO.aff_pk, AO.parent_aff_fk, AO.aff_type, AO.aff_localSubChapter, AO.aff_stateNat_type, AO.aff_subUnit, AO.aff_councilRetiree_chap, AO.old_aff_unit_cd_legacy
	ORDER BY AO.aff_pk, AO.parent_aff_fk'
	-- 
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT ''Num_Mbrs'' = count(*), AM.mbr_status, ''Mbr_Status_Description'' = CC.com_cd_desc
		FROM 	Aff_Members AM
			LEFT OUTER JOIN Common_Codes CC
				ON CC.com_cd_pk = AM.mbr_status
		WHERE AM.aff_pk IN (SELECT 	VAT.aff_pk
						FROM 	dbo.V_aff_tree_AdminC VAT
						WHERE	VAT.Root_aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						OR 	VAT.GParent_aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						OR	VAT.Parent_aff_pk  = '+CAST(@reporting_aff_pk AS varchar)+'
						OR 	VAT.aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						)
		GROUP BY AM.mbr_status, CC.com_cd_desc
		ORDER BY AM.mbr_status'
	-- 
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'SELECT ''Num_Mbrs'' = count(*), AM.mbr_type, ''Mbr_Type_Description'' = CC.com_cd_desc
		FROM 	Aff_Members AM
			LEFT OUTER JOIN Common_Codes CC
				ON CC.com_cd_pk = AM.mbr_type
		WHERE AM.aff_pk IN (SELECT 	VAT.aff_pk
						FROM 	dbo.V_aff_tree_AdminC VAT
						WHERE	VAT.Root_aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						OR 	VAT.GParent_aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						OR	VAT.Parent_aff_pk  = '+CAST(@reporting_aff_pk AS varchar)+'
						OR 	VAT.aff_pk = '+CAST(@reporting_aff_pk AS varchar)+'
						)
		GROUP BY AM.mbr_type, CC.com_cd_desc
		ORDER BY AM.mbr_type'
	
	--								 -- 20070416 GP
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
	--
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
*/
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'UPDATE AUP_Job_Log SET [posted] = 1, lst_mod_dt = GETDATE() WHERE Job_Type = ''Affiliate Transmittal'' AND Load_ID = '+CAST(@load_id AS varchar)
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	--
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'
	-- print 'Final Stats';
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Counts numbers of active members by type after the transmittal and generates report (only on AFSSQL1604 and AFSSQL1604\AFSCME_TEST)'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'IF @@SERVERNAME IN (''AFSSQL1604'', ''AFSSQL1604\AFSCME_TEST'')'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    BEGIN'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''' + ''''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''Use the following SQL query to view the report of Changes in Number of Active Members by Bype (only on AFSSQL1604 and AFSSQL1604\AFSCME_TEST):'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''' + ''''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''    EXEC [OPERATIONS].[dbo].[sp_Member_Count_Before_&_After_Transmittal] @load_id = ' + CAST(@load_id AS varchar)  + ', @reporting_aff_pk = ' + CAST(@reporting_aff_pk AS varchar) + ', @view_only = 1'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    --'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    EXEC [OPERATIONS].[dbo].[sp_Member_Count_Before_&_After_Transmittal] @load_id = ' + CAST(@load_id AS varchar) + ', @reporting_aff_pk = ' + CAST(@reporting_aff_pk AS varchar) + ', @before_after = ' + '''after'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    END'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '-- Generates report of member type conversions (only on AFSSQL1604 and AFSSQL1604\AFSCME_TEST)'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'IF @@SERVERNAME IN (''AFSSQL1604'', ''AFSSQL1604\AFSCME_TEST'')'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    BEGIN'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''' + ''''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''Use the following SQL query to view the report of Member Type Conversions:'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''' + ''''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''    EXEC [OPERATIONS].[dbo].[sp_Member_Type_Conversion_by_Transmittal] @load_id = ' + CAST(@load_id AS varchar)  + ', @reporting_aff_pk = ' + CAST(@reporting_aff_pk AS varchar) + ', @view_only = 1'''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    --'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    EXEC [OPERATIONS].[dbo].[sp_Member_Type_Conversion_by_Transmittal] @load_id = ' + CAST(@load_id AS varchar) + ', @reporting_aff_pk = ' + CAST(@reporting_aff_pk AS varchar) + ', @transmittal_date = ''' + CONVERT(varchar, @as_of_date, 101) + ''''
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    END'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '--'

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'END'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'ELSE'
	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '    PRINT ''Load_ID = '+CAST(@load_id AS varchar)+' has already been executed in this database.'''
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'
--	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, '------------------'

	INSERT INTO [dbo].[AUP_Code_BAR] ([Load_ID],[Code]) SELECT @load_id, 'NOEXECUTION:'	
	----------------------------------------------------------------------
	----------------------------------------------------------------------
	----------------------------------------------------------------------
	--
	--
	SET NOCOUNT OFF
	--	
	RETURN (SELECT count(*) FROM [dbo].[AUP_Code_BAR])
	--
END
/*
USAGE:

	SET NOCOUNT ON
	DECLARE @load_id int
	SELECT @load_id = (select TOP 1 [Load_ID] from AUP_Report_Log_Bar ORDER BY [Load_ID] DESC)
	SELECT '@load_id = '+ CAST(@load_id AS varchar)
	--
	DECLARE @return_value int
	EXEC  @return_value = [dbo].[AUP_Generate_Code_BAR] @load_id
	SELECT '@return_value = '+CAST(@return_value AS varchar)
	-- SET NOCOUNT OFF

---------------------------------------------------------------
-- EXEC master..xp_cmdshell 'DTSRUN /SVNA016\DEVDB1 /E /NAUP_Generated_Code'
-- EXEC master..xp_cmdshell 'DTSRUN /SVNA016\DEVDB1 /E /NAUP_Generated_Code_Summary'
-- 
-- TRUNCATE TABLE AUP_Code_BAR
-- select * from AUP_PK_Admin
-- select * from AUP_Code_BAR
-- select * from AUP_RAW_BAR_Hist
-- select * from AUP_Input_Hist
-- select * from AUP_Input_BAR_Hist
-- SELECT * FROM [dbo].[AUP_Report_Log_Bar]
---------------------------------------------------------------
*/
GO

