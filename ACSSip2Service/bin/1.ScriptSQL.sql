
/****** Object:  UserDefinedFunction [dbo].[DecodeUTF8String]    Script Date: 03/26/2021 4:44:05 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create FUNCTION [dbo].[DecodeUTF8String] (@value varchar(max))
RETURNS nvarchar(max)
AS
BEGIN
    -- Transforms a UTF-8 encoded varchar string into Unicode
    -- By Anthony Faull 2014-07-31
    DECLARE @result nvarchar(max);

    -- If ASCII or null there's no work to do
    IF (@value IS NULL
        OR @value NOT LIKE '%[^ -~]%' COLLATE Latin1_General_BIN
    )
        RETURN @value;

    -- Generate all integers from 1 to the length of string
    WITH e0(n) AS (SELECT TOP(POWER(2,POWER(2,0))) NULL FROM (VALUES (NULL),(NULL)) e(n))
        , e1(n) AS (SELECT TOP(POWER(2,POWER(2,1))) NULL FROM e0 CROSS JOIN e0 e)
        , e2(n) AS (SELECT TOP(POWER(2,POWER(2,2))) NULL FROM e1 CROSS JOIN e1 e)
        , e3(n) AS (SELECT TOP(POWER(2,POWER(2,3))) NULL FROM e2 CROSS JOIN e2 e)
        , e4(n) AS (SELECT TOP(POWER(2,POWER(2,4))) NULL FROM e3 CROSS JOIN e3 e)
        , e5(n) AS (SELECT TOP(POWER(2.,POWER(2,5)-1)-1) NULL FROM e4 CROSS JOIN e4 e)
    , numbers(position) AS
    (
        SELECT TOP(DATALENGTH(@value)) ROW_NUMBER() OVER (ORDER BY (SELECT NULL))
        FROM e5
    )
    -- UTF-8 Algorithm (http://en.wikipedia.org/wiki/UTF-8)
    -- For each octet, count the high-order one bits, and extract the data bits.
    , octets AS
    (
        SELECT position, highorderones, partialcodepoint
        FROM numbers a
        -- Split UTF8 string into rows of one octet each.
        CROSS APPLY (SELECT octet = ASCII(SUBSTRING(@value, position, 1))) b
        -- Count the number of leading one bits
        CROSS APPLY (SELECT highorderones = 8 - FLOOR(LOG( ~CONVERT(tinyint, octet) * 2 + 1)/LOG(2))) c
        CROSS APPLY (SELECT databits = 7 - highorderones) d
        CROSS APPLY (SELECT partialcodepoint = octet % POWER(2, databits)) e
    )
    -- Compute the Unicode codepoint for each sequence of 1 to 4 bytes
    , codepoints AS
    (
        SELECT position, codepoint
        FROM
        (
            -- Get the starting octect for each sequence (i.e. exclude the continuation bytes)
            SELECT position, highorderones, partialcodepoint
            FROM octets
            WHERE highorderones <> 1
        ) lead
        CROSS APPLY (SELECT sequencelength = CASE WHEN highorderones in (1,2,3,4) THEN highorderones ELSE 1 END) b
        CROSS APPLY (SELECT endposition = position + sequencelength - 1) c
        CROSS APPLY
        (
            -- Compute the codepoint of a single UTF-8 sequence
            SELECT codepoint = SUM(POWER(2, shiftleft) * partialcodepoint)
            FROM octets
            CROSS APPLY (SELECT shiftleft = 6 * (endposition - position)) b
            WHERE position BETWEEN lead.position AND endposition
        ) d
    )
    -- Concatenate the codepoints into a Unicode string
    SELECT @result = CONVERT(xml,
        (
            SELECT NCHAR(codepoint)
            FROM codepoints
            ORDER BY position
            FOR XML PATH('')
        )).value('.', 'nvarchar(max)');

    RETURN @result;
END
 

 --			[SEARCH_CODE_TITLE_SEL] '15912,15905',1
GO
/****** Object:  StoredProcedure [dbo].[SP_ADMIN_GET_RIGHTS]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create PROCEDURE [dbo].[SP_ADMIN_GET_RIGHTS]
-- Purpose: Get the basic user rights 
-- MODIFICATION HISTORY
-- Person      Date    Comments
-- Tuannn      231104  Get the basic user rights 
-- ---------   ------  -------------------------------------------
	@intModuleID int,
	@intUID int,
	@intParentID int
AS	
	DECLARE @intCount int

	IF @intModuleID <> 0 
	   BEGIN
		-- Second Grade Admin (Parent ID is not 0) or not
	/*	IF @intParentID = 0 
			SELECT ID, [Right] FROM SYS_USER_RIGHT WHERE ModuleID = @intModuleID
		ELSE
		-- Nguoi dung duoi quyen chi duoc cap nhung quyen cua admin cap 2
			SELECT B.ID, B.[Right] FROM SYS_USER_RIGHT_DETAIL A, SYS_USER_RIGHT B
			WHERE ModuleID = @intModuleID and A.UserID = @intParentID and A.RightID = B.ID
 */
	 select 1
	   END
	  
	   
	ELSE
		BEGIN
		/*
		   IF @intUID = 0 
			BEGIN
				-- Second Grade Admin (Parent ID is not 0) or not
				IF @intParentID = 0 
					SELECT ID FROM SYS_USER_RIGHT WHERE isBasic = 1
				ELSE
				-- Nguoi dung duoi quyen chi duoc cap nhung quyen co ban (luc dau) cua admin cap 2
					SELECT B.ID FROM SYS_USER_RIGHT_DETAIL A, SYS_USER_RIGHT B
					WHERE isBasic = 1 AND A.UserID = 1 and A.RightID = B.ID
			END
		   ELSE
			BEGIN
				SELECT @intCount = ISNULL(Count(RightID),0) FROM SYS_USER_RIGHT_DETAIL 
					WHERE UserID = @intUID
				IF @intCount <> 0 
					SELECT RightID FROM SYS_USER_RIGHT_DETAIL WHERE UserID = @intUID
				ELSE
					SELECT ID FROM SYS_USER_RIGHT WHERE isBasic = 1
			END
			*/
			select 1
		END

GO
/****** Object:  StoredProcedure [dbo].[SP_CHECKCOPYNUMBER]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[SP_CHECKCOPYNUMBER]
-- Purpose: Check copynumber, return:
	-- 0: OK
	-- 1: Copynumber doesn't exists
	-- 2: Copynumber is locked or not in circulation
	-- 3: Copynumber is on load
	-- 4: Copynumber is on hold
	-- 5: Librarian has not permission to manage location of the CopyNumber
	-- 6: Librarian has not permission to manage location of Patron
-- MODIFICATION HISTORY
-- Person      Date    Comments
-- Oanhtn	190804		Create
-- Sondp	20/11/2005	Modified

-- Person      Date             Comments
-- chuyenpt    15-01-2007       Edit(Replace table: SYS_USER_LOCATION by table: SYS_USER_CIR_LOCATION )
--                              (Loc ra cac kho ma phan he MUON TRA duoc phep quan ly)

-- ---------   ------  -------------------------------------------       [SP_CHECKCOPYNUMBER] 1,'1725202010040','DH14001350',0
	@intUserID	int,
	@strPatronCode	varchar(30),
	@strCopyNumber	varchar(30),
	@intOutPut	int	OUT
   -- Declare program variables as shown above    
AS
	DECLARE @intCount int
		SET @intOutPut = 0	
	SELECT @intCount=COUNT(*) FROM Block_Ban_doc l inner join Ban_doc b on l.The_ID=b.ID	WHERE So_the like @strPatronCode
	if @intCount>0 
	begin
		set @intOutPut=7
		return
	end
	SELECT @intCount=COUNT(*) FROM Ma_xep_gia WHERE Ma_xep_gia like @strCopyNumber and LoanTypeID=5;
IF @intCount > 0
begin
print 'a'
			SET @intOutPut = 8  
			--print @intOutPut
			--RETURN 
	end

			SELECT @intCount=COUNT(*) FROM Lich_su_muon_sach WHERE Tai_lieu_ID in (select Tai_lieu_ID from Ma_xep_gia where Ma_xep_gia like @strCopyNumber) and So_the_ID in (SELECT ID FROM Ban_doc WHERE So_the  like @strPatronCode) and DATEDIFF(HH,Ngay_tra,getdate())<24
IF @intCount > 0
begin
			SET @intOutPut = 9 
			RETURN 
			end
	-- Check permission manage location
	-- Check permission manage location
	if @strPatronCode <>''
	begin
		SELECT @intCount = COUNT(*) FROM Ma_xep_gia WHERE Ma_xep_gia like @strCopyNumber
		/*AND LocationID IN
		(SELECT HOLDING_LOCATION.ID FROM HOLDING_LOCATION  Where ID IN
		(SELECT A.LocationID
		FROM CIR_PATRON_GROUP_LOC A     
		LEFT JOIN HOLDING_LOCATION B    
		ON A.LocationID=B.ID    
		WHERE PatronGroupID  = @intPatronGroupID))
		*/
	
		IF @intCount = 0
		BEGIN
			SET @intOutPut = 6
			
			RETURN
		END
	end
	else
	begin
		-- check for sip2server
		-- copynumber is not in circulation
		SELECT @intCount=COUNT(*) FROM An_pham_cho_muon WHERE Ma_xep_gia like @strCopyNumber;
IF @intCount > 0
begin
			SET @intOutPut = 3  
			RETURN 
			end
			
	end

	

	
	print @intOutPut
GO
/****** Object:  StoredProcedure [dbo].[SP_CHECKIN]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
    CREATE   PROCEDURE [dbo].[SP_CHECKIN]
-- Purpose: Checkin copies
-- MODIFICATION HISTORY
-- Person      Date    Comments
-- Oanhtn      060904  Create
-- ---------   ------  -------------------------------------------
@intAutoPaid		int,
	@intUserID		int,	
	@strCopyNumbers		varchar(500),
	@strCheckInDate		varchar(30),
	@strPatronCode		varchar(30) OUT,
	@strTransIDs		varchar(300) OUT,
	@intError		int	OUT,
	@intLoanFees 	int	OUT,
	@intTotalLoanFees	int	OUT
AS
set @intLoanFees=0
set @intTotalLoanFees=0
	DECLARE @strCopyNumber	varchar(30)
	DECLARE @dblFees	money
	DECLARE @dblFines	money
	DECLARE @dblTemp	money
	DECLARE @intNextID	int
	DECLARE @intPatronRes	int
	DECLARE @intPos		int
	DECLARE @intTransID	int
	DECLARE @intNextTransID	int
	DECLARE @strTempIDs	varchar(300)
	DECLARE @intOutPut      int
	DECLARE @intTimeUnit	int

	-- For UPDATE RESERVATION HOLDING ONLY
	DECLARE @intItemID int
	DECLARE @intCount int
	DECLARE @intTargetID int
	DECLARE @strLocationID varchar(30)
	DECLARE @intHoldTurnTimeOut int
	DECLARE @dteDate DateTime

	SET @strTransIDs=''
	SET @strPatronCode=''
	SET @intError=0
	SET @intOutPut=0
	BEGIN
	-- New transaction
	BEGIN TRAN

	SELECT @intTransID = ID FROM An_pham_cho_muon WHERE Ma_xep_gia IN (@strCopyNumber)

			-- Move onloan copies to loan history copies
			SELECT @intTimeUnit = C.TimeUnit FROM An_pham_cho_muon A, LoanType C WHERE A.Ma_xep_gia IN (@strCopyNumber) AND C.LoanTypeID = A.LoanType

			IF @intTimeUnit = 1 -- Ngay
			BEGIN
				
					INSERT INTO Lich_su_muon_sach( Ma_xep_gia,Ngay_muon,Ngay_tra,So_the_ID,Tai_lieu_ID,So_ngay_qua_han,Tien_phat) 
					SELECT  Ma_xep_gia, Ngay_muon, @strCheckInDate  AS CheckInDate, So_the_ID, Tai_lieu_ID,  
					DATEDIFF(Day, ISNULL(Ngay_tra, @strCheckInDate), GetDate()) AS OverdueDays, DATEDIFF(Day, ISNULL(Ngay_tra,@strCheckInDate), 
					GetDate())*OverdueFine AS OverdueFine FROM An_pham_cho_muon A, LoanType B WHERE A.Ma_xep_gia IN (@strCopyNumber) 
					--AND A.LoanTypeID = B.ID
			
			UPDATE Ma_xep_gia SET INUSED = 0 WHERE Ma_xep_gia IN (@strCopyNumber)


			-- Update Holding_copies
			UPDATE So_luong_an_pham SET So_ban_roi = So_ban_roi - 1 WHERE id IN (SELECT Tai_lieu_ID FROM An_pham_cho_muon WHERE Ma_xep_gia IN (@strCopyNumber))
			-- Delete onloan copies
			DELETE An_pham_cho_muon WHERE Ma_xep_gia IN (@strCopyNumber)
			END
			
	
SET @intError = 0
	PRINT @strTempIDs
	IF NOT @@ERROR = 0
	BEGIN
		SET @strTransIDs = ''
		SET @intError = @@ERROR
		ROLLBACK TRAN
	END
	ELSE
	BEGIN
		SET @intError = 0
		COMMIT TRAN
	END
	END
	


	
GO
/****** Object:  StoredProcedure [dbo].[SP_CHECKOUT]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
 CREATE PROCEDURE [dbo].[SP_CHECKOUT]  
-- Purpose: Checkout copy  
--  ALTER Loan record  
--  Update Copynumber' status  
--  Calculation FreeCopies  
--  Delete the request record if it exist  
--  Remove the holding record  
--  Process holding queue  
-- MODIFICATION HISTORY  
-- Person      Date    Comments  
-- Oanhtn      230804  ALTER  
--	@strPatronCode='cb0005',@intLoanMode=1,@intUserID=6,@strCopyNumber='DH10002662',@strDueDate='2/20/2020 10:31:33',@strCheckOutDate='2/20/2020 10:31:33',
--@intHoldIgnore=1,@strFees='',@intOutValue=@p9 output,@intOutID=@p10 output
-- ---------   ------  -------------------------------------------         
	 @strPatronCode  	varchar(30),  
	  @intLoanMode  		INT,
	 @intUserID  		INT,
	  	 @strCopyNumber  	varchar(30),  
	 @strDueDate  		varchar(30),  
	 @strCheckOutDate 	varchar(30),  
	 @intHoldIgnore  	INT,  
	 @strFees varchar(30),
	 @intOutValue  		INT OUT,
	 @intOutID		INT OUT  
AS  
	 DECLARE @lngID 	INT  
	 DECLARE @lngPatronID 	INT  
	 DECLARE @intCount 	INT
	 DECLARE @dteDate 	DateTime
	 DECLARE @dteSDate 	DateTime
	 DECLARE @intLoanTypeID INT
	 DECLARE @intLocationID INT
	 DECLARE @intItemID  	INT
	 DECLARE @intTimeUnit 	INT
	 DECLARE @intLOANPERIOD INT  
	 DECLARE @intOutValue1 	INT


	 SET @intOutID=0
	 SET @dteDate = convert(datetime,@strCheckOutDate)
	 
	 -- kiem tra tinh hop le cua copynuber trong truong hop khong check o form
    	 --lent : 12-5-2007
    	 IF @intUserID>0 
	 BEGIN
		--EXEC SP_CHECKCOPYNUMBER @intUserID,@strPatronCode,@strCopyNumber,@intOutPut=@intOutValue output
		set @intOutValue=0
         	IF @intOutValue<>0  -- khong hop le
		BEGIN
            		SET @intOutValue=@intOutValue
         		RETURN
		END
		print @intOutValue
		SET @intOutValue1 = @intOutValue

		--EXEC SP_CHECKPATRONCODE @intUserID,@strPatronCode,@intLoanMode,@intOutPut=@intOutValue output
		--print 1
		IF @intOutValue=3 OR @intOutValue=6
		BEGIN
			SET @intOutValue=10
			RETURN
		END
		ELSE
			SET @intOutValue=@intOutValue1
		SELECT @intLoanTypeID=loantypeid,@intLocationID=Kho_ID,@intItemID=Tai_lieu_ID FROM  Ma_xep_gia WHERE Ma_xep_gia like (@strCopyNumber)
		SELECT @intTimeUnit=TimeUnit,@intLOANPERIOD=LOANPERIOD FROM LoanType WHERE LoanTypeID=@intLoanTypeID
         	IF @intTimeUnit=1  -- NGAY
            		SET @dteDate = DateAdd("d", @intLOANPERIOD, @dteDate)
         	ELSE -- GIO
			--SET @dteDate = DateAdd("hh", @intLOANPERIOD/24, @dteDate) 
            		SET @dteDate = DateAdd("hh", @intLOANPERIOD, @dteDate) 

						 print 2
         END
   	 ELSE
		SELECT @intLoanTypeID=loantypeid,@intLocationID=Kho_ID,@intItemID=Tai_lieu_ID FROM Ma_xep_gia WHERE Ma_xep_gia like (@strCopyNumber)
	 -- Get NextID   
	 SELECT @lngID = ISNULL(MAX(ID), 0) + 1 FROM An_pham_cho_muon  
	 SET @intOutID = @lngID  
  	 print 3
 	 -- Get PatronID  
 	 SELECT @lngPatronID = ID FROM Ban_doc WHERE So_the like (@strPatronCode)  
 	 
	 -- ALTER Loan record  
	 --IF (@strDueDate = 'NULL' OR @strDueDate = ''  ) and (@intUserID=0)
	 -- print @strDueDate + '---'
	 IF (@strDueDate = 'NULL' OR @strDueDate = '') -- Han tra mo
		INSERT INTO An_pham_cho_muon(ID, Tai_lieu_ID, Ma_xep_gia, Ngay_muon,Ngay_tra,LoanType, CirID, So_the_ID,So_luot_gia_han) VALUES (@lngID, @intItemID, UPPER(@strCopyNumber), @strCheckOutDate,DATEADD(d,10,getdate()),  @intLoanMode, @intLoanTypeID, @lngPatronID,0)  
 	 ELSE  
  	 BEGIN  
		-- Ngay tra duoc nhap vao
	 	IF @strDueDate <> '' --AND CONVERT(DATETIME,@strDueDate) > CONVERT(DATETIME,@strCheckOutDate)
			INSERT INTO An_pham_cho_muon(ID, Tai_lieu_ID, Ma_xep_gia, Ngay_muon,Ngay_tra,LoanType, CirID, So_the_ID,So_luot_gia_han)  VALUES (@lngID, @intItemID, UPPER(@strCopyNumber), @strCheckOutDate, @strDueDate, @intLoanMode, @intLoanTypeID, @lngPatronID,0)  
		ELSE
		-- Ngay tra khong duoc nhap
		BEGIN
			IF @intUserID=0 
            			SET @dteDate = CONVERT(DATETIME,@strDueDate,103)
/*
        		SET @dteSDate = @dteDate 
	        	WHILE (SELECT DISTINCT(OffDay) from CIR_CALENDAR WHERE OffDay = @dteDate AND LOCATIONID = @intLocationID) IS NOT NULL  
			BEGIN
		     		IF (SELECT DISTINCT(OffDay) from CIR_CALENDAR WHERE OffDay = @dteDate AND LOCATIONID = @intLocationID) IS NULL  
      					BREAK  
     				ELSE  
     				BEGIN  
					SET @dteSDate = @dteDate 
      					SELECT @dteDate = DateAdd("d", 1, @dteDate)  
      					CONTINUE   
     				END  
		 	END
	 		-- ALTER Loan record   
	 		INSERT INTO CIR_LOAN (ID, ItemID, CopyNumber, CheckOutDate, DueDate, LoanMode, LoanTypeID, PatronID, LocationID) VALUES (@lngID, @intItemID, UPPER(@strCopyNumber), @strCheckOutDate, @dteDate, @intLoanMode, @intLoanTypeID, @lngPatronID, @intLocationID) 
	*/
		END
	 END  
  
	 -- Update Copynumber' status  
	 UPDATE Ma_xep_gia SET InUsed = 1, OnHold = 0, UseCount = UseCount + 1, DateLastUsed = GetDate() WHERE Ma_xep_gia like (@strCopyNumber)  
	 -- Calculation FreeCopies  
	 UPDATE So_luong_an_pham SET So_ban_roi = So_ban_roi - 1 WHERE ID = @intItemID  

	 set @intOutValue=1
	 /*
	 -- Delete the request record if it exist  
	 DELETE CIR_RESERVE_REQUEST WHERE PatronID = @lngPatronID AND ItemID = @intItemID  
	 -- Remove the holding record  
	 DELETE CIR_HOLDING WHERE UPPER(PatronCode) = UPPER(@strPatronCode) AND ItemID = @intItemID  
 
 	 -- If the item is on hold and it is checked out to another patron, so push the holding patron back to queue  
	 IF @intHoldIgnore = 1  
	 BEGIN  
		 -- SELECT @lngID = ID FROM CIR_HOLDING WHERE ItemID = @lngItemID AND inturn = 1 AND UPPER(CopyNumber) = UPPER(@strCopyNumber)  
		  IF NOT @lngID IS NULL  
		  BEGIN  
			   UPDATE CIR_HOLDING SET inturn = 0, TimeOutDate = NULL, CopyNumber = '' WHERE ID = @lngID AND CopySpecified = 0  
			   UPDATE CIR_HOLDING SET inturn = 0, TimeOutDate = NULL WHERE ID = @lngID AND CopySpecified = 1  
		  END  
	 END
 */
GO
/****** Object:  StoredProcedure [dbo].[SP_CHECKPATRONCODE]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[SP_CHECKPATRONCODE]
-- Purpose: Check Patron by PatronCode & return all main information when found:
	-- 0: OK
	-- 1: Card doesn't exists
	-- 2: Card expired
	-- 3: Quota exceeded (LoanMode=1)
	-- 4: Card is locked
	-- 5: Patron doesn't has access permission to one of the locations which this librarian has manage permission
	-- 6: Quota exceeded (LoanMode=2)
-- MODIFICATION HISTORY
-- Person      Date    Comments
-- Oanhtn      190804  Create
-- Oanhtn      270804  Update (return 5)
-- ---------   ------  -------------------------------------------       
	@intUserID	int,
	@strPatronCode	varchar(30),
	@intLoanMode	int,
	@intOutPut	int	OUT
   -- Declare program variables as shown above    
AS
	DECLARE @intCount int
	DECLARE @intPatronGroupID int

	SET @intOutPut = 0 -- Default
	-- Check exists
	SELECT @intCount = COUNT(*) FROM Ban_doc WHERE UPPER(So_the) = UPPER(@strPatronCode)
	IF @intCount = 0
	BEGIN
		SET @intOutPut = 1
--print(@intOutPut)
		RETURN @intOutPut
	END
	
	-- Check Locked
	SELECT @intCount = COUNT(*) FROM Block_Ban_doc l inner join Ban_doc b on l.The_ID=b.ID WHERE UPPER(So_the) = UPPER(@strPatronCode)
	IF @intCount > 0
	BEGIN
		SET @intOutPut = 4
--print(@intOutPut)
		RETURN @intOutPut
	END

	-- Check expired
	SELECT @intCount = COUNT(*) FROM Ban_doc WHERE UPPER(So_the) = UPPER(@strPatronCode) AND Ngay_het_han < GetDate()
	IF @intCount > 0
	BEGIN
		SET @intOutPut = 2
--print(@intOutPut)
		RETURN @intOutPut
	END
	-- Check quota exceed
	IF @intLoanMode = 1
	BEGIN
	/*Tam bo lai
		SELECT @intCount = COUNT(A.PatronID) FROM LoanType A, Ban_doc B WHERE A.LoanTypeID = B.ty AND UPPER(B.CODE) = UPPER(@strPatronCode) AND LoanMode = 1

		IF @intCount >= 0
		BEGIN
			SELECT @intCount = COUNT(*) FROM CIR_PATRON A, CIR_PATRON_GROUP B WHERE A.PatronGroupID = B.ID AND UPPER(A.CODE) = UPPER(@strPatronCode) AND B.LoanQuota <= @intCount
			IF @intCount > 0
			BEGIN
				SET @intOutPut = 3
--print(@intOutPut)
				RETURN @intOutPut
			END
		END
		*/
		SELECT @intCount =1
	END
	ELSE IF @intLoanMode = 2
	BEGIN
	/* tam Bo
		SELECT @intCount = COUNT(A.PatronID) FROM CIR_LOAN A, CIR_PATRON B WHERE A.PatronID = B.ID AND UPPER(B.CODE) = UPPER(@strPatronCode) AND LoanMode = 2
		IF @intCount >= 0
		BEGIN
			SELECT @intCount = COUNT(*) FROM CIR_PATRON A, CIR_PATRON_GROUP B WHERE A.PatronGroupID = B.ID AND UPPER(A.CODE) = UPPER(@strPatronCode) AND B.InLibraryQuota <= @intCount
			IF @intCount > 0
			BEGIN
				SET @intOutPut = 6
--print(@intOutPut)
				RETURN @intOutPut
			END
		END
		*/
		SELECT @intCount =1
	END
	/*
	-- Check permission
	SELECT @intPatronGroupID = A.PatronGroupID FROM CIR_PATRON A, CIR_PATRON_GROUP B WHERE A.PatronGroupID = B.ID AND UPPER(A.CODE) = UPPER(@strPatronCode)
	IF @intPatronGroupID > 0
	BEGIN
		SELECT @intCount = COUNT(*) FROM CIR_PATRON_GROUP_LOC WHERE PatronGroupID = @intPatronGroupID AND LocationID IN (SELECT LocID FROM SYS_USER_LOCATION WHERE UserID = @intuserID)
		IF @intPatronGroupID = 0
		BEGIN
			SET @intOutPut = 5
--print(@intOutPut)
			RETURN @intOutPut
		END
	END
	*/
	RETURN @intOutPut

GO
/****** Object:  StoredProcedure [dbo].[SP_GET_PATRON_INFOR]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[SP_GET_PATRON_INFOR] 
 @strFullName nvarchar(150),      
 @strPatronCode  varchar(30),
 @strFixDueDate varchar(30)      
AS   
	DECLARE @intByte int

SET    @intByte = 0

IF @strFixDueDate <>'' --Kiem tra ngay gia han lon hon ngay het han the
Begin
	SELECT 	@intByte=1 FROM Ban_doc WHERE So_the =  UPPER(@strPatronCode) AND Ngay_het_han < @strFixDueDate
	IF @intByte=0 --Kiem tra ngay gia han nho hon ngay cap the
		SELECT 	@intByte=1 FROM Ban_doc WHERE So_the =  UPPER(@strPatronCode) AND Ngay_cap > @strFixDueDate
	SELECT *, @intByte AS intByte FROM Ban_doc WHERE So_the =  UPPER(@strPatronCode)
End
ELSE
Begin
 	IF @strFullName <> ''
		SET @strFullName = '%' + @strFullName
 
 	SELECT dbo.fnChuyenCoDauThanhKhongDau(dbo.DecodeUTF8String(Ho_ten))  AS FullName      
  	, '' as Ethnic, CP.ngay_sinh as Dob, CP.Gioi_tinh as Sex, '' EducationLevel, CP.So_the as Code, CP.Ngay_cap ValidDate, CONVERT(varchar,Ngay_het_han,103)AS ExpiredDate, '' as  GroupName,  
 	'' as Grade, '' as Faculty, '' as Class, '' as Occupation, NULL as OccupationChinh, '' College      
 	,'' as WorkPlace, Dia_chi as Address, CP.So_dien_thoai as Telephone, CP.Mobile, CP.Email, '' Portrait, 
 	CP.Debt, 0 as LoanQuota, 10 as InLibraryQuota, CP.ID,CP.Mat_khau as password, CP.Ghi_chu as Note
  	FROM  ban_doc CP   
	WHERE  CP.So_the = UPPER(@strPatronCode)   
       
End

GO
/****** Object:  StoredProcedure [dbo].[SP_SYS_USER_LOGIN]    Script Date: 24/03/2021 8:55:29 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[SP_SYS_USER_LOGIN]  
 @strUserName varchar(100),  
 @strPassword varchar(100)
AS
	SELECT * FROM Users
	WHERE Username=@strUserName	-- AND Pass=@strPassword
GO
