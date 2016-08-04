/*##################################################################################################
# Name:                                                                                   2016-07-29
#    ut_zzVBA
#
####################################################################################################
# Description: 
#    List basic VBA code statements
#
####################################################################################################
# Utilities:
EXEC ut_zzANL VAT
EXEC ut_zzANL VAV
EXEC ut_zzANL VAP
EXEC ut_zzANL VAF
##################################################################################################*/
SET QUOTED_IDENTIFIER OFF
SET NOCOUNT           ON
GO
----------------------------------------------------------------------------------------------------
--USE DBGen
GO
----------------------------------------------------------------------------------------------------
IF OBJECT_ID('dbo.ut_zzVBA') IS NOT NULL DROP PROCEDURE dbo.ut_zzVBA  -- (IXS)
GO
----------------------------------------------------------------------------------------------------
CREATE PROCEDURE dbo.ut_zzVBA (    -- (PSB)
    @BldLST varchar(2000) = '',             -- Build code list (comma delimited; see below)
    @InpTxt varchar(max)  = '',             -- Input Object text (comma delimited)
    @StdTx1 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx2 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx3 varchar(max)  = '',             -- Miscellaneous text value
    @StdFlg tinyint       = 0,              -- Miscellaneous flag value
    @StdCnt int           = 0               -- Miscellaneous count value
) AS BEGIN
    ------------------------------------------------------------------------------------------------
    -- Signature Template  (PIF)
    /*----------------------------------------------------------------------------------------------
    --   ut_zzVBA BldTxtTx1Tx2Tx3FlgCnt
    EXEC ut_zzVBA ,'','','','',0,0
    --   ut_zzVBA BldTxtTx1Tx2Tx3FlgCnt
    EXEC ut_zzVBA @BldLST,@InpTxt,@StdTx1,@StdTx2,@StdTx3,@StdFlg,@StdCnt
    ----------------------------------------------------------------------------------------------*/
    -- Set the Environment
    ------------------------------------------------------------------------------------------------
    SET NOCOUNT       ON   -- ON OFF
    SET ANSI_WARNINGS OFF  -- ON OFF
    SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
    ------------------------------------------------------------------------------------------------
    -- Prep Parameters
    ------------------------------------------------------------------------------------------------
    SET @BldLST = UPPER(@BldLST)
    ------------------------------------------------------------------------------------------------
    -- Assign Procedure Profile Values
    ------------------------------------------------------------------------------------------------
    DECLARE @CurUSP    varchar(30)     ; SET @CurUSP = 'ut_zzVBA'
    DECLARE @CurREF    varchar(30)     ; SET @CurREF = 'dbo.ut_zzVBA'
    DECLARE @CurDSC    varchar(100)    ; SET @CurDSC = 'List basic VBA code statements'
    DECLARE @CurCAT    varchar(10)     ; SET @CurCAT = 'GEN'
    DECLARE @CurFMT    char(3)         ; SET @CurFMT = RIGHT(@CurUSP,3)
    ------------------------------------------------------------------------------------------------
    -- Manage Execution Flags
    ------------------------------------------------------------------------------------------------
    DECLARE @CurEXC    tinyint         ; SET @CurEXC = 1                                          -- Execution: 0=Disabled 1=Enabled
    DECLARE @CurDBG    tinyint         ; SET @CurDBG = 0                                          -- DebugMode: 0=Disabled 1=Enabled
    ------------------------------------------------------------------------------------------------
    DECLARE @DbgLvl    varchar(9)      ; SET @DbgLvl = ''                                         -- DebugText: Customize for Debug Tracking
    DECLARE @DbgFlg    tinyint         ; SET @DbgFlg = @CurDBG                                    -- Backward Compatibility; Assign @CurDBG
    ------------------------------------------------------------------------------------------------
    SET @DbgFlg = CASE WHEN @BldLST = 'ZZZ' THEN 1 ELSE @DbgFlg END
    ------------------------------------------------------------------------------------------------
    -- Display text based on Debug/Execution modes
    ------------------------------------------------------------------------------------------------
    IF @CurDBG = 1 OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT
        LEFT(@CurUSP ,15) AS CurUSP,
        LEFT(@BldLST ,30) AS BldLST,
        LEFT(@InpTxt ,30) AS InpTxt,
        LEFT(@StdTx1 ,30) AS StdTx1,
        LEFT(@StdTx2 ,30) AS StdTx2,
        LEFT(@StdTx3 ,30) AS StdTx3,
        @StdFlg           AS StdFlg,
        @StdCnt           AS StdCnt
    ------------------------------------------------------------------------------------------------
    END ELSE IF @CurEXC = 0 OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT ''
        PRINT 'Procedure   :  '+@CurREF
        PRINT 'Description :  '+@CurDSC
        PRINT 'Status      :  Disabled'
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        RETURN
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldLST IN ('', 'h','/h','help','*') OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT ''
        PRINT 'Procedure   :  '+@CurREF
        PRINT 'Description :  '+@CurDSC
        PRINT 'Parameters  :'
        PRINT "    @BldLST varchar(2000) = '',             -- Build code list (comma delimited; see below)"
        PRINT "    @InpTxt varchar(max)  = '',             -- Input Object text (comma delimited)"
        PRINT "    @StdTx1 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @StdTx2 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @StdTx3 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @StdFlg tinyint       = 0,              -- Miscellaneous flag value"
        PRINT "    @StdCnt int           = 0               -- Miscellaneous count value"
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT 'Build Codes:'
        PRINT ''
        PRINT '    LCP    = Lookup Constants: PKey'
        PRINT '    LCC    = Lookup Constants: Code'
        PRINT '    LCN    = Lookup Constants: Name'
        PRINT '    LCX    = Lookup Constants: CmdTxt'
        PRINT '    '
        PRINT '    LPP    = Lookup Properties: PKey'
        PRINT '    LPC    = Lookup Properties: Code'
        PRINT '    LPN    = Lookup Properties: Name'
        PRINT '    LPX    = Lookup Properties: CmdTxt'
        PRINT '    '
        PRINT '    MLV    = Module Level Variables'
        PRINT '    MLP    = Module Level Properties - LET/GET'
        PRINT '    MLL    = Module Level Properties - LET'
        PRINT '    MLG    = Module Level Properties - GET'
        PRINT '    '
        PRINT '    MAV    = Module Level AssignSQL (from variables)'
        PRINT '    MAC    = Module Level AssignSQL (from controls)'
        PRINT '    MAN    = Module Level AssignSQL (from nulls)'
        PRINT '    '
        PRINT '    MRV    = Module Level ReadSQL (into variables)'
        PRINT '    MRC    = Module Level ReadSQL (into controls)'
        PRINT '    '
        PRINT '    MCV    = Module Level ClearSQL (variables)'
        PRINT '    MCC    = Module Level ClearSQL (controls)'
        PRINT '    '
        PRINT '    MIF    = Module IF Criteria'
        PRINT '    '
        PRINT '    DTV    = Declare fields column variables'
        PRINT '    ITV    = Initialize fields column variables'
        PRINT '    ATV    = Assign fields column variables'
        PRINT '    '
        PRINT '    DSV    = Declare statement column variables'
        PRINT '    ISV    = Initialize statement column variables'
        PRINT '    ASV    = Assign statement column variables'
        PRINT '    '
        PRINT '    DKV    = Declare primary key column variables'
        PRINT '    IKV    = Initialize primary key column variables'
        PRINT '    '
        PRINT '    AKV    = Assign primary key column variables'
        PRINT '    DGV    = Declare function parameter variables'
        PRINT '    IGV    = Initialize parameter variables'
        PRINT '    '
        PRINT '    AGV    = Assign parameter variables'
        PRINT '    '
        PRINT '    RSV    = Recordset variables'
        PRINT '    RSL    = Recordset standard loop'
        PRINT '    RSW    = Recordset write loop'
        PRINT '    '
        PRINT '    BASWTL = Build module:  bas_UtlWTL'
        PRINT '    CLSWTL = Build module:  clsUtlWTL'
        PRINT '    '
        PRINT '    BASAPC = Build module:  bas_AppCons'
        PRINT '    BASAPF = Build module:  bas_AppFunc'
        PRINT '    BASAPT = Build module:  bas_AppTest'
        PRINT '    BASAPV = Build module:  bas_AppVars'
        PRINT '    '
        PRINT '    BASGLB = Build module:  bas_Global'
        PRINT '    BASIMX = Build module:  bas_ImpExp'
        PRINT '    BASTST = Build module:  bas_Test01'
        PRINT '    BASTBM = Build module:  bas_TblMnt'
        PRINT '    '
        PRINT '    UTLASC = Build module:  clsUtlASC'
        PRINT '    UTLFMT = Build module:  clsUtlFMT'
        PRINT '    UTLVBG = Build module:  clsUtlVBG'
        PRINT '    UTLWSH = Build module:  clsUtlWSH'
        PRINT '    UTLWTX = Build module:  clsUtlWTX'
        PRINT '    '
        PRINT '    GENGLB = Build module:  vba_Global'
        PRINT '    GENSTD = Build module:  vbaGenSTD'
        PRINT '    GENJET = Build module:  vbaGenJET'
        PRINT '    '
        PRINT '    GEN_IT = Build module:  vbaGen_IT'
        PRINT '    GENFRM = Build module:  vbaGenFRM'
        PRINT '    GENCTL = Build module:  vbaGenCTL'
        PRINT '    GENTBL = Build module:  vbaGenTBL'
        PRINT '    GENPRP = Build module:  vbaGenPRP'
        PRINT '    GENCMD = Build module:  vbaGenCMD'
        PRINT '    GENRPT = Build module:  vbaGenRPT'
        PRINT '    GENPTH = Build module:  vbaGenPTH'
        PRINT '    GENSQL = Build module:  vbaGenSQL'
        PRINT '    GENSBY = Build module:  vbaGenSBY'
        PRINT '    GENGBY = Build module:  vbaGenGBY'
        PRINT '    GENSLO = Build module:  vbaGenSLO'
        PRINT '    '
        PRINT '    CLSAPC = Build module:  clsAppCons'
        PRINT '    CLSAPV = Build module:  clsAppVals'
        PRINT '    '
        PRINT '    BASCMG = Build module:  bas_CmgCons'
        PRINT '    CLSCMG = Build module:  clsCtlMgr'
        PRINT '    '
        PRINT '    REGTBL = Build module:  clsRegTBL'
        PRINT '    REGPRP = Build module:  clsRegPRP'
        PRINT '    REGCMD = Build module:  clsRegCMD'
        PRINT '    REGRPT = Build module:  clsRegRPT'
        PRINT '    REGPTH = Build module:  clsRegPTH'
        PRINT '    REGSRC = Build module:  clsRegSRC'
        PRINT '    '
        PRINT '    SQLSTM = Build module:  clsSqlSTM'
        PRINT '    SQLOBY = Build module:  clsSqlOBY'
        PRINT '    RUNWHR = Build module:  clsRunWHR'
        PRINT '    '
        PRINT '    RUNCMD = Build module:  clsRunCMD'
        PRINT '    RUNCMM = Build module:  Run_Process_0000 (CALL cls_Method)'
        PRINT '    RUNCMF = Build module:  Run_Process_0000 (OPEN frm_FrmNam)'
        PRINT '    '
        PRINT '    RUNRPT = Build module:  clsRunRPT'
        PRINT '    RUNRPR = Build module:  Run_Report_0000'
        PRINT '    RUNRPX = Build module:  Print_rpt_ReportName'
        PRINT '    '
        PRINT '    RUNUSP = Build module:  clsRunUSP'
        PRINT '    RUNUSR = Build module:  Run_Process_0000 (EXEC PROC)'
        PRINT '    RUNUSF = Build module:  Run_Process_0000 (OPEN FORM)'
        PRINT '    RUNUSX = Build SProcs:  Execute_usp_ProcedureName'
        PRINT '    '
        PRINT '    RUNRST = Build module:  clsRunRST'
        PRINT '    RUNSQL = Build module:  clsRunSQL'
        PRINT '    RUNSBY = Build module:  clsRunSBY'
        PRINT '    RUNGBY = Build module:  clsRunGBY'
        PRINT '    '
        PRINT '    FRMCLR = Build module:  frm_FrmName'
        PRINT '    FRMLNK = Build module:  sys_LinkAPP'
        PRINT '    '
        PRINT '    RPTNAR = Build module:  tpl_NARROW'
        PRINT '    RPTWID = Build module:  tpl_WIDE'
        PRINT '    '
        PRINT '    ANYFRM = Build module:  frm_FrmName'
        PRINT '    ANYTAB = Build module:  frm_FrmName'
        PRINT '    ANYLST = Build module:  lst_FrmName'
        PRINT '    ANYPOP = Build module:  pop_FrmName'
        PRINT '    ANYSUB = Build module:  sub_FrmName'
        PRINT '    ANYBAS = Build module:  basBasName'
        PRINT '    ANYCLS = Build module:  clsClsName'
        PRINT '    ANYRPT = Build module:  rpt_RptNam'
        PRINT '    '
        PRINT '    CLSTCN = Build module:  clsTxtCon'
        PRINT '    '
        PRINT '    XTDXYR = Extend Tax Year'
        PRINT '    XTDXPD = Extend Tax Period'
        PRINT '    XTDXMN = Extend Tax Month'
        PRINT '    XTDXAY = Extend Active Tax Year'
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT 'Example Code:'
        PRINT ''
        PRINT '    --   ut_zzVBA Bld    Obj         Tx1 Tx2 Tx3'
        PRINT '    EXEC ut_zzVBA BldCod,InputObject,'''''''' ,'''''''' ,'''''''''
        PRINT '    '
        PRINT '    --   ut_zzVBA Bld     Obj     Tx1     Tx2     Tx3'
        PRINT '    EXEC ut_zzVBA @BldLST,@InpTxt,@StdTx1,@StdTx2,@StdTx3'
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        RETURN
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldLST IN ('?','?N','?Y') OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT CASE @BldLST WHEN '?N' THEN '*' WHEN '?Y' THEN ' ' ELSE '' END+@CurUSP+' ('+@CurCAT+') = '+@CurDSC; RETURN
    END
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Debug flags
    ------------------------------------------------------------------------------------------------
    DECLARE @DbgCon    smallint        ; SET @DbgCon    = 0                                       -- Debug tracking
    DECLARE @DbgInd    smallint        ; SET @DbgInd    = 0                                       -- Debug tracking
    ------------------------------------------------------------------------------------------------
    -- Build codes formatting
    ------------------------------------------------------------------------------------------------
    SET @BldLST = UPPER(@BldLST)                                                                  -- Build list
    ------------------------------------------------------------------------------------------------
    -- Build codes tracking
    ------------------------------------------------------------------------------------------------
    DECLARE @BldCOD    varchar(100)    ; SET @BldCOD    = @BldLST                                 -- Build code
    DECLARE @BldVAL    varchar(100)    ; SET @BldVAL    = ''                                      -- Build value (=Xxx)
    DECLARE @BldFlg    bit             ; SET @BldFlg    = 0                                       -- Build flag
    DECLARE @PrvBld    varchar(100)    ; SET @PrvBld    = ''                                      -- Previous build code
    DECLARE @PrvBlv    varchar(100)    ; SET @PrvBlv    = ''                                      -- Previous build value
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize Standard default objects  (ID1)
    ------------------------------------------------------------------------------------------------
    DECLARE @InpObj    sysname         ; SET @InpObj    = @InpTxt                                 -- Input object name
    DECLARE @NamFmt    varchar(3)      ; SET @NamFmt    = ''                                      -- Name format (revised from @BldCOD)
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize Output default objects  (ID2)
    ------------------------------------------------------------------------------------------------
    DECLARE @OupObj    sysname         ; SET @OupObj    = ''                                      -- Output object (what is created)
    DECLARE @OupDsc    sysname         ; SET @OupDsc    = ''                                      -- Output description
    DECLARE @OvrOup    tinyint         ; SET @OvrOup    = 0                                       -- Output override flag
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize ut_zzVBA default objects  (ID3)
    ------------------------------------------------------------------------------------------------
    DECLARE @OupCls    varchar(11)     ; SET @OupCls    = ''                                      -- Output class  (TBL,VEW,etc)
    DECLARE @OupFmt    varchar(11)     ; SET @OupFmt    = ''                                      -- Output format (SEL,INS,etc)
    DECLARE @OupAls    varchar(11)     ; SET @OupAls    = ''                                      -- Output table alias (abc,xyz,...)
    DECLARE @OupSfx    varchar(11)     ; SET @OupSfx    = ''                                      -- Output table suffix (ABC,XYZ,...)
    DECLARE @OupCpx    varchar(20)     ; SET @OupCpx    = ''                                      -- Output column prefix (Abcdef,...)
    DECLARE @OupUtl    varchar(11)     ; SET @OupUtl    = RIGHT(@CurUSP,3)                        -- Output utility (TBX,USX,etc)
    DECLARE @SqlExc    tinyint         ; SET @SqlExc    = 0                                       -- Execute the dynamic SQL statement
    DECLARE @LftMrg    smallint        ; SET @LftMrg    = 0                                       -- Increase left margin (4x)
    DECLARE @IncSpc    tinyint         ; SET @IncSpc    = 0                                       -- Include space(s) before the header
    DECLARE @IncTtl    tinyint         ; SET @IncTtl    = 1                                       -- Include code segment titles
    DECLARE @IncCmt    tinyint         ; SET @IncCmt    = 0                                       -- Include comment text
    DECLARE @IncHdr    tinyint         ; SET @IncHdr    = 0                                       -- Include header lines/text
    DECLARE @IncTpl    tinyint         ; SET @IncTpl    = 0                                       -- Include templates
    DECLARE @IncDbg    tinyint         ; SET @IncDbg    = 0                                       -- Include debug logic
    DECLARE @IncMsg    tinyint         ; SET @IncMsg    = 0                                       -- Include information message
    DECLARE @IncErm    tinyint         ; SET @IncErm    = 0                                       -- Include error message
    DECLARE @IncSep    tinyint         ; SET @IncSep    = 1                                       -- Include separator line between objects
    DECLARE @IncDrp    tinyint         ; SET @IncDrp    = 0                                       -- Include drop statement
    DECLARE @IncAdd    tinyint         ; SET @IncAdd    = 0                                       -- Include add statement
    DECLARE @IncBat    tinyint         ; SET @IncBat    = 1                                       -- Include batch GO statement
    DECLARE @IncTcd    tinyint         ; SET @IncTcd    = 0                                       -- Include test code
    DECLARE @IncDat    tinyint         ; SET @IncDat    = NULL                                    -- Include data insert statements
    DECLARE @IncPrm    tinyint         ; SET @IncPrm    = 0                                       -- Include permissions statements
    DECLARE @SelStm    varchar(100)    ; SET @SelStm    = ''                                      -- SELECT statement (DISTINCT, TOP, etc)
    DECLARE @SetLst    varchar(2000)   ; SET @SetLst    = ''                                      -- SET Column = Value list (colon delimited)
    DECLARE @JnnLst    varchar(2000)   ; SET @JnnLst    = ''                                      -- JOIN list (colon delimited)
    DECLARE @WhrLst    varchar(2000)   ; SET @WhrLst    = ''                                      -- WHERE list (colon delimited)
    DECLARE @GbyLst    varchar(2000)   ; SET @GbyLst    = ''                                      -- GROUP BY list (colon delimited)
    DECLARE @HavLst    varchar(2000)   ; SET @HavLst    = ''                                      -- HAVING list (colon delimited)
    DECLARE @ObyLst    varchar(2000)   ; SET @ObyLst    = ''                                      -- ORDER BY list (comma delimited)
    DECLARE @LkpLst    varchar(2000)   ; SET @LkpLst    = ''                                      -- Lookup parameters (comma delimited list)
    DECLARE @IncTrn    tinyint         ; SET @IncTrn    = NULL                                    -- Include transaction logic
    DECLARE @IncIdn    tinyint         ; SET @IncIdn    = NULL                                    -- Include identity column logic
    DECLARE @IncHtk    tinyint         ; SET @IncHtk    = NULL                                    -- Include record history tracking columns
    DECLARE @IncDim    tinyint         ; SET @IncDim    = NULL                                    -- Include record dimension columns
    DECLARE @IncFct    tinyint         ; SET @IncFct    = NULL                                    -- Include record fact columns
    DECLARE @IncUsd    tinyint         ; SET @IncUsd    = NULL                                    -- Include record used column
    DECLARE @IncLkd    tinyint         ; SET @IncLkd    = NULL                                    -- Include record locked column
    DECLARE @IncDsb    tinyint         ; SET @IncDsb    = NULL                                    -- Include record disabled column
    DECLARE @IncDlt    tinyint         ; SET @IncDlt    = NULL                                    -- Include record delflag column
    DECLARE @IncLok    tinyint         ; SET @IncLok    = NULL                                    -- Include record locking columns
    DECLARE @IncCrt    tinyint         ; SET @IncCrt    = NULL                                    -- Include record created columns
    DECLARE @IncUpd    tinyint         ; SET @IncUpd    = NULL                                    -- Include record updated columns
    DECLARE @IncExp    tinyint         ; SET @IncExp    = NULL                                    -- Include record expired columns
    DECLARE @IncDel    tinyint         ; SET @IncDel    = NULL                                    -- Include record deleted columns
    DECLARE @IncAud    tinyint         ; SET @IncAud    = NULL                                    -- Include record auditing columns
    DECLARE @IncHst    tinyint         ; SET @IncHst    = NULL                                    -- Include record history columns
    DECLARE @IncMod    tinyint         ; SET @IncMod    = NULL                                    -- Include record modified columns
    DECLARE @BldSfx    varchar(11)     ; SET @BldSfx    = RIGHT(@OupFmt,3)                        -- Output suffix
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize Margin Values  (AMV)                                         EXEC ut_zzUTL zzz,AMV
    ------------------------------------------------------------------------------------------------
    DECLARE @MrgInc    tinyint         ; SET @MrgInc    = 4                                       -- Left margin space increment
    DECLARE @StdLen    tinyint         ; SET @StdLen    = 100                                     -- Standard line length
    ------------------------------------------------------------------------------------------------
    DECLARE @LftWid    tinyint         ; SET @LftWid    = @LftMrg * @MrgInc                       -- Left margin space length Beg
    DECLARE @LftLen    tinyint         ; SET @LftLen    = @StdLen - @LftWid                       -- Left length
    ------------------------------------------------------------------------------------------------
    DECLARE @StmMrg    smallint        ; SET @StmMrg    = @LftMrg+1                             -- Code statement margin
    DECLARE @StmWid    smallint        ; SET @StmWid    = @StmMrg * @MrgInc                       -- Statement margin space length
    DECLARE @StmLen    smallint        ; SET @StmLen    = @StdLen - @StmWid                       -- Statement line length
    ------------------------------------------------------------------------------------------------
    DECLARE @M         varchar(50)     ; SET @M         = REPLICATE(' ', @LftWid)                 -- Left margin
    DECLARE @T         varchar(50)     ; SET @T         = REPLICATE(' ', @StmWid)                 -- Statement margin
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Standard Working Constants  (SWC)                                       EXEC ut_zzUTL zzz,SWC
    ------------------------------------------------------------------------------------------------
    DECLARE @NOP       bit             ; SET @NOP       = 0                                       -- Standard Flag = False
    DECLARE @YUP       bit             ; SET @YUP       = 1                                       -- Standard Flag = True
    DECLARE @False     bit             ; SET @False     = 0                                       -- Standard Flag = False
    DECLARE @True      bit             ; SET @True      = 1                                       -- Standard Flag = True
    ------------------------------------------------------------------------------------------------
    DECLARE @CRT       char(1)         ; SET @CRT       = CHAR(13)                                -- Standard Character: CarrRtn
    DECLARE @LFD       char(1)         ; SET @LFD       = CHAR(10)                                -- Standard Character: LineFeed
    DECLARE @NLN       char(2)         ; SET @NLN       = @CRT+@LFD                               -- Standard Character: Newline
    DECLARE @SPC       char(1)         ; SET @SPC       = ' '                                     -- Standard Character: Space
    DECLARE @MTY       varchar(1)      ; SET @MTY       = ''                                      -- Standard Character: Empty
    ------------------------------------------------------------------------------------------------
    DECLARE @DOT       char(1)         ; SET @DOT       = '.'                                     -- Standard Character: Period
    DECLARE @CMA       char(1)         ; SET @CMA       = ','                                     -- Standard Character: Comma
    DECLARE @CLN       char(1)         ; SET @CLN       = ':'                                     -- Standard Character: Colon
    DECLARE @SCN       char(1)         ; SET @SCN       = ';'                                     -- Standard Character: SemiColon
    DECLARE @UBR       char(1)         ; SET @UBR       = '_'                                     -- Standard Character: Underbar
    DECLARE @PCT       char(1)         ; SET @PCT       = '%'                                     -- Standard Character: Percent
    ------------------------------------------------------------------------------------------------
    DECLARE @SGL       char(1)         ; SET @SGL       = '-'                                     -- Standard Character: Single
    DECLARE @DBL       char(1)         ; SET @DBL       = '='                                     -- Standard Character: Double
    DECLARE @AST       char(1)         ; SET @AST       = '*'                                     -- Standard Character: Asterisk
    DECLARE @PND       char(1)         ; SET @PND       = '#'                                     -- Standard Character: Pound
    DECLARE @ATS       char(1)         ; SET @ATS       = '@'                                     -- Standard Character: AtSign
    DECLARE @TLD       char(1)         ; SET @TLD       = '~'                                     -- Standard Character: Tilde
    DECLARE @BNG       char(1)         ; SET @BNG       = '!'                                     -- Standard Character: Bang
    DECLARE @VBR       char(1)         ; SET @VBR       = '|'                                     -- Standard Character: VertBar
    ------------------------------------------------------------------------------------------------
    DECLARE @NQT       varchar(1)      ; SET @NQT       = ''                                      -- Standard Character: Quote Empty
    DECLARE @SQT       char(1)         ; SET @SQT       = "'"                                     -- Standard Character: Quote Single
    DECLARE @DQT       char(1)         ; SET @DQT       = '"'                                     -- Standard Character: Quote Double
    DECLARE @BTK       char(1)         ; SET @BTK       = '`'                                     -- Standard Character: Quote BackTick
    ------------------------------------------------------------------------------------------------
    DECLARE @AND       char(4)         ; SET @AND       = 'AND '                                  -- Standard Statement: And
    DECLARE @ONN       char(4)         ; SET @ONN       = ' ON '                                  -- Standard Statement: On
    ------------------------------------------------------------------------------------------------
    DECLARE @WUB       char(1)         ; SET @WUB       = '_'                                     -- WildCard: Underbar (single char)
    DECLARE @WPC       char(1)         ; SET @WPC       = '%'                                     -- WildCard: Percent  (multi chars)
    DECLARE @WBG       char(1)         ; SET @WBG       = '!'                                     -- WildCard: Bang     (escape char)
    ------------------------------------------------------------------------------------------------
    DECLARE @ITX       char(500)       ; SET @ITX       = ''                                      -- Standard String: Pad with Spaces
    DECLARE @ITZ       char(100)       ; SET @ITZ       = REPLICATE('Z',100)                      -- Standard String: Pad with Zzzzs
    DECLARE @IT0       char(100)       ; SET @IT0       = REPLICATE('0',100)                      -- Standard String: Pad with Zeros
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Standard Working Variables  (SWV)                                       EXEC ut_zzUTL zzz,SWV
    ------------------------------------------------------------------------------------------------
    DECLARE @CNT       int             ; SET @CNT       = 0                                       -- Working CountValue
    DECLARE @CNX       varchar(10)     ; SET @CNX       = ''                                      -- Working CountText
    ------------------------------------------------------------------------------------------------
    DECLARE @IDN       int             ; SET @IDN       = 0                                       -- Working IndexValue
    DECLARE @IDX       varchar(10)     ; SET @IDX       = ''                                      -- Working IndexText
    ------------------------------------------------------------------------------------------------
    DECLARE @FLG       bit             ; SET @FLG       = 0                                       -- Working FlagValue
    DECLARE @FLX       varchar(10)     ; SET @FLX       = ''                                      -- Working FlagText
    ------------------------------------------------------------------------------------------------
    DECLARE @LVL       int             ; SET @LVL       = 0                                       -- Working Level
    DECLARE @SIZ       dec(15,2)       ; SET @SIZ       = 0                                       -- Working Size
    DECLARE @SQX       varchar(10)     ; SET @SQX       = ''                                      -- Working SequenceText
    DECLARE @QOT       varchar(1)      ; SET @QOT       = ''                                      -- Working QuoteMark
    DECLARE @NUL       varchar(1)      ; SET @NUL       = ''                                      -- Working Null Output
    ------------------------------------------------------------------------------------------------
    DECLARE @MIN       int             ; SET @MIN       = 0                                       -- Working Minimum
    DECLARE @MAX       int             ; SET @MAX       = 0                                       -- Working Maximum
    ------------------------------------------------------------------------------------------------
    DECLARE @LSP       varchar(20)     ; SET @LSP       = ''                                      -- Working Line Space
    ------------------------------------------------------------------------------------------------
    DECLARE @VAL       varchar(200)    ; SET @VAL       = ''                                      -- Working General Value
    DECLARE @PRV       varchar(200)    ; SET @PRV       = ''                                      -- Working Previous Value
    DECLARE @CUR       varchar(200)    ; SET @CUR       = ''                                      -- Working Current Value
    ------------------------------------------------------------------------------------------------
    DECLARE @BLD       varchar(200)    ; SET @BLD       = ''                                      -- Working Build
    DECLARE @TYP       varchar(200)    ; SET @TYP       = ''                                      -- Working Type
    DECLARE @PTN       varchar(200)    ; SET @PTN       = ''                                      -- Working Pattern
    DECLARE @UNK       varchar(200)    ; SET @UNK       = ''                                      -- Working Unknown
    DECLARE @ZZZ       varchar(200)    ; SET @ZZZ       = ''                                      -- Working Placeholder
    ------------------------------------------------------------------------------------------------
    DECLARE @TST       bit             ; SET @TST       = 0                                       -- Working Flag
    DECLARE @RUN       bit             ; SET @RUN       = 0                                       -- Working Flag
    ------------------------------------------------------------------------------------------------
    DECLARE @TXT       varchar(max)    ; SET @TXT       = ''                                      -- Working Text
    DECLARE @LEN       int             ; SET @LEN       = 0                                       -- Working Length (numeric)
    DECLARE @LEX       varchar(10)     ; SET @LEX       = ''                                      -- Working Length (text)
    ------------------------------------------------------------------------------------------------
    DECLARE @CMX       varchar(2)      ; SET @CMX       = ''                                      -- Working SQL Comma
    DECLARE @ANX       varchar(40)     ; SET @ANX       = ''                                      -- Working SQL AND
    DECLARE @ONX       varchar(40)     ; SET @ONX       = ''                                      -- Working SQL ON
    ------------------------------------------------------------------------------------------------
    DECLARE @OPN       varchar(11)     ; SET @OPN       = ''                                      -- Working Paren: Open
    DECLARE @CPN       varchar(11)     ; SET @CPN       = ''                                      -- Working Paren: Close
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Extended Working Variables  (XWV)                                       EXEC ut_zzUTL zzz,XWV
    ------------------------------------------------------------------------------------------------
    DECLARE @CN1       int             ; SET @CN1       = 0                                       -- Working Count
    DECLARE @CN2       int             ; SET @CN2       = 0                                       -- Working Count
    DECLARE @CN3       int             ; SET @CN3       = 0                                       -- Working Count
    ------------------------------------------------------------------------------------------------
    DECLARE @ID1       int             ; SET @ID1       = 0                                       -- Working Index
    DECLARE @ID2       int             ; SET @ID2       = 0                                       -- Working Index
    DECLARE @ID3       int             ; SET @ID3       = 0                                       -- Working Index
    ------------------------------------------------------------------------------------------------
    DECLARE @PS1       int             ; SET @PS1       = 0                                       -- Working Position
    DECLARE @PS2       int             ; SET @PS2       = 0                                       -- Working Position
    DECLARE @PS3       int             ; SET @PS3       = 0                                       -- Working Position
    ------------------------------------------------------------------------------------------------
    DECLARE @TX0       varchar(max)    ; SET @TX0       = ''                                      -- Working text
    DECLARE @TX1       varchar(max)    ; SET @TX1       = ''                                      -- Working text
    DECLARE @TX2       varchar(max)    ; SET @TX2       = ''                                      -- Working text
    DECLARE @TX3       varchar(max)    ; SET @TX3       = ''                                      -- Working text
    DECLARE @TX4       varchar(max)    ; SET @TX4       = ''                                      -- Working text
    DECLARE @TX5       varchar(max)    ; SET @TX5       = ''                                      -- Working text
    DECLARE @TX6       varchar(max)    ; SET @TX6       = ''                                      -- Working text
    DECLARE @TX7       varchar(max)    ; SET @TX7       = ''                                      -- Working text
    DECLARE @TX8       varchar(max)    ; SET @TX8       = ''                                      -- Working text
    DECLARE @TX9       varchar(max)    ; SET @TX9       = ''                                      -- Working text
    ------------------------------------------------------------------------------------------------
    DECLARE @LN0       int             ; SET @LN0       = 0                                       -- Working length
    DECLARE @LN1       int             ; SET @LN1       = 0                                       -- Working length
    DECLARE @LN2       int             ; SET @LN2       = 0                                       -- Working length
    DECLARE @LN3       int             ; SET @LN3       = 0                                       -- Working length
    DECLARE @LN4       int             ; SET @LN4       = 0                                       -- Working length
    DECLARE @LN5       int             ; SET @LN5       = 0                                       -- Working length
    DECLARE @LN6       int             ; SET @LN6       = 0                                       -- Working length
    ------------------------------------------------------------------------------------------------
    DECLARE @FG0       bit             ; SET @FG0       = 0                                       -- Working flag
    DECLARE @FG1       bit             ; SET @FG1       = 0                                       -- Working flag
    DECLARE @FG2       bit             ; SET @FG2       = 0                                       -- Working flag
    DECLARE @FG3       bit             ; SET @FG3       = 0                                       -- Working flag
    DECLARE @FG4       bit             ; SET @FG4       = 0                                       -- Working flag
    DECLARE @FG5       bit             ; SET @FG5       = 0                                       -- Working flag
    DECLARE @FG6       bit             ; SET @FG6       = 0                                       -- Working flag
    ------------------------------------------------------------------------------------------------
    DECLARE @SPX       varchar(20)     ; SET @SPX       = ''                                      -- Working  Space
    DECLARE @SP0       varchar(1)      ; SET @SP0       = ''                                      -- Constant Space 00
    DECLARE @SP1       char(1)         ; SET @SP1       = ''                                      -- Constant Space 01
    DECLARE @SP2       char(2)         ; SET @SP2       = ''                                      -- Constant Space 02
    DECLARE @SP3       char(3)         ; SET @SP3       = ''                                      -- Constant Space 03
    DECLARE @SP4       char(4)         ; SET @SP4       = ''                                      -- Constant Space 04
    ------------------------------------------------------------------------------------------------
    DECLARE @MRG       tinyint         ; SET @MRG       = 0                                       -- Working  Margin Increment
    DECLARE @MG0       tinyint         ; SET @MG0       = 0                                       -- Constant Margin Increment 00
    DECLARE @MG1       tinyint         ; SET @MG1       = 1                                       -- Constant Margin Increment 01
    DECLARE @MG2       tinyint         ; SET @MG2       = 2                                       -- Constant Margin Increment 02
    DECLARE @MG3       tinyint         ; SET @MG3       = 3                                       -- Constant Margin Increment 03
    DECLARE @MG4       tinyint         ; SET @MG4       = 4                                       -- Constant Margin Increment 04
    DECLARE @MG5       tinyint         ; SET @MG5       = 5                                       -- Constant Margin Increment 05
    ------------------------------------------------------------------------------------------------
    DECLARE @MWD       tinyint         ; SET @MWD       = 0                                       -- Working  Margin Width
    DECLARE @MW0       tinyint         ; SET @MW0       = 00                                      -- Constant Margin Width 00
    DECLARE @MW1       tinyint         ; SET @MW1       = 04                                      -- Constant Margin Width 01
    DECLARE @MW2       tinyint         ; SET @MW2       = 08                                      -- Constant Margin Width 02
    DECLARE @MW3       tinyint         ; SET @MW3       = 12                                      -- Constant Margin Width 03
    DECLARE @MW4       tinyint         ; SET @MW4       = 16                                      -- Constant Margin Width 04
    DECLARE @MW5       tinyint         ; SET @MW5       = 20                                      -- Constant Margin Width 05
    ------------------------------------------------------------------------------------------------
    DECLARE @MRX       varchar(20)     ; SET @MRX       = ''                                      -- Working  Margin Space
    DECLARE @MXX       varchar(20)     ; SET @MXX       = ''                                      -- Working  Margin Space
    DECLARE @MX0       varchar(1)      ; SET @MX0       = ''                                      -- Constant Margin Space 00
    DECLARE @MX1       char(04)        ; SET @MX1       = ''                                      -- Constant Margin Space 01
    DECLARE @MX2       char(08)        ; SET @MX2       = ''                                      -- Constant Margin Space 02
    DECLARE @MX3       char(12)        ; SET @MX3       = ''                                      -- Constant Margin Space 03
    DECLARE @MX4       char(16)        ; SET @MX4       = ''                                      -- Constant Margin Space 04
    DECLARE @MX5       char(20)        ; SET @MX5       = ''                                      -- Constant Margin Space 05
    ------------------------------------------------------------------------------------------------
    DECLARE @LXX       varchar(500)    ; SET @LXX       = ''                                      -- Working  Line
    DECLARE @LX0       char(100)       ; SET @LX0       = @MX0+REPLICATE('-',100)                 -- Constant Line 00
    DECLARE @LX1       char(100)       ; SET @LX1       = @MX1+REPLICATE('-',096)                 -- Constant Line 01
    DECLARE @LX2       char(100)       ; SET @LX2       = @MX2+REPLICATE('-',092)                 -- Constant Line 02
    DECLARE @LX3       char(100)       ; SET @LX3       = @MX3+REPLICATE('-',088)                 -- Constant Line 03
    DECLARE @LX4       char(100)       ; SET @LX4       = @MX4+REPLICATE('-',084)                 -- Constant Line 04
    DECLARE @LX5       char(100)       ; SET @LX5       = @MX5+REPLICATE('-',080)                 -- Constant Line 05
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Extended Construction Variables (XCV)                                   EXEC ut_zzUTL zzz,XCV
    ------------------------------------------------------------------------------------------------
    -- Manage Variable Declarations
    ------------------------------------------------------------------------------------------------
    DECLARE @VAR       varchar(100)    ; SET @VAR       = ''                                      -- Working Variable
    DECLARE @VAX       char(9)         ; SET @VAX       = ''                                      -- Working Variable
    DECLARE @DTP       char(16)        ; SET @DTP       = ''                                      -- Working DataType
    DECLARE @STM       varchar(max)    ; SET @STM       = ''                                      -- Working Statement
    DECLARE @CMT       varchar(200)    ; SET @CMT       = ''                                      -- Working Comment
    DECLARE @VER       varchar(20)     ; SET @VER       = ''                                      -- Working VersionText
    ------------------------------------------------------------------------------------------------
    -- Manage Objects
    ------------------------------------------------------------------------------------------------
    DECLARE @SRV       varchar(100)    ; SET @SRV       = ''                                      -- Working Server
    DECLARE @DBS       varchar(100)    ; SET @DBS       = ''                                      -- Working Database
    DECLARE @SCM       varchar(100)    ; SET @SCM       = ''                                      -- Working Schema
    DECLARE @OBJ       varchar(100)    ; SET @OBJ       = ''                                      -- Working Object
    DECLARE @REF       varchar(100)    ; SET @REF       = ''                                      -- Working scm.obj
    DECLARE @FQD       varchar(100)    ; SET @FQD       = ''                                      -- Working dbs.scm.obj
    DECLARE @FQS       varchar(100)    ; SET @FQS       = ''                                      -- Working srv.dbs.scm.obj
    ------------------------------------------------------------------------------------------------
    DECLARE @ALS       varchar(10)     ; SET @ALS       = ''                                      -- Working Alias
    DECLARE @ALD       varchar(10)     ; SET @ALD       = ''                                      -- Working als.
    DECLARE @ALN       varchar(10)     ; SET @ALN       = ''                                      -- Working spc als
    ------------------------------------------------------------------------------------------------
    DECLARE @TBL       varchar(100)    ; SET @OBJ       = ''                                      -- Working Table
    DECLARE @VEW       varchar(100)    ; SET @VEW       = ''                                      -- Working View
    DECLARE @USP       varchar(100)    ; SET @USP       = ''                                      -- Working SProc
    DECLARE @UFN       varchar(100)    ; SET @UFN       = ''                                      -- Working Function
    DECLARE @EXC       varchar(100)    ; SET @EXC       = ''                                      -- Working Execute
    ------------------------------------------------------------------------------------------------
    -- Manage Columns
    ------------------------------------------------------------------------------------------------
    DECLARE @CLM       varchar(100)    ; SET @CLM       = ''                                      -- Working Column
    DECLARE @ALM       varchar(100)    ; SET @ALM       = ''                                      -- Working als.clm
    DECLARE @NLX       char(9)         ; SET @NLX       = ''                                      -- Working Nullable
    DECLARE @IDT       varchar(9)      ; SET @IDT       = ''                                      -- Working Identity
    ------------------------------------------------------------------------------------------------
    -- Manage KeyIDs
    ------------------------------------------------------------------------------------------------
    DECLARE @KEY       varchar(100)    ; SET @KEY       = ''                                      -- Working KeyIDColumn
    DECLARE @KID       int             ; SET @KID       = 0                                       -- Working KeyIDValue
    DECLARE @KIX       varchar(100)    ; SET @KIX       = ''                                      -- Working KeyIDText
    ------------------------------------------------------------------------------------------------
    -- Manage Signatures
    ------------------------------------------------------------------------------------------------
    DECLARE @SIG       varchar(max)    ; SET @SIG       = ''                                      -- Working Signature
    DECLARE @MOD       varchar(100)    ; SET @MOD       = ''                                      -- Working Module
    DECLARE @TSK       varchar(100)    ; SET @TSK       = ''                                      -- Working Task
    DECLARE @PRM       varchar(100)    ; SET @PRM       = ''                                      -- Working Parameters
    DECLARE @PRX       varchar(100)    ; SET @PRX       = ''                                      -- Working ParamsText
    DECLARE @DBG       varchar(100)    ; SET @DBG       = ''                                      -- Working DebugFlags
    --CLARE @NAM       varchar(100)    ; SET @NAM       = ''                                      -- Working ProcessName
    DECLARE @NAS       varchar(100)    ; SET @NAS       = ''                                      -- Working ProcessName (plural)
    ------------------------------------------------------------------------------------------------
    -- Manage Paramaters
    ------------------------------------------------------------------------------------------------
    DECLARE @OUP       varchar(20)     ; SET @OUP       = ''                                      -- Working Output
    ------------------------------------------------------------------------------------------------
    DECLARE @PFX       varchar(20)     ; SET @PFX       = ''                                      -- Working Prefix
    DECLARE @SFX       varchar(20)     ; SET @SFX       = ''                                      -- Working Suffix
    DECLARE @COD       varchar(100)    ; SET @COD       = ''                                      -- Working Code
    DECLARE @SYS       varchar(100)    ; SET @SYS       = ''                                      -- Working System
    DECLARE @BAS       varchar(100)    ; SET @BAS       = ''                                      -- Working Base
    DECLARE @NAM       varchar(100)    ; SET @NAM       = ''                                      -- Working Name
    DECLARE @TTL       varchar(100)    ; SET @TTL       = ''                                      -- Working Title
    DECLARE @TTX       varchar(200)    ; SET @TTX       = ''                                      -- Working TitleText
    DECLARE @DSC       varchar(200)    ; SET @DSC       = ''                                      -- Working Description
    ------------------------------------------------------------------------------------------------
    -- Manage Lists
    ------------------------------------------------------------------------------------------------
    DECLARE @SEP       varchar(10)     ; SET @SEP       = @CMA                                    -- Working Separator
    DECLARE @DLM       varchar(10)     ; SET @DLM       = ','                                     -- Working Delimiter
    DECLARE @POS       smallint        ; SET @POS       = 0                                       -- Working Position
    DECLARE @LST       varchar(max)    ; SET @LST       = ''                                      -- Working List
    DECLARE @ITM       varchar(500)    ; SET @ITM       = ''                                      -- Working Item (variable length)
    ------------------------------------------------------------------------------------------------
    -- Manage Concatenation
    ------------------------------------------------------------------------------------------------
    DECLARE @N         varchar(2)      ; SET @N         = CHAR(13)+CHAR(10)                       -- Working newline characters
    DECLARE @D         varchar(9)      ; SET @D         = @N                                      -- Working delimiter
    DECLARE @X         varchar(max)    ; SET @X         = ''                                      -- Working dynamic SQL text
    DECLARE @Z         varchar(max)    ; SET @Z         = ''                                      -- Working dynamic SQL text
    DECLARE @S         varchar(2)      ; SET @S         = ' '                                     -- Working dynamic SQL space
    DECLARE @B         varchar(2)      ; SET @B         = ''                                      -- Working dynamic SQL blank
    /*----------------------------------------------------------------------------------------------
    SET @X = @B+@B+""                                                                             -- Firstline initialized with blank
    SET @X = @X+@N+""                                                                             -- Next lines accumulate the text
    PRINT @X                                                                                      -- Print   the text (8000 max)
    EXEC (@X)                                                                                     -- Execute the text (unlimited)
    ----------------------------------------------------------------------------------------------*/
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize Display Size Variables  (DSV)                                EXEC ut_zzUTL zzz,DSV
    ------------------------------------------------------------------------------------------------
    -- Notepad Setup: Font=Fixedsys 8pt; Margins=.5x.5x.5x.5;  Footer='&f   Page &p'
    ------------------------------------------------------------------------------------------------
    DECLARE @811WidPOR smallint        ; SET @811WidPOR = 100                                    -- 08.5 x 11.0 Protrait
    DECLARE @811WidLND smallint        ; SET @811WidLND = 149                                    -- 08.5 x 11.0 Landscape
    ------------------------------------------------------------------------------------------------
    DECLARE @811HgtPOR smallint        ; SET @811HgtPOR =  94                                    -- 08.5 x 11.0 Protrait
    DECLARE @811HgtLND smallint        ; SET @811HgtLND =  63                                    -- 08.5 x 11.0 Landscape
    ------------------------------------------------------------------------------------------------
    DECLARE @RptWid    smallint        ; SET @RptWid    = @811WidPOR                             -- Report width
    DECLARE @WidMn0    smallint        ; SET @WidMn0    = @RptWid - 0                            -- Report width minus 0
    DECLARE @WidMn1    smallint        ; SET @WidMn1    = @RptWid - 1                            -- Report width minus 1
    DECLARE @WidMn2    smallint        ; SET @WidMn2    = @RptWid - 2                            -- Report width minus 2
    DECLARE @WidMn4    smallint        ; SET @WidMn4    = @RptWid - 4                            -- Report width minus 4
    ------------------------------------------------------------------------------------------------
    DECLARE @RptHgt    smallint        ; SET @RptHgt    = @811HgtPOR                             -- Report height
    DECLARE @RptAdj    smallint        ; SET @RptAdj    = 0                                      -- Report height adjustment
    DECLARE @AdjHgt    smallint        ; SET @AdjHgt    = 0                                      -- Adjusted report count
    DECLARE @LinCnt    smallint        ; SET @LinCnt    = 0                                      -- Line count
    ------------------------------------------------------------------------------------------------
    DECLARE @WidSLT    varchar(10)     ; SET @WidSLT    = 'SLT'                                  -- Report width prefix for SQX
    DECLARE @WidDLT    varchar(10)     ; SET @WidDLT    = 'DLT'                                  -- Report width prefix for SQX
    ------------------------------------------------------------------------------------------------
    SET @LftLen = @RptWid - @LftWid                                                              -- Initial Value
    SET @StmLen = @RptWid - @StmWid                                                              -- Initial Value
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Assign Standard Line Variables  (LNV)                                   EXEC ut_zzUTL zzz,LNV
    ------------------------------------------------------------------------------------------------
    DECLARE @LinSgl    varchar(200)    ; SET @LinSgl    = "'"+REPLICATE(@SGL,@WidMn1 - @LftWid)
    DECLARE @LinDbl    varchar(200)    ; SET @LinDbl    = "'"+REPLICATE(@DBL,@WidMn1 - @LftWid)
    DECLARE @LinAst    varchar(200)    ; SET @LinAst    = "'"+REPLICATE(@AST,@WidMn1 - @LftWid)
    DECLARE @LinPnd    varchar(200)    ; SET @LinPnd    = "'"+REPLICATE(@PND,@WidMn1 - @LftWid)
    DECLARE @LinAts    varchar(200)    ; SET @LinAts    = "'"+REPLICATE(@ATS,@WidMn1 - @LftWid)
    DECLARE @LinTld    varchar(200)    ; SET @LinTld    = "'"+REPLICATE(@TLD,@WidMn1 - @LftWid)
    DECLARE @LinBng    varchar(200)    ; SET @LinBng    = "'"+REPLICATE(@BNG,@WidMn1 - @LftWid)
    DECLARE @LinCmt    varchar(200)    ; SET @LinCmt    = "' "
    ------------------------------------------------------------------------------------------------
    DECLARE @HdrBeg    varchar(200)    ; SET @HdrBeg    = ''
    DECLARE @HdrEnd    varchar(200)    ; SET @HdrEnd    = ''
    DECLARE @HdrCmt    varchar(200)    ; SET @HdrCmt    = ''
    DECLARE @HdrSep    varchar(200)    ; SET @HdrSep    = ''
    DECLARE @LinWid    smallint        ; SET @LinWid    = LEN(@LinSgl)
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Declare Extended Utility Variables  (XUV)                               EXEC ut_zzUTL zzz,XUV
    ------------------------------------------------------------------------------------------------
    DECLARE @LM        tinyint         ; SET @LM        = 0                                      -- Utility Flag:  Left margin
    DECLARE @SP        tinyint         ; SET @SP        = 0                                      -- Utility Flag:  Space lines
    DECLARE @TL        tinyint         ; SET @TL        = 0                                      -- Utility Flag:  Include header title
    DECLARE @BT        tinyint         ; SET @BT        = 0                                      -- Utility Flag:  Include batch GO
    DECLARE @TR        tinyint         ; SET @TR        = 0                                      -- Utility Flag:  Include transaction logic
    DECLARE @ID        tinyint         ; SET @ID        = 0                                      -- Utility Flag:  Include identity column
    DECLARE @EM        tinyint         ; SET @EM        = 0                                      -- Utility Flag:  Include error message
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Development utility constants (EXEC ut_zzNAX DEVUTL,DVU)
    ------------------------------------------------------------------------------------------------
    DECLARE @DevUtlTBX varchar(03)     ; SET @DevUtlTBX = 'TBX'                                   -- Table Scripts
    DECLARE @DevUtlVEX varchar(03)     ; SET @DevUtlVEX = 'VEX'                                   -- View  Scripts
    DECLARE @DevUtlUSX varchar(03)     ; SET @DevUtlUSX = 'USX'                                   -- SProc Scripts
    DECLARE @DevUtlTRX varchar(03)     ; SET @DevUtlTRX = 'TRX'                                   -- Trigger Scripts
    DECLARE @DevUtlUFX varchar(03)     ; SET @DevUtlUFX = 'UFX'                                   -- Function Scripts
    DECLARE @DevUtlLKX varchar(03)     ; SET @DevUtlLKX = 'LKX'                                   -- Lookup Scripts
    DECLARE @DevUtlDMX varchar(03)     ; SET @DevUtlDMX = 'DMX'                                   -- Dimension Scripts
    DECLARE @DevUtlFTX varchar(03)     ; SET @DevUtlFTX = 'FTX'                                   -- Fact Scripts
    DECLARE @DevUtlPOX varchar(03)     ; SET @DevUtlPOX = 'POX'                                   -- Population Scripts
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Object Class (EXEC ut_zzNAX OBJCLS,JCL)
    ------------------------------------------------------------------------------------------------
    DECLARE @ObjClsTBL varchar(03)     ; SET @ObjClsTBL = 'TBL'                                   -- Table
    DECLARE @ObjClsVEW varchar(03)     ; SET @ObjClsVEW = 'VEW'                                   -- View
    DECLARE @ObjClsUSP varchar(03)     ; SET @ObjClsUSP = 'USP'                                   -- SProc
    DECLARE @ObjClsTRG varchar(03)     ; SET @ObjClsTRG = 'TRG'                                   -- Trigger
    DECLARE @ObjClsUFN varchar(03)     ; SET @ObjClsUFN = 'UFN'                                   -- Function
    DECLARE @ObjClsPKY varchar(03)     ; SET @ObjClsPKY = 'PKY'                                   -- PrimaryKey
    DECLARE @ObjClsUKY varchar(03)     ; SET @ObjClsUKY = 'UKY'                                   -- UniqueKey
    DECLARE @ObjClsIND varchar(03)     ; SET @ObjClsIND = 'IND'                                   -- Index
    DECLARE @ObjClsSTT varchar(03)     ; SET @ObjClsSTT = 'STT'                                   -- Statistic
    DECLARE @ObjClsFKY varchar(03)     ; SET @ObjClsFKY = 'FKY'                                   -- ForeignKey
    DECLARE @ObjClsDEF varchar(03)     ; SET @ObjClsDEF = 'DEF'                                   -- Default
    DECLARE @ObjClsCHK varchar(03)     ; SET @ObjClsCHK = 'CHK'                                   -- Check
    DECLARE @ObjClsDDL varchar(03)     ; SET @ObjClsDDL = 'DDL'                                   -- DataDict
    DECLARE @ObjClsSCP varchar(03)     ; SET @ObjClsSCP = 'SCP'                                   -- Script
    DECLARE @ObjClsVDN varchar(03)     ; SET @ObjClsVDN = 'VDN'                                   -- VB.NET
    DECLARE @ObjClsUNK varchar(03)     ; SET @ObjClsUNK = 'UNK'                                   -- Unknown
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Object Type (sys.objects.type) (EXEC ut_zzNAX OBJTYP,JTP)
    ------------------------------------------------------------------------------------------------
    DECLARE @ObjTypTBL varchar(03)     ; SET @ObjTypTBL = 'U'                                     -- Table
    DECLARE @ObjTypVEW varchar(03)     ; SET @ObjTypVEW = 'V'                                     -- View
    DECLARE @ObjTypUSP varchar(03)     ; SET @ObjTypUSP = 'P'                                     -- SProc
    DECLARE @ObjTypTRG varchar(03)     ; SET @ObjTypTRG = 'TR'                                    -- Trigger
    DECLARE @ObjTypUFN varchar(03)     ; SET @ObjTypUFN = 'FN'                                    -- Function
    DECLARE @ObjTypPKY varchar(03)     ; SET @ObjTypPKY = 'PK'                                    -- PrimaryKey
    DECLARE @ObjTypUKY varchar(03)     ; SET @ObjTypUKY = 'UQ'                                    -- UniqueKey
    DECLARE @ObjTypIND varchar(03)     ; SET @ObjTypIND = ''                                      -- Index
    DECLARE @ObjTypSTT varchar(03)     ; SET @ObjTypSTT = ''                                      -- Statistic
    DECLARE @ObjTypFKY varchar(03)     ; SET @ObjTypFKY = 'F'                                     -- ForeignKey
    DECLARE @ObjTypDEF varchar(03)     ; SET @ObjTypDEF = 'D'                                     -- Default
    DECLARE @ObjTypCHK varchar(03)     ; SET @ObjTypCHK = 'C'                                     -- Check
    DECLARE @ObjTypDDL varchar(03)     ; SET @ObjTypDDL = ''                                      -- DataDict
    DECLARE @ObjTypSCP varchar(03)     ; SET @ObjTypSCP = ''                                      -- Script
    DECLARE @ObjTypVDN varchar(03)     ; SET @ObjTypVDN = ''                                      -- VB.NET
    DECLARE @ObjTypUNK varchar(03)     ; SET @ObjTypUNK = ''                                      -- Unknown
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Synchronize Display object with Source object  (DJS)
    ------------------------------------------------------------------------------------------------
    DECLARE @DspObj    sysname      = ''                                                          -- Display Object (replaces @SrcObj)
    IF LEN(@DspObj) = 0 SET @DspObj = @InpObj
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT DJS='DJS',SrcObj=LEFT(@InpObj,30),DspObj=LEFT(@DspObj,30)
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Assign Display object values  (DJV)
    ------------------------------------------------------------------------------------------------
    DECLARE @DspNam    sysname      = @DspObj                                                     -- Display Name (replaces @ObjNam)
    DECLARE @DspTbl    sysname      ; EXEC ut_zzNAM TBL,NAM,XXX,@DspNam,@DspTbl OUTPUT
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT DJV='DJV',DspObj=LEFT(@DspObj,30),DspNam=LEFT(@DspNam,30),DspTbl=LEFT(@DspTbl,30)
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Synchronize Empty Object With FirstTable  (ROT)
    ------------------------------------------------------------------------------------------------
    IF LEN(@OupObj) = 0 SET @OupObj = @InpObj
    ------------------------------------------------------------------------------------------------
    -- Resolve Ouput object (from Display name - ROT)
    EXEC dbo.ut_zzNAM OBJ,NAM,@NamFmt,@InpObj,@OupObj OUTPUT,0  --','',1
    EXEC dbo.ut_zzNAM OBJ,DSC,@NamFmt,@InpObj,@OupDsc OUTPUT,0  --','',1
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'ROT' AS 'ROT',InpObj=LEFT(@InpObj,30),NamFmt=@NamFmt,OupObj=LEFT(@OupObj,30),OupDsc=@OupDsc
    -- DECLARE @OUP sysname; EXEC dbo.ut_zzNAM OBJ,NAM,@NamFmt,pfx_ObjNam,@OUP OUTPUT; PRINT OUP=@OUP
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize standard OutPut Code Constants  (OCC)
    ------------------------------------------------------------------------------------------------
    DECLARE @SecALL    varchar(11)     ; SET @SecALL    = 'ALL'                                   -- Include all code sections
    DECLARE @SecXXX    varchar(11)     ; SET @SecXXX    = 'XXX'                                   -- New code section
    DECLARE @SecZZZ    varchar(11)     ; SET @SecZZZ    = 'ZZZ'                                   -- Invalid code error message
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Initialize parameter OutPut Code Constants
    ------------------------------------------------------------------------------------------------
    DECLARE @SecLCP    varchar(11)     ; SET @SecLCP    = 'LCP'                                   -- Lookup Constants: PKey
    DECLARE @SecLCC    varchar(11)     ; SET @SecLCC    = 'LCC'                                   -- Lookup Constants: Code
    DECLARE @SecLCN    varchar(11)     ; SET @SecLCN    = 'LCN'                                   -- Lookup Constants: Name
    DECLARE @SecLCX    varchar(11)     ; SET @SecLCX    = 'LCX'                                   -- Lookup Constants: CmdTxt
 
    DECLARE @SecLPP    varchar(11)     ; SET @SecLPP    = 'LPP'                                   -- Lookup Properties: PKey
    DECLARE @SecLPC    varchar(11)     ; SET @SecLPC    = 'LPC'                                   -- Lookup Properties: Code
    DECLARE @SecLPN    varchar(11)     ; SET @SecLPN    = 'LPN'                                   -- Lookup Properties: Name
    DECLARE @SecLPX    varchar(11)     ; SET @SecLPX    = 'LPX'                                   -- Lookup Properties: CmdTxt
 
    DECLARE @SecMLV    varchar(11)     ; SET @SecMLV    = 'MLV'                                   -- Module Level Variables
    DECLARE @SecMLP    varchar(11)     ; SET @SecMLP    = 'MLP'                                   -- Module Level Properties - LET/GET
    DECLARE @SecMLL    varchar(11)     ; SET @SecMLL    = 'MLL'                                   -- Module Level Properties - LET
    DECLARE @SecMLG    varchar(11)     ; SET @SecMLG    = 'MLG'                                   -- Module Level Properties - GET
 
    DECLARE @SecMAV    varchar(11)     ; SET @SecMAV    = 'MAV'                                   -- Module Level AssignSQL (from variables)
    DECLARE @SecMAC    varchar(11)     ; SET @SecMAC    = 'MAC'                                   -- Module Level AssignSQL (from controls)
    DECLARE @SecMAN    varchar(11)     ; SET @SecMAN    = 'MAN'                                   -- Module Level AssignSQL (from nulls)
 
    DECLARE @SecMRV    varchar(11)     ; SET @SecMRV    = 'MRV'                                   -- Module Level ReadSQL (into variables)
    DECLARE @SecMRC    varchar(11)     ; SET @SecMRC    = 'MRC'                                   -- Module Level ReadSQL (into controls)
 
    DECLARE @SecMCV    varchar(11)     ; SET @SecMCV    = 'MCV'                                   -- Module Level ClearSQL (variables)
    DECLARE @SecMCC    varchar(11)     ; SET @SecMCC    = 'MCC'                                   -- Module Level ClearSQL (controls)
 
    DECLARE @SecMIF    varchar(11)     ; SET @SecMIF    = 'MIF'                                   -- Module IF Criteria
 
    DECLARE @SecDTV    varchar(11)     ; SET @SecDTV    = 'DTV'                                   -- Declare fields column variables
    DECLARE @SecITV    varchar(11)     ; SET @SecITV    = 'ITV'                                   -- Initialize fields column variables
    DECLARE @SecATV    varchar(11)     ; SET @SecATV    = 'ATV'                                   -- Assign fields column variables
 
    DECLARE @SecDSV    varchar(11)     ; SET @SecDSV    = 'DSV'                                   -- Declare statement column variables
    DECLARE @SecISV    varchar(11)     ; SET @SecISV    = 'ISV'                                   -- Initialize statement column variables
    DECLARE @SecASV    varchar(11)     ; SET @SecASV    = 'ASV'                                   -- Assign statement column variables
 
    DECLARE @SecDKV    varchar(11)     ; SET @SecDKV    = 'DKV'                                   -- Declare primary key column variables
    DECLARE @SecIKV    varchar(11)     ; SET @SecIKV    = 'IKV'                                   -- Initialize primary key column variables
 
    DECLARE @SecAKV    varchar(11)     ; SET @SecAKV    = 'AKV'                                   -- Assign primary key column variables
    DECLARE @SecDGV    varchar(11)     ; SET @SecDGV    = 'DGV'                                   -- Declare function parameter variables
    DECLARE @SecIGV    varchar(11)     ; SET @SecIGV    = 'IGV'                                   -- Initialize parameter variables
 
    DECLARE @SecAGV    varchar(11)     ; SET @SecAGV    = 'AGV'                                   -- Assign parameter variables
 
    DECLARE @SecRSV    varchar(11)     ; SET @SecRSV    = 'RSV'                                   -- Recordset variables
    DECLARE @SecRSL    varchar(11)     ; SET @SecRSL    = 'RSL'                                   -- Recordset standard loop
    DECLARE @SecRSW    varchar(11)     ; SET @SecRSW    = 'RSW'                                   -- Recordset write loop
 
    DECLARE @SecBASWTL varchar(11)     ; SET @SecBASWTL = 'BASWTL'                                -- Build module:  bas_UtlWTL
    DECLARE @SecCLSWTL varchar(11)     ; SET @SecCLSWTL = 'CLSWTL'                                -- Build module:  clsUtlWTL
 
    DECLARE @SecBASAPC varchar(11)     ; SET @SecBASAPC = 'BASAPC'                                -- Build module:  bas_AppCons
    DECLARE @SecBASAPF varchar(11)     ; SET @SecBASAPF = 'BASAPF'                                -- Build module:  bas_AppFunc
    DECLARE @SecBASAPT varchar(11)     ; SET @SecBASAPT = 'BASAPT'                                -- Build module:  bas_AppTest
    DECLARE @SecBASAPV varchar(11)     ; SET @SecBASAPV = 'BASAPV'                                -- Build module:  bas_AppVars
 
    DECLARE @SecBASGLB varchar(11)     ; SET @SecBASGLB = 'BASGLB'                                -- Build module:  bas_Global
    DECLARE @SecBASIMX varchar(11)     ; SET @SecBASIMX = 'BASIMX'                                -- Build module:  bas_ImpExp
    DECLARE @SecBASTST varchar(11)     ; SET @SecBASTST = 'BASTST'                                -- Build module:  bas_Test01
    DECLARE @SecBASTBM varchar(11)     ; SET @SecBASTBM = 'BASTBM'                                -- Build module:  bas_TblMnt
 
    DECLARE @SecUTLASC varchar(11)     ; SET @SecUTLASC = 'UTLASC'                                -- Build module:  clsUtlASC
    DECLARE @SecUTLFMT varchar(11)     ; SET @SecUTLFMT = 'UTLFMT'                                -- Build module:  clsUtlFMT
    DECLARE @SecUTLVBG varchar(11)     ; SET @SecUTLVBG = 'UTLVBG'                                -- Build module:  clsUtlVBG
    DECLARE @SecUTLWSH varchar(11)     ; SET @SecUTLWSH = 'UTLWSH'                                -- Build module:  clsUtlWSH
    DECLARE @SecUTLWTX varchar(11)     ; SET @SecUTLWTX = 'UTLWTX'                                -- Build module:  clsUtlWTX
 
    DECLARE @SecGENGLB varchar(11)     ; SET @SecGENGLB = 'GENGLB'                                -- Build module:  vba_Global
    DECLARE @SecGENSTD varchar(11)     ; SET @SecGENSTD = 'GENSTD'                                -- Build module:  vbaGenSTD
    DECLARE @SecGENJET varchar(11)     ; SET @SecGENJET = 'GENJET'                                -- Build module:  vbaGenJET
 
    DECLARE @SecGEN_IT varchar(11)     ; SET @SecGEN_IT = 'GEN_IT'                                -- Build module:  vbaGen_IT
    DECLARE @SecGENFRM varchar(11)     ; SET @SecGENFRM = 'GENFRM'                                -- Build module:  vbaGenFRM
    DECLARE @SecGENCTL varchar(11)     ; SET @SecGENCTL = 'GENCTL'                                -- Build module:  vbaGenCTL
    DECLARE @SecGENTBL varchar(11)     ; SET @SecGENTBL = 'GENTBL'                                -- Build module:  vbaGenTBL
    DECLARE @SecGENPRP varchar(11)     ; SET @SecGENPRP = 'GENPRP'                                -- Build module:  vbaGenPRP
    DECLARE @SecGENCMD varchar(11)     ; SET @SecGENCMD = 'GENCMD'                                -- Build module:  vbaGenCMD
    DECLARE @SecGENRPT varchar(11)     ; SET @SecGENRPT = 'GENRPT'                                -- Build module:  vbaGenRPT
    DECLARE @SecGENPTH varchar(11)     ; SET @SecGENPTH = 'GENPTH'                                -- Build module:  vbaGenPTH
    DECLARE @SecGENSQL varchar(11)     ; SET @SecGENSQL = 'GENSQL'                                -- Build module:  vbaGenSQL
    DECLARE @SecGENSBY varchar(11)     ; SET @SecGENSBY = 'GENSBY'                                -- Build module:  vbaGenSBY
    DECLARE @SecGENGBY varchar(11)     ; SET @SecGENGBY = 'GENGBY'                                -- Build module:  vbaGenGBY
    DECLARE @SecGENSLO varchar(11)     ; SET @SecGENSLO = 'GENSLO'                                -- Build module:  vbaGenSLO
 
    DECLARE @SecCLSAPC varchar(11)     ; SET @SecCLSAPC = 'CLSAPC'                                -- Build module:  clsAppCons
    DECLARE @SecCLSAPV varchar(11)     ; SET @SecCLSAPV = 'CLSAPV'                                -- Build module:  clsAppVals
 
    DECLARE @SecBASCMG varchar(11)     ; SET @SecBASCMG = 'BASCMG'                                -- Build module:  bas_CmgCons
    DECLARE @SecCLSCMG varchar(11)     ; SET @SecCLSCMG = 'CLSCMG'                                -- Build module:  clsCtlMgr
 
    DECLARE @SecREGTBL varchar(11)     ; SET @SecREGTBL = 'REGTBL'                                -- Build module:  clsRegTBL
    DECLARE @SecREGPRP varchar(11)     ; SET @SecREGPRP = 'REGPRP'                                -- Build module:  clsRegPRP
    DECLARE @SecREGCMD varchar(11)     ; SET @SecREGCMD = 'REGCMD'                                -- Build module:  clsRegCMD
    DECLARE @SecREGRPT varchar(11)     ; SET @SecREGRPT = 'REGRPT'                                -- Build module:  clsRegRPT
    DECLARE @SecREGPTH varchar(11)     ; SET @SecREGPTH = 'REGPTH'                                -- Build module:  clsRegPTH
    DECLARE @SecREGSRC varchar(11)     ; SET @SecREGSRC = 'REGSRC'                                -- Build module:  clsRegSRC
 
    DECLARE @SecSQLSTM varchar(11)     ; SET @SecSQLSTM = 'SQLSTM'                                -- Build module:  clsSqlSTM
    DECLARE @SecSQLOBY varchar(11)     ; SET @SecSQLOBY = 'SQLOBY'                                -- Build module:  clsSqlOBY
    DECLARE @SecRUNWHR varchar(11)     ; SET @SecRUNWHR = 'RUNWHR'                                -- Build module:  clsRunWHR
 
    DECLARE @SecRUNCMD varchar(11)     ; SET @SecRUNCMD = 'RUNCMD'                                -- Build module:  clsRunCMD
    DECLARE @SecRUNCMM varchar(11)     ; SET @SecRUNCMM = 'RUNCMM'                                -- Build module:  Run_Process_0000 (CALL cls_Method)
    DECLARE @SecRUNCMF varchar(11)     ; SET @SecRUNCMF = 'RUNCMF'                                -- Build module:  Run_Process_0000 (OPEN frm_FrmNam)
 
    DECLARE @SecRUNRPT varchar(11)     ; SET @SecRUNRPT = 'RUNRPT'                                -- Build module:  clsRunRPT
    DECLARE @SecRUNRPR varchar(11)     ; SET @SecRUNRPR = 'RUNRPR'                                -- Build module:  Run_Report_0000
    DECLARE @SecRUNRPX varchar(11)     ; SET @SecRUNRPX = 'RUNRPX'                                -- Build module:  Print_rpt_ReportName
 
    DECLARE @SecRUNUSP varchar(11)     ; SET @SecRUNUSP = 'RUNUSP'                                -- Build module:  clsRunUSP
    DECLARE @SecRUNUSR varchar(11)     ; SET @SecRUNUSR = 'RUNUSR'                                -- Build module:  Run_Process_0000 (EXEC PROC)
    DECLARE @SecRUNUSF varchar(11)     ; SET @SecRUNUSF = 'RUNUSF'                                -- Build module:  Run_Process_0000 (OPEN FORM)
    DECLARE @SecRUNUSX varchar(11)     ; SET @SecRUNUSX = 'RUNUSX'                                -- Build SProcs:  Execute_usp_ProcedureName
 
    DECLARE @SecRUNRST varchar(11)     ; SET @SecRUNRST = 'RUNRST'                                -- Build module:  clsRunRST
    DECLARE @SecRUNSQL varchar(11)     ; SET @SecRUNSQL = 'RUNSQL'                                -- Build module:  clsRunSQL
    DECLARE @SecRUNSBY varchar(11)     ; SET @SecRUNSBY = 'RUNSBY'                                -- Build module:  clsRunSBY
    DECLARE @SecRUNGBY varchar(11)     ; SET @SecRUNGBY = 'RUNGBY'                                -- Build module:  clsRunGBY
 
    DECLARE @SecFRMCLR varchar(11)     ; SET @SecFRMCLR = 'FRMCLR'                                -- Build module:  frm_FrmName
    DECLARE @SecFRMLNK varchar(11)     ; SET @SecFRMLNK = 'FRMLNK'                                -- Build module:  sys_LinkAPP
 
    DECLARE @SecRPTNAR varchar(11)     ; SET @SecRPTNAR = 'RPTNAR'                                -- Build module:  tpl_NARROW
    DECLARE @SecRPTWID varchar(11)     ; SET @SecRPTWID = 'RPTWID'                                -- Build module:  tpl_WIDE
 
    DECLARE @SecANYFRM varchar(11)     ; SET @SecANYFRM = 'ANYFRM'                                -- Build module:  frm_FrmName
    DECLARE @SecANYTAB varchar(11)     ; SET @SecANYTAB = 'ANYTAB'                                -- Build module:  frm_FrmName
    DECLARE @SecANYLST varchar(11)     ; SET @SecANYLST = 'ANYLST'                                -- Build module:  lst_FrmName
    DECLARE @SecANYPOP varchar(11)     ; SET @SecANYPOP = 'ANYPOP'                                -- Build module:  pop_FrmName
    DECLARE @SecANYSUB varchar(11)     ; SET @SecANYSUB = 'ANYSUB'                                -- Build module:  sub_FrmName
    DECLARE @SecANYBAS varchar(11)     ; SET @SecANYBAS = 'ANYBAS'                                -- Build module:  basBasName
    DECLARE @SecANYCLS varchar(11)     ; SET @SecANYCLS = 'ANYCLS'                                -- Build module:  clsClsName
    DECLARE @SecANYRPT varchar(11)     ; SET @SecANYRPT = 'ANYRPT'                                -- Build module:  rpt_RptNam

    DECLARE @SecMACMOD varchar(11)     ; SET @SecMACMOD = 'MACMOD'                                -- Build module:  clsClsName
    DECLARE @SecMACTST varchar(11)     ; SET @SecMACTST = 'MACTST'                                -- Build module:  clsClsName
    DECLARE @SecMAFPOP varchar(11)     ; SET @SecMAFPOP = 'MAFPOP'                                -- Build module:  clsClsName
    DECLARE @SecMAFRVW varchar(11)     ; SET @SecMAFRVW = 'MAFRVW'                                -- Build module:  clsClsName
 
    DECLARE @SecCLSTCN varchar(11)     ; SET @SecCLSTCN = 'CLSTCN'                                -- Build module:  clsTxtCon
 
    DECLARE @SecXTDXYR varchar(11)     ; SET @SecXTDXYR = 'XTDXYR'                                -- Extend Tax Year
    DECLARE @SecXTDXPD varchar(11)     ; SET @SecXTDXPD = 'XTDXPD'                                -- Extend Tax Period
    DECLARE @SecXTDXMN varchar(11)     ; SET @SecXTDXMN = 'XTDXMN'                                -- Extend Tax Month
    DECLARE @SecXTDXAY varchar(11)     ; SET @SecXTDXAY = 'XTDXAY'                                -- Extend Active Tax Year
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################


    --XGM@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@XGM

 
    --##############################################################################################
 

    ------------------------------------------------------------------------------------------------
    -- Declare working variables
    ------------------------------------------------------------------------------------------------
    DECLARE @ClsLst    varchar(200) ; SET @ClsLst    = ""
    DECLARE @DefTyp    varchar(11)  ; SET @DefTyp    = ""                  -- Default object type code
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- LCP = Lookup Constants: PKey
    -- LCC = Lookup Constants: Code
    -- LCN = Lookup Constants: Name
    -- LCX = Lookup Constants: CmdTxt
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- LPP = Lookup Properties: PKey
    -- LPC = Lookup Properties: Code
    -- LPN = Lookup Properties: Name
    -- LPX = Lookup Properties: CmdTxt
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA LCP,zzz_TEST01
        EXEC ut_zzVBA LCC,zzz_TEST01
        EXEC ut_zzVBA LCN,zzz_TEST01
        EXEC ut_zzVBA LCX,zzz_TEST01
        EXEC ut_zzVBA LPP,zzz_TEST01
        EXEC ut_zzVBA LPC,zzz_TEST01
        EXEC ut_zzVBA LPN,zzz_TEST01
        EXEC ut_zzVBA LPX,zzz_TEST01
        --------------------------------------------------------------------------------------------
                       Oup Stx       Lft Spc Ttl Bat  BAS,VAR,SFX,VAL,COD,DSC,OBY            Tx2 Tx3
        EXEC ut_zzVBA LCP,zzz_TEST01,'SysTyp,SysTyp,TypCod,SysTypID,SysCod,TypNam,SysTypID','' ,''
        EXEC ut_zzVBA LCP,zzz_TEST01,'SysTyp,SysTyx,TypCod,SysCod  ,TypCod,TypNam,SysTypID','' ,''
        EXEC ut_zzVBA LCP,zzz_TEST01,'SysTyp,SysTyd,TypCod,TypNam  ,TypCod,SysCod,SysTypID','' ,''
        EXEC ut_zzVBA LCP,zzz_TEST01,0  ,0  ,0  ,0  ,'VbaTyp,VbaPfx,VbaPfx,VbaPfx  ,VbaPfx,VbaDtp,VbaTypID','' ,'' ,0  ,0  ,0
        EXEC ut_zzVBA LCP,zzz_TEST01,0  ,0  ,0  ,0  ,'VbaTyp,VbaTyp,VbaPfx,VbaTypID,VbaPfx,VbaDtp,VbaTypID','' ,'' ,0  ,0  ,0
    ----------------------------------------------------------------------------------------------*/
    IF @BldLST IN (@SecLCP,@SecLCC,@SecLCN,@SecLCX,@SecLPP,@SecLPC,@SecLPN,@SecLPX) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @LftMrg = 1
        SET @IncSpc = 1
        SET @IncTtl = 1
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft     Spc     Ttl     Bat     Tx1     Tx2     Tx3     Trn     Idn     Erm
        EXEC ut_zzVBX @BldLST,@InpObj,@LftMrg,@IncSpc,@IncTtl,@IncBat,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncErm
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- MLV = Module Level Variables
    -- MLP = Module Level Properties - LET/GET
    -- MLL = Module Level Properties - LET
    -- MLG = Module Level Properties - GET
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- MAV = Module Level AssignSQL (from variables)
    -- MAC = Module Level AssignSQL (from controls)
    -- MAN = Module Level AssignSQL (from nulls)
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- MRV = Module Level ReadSQL (into variables)
    -- MRC = Module Level ReadSQL (into controls)
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- MCV = Module Level ClearSQL (variables)
    -- MCC = Module Level ClearSQL (controls)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MLV,zzz_TEST01    -- Module Level Variables
        EXEC ut_zzVBA MLP,zzz_TEST01    -- Module Level Properties - LET/GET
        EXEC ut_zzVBA MLL,zzz_TEST01    -- Module Level Properties - LET
        EXEC ut_zzVBA MLG,zzz_TEST01    -- Module Level Properties - GET
        EXEC ut_zzVBA MAV,zzz_TEST01    -- Module Level AssignSQL (from variables)
        EXEC ut_zzVBA MAC,zzz_TEST01    -- Module Level AssignSQL (from controls)
        EXEC ut_zzVBA MAN,zzz_TEST01    -- Module Level AssignSQL (from nulls)
        EXEC ut_zzVBA MRV,zzz_TEST01    -- Module Level ReadSQL (into variables)
        EXEC ut_zzVBA MRC,zzz_TEST01    -- Module Level ReadSQL (into controls)
        EXEC ut_zzVBA MCV,zzz_TEST01    -- Module Level ClearSQL (variables)
        EXEC ut_zzVBA MCC,zzz_TEST01    -- Module Level ClearSQL (controls)
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMLV,@SecMLP,@SecMLL,@SecMLG,@SecMAV,@SecMAC,@SecMAN,@SecMRV,@SecMRC,@SecMCV,@SecMCC) BEGIN  -- (OCT)
    ------------------------------------------------------------------------------------------------
        SET @LftMrg = 0
        SET @IncSpc = 2
        SET @IncTtl = 2
        SET @IncHdr = 1
        --------------------------------------------------------------------------------------------
        --            Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
        EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- MIF = Module IF Criteria
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MIF,zzz_TEST01    -- Module IF Criteria
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMIF) BEGIN  -- (OCT)
    ------------------------------------------------------------------------------------------------
        SET @LftMrg = 1
        SET @IncSpc = 0
        SET @IncTtl = 0
        SET @IncHdr = 0
        --------------------------------------------------------------------------------------------
        --            Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
        EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DTV = Declare fields column variables
    -- ITV = Initialize fields column variables
    -- ATV = Assign fields column variables
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- DSV = Declare statement column variables
    -- ISV = Initialize statement column variables
    -- ASV = Assign statement column variables
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- DKV = Declare primary key column variables
    -- IKV = Initialize primary key column variables
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- AKV = Assign primary key column variables
    -- DGV = Declare function parameter variables
    -- IGV = Initialize parameter variables
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- AGV = Assign parameter variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA DTV
        EXEC ut_zzVBA ITV
        EXEC ut_zzVBA ATV
        EXEC ut_zzVBA DSV
        EXEC ut_zzVBA ISV
        EXEC ut_zzVBA ASV
        EXEC ut_zzVBA DKV
        EXEC ut_zzVBA IKV
        EXEC ut_zzVBA AKV
        EXEC ut_zzVBA DGV
        EXEC ut_zzVBA IGV
        EXEC ut_zzVBA AGV
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecDTV,@SecITV,@SecATV,@SecDSV,@SecISV,@SecASV,@SecDKV,@SecIKV,@SecAKV,@SecDGV,@SecIGV,@SecAGV) BEGIN  -- (OCT)
    ------------------------------------------------------------------------------------------------
        SET @LftMrg = 1
        SET @IncSpc = 2
        SET @IncTtl = 1
        --------------------------------------------------------------------------------------------
        --            Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
        EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RSV = Recordset variables
    -- RSL = Recordset standard loop
    -- RSW = Recordset write loop
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RSV,vba_SrcDfn
        EXEC ut_zzVBA RSL,vba_TblDfn
        EXEC ut_zzVBA RSW,vba_TblDfn
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRSV,@SecRSL,@SecRSW) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @LftMrg = 1
        SET @IncSpc = 1
        SET @IncTtl = 1
        --------------------------------------------------------------------------------------------
        -- Adjust values
        --------------------------------------------------------------------------------------------
        SET @BldLST = CASE @BldLST
            WHEN @SecRSV THEN 'VLN=20,DSV'
            WHEN @SecRSL THEN 'DSV'
            WHEN @SecRSW THEN 'DSV'
            ELSE @BldLST
        END
        --------------------------------------------------------------------------------------------
        --            Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
        EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
        RETURN
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- BASWTL = Build module:  bas_UtlWTL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASWTL,bas_UtlWTL,'Manage Text Output'
    ----------------------------------------------------------------------------------------------*/
    IF @BldLST IN (@SecBASWTL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_UtlWTL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Text Output'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1       Tx2     Tx3 Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1 ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TFL    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TCV    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXC    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXM    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXW    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TOB    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- CLSWTL = Build module:  clsUtlWTL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA CLSWTL,clsUtlWTL,'Manage Text Output - Light'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecCLSWTL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlWTL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Text Output - Light'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'wtl'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1       Tx2     Tx3 Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1 ,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TFL    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TCV    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX IWT    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXP    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXC    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXM    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TXW    ,''     ,0  ,0  ,0  ,0  ,'Public',''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX TOC    ,''     ,0  ,0  ,0  ,0  ,''      ,''     ,'' ,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- BASAPC = Build module:  bas_AppCons
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASAPC
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASAPC) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_AppCons'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Constants'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASAPF = Build module:  bas_AppFunc
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASAPF
        EXEC ut_zzVBA BASAPF,'','','UtlWSH'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASAPF) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_AppFunc'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Functions'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASAPT = Build module:  bas_AppTest
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASAPT
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASAPT) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_AppTest'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Testing'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASAPV = Build module:  bas_AppVars
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASAPV
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASAPV) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_AppVars'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Variables'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- BASGLB = Build module:  bas_Global
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASGLB
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASGLB,'','','Data'
        EXEC ut_zzVBA BASGLB,'','','Basic'
        EXEC ut_zzVBA BASGLB,'','','Forms'
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASGLB,'','','Data' ,'Gen'
        EXEC ut_zzVBA BASGLB,'','','Basic','Gen'
        EXEC ut_zzVBA BASGLB,'','','Forms','Gen'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASGLB) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_Global'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Global Objects'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'Basic'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'Std'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,''     ,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASIMX = Build module:  bas_ImpExp
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASIMX
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASIMX) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_ImpExp'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Import/Export Data'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASTST = Build module:  bas_Test01
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASTST
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASTST) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_Test01'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Test Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- BASTBM = Build module:  bas_TblMnt
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASTBM
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASTBM) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_TblMnt'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Table Maintenance'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- UTLASC = Build module:  clsUtlASC
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA UTLASC,clsUtlASC,'Manage KeyASCii Assignments',kys
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecUTLASC) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlASC'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage KeyASCii Assignments'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'kys'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- UTLFMT = Build module:  clsUtlFMT
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA UTLFMT,clsUtlFMT,'Provide Standard Formatting',fmt
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecUTLFMT) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlFMT'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Provide Standard Formatting'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'fmt'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- UTLVBG = Build module:  clsUtlVBG
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA UTLVBG,clsUtlVBG,'Manage VB Code Generation',vbg
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecUTLVBG) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlVBG'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage VB Code Generation'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'vbg'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- UTLWSH = Build module:  clsUtlWSH
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA UTLWSH,clsUtlWSH,'WinScriptHost Functions',wsh
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecUTLWSH) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlWSH'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'WinScriptHost Functions'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'wsh'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- UTLWTX = Build module:  clsUtlWTX
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA UTLWTX,clsUtlWTX,'Manage Text Output',wtx
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecUTLWTX) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsUtlWTX'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Text Output'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'wtx'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- GENGLB = Build module:  vba_Global
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENGLB,vba_Global ,'Global Generation'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENGLB) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vba_Global'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Global Generation'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENSTD = Build module:  vbaGenSTD
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENSTD,vbaGenSTD ,'Generate Standard Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENSTD) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenSTD'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Standard Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENJET = Build module:  vbaGenJET
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENJET,vbaGenJET,'Generate JET Objects Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENJET) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenJET'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate JET Objects Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- GEN_IT = Build module:  vbaGen_IT
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GEN_IT,vbaGen_IT ,'Generate Code'
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GEN_IT,'','','sub_MntSRC|pop_SrcOBJ|lst_SrcCLM|pop_SrcCLM'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGEN_IT) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGen_IT'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXTBL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXPRP ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXCMD ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXRPT ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXSQL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXSBY ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXGBY ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXCTL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SGXFRM ,@BldLST,0  ,0  ,0  ,0  ,@StdTx2,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENFRM = Build module:  vbaGenFRM
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENFRM,vbaGenFRM ,'Generate Form Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENFRM) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenFRM'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Form Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXFRM ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENCTL = Build module:  vbaGenCTL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENCTL,vbaGenCTL ,'Generate Control Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENCTL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenCTL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Control Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXCTL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENTBL = Build module:  vbaGenTBL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENTBL,vbaGenTBL ,'Generate Table Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENTBL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenTBL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Table Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXTBL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENPRP = Build module:  vbaGenPRP
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENPRP,vbaGenPRP,'Generate Property Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENPRP) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenPRP'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Property Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXPRP ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENCMD = Build module:  vbaGenCMD
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENCMD,vbaGenCMD ,'Generate Command Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENCMD) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenCMD'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Command Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXCMD ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENRPT = Build module:  vbaGenRPT
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENRPT,vbaGenRPT ,'Generate Report Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENRPT) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenRPT'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Report Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXRPT ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENPTH = Build module:  vbaGenPTH
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENPTH,vbaGenPTH ,'Generate Path Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENPTH) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenPTH'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate Path Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXPTH ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENSQL = Build module:  vbaGenSQL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENSQL,vbaGenSQL,'Generate RunSQL Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENSQL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenSQL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate RunSQL Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXSQL ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENSBY = Build module:  vbaGenSBY
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENSBY,vbaGenSBY,'Generate SrtBy Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENSBY) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenSBY'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate SrtBy Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXSBY ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENGBY = Build module:  vbaGenGBY
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENGBY,vbaGenGBY,'Generate GrpBy Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENGBY) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenGBY'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate GrpBy Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXGBY ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- GENSLO = Build module:  vbaGenSLO
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA GENSLO,vbaGenSLO,'Generate SelOn Code'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecGENSLO) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'vbaGenSLO'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Generate SelOn Code'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGNFLG ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SGXSLO ,@BldLST,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- CLSAPC = Build module:  clsAppCons
    -- CLSAPV = Build module:  clsAppVals
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA CLSAPC,clsAppCons,'Standard Application Constants',apc
        EXEC ut_zzVBA CLSAPV,clsAppVals,'Standard Application Values'   ,apv
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecCLSAPC) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsAppCons'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Constants'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'apc'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX PPH    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecCLSAPV) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsAppVals'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Standard Application Values'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'apv'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX PPH    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- BASCMG = Build module:  bas_CmgCons
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA BASCMG,bas_CmgCons,'Control Manager Constants'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecBASCMG) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'bas_CmgCons'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Control Manager Constants'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX FIN    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- CLSCMG = Build module:  clsCtlMgr
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA CLSCMG,clsCtlMgr,'Manage Form Controls',cmg
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecCLSCMG) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsCtlMgr'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Form Controls'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'cmg'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- REGTBL = Build module:  clsRegTBL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGTBL,clsRegTBL,'Register Table Information'  ,rtb
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGTBL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegTBL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Table Information'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rtb'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- REGPRP = Build module:  clsRegPRP
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGPRP,clsRegPRP,'Register Property Information'  ,rtb
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGPRP) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegPRP'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Property Information'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rtb'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- REGCMD = Build module:  clsRegCMD
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGCMD,clsRegCMD,'Register Command Information',rcm
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGCMD) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegCMD'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Command Information'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rcm'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- REGRPT = Build module:  clsRegRPT
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGRPT,clsRegRPT,'Register Report Information' ,rrp
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGRPT) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegRPT'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Report Information'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rrp'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- REGPTH = Build module:  clsRegPTH
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGPTH,clsRegPTH,'Register Path Information' ,rrp
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGPTH) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegPTH'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Path Information'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rph'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- REGSRC = Build module:  clsRegSRC
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA REGSRC,clsRegSRC,'Register Source Statements'  ,rsc
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecREGSRC) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRegSRC'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Register Source Statements'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rsc'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DON    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- SQLSTM = Build module:  clsSqlSTM
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA SQLSTM,'
            zzz_TEST01
            vba_TblDfn
        ',clsSqlSTM,'Build SQL Statements',stm
        EXEC ut_zzVBA SQLSTM
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecSQLSTM) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = ''
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'clsSqlSTM'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'Build SQL Statements'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'stm'
        SET @ClsLst = 'rtb:clsRegTBL'
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@StdTx1,0  ,0  ,0  ,0  ,@StdTx2,@StdTx3,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX DCS    ,@ClsLst
        EXEC ut_zzVBX TCV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX STV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX ICS    ,@ClsLst
        EXEC ut_zzVBX TXC    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        SET @BldLST = "SQLSTM"
        SET @OupFmt = ""
        SET @OupObj = ""
        SET @OupDsc = ""
        SET @DspObj = ""
        --------------------------------------------------------------------------------------------
        SET @SEP = @SEP; SET @LST = @InpTxt
        SET @LST = LTRIM(RTRIM(REPLACE(REPLACE(@LST," ",""),@NLN,@SEP)))
        WHILE LEFT (@LST,LEN(@SEP)) = @SEP SET @LST = RIGHT(@LST,LEN(@LST)-LEN(@SEP))
        WHILE RIGHT(@LST,LEN(@SEP)) = @SEP SET @LST = LEFT (@LST,LEN(@LST)-LEN(@SEP))
        WHILE @LST LIKE "%"+@SEP+@SEP+"%"  SET @LST = REPLACE(@LST,@SEP+@SEP,@SEP)
        WHILE LEN(@LST) > 0 BEGIN
            SET @POS = CHARINDEX(@SEP,@LST)
            IF @POS > 0 BEGIN
                SET @ITM = LTRIM(RTRIM(LEFT(@LST,@POS-1)))
                SET @LST = LTRIM(RIGHT(@LST,LEN(@LST)-@POS-(LEN(@SEP)-1)))
            END ELSE BEGIN
                SET @ITM = LTRIM(RTRIM(@LST))
                SET @LST = ""
            END
            ----------------------------------------------------------------------------------------
            SET @InpObj = @ITM
            --   ut_zzVBJ Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
            EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
            ----------------------------------------------------------------------------------------
        END
        RETURN
    ------------------------------------------------------------------------------------------------
    -- SQLOBY = Build module:  clsSqlOBY
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA SQLOBY,clsSqlOBY,'Build SQL ORDER BY clause',oby
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecSQLOBY) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsSqlOBY'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Build SQL ORDER BY clause'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'oby'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- RUNWHR = Build module:  clsRunWHR
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNWHR,clsRunWHR,'Build SQL WHERE clause',whr
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNWHR) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRunWHR'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Build SQL WHERE clause'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'whr'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RUNCMD = Build module:  clsRunCMD
    -- RUNCMM = Build module:  Run_Process_0000 (CALL cls_Method)
    -- RUNCMF = Build module:  Run_Process_0000 (OPEN frm_FrmNam)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNCMD,clsRunCMD,'Run MSAccess Commands',cmd
        EXEC ut_zzVBA RUNCMM,0002,cls_Method
        EXEC ut_zzVBA RUNCMF,0002,frm_FrmNam
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNCMD) BEGIN
        IF LEN(@StdTx1) = 0 SET @InpTxt = 'clsRunCMD'
        IF LEN(@StdTx2) = 0 SET @StdTx1 = 'Run MSAccess Commands'
        IF LEN(@StdTx3) = 0 SET @StdTx2 = 'cmd'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecRUNCMM) BEGIN
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecRUNCMF) BEGIN
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RUNRPT = Build module:  clsRunRPT
    -- RUNRPR = Build module:  Run_Report_0000
    -- RUNRPX = Build module:  Print_rpt_ReportName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNRPT,clsRunRPT,'Run MSAccess Reports',run
        EXEC ut_zzVBA RUNRPR,0002,rpt_NewReport
        EXEC ut_zzVBA RUNRPX,rpt_NewReport
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNRPT) BEGIN
        IF LEN(@StdTx1) = 0 SET @InpTxt = 'clsRunRPT'
        IF LEN(@StdTx2) = 0 SET @StdTx1 = 'Run MSAccess Reports'
        IF LEN(@StdTx3) = 0 SET @StdTx2 = 'rpt'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecRUNRPR) BEGIN
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecRUNRPX) BEGIN
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RUNUSP = Build module:  clsRunUSP
    -- RUNUSR = Build module:  Run_Process_0000 (EXEC PROC)
    -- RUNUSF = Build module:  Run_Process_0000 (OPEN FORM)
    -- RUNUSX = Build SProcs:  Execute_usp_ProcedureName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNUSP,'
            usp_Get_DSLineage
        ',clsRunUSP,'Run Stored Procedures',run
        EXEC ut_zzVBA RUNUSR,0002,usp_NewProcName
        EXEC ut_zzVBA RUNUSF,0002,frm_NewFormName
        EXEC ut_zzVBA RUNUSX,usp_Get_DSLineage
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNUSP,@SecRUNUSX) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = ''
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'clsRunUSP'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'Run Stored Procedures'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'usp'
        SET @ClsLst = ''
        IF @BldLST IN (@SecRUNUSP) BEGIN
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC   ,@StdTx1,0  ,0  ,0  ,0  ,@StdTx2,@StdTx3,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV   ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX ORV   ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX SCV   ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        END
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX MOJVAR,@InpTxt,0  ,0  ,1  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX MOJPRP,@InpTxt,0  ,0  ,1  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        IF LEN(@InpTxt) = 0 BEGIN
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,''    ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        END
        --------------------------------------------------------------------------------------------
        SET @BldLST = "RUNUSX"
        SET @OupFmt = ""
        SET @OupObj = ""
        SET @OupDsc = ""
        SET @DspObj = ""
        --------------------------------------------------------------------------------------------
        SET @SEP = @SEP; SET @LST = @InpTxt
        SET @LST = LTRIM(RTRIM(REPLACE(REPLACE(@LST," ",""),@NLN,@SEP)))
        WHILE LEFT (@LST,LEN(@SEP)) = @SEP SET @LST = RIGHT(@LST,LEN(@LST)-LEN(@SEP))
        WHILE RIGHT(@LST,LEN(@SEP)) = @SEP SET @LST = LEFT (@LST,LEN(@LST)-LEN(@SEP))
        WHILE @LST LIKE "%"+@SEP+@SEP+"%"  SET @LST = REPLACE(@LST,@SEP+@SEP,@SEP)
        WHILE LEN(@LST) > 0 BEGIN
            SET @POS = CHARINDEX(@SEP,@LST)
            IF @POS > 0 BEGIN
                SET @ITM = LTRIM(RTRIM(LEFT(@LST,@POS-1)))
                SET @LST = LTRIM(RIGHT(@LST,LEN(@LST)-@POS-(LEN(@SEP)-1)))
            END ELSE BEGIN
                SET @ITM = LTRIM(RTRIM(@LST))
                SET @LST = ""
            END
            ----------------------------------------------------------------------------------------
            SET @InpObj = @ITM
            --   ut_zzVBJ Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
            EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
            ----------------------------------------------------------------------------------------
        END
        RETURN
    ------------------------------------------------------------------------------------------------
   END ELSE IF @BldCOD IN (@SecRUNUSR,@SecRUNUSF) BEGIN
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RUNRST = Build module:  clsRunRST
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNRST,clsRunRST,'Manage Criteria Persistence',rst
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNRST) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRunRST'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Criteria Persistence'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rst'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- RUNSQL = Build module:  clsRunSQL
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNSQL,clsRunSQL,'Manage SQL Statements',sql
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNSQL) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRunSQL'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage SQL Statements'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'sql'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- RUNSBY = Build module:  clsRunSBY
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNSBY,clsRunSBY,'Manage SortBy Array',sby
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNSBY) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRunSBY'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage SrtBy Array'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'sby'
        SET @ClsLst = ''
        --   ut_zzVBX Oup    Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- RUNGBY = Build module:  clsRunGBY
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RUNSBY,clsRunGBY,'Manage GroupBy Array',gby
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecRUNGBY) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsRunGBY'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage GrpBy Array'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'gby'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- FRMCLR = Build module:  frm_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA FRMCLR,sys_Colors,'Manage Color Schemes','frm'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecFRMCLR) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'sys_Colors'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Color Schemes'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'frm'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- FRMLNK = Build module:  sys_LinkAPP
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA FRMLNK,sys_LinkAPP,'Link Data Objects','lnk'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecFRMLNK) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'sys_LinkAPP'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Link Data Objects'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'lnk'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RPTNAR = Build module:  tpl_NARROW
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RPTNAR,tpl_NARROW,'Report_Description','rpt'
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecRPTNAR) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'tpl_NARROW'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Report_Description'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rpt'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX RMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- RPTWID = Build module:  tpl_WIDE
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA RPTWID,tpl_WIDE,'Report_Description','rpt'
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecRPTWID) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'tpl_WIDE'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Report_Description'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rpt'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX RMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- ANYFRM = Build module:  frm_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYFRM,frm_FrmName,'StdForm Module','CTCT','XXII'
        EXEC ut_zzVBA ANYFRM,tpl_FrmStd ,'StdForm Template'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYFRM) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'frm_FrmName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'StdForm Module'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'CTCTCTCTCT'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'XXIIXXIIXX'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYTAB = Build module:  frm_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYTAB,tab_FrmName,'TabForm Module'
        EXEC ut_zzVBA ANYTAB,tpl_FrmTab ,'TabForm Template'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYTAB) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'tab_FrmName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'TabForm Module'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'CTCTCTCTCT'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'XXIIXXIIXX'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYLST = Build module:  lst_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYLST,lst_FrmName,'ListForm Module'
        EXEC ut_zzVBA ANYLST,tpl_Lst010 ,'ListForm Template'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYLST) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'lst_FrmName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'ListForm Module'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'CTCTCTCTCT'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'XXIIXXIIXX'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYPOP = Build module:  pop_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYPOP,Cno_Modify_DataFeedChangeLog,'Modify CanalOne Data Feed Change Log'
        EXEC ut_zzVBA ANYPOP,pop_FrmName ,'Popup Module'
        EXEC ut_zzVBA ANYPOP,tpl_PopAdd  ,"Add New Data"
        EXEC ut_zzVBA ANYPOP,tpl_TestPops,'Test Popup Forms'
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecANYPOP) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'pop_FrmName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Popup Module'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'TC'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'II'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYSUB = Build module:  sub_FrmName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYSUB,sub_FrmName,'SubForm Module'
        EXEC ut_zzVBA ANYSUB,tpl_Sub010 ,'SubForm Template'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYSUB) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'sub_FrmName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'SubForm Module'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'CTCTCTCTCT'
        IF LEN(@StdTx3) = 0 SET @StdTx3 = 'IIIIIIIIII'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX FMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYBAS = Build module:  basBasName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYBAS,basBasName,'Manage AnyBase'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYBAS) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'basBasName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage AnyBase'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX BMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYCLS = Build module:  clsClsName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYCLS,DataFeedChangeLog,'Manage CanalOne Data Feed Change Log',cls
        EXEC ut_zzVBA ANYCLS,clsClsName,'Manage Any Class',cls
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYCLS) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsClsName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Any Class'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'cls'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX BTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX XTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SQC    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SNC    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------
    -- ANYRPT = Build module:  rpt_RptNam
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA ANYRPT,rpt_RptNam,'Manage Report Values'
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecANYRPT) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'rpt_RptNam'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage Report Values'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,''     ,@BldLST,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,@InpTxt,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MACMOD = MSAccess Class Module:  clsTableName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MACMOD,zzz_Test01,'Manage Test01 Records',cls
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MACMOD,DataFeedChangeLog,'Manage DataFeedChangeLog Records',cls
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMACMOD) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpObj) = 0 SET @InpObj = 'TableName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Manage TableName Records'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'cls'
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX SHD    ,@InpObj,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldCOD,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX BTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX XTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SNC    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX CLSMIT ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld     Inp     Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MLV    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSCIN ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MTP    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSCLP ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MLP    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSCLR ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MLC    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,1  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSXST ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MRV    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,2  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSANW ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MAD    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,1  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSUPD ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MUP    ,@InpObj,'' ,'' ,'' ,'' ,'' ,0  ,1  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX CLSDEL ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        RETURN
    ------------------------------------------------------------------------------------------------
    -- MACTST = MSAccess Class Test:  basTableName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MACTST,zzz_Test01,'Manage Test01 Records',cls
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MACTST,DataFeedChangeLog,'Manage DataFeedChangeLog Records',cls
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMACTST) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpObj) = 0 SET @InpObj = 'TableName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Test TableName Class'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'cls'
        --------------------------------------------------------------------------------------------
        SET @DspObj = @StdTx2+@InpObj
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX SHD    ,@InpObj,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldCOD,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX TSTCPP ,@InpObj,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld     Inp     Fmt Oup Dsc Dsp     Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ MTF    ,@InpObj,'' ,'' ,'' ,@DspObj,'' ,0  ,0  ,2  ,0  ,0  ,0  ,0  ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        RETURN
    ------------------------------------------------------------------------------------------------
    -- MAFPOP = MSAccess Form Module:  PopUp Form Using clsTableName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MAFPOP,zzz_Test01,'Modify Test01 Records',pop
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MAFPOP,DataFeedChangeLog,'Modify DataFeedChangeLog Records',pop
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMAFPOP) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpObj) = 0 SET @InpObj = 'TableName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Modify TableName Records'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'pop'
        --------------------------------------------------------------------------------------------
        SET @DspObj = @StdTx2+@InpObj
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX SHD    ,@InpObj,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldCOD,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX BTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX XTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SNC    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX POPMIT ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        RETURN
    ------------------------------------------------------------------------------------------------
    -- MAFRVW = MSAccess Form Module:  Review Form Using clsTableName
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MAFRVW,zzz_Test01,'Review Test01 Records',lst
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBA MAFRVW,DataFeedChangeLog,'Review DataFeedChangeLog Records',lst
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecMAFRVW) BEGIN
    ------------------------------------------------------------------------------------------------
        IF LEN(@InpObj) = 0 SET @InpObj = 'TableName'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Review TableName Records'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'rvw'
        --------------------------------------------------------------------------------------------
        SET @DspObj = @StdTx2+@InpObj
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1     Tx2     Tx3     Trn Idn Erm
        --------------------------------------------------------------------------------------------
        EXEC ut_zzVBX SHD    ,@InpObj,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@BldCOD,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX BTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX XTX    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX SNC    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX ORV    ,''     ,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        EXEC ut_zzVBX LSTMIT ,@InpObj,0  ,0  ,0  ,0  ,''     ,''     ,''     ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- CLSTCN = Build module:  clsTxtCon
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA CLSTCN,clsTxtCon,'Provide common constants for code generation',tcn
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecCLSTCN) BEGIN
        IF LEN(@InpTxt) = 0 SET @InpTxt = 'clsTxtCon'
        IF LEN(@StdTx1) = 0 SET @StdTx1 = 'Code Generation Constants'
        IF LEN(@StdTx2) = 0 SET @StdTx2 = 'tcn'
        SET @ClsLst = ''
        --   ut_zzVBX Oup     Stx    Lft Spc Ttl Bat Tx1      Tx2     Tx3     Trn Idn Erm
        EXEC ut_zzVBX CMC    ,@InpTxt,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,'' ,0  ,0  ,0
        EXEC ut_zzVBX CEV    ,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        EXEC ut_zzVBX @BldLST,''     ,0  ,0  ,0  ,0  ,@StdTx1,@StdTx2,@StdTx3,0  ,0  ,0
        RETURN
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

 
    ------------------------------------------------------------------------------------------------
    -- XTDXYR = Extend Tax Year
    -- XTDXPD = Extend Tax Period
    -- XTDXMN = Extend Tax Month
    -- XTDXAY = Extend Active Tax Year
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBA XTDXYR,trx_DfnDbs
        EXEC ut_zzVBA XTDXPD,trx_DfnDbs
        EXEC ut_zzVBA XTDXMN,trx_DfnDbs
        EXEC ut_zzVBA XTDXAY,trx_DfnDbs
    ----------------------------------------------------------------------------------------------*/
   END ELSE IF @BldCOD IN (@SecXTDXYR,@SecXTDXPD,@SecXTDXMN,@SecXTDXAY) BEGIN
        SET @LftMrg = 0
        SET @IncSpc = 0
        SET @IncTtl = 0
        SET @IncTrn = 0
        --------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup     Stx     Lft     Spc     Ttl     Bat     Tx1     Tx2     Tx3     Trn     Idn     Erm
        EXEC ut_zzVBX @BldLST,@InpObj,@LftMrg,@IncSpc,@IncTtl,@IncBat,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncErm
        RETURN
    ------------------------------------------------------------------------------------------------
 
 
    ------------------------------------------------------------------------------------------------
    -- Combined output codes
    ------------------------------------------------------------------------------------------------
    END ELSE BEGIN
        SET @LftMrg = 1
        SET @IncSpc = 2
        SET @IncTtl = 1
        --------------------------------------------------------------------------------------------
        --   ut_zzVBJ Obj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod
        EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
    END
    ------------------------------------------------------------------------------------------------

 
    --##############################################################################################
 
END
GO
 
/*--(LSP)-------------------------------------------------------------------------------------------
 
    --  (Oup: PHL SIG UTL URP LSP PVL)
 
    --   ut_zzUTL Soj      Oup Dbg Obj Dsc                              Dsp Par Cod Exm Tcd
    EXEC ut_zzUTL ut_zzVBA,UTL,1  ,'' ,'List basic VBA code statements','' ,"
    @BldLST varchar(2000) = '',             -- Build code list (comma delimited; see below)
    @InpTxt varchar(max)  = '',             -- Input Object text (comma delimited)
    @StdTx1 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx2 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx3 varchar(max)  = '',             -- Miscellaneous text value
    @StdFlg tinyint       = 0,              -- Miscellaneous flag value
    @StdCnt int           = 0               -- Miscellaneous count value
    ","
        PRINT '    LCP    = Lookup Constants: PKey'
        PRINT '    LCC    = Lookup Constants: Code'
        PRINT '    LCN    = Lookup Constants: Name'
        PRINT '    LCX    = Lookup Constants: CmdTxt'
        PRINT ''
        PRINT '    LPP    = Lookup Properties: PKey'
        PRINT '    LPC    = Lookup Properties: Code'
        PRINT '    LPN    = Lookup Properties: Name'
        PRINT '    LPX    = Lookup Properties: CmdTxt'
        PRINT ''
        PRINT '    MLV    = Module Level Variables'
        PRINT '    MLP    = Module Level Properties - LET/GET'
        PRINT '    MLL    = Module Level Properties - LET'
        PRINT '    MLG    = Module Level Properties - GET'
        PRINT ''
        PRINT '    MAV    = Module Level AssignSQL (from variables)'
        PRINT '    MAC    = Module Level AssignSQL (from controls)'
        PRINT '    MAN    = Module Level AssignSQL (from nulls)'
        PRINT ''
        PRINT '    MRV    = Module Level ReadSQL (into variables)'
        PRINT '    MRC    = Module Level ReadSQL (into controls)'
        PRINT ''
        PRINT '    MCV    = Module Level ClearSQL (variables)'
        PRINT '    MCC    = Module Level ClearSQL (controls)'
        PRINT ''
        PRINT '    MIF    = Module IF Criteria'
        PRINT ''
        PRINT '    DTV    = Declare fields column variables'
        PRINT '    ITV    = Initialize fields column variables'
        PRINT '    ATV    = Assign fields column variables'
        PRINT ''
        PRINT '    DSV    = Declare statement column variables'
        PRINT '    ISV    = Initialize statement column variables'
        PRINT '    ASV    = Assign statement column variables'
        PRINT ''
        PRINT '    DKV    = Declare primary key column variables'
        PRINT '    IKV    = Initialize primary key column variables'
        PRINT ''
        PRINT '    AKV    = Assign primary key column variables'
        PRINT '    DGV    = Declare function parameter variables'
        PRINT '    IGV    = Initialize parameter variables'
        PRINT ''
        PRINT '    AGV    = Assign parameter variables'
        PRINT ''
        PRINT '    RSV    = Recordset variables'
        PRINT '    RSL    = Recordset standard loop'
        PRINT '    RSW    = Recordset write loop'
        PRINT ''
        PRINT '    BASWTL = Build module:  bas_UtlWTL'
        PRINT '    CLSWTL = Build module:  clsUtlWTL'
        PRINT ''
        PRINT '    BASAPC = Build module:  bas_AppCons'
        PRINT '    BASAPF = Build module:  bas_AppFunc'
        PRINT '    BASAPT = Build module:  bas_AppTest'
        PRINT '    BASAPV = Build module:  bas_AppVars'
        PRINT ''
        PRINT '    BASGLB = Build module:  bas_Global'
        PRINT '    BASIMX = Build module:  bas_ImpExp'
        PRINT '    BASTST = Build module:  bas_Test01'
        PRINT '    BASTBM = Build module:  bas_TblMnt'
        PRINT ''
        PRINT '    UTLASC = Build module:  clsUtlASC'
        PRINT '    UTLFMT = Build module:  clsUtlFMT'
        PRINT '    UTLVBG = Build module:  clsUtlVBG'
        PRINT '    UTLWSH = Build module:  clsUtlWSH'
        PRINT '    UTLWTX = Build module:  clsUtlWTX'
        PRINT ''
        PRINT '    GENGLB = Build module:  vba_Global'
        PRINT '    GENSTD = Build module:  vbaGenSTD'
        PRINT '    GENJET = Build module:  vbaGenJET'
        PRINT ''
        PRINT '    GEN_IT = Build module:  vbaGen_IT'
        PRINT '    GENFRM = Build module:  vbaGenFRM'
        PRINT '    GENCTL = Build module:  vbaGenCTL'
        PRINT '    GENTBL = Build module:  vbaGenTBL'
        PRINT '    GENPRP = Build module:  vbaGenPRP'
        PRINT '    GENCMD = Build module:  vbaGenCMD'
        PRINT '    GENRPT = Build module:  vbaGenRPT'
        PRINT '    GENPTH = Build module:  vbaGenPTH'
        PRINT '    GENSQL = Build module:  vbaGenSQL'
        PRINT '    GENSBY = Build module:  vbaGenSBY'
        PRINT '    GENGBY = Build module:  vbaGenGBY'
        PRINT '    GENSLO = Build module:  vbaGenSLO'
        PRINT ''
        PRINT '    CLSAPC = Build module:  clsAppCons'
        PRINT '    CLSAPV = Build module:  clsAppVals'
        PRINT ''
        PRINT '    BASCMG = Build module:  bas_CmgCons'
        PRINT '    CLSCMG = Build module:  clsCtlMgr'
        PRINT ''
        PRINT '    REGTBL = Build module:  clsRegTBL'
        PRINT '    REGPRP = Build module:  clsRegPRP'
        PRINT '    REGCMD = Build module:  clsRegCMD'
        PRINT '    REGRPT = Build module:  clsRegRPT'
        PRINT '    REGPTH = Build module:  clsRegPTH'
        PRINT '    REGSRC = Build module:  clsRegSRC'
        PRINT ''
        PRINT '    SQLSTM = Build module:  clsSqlSTM'
        PRINT '    SQLOBY = Build module:  clsSqlOBY'
        PRINT '    RUNWHR = Build module:  clsRunWHR'
        PRINT ''
        PRINT '    RUNCMD = Build module:  clsRunCMD'
        PRINT '    RUNCMM = Build module:  Run_Process_0000 (CALL cls_Method)'
        PRINT '    RUNCMF = Build module:  Run_Process_0000 (OPEN frm_FrmNam)'
        PRINT ''
        PRINT '    RUNRPT = Build module:  clsRunRPT'
        PRINT '    RUNRPR = Build module:  Run_Report_0000'
        PRINT '    RUNRPX = Build module:  Print_rpt_ReportName'
        PRINT ''
        PRINT '    RUNUSP = Build module:  clsRunUSP'
        PRINT '    RUNUSR = Build module:  Run_Process_0000 (EXEC PROC)'
        PRINT '    RUNUSF = Build module:  Run_Process_0000 (OPEN FORM)'
        PRINT '    RUNUSX = Build SProcs:  Execute_usp_ProcedureName'
        PRINT ''
        PRINT '    RUNRST = Build module:  clsRunRST'
        PRINT '    RUNSQL = Build module:  clsRunSQL'
        PRINT '    RUNSBY = Build module:  clsRunSBY'
        PRINT '    RUNGBY = Build module:  clsRunGBY'
        PRINT ''
        PRINT '    FRMCLR = Build module:  frm_FrmName'
        PRINT '    FRMLNK = Build module:  sys_LinkAPP'
        PRINT ''
        PRINT '    RPTNAR = Build module:  tpl_NARROW'
        PRINT '    RPTWID = Build module:  tpl_WIDE'
        PRINT ''
        PRINT '    ANYFRM = Build module:  frm_FrmName'
        PRINT '    ANYTAB = Build module:  frm_FrmName'
        PRINT '    ANYLST = Build module:  lst_FrmName'
        PRINT '    ANYPOP = Build module:  pop_FrmName'
        PRINT '    ANYSUB = Build module:  sub_FrmName'
        PRINT '    ANYBAS = Build module:  basBasName'
        PRINT '    ANYCLS = Build module:  clsClsName'
        PRINT '    ANYRPT = Build module:  rpt_RptNam'
        PRINT ''
        PRINT '    CLSTCN = Build module:  clsTxtCon'
        PRINT ''
        PRINT '    XTDXYR = Extend Tax Year'
        PRINT '    XTDXPD = Extend Tax Period'
        PRINT '    XTDXMN = Extend Tax Month'
        PRINT '    XTDXAY = Extend Active Tax Year'
    ","
        PRINT '    --   ut_zzVBA Bld    Obj         Tx1 Tx2 Tx3'
        PRINT '    EXEC ut_zzVBA BldCod,InputObject,'''''''' ,'''''''' ,'''''''''
        PRINT '    '
        PRINT '    --   ut_zzVBA Bld     Obj     Tx1     Tx2     Tx3'
        PRINT '    EXEC ut_zzVBA @BldLST,@InpTxt,@StdTx1,@StdTx2,@StdTx3'
    ","
    "
 
--------------------------------------------------------------------------------------------------*/
