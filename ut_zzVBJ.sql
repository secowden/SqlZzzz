/*##################################################################################################
# Name:                                                                                   2016-08-01
#    ut_zzVBJ
#
####################################################################################################
# Description: 
#    Build core VBA code logic statements
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
IF OBJECT_ID('dbo.ut_zzVBJ') IS NOT NULL DROP PROCEDURE dbo.ut_zzVBJ  -- (IXS)
GO
----------------------------------------------------------------------------------------------------
CREATE PROCEDURE dbo.ut_zzVBJ (    -- (PSB)
    @BldLST varchar(2000) = '',             -- Build code list (comma delimited; see below)
    @InpObj sysname       = '',             -- Input object name
    @OupFmt varchar(11)   = '',             -- Output format (SEL,INS,etc)
    @OupObj sysname       = '',             -- Output object name
    @OupDsc sysname       = '',             -- Output object description
    @DspObj sysname       = '',             -- Display object (replaces @InpObj)
    @DefTyp varchar(11)   = '',             -- Default object type code
    @SqlExc tinyint       = 0,              -- Execute the dynamic SQL statement
    @LftMrg smallint      = 1,              -- Increase left margin (4x)
    @IncSpc tinyint       = 2,              -- Include space(s) before the header
    @IncTtl tinyint       = 1,              -- Include code segment titles
    @IncHdr tinyint       = 1,              -- Include header lines/text
    @IncTpl tinyint       = 0,              -- Include templates
    @IncMsg tinyint       = 0,              -- Include information message
    @IncDrp tinyint       = 0,              -- Include drop statement
    @IncAdd tinyint       = 1,              -- Include add statement
    @IncBat tinyint       = 1,              -- Include batch GO statement
    @IncDat tinyint       = 1,              -- Include data insert statements
    @SelStm varchar(100)  = '',             -- SELECT statement (DISTINCT, TOP, etc)
    @SetLst varchar(2000) = '',             -- SET Column = Value list (colon delimited)
    @JnnLst varchar(2000) = '',             -- JOIN list (colon delimited)
    @WhrLst varchar(2000) = '',             -- WHERE list (colon delimited)
    @GbyLst varchar(2000) = '',             -- GROUP BY list (colon delimited)
    @HavLst varchar(2000) = '',             -- HAVING list (colon delimited)
    @ObyLst varchar(2000) = '',             -- ORDER BY list (comma delimited)
    @LkpLst varchar(2000) = '',             -- Lookup parameters (comma delimited list)
    @StdTx1 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx2 varchar(max)  = '',             -- Miscellaneous text value
    @StdTx3 varchar(max)  = '',             -- Miscellaneous text value
    @IncTrn tinyint       = 0,              -- Include transaction logic
    @IncIdn tinyint       = 1,              -- Include identity column logic
    @IncDsb tinyint       = NULL,           -- Include record disabled column
    @IncDlt tinyint       = NULL,           -- Include record delflag column
    @IncLok tinyint       = NULL,           -- Include record locking columns
    @IncAud tinyint       = NULL,           -- Include record auditing columns
    @IncHst tinyint       = NULL,           -- Include record history columns
    @IncMod tinyint       = NULL            -- Include record modified columns
) AS BEGIN
    ------------------------------------------------------------------------------------------------
    -- Signature Template  (PIF)
    /*----------------------------------------------------------------------------------------------
    --   ut_zzVBJ BldObjFmtOupDscDspTypSqxLftSpcTtlHdrTplMsgDrpAddBatDatStmSetJnnWhrGbyHavObyLkpTx1Tx2Tx3TrnIdnDsbDltLokAudHstMod
    EXEC ut_zzVBJ ,'','','','','','',0,1,2,1,1,0,0,0,1,1,1,'','','','','','','','','','','',0,1,NULL,NULL,NULL,NULL,NULL,NULL
    --   ut_zzVBJ BldObjFmtOupDscDspTypSqxLftSpcTtlHdrTplMsgDrpAddBatDatStmSetJnnWhrGbyHavObyLkpTx1Tx2Tx3TrnIdnDsbDltLokAudHstMod
    EXEC ut_zzVBJ @BldLST,@InpObj,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod
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
    DECLARE @CurUSP    varchar(30)     ; SET @CurUSP = 'ut_zzVBJ'
    DECLARE @CurREF    varchar(30)     ; SET @CurREF = 'dbo.ut_zzVBJ'
    DECLARE @CurDSC    varchar(100)    ; SET @CurDSC = 'Build core VBA code logic statements'
    DECLARE @CurCAT    varchar(10)     ; SET @CurCAT = 'GEN'
    DECLARE @CurFMT    char(3)         ; SET @CurFMT = RIGHT(@CurUSP,3)
    ------------------------------------------------------------------------------------------------
    -- Manage Execution Flags
    ------------------------------------------------------------------------------------------------
    DECLARE @CurEXC    tinyint         ; SET @CurEXC = 1                                          -- Execution: 0=Disabled 1=Enabled
    DECLARE @CurDBG    tinyint         ; SET @CurDBG = 0                                          -- DebugMode: 0=Disabled 1=Enabled
    ------------------------------------------------------------------------------------------------
    DECLARE @DbgLvl    varchar(9)      ; SET @DbgLvl = ''                                         -- DebugText: Customize for Debug Tracking
    DECLARE @DbgFlg    tinyint         ; SET @DbgFlg = @CurDBG                                    -- DebugFlag: Backward Compatibility; Assign @CurDBG
    DECLARE @DbgRsj    tinyint         ; SET @DbgRsj = 0                                          -- DebugMode: 0=Disabled 1=Enabled; Debug ut_zzRSJ Output
    ------------------------------------------------------------------------------------------------
    SET @DbgFlg = CASE WHEN @BldLST = 'ZZZ' AND LEN(@InpObj) > 4 THEN 1 ELSE @DbgFlg END
    ------------------------------------------------------------------------------------------------
    -- Display text based on Debug/Execution modes
    ------------------------------------------------------------------------------------------------
    IF @CurDBG = 1 OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT
        LEFT(@CurUSP ,15) AS CurUSP,
        LEFT(@BldLST ,30) AS BldLST,
        LEFT(@InpObj ,30) AS InpObj,
        LEFT(@OupFmt ,03) AS OupFmt,
        LEFT(@OupObj ,30) AS OupObj,
        LEFT(@OupDsc ,30) AS OupDsc,
        LEFT(@DspObj ,30) AS DspObj,
        @DefTyp           AS DefTyp,
        @SqlExc           AS SqlExc,
        @LftMrg           AS LftMrg,
        @IncSpc           AS IncSpc,
        @IncTtl           AS IncTtl,
        @IncHdr           AS IncHdr,
        @IncTpl           AS IncTpl,
        @IncMsg           AS IncMsg,
        @IncDrp           AS IncDrp,
        @IncAdd           AS IncAdd,
        @IncBat           AS IncBat,
        @IncDat           AS IncDat,
        LEFT(@SelStm ,30) AS SelStm,
        LEFT(@SetLst ,30) AS SetLst,
        LEFT(@JnnLst ,30) AS JnnLst,
        LEFT(@WhrLst ,30) AS WhrLst,
        LEFT(@GbyLst ,30) AS GbyLst,
        LEFT(@HavLst ,30) AS HavLst,
        LEFT(@ObyLst ,30) AS ObyLst,
        LEFT(@LkpLst ,30) AS LkpLst,
        LEFT(@StdTx1 ,30) AS StdTx1,
        LEFT(@StdTx2 ,30) AS StdTx2,
        LEFT(@StdTx3 ,30) AS StdTx3,
        @IncTrn           AS IncTrn,
        @IncIdn           AS IncIdn,
        @IncDsb           AS IncDsb,
        @IncDlt           AS IncDlt,
        @IncLok           AS IncLok,
        @IncAud           AS IncAud,
        @IncHst           AS IncHst,
        @IncMod           AS IncMod
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
        PRINT "    @InpObj sysname       = '',             -- Input object name"
        PRINT "    @OupFmt varchar(11)   = '',             -- Output format (SEL,INS,etc)"
        PRINT "    @OupObj sysname       = '',             -- Output object name"
        PRINT "    @OupDsc sysname       = '',             -- Output object description"
        PRINT "    @DspObj sysname       = '',             -- Display object (replaces @InpObj)"
        PRINT "    @DefTyp varchar(11)   = '',             -- Default object type code"
        PRINT "    @SqlExc tinyint       = 0,              -- Execute the dynamic SQL statement"
        PRINT "    @LftMrg smallint      = 1,              -- Increase left margin (4x)"
        PRINT "    @IncSpc tinyint       = 2,              -- Include space(s) before the header"
        PRINT "    @IncTtl tinyint       = 1,              -- Include code segment titles"
        PRINT "    @IncHdr tinyint       = 1,              -- Include header lines/text"
        PRINT "    @IncTpl tinyint       = 0,              -- Include templates"
        PRINT "    @IncMsg tinyint       = 0,              -- Include information message"
        PRINT "    @IncDrp tinyint       = 0,              -- Include drop statement"
        PRINT "    @IncAdd tinyint       = 1,              -- Include add statement"
        PRINT "    @IncBat tinyint       = 1,              -- Include batch GO statement"
        PRINT "    @IncDat tinyint       = 1,              -- Include data insert statements"
        PRINT "    @SelStm varchar(100)  = '',             -- SELECT statement (DISTINCT, TOP, etc)"
        PRINT "    @SetLst varchar(2000) = '',             -- SET Column = Value list (colon delimited)"
        PRINT "    @JnnLst varchar(2000) = '',             -- JOIN list (colon delimited)"
        PRINT "    @WhrLst varchar(2000) = '',             -- WHERE list (colon delimited)"
        PRINT "    @GbyLst varchar(2000) = '',             -- GROUP BY list (colon delimited)"
        PRINT "    @HavLst varchar(2000) = '',             -- HAVING list (colon delimited)"
        PRINT "    @ObyLst varchar(2000) = '',             -- ORDER BY list (comma delimited)"
        PRINT "    @LkpLst varchar(2000) = '',             -- Lookup parameters (comma delimited list)"
        PRINT "    @StdTx1 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @StdTx2 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @StdTx3 varchar(max)  = '',             -- Miscellaneous text value"
        PRINT "    @IncTrn tinyint       = 0,              -- Include transaction logic"
        PRINT "    @IncIdn tinyint       = 1,              -- Include identity column logic"
        PRINT "    @IncDsb tinyint       = NULL,           -- Include record disabled column"
        PRINT "    @IncDlt tinyint       = NULL,           -- Include record delflag column"
        PRINT "    @IncLok tinyint       = NULL,           -- Include record locking columns"
        PRINT "    @IncAud tinyint       = NULL,           -- Include record auditing columns"
        PRINT "    @IncHst tinyint       = NULL,           -- Include record history columns"
        PRINT "    @IncMod tinyint       = NULL            -- Include record modified columns"
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT 'Build Codes:'
        PRINT ''
        PRINT '    PSH    = Push margin 4 spaces right'
        PRINT '    PUL    = Pull margin 4 spaces left'
        PRINT '    LM0    = Set left margin to zero'
        PRINT '    LM1    = Set left margin to one'
        PRINT '    LM2    = Set left margin to two'
        PRINT '    RWP    = Set report width Portrait'
        PRINT '    RWL    = Set report width Landscape'
        PRINT '    '
        PRINT '    LSG    = Set lines for single lines'
        PRINT '    LDB    = Set lines for double lines'
        PRINT '    LPD    = Set lines for pound  lines'
        PRINT '    HSG    = Set header for single lines'
        PRINT '    HDB    = Set header for double lines'
        PRINT '    HPD    = Set header for pound  lines'
        PRINT '    '
        PRINT '    SLN    = Print single line'
        PRINT '    DLN    = Print double line'
        PRINT '    ALN    = Print asterick line'
        PRINT '    PLN    = Print pound line'
        PRINT '    MLN    = Print ampersand line'
        PRINT '    TLN    = Print tilde line'
        PRINT '    '
        PRINT '    LN0    = Print current LinSpc'
        PRINT '    LN1    = Print empty lines (1)'
        PRINT '    LN2    = Print empty lines (2)'
        PRINT '    PLP    = Pound line (prefixed)'
        PRINT '    ALP    = AtSign line (prefixed)'
        PRINT '    PHB    = Print begin header (previously set)'
        PRINT '    PHE    = Print end header (previously set)'
        PRINT '    '
        PRINT '    T0N    = Set Title: Space 0 Lines N'
        PRINT '    T0Y    = Set Title: Space 0 Lines Y'
        PRINT '    T1N    = Set Title: Space 1 Lines N'
        PRINT '    T1Y    = Set Title: Space 1 Lines Y'
        PRINT '    T2N    = Set Title: Space 2 Lines N'
        PRINT '    T2Y    = Set Title: Space 2 Lines Y'
        PRINT '    '
        PRINT '    IST    = Toggle   Include space(s) before the header'
        PRINT '    ISR    = Reset    Include space(s) before the header'
        PRINT '    IS0    = Set OFF  Include space(s) before the header'
        PRINT '    IS1    = Set Opt1 Include space(s) before the header'
        PRINT '    IS2    = Set Opt2 Include space(s) before the header'
        PRINT '    '
        PRINT '    ITT    = Toggle   Include code segment titles'
        PRINT '    ITR    = Reset    Include code segment titles'
        PRINT '    IT0    = Set OFF  Include code segment titles'
        PRINT '    IT1    = Set Opt1 Include code segment titles'
        PRINT '    IT2    = Set Opt2 Include code segment titles'
        PRINT '    IT3    = Set Opt3 Include code segment titles'
        PRINT '    '
        PRINT '    IPT    = Toggle   Include templates'
        PRINT '    IPR    = Reset    Include templates'
        PRINT '    IP0    = Set OFF  Include templates'
        PRINT '    IP1    = Set Opt1 Include templates'
        PRINT '    IP2    = Set Opt2 Include templates'
        PRINT '    '
        PRINT '    IGT    = Toggle   Include debug logic'
        PRINT '    IGR    = Reset    Include debug logic'
        PRINT '    IG0    = Disable  Include debug logic'
        PRINT '    IG1    = Enable   Include debug logic'
        PRINT '    '
        PRINT '    IFT    = Toggle   Include information message'
        PRINT '    IFR    = Reset    Include information message'
        PRINT '    IF0    = Disable  Include information message'
        PRINT '    IF1    = Enable   Include information message'
        PRINT '    '
        PRINT '    IQT    = Toggle   Include error message'
        PRINT '    IQR    = Reset    Include error message'
        PRINT '    IQ0    = Disable  Include error message'
        PRINT '    IQ1    = Enable   Include error message'
        PRINT '    '
        PRINT '    IVT    = Toggle   Include separator line between objects'
        PRINT '    IVR    = Reset    Include separator line between objects'
        PRINT '    IV0    = Disable  Include separator line between objects'
        PRINT '    IV1    = Enable   Include separator line between objects'
        PRINT '    '
        PRINT '    IDT    = Toggle   Include drop statement'
        PRINT '    IDR    = Reset    Include drop statement'
        PRINT '    ID0    = Disable  Include drop statement'
        PRINT '    ID1    = Enable   Include drop statement'
        PRINT '    '
        PRINT '    IBT    = Toggle   Include batch GO statement'
        PRINT '    IBR    = Reset    Include batch GO statement'
        PRINT '    IB0    = Disable  Include batch GO statement'
        PRINT '    IB1    = Enable   Include batch GO statement'
        PRINT '    '
        PRINT '    INT    = Toggle   Include batch GO statement'
        PRINT '    INR    = Reset    Include batch GO statement'
        PRINT '    IN0    = Disable  Include batch GO statement'
        PRINT '    IN1    = Enable   Include batch GO statement'
        PRINT '    '
        PRINT '    IMT    = Toggle   Include permissions statements'
        PRINT '    IMR    = Reset    Include permissions statements'
        PRINT '    IM0    = Disable  Include permissions statements'
        PRINT '    IM1    = Enable   Include permissions statements'
        PRINT '    '
        PRINT '    IIT    = Toggle   Include identity column logic'
        PRINT '    IIR    = Reset    Include identity column logic'
        PRINT '    II0    = Disable  Include identity column logic'
        PRINT '    II1    = Enable   Include identity column logic'
        PRINT '    '
        PRINT '    IXT    = Toggle   Include record expires columns'
        PRINT '    IXR    = Reset    Include record expires columns'
        PRINT '    IX0    = Disable  Include record expires columns'
        PRINT '    IX1    = Enable   Include record expires columns'
        PRINT '    '
        PRINT '    OTX    = Object text (SysComments)'
        PRINT '    '
        PRINT '    DVA    = Developer action history'
        PRINT '    VLN    = Set VbaVln length'
        PRINT '    TMB    = Initialize text management objects (Begin)'
        PRINT '    TME    = Initialize text management objects (End)'
        PRINT '    '
        PRINT '    DTV    = Declare field column variables'
        PRINT '    ITV    = Initialize field column variables'
        PRINT '    ATV    = Assign standard field variables'
        PRINT '    '
        PRINT '    DSV    = Declare statement column variables'
        PRINT '    ISV    = Initialize statement column variables'
        PRINT '    ASV    = Assign statement column variables'
        PRINT '    '
        PRINT '    DKV    = Declare primary key column variables'
        PRINT '    IKV    = Initialize primary key column variables'
        PRINT '    AKV    = Assign primary key column variables'
        PRINT '    '
        PRINT '    DGV    = Declare function parameter variables'
        PRINT '    IGV    = Initialize parameter variables'
        PRINT '    AGV    = Assign parameter variables'
        PRINT '    '
        PRINT '    IRV    = Initialize recordset variables'
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
        PRINT '    RUNUSX = Run SProc commands'
        PRINT '    '
        PRINT '    SQLSTM = Build the full SQL statement'
        PRINT '    '
        PRINT '    FRMHDR = Header form'
        PRINT '    FRMDTL = Detail form'
        PRINT '    FRMLST = List form'
        PRINT '    '
        PRINT '    POPADD = PopAdd form'
        PRINT '    POPUPD = PopUpd form'
        PRINT '    '
        PRINT '    RECINS = Insert record function'
        PRINT '    RECUPD = Update record function'
        PRINT ''
        PRINT '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
        PRINT 'Example Code:'
        PRINT ''
        PRINT '    --   ut_zzVBJ Soj     Oup Fmt Obj Dsc Dsp Oup Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat Dat Stm Set Jnn Whr Gby Hav Oby Lkp Tx1 Tx2 Tx3 Trn Idn Dsb  Dlt  Lok  Aud  Hst  Mod'
        PRINT '    EXEC ut_zzVBJ ''''     ,ZZZ,   ,'''' ,'''' ,'''' ,'''' ,0  ,1  ,2  ,1  ,1  ,0  ,0  ,0  ,1  ,1  ,0  ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,0  ,1  ,NULL,NULL,NULL,NULL,NULL,NULL'
        PRINT '    '
        PRINT '    --   ut_zzVBJ Soj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod'
        PRINT '    EXEC ut_zzVBJ @InpObj,@BldLST,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod'
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
    -- Output codes tracking
    ------------------------------------------------------------------------------------------------
    SET @OupFmt = UPPER(@OupFmt)                                                                  -- Output format
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize Standard Default Objects  (ID1)                              
    ------------------------------------------------------------------------------------------------
    DECLARE @InpLst    varchar(max)    ; SET @InpLst    = @InpObj                                 -- Input object list
    DECLARE @NamFmt    varchar(3)      ; SET @NamFmt    = ''                                      -- Name format (revised from @BldCOD)
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize Output Default Objects  (ID2)                                
    ------------------------------------------------------------------------------------------------
    DECLARE @OvrOup    tinyint         ; SET @OvrOup    = 0                                       -- Output override flag
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize Code Generation Default Objects  (ID3)                       
    ------------------------------------------------------------------------------------------------
    DECLARE @OupCls    varchar(11)     ; SET @OupCls    = ''                                      -- Output class  (TBL,VEW,etc)
    DECLARE @OupAls    varchar(11)     ; SET @OupAls    = ''                                      -- Output table alias (abc,xyz,...)
    DECLARE @OupSfx    varchar(11)     ; SET @OupSfx    = ''                                      -- Output table suffix (ABC,XYZ,...)
    DECLARE @OupCpx    varchar(20)     ; SET @OupCpx    = ''                                      -- Output column prefix (Abcdef,...)
    DECLARE @OupUtl    varchar(11)     ; SET @OupUtl    = RIGHT(@CurUSP,3)                        -- Output utility (TBX,USX,etc)
    DECLARE @IncCmt    tinyint         ; SET @IncCmt    = 0                                       -- Include comment text
    DECLARE @IncDbg    tinyint         ; SET @IncDbg    = 0                                       -- Include debug logic
    DECLARE @IncErm    tinyint         ; SET @IncErm    = 0                                       -- Include error message
    DECLARE @IncSep    tinyint         ; SET @IncSep    = 1                                       -- Include separator line between objects
    DECLARE @IncTcd    tinyint         ; SET @IncTcd    = 0                                       -- Include test code
    DECLARE @IncPrm    tinyint         ; SET @IncPrm    = 0                                       -- Include permissions statements
    DECLARE @StdFlg    tinyint         ; SET @StdFlg    = 0                                       -- Miscellaneous flag value
    DECLARE @StdCnt    int             ; SET @StdCnt    = 0                                       -- Miscellaneous count value
    DECLARE @IncHtk    tinyint         ; SET @IncHtk    = NULL                                    -- Include record history tracking columns
    DECLARE @IncDim    tinyint         ; SET @IncDim    = NULL                                    -- Include record dimension columns
    DECLARE @IncFct    tinyint         ; SET @IncFct    = NULL                                    -- Include record fact columns
    DECLARE @IncUsd    tinyint         ; SET @IncUsd    = NULL                                    -- Include record used column
    DECLARE @IncLkd    tinyint         ; SET @IncLkd    = NULL                                    -- Include record locked column
    DECLARE @IncCrt    tinyint         ; SET @IncCrt    = NULL                                    -- Include record created columns
    DECLARE @IncUpd    tinyint         ; SET @IncUpd    = NULL                                    -- Include record updated columns
    DECLARE @IncExp    tinyint         ; SET @IncExp    = NULL                                    -- Include record expired columns
    DECLARE @IncDel    tinyint         ; SET @IncDel    = NULL                                    -- Include record deleted columns
    DECLARE @BldSfx    varchar(11)     ; SET @BldSfx    = RIGHT(@OupFmt,3)                        -- Output suffix
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Track Original Values  (TOV)                                            
    ------------------------------------------------------------------------------------------------
    DECLARE @OrgAdd    tinyint         ; SET @OrgAdd    = @IncAdd                                 -- Include add statement
    DECLARE @OrgAud    tinyint         ; SET @OrgAud    = @IncAud                                 -- Include record auditing columns
    DECLARE @OrgBat    tinyint         ; SET @OrgBat    = @IncBat                                 -- Include batch GO statement
    DECLARE @OrgCmt    tinyint         ; SET @OrgCmt    = @IncCmt                                 -- Include comment text
    DECLARE @OrgCrt    tinyint         ; SET @OrgCrt    = @IncCrt                                 -- Include record created columns
    DECLARE @OrgDat    tinyint         ; SET @OrgDat    = @IncDat                                 -- Include data insert statements
    DECLARE @OrgDbg    tinyint         ; SET @OrgDbg    = @IncDbg                                 -- Include debug logic
    DECLARE @OrgDel    tinyint         ; SET @OrgDel    = @IncDel                                 -- Include record deleted columns
    DECLARE @OrgDim    tinyint         ; SET @OrgDim    = @IncDim                                 -- Include record dimension columns
    DECLARE @OrgDlt    tinyint         ; SET @OrgDlt    = @IncDlt                                 -- Include record delflag column
    DECLARE @OrgDrp    tinyint         ; SET @OrgDrp    = @IncDrp                                 -- Include drop statement
    DECLARE @OrgDsb    tinyint         ; SET @OrgDsb    = @IncDsb                                 -- Include record disabled column
    DECLARE @OrgErm    tinyint         ; SET @OrgErm    = @IncErm                                 -- Include error message
    DECLARE @OrgExp    tinyint         ; SET @OrgExp    = @IncExp                                 -- Include record expired columns
    DECLARE @OrgFct    tinyint         ; SET @OrgFct    = @IncFct                                 -- Include record fact columns
    DECLARE @OrgHdr    tinyint         ; SET @OrgHdr    = @IncHdr                                 -- Include header lines/text
    DECLARE @OrgHst    tinyint         ; SET @OrgHst    = @IncHst                                 -- Include record history columns
    DECLARE @OrgHtk    tinyint         ; SET @OrgHtk    = @IncHtk                                 -- Include record history tracking columns
    DECLARE @OrgIdn    tinyint         ; SET @OrgIdn    = @IncIdn                                 -- Include identity column logic
    DECLARE @OrgLkd    tinyint         ; SET @OrgLkd    = @IncLkd                                 -- Include record locked column
    DECLARE @OrgLok    tinyint         ; SET @OrgLok    = @IncLok                                 -- Include record locking columns
    DECLARE @OrgMod    tinyint         ; SET @OrgMod    = @IncMod                                 -- Include record modified columns
    DECLARE @OrgMsg    tinyint         ; SET @OrgMsg    = @IncMsg                                 -- Include information message
    DECLARE @OrgPrm    tinyint         ; SET @OrgPrm    = @IncPrm                                 -- Include permissions statements
    DECLARE @OrgSep    tinyint         ; SET @OrgSep    = @IncSep                                 -- Include separator line between objects
    DECLARE @OrgSpc    tinyint         ; SET @OrgSpc    = @IncSpc                                 -- Include space(s) before the header
    DECLARE @OrgTcd    tinyint         ; SET @OrgTcd    = @IncTcd                                 -- Include test code
    DECLARE @OrgTpl    tinyint         ; SET @OrgTpl    = @IncTpl                                 -- Include templates
    DECLARE @OrgTrn    tinyint         ; SET @OrgTrn    = @IncTrn                                 -- Include transaction logic
    DECLARE @OrgTtl    tinyint         ; SET @OrgTtl    = @IncTtl                                 -- Include code segment titles
    DECLARE @OrgUpd    tinyint         ; SET @OrgUpd    = @IncUpd                                 -- Include record updated columns
    DECLARE @OrgUsd    tinyint         ; SET @OrgUsd    = @IncUsd                                 -- Include record used column
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
    -- Extended Construction Variables  (XCV)                                  EXEC ut_zzUTL zzz,XCV
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
    SET @X = @B+@B+''                                                                             -- Firstline initialized with blank
    SET @X = @X+@N+''                                                                             -- Next lines accumulate the text
    PRINT @X                                                                                      -- Print   the text (8000 max)
    EXEC (@X)                                                                                     -- Execute the text (unlimited)
    ----------------------------------------------------------------------------------------------*/

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Working Parameter Variables  (WPV)                                      EXEC ut_zzUTL zzz,WPV
    ------------------------------------------------------------------------------------------------
    DECLARE @WrkMrg    tinyint         ; SET @WrkMrg    = 0                                       -- Left margin quantity
    DECLARE @WrkSpc    tinyint         ; SET @WrkSpc    = 0                                       -- Left margin space increment
    DECLARE @WrkTtl    tinyint         ; SET @WrkTtl    = 0                                       -- Include code segment titles
    DECLARE @WrkBat    tinyint         ; SET @WrkBat    = 0                                       -- Include batch GO statement
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Adjust Margin Values  (AMV)                                             EXEC ut_zzUTL zzz,AMV
    ------------------------------------------------------------------------------------------------
    DECLARE @MrgInc    tinyint         ; SET @MrgInc    = 4                                       -- Left margin space increment
    DECLARE @StdLen    tinyint         ; SET @StdLen    = 100                                     -- Standard line length
    ------------------------------------------------------------------------------------------------
    DECLARE @LftWid    tinyint         ; SET @LftWid    = @LftMrg * @MrgInc                       -- Left margin space length Beg
    DECLARE @LftLen    tinyint         ; SET @LftLen    = @StdLen - @LftWid                       -- Left length
    ------------------------------------------------------------------------------------------------
    DECLARE @StmMrg    smallint        ; SET @StmMrg    = @LftMrg+1                               -- Code statement margin
    DECLARE @StmWid    smallint        ; SET @StmWid    = @StmMrg * @MrgInc                       -- Statement margin space length
    DECLARE @StmLen    smallint        ; SET @StmLen    = @StdLen - @StmWid                       -- Statement line length
    ------------------------------------------------------------------------------------------------
    DECLARE @M         varchar(50)     ; SET @M         = REPLICATE(' ', @LftWid)                 -- Left margin
    DECLARE @T         varchar(50)     ; SET @T         = REPLICATE(' ', @StmWid)                 -- Statement margin
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Display Size Variables  (DSV)                                           EXEC ut_zzUTL zzz,DSV
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
    DECLARE @LinCmt    varchar(200)    ; SET @LinCmt    = "' "
    ------------------------------------------------------------------------------------------------
    DECLARE @LinSgl    varchar(200)    ; SET @LinSgl    = "'"+REPLICATE(@SGL,@WidMn1 - @LftWid)
    DECLARE @LinDbl    varchar(200)    ; SET @LinDbl    = "'"+REPLICATE(@DBL,@WidMn1 - @LftWid)
    DECLARE @LinAst    varchar(200)    ; SET @LinAst    = "'"+REPLICATE(@AST,@WidMn1 - @LftWid)
    DECLARE @LinPnd    varchar(200)    ; SET @LinPnd    = "'"+REPLICATE(@PND,@WidMn1 - @LftWid)
    DECLARE @LinAts    varchar(200)    ; SET @LinAts    = "'"+REPLICATE(@ATS,@WidMn1 - @LftWid)
    DECLARE @LinTld    varchar(200)    ; SET @LinTld    = "'"+REPLICATE(@TLD,@WidMn1 - @LftWid)
    DECLARE @LinBng    varchar(200)    ; SET @LinBng    = "'"+REPLICATE(@BNG,@WidMn1 - @LftWid)
    ------------------------------------------------------------------------------------------------
    DECLARE @HdrBeg    varchar(200)    ; SET @HdrBeg    = ''
    DECLARE @HdrEnd    varchar(200)    ; SET @HdrEnd    = ''
    DECLARE @HdrCmt    varchar(200)    ; SET @HdrCmt    = ''
    DECLARE @HdrSep    varchar(200)    ; SET @HdrSep    = ''
    DECLARE @LinWid    smallint        ; SET @LinWid    = LEN(@LinSgl)
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Extended Utility Variables  (XUV)                                       EXEC ut_zzUTL zzz,XUV
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
    --  Working Report Variables  (WRV)                                        EXEC ut_zzUTL zzz,WRV
    ------------------------------------------------------------------------------------------------
    DECLARE @RptLin    varchar(500)    ; SET @RptLin    = ''                                      -- Report line
    DECLARE @RptFil    varchar(100)    ; SET @RptFil    = ''                                      -- Report filename
    DECLARE @RptMsg    varchar(100)    ; SET @RptMsg    = ''                                      -- Report message
    DECLARE @RptFlt    varchar(100)    ; SET @RptFlt    = ''                                      -- Report filter  text
    DECLARE @RptOrd    varchar(100)    ; SET @RptOrd    = ''                                      -- Report orderby text
    DECLARE @RptTtl    varchar(500)    ; SET @RptTtl    = ''                                      -- Report title
    DECLARE @RptHdr    varchar(4000)   ; SET @RptHdr    = ''                                      -- Report header
    DECLARE @RptCmd    varchar(4000)   ; SET @RptCmd    = ''                                      -- Report command
    DECLARE @RptFtr    varchar(4000)   ; SET @RptFtr    = ''                                      -- Report footer
    DECLARE @RptCnt    int             ; SET @RptCnt    = 0                                       -- Report record count (numeric)
    DECLARE @RptCnx    varchar(100)    ; SET @RptCnx    = ''                                      -- Report record count (text)
    DECLARE @RptTsz    dec(15,2)       ; SET @RptTsz    = 0                                       -- Report table size   (numeric)
    DECLARE @RptTsx    varchar(100)    ; SET @RptTsx    = ''                                      -- Report table size   (text)
    DECLARE @RptNul    varchar(20)     ; SET @RptNul    = ''                                      -- Report NULL display value
    DECLARE @RptDtm    datetime        ; SET @RptDtm    = GETDATE()                               -- Report Date/Time value
    DECLARE @RptDat    varchar(10)     ; SET @RptDat    = CONVERT(varchar(10),@RptDtm,121)        -- 121 yyyy-mm-dd hh:mi:ss.mmm | 110 = mm-dd-yyyy | 101 mm/dd/yyyy
    DECLARE @RptTim    varchar(05)     ; SET @RptTim    = CONVERT(varchar(08),@RptDtm,114)        -- 114 (05/08/12) hh:mi:ss:mmm
    DECLARE @RptDtx    varchar(20)     ; SET @RptDtx    = @RptDat+' '+@RptTim                 -- Report Date+Time
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Get Environment Values  (GEV)                                           EXEC ut_zzUTL zzz,GEV
    ------------------------------------------------------------------------------------------------
    DECLARE @PrjCod    varchar(3)      ; SET @PrjCod    = ''                                      -- Project Code
    DECLARE @DvpNam    varchar(12)     ; SET @DvpNam    = ''                                      -- Developer name
    DECLARE @ClnNam    varchar(30)     ; SET @ClnNam    = ''                                      -- Client Name
    DECLARE @WksNam    varchar(20)     ; SET @WksNam    = ''                                      -- Workstation Name
    DECLARE @SinNam    sysname         ; SET @SinNam    = ''                                      -- Instance Name
    DECLARE @SrvNam    sysname         ; SET @SrvNam    = ''                                      -- Server Name
    DECLARE @DbsNam    sysname         ; SET @DbsNam    = ''                                      -- Database Name
    DECLARE @DbsPfx    sysname         ; SET @DbsPfx    = ''                                      -- Database Prefix
    DECLARE @ImpLvl    varchar(3)      ; SET @ImpLvl    = ''                                      -- Implementation Level
    ------------------------------------------------------------------------------------------------
    -- Assign environment values
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzENV GET,
        @PrjCod OUTPUT,
        @DvpNam OUTPUT,
        @ClnNam OUTPUT,
        @WksNam OUTPUT,
        @SinNam OUTPUT,
        @SrvNam OUTPUT,
        @DbsNam OUTPUT,
        @DbsPfx OUTPUT,
        @ImpLvl OUTPUT
    ------------------------------------------------------------------------------------------------

    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Set line space values  (SLS)
    ------------------------------------------------------------------------------------------------
    DECLARE @LinSpc    varchar(20)     ; SET @LinSpc    = NULL                                    -- Line Space
    DECLARE @PrnSpc    bit             ; SET @PrnSpc    = 0                                       -- Print Space Flag
    DECLARE @SpcCnt    smallint        ; SET @SpcCnt    = 0                                       -- Space Count
    SET @CNT = @IncSpc; WHILE @CNT > 0 BEGIN
        SET @LinSpc = ISNULL(@LinSpc,'')+@N; SET @PrnSpc = 1; SET @SpcCnt += 1; SET @CNT -= 1
    END; SET @LinSpc = ISNULL(@LinSpc,'')
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
    -- Table Category (EXEC ut_zzNAX TBLCAT,TCT)
    ------------------------------------------------------------------------------------------------
    DECLARE @TblCatSTP varchar(03)     ; SET @TblCatSTP = 'STP'                                   -- Setup
    DECLARE @TblCatLKP varchar(03)     ; SET @TblCatLKP = 'LKP'                                   -- Lookup
    DECLARE @TblCatSEC varchar(03)     ; SET @TblCatSEC = 'SEC'                                   -- Security
    DECLARE @TblCatREF varchar(03)     ; SET @TblCatREF = 'REF'                                   -- Reference
    DECLARE @TblCatTRX varchar(03)     ; SET @TblCatTRX = 'TRX'                                   -- Transaction
    DECLARE @TblCatLNK varchar(03)     ; SET @TblCatLNK = 'LNK'                                   -- Link
    DECLARE @TblCatDSS varchar(03)     ; SET @TblCatDSS = 'DSS'                                   -- DecisionSupport
    DECLARE @TblCatHIS varchar(03)     ; SET @TblCatHIS = 'HIS'                                   -- History
    DECLARE @TblCatARC varchar(03)     ; SET @TblCatARC = 'ARC'                                   -- Archive
    DECLARE @TblCatTBL varchar(03)     ; SET @TblCatTBL = 'TBL'                                   -- Table
    DECLARE @TblCatETL varchar(03)     ; SET @TblCatETL = 'ETL'                                   -- Transform
    DECLARE @TblCatDIM varchar(03)     ; SET @TblCatDIM = 'DIM'                                   -- Dimension
    DECLARE @TblCatFCT varchar(03)     ; SET @TblCatFCT = 'FCT'                                   -- Fact
    DECLARE @TblCatVBA varchar(03)     ; SET @TblCatVBA = 'VBA'                                   -- Table (VBAGen)
    DECLARE @TblCatUNK varchar(03)     ; SET @TblCatUNK = 'UNK'                                   -- Unknown
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Table Prefix (EXEC ut_zzNAX TBLPFX,TPX)
    ------------------------------------------------------------------------------------------------
    DECLARE @TblPfxSTP varchar(20)     ; SET @TblPfxSTP = 'stp_'                                  -- @TblPfxSTP + @TblBas
    DECLARE @TblPfxSTB varchar(20)     ; SET @TblPfxSTB = 'Stp'                                   -- @TblPfxSTB + @TblBas
    DECLARE @TblPfxLKP varchar(20)     ; SET @TblPfxLKP = 'lkp_'                                  -- @TblPfxLKP + @TblBas
    DECLARE @TblPfxZLK varchar(20)     ; SET @TblPfxZLK = 'zlk_'                                  -- @TblPfxZLB + @TblBas
    DECLARE @TblPfxZLB varchar(20)     ; SET @TblPfxZLB = 'zlk'                                   -- @TblPfxZLK + @TblBas
    DECLARE @TblPfxSEC varchar(20)     ; SET @TblPfxSEC = 'sec_'                                  -- @TblPfxSEC + @TblBas
    DECLARE @TblPfxREF varchar(20)     ; SET @TblPfxREF = 'ref_'                                  -- @TblPfxREF + @TblBas
    DECLARE @TblPfxTRX varchar(20)     ; SET @TblPfxTRX = 'trx_'                                  -- @TblPfxTRX + @TblBas
    DECLARE @TblPfxDAT varchar(20)     ; SET @TblPfxDAT = 'dat_'                                  -- @TblPfxDAT + @TblBas
    DECLARE @TblPfxTBB varchar(20)     ; SET @TblPfxTBB = 'tbl'                                   -- @TblPfxTBL + @TblBas
    DECLARE @TblPfxTBL varchar(20)     ; SET @TblPfxTBL = 'tbl_'                                  -- @TblPfxTBB + @TblBas
    DECLARE @TblPfxTXX varchar(20)     ; SET @TblPfxTXX = 'Tx'                                    -- @TblPfxTXX + @TblBas
    DECLARE @TblPfxENT varchar(20)     ; SET @TblPfxENT = 'ent_'                                  -- @TblPfxENT + @TblBas
    DECLARE @TblPfxAPP varchar(20)     ; SET @TblPfxAPP = 'app_'                                  -- @TblPfxAPP + @TblBas
    DECLARE @TblPfxAPX varchar(20)     ; SET @TblPfxAPX = 'apx_'                                  -- @TblPfxAPX + @TblBas
    DECLARE @TblPfxLNK varchar(20)     ; SET @TblPfxLNK = 'lnk_'                                  -- @TblPfxLNK + @TblBas
    DECLARE @TblPfxDSS varchar(20)     ; SET @TblPfxDSS = 'dss_'                                  -- @TblPfxDSS + @TblBas
    DECLARE @TblPfxHIS varchar(20)     ; SET @TblPfxHIS = 'his_'                                  -- @TblPfxHIS + @TblBas
    DECLARE @TblPfxARC varchar(20)     ; SET @TblPfxARC = 'arc_'                                  -- @TblPfxARC + @TblBas
    DECLARE @TblPfxZAR varchar(20)     ; SET @TblPfxZAR = 'zar_'                                  -- @TblPfxZAR + @TblBas
    DECLARE @TblPfxVBA varchar(20)     ; SET @TblPfxVBA = 'vba_'                                  -- @TblPfxVBA + @TblBas
    DECLARE @TblPfxLKM varchar(20)     ; SET @TblPfxLKM = 'lkm_'                                  -- @TblPfxLKM + @TblBas
    DECLARE @TblPfxLKX varchar(20)     ; SET @TblPfxLKX = 'lkx_'                                  -- @TblPfxLKX + @TblBas
    DECLARE @TblPfxDMM varchar(20)     ; SET @TblPfxDMM = 'dmm_'                                  -- @TblPfxDMM + @TblBas
    DECLARE @TblPfxDMX varchar(20)     ; SET @TblPfxDMX = 'dmx_'                                  -- @TblPfxDMX + @TblBas
    DECLARE @TblPfxDIM varchar(20)     ; SET @TblPfxDIM = 'dim_'                                  -- @TblPfxDIM + @TblBas
    DECLARE @TblPfxFTS varchar(20)     ; SET @TblPfxFTS = 'fts_'                                  -- @TblPfxFTS + @TblBas
    DECLARE @TblPfxFTM varchar(20)     ; SET @TblPfxFTM = 'ftm_'                                  -- @TblPfxFTM + @TblBas
    DECLARE @TblPfxFTX varchar(20)     ; SET @TblPfxFTX = 'ftx_'                                  -- @TblPfxFTX + @TblBas
    DECLARE @TblPfxFCT varchar(20)     ; SET @TblPfxFCT = 'fct_'                                  -- @TblPfxFCT + @TblBas
    ------------------------------------------------------------------------------------------------
    DECLARE @TblPfxALL varchar(20)     ; SET @TblPfxALL = ''                                      -- @TblPfxFCT + @TblBas
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Datatype categories (EXEC ut_zzNAX DTPCAT,DTC)
    ------------------------------------------------------------------------------------------------
    DECLARE @DtpCatBLN varchar(03)     ; SET @DtpCatBLN = 'BLN'                                   -- Boolean
    DECLARE @DtpCatNBR varchar(03)     ; SET @DtpCatNBR = 'NBR'                                   -- Numeric
    DECLARE @DtpCatDAT varchar(03)     ; SET @DtpCatDAT = 'DAT'                                   -- Date
    DECLARE @DtpCatTXT varchar(03)     ; SET @DtpCatTXT = 'TXT'                                   -- Text
    DECLARE @DtpCatBIN varchar(03)     ; SET @DtpCatBIN = 'BIN'                                   -- Binary
    DECLARE @DtpCatVRN varchar(03)     ; SET @DtpCatVRN = 'VRN'                                   -- Variant
    DECLARE @DtpCatUNK varchar(03)     ; SET @DtpCatUNK = 'UNK'                                   -- Unknown
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Datatype values (EXEC ut_zzNAX DTPVAL,DTV)
    ------------------------------------------------------------------------------------------------
    DECLARE @DtpValBLN varchar(10)     ; SET @DtpValBLN = '0'                                     -- Boolean
    DECLARE @DtpValNBR varchar(10)     ; SET @DtpValNBR = '0'                                     -- Numeric
    DECLARE @DtpValDAT varchar(10)     ; SET @DtpValDAT = 'NULL'                                  -- Date
    DECLARE @DtpValTXT varchar(10)     ; SET @DtpValTXT = ''''                                    -- Text
    DECLARE @DtpValBIN varchar(10)     ; SET @DtpValBIN = ''''                                    -- Binary
    DECLARE @DtpValVRN varchar(10)     ; SET @DtpValVRN = 'NULL'                                  -- Variant
    DECLARE @DtpValUNK varchar(10)     ; SET @DtpValUNK = ''''                                    -- Unknown
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize FieldLevel categories (EXEC ut_zzNAX FLVCAT,FLV)
    ------------------------------------------------------------------------------------------------
    DECLARE @FlvPKY    tinyint         ; SET @FlvPKY    = 01                                      -- Primary Keys
    DECLARE @FlvFKY    tinyint         ; SET @FlvFKY    = 02                                      -- Foreign Keys
    DECLARE @FlvMKY    tinyint         ; SET @FlvMKY    = 03                                      -- Mapping Keys
    DECLARE @FlvRTN    tinyint         ; SET @FlvRTN    = 04                                      -- Return Values
    DECLARE @FlvHTK    tinyint         ; SET @FlvHTK    = 05                                      -- History Tracking
    DECLARE @FlvLKX    tinyint         ; SET @FlvLKX    = 06                                      -- Lookup Exceptons
    DECLARE @FlvLKM    tinyint         ; SET @FlvLKM    = 07                                      -- Lookup KeyMap
    DECLARE @FlvDMM    tinyint         ; SET @FlvDMM    = 08                                      -- Dimension Master
    DECLARE @FlvFTM    tinyint         ; SET @FlvFTM    = 09                                      -- FactTable Master
    DECLARE @FlvFTX    tinyint         ; SET @FlvFTX    = 10                                      -- FactTable Exceptons
    DECLARE @FlvELD    tinyint         ; SET @FlvELD    = 11                                      -- Load ID
    DECLARE @FlvSRC    tinyint         ; SET @FlvSRC    = 12                                      -- Dimension Source IDs
    DECLARE @FlvST1    tinyint         ; SET @FlvST1    = 13                                      -- Standard Fields 1
    DECLARE @FlvSTD    tinyint         ; SET @FlvSTD    = 14                                      -- Standard Fields 2
    DECLARE @FlvST3    tinyint         ; SET @FlvST3    = 15                                      -- Standard Fields 3
    DECLARE @FlvLKP    tinyint         ; SET @FlvLKP    = 16                                      -- Lookup Fields
    DECLARE @FlvLNK    tinyint         ; SET @FlvLNK    = 17                                      -- Link Fields
    DECLARE @FlvPRS    tinyint         ; SET @FlvPRS    = 18                                      -- Parsing Fields
    DECLARE @FlvSEC    tinyint         ; SET @FlvSEC    = 19                                      -- Security Fields
    DECLARE @FlvLKD    tinyint         ; SET @FlvLKD    = 20                                      -- IsLocked Flag
    DECLARE @FlvUSD    tinyint         ; SET @FlvUSD    = 21                                      -- IsUsed Flag
    DECLARE @FlvDSB    tinyint         ; SET @FlvDSB    = 22                                      -- IsDisabled Flag
    DECLARE @FlvDLT    tinyint         ; SET @FlvDLT    = 23                                      -- IsDeleted Flag
    DECLARE @FlvLOK    tinyint         ; SET @FlvLOK    = 24                                      -- Locked  By/On
    DECLARE @FlvCRT    tinyint         ; SET @FlvCRT    = 25                                      -- Created By/On
    DECLARE @FlvUPD    tinyint         ; SET @FlvUPD    = 26                                      -- Updated By/On
    DECLARE @FlvEXP    tinyint         ; SET @FlvEXP    = 27                                      -- Expired By/On
    DECLARE @FlvDEL    tinyint         ; SET @FlvDEL    = 28                                      -- Deleted By/On
    DECLARE @FlvHST    tinyint         ; SET @FlvHST    = 29                                      -- History By/On
    DECLARE @FlvFLG    tinyint         ; SET @FlvFLG    = 30                                      -- History Flags
    DECLARE @FlvCLU    tinyint         ; SET @FlvCLU    = 31                                      -- Clear Used Flag
    DECLARE @FlvXPF    tinyint         ; SET @FlvXPF    = 32                                      -- Expire Record Flag
    DECLARE @FlvMDF    tinyint         ; SET @FlvMDF    = 33                                      -- Modify Record Flag
    DECLARE @FlvANF    tinyint         ; SET @FlvANF    = 34                                      -- AddNew Record Flag
    DECLARE @FlvDLF    tinyint         ; SET @FlvDLF    = 35                                      -- Delete Record Flag
    DECLARE @FlvRNT    tinyint         ; SET @FlvRNT    = 36                                      -- Runit Flag
    DECLARE @FlvRST    tinyint         ; SET @FlvRST    = 37                                      -- Reset Flag
    DECLARE @FlvTMR    tinyint         ; SET @FlvTMR    = 38                                      -- Timer Flag
    DECLARE @FlvDBG    tinyint         ; SET @FlvDBG    = 39                                      -- Debug Flag
    DECLARE @FlvTST    tinyint         ; SET @FlvTST    = 40                                      -- Test Flag
    DECLARE @FlvFWK    tinyint         ; SET @FlvFWK    = 41                                      -- Framework Flags
    DECLARE @FlvUTL    tinyint         ; SET @FlvUTL    = 42                                      -- Utility Flags
    DECLARE @FlvZZZ    tinyint         ; SET @FlvZZZ    = 43                                      -- Template
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize FieldList Constants  (FLC)                                   EXEC ut_zzUTL zzz,FLC
    ------------------------------------------------------------------------------------------------
    DECLARE @DecStmSTD char(8)         ; SET @DecStmSTD = 'DECLARE '                              -- Constant
    DECLARE @DecStmCMT char(8)         ; SET @DecStmCMT = '--CLARE '                              -- Constant
    DECLARE @DecAtsSTD char(9)         ; SET @DecAtsSTD = 'DECLARE @'                             -- Constant
    DECLARE @DecAtsCMT char(9)         ; SET @DecAtsCMT = '--CLARE @'                             -- Constant
    DECLARE @DecStmTXT varchar(9)      ; SET @DecStmTXT = ''                                      -- Reserved
    ------------------------------------------------------------------------------------------------
    DECLARE @SetStmSTD char(4)         ; SET @SetStmSTD = 'SET '                                  -- Constant
    DECLARE @SetStmCMT char(4)         ; SET @SetStmCMT = '--T '                                  -- Constant
    DECLARE @SetAtsSTD char(5)         ; SET @SetAtsSTD = 'SET @'                                 -- Constant
    DECLARE @SetAtsCMT char(5)         ; SET @SetAtsCMT = '--T @'                                 -- Constant
    DECLARE @SetStmTXT varchar(5)      ; SET @SetStmTXT = ''                                      -- Reserved
    ------------------------------------------------------------------------------------------------
    DECLARE @SetStmASN char(7)         ; SET @SetStmASN = '; SET '                                -- Constant
    DECLARE @SetAtsASN char(7)         ; SET @SetAtsASN = '; SET @'                               -- Constant
    DECLARE @SetStmSCN char(2)         ; SET @SetStmSCN = '; '                                    -- Constant
    DECLARE @SetStmEQL char(3)         ; SET @SetStmEQL = ' = '                                   -- Constant
    ------------------------------------------------------------------------------------------------
    DECLARE @ClmNulALN char(9)         ; SET @ClmNulALN = '     NULL'                             -- Constant
    DECLARE @ClmNulNNL char(9)         ; SET @ClmNulNNL = ' NOT NULL'                             -- Constant
    ------------------------------------------------------------------------------------------------
    DECLARE @ClmIdtYID char(9)         ; SET @ClmIdtYID = ' IDENTITY'                             -- Constant
    DECLARE @ClmIdtNID varchar(1)      ; SET @ClmIdtNID = ''                                      -- Constant
    ------------------------------------------------------------------------------------------------
    DECLARE @CfdTxtPFX char(3)         ; SET @CfdTxtPFX = 'AS '                                   -- Constant
    ------------------------------------------------------------------------------------------------
    DECLARE @CmtTxtPFX char(4)         ; SET @CmtTxtPFX = ' -- '                                  -- Constant
    DECLARE @CmtTxtSEP char(2)         ; SET @CmtTxtSEP = ': '                                    -- Constant
    ------------------------------------------------------------------------------------------------
    DECLARE @PrmOupTXT char(7)         ; SET @PrmOupTXT = ' OUTPUT'                               -- Constant
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize Statement Constants  (STC)                                   EXEC ut_zzUTL zzz,STC
    ------------------------------------------------------------------------------------------------
    -- For Variable declaration See ut_zzSQJ -> SecDTV
    -- For Column   definition  See ut_zzSQJ -> SecTSC
    /*----------------------------------------------------------------------------------------------
    +++++++++1+++++++++2+++++++++3+++++++++4+++++++++5+++++++++6+++++++++7+++++++++8+++++++++9++++++
    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456
    ------------------------------------------------------------------------------------------------
    DEC     VAR       _DTP             ;_SET VAR       EQLVAL                 _CMT    (SPC=_ SCN=;_)
    ----------------------------------------------------------------------------------------------*/
    DECLARE @MinLenVAR tinyint         ; SET @MinLenVAR = 9                                       -- Constant Length: Variable Minimum
    DECLARE @MaxLenDTP tinyint         ; SET @MaxLenDTP = 16                                      -- Constant Length: DataType Maximum = LEN('uniqueidentifier')
    ------------------------------------------------------------------------------------------------
    DECLARE @StdRmgVAR tinyint         ; SET @StdRmgVAR = 23                                      -- Constant Margin: Variable   (LM4+DEC+VAR+SPC                        ; to  DTP)
    DECLARE @StdRmgDEC tinyint         ; SET @StdRmgDEC = 39                                      -- Constant Margin: Declare    (LM4+DEC+VAR+SPC+DTP                    ; to  SCN)
    DECLARE @StdRmgEQX tinyint         ; SET @StdRmgEQX = 38                                      -- Constant Margin: Assign     (LM4+DEC+VAR+SPC+DTP+SCN+EQL            ; to  VAL)
    DECLARE @StdRmgSET tinyint         ; SET @StdRmgSET = 55                                      -- Constant Margin: Assign     (LM4+DEC+VAR+SPC+DTP+SCN+SET+VAR        ; to _EQL)
    DECLARE @StdRmgEQL tinyint         ; SET @StdRmgEQL = 58                                      -- Constant Margin: Assign     (LM4+DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL    ; to  VAL)
    DECLARE @StdRmgSTM tinyint         ; SET @StdRmgSTM = 97                                      -- Constant Margin: Statement  (LM4+DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL+VAL; to _CMT)
    ------------------------------------------------------------------------------------------------
    DECLARE @StmLenVAR tinyint         ; SET @StmLenVAR = 19                                      -- Constant Length: Variable   (DEC+VAR                                ; to  SCN)
    DECLARE @StmLenSVR tinyint         ; SET @StmLenSVR = 14                                      -- Constant Length: Assign     (SET+VAR                                ; to _EQL)
    DECLARE @StmLenAVR tinyint         ; SET @StmLenAVR = 17                                      -- Constant Length: Assign     (SET+VAR+EQL                            ; to  VAL)
    DECLARE @StmLenSET tinyint         ; SET @StmLenSET = 19                                      -- Constant Length: Assign     (SCN+SET+VAR+EQL                        ; to  VAL)
    DECLARE @StmLenDEC tinyint         ; SET @StmLenDEC = 35                                      -- Constant Length: Declare    (DEC+VAR+SPC+DTP                        ; to  SCN)
    DECLARE @StmLenINT tinyint         ; SET @StmLenINT = 54                                      -- Constant Length: Initialize (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL        ; to  VAL)
    DECLARE @StmLenSTM tinyint         ; SET @StmLenSTM = 97                                      -- Constant Length: Statement  (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL+VAL    ; to _CMT)
    ------------------------------------------------------------------------------------------------
    DECLARE @ClmLenFLD tinyint         ; SET @ClmLenFLD = 55                                      -- Constant Length: ColumnName (CLM                                    ; to  SPC)
    DECLARE @ClmLenCLM tinyint         ; SET @ClmLenCLM = @ClmLenFLD+1                            -- Constant Length: ColumnName (CLM+SPC                                ; to  DTP)
    DECLARE @ClmLenDFN tinyint         ; SET @ClmLenDFN = 93                                      -- Constant Length: Comment    (MX1+CLM+SPC+DTP+NUL+IDT                ; to _CMT)
    /*----------------------------------------------------------------------------------------------
    SET @PFX = 'VarPfx'; SET @COD = 'COD'; SET @DTP = 'varchar(11)'; SET @TTL = 'Description;
    PRINT LEFT(LEFT(LEFT(LEFT(@MX1+@DecAtsSTD+@PFX+@COD+@ITX,@StdRmgVAR)+@DTP+@ITX,@StdRmgDEC)+@SetStmASN+@PFX+@COD+@ITX,@StdRmgSET)+@SetStmEQL+@SQT+@COD+@SQT+@ITX,@StdRmgSTM)+@CmtTxtPFX+@TTL
    ----------------------------------------------------------------------------------------------*/
    DECLARE @WrkLenSTM tinyint         ; SET @WrkLenSTM = @StmLenSTM-@LftWid                      -- Constant Length
    DECLARE @WrkLenCLM tinyint         ; SET @WrkLenCLM = @ClmLenCLM-@LftWid                      -- Constant Length
    DECLARE @WrkLenDFN tinyint         ; SET @WrkLenDFN = @ClmLenDFN-@LftWid                      -- Constant Length
    ------------------------------------------------------------------------------------------------
    DECLARE @VfxMTY    varchar(1)      ; SET @VfxMTY    = ''                                      -- Constant PlaceHolder
    DECLARE @StmMTY    varchar(1)      ; SET @StmMTY    = ''                                      -- Constant PlaceHolder
    DECLARE @StmSPC    char(1)         ; SET @StmSPC    = ' '                                     -- Constant PlaceHolder
    ------------------------------------------------------------------------------------------------
    DECLARE @VlnDEC    tinyint         ; SET @VlnDEC    = 0                                       -- Working Length : Declare    (DEC+VAR+SPC+DTP                        ; to  SCN)
    DECLARE @VlnSET    tinyint         ; SET @VlnSET    = 0                                       -- Working Length : Set        (SCN+SET+VAR+EQL                        ; to  VAL)
    DECLARE @VlnINT    tinyint         ; SET @VlnINT    = 0                                       -- Working Length : Initialize (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL        ; to  VAL)
    DECLARE @VlnSTM    smallint        ; SET @VlnSTM    = 0                                       -- Working Length : Statement  (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL+VAL    ; to _CMT)
    ------------------------------------------------------------------------------------------------
    DECLARE @TxtDEC    varchar(300)    ; SET @TxtDEC    = ''                                      -- Working Text   : Declare    (DEC+VAR+SPC+DTP                        ; to  SCN)
    DECLARE @TxtSET    varchar(300)    ; SET @TxtSET    = ''                                      -- Working Text   : Set        (SCN+SET+VAR+EQL                        ; to  VAL)
    DECLARE @TxtINT    varchar(300)    ; SET @TxtINT    = ''                                      -- Working Text   : Initialize (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL        ; to  VAL)
    DECLARE @TxtSTM    varchar(300)    ; SET @TxtSTM    = ''                                      -- Working Text   : Statement  (DEC+VAR+SPC+DTP+SCN+SET+VAR+EQL+VAL    ; to _CMT)
    ------------------------------------------------------------------------------------------------
    DECLARE @StmOBJ    sysname         ; SET @StmOBJ    = ''                                      -- Working
    DECLARE @StmFLD    sysname         ; SET @StmFLD    = ''                                      -- Working
    DECLARE @StmCLM    sysname         ; SET @StmCLM    = ''                                      -- Working
    DECLARE @StmDTX    char(16)        ; SET @StmDTX    = ''                                      -- Working (Keep this in sync with GenSQL.clsUtlFMT.mcCrtLenDTP = 16)
    DECLARE @StmCFX    varchar(300)    ; SET @StmCFX    = ''                                      -- Working (Virtual field calculation text)
    DECLARE @StmNUL    char(9)         ; SET @StmNUL    = ''                                      -- Working
    DECLARE @StmIDT    varchar(9)      ; SET @StmIDT    = ''                                      -- Working
    DECLARE @StmVAL    varchar(100)    ; SET @StmVAL    = ''                                      -- Working
    DECLARE @StmCMT    varchar(200)    ; SET @StmCMT    = ''                                      -- Working
    ------------------------------------------------------------------------------------------------
    DECLARE @StmFLN    smallint        ; SET @StmFLN    = 0                                       -- Working FldLen
    DECLARE @StmCLN    smallint        ; SET @StmCLN    = 0                                       -- Working ClmLen
    ------------------------------------------------------------------------------------------------
    DECLARE @StmTMP    char(300)       ; SET @StmTMP    = ''                                      -- Working TempText
    DECLARE @FldTMP    char(300)       ; SET @FldTMP    = ''                                      -- Working TempText
    ------------------------------------------------------------------------------------------------
    DECLARE @HasCFX    bit             ; SET @HasCFX    = 0                                       -- Working
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Synchronize Input Object with FirstTable  (IJF)
    ------------------------------------------------------------------------------------------------
    IF LEN(@InpObj) = 0 EXEC dbo.ut_zzNAM TBN,NEW,TBL,'',@InpObj OUTPUT                           -- Get Default Object Name
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJF' AS 'IJF',LEFT(@InpObj,30) AS InpObj
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Store Original Object Values  (QJS)
    ------------------------------------------------------------------------------------------------
    IF LEN(@OupObj) = 0 SET @OupObj = @InpObj                                                     -- Synchronize Objects
    ------------------------------------------------------------------------------------------------
    DECLARE @InpTxt    sysname         ; SET @InpTxt    = @InpObj                                 -- Working Text
    DECLARE @OupTxt    sysname         ; SET @OupTxt    = @OupObj                                 -- Working Text
    DECLARE @WrkTxt    sysname         ; SET @WrkTxt    = ''                                      -- Working Text
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'QJS' AS 'QJS',LEFT(@InpTxt,30) AS InpTxt,LEFT(@OupTxt,30) AS OupTxt
    ------------------------------------------------------------------------------------------------
    -- Default Output Format  (QJF)
    ------------------------------------------------------------------------------------------------
    IF LEN(@OupFmt) = 0 SET @OupFmt = 'BAS'
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'QJF' AS 'QJF',@OupFmt AS OupFmt
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Initialize Input Object Attributes (EXEC ut_zzNAX INTATT,IJI)
    ------------------------------------------------------------------------------------------------
    DECLARE @InpSID    int             ; SET @InpSID = 0                                          -- Object sys.objects Identity
    DECLARE @InpTyp    varchar(2)      ; SET @InpTyp = ''                                         -- Object Type
    DECLARE @InpCls    varchar(3)      ; SET @InpCls = ''                                         -- Object Class
    --CLARE @InpObj    sysname         ; SET @InpObj = ''                                         -- Object Specification
    DECLARE @InpSrv    sysname         ; SET @InpSrv = ''                                         -- Object ServerName
    DECLARE @InpDbs    sysname         ; SET @InpDbs = ''                                         -- Object DBName
    DECLARE @InpScm    sysname         ; SET @InpScm = ''                                         -- Object SchemaName
    DECLARE @InpNam    sysname         ; SET @InpNam = ''                                         -- Object Name
    DECLARE @InpTbl    sysname         ; SET @InpTbl = ''                                         -- Object Table
    DECLARE @InpRef    sysname         ; SET @InpRef = ''                                         -- Object Reference
    DECLARE @InpFqd    sysname         ; SET @InpFqd = ''                                         -- Object FullyQualifiedDBName
    DECLARE @InpFqs    sysname         ; SET @InpFqs = ''                                         -- Object FullyQualifiedServer
    DECLARE @InpDtd    sysname         ; SET @InpDtd = ''                                         -- Object Default Table Desc
    DECLARE @InpBpx    varchar(3)      ; SET @InpBpx = ''                                         -- Object Base Prefix
    DECLARE @InpRpx    varchar(3)      ; SET @InpRpx = ''                                         -- Object Reference Prefix
    DECLARE @InpExs    bit             ; SET @InpExs = 0                                          -- Object Exists Flag
    DECLARE @InpPar    bit             ; SET @InpPar = 0                                          -- Object Has Parameters Flag
    DECLARE @InpItp    bit             ; SET @InpItp = 0                                          -- Object Is Temp Table
    DECLARE @InpIvb    bit             ; SET @InpIvb = 0                                          -- Object Is Variable Table
    DECLARE @InpAls    varchar(11)     ; SET @InpAls = ''                                         -- Table Alias
    DECLARE @InpPfx    varchar(11)     ; SET @InpPfx = ''                                         -- Table Prefix
    DECLARE @InpSfx    varchar(11)     ; SET @InpSfx = ''                                         -- Table Suffix
    DECLARE @InpBas    sysname         ; SET @InpBas = ''                                         -- Table Base
    DECLARE @InpCpx    varchar(20)     ; SET @InpCpx = ''                                         -- Table Column Prefix
    DECLARE @InpCur    varchar(50)     ; SET @InpCur = ''                                         -- Table Cursor Name
    DECLARE @InpCat    varchar(03)     ; SET @InpCat = ''                                         -- Table Category
    DECLARE @InpDat    bit             ; SET @InpDat = 0                                          -- Table IncludeData Flag
    DECLARE @InpAud    bit             ; SET @InpAud = 0                                          -- Table IncludeAudit Flag
    DECLARE @InpCnt    int             ; SET @InpCnt = 0                                          -- Table Record Count
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJI' AS 'IJI',@InpSID AS InpSID,@InpTyp AS InpTyp,@InpCls AS InpCls,LEFT(@InpObj,30) AS InpObj,LEFT(@InpSrv,30) AS InpSrv,LEFT(@InpDbs,30) AS InpDbs,LEFT(@InpScm,30) AS InpScm,LEFT(@InpNam,30) AS InpNam,LEFT(@InpTbl,30) AS InpTbl,LEFT(@InpRef,30) AS InpRef,LEFT(@InpFqd,30) AS InpFqd,LEFT(@InpFqs,30) AS InpFqs,LEFT(@InpDtd,30) AS InpDtd,@InpBpx AS InpBpx,@InpRpx AS InpRpx,@InpExs AS InpExs,@InpPar AS InpPar,@InpItp AS InpItp,@InpIvb AS InpIvb,@InpAls AS InpAls,@InpPfx AS InpPfx,@InpSfx AS InpSfx,LEFT(@InpBas,30) AS InpBas,@InpCpx AS InpCpx,LEFT(@InpCur,30) AS InpCur,@InpCat AS InpCat,@InpDat AS InpDat,@InpAud AS InpAud,@InpCnt AS InpCnt
    ------------------------------------------------------------------------------------------------
    -- Parse Object Name Into QualifiedName Parts
    ------------------------------------------------------------------------------------------------
    SET @WrkTxt = REVERSE(@InpTxt)                                                                -- Assign Working Text
    ------------------------------------------------------------------------------------------------
    -- Split Name Components On Dots  (from reversed object text)
    ------------------------------------------------------------------------------------------------
    IF CHARINDEX('.',@WrkTxt) > 0 BEGIN  -- ObjectName
        SET @InpNam = REVERSE(LEFT(@WrkTxt,CHARINDEX('.',@WrkTxt) - 1))
        SET @WrkTxt = SUBSTRING(@WrkTxt,CHARINDEX('.',@WrkTxt)+1,999)
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJQ' AS 'IJQ','OBJ' AS 'OBJ', @WrkTxt AS ORG_Name, @InpObj AS InpObj
    ------------------------------------------------------------------------------------------------
    IF CHARINDEX('.',@WrkTxt) > 0 BEGIN  -- SchemaName
        SET @InpScm = REVERSE(LEFT(@WrkTxt,CHARINDEX('.',@WrkTxt) - 1))
        SET @WrkTxt = SUBSTRING(@WrkTxt,CHARINDEX('.',@WrkTxt)+1,999)
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJQ' AS 'IJQ','SCM' AS 'SCM', @WrkTxt AS ORG_Name, @InpScm AS InpScm
    ------------------------------------------------------------------------------------------------
    IF CHARINDEX('.',@WrkTxt) > 0 BEGIN  -- DatabaseName
        SET @InpDbs = REVERSE(LEFT(@WrkTxt,CHARINDEX('.',@WrkTxt) - 1))
        SET @WrkTxt = SUBSTRING(@WrkTxt,CHARINDEX('.',@WrkTxt)+1,999)
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJQ' AS 'IJQ','DBS' AS 'DBS', @WrkTxt AS ORG_Name, @InpDbs AS InpDbs
    ------------------------------------------------------------------------------------------------
    IF LEN(@WrkTxt) > 0 BEGIN            -- ServerName
        SET @WrkTxt = REVERSE(REPLACE(@WrkTxt,'.',''))
        IF LEN(@InpNam) = 0 BEGIN
            SET @InpNam = @WrkTxt
        END ELSE IF LEN(@InpScm) = 0 BEGIN
            SET @InpScm = @WrkTxt
        END ELSE IF LEN(@InpDbs) = 0 BEGIN
            SET @InpDbs = @WrkTxt
        END ELSE BEGIN
            SET @InpSrv = @WrkTxt
        END
    END
    ------------------------------------------------------------------------------------------------
    SET @InpObj = @InpNam
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJQ' AS 'IJQ','MTY' AS 'MTY', @WrkTxt AS ORG_Name, @InpSrv+'.'+@InpDbs+'.'+@InpScm+'.'+@InpObj AS FQN_Name
    ------------------------------------------------------------------------------------------------
    -- Assign Defaults to Empty Values
    ------------------------------------------------------------------------------------------------
    IF LEN(@InpObj) = 0 SET @InpObj = 'ZzzZzz'
    IF LEN(@InpScm) = 0 SET @InpScm = ISNULL((SELECT TOP 1 scm.name FROM sys.objects soj INNER JOIN sys.schemas scm ON scm.schema_id = soj.schema_id WHERE soj.name = @InpObj),'dbo')
    IF LEN(@InpDbs) = 0 SET @InpDbs = DB_NAME()
    IF LEN(@InpSrv) = 0 SET @InpSrv = @@SERVERNAME
    ------------------------------------------------------------------------------------------------
    -- Compose Qualified Names
    ------------------------------------------------------------------------------------------------
    SET @InpRef = @InpScm+'.'+@InpObj                                                             -- Input Reference
    SET @InpFqd = @InpSrv+'.'+@InpRef                                                             -- Input Fully Qualified Database
    SET @InpFqs = @InpSrv+'.'+@InpFqd                                                             -- Input Fully Qualified Server
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJQ' AS 'IJQ','FQN' AS 'FQN', @InpSrv+'.'+@InpDbs+'.'+@InpScm+'.'+@InpObj AS FQN_Name
    ------------------------------------------------------------------------------------------------
    -- Set Input object identifier  (IJX)
    ------------------------------------------------------------------------------------------------
    DECLARE @InpSIX    varchar(20)     ; SET @InpSIX    = ''                                      -- Input sys.objects Identity Text
    ------------------------------------------------------------------------------------------------
    SET @InpSID = ISNULL(OBJECT_ID(@InpRef),0)                                                    -- Input sys.objects Identity Number
    SET @InpSIX = CAST(ISNULL(@InpSID,0) AS varchar(20))                                          -- Input sys.objects Identity Text
    SET @InpExs = CASE WHEN @InpSID > 0 THEN 1 ELSE 0 END                                         -- Input Object Exists
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJX' AS 'IJX',LEFT(@InpRef,30) AS InpObj,@InpSID AS InpSID,@InpExs AS InpExs
    ------------------------------------------------------------------------------------------------
    -- Lookup Input Object Type/Class  (IJY)
    ------------------------------------------------------------------------------------------------
    SET @InpTyp = ISNULL((SELECT type FROM sys.objects WHERE object_id = @InpSID),@ObjTypTBL)
    ------------------------------------------------------------------------------------------------
    SET @InpCls = CASE @InpTyp
        WHEN @ObjTypTBL THEN @ObjClsTBL
        WHEN @ObjTypVEW THEN @ObjClsVEW
        WHEN @ObjTypUSP THEN @ObjClsUSP
        WHEN @ObjTypUFN THEN @ObjClsUFN
        ELSE @MTY
    END
    ------------------------------------------------------------------------------------------------
    IF LEN(@InpCls) = 0 BEGIN
        EXEC dbo.ut_zzNAM LKP,CLS,XXX,@InpObj,@InpCls OUTPUT
    END
    ------------------------------------------------------------------------------------------------
    SET @InpPar = CASE WHEN @InpCls IN (@ObjClsUSP,@ObjClsTRG,@ObjClsUFN) THEN 1 ELSE 0 END       -- Assign HasParameter Flag
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJY' AS 'IJY',InpObj=LEFT(@InpObj,30),InpTyp=@InpTyp,InpCls=@InpCls,InpPar=@InpPar
    ------------------------------------------------------------------------------------------------
    -- Assign Input Object TableName  (IJN)
    ------------------------------------------------------------------------------------------------
    IF @InpCls = @ObjClsVEW BEGIN
        EXEC dbo.ut_zzNAM VWN,NAM,XXX,@InpNam,@InpTbl OUTPUT
    END ELSE BEGIN
        EXEC dbo.ut_zzNAM TBN,NAM,XXX,@InpNam,@InpTbl OUTPUT
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJN' AS 'IJN',LEFT(@InpObj,30) AS InpObj,LEFT(@InpNam,30) AS InpNam,LEFT(@InpTbl,30) AS InpTbl
 
    ------------------------------------------------------------------------------------------------
    -- Lookup Input Object Attributes (EXEC ut_zzNAX LKPATT,IJB)
    ------------------------------------------------------------------------------------------------
    EXEC dbo.ut_zzNAM TBL,ALS,AL0,@InpTbl,@InpAls OUTPUT                                          -- Table Alias
    EXEC dbo.ut_zzNAM TBL,PFX,XXX,@InpTbl,@InpPfx OUTPUT                                          -- Table Prefix
    EXEC dbo.ut_zzNAM TBL,BAS,XXX,@InpTbl,@InpBas OUTPUT                                          -- Table Base
    EXEC dbo.ut_zzNAM TBL,CPX,XXX,@InpTbl,@InpCpx OUTPUT                                          -- Table Column Prefix
    EXEC dbo.ut_zzNAM TBL,CUR,XXX,@InpTbl,@InpCur OUTPUT                                          -- Table Cursor Name
    EXEC dbo.ut_zzNAM TBL,CAT,XXX,@InpTbl,@InpCat OUTPUT                                          -- Table Category
    EXEC dbo.ut_zzNAM TBL,DAT,XXX,@InpTbl,@InpDat OUTPUT                                          -- Table IncludeData Flag
    EXEC dbo.ut_zzNAM TBL,AUD,XXX,@InpTbl,@InpAud OUTPUT                                          -- Table IncludeAudit Flag
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJB' AS 'IJB',InpAls=@InpAls,InpPfx=@InpPfx,InpBas=LEFT(@InpBas,30),InpCpx=@InpCpx,InpCur=LEFT(@InpCur,30),InpCat=@InpCat,InpDat=@InpDat,InpAud=@InpAud
    ------------------------------------------------------------------------------------------------
    ------------------------------------------------------------------------------------------------
    -- Get Input Count (IJC)
    ------------------------------------------------------------------------------------------------
    IF @InpCls = @ObjClsTBL BEGIN
        SET @InpCnt = ISNULL((SELECT rows FROM sys.partitions WHERE object_id = OBJECT_ID(@InpRef) AND index_id IN (0,1)),0)
    END ELSE IF @InpCls = @ObjClsVEW BEGIN
        IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(@InpRef)) BEGIN
            CREATE TABLE #tmp_InpCnt (RecCnt int)
            EXEC (N'INSERT INTO #tmp_InpCnt VALUES (ISNULL((SELECT COUNT(*) FROM '+@InpRef+'),0))'); SET @InpCnt = ISNULL((SELECT RecCnt FROM #tmp_InpCnt),0)
            DROP TABLE   #tmp_InpCnt
        END
    END
    ------------------------------------------------------------------------------------------------
    SET @InpDat = ISNULL(CASE WHEN @InpDat = 1 OR @IncDat = 1 OR (@IncDat = 2 AND @InpCnt BETWEEN 1 AND 10) THEN 1 ELSE 0 END,0); SET @IncDat = @InpDat
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJC' AS 'IJC',InpObj=LEFT(@InpObj,30),InpCnt=@InpCnt,InpCls=@InpCls,InpRef=@InpRef,InpDat=@InpDat,IncDat=@IncDat
    ------------------------------------------------------------------------------------------------
    -- Assign TableFormat References  (IJT)
    ------------------------------------------------------------------------------------------------
    IF          LEFT(@InpNam,1) = '#' BEGIN
        SET @InpItp = 1
        SET @InpDtd = 'temp table'
    END ELSE IF LEFT(@InpNam,1) = '@' BEGIN
        SET @InpIvb = 1
        SET @InpDtd = 'variable table'
    END ELSE BEGIN
        SET @InpDtd = 'standard table'
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJT' AS 'IJT',InpObj=LEFT(@InpObj,30),InpNam=LEFT(@InpNam,30),InpItp=@InpItp,InpIvb=@InpIvb,InpDtd=LEFT(@InpDtd,30)
    ------------------------------------------------------------------------------------------------
    -- Assign TableFormat Prefixes  (IJP)
    ------------------------------------------------------------------------------------------------
    SET @InpBpx = UPPER(LEFT(@InpBas,3))                                                          -- Base prefix
    SET @InpRpx = LEFT(@InpBpx,1)+LOWER(RIGHT(@InpBpx,2))                                         -- Reference prefix
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'IJP' AS 'IJP',InpObj=LEFT(@InpObj,30),InpBas=LEFT(@InpBas,30),InpBpx=@InpBpx,InpRpx=@InpRpx
    ------------------------------------------------------------------------------------------------
    -- Debug Input Object Attributes (EXEC ut_zzNAX DBGATT,IJG)
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 BEGIN
        PRINT REPLICATE('-',80)
        PRINT '-- Input  Values'
        PRINT REPLICATE('-',80)
        PRINT 'InpSID=' + CONVERT(varchar(200),@InpSID)                                           -- Object sys.objects Identity
        PRINT 'InpTyp=' + CONVERT(varchar(200),@InpTyp)                                           -- Object Type
        PRINT 'InpCls=' + CONVERT(varchar(200),@InpCls)                                           -- Object Class
        PRINT 'InpObj=' + CONVERT(varchar(200),@InpObj)                                           -- Object Specification
        PRINT 'InpSrv=' + CONVERT(varchar(200),@InpSrv)                                           -- Object ServerName
        PRINT 'InpDbs=' + CONVERT(varchar(200),@InpDbs)                                           -- Object DBName
        PRINT 'InpScm=' + CONVERT(varchar(200),@InpScm)                                           -- Object SchemaName
        PRINT 'InpNam=' + CONVERT(varchar(200),@InpNam)                                           -- Object Name
        PRINT 'InpTbl=' + CONVERT(varchar(200),@InpTbl)                                           -- Object Table
        PRINT 'InpRef=' + CONVERT(varchar(200),@InpRef)                                           -- Object Reference
        PRINT 'InpFqd=' + CONVERT(varchar(200),@InpFqd)                                           -- Object FullyQualifiedDBName
        PRINT 'InpFqs=' + CONVERT(varchar(200),@InpFqs)                                           -- Object FullyQualifiedServer
        PRINT 'InpDtd=' + CONVERT(varchar(200),@InpDtd)                                           -- Object Default Table Desc
        PRINT 'InpBpx=' + CONVERT(varchar(200),@InpBpx)                                           -- Object Base Prefix
        PRINT 'InpRpx=' + CONVERT(varchar(200),@InpRpx)                                           -- Object Reference Prefix
        PRINT 'InpExs=' + CONVERT(varchar(200),@InpExs)                                           -- Object Exists Flag
        PRINT 'InpPar=' + CONVERT(varchar(200),@InpPar)                                           -- Object Has Parameters Flag
        PRINT 'InpItp=' + CONVERT(varchar(200),@InpItp)                                           -- Object Is Temp Table
        PRINT 'InpIvb=' + CONVERT(varchar(200),@InpIvb)                                           -- Object Is Variable Table
        PRINT 'InpAls=' + CONVERT(varchar(200),@InpAls)                                           -- Table Alias
        PRINT 'InpPfx=' + CONVERT(varchar(200),@InpPfx)                                           -- Table Prefix
        PRINT 'InpSfx=' + CONVERT(varchar(200),@InpSfx)                                           -- Table Suffix
        PRINT 'InpBas=' + CONVERT(varchar(200),@InpBas)                                           -- Table Base
        PRINT 'InpCpx=' + CONVERT(varchar(200),@InpCpx)                                           -- Table Column Prefix
        PRINT 'InpCur=' + CONVERT(varchar(200),@InpCur)                                           -- Table Cursor Name
        PRINT 'InpCat=' + CONVERT(varchar(200),@InpCat)                                           -- Table Category
        PRINT 'InpDat=' + CONVERT(varchar(200),@InpDat)                                           -- Table IncludeData Flag
        PRINT 'InpAud=' + CONVERT(varchar(200),@InpAud)                                           -- Table IncludeAudit Flag
        PRINT 'InpCnt=' + CONVERT(varchar(200),@InpCnt)                                           -- Table Record Count
    END
    ------------------------------------------------------------------------------------------------
 
    --##############################################################################################


    --XGM@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@XGM


    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Synchronize Display object with Source object  (DJS)
    ------------------------------------------------------------------------------------------------
    IF LEN(@DspObj) = 0 SET @DspObj = @InpObj
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT DJS='DJS',SrcObj=LEFT(@InpObj,30),DspObj=LEFT(@DspObj,30)
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Set Display object identifier  (DJI)
    ------------------------------------------------------------------------------------------------
    DECLARE @DspID     int          = ISNULL(OBJECT_ID(@DspObj),0)
    DECLARE @DspExs    bit          = CASE WHEN @DspID > 0 THEN 1 ELSE 0 END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT DJI='DJI',DspObj=LEFT(@DspObj,30),DspID=@DspID,DspExs=@DspExs
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Assign Display object values  (DJV)
    ------------------------------------------------------------------------------------------------
    DECLARE @DspNam    sysname      = @DspObj                                                     -- Display name (replaces @ObjNam)
    DECLARE @DspTbl    sysname      ; EXEC ut_zzNAM TBL,NAM,XXX,@DspNam,@DspTbl OUTPUT
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT DJV='DJV',DspObj=LEFT(@DspObj,30),DspNam=LEFT(@DspNam,30),DspTbl=LEFT(@DspTbl,30)
    ------------------------------------------------------------------------------------------------
 
    ------------------------------------------------------------------------------------------------
    -- Lookup Display Object Attributes (EXEC ut_zzNAX LKPATT,DJT)
    ------------------------------------------------------------------------------------------------
    DECLARE @DspAls varchar(11)     ; EXEC dbo.ut_zzNAM TBL,ALS,AL0,@DspTbl,@DspAls OUTPUT        -- Table Alias
    DECLARE @DspPfx varchar(11)     ; EXEC dbo.ut_zzNAM TBL,PFX,XXX,@DspTbl,@DspPfx OUTPUT        -- Table Prefix
    DECLARE @DspBas sysname         ; EXEC dbo.ut_zzNAM TBL,BAS,XXX,@DspTbl,@DspBas OUTPUT        -- Table Base
    DECLARE @DspCpx varchar(20)     ; EXEC dbo.ut_zzNAM TBL,CPX,XXX,@DspTbl,@DspCpx OUTPUT        -- Table Column Prefix
    DECLARE @DspCur varchar(50)     ; EXEC dbo.ut_zzNAM TBL,CUR,XXX,@DspTbl,@DspCur OUTPUT        -- Table Cursor Name
    DECLARE @DspCat varchar(03)     ; EXEC dbo.ut_zzNAM TBL,CAT,XXX,@DspTbl,@DspCat OUTPUT        -- Table Category
    DECLARE @DspDat bit             ; EXEC dbo.ut_zzNAM TBL,DAT,XXX,@DspTbl,@DspDat OUTPUT        -- Table IncludeData Flag
    DECLARE @DspAud bit             ; EXEC dbo.ut_zzNAM TBL,AUD,XXX,@DspTbl,@DspAud OUTPUT        -- Table IncludeAudit Flag
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'DJT' AS 'DJT',DspAls=@DspAls,DspPfx=@DspPfx,DspBas=LEFT(@DspBas,30),DspCpx=@DspCpx,DspCur=LEFT(@DspCur,30),DspCat=@DspCat,DspDat=@DspDat,DspAud=@DspAud
    ------------------------------------------------------------------------------------------------
  
    ------------------------------------------------------------------------------------------------
    -- Resolve Ouput object (from Display name - ROJ)
    ------------------------------------------------------------------------------------------------
    IF LEN(@OupObj) = 0 BEGIN
        IF LEN(@DspNam) > 0 BEGIN
            SET @OupObj = @DspNam
        END ELSE BEGIN
            --EC ut_zzNAM LKP,NAM,XXX    ,@DspNam,@OupObj OUTPUT,0  --','',1
            EXEC dbo.ut_zzNAM OBJ,NAM,@NamFmt,@DspNam,@OupObj OUTPUT,0  --','',1
        END
    END
    -- Resolve Ouput description (from Display name - ROJ)
    IF LEN(@OupDsc) = 0 BEGIN
        EXEC dbo.ut_zzNAM LKP,DSC,XXX,@OupObj,@OupDsc OUTPUT,0  --','',1
    END
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT 'ROJ' AS 'ROJ',DspNam=LEFT(@DspNam,30),NamFmt=@NamFmt,OupObj=LEFT(@OupObj,30),OupDsc=@OupDsc
    ------------------------------------------------------------------------------------------------


    --##############################################################################################

    -- Set Source object identifier  (SJI)
    DECLARE @SrcID     int             ; SET @SrcID     = ISNULL(OBJECT_ID(@InpObj),0)
    DECLARE @SrcExs    bit             ; SET @SrcExs    = CASE WHEN @SrcID > 0 THEN 1 ELSE 0 END
    IF @DbgFlg = 1 OR 0=9 SELECT SJI='SJI',SrcObj=LEFT(@InpObj,30),SrcID=@SrcID,SrcExs=@SrcExs
 
    --##############################################################################################
 
    -- Assign Source object values  (SJV)
    DECLARE @SrcNam    sysname         ; SET @SrcNam    = @InpObj
    DECLARE @SrcTbl    sysname      ; EXEC ut_zzNAM TBL,NAM,XXX,@SrcNam,@SrcTbl OUTPUT
    IF @DbgFlg = 1 OR 0=9 SELECT SJV='SJV',SrcObj=LEFT(@InpObj,30),SrcNam=LEFT(@SrcNam,30),SrcTbl=LEFT(@SrcTbl,30)
 
    --##############################################################################################
 
    -- Lookup Source object type  (SJY - see ut_zzNAM)
    DECLARE @SrcTyp    varchar(3)   ; EXEC ut_zzNAM LKP,CLS,XXX,@InpObj,@SrcTyp OUTPUT
    DECLARE @StpTBL    bit             ; SET @StpTBL = CASE WHEN @SrcTyp = @ObjTypTBL THEN 1 ELSE 0 END  -- Table
    DECLARE @StpVEW    bit             ; SET @StpVEW = CASE WHEN @SrcTyp = @ObjTypVEW THEN 1 ELSE 0 END  -- View
    DECLARE @StpUSP    bit             ; SET @StpUSP = CASE WHEN @SrcTyp = @ObjTypUSP THEN 1 ELSE 0 END  -- SProc
    DECLARE @StpTRG    bit             ; SET @StpTRG = CASE WHEN @SrcTyp = @ObjTypTRG THEN 1 ELSE 0 END  -- Trigger
    DECLARE @StpUFN    bit             ; SET @StpUFN = CASE WHEN @SrcTyp = @ObjTypUFN THEN 1 ELSE 0 END  -- Function
    DECLARE @StpPKY    bit             ; SET @StpPKY = CASE WHEN @SrcTyp = @ObjTypPKY THEN 1 ELSE 0 END  -- PrimaryKey
    DECLARE @StpUKY    bit             ; SET @StpUKY = CASE WHEN @SrcTyp = @ObjTypUKY THEN 1 ELSE 0 END  -- UniqueKey
    DECLARE @StpIND    bit             ; SET @StpIND = CASE WHEN @SrcTyp = @ObjTypIND THEN 1 ELSE 0 END  -- Index
    DECLARE @StpFKY    bit             ; SET @StpFKY = CASE WHEN @SrcTyp = @ObjTypFKY THEN 1 ELSE 0 END  -- ForeignKey
    DECLARE @StpDEF    bit             ; SET @StpDEF = CASE WHEN @SrcTyp = @ObjTypDEF THEN 1 ELSE 0 END  -- Default
    DECLARE @StpCHK    bit             ; SET @StpCHK = CASE WHEN @SrcTyp = @ObjTypCHK THEN 1 ELSE 0 END  -- Check
    DECLARE @StpDDL    bit             ; SET @StpDDL = CASE WHEN @SrcTyp = @ObjTypDDL THEN 1 ELSE 0 END  -- DataDict
    DECLARE @StpSCP    bit             ; SET @StpSCP = CASE WHEN @SrcTyp = @ObjTypSCP THEN 1 ELSE 0 END  -- Script
    DECLARE @StpVDN    bit             ; SET @StpVDN = CASE WHEN @SrcTyp = @ObjTypVDN THEN 1 ELSE 0 END  -- VB.NET
    DECLARE @StpUNK    bit             ; SET @StpUNK = CASE WHEN @SrcTyp = @ObjTypUNK THEN 1 ELSE 0 END  -- Unknown
    DECLARE @SrcPrm    bit             ; SET @SrcPrm = CASE WHEN @SrcTyp IN (@ObjTypUSP,@ObjTypTRG,@ObjTypUFN) THEN 1 ELSE 0 END
    IF @DbgFlg = 1 OR 0=9 SELECT SJY='SJY',SrcObj=LEFT(@InpObj,30),SrcTyp=@SrcTyp,SrcPrm=@SrcPrm,StpTBL=@StpTBL,StpVEW=@StpVEW,StpUSP=@StpUSP,StpTRG=@StpTRG,StpUFN=@StpUFN,StpPKY=@StpPKY,StpUKY=@StpUKY,StpIND=@StpIND,StpFKY=@StpFKY,StpDEF=@StpDEF,StpCHK=@StpCHK,StpDDL=@StpDDL,StpSCP=@StpSCP,StpVDN=@StpVDN,StpUNK=@StpUNK
 
    --##############################################################################################
 
    -- Synchronize Table object with Source object  (TJS)
    DECLARE @TblObj    sysname         ; SET @TblObj    = @InpObj
    IF @StpTBL <> 1 BEGIN
        EXEC ut_zzNAM TBN,'','',@InpObj,@TblObj OUTPUT
    END
    IF @DbgFlg = 1 OR 0=9 SELECT TJS='TJS',SrcObj=LEFT(@InpObj,30),TblObj=LEFT(@TblObj,30)
 
    --##############################################################################################
 
    -- Set Table object identifier  (TJI)
    DECLARE @TblID     int             ; SET @TblID     = ISNULL(OBJECT_ID(@TblObj),0)
    DECLARE @TblExs    bit             ; SET @TblExs    = CASE WHEN @TblID > 0 THEN 1 ELSE 0 END
    IF @DbgFlg = 1 OR 0=9 SELECT TJI='TJI',TblObj=LEFT(@TblObj,30),TblID=@TblID,TblExs=@TblExs
 
    --##############################################################################################
 
    -- Assign Table object values  (TJV)
    DECLARE @TblNam    sysname         ; SET @TblNam    = @TblObj
    IF @DbgFlg = 1 OR 0=9 SELECT TJV='TJV',TblObj=LEFT(@TblObj,30),TblNam=LEFT(@TblNam,30)
 
    --##############################################################################################
 
    -- Lookup Table object attributes (TJT - see ut_zzNAM)
    DECLARE @TblAls    varchar(10)  ; EXEC ut_zzNAM TBL,ALS,AL0,@TblNam,@TblAls OUTPUT  -- Table Alias
    DECLARE @TblBas    sysname      ; EXEC ut_zzNAM TBL,BAS,XXX,@TblNam,@TblBas OUTPUT  -- Table Base
    DECLARE @TblPfx    varchar(10)  ; EXEC ut_zzNAM TBL,PFX,XXX,@TblNam,@TblPfx OUTPUT  -- Table Prefix
    DECLARE @TblCpx    varchar(20)  ; EXEC ut_zzNAM TBL,CPX,XXX,@TblNam,@TblCpx OUTPUT  -- Column Prefix
    DECLARE @TblCur    varchar(50)  ; EXEC ut_zzNAM TBL,CUR,XXX,@TblNam,@TblCur OUTPUT  -- Cursor Name
    DECLARE @TblHst    sysname      ; EXEC ut_zzNAM TBL,HST,XXX,@TblNam,@TblHst OUTPUT  -- History Name
    DECLARE @TblCat    varchar(03)  ; EXEC ut_zzNAM TBL,CAT,XXX,@TblNam,@TblCat OUTPUT  -- Table Category
    DECLARE @TblFmx    varchar(03)  ; EXEC ut_zzNAM TBL,FMX,XXX,@TblNam,@TblFmx OUTPUT  -- Table Format
    DECLARE @TblDat    bit          ; EXEC ut_zzNAM TBL,DAT,XXX,@TblNam,@TblDat OUTPUT  -- IncludeData Flag
    DECLARE @TblAud    bit          ; EXEC ut_zzNAM TBL,AUD,XXX,@TblNam,@TblAud OUTPUT  -- IncludeAudit Flag
    IF @DbgFlg = 1 OR 0=9 SELECT TJT='TJT',TblNam=LEFT(@TblNam,30),TblAls=@TblAls,TblBas=@TblBas,TblPfx=@TblPfx,TblCpx=@TblCpx,TblCur=@TblCur,TblHst=@TblHst,TblCat=@TblCat,TblFmx=@TblFmx,TblDat=@TblDat,TblAud=@TblAud
 
    --##############################################################################################
 
    -- Lookup Table object category (TJC - see ut_zzNAM)
    DECLARE @TblCtg    varchar(3)   ; EXEC ut_zzNAM TBL,CAT,XXX,@TblNam,@TblCtg OUTPUT
    DECLARE @TctSTP    bit             ; SET @TctSTP = CASE WHEN @TblCtg = @TblCatSTP THEN 1 ELSE 0 END  -- Setup
    DECLARE @TctLKP    bit             ; SET @TctLKP = CASE WHEN @TblCtg = @TblCatLKP THEN 1 ELSE 0 END  -- Lookup
    DECLARE @TctSEC    bit             ; SET @TctSEC = CASE WHEN @TblCtg = @TblCatSEC THEN 1 ELSE 0 END  -- Security
    DECLARE @TctREF    bit             ; SET @TctREF = CASE WHEN @TblCtg = @TblCatREF THEN 1 ELSE 0 END  -- Reference
    DECLARE @TctTRX    bit             ; SET @TctTRX = CASE WHEN @TblCtg = @TblCatTRX THEN 1 ELSE 0 END  -- Transaction
    DECLARE @TctLNK    bit             ; SET @TctLNK = CASE WHEN @TblCtg = @TblCatLNK THEN 1 ELSE 0 END  -- Link
    DECLARE @TctDSS    bit             ; SET @TctDSS = CASE WHEN @TblCtg = @TblCatDSS THEN 1 ELSE 0 END  -- DecisionSupport
    DECLARE @TctHIS    bit             ; SET @TctHIS = CASE WHEN @TblCtg = @TblCatHIS THEN 1 ELSE 0 END  -- History
    DECLARE @TctARC    bit             ; SET @TctARC = CASE WHEN @TblCtg = @TblCatARC THEN 1 ELSE 0 END  -- Archive
    DECLARE @TctTBL    bit             ; SET @TctTBL = CASE WHEN @TblCtg = @TblCatTBL THEN 1 ELSE 0 END  -- Table
    DECLARE @TctFCT    bit             ; SET @TctFCT = CASE WHEN @TblCtg = @TblCatFCT THEN 1 ELSE 0 END  -- Fact
    DECLARE @TctDIM    bit             ; SET @TctDIM = CASE WHEN @TblCtg = @TblCatDIM THEN 1 ELSE 0 END  -- Dimension
    DECLARE @TctETL    bit             ; SET @TctETL = CASE WHEN @TblCtg = @TblCatETL THEN 1 ELSE 0 END  -- Transform
    DECLARE @TctVBA    bit             ; SET @TctVBA = CASE WHEN @TblCtg = @TblCatVBA THEN 1 ELSE 0 END  -- Table (VBAGen)
    DECLARE @TctUNK    bit             ; SET @TctUNK = CASE WHEN @TblCtg = @TblCatUNK THEN 1 ELSE 0 END  -- Unknown
    IF @DbgFlg = 1 OR 0=9 SELECT TJC='TJC',TblNam=LEFT(@TblNam,30),TblCat=@TblCat,TctSTP=@TctSTP,TctLKP=@TctLKP,TctSEC=@TctSEC,TctREF=@TctREF,TctTRX=@TctTRX,TctLNK=@TctLNK,TctDSS=@TctDSS,TctHIS=@TctHIS,TctARC=@TctARC,TctTBL=@TctTBL,TctFCT=@TctFCT,TctDIM=@TctDIM,TctETL=@TctETL,TctVBA=@TctVBA,TctUNK=@TctUNK
 
    --##############################################################################################
 
    -- Set Output object identifier  (OJI)
    DECLARE @OupID     int             ; SET @OupID     = ISNULL(OBJECT_ID(@OupObj),0)
    DECLARE @OupExs    bit             ; SET @OupExs    = CASE WHEN @OupID > 0 THEN 1 ELSE 0 END
    IF @DbgFlg = 1 OR 0=9 SELECT OJI='OJI',OupObj=LEFT(@OupObj,30),OupID=@OupID,OupExs=@OupExs
 
    --##############################################################################################
 
    -- Assign Output object values  (OJV)
    DECLARE @OupNam    sysname         ; SET @OupNam    = @OupObj
    DECLARE @OupTbl    sysname      ; EXEC ut_zzNAM TBL,NAM,XXX,@OupNam,@OupTbl OUTPUT
    IF @DbgFlg = 1 OR 0=9 SELECT OJV='OJV',OupObj=LEFT(@OupObj,30),OupNam=LEFT(@OupNam,30),OupTbl=LEFT(@OupTbl,30)
 
    --##############################################################################################
 
    -- Lookup Output object attributes (OJT - see ut_zzNAM)
    --CLARE @OupAls    varchar(10)  ; EXEC ut_zzNAM TBL,ALS,AL0,@OupTbl,@OupAls OUTPUT  -- Table Alias
    DECLARE @OupBas    sysname      ; EXEC ut_zzNAM TBL,BAS,XXX,@OupTbl,@OupBas OUTPUT  -- Table Base
    DECLARE @OupPfx    varchar(10)  ; EXEC ut_zzNAM TBL,PFX,XXX,@OupTbl,@OupPfx OUTPUT  -- Table Prefix
    --CLARE @OupCpx    varchar(20)  ; EXEC ut_zzNAM TBL,CPX,XXX,@OupTbl,@OupCpx OUTPUT  -- Column Prefix
    DECLARE @OupCur    varchar(50)  ; EXEC ut_zzNAM TBL,CUR,XXX,@OupTbl,@OupCur OUTPUT  -- Cursor Name
    DECLARE @OupHst    sysname      ; EXEC ut_zzNAM TBL,HST,XXX,@OupTbl,@OupHst OUTPUT  -- History Name
    DECLARE @OupCat    varchar(03)  ; EXEC ut_zzNAM TBL,CAT,XXX,@OupTbl,@OupCat OUTPUT  -- Table Category
    DECLARE @OupFmx    varchar(03)  ; EXEC ut_zzNAM TBL,FMX,XXX,@OupTbl,@OupFmx OUTPUT  -- Table Format
    DECLARE @OupDat    bit          ; EXEC ut_zzNAM TBL,DAT,XXX,@OupTbl,@OupDat OUTPUT  -- IncludeData Flag
    DECLARE @OupAud    bit          ; EXEC ut_zzNAM TBL,AUD,XXX,@OupTbl,@OupAud OUTPUT  -- IncludeAudit Flag
    IF @DbgFlg = 1 OR 0=9 SELECT OJT='OJT',OupTbl=LEFT(@OupTbl,30),OupAls=@OupAls,OupBas=@OupBas,OupPfx=@OupPfx,OupCpx=@OupCpx,OupCur=@OupCur,OupHst=@OupHst,OupCat=@OupCat,OupFmx=@OupFmx,OupDat=@OupDat,OupAud=@OupAud
 
    --##############################################################################################
 
    -- Lookup Output object type  (OJY - see ut_zzNAM)
    DECLARE @OupTyp    varchar(3)   ; EXEC ut_zzNAM LKP,CLS,XXX,@OupObj,@OupTyp OUTPUT
    DECLARE @OtpTBL    bit             ; SET @OtpTBL = CASE WHEN @OupTyp = @ObjTypTBL THEN 1 ELSE 0 END  -- Table
    DECLARE @OtpVEW    bit             ; SET @OtpVEW = CASE WHEN @OupTyp = @ObjTypVEW THEN 1 ELSE 0 END  -- View
    DECLARE @OtpUSP    bit             ; SET @OtpUSP = CASE WHEN @OupTyp = @ObjTypUSP THEN 1 ELSE 0 END  -- SProc
    DECLARE @OtpTRG    bit             ; SET @OtpTRG = CASE WHEN @OupTyp = @ObjTypTRG THEN 1 ELSE 0 END  -- Trigger
    DECLARE @OtpUFN    bit             ; SET @OtpUFN = CASE WHEN @OupTyp = @ObjTypUFN THEN 1 ELSE 0 END  -- Function
    DECLARE @OtpPKY    bit             ; SET @OtpPKY = CASE WHEN @OupTyp = @ObjTypPKY THEN 1 ELSE 0 END  -- PrimaryKey
    DECLARE @OtpUKY    bit             ; SET @OtpUKY = CASE WHEN @OupTyp = @ObjTypUKY THEN 1 ELSE 0 END  -- UniqueKey
    DECLARE @OtpIND    bit             ; SET @OtpIND = CASE WHEN @OupTyp = @ObjTypIND THEN 1 ELSE 0 END  -- Index
    DECLARE @OtpFKY    bit             ; SET @OtpFKY = CASE WHEN @OupTyp = @ObjTypFKY THEN 1 ELSE 0 END  -- ForeignKey
    DECLARE @OtpDEF    bit             ; SET @OtpDEF = CASE WHEN @OupTyp = @ObjTypDEF THEN 1 ELSE 0 END  -- Default
    DECLARE @OtpCHK    bit             ; SET @OtpCHK = CASE WHEN @OupTyp = @ObjTypCHK THEN 1 ELSE 0 END  -- Check
    DECLARE @OtpDDL    bit             ; SET @OtpDDL = CASE WHEN @OupTyp = @ObjTypDDL THEN 1 ELSE 0 END  -- DataDict
    DECLARE @OtpSCP    bit             ; SET @OtpSCP = CASE WHEN @OupTyp = @ObjTypSCP THEN 1 ELSE 0 END  -- Script
    DECLARE @OtpVDN    bit             ; SET @OtpVDN = CASE WHEN @OupTyp = @ObjTypVDN THEN 1 ELSE 0 END  -- VB.NET
    DECLARE @OtpUNK    bit             ; SET @OtpUNK = CASE WHEN @OupTyp = @ObjTypUNK THEN 1 ELSE 0 END  -- Unknown
    DECLARE @OupPrm    bit             ; SET @OupPrm = CASE WHEN @OupTyp IN (@ObjTypUSP,@ObjTypTRG,@ObjTypUFN) THEN 1 ELSE 0 END
    IF @DbgFlg = 1 OR 0=9 SELECT OJY='OJY',OupObj=LEFT(@OupObj,30),OupTyp=@OupTyp,OupPrm=@OupPrm,OtpTBL=@OtpTBL,OtpVEW=@OtpVEW,OtpUSP=@OtpUSP,OtpTRG=@OtpTRG,OtpUFN=@OtpUFN,OtpPKY=@OtpPKY,OtpUKY=@OtpUKY,OtpIND=@OtpIND,OtpFKY=@OtpFKY,OtpDEF=@OtpDEF,OtpCHK=@OtpCHK,OtpDDL=@OtpDDL,OtpSCP=@OtpSCP,OtpVDN=@OtpVDN,OtpUNK=@OtpUNK
 
    --##############################################################################################
 

    --ZZZ@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@XGM


    --##############################################################################################

    -- Assign developer name to fixed width
    DECLARE @DvpTxt    char(12)        ; SET @DvpTxt    = @DvpNam             -- Developer name
 
    --##############################################################################################

    -- Adjust column prefix values based on table type and name
    SET @TblCpx = CASE
        WHEN @TctLKP = 1 AND @TblNam LIKE '%Status' THEN 'Status'
        ELSE @TblCpx
    END

    --##############################################################################################

    -- Declare Statement object variables
    DECLARE @StmTyp    char(21)        ; SET @StmTyp    = ''
    DECLARE @StmIdn    char(9)         ; SET @StmIdn    = ''
    DECLARE @StmOup    char(9)         ; SET @StmOup    = ''

    -- Declare Column datatype variables
    DECLARE @TypNam    char(200)       ; SET @TypNam    = ''
    DECLARE @TypTxt    varchar(21)     ; SET @TypTxt    = ''

    -- Declare PrimaryKeyList object variables
    DECLARE @PkyNam    sysname         ; SET @PkyNam    = ''
    DECLARE @PkyFch    varchar(200)    ; SET @PkyFch    = ''
    DECLARE @PkyWhr    varchar(200)    ; SET @PkyWhr    = ''
    DECLARE @UseFst    bit             ; SET @UseFst    = 0

    -- Declare ForeignKeyList object variables
    DECLARE @UspNam    sysname         ; SET @UspNam    = ''
    DECLARE @FkyNam    sysname         ; SET @FkyNam    = ''
    DECLARE @RkyNam    sysname         ; SET @RkyNam    = ''
 
    --##############################################################################################

    -- Assign alternative aliases
    /*----------------------------------------------------------------------------------------------
        DECLARE @OUP varchar(10); EXEC ut_zzNAM TBL,ALS,AL0,zzz_TEST01,@OUP OUTPUT; SELECT 'TBL-ALS-AL0'=@OUP
        DECLARE @OUP varchar(10); EXEC ut_zzNAM TBL,ALS,LB3,zzz_TEST01,@OUP OUTPUT; SELECT 'TBL-ALS-LB3'=@OUP
        DECLARE @OUP varchar(10); EXEC ut_zzNAM TBL,ALS,RB3,zzz_TEST01,@OUP OUTPUT; SELECT 'TBL-ALS-RB3'=@OUP
    ----------------------------------------------------------------------------------------------*/
    IF @OupFmt IN ('XXX') BEGIN
        EXEC ut_zzNAM TBL,ALS,LB3,@TblNam,@TblAls OUTPUT
    END
 
    --##############################################################################################

    -- Assign minimum variable length
    DECLARE @VarLex    varchar(2)      ; SET @VarLex    = CAST(@MinLenVAR AS varchar(2))
    DECLARE @CmtVln    smallint        ; SET @CmtVln    = 0

    -- Declare Standard Return Variable tracking variables
    DECLARE @RetNam    varchar(50)     ; SET @RetNam    = 'RetVal' 
    DECLARE @RetTxt    char(200)       ; SET @RetTxt    = @RetNam 
    DECLARE @RetDtp    varchar(20)     ; SET @RetDtp    = 'int' 
    DECLARE @RetVal    varchar(50)     ; SET @RetVal    = '0'
    DECLARE @RetFln    smallint        ; SET @RetFln    = LEN(@RetNam)
    DECLARE @RetVln    smallint        ; SET @RetVln    = 0
    DECLARE @RetTln    smallint        ; SET @RetTln    = LEN(@RetDtp)
    SET @RetVln = CASE WHEN @RetFln < @MinLenVAR THEN @MinLenVAR ELSE @RetFln END

    -- Declare Function Return Variable tracking variables
    DECLARE @UfnNam    varchar(50)     ; SET @UfnNam    = 'UfnVal' 
    DECLARE @UfnTxt    char(200)       ; SET @UfnTxt    = @UfnNam 
    DECLARE @UfnDtp    varchar(20)     ; SET @UfnDtp    = 'int' 
    DECLARE @UfnVal    varchar(50)     ; SET @UfnVal    = '0'
    DECLARE @UfnFln    smallint        ; SET @UfnFln    = LEN(@UfnNam)
    DECLARE @UfnVln    smallint        ; SET @UfnVln    = 0
    DECLARE @UfnTln    smallint        ; SET @UfnTln    = LEN(@UfnDtp)
    SET @UfnVln = CASE WHEN @UfnFln < @MinLenVAR THEN @MinLenVAR ELSE @UfnFln END

    -- Initialize DataType constants
    DECLARE @NT        varchar(100)    ; EXEC ut_zzTYP VRN,DTK,NUM, NT, @NT OUTPUT
    DECLARE @DT        varchar(100)    ; EXEC ut_zzTYP VRN,DTK,DAT, DT, @DT OUTPUT
    DECLARE @CT        varchar(100)    ; EXEC ut_zzTYP VRN,DTK,TXT, CT, @CT OUTPUT
    DECLARE @ST        varchar(100)    ; EXEC ut_zzTYP VRN,DTK,APX, ST, @ST OUTPUT
    DECLARE @YT        varchar(100)    ; EXEC ut_zzTYP VRN,DTK,SYS, YT, @YT OUTPUT

    -- Track new objects
    DECLARE @DecTln    int             ; SET @DecTln    = 13
 
    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Assign output values
    ------------------------------------------------------------------------------------------------
    IF LEN(@OupFmt) = 0 SET @OupFmt = LEFT(@BldCOD,3)
    IF LEN(@OupSfx) = 0 SET @OupSfx = @OupFmt
    IF LEN(@DefTyp) = 0 SET @DefTyp = RIGHT(@BldCOD,3)
 
    -- Set name format
    SET @NamFmt = @DefTyp

    --##############################################################################################
 
    ------------------------------------------------------------------------------------------------
    -- Initialize module level objects
    ------------------------------------------------------------------------------------------------
    DECLARE @SrcAls    varchar(10)     ; SET @SrcAls    = CASE WHEN LEN(@StdTx2) > 0 THEN @StdTx2                   ELSE @TblAls END
    DECLARE @QryHnt    varchar(200)    ; SET @QryHnt    = CASE WHEN LEN(@StdTx3) > 0 THEN ' WITH ('+@StdTx3+')' ELSE ''      END
    DECLARE @JnnHnt    varchar(200)    ; SET @JnnHnt    = CASE WHEN LEN(@StdTx3) > 0 THEN ' WITH ('+@StdTx3+')' ELSE ''      END
    IF LEN(@SrcAls) > 0 SET @DspAls = @SrcAls
 
    ------------------------------------------------------------------------------------------------
    -- Set module level flags and defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @MaxFld    sysname         ; SET @MaxFld    = '' 
    DECLARE @MaxTbl    sysname         ; SET @MaxTbl    = '' 
    DECLARE @MaxFln    smallint        ; SET @MaxFln    = 0

    DECLARE @HstObj    sysname         ; SET @HstObj    = ''
    DECLARE @VarLin    varchar(1000)   ; SET @VarLin    = ''
    DECLARE @FldLin    varchar(1000)   ; SET @FldLin    = ''

    DECLARE @MaxLn1    smallint        ; SET @MaxLn1    = 0
    DECLARE @MaxLn2    smallint        ; SET @MaxLn2    = 0
    DECLARE @MaxLn3    smallint        ; SET @MaxLn3    = 0
    DECLARE @MaxLn4    smallint        ; SET @MaxLn4    = 0
    DECLARE @MinObj    smallint        ; SET @MinObj    = 40

    ------------------------------------------------------------------------------------------------
    -- Assign object flags
    ------------------------------------------------------------------------------------------------
    DECLARE @UspSEL    bit             ; SET @UspSEL    = CASE WHEN @OupObj LIKE 'usp_Select%' THEN 1 ELSE 0 END
    DECLARE @UspINS    bit             ; SET @UspINS    = CASE WHEN @OupObj LIKE 'usp_Insert%' THEN 1 ELSE 0 END
    DECLARE @UspUPD    bit             ; SET @UspUPD    = CASE WHEN @OupObj LIKE 'usp_Update%' THEN 1 ELSE 0 END
    DECLARE @UspDEL    bit             ; SET @UspDEL    = CASE WHEN @OupObj LIKE 'usp_Delete%' THEN 1 ELSE 0 END
 
    ------------------------------------------------------------------------------------------------
    -- Set table owner prefix
    ------------------------------------------------------------------------------------------------
    DECLARE @IsTemp    bit             ; SET @IsTemp    = CASE WHEN LEFT(@DspNam,1) = '#' THEN 1 ELSE 0 END
    DECLARE @IsVarb    bit             ; SET @IsVarb    = CASE WHEN LEFT(@DspNam,1) = '@' THEN 1 ELSE 0 END
    DECLARE @OwnPfx    varchar(10)     ; SET @OwnPfx    = CASE WHEN @IsTemp = 1 THEN '#' WHEN @IsVarb = 1 THEN '@' ELSE 'dbo.' END
    SET @IsVarb = CASE WHEN @BldLST IN ('TMPVAR') THEN 1 ELSE @IsVarb END

    ------------------------------------------------------------------------------------------------
    -- Set Object reference name
    ------------------------------------------------------------------------------------------------
    DECLARE @RefBas    sysname         ; SET @RefBas    = CASE 
        WHEN @IsVarb = 1 THEN @DspBas
        WHEN @IsTemp = 1 THEN @DspBas
        ELSE                  @DspBas
    END
    DECLARE @RefNam    sysname         ; SET @RefNam    = CASE 
        WHEN @IsVarb = 1 THEN @DspBas
        WHEN @IsTemp = 1 THEN @DspBas
        ELSE                  @DspNam
    END
    DECLARE @RefObj    sysname         ; SET @RefObj    = CASE 
        WHEN @IsVarb = 1 THEN '@'+@DspBas
        WHEN @IsTemp = 1 THEN '#'+@DspBas
        ELSE               'dbo.'+@DspNam
    END
    DECLARE @RefDsc    varchar(100)    ; SET @RefDsc    = CASE 
        WHEN @IsVarb = 1 THEN 'variable table'
        WHEN @IsTemp = 1 THEN 'temp table'
        ELSE                  'standard table'
    END
    DECLARE @RefOup    sysname         ; SET @RefOup    = 'dbo.'+@OupObj
    DECLARE @RefCod    varchar(3)      ; SET @RefCod    = UPPER(LEFT(REPLACE(@RefObj,'#',''),3))
    DECLARE @RefPfx    varchar(3)      ; SET @RefPfx    = LEFT(@RefCod,1)+LOWER(RIGHT(@RefCod,2))
    DECLARE @TmpObj    sysname         ; SET @TmpObj    = ''

    ------------------------------------------------------------------------------------------------
    -- Standard field variables
    ------------------------------------------------------------------------------------------------
    DECLARE @XbyVar    varchar(50)     ; SET @XbyVar    = 'ExpiredBy'
    DECLARE @XbySig    bit             ; SET @XbySig    = 0  
    DECLARE @XbyStm    bit             ; SET @XbyStm    = 0 
    DECLARE @XbyHst    bit             ; SET @XbyHst    = 0 
    DECLARE @XonVar    varchar(50)     ; SET @XonVar    = 'ExpiredOn'
    DECLARE @XonSig    bit             ; SET @XonSig    = 0  
    DECLARE @XonStm    bit             ; SET @XonStm    = 0 
    DECLARE @XonHst    bit             ; SET @XonHst    = 0 

    DECLARE @CbyVar    varchar(50)     ; SET @CbyVar    = 'CreatedBy'
    DECLARE @CbySig    bit             ; SET @CbySig    = 0  
    DECLARE @CbyStm    bit             ; SET @CbyStm    = 0 
    DECLARE @CbyHst    bit             ; SET @CbyHst    = 0 
    DECLARE @ConVar    varchar(50)     ; SET @ConVar    = 'CreatedOn'
    DECLARE @ConSig    bit             ; SET @ConSig    = 0  
    DECLARE @ConStm    bit             ; SET @ConStm    = 0 
    DECLARE @ConHst    bit             ; SET @ConHst    = 0 

    DECLARE @UbyVar    varchar(50)     ; SET @UbyVar    = 'UpdatedBy'
    DECLARE @UbySig    bit             ; SET @UbySig    = 0  
    DECLARE @UbyStm    bit             ; SET @UbyStm    = 0 
    DECLARE @UbyHst    bit             ; SET @UbyHst    = 0 
    DECLARE @UonVar    varchar(50)     ; SET @UonVar    = 'UpdatedOn'
    DECLARE @UonSig    bit             ; SET @UonSig    = 0  
    DECLARE @UonStm    bit             ; SET @UonStm    = 0 
    DECLARE @UonHst    bit             ; SET @UonHst    = 0 

    DECLARE @HbyVar    varchar(50)     ; SET @HbyVar    = 'HistoryBy'
    DECLARE @HbySig    bit             ; SET @HbySig    = 0  
    DECLARE @HbyStm    bit             ; SET @HbyStm    = 0 
    DECLARE @HbyHst    bit             ; SET @HbyHst    = 0 
    DECLARE @HonVar    varchar(50)     ; SET @HonVar    = 'HistoryOn'
    DECLARE @HonSig    bit             ; SET @HonSig    = 0  
    DECLARE @HonStm    bit             ; SET @HonStm    = 0 
    DECLARE @HonHst    bit             ; SET @HonHst    = 0 

    DECLARE @HttVar    varchar(50)     ; SET @HttVar    = 'HstTrnTyp'
    DECLARE @HttSig    bit             ; SET @HttSig    = 0  
    DECLARE @HttStm    bit             ; SET @HttStm    = 0 
    DECLARE @HttHst    bit             ; SET @HttHst    = 0 

    DECLARE @HtuVar    varchar(50)     ; SET @HtuVar    = 'HstTrnUtc'
    DECLARE @HtuSig    bit             ; SET @HtuSig    = 0  
    DECLARE @HtuStm    bit             ; SET @HtuStm    = 0 
    DECLARE @HtuHst    bit             ; SET @HtuHst    = 0 

    ------------------------------------------------------------------------------------------------
    -- Parameter field variables
    ------------------------------------------------------------------------------------------------
    DECLARE @AlsVar    varchar(50)     ; SET @AlsVar    = 'AddList'
    DECLARE @AlsTxt    char(50)        ; SET @AlsTxt    = @AlsVar 
    DECLARE @AlsSig    bit             ; SET @AlsSig    = 0 
    DECLARE @AlsStm    bit             ; SET @AlsStm    = 0 
    DECLARE @AlsHst    bit             ; SET @AlsHst    = 0 

    DECLARE @DlsVar    varchar(50)     ; SET @DlsVar    = 'DelList'
    DECLARE @DlsTxt    char(50)        ; SET @DlsTxt    = @DlsVar 
    DECLARE @DlsSig    bit             ; SET @DlsSig    = 0 
    DECLARE @DlsStm    bit             ; SET @DlsStm    = 0 
    DECLARE @DlsHst    bit             ; SET @DlsHst    = 0 

    DECLARE @TmdVar    varchar(50)     ; SET @TmdVar    = 'TestMode'
    DECLARE @TmdTxt    char(50)        ; SET @TmdTxt    = @TmdVar 
    DECLARE @TmdSig    bit             ; SET @TmdSig    = 0 
    DECLARE @TmdStm    bit             ; SET @TmdStm    = 0 
    DECLARE @TmdHst    bit             ; SET @TmdHst    = 0 

    DECLARE @CldVar    varchar(50)     ; SET @CldVar    = 'ClearOld'
    DECLARE @CldTxt    char(50)        ; SET @CldTxt    = @CldVar 
    DECLARE @CldSig    bit             ; SET @CldSig    = 0 
    DECLARE @CldStm    bit             ; SET @CldStm    = 0 
    DECLARE @CldHst    bit             ; SET @CldHst    = 0 

    DECLARE @TkhVar    varchar(50)     ; SET @TkhVar    = 'TrakHist'
    DECLARE @TkhTxt    char(50)        ; SET @TkhTxt    = @TkhVar 
    DECLARE @TkhSig    bit             ; SET @TkhSig    = 0 
    DECLARE @TkhStm    bit             ; SET @TkhStm    = 0 
    DECLARE @TkhHst    bit             ; SET @TkhHst    = 0 
    DECLARE @TkhDec    bit             ; SET @TkhDec    = 0 

    DECLARE @IdnClm    varchar(50)     ; SET @IdnClm    = ''
    DECLARE @HasIdn    bit             ; SET @HasIdn    = 0
    DECLARE @OupClm    varchar(50)     ; SET @OupClm    = ''
    DECLARE @HasOup    bit             ; SET @HasOup    = 0
    DECLARE @HasQot    bit             ; SET @HasQot    = 0
    
    ------------------------------------------------------------------------------------------------
    -- Population variables (POP and TMP)
    ------------------------------------------------------------------------------------------------
    DECLARE @QT1       char(06)        ; SET @QT1       = ''
    DECLARE @QT2       char(06)        ; SET @QT2       = ''
    DECLARE @LQT       varchar(06)     ; SET @LQT       = """'""+"
    DECLARE @RQT       varchar(06)     ; SET @RQT       = "+""'"""
    DECLARE @DFV       varchar(10)     ; SET @DFV       = ''
    DECLARE @DFN       varchar(10)     ; SET @DFN       = ''
    DECLARE @DFC       varchar(10)     ; SET @DFC       = ''
    DECLARE @DFX       varchar(200)    ; SET @DFX       = ''
    DECLARE @LNV       sysname         ; SET @LNV       = ''

    DECLARE @FLQ       varchar(200)    ; SET @FLQ       = ''
    
    DECLARE @XLV       tinyint         ; SET @XLV       = 0
    DECLARE @X01       varchar(max)    ; SET @X01       = ''
    DECLARE @X02       varchar(max)    ; SET @X02       = ''
    DECLARE @X03       varchar(max)    ; SET @X03       = ''
    DECLARE @X04       varchar(max)    ; SET @X04       = ''
    DECLARE @X05       varchar(max)    ; SET @X05       = ''
    DECLARE @X06       varchar(max)    ; SET @X06       = ''
    DECLARE @X07       varchar(max)    ; SET @X07       = ''

    DECLARE @PsqLN1    tinyint         ; SET @PsqLN1    = 01  -- Placeholder for first lines
    DECLARE @PsqDCB    tinyint         ; SET @PsqDCB    = 02  -- DECLARE Base variables
    DECLARE @PsqDCL    tinyint         ; SET @PsqDCL    = 03  -- DECLARE Line variables
    DECLARE @PsqDLV    tinyint         ; SET @PsqDLV    = 04  -- Assign dynamic length variables
    DECLARE @PsqINS    tinyint         ; SET @PsqINS    = 05  -- Build INSERT Statement
    DECLARE @PsqTLN    tinyint         ; SET @PsqTLN    = 06  -- Build Title Line
    DECLARE @PsqTTX    tinyint         ; SET @PsqTTX    = 07  -- Build Title Text
    DECLARE @PsqHDR    tinyint         ; SET @PsqHDR    = 08  -- PRINT Header Lines
    DECLARE @PsqBEG    tinyint         ; SET @PsqBEG    = 09  -- IF Rows > 0 BEGIN
    DECLARE @PsqSEL    tinyint         ; SET @PsqSEL    = 10  -- SELECT [Records]
    DECLARE @PsqELS    tinyint         ; SET @PsqELS    = 11  -- END ELSE BEGIN
    DECLARE @PsqTPL    tinyint         ; SET @PsqTPL    = 12  -- PRINT [Template]
    DECLARE @PsqEND    tinyint         ; SET @PsqEND    = 13  -- END
    DECLARE @PsqFTR    tinyint         ; SET @PsqFTR    = 14  -- PRINT Footer Lines
    DECLARE @PsqLN2    tinyint         ; SET @PsqLN2    = 15  -- Placeholder for last lines

    DECLARE @PopMrx    varchar(20)     ; SET @PopMrx    = ''

    DECLARE @PpxFld    varchar(max)    ; SET @PpxFld    = ''
    DECLARE @PpxVal    varchar(max)    ; SET @PpxVal    = ''
    DECLARE @PpxVar    varchar(max)    ; SET @PpxVar    = ''
    DECLARE @PpxStm    varchar(max)    ; SET @PpxStm    = ''
    DECLARE @PpxLn1    varchar(max)    ; SET @PpxLn1    = ''
    DECLARE @PpxLn2    varchar(max)    ; SET @PpxLn2    = ''
    DECLARE @PpxLn3    varchar(max)    ; SET @PpxLn3    = ''

    ------------------------------------------------------------------------------------------------
    -- Miscellaneous variables
    ------------------------------------------------------------------------------------------------
    DECLARE @ExsVar    varchar(50)     ; SET @ExsVar    = 'RecExists'
    DECLARE @ExsTxt    char(50)        ; SET @ExsTxt    = @ExsVar 

    DECLARE @ChgNam    varchar(20)     ; SET @ChgNam   = 'ChgKeys'

    DECLARE @ModCur    smallint        ; SET @ModCur    = 0              -- Track module OPEN cursor statements
    DECLARE @ModBeg    smallint        ; SET @ModBeg    = 0              -- Track module BEGIN statements
    DECLARE @FncBeg    smallint        ; SET @FncBeg    = 0              -- Track function BEGIN statements

    DECLARE @RowCnt    int             ; SET @RowCnt    = 0              -- Track global row count

    DECLARE @TmpAls    varchar(10)     ; SET @TmpAls    = ''
    DECLARE @TmpFln    smallint        ; SET @TmpFln    = 0

    DECLARE @CRS       cursor

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Parameter values
    ------------------------------------------------------------------------------------------------
    DECLARE @AlnFlg    bit             ; SET @AlnFlg    = 0              -- Should parameter data types be aligned?
    DECLARE @TrkHst    bit             ; SET @TrkHst    = 0              -- Is history being tracked?

    ------------------------------------------------------------------------------------------------
    -- Formatting comments
    ------------------------------------------------------------------------------------------------
    DECLARE @CPD       char(200)       ; SET @CPD       = ''             -- Padding for parameter text
    DECLARE @SCS       smallint        ; SET @SCS       = 10             -- Space for parameter text

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Working text list temp table
    ------------------------------------------------------------------------------------------------
    DECLARE @TxtID     int
    DECLARE @TxtLvl    int
    DECLARE @TxtLin    varchar(8000)
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #TxtLst (               -- DROP TABLE dbo.zzz_TxtLst; CREATE TABLE dbo.zzz_TxtLst (
    ------------------------------------------------------------------------------------------------
        TxtID                                             int                  NOT NULL IDENTITY,
        TxtLvl                                            int                  NOT NULL DEFAULT 0,
        TxtLin                                            varchar(max)             NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------


    --DFX@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- Object definition parameter variables (JRS)
    ------------------------------------------------------------------------------------------------
    DECLARE @TblLst    sysname         ; SET @TblLst    = @InpObj        -- Source Object list (comma delimited)
    DECLARE @OupRSJ    varchar(11)     ; SET @OupRSJ    = ''             -- Output code
    DECLARE @ClmLst    varchar(max)    ; SET @ClmLst    = ''             -- Source Column list (comma delimited)
    DECLARE @DtpLst    varchar(max)    ; SET @DtpLst    = ''             -- Source DataType list (comma delimited)
    DECLARE @BlkLst    varchar(max)    ; SET @BlkLst    = ''             -- Block these Objects (comma delimited list)
    DECLARE @StpLst    varchar(max)    ; SET @StpLst    = ''             -- Stop at these Objects (comma delimited list)
    DECLARE @SetFmt    tinyint         ; SET @SetFmt    = 1              -- Set Format values in result set
    DECLARE @AsnDef    tinyint         ; SET @AsnDef    = 1              -- Assign default values/records
    DECLARE @AllLst    sysname         ; SET @AllLst    = '%'            -- Source Object list (comma delimited)
    DECLARE @PrmLst    sysname         ; SET @PrmLst    = CASE           -- Permissions Object list
        WHEN @OupFmt = 'TBL' THEN @DspTbl
        ELSE @OupObj
    END
    ------------------------------------------------------------------------------------------------


    --DFN@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- Initialize default assignment flags
    ------------------------------------------------------------------------------------------------
    DECLARE @AsnDefNON smallint        ; SET @AsnDefNON = 0              -- 0 - No default assignment
    DECLARE @AsnDefMTY smallint        ; SET @AsnDefMTY = 1              -- 1 - Assign defaults when temp table is empty
    DECLARE @AsnDefABS smallint        ; SET @AsnDefABS = 2              -- 2 - Assign defaults absolutely
    ------------------------------------------------------------------------------------------------
    DECLARE @AsnDefDST tinyint         ; SET @AsnDefDST = 3              -- 3 - Assign default SProcs for table
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare common definition tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE @DfnID     smallint
    DECLARE @DfnCls    char(3)
    DECLARE @DfnExs    bit
    DECLARE @DfnFix    bit
    DECLARE @DfnFmt    smallint
    DECLARE @DfnFmx    varchar(3)
    DECLARE @DfnStd    sysname
    DECLARE @DfnCur    CURSOR
    ------------------------------------------------------------------------------------------------
    DECLARE @ConTbl    sysname
    DECLARE @ConNam    sysname
    DECLARE @ConClm    sysname
    DECLARE @ConDsc    varchar(200)
    DECLARE @ConKys    varchar(1000)
    DECLARE @ConKyx    varchar(1000)
    ------------------------------------------------------------------------------------------------
    --CLARE @TblNam    sysname
    ------------------------------------------------------------------------------------------------
    --CLARE @TblID     int
    DECLARE @ConID     int
    DECLARE @ClmID     int
    ------------------------------------------------------------------------------------------------
    DECLARE @KeyOrd    smallint
    DECLARE @KeyClm    sysname
    DECLARE @IndDir    varchar(5)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Object definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopObj    bit                ; SET @PopObj    = 1 --@InpExs        -- Populate System Objects
    DECLARE @AdfObj    bit                ; SET @AdfObj    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @ObjNAM    sysname            ; SET @ObjNAM    = ''             -- Object name
    DECLARE @ObjTxt    sysname            ; SET @ObjTxt    = ''             -- Object text
    DECLARE @ObjCls    char(3)            ; SET @ObjCls    = ''             -- Object class
    DECLARE @ObjPrp    varchar(100)       ; SET @ObjPrp    = ''             -- Object property
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Table definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopTbl    bit                ; SET @PopTbl    = 0              -- Populate User Tables
    DECLARE @AdfTbl    bit                ; SET @AdfTbl    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    --CLARE @TblNam    sysname
    DECLARE @TblSiz    dec(15,2)
    DECLARE @RecQty    int
    DECLARE @RecSiz    dec(15,0)
    DECLARE @HasPky    bit
    DECLARE @UkyQty    smallint
    DECLARE @FkyQty    smallint
    DECLARE @RkyQty    smallint
    DECLARE @IndQty    smallint
    DECLARE @DefQty    smallint
    DECLARE @ChkQty    smallint
    DECLARE @ClxNam    sysname
    DECLARE @HasHtk    bit
    DECLARE @HasDim    bit
    DECLARE @HasFct    bit
    DECLARE @HasLkd    bit
    DECLARE @HasDsb    bit
    DECLARE @HasDlt    bit
    DECLARE @HasLok    bit
    DECLARE @HasCrt    bit
    DECLARE @HasUpd    bit
    DECLARE @HasExp    bit
    DECLARE @HasDel    bit
    DECLARE @HasHst    bit
    DECLARE @HasAud    bit
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare View definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopVew    bit             ; SET @PopVew    = 0              -- Populate User Views
    DECLARE @AdfVew    bit             ; SET @AdfVew    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @VewNam    sysname
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare SProc definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopUsp    bit             ; SET @PopUsp    = 0              -- Populate User Stored Procedures
    DECLARE @AdfUsp    bit             ; SET @AdfUsp    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    --CLARE @UspNam    sysname
    DECLARE @DepUsp    smallint
    DECLARE @DepFnc    smallint
    DECLARE @DepTrg    smallint
    DECLARE @UseTbl    smallint
    DECLARE @UseVew    smallint
    DECLARE @UseUsp    smallint
    DECLARE @UseFnc    smallint
    DECLARE @UseTrg    smallint
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare SProc text lines variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopUln    bit             ; SET @PopUln    = 0              -- Populate 
    DECLARE @AdfUln    bit             ; SET @AdfUln    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @LinID     int
    DECLARE @LinUsp    sysname
    DECLARE @LinNbr    smallint
    DECLARE @LinSqn    char(4)
    DECLARE @LinLen    smallint
    DECLARE @LinSiz    char(1)
    DECLARE @LinTxt    varchar(512)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Trigger definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopTrg    bit             ; SET @PopTrg    = 0              -- Populate 
    DECLARE @AdfTrg    bit             ; SET @AdfTrg    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare User Function definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopUfn    bit             ; SET @PopUfn    = 0              -- Populate 
    DECLARE @AdfUfn    bit             ; SET @AdfUfn    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Permissions definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopPrm    bit             ; SET @PopPrm    = 0              -- Populate 
    DECLARE @AdfPrm    bit             ; SET @AdfPrm    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PrmLvl    tinyint
    DECLARE @ProTyp    tinyint
    DECLARE @ProTxt    varchar(20)
    DECLARE @UsrNam    sysname
    DECLARE @ActItm    tinyint
    DECLARE @ActTxt    varchar(20)
    DECLARE @UsrDsc    varchar(80)
    ------------------------------------------------------------------------------------------------
    DECLARE @ActLst    varchar(300)
    DECLARE @ProDir    varchar(6)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Dependencies definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopDep    bit             ; SET @PopDep    = 0              -- Populate 
    DECLARE @AdfDep    bit             ; SET @AdfDep    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Field definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopFld    bit             ; SET @PopFld    = 0              -- Populate 
    DECLARE @AdfFld    bit             ; SET @AdfFld    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Column definition variables                              EXEC ut_zzSQL DTV,zzz_ClmDfn
    ------------------------------------------------------------------------------------------------
    DECLARE @PopClm    bit                ; SET @PopClm    = 0              -- Populate 
    DECLARE @AdfClm    bit                ; SET @AdfClm    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @ClmLvl    tinyint
    DECLARE @ClmObj    sysname
    DECLARE @ClmOrd    smallint
    DECLARE @ClmNam    sysname
    --------------------------------
    DECLARE @ClmUtp    int
    DECLARE @ClmUtx    sysname
    DECLARE @ClmStp    tinyint
    DECLARE @ClmStx    sysname
    DECLARE @ClmDtp    int
    DECLARE @ClmDTX    sysname
    --------------------------------
    DECLARE @ClmLen    smallint
    DECLARE @ClmWid    smallint
    DECLARE @ClmPrc    tinyint
    DECLARE @ClmScl    tinyint
    DECLARE @ClmDsp    smallint
    --------------------------------
    DECLARE @ClmDct    varchar(3)
    DECLARE @ClmQot    bit
    DECLARE @ClmNul    bit
    DECLARE @ClmIdn    bit
    DECLARE @ClmCmp    bit
    DECLARE @ClmOup    bit
    DECLARE @ClmVwd    bit
    DECLARE @ClmMax    bit
    DECLARE @ClmAud    bit
    --------------------------------
    DECLARE @ClmPky    tinyint
    DECLARE @ClmUky    tinyint
    DECLARE @ClmXky    tinyint
    DECLARE @ClmFky    bit
    --------------------------------
    DECLARE @ClmDef    bit
    DECLARE @ClmDfv    nvarchar(max)
    DECLARE @ClmEmv    varchar(100)
    DECLARE @ClmCpx    nvarchar(max)
    ------------------------------------------------------------------------------------------------
    DECLARE @ClmSiz    smallint  -- Deprecate!
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Parameter definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopPar    bit             ; SET @PopPar    = 0              -- Populate 
    DECLARE @AdfPar    bit             ; SET @AdfPar    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @ParObj    sysname
    DECLARE @ParOrd    smallint
    DECLARE @ParNam    sysname
    DECLARE @ParDtp    varchar(21)
    DECLARE @ParUtp    sysname
    DECLARE @ParLen    smallint
    DECLARE @ParQot    bit
    DECLARE @ParOup    bit
    DECLARE @ParDef    bit
    DECLARE @ParDfx    varchar(4000)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Default Column definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopDcv    bit             ; SET @PopDcv    = 0              -- Populate 
    DECLARE @AdfDcv    bit             ; SET @AdfDcv    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @DcvLvl    tinyint
    DECLARE @DcvObj    sysname
    DECLARE @DcvOrd    smallint
    DECLARE @DcvNam    sysname
    DECLARE @DcvDtp    varchar(21)
    DECLARE @DcvSiz    smallint
    DECLARE @DcvLen    smallint
    DECLARE @DcvDct    varchar(3)
    DECLARE @DcvQot    bit
    DECLARE @DcvNul    bit
    DECLARE @DcvIdn    bit
    DECLARE @DcvCmp    bit
    DECLARE @DcvOup    bit
    DECLARE @DcvPky    tinyint
    DECLARE @DcvUky    tinyint
    DECLARE @DcvFky    bit
    DECLARE @DcvInd    bit
    DECLARE @DcvDef    bit
    DECLARE @DcvChk    bit
    DECLARE @DcvVwd    bit
    DECLARE @DcvAud    bit
    DECLARE @DcvVal    varchar(100)
    DECLARE @DcvDfv    varchar(100)
    DECLARE @DcvVfx    varchar(4000)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare PrimaryKey, UniqueKey, Standard Index, Dynamic Statistics definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopPky    bit             ; SET @PopPky    = 0              -- Populate Primary Keys
    DECLARE @AdfPky    bit             ; SET @AdfPky    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PopUky    bit             ; SET @PopUky    = 0              -- Populate 
    DECLARE @AdfUky    bit             ; SET @AdfUky    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PopInd    bit             ; SET @PopInd    = 0              -- Populate 
    DECLARE @AdfInd    bit             ; SET @AdfInd    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PopStt    bit             ; SET @PopStt    = 0              -- Populate 
    DECLARE @AdfStt    bit             ; SET @AdfStt    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @IsUniq    bit
    DECLARE @IsClus    bit
    DECLARE @FilFct    smallint
    ------------------------------------------------------------------------------------------------
    DECLARE @UnqTxt    varchar(40)
    DECLARE @CluTxt    varchar(40)
    DECLARE @FilTxt    varchar(40)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Foreign Keys/References definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopFky    bit             ; SET @PopFky    = 0              -- Populate Foreign Keys
    DECLARE @AdfFky    bit             ; SET @AdfFky    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PopRky    bit             ; SET @PopRky    = 0              -- Populate Foreign References
    DECLARE @AdfRky    bit             ; SET @AdfRky    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @FkyTbl    sysname
    DECLARE @FkyKys    varchar(1000)
    DECLARE @RkyTbl    sysname
    DECLARE @RkyKys    varchar(1000)
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Declare Default/Check constraints definitions variables
    ------------------------------------------------------------------------------------------------
    DECLARE @PopDef    bit             ; SET @PopDef    = 0              -- Populate 
    DECLARE @AdfDef    bit             ; SET @AdfDef    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @PopChk    bit             ; SET @PopChk    = 0              -- Populate 
    DECLARE @AdfChk    bit             ; SET @AdfChk    = 0              -- Assign Defaults
    ------------------------------------------------------------------------------------------------
    DECLARE @ConTxt    varchar(1000)
    ------------------------------------------------------------------------------------------------


    --VBA@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  VBA Values
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaID     smallint
    DECLARE @OutFlg    bit
    DECLARE @IdnFlg    bit
    DECLARE @UpcFlg    bit
    DECLARE @CboFlg    bit
    DECLARE @ChkFlg    bit
    DECLARE @DirPrp    tinyint
    DECLARE @VarCat    varchar(10)
    DECLARE @VarPfx    varchar(10)
    DECLARE @VarNam    sysname
    DECLARE @VarDtp    varchar(30)
    DECLARE @VarDfc    varchar(30)
    DECLARE @VarVal    varchar(30)
    DECLARE @VarNul    varchar(30)
    DECLARE @CtlPfx    varchar(10)
    DECLARE @CtlNam    sysname
    DECLARE @PrmDtp    varchar(20)
    DECLARE @PrmDir    varchar(20)
    DECLARE @PrmLen    varchar(10)
    ------------------------------------------------------------------------------------------------
    -- Cursor Tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaCnt    smallint     ; SET @VbaCnt    = 0
    DECLARE @VbaVln    smallint     ; SET @VbaVln    = 0
    DECLARE @VbaTln    smallint     ; SET @VbaTln    = 0
    DECLARE @VbaCln    smallint     ; SET @VbaCln    = 0
    DECLARE @VbaPln    smallint     ; SET @VbaPln    = 0
    DECLARE @VbaDln    smallint     ; SET @VbaDln    = 0
    DECLARE @VbaLln    smallint     ; SET @VbaLln    = 0
    ------------------------------------------------------------------------------------------------
    --CLARE @TmpFln    smallint     ; SET @TmpFln    = 0
    DECLARE @TmpVln    smallint     ; SET @TmpVln    = 0
    DECLARE @TmpTln    smallint     ; SET @TmpTln    = 0
    DECLARE @TmpCln    smallint     ; SET @TmpCln    = 0
    DECLARE @TmpPln    smallint     ; SET @TmpPln    = 0
    DECLARE @TmpDln    smallint     ; SET @TmpDln    = 0
    DECLARE @TmpLln    smallint     ; SET @TmpLln    = 0
    ------------------------------------------------------------------------------------------------
    DECLARE @FncTyp    varchar(20)  ; SET @FncTyp    = ''
    DECLARE @FncNam    sysname      ; SET @FncNam    = ''
    DECLARE @RunNam    sysname      ; SET @RunNam    = ''
    DECLARE @MthNam    sysname      ; SET @MthNam    = ''
    DECLARE @RtnVar    sysname      ; SET @RtnVar    = ''
    DECLARE @RtnDtp    sysname      ; SET @RtnDtp    = ''
    DECLARE @FstFld    sysname      ; SET @FstFld    = ''
    ------------------------------------------------------------------------------------------------
    -- VBA variable categories
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaCatTXT varchar(03)  ; SET @VbaCatTXT = 'TXT'
    DECLARE @VbaCatNUM varchar(03)  ; SET @VbaCatNUM = 'NUM'
    DECLARE @VbaCatBLN varchar(03)  ; SET @VbaCatBLN = 'BLN'
    DECLARE @VbaCatDAT varchar(03)  ; SET @VbaCatDAT = 'DAT'
    DECLARE @VbaCatVRN varchar(03)  ; SET @VbaCatVRN = 'VRN'
    ------------------------------------------------------------------------------------------------
    -- VBA variable default constants
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaDfcTXT varchar(09)  ; SET @VbaDfcTXT = 'gcNullTXT'
    DECLARE @VbaDfcNUM varchar(09)  ; SET @VbaDfcNUM = 'gcNullNUM'
    DECLARE @VbaDfcBLN varchar(09)  ; SET @VbaDfcBLN = 'gcNullBLN'
    DECLARE @VbaDfcDAT varchar(09)  ; SET @VbaDfcDAT = 'gcNullDAT'
    DECLARE @VbaDfcVRN varchar(09)  ; SET @VbaDfcVRN = 'gcNullVRN'
    ------------------------------------------------------------------------------------------------
    -- VBA variable default values
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaDfvTXT varchar(03)  ; SET @VbaDfvTXT = '""'
    DECLARE @VbaDfvNUM varchar(03)  ; SET @VbaDfvNUM = '0'
    DECLARE @VbaDfvBLN varchar(03)  ; SET @VbaDfvBLN = '0'
    DECLARE @VbaDfvDAT varchar(03)  ; SET @VbaDfvDAT = '0'
    DECLARE @VbaDfvVRN varchar(03)  ; SET @VbaDfvVRN = '""'
    ------------------------------------------------------------------------------------------------
    -- VBA variable null values
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaNulTXT varchar(03)  ; SET @VbaNulTXT = '""'
    DECLARE @VbaNulNUM varchar(03)  ; SET @VbaNulNUM = '0'
    DECLARE @VbaNulBLN varchar(03)  ; SET @VbaNulBLN = '0'
    DECLARE @VbaNulDAT varchar(03)  ; SET @VbaNulDAT = '0'
    DECLARE @VbaNulVRN varchar(03)  ; SET @VbaNulVRN = 'vbNULL'
    ------------------------------------------------------------------------------------------------
    -- VBA variable single quote
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaSqtTXT varchar(03)  ; SET @VbaSqtTXT = "'"
    DECLARE @VbaSqtNUM varchar(03)  ; SET @VbaSqtNUM = ''
    DECLARE @VbaSqtBLN varchar(03)  ; SET @VbaSqtBLN = ''
    DECLARE @VbaSqtDAT varchar(03)  ; SET @VbaSqtDAT = "'"
    DECLARE @VbaSqtVRN varchar(03)  ; SET @VbaSqtVRN = "'"
    ------------------------------------------------------------------------------------------------
    -- VBA variable double quote
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaDqtTXT varchar(03)  ; SET @VbaDqtTXT = '"'
    DECLARE @VbaDqtNUM varchar(03)  ; SET @VbaDqtNUM = ''
    DECLARE @VbaDqtBLN varchar(03)  ; SET @VbaDqtBLN = ''
    DECLARE @VbaDqtDAT varchar(03)  ; SET @VbaDqtDAT = '"'
    DECLARE @VbaDqtVRN varchar(03)  ; SET @VbaDqtVRN = '"'
    ------------------------------------------------------------------------------------------------
    -- VBA variable prefixes
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaPfxSTR varchar(03)  ; SET @VbaPfxSTR = 'str'
    DECLARE @VbaPfxBLN varchar(03)  ; SET @VbaPfxBLN = 'bln'
    DECLARE @VbaPfxINT varchar(03)  ; SET @VbaPfxINT = 'int'
    DECLARE @VbaPfxLNG varchar(03)  ; SET @VbaPfxLNG = 'lng'
    DECLARE @VbaPfxCUR varchar(03)  ; SET @VbaPfxCUR = 'cur'
    DECLARE @VbaPfxSGL varchar(03)  ; SET @VbaPfxSGL = 'sgl'
    DECLARE @VbaPfxDBL varchar(03)  ; SET @VbaPfxDBL = 'dbl'
    DECLARE @VbaPfxDAT varchar(03)  ; SET @VbaPfxDAT = 'dat'
    DECLARE @VbaPfxVRN varchar(03)  ; SET @VbaPfxVRN = 'vrn'
    ------------------------------------------------------------------------------------------------
    -- VBA variable data types
    ------------------------------------------------------------------------------------------------
    DECLARE @VbaDtpSTR varchar(11)  ; SET @VbaDtpSTR = 'String'
    DECLARE @VbaDtpBLN varchar(11)  ; SET @VbaDtpBLN = 'Boolean'
    DECLARE @VbaDtpINT varchar(11)  ; SET @VbaDtpINT = 'Integer'
    DECLARE @VbaDtpLNG varchar(11)  ; SET @VbaDtpLNG = 'Long'
    DECLARE @VbaDtpCUR varchar(11)  ; SET @VbaDtpCUR = 'Currency'
    DECLARE @VbaDtpSGL varchar(11)  ; SET @VbaDtpSGL = 'Single'
    DECLARE @VbaDtpDBL varchar(11)  ; SET @VbaDtpDBL = 'Double'
    DECLARE @VbaDtpDAT varchar(11)  ; SET @VbaDtpDAT = 'Date'
    DECLARE @VbaDtpVRN varchar(11)  ; SET @VbaDtpVRN = 'Variant'
    ------------------------------------------------------------------------------------------------
    -- VBA parameter direction
    ------------------------------------------------------------------------------------------------
    DECLARE @PrmDirINP varchar(20)  ; SET @PrmDirINP = 'adParamInput'
    DECLARE @PrmDirOUP varchar(20)  ; SET @PrmDirOUP = 'adParamOutput'
    DECLARE @PrmDirIOP varchar(20)  ; SET @PrmDirIOP = 'adParamInputOutput'
    DECLARE @PrmDirRTN varchar(20)  ; SET @PrmDirRTN = 'adParamRetunValue'
    DECLARE @PrmDirUNK varchar(20)  ; SET @PrmDirUNK = 'adParamUnknown'
    ------------------------------------------------------------------------------------------------
    -- FORM control prefixes
    ------------------------------------------------------------------------------------------------
    DECLARE @CtlPfxTXT varchar(03)  ; SET @CtlPfxTXT = 'txt'
    DECLARE @CtlPfxCBO varchar(03)  ; SET @CtlPfxCBO = 'cbo'
    DECLARE @CtlPfxCHK varchar(03)  ; SET @CtlPfxCHK = 'chk'
    DECLARE @CtlPfxCMD varchar(03)  ; SET @CtlPfxCMD = 'cmd'
    ------------------------------------------------------------------------------------------------
    -- FORM control object name
    ------------------------------------------------------------------------------------------------
    DECLARE @CtlTxtSTR varchar(20)  ; SET @CtlTxtSTR = 'TextBox'
    DECLARE @CtlTxtCBO varchar(20)  ; SET @CtlTxtCBO = 'ComboBox'
    DECLARE @CtlTxtCHK varchar(20)  ; SET @CtlTxtCHK = 'CheckBox'
    DECLARE @CtlTxtCMD varchar(20)  ; SET @CtlTxtCMD = 'CommandButton'
    ------------------------------------------------------------------------------------------------
    DECLARE @DtpNam    varchar(30)  ; SET @DtpNam    = ''
    ------------------------------------------------------------------------------------------------


    --RSJ@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- Display temp table values
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT
           'CUR'     = 'CUR'
           ,OBJ      = @PopObj
           ,TblLst   = LEFT(@TblLst,40)
           ,ClmLst   = LEFT(@ClmLst,40)
           ,DtpLst   = LEFT(@DtpLst,40)
           ,BlkLst   = LEFT(@BlkLst,40)
           ,StpLst   = LEFT(@StpLst,40)
           ,SetFmt   = LEFT(@SetFmt,40)
           ,AsnDef   = LEFT(@AsnDef,40)
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  System Objects
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #ObjDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #ObjDfn (               -- DROP TABLE dbo.zzz_ObjDfn; CREATE TABLE dbo.zzz_ObjDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ObjNam                                            sysname              NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_ObjDfn
    ------------------------------------------------------------------------------------------------
    IF @PopObj = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfObj; EXEC ut_zzRSJ OBJ,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ObjNam,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #ObjDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ObjNam, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ObjNam,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_ObjDefs CURSOR LOCAL FOR SELECT *                              FROM #ObjDfn ORDER BY ObjNam  -- Cursor
    DECLARE @ObjCnt smallint; SET @ObjCnt = (SELECT COUNT(*)                   FROM #ObjDfn)                 -- Record Count
    DECLARE @ObjFln smallint; SET @ObjFln = (SELECT ISNULL(MAX(LEN(ObjNam)),0) FROM #ObjDfn)                 -- Max FieldName Length
    DECLARE @ObjTln smallint; SET @ObjTln = (SELECT ISNULL(MAX(LEN(ObjNam)),0) FROM #ObjDfn)                 -- Max TableName Length
    DECLARE @ObjVln smallint; SET @ObjVln = CASE WHEN @ObjFln < @MinLenVAR THEN @MinLenVAR ELSE @ObjFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ObjDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#ObjDfn' = '#ObjDfn', ObjCnt = @ObjCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ObjNam = LEFT(dfn.ObjNam,60)
            ,DfnStd = dfn.DfnStd
        FROM
            #ObjDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Set object profile values
    ------------------------------------------------------------------------------------------------
    SET @DfnCls = ISNULL(@DfnCls,@SrcTyp)  -- @SrcTyp @OupTyp
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT
           'POP'      = 'POP'
           ,DfnCls    = @DfnCls
           ,ObjClsTBL = @ObjClsTBL
           ,ObjClsVEW = @ObjClsVew
           ,ObjClsUSP = @ObjClsUsp
           ,ObjClsTRG = @ObjClsTrg
           ,ObjClsUFN = @ObjClsUfn
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    SET @PopTbl = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0 
    END
    SET @PopVew = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopUsp = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        WHEN @ObjClsUSP THEN 1
        ELSE                 0 
    END
    SET @PopTrg = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopUfn = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopPrm = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        WHEN @ObjClsVEW THEN 1
        WHEN @ObjClsUSP THEN 1
        WHEN @ObjClsTRG THEN 1
        WHEN @ObjClsUFN THEN 1
        ELSE                 0
    END
    SET @PopDep = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopFld = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopClm = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        WHEN @ObjClsVEW THEN 1
        WHEN @ObjClsUSP THEN 1
        WHEN @ObjClsTRG THEN 1
        WHEN @ObjClsUFN THEN 1
        ELSE                 0
    END
    SET @PopPar = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopDcv = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopPky = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopUky = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopInd = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopStt = CASE @DfnCls
        WHEN @ObjClsTBL THEN 0
        ELSE                 0
    END
    SET @PopFky = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopRky = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopDef = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    SET @PopChk = CASE @DfnCls
        WHEN @ObjClsTBL THEN 1
        ELSE                 0
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT
           'POP'     = 'POP'
           ,DfnCls   = @DfnCls
           ,OBJ      = @PopObj
           ,TBL      = @PopTbl
           ,VEW      = @PopVew
           ,USP      = @PopUsp
           ,TRG      = @PopTrg
           ,UFN      = @PopUfn
           ,PRM      = @PopPrm
           ,DEP      = @PopDep
           ,FLD      = @PopFld
           ,CLM      = @PopClm
           ,PAR      = @PopPar
           ,DVV      = @PopDcv
           ,PKY      = @PopPky
           ,UKY      = @PopUky
           ,IND      = @PopInd
           ,STT      = @PopStt
           ,FKY      = @PopFky
           ,RKY      = @PopRky
           ,DEF      = @PopDef
           ,CHK      = @PopChk
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    SET @AdfTbl = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON 
    END
    SET @AdfObj = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfTbl = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfVew = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfUsp = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        WHEN @ObjClsUSP THEN @AsnDefNON
        ELSE                 @AsnDefNON 
    END
    SET @AdfUln = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfTrg = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfUfn = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfPrm = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefMTY
        WHEN @ObjClsVEW THEN @AsnDefMTY
        WHEN @ObjClsUSP THEN @AsnDefMTY
        WHEN @ObjClsTRG THEN @AsnDefMTY
        WHEN @ObjClsUFN THEN @AsnDefMTY
        ELSE                 @AsnDefNON
    END
    SET @AdfDep = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfFld = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfClm = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefMTY
        WHEN @ObjClsVEW THEN @AsnDefMTY
        WHEN @ObjClsUSP THEN @AsnDefMTY
        WHEN @ObjClsTRG THEN @AsnDefMTY
        WHEN @ObjClsUFN THEN @AsnDefMTY
        ELSE                 @AsnDefNON
    END
    SET @AdfPar = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfDcv = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfPky = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfUky = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfInd = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfStt = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfFky = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfRky = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfDef = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    SET @AdfChk = CASE @DfnCls
        WHEN @ObjClsTBL THEN @AsnDefNON
        ELSE                 @AsnDefNON
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  User Tables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #TblDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #TblDfn (               -- DROP TABLE dbo.zzz_TblDfn; CREATE TABLE dbo.zzz_TblDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        TblNam                                            sysname              NOT NULL,
        TblSiz                                            dec(15,2)            NOT NULL,
        RecQty                                            int                  NOT NULL,
        RecSiz                                            dec(15,0)            NOT NULL,
        HasPky                                            bit                  NOT NULL,
        UkyQty                                            smallint             NOT NULL,
        FkyQty                                            smallint             NOT NULL,
        RkyQty                                            smallint             NOT NULL,
        IndQty                                            smallint             NOT NULL,
        DefQty                                            smallint             NOT NULL,
        ChkQty                                            smallint             NOT NULL,
        ClxNam                                            sysname              NOT NULL,
        HasHtk                                            bit                  NOT NULL,
        HasDim                                            bit                  NOT NULL,
        HasFct                                            bit                  NOT NULL,
        HasDsb                                            bit                  NOT NULL,
        HasDlt                                            bit                  NOT NULL,
        HasLok                                            bit                  NOT NULL,
        HasCrt                                            bit                  NOT NULL,
        HasUpd                                            bit                  NOT NULL,
        HasExp                                            bit                  NOT NULL,
        HasDel                                            bit                  NOT NULL,
        HasHst                                            bit                  NOT NULL,
        HasAud                                            bit                  NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_TblDfn
    ------------------------------------------------------------------------------------------------
    IF @PopTbl = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfTbl; EXEC ut_zzRSJ TBL,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@TblNam,@TblSiz,@RecQty,@RecSiz,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #TblDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, TblNam, TblSiz, RecQty, RecSiz, HasPky, UkyQty, FkyQty, RkyQty, IndQty, DefQty, ChkQty, ClxNam, HasHtk, HasDim, HasFct, HasDsb, HasDlt, HasLok, HasCrt, HasUpd, HasExp, HasDel, HasHst, HasAud, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@TblNam,@TblSiz,@RecQty,@RecSiz,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_TblDefs CURSOR LOCAL FOR SELECT *                              FROM #TblDfn ORDER BY TblNam  -- Cursor
    DECLARE @TblCnt smallint; SET @TblCnt = (SELECT COUNT(*)                   FROM #TblDfn)                 -- Record Count
    DECLARE @TblFln smallint; SET @TblFln = (SELECT ISNULL(MAX(LEN(TblNam)),0) FROM #TblDfn)                 -- Max FieldName Length
    DECLARE @TblTln smallint; SET @TblTln = (SELECT ISNULL(MAX(LEN(TblNam)),0) FROM #TblDfn)                 -- Max TableName Length
    DECLARE @TblVln smallint; SET @TblVln = CASE WHEN @TblFln < @MinLenVAR THEN @MinLenVAR ELSE @TblFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_TblDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#TblDfn' = '#TblDfn', TblCnt = @TblCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,TblNam = LEFT(dfn.TblNam,40)
            ,TblSiz = dfn.TblSiz
            ,RecQty = dfn.RecQty
            ,RecSiz = dfn.RecSiz
            ,HasPky = dfn.HasPky
            ,UkyQty = dfn.UkyQty
            ,FkyQty = dfn.FkyQty
            ,RkyQty = dfn.RkyQty
            ,IndQty = dfn.IndQty
            ,DefQty = dfn.DefQty
            ,ChkQty = dfn.ChkQty
            ,ClxNam = LEFT(dfn.ClxNam,40)
            ,HasHtk = dfn.HasHtk
            ,HasDim = dfn.HasDim
            ,HasFct = dfn.HasFct
            ,HasDsb = dfn.HasDsb
            ,HasDlt = dfn.HasDlt
            ,HasLok = dfn.HasLok
            ,HasCrt = dfn.HasCrt
            ,HasUpd = dfn.HasUpd
            ,HasExp = dfn.HasExp
            ,HasDel = dfn.HasDel
            ,HasHst = dfn.HasHst
            ,HasAud = dfn.HasAud
            ,DfnStd = dfn.DfnStd
        FROM
            #TblDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Views
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #VewDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #VewDfn (               -- DROP TABLE dbo.zzz_VewDfn; CREATE TABLE dbo.zzz_VewDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        VewNam                                            sysname              NOT NULL,
        RecQty                                            int                  NOT NULL,
        HasPky                                            bit                  NOT NULL,
        UkyQty                                            smallint             NOT NULL,
        FkyQty                                            smallint             NOT NULL,
        RkyQty                                            smallint             NOT NULL,
        IndQty                                            smallint             NOT NULL,
        DefQty                                            smallint             NOT NULL,
        ChkQty                                            smallint             NOT NULL,
        ClxNam                                            sysname              NOT NULL,
        HasHtk                                            bit                  NOT NULL,
        HasDim                                            bit                  NOT NULL,
        HasFct                                            bit                  NOT NULL,
        HasDsb                                            bit                  NOT NULL,
        HasDlt                                            bit                  NOT NULL,
        HasLok                                            bit                  NOT NULL,
        HasCrt                                            bit                  NOT NULL,
        HasUpd                                            bit                  NOT NULL,
        HasExp                                            bit                  NOT NULL,
        HasDel                                            bit                  NOT NULL,
        HasHst                                            bit                  NOT NULL,
        HasAud                                            bit                  NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_VewDfn
    ------------------------------------------------------------------------------------------------
    IF @PopVew = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfVew; EXEC ut_zzRSJ VEW,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@VewNam,@RecQty,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #VewDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, VewNam, RecQty, HasPky, UkyQty, FkyQty, RkyQty, IndQty, DefQty, ChkQty, ClxNam, HasHtk, HasDim, HasFct, HasDsb, HasDlt, HasLok, HasCrt, HasUpd, HasExp, HasDel, HasHst, HasAud, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@VewNam,@RecQty,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_VewDefs CURSOR LOCAL FOR SELECT *                              FROM #VewDfn ORDER BY VewNam  -- Cursor
    DECLARE @VewCnt smallint; SET @VewCnt = (SELECT COUNT(*)                   FROM #VewDfn)                 -- Record Count
    DECLARE @VewFln smallint; SET @VewFln = (SELECT ISNULL(MAX(LEN(VewNam)),0) FROM #VewDfn)                 -- Max FieldName Length
    DECLARE @VewTln smallint; SET @VewTln = (SELECT ISNULL(MAX(LEN(VewNam)),0) FROM #VewDfn)                 -- Max TableName Length
    DECLARE @VewVln smallint; SET @VewVln = CASE WHEN @VewFln < @MinLenVAR THEN @MinLenVAR ELSE @VewFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_VewDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#VewDfn' = '#VewDfn', VewCnt = @VewCnt
            ,DfnID    = dfn.DfnID
            ,DfnCls   = dfn.DfnCls
            ,DfnExs   = dfn.DfnExs
            ,DfnFmt   = dfn.DfnFmt
            ,DfnFmx   = dfn.DfnFmx
            ,DfnFix   = dfn.DfnFix
            ,VewNam   = LEFT(dfn.VewNam,30)
            ,RecQty   = dfn.RecQty
            ,HasPky   = dfn.HasPky
            ,UkyQty   = dfn.UkyQty
            ,FkyQty   = dfn.FkyQty
            ,RkyQty   = dfn.RkyQty
            ,IndQty   = dfn.IndQty
            ,DefQty   = dfn.DefQty
            ,ChkQty   = dfn.ChkQty
            ,ClxNam   = LEFT(dfn.ClxNam,30)
            ,HasHtk   = dfn.HasHtk
            ,HasDim   = dfn.HasDim
            ,HasFct   = dfn.HasFct
            ,HasDsb   = dfn.HasDsb
            ,HasDlt   = dfn.HasDlt
            ,HasLok   = dfn.HasLok
            ,HasCrt   = dfn.HasCrt
            ,HasUpd   = dfn.HasUpd
            ,HasExp   = dfn.HasExp
            ,HasDel   = dfn.HasDel
            ,HasHst   = dfn.HasHst
            ,HasAud   = dfn.HasAud
            ,DfnStd   = dfn.DfnStd
        FROM
            #VewDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  User Stored Procedures
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #UspDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #UspDfn (               -- DROP TABLE dbo.zzz_UspDfn; CREATE TABLE dbo.zzz_UspDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        UspNam                                            sysname              NOT NULL,
        DepUsp                                            smallint             NOT NULL,
        DepFnc                                            smallint             NOT NULL,
        DepTrg                                            smallint             NOT NULL,
        UseTbl                                            smallint             NOT NULL,
        UseVew                                            smallint             NOT NULL,
        UseUsp                                            smallint             NOT NULL,
        UseFnc                                            smallint             NOT NULL,
        UseTrg                                            smallint             NOT NULL,
        HasHtk                                            bit                  NOT NULL,
        HasDim                                            bit                  NOT NULL,
        HasFct                                            bit                  NOT NULL,
        HasDsb                                            bit                  NOT NULL,
        HasDlt                                            bit                  NOT NULL,
        HasLok                                            bit                  NOT NULL,
        HasCrt                                            bit                  NOT NULL,
        HasUpd                                            bit                  NOT NULL,
        HasExp                                            bit                  NOT NULL,
        HasDel                                            bit                  NOT NULL,
        HasHst                                            bit                  NOT NULL,
        HasAud                                            bit                  NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_UspDfn
    ------------------------------------------------------------------------------------------------
    IF @PopUsp = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfUsp; EXEC ut_zzRSJ USP,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@UspNam,@DepUsp,@DepFnc,@DepTrg,@UseTbl,@UseVew,@UseUsp,@UseFnc,@UseTrg,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #UspDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, UspNam, DepUsp, DepFnc, DepTrg, UseTbl, UseVew, UseUsp, UseFnc, UseTrg, HasHtk, HasDim, HasFct, HasDsb, HasDlt, HasLok, HasCrt, HasUpd, HasExp, HasDel, HasHst, HasAud, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@UspNam,@DepUsp,@DepFnc,@DepTrg,@UseTbl,@UseVew,@UseUsp,@UseFnc,@UseTrg,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_UspDefs CURSOR LOCAL FOR SELECT *                              FROM #UspDfn ORDER BY UspNam  -- Cursor
    DECLARE @UspCnt smallint; SET @UspCnt = (SELECT COUNT(*)                   FROM #UspDfn)                 -- Record Count
    DECLARE @UspFln smallint; SET @UspFln = (SELECT ISNULL(MAX(LEN(UspNam)),0) FROM #UspDfn)                 -- Max FieldName Length
    DECLARE @UspTln smallint; SET @UspTln = (SELECT ISNULL(MAX(LEN(UspNam)),0) FROM #UspDfn)                 -- Max TableName Length
    DECLARE @UspVln smallint; SET @UspVln = CASE WHEN @UspFln < @MinLenVAR THEN @MinLenVAR ELSE @UspFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_UspDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#UspDfn' = '#UspDfn', UspCnt = @UspCnt
            ,DfnID    = dfn.DfnID
            ,DfnCls   = dfn.DfnCls
            ,DfnExs   = dfn.DfnExs
            ,DfnFmt   = dfn.DfnFmt
            ,DfnFmx   = dfn.DfnFmx
            ,DfnFix   = dfn.DfnFix
            ,UspNam   = LEFT(dfn.UspNam,30)
            ,DepUsp   = dfn.DepUsp
            ,DepFnc   = dfn.DepFnc
            ,DepTrg   = dfn.DepTrg
            ,UseTbl   = dfn.UseTbl
            ,UseVew   = dfn.UseVew
            ,UseUsp   = dfn.UseUsp
            ,UseFnc   = dfn.UseFnc
            ,UseTrg   = dfn.UseTrg
            ,HasHtk   = dfn.HasHtk
            ,HasDim   = dfn.HasDim
            ,HasFct   = dfn.HasFct
            ,HasDsb   = dfn.HasDsb
            ,HasDlt   = dfn.HasDlt
            ,HasLok   = dfn.HasLok
            ,HasCrt   = dfn.HasCrt
            ,HasUpd   = dfn.HasUpd
            ,HasExp   = dfn.HasExp
            ,HasDel   = dfn.HasDel
            ,HasHst   = dfn.HasHst
            ,HasAud   = dfn.HasAud
            ,DfnStd   = dfn.DfnStd
        FROM
            #UspDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Triggers
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #TrgDfn temp table
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  User Functions
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #UfnDfn temp table
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 

    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Object Permissions
    ------------------------------------------------------------------------------------------------
    -- NOTE:  This uses a potentially different object list (@PrmLst) to allow for SProcs/Tables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #PrmDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #PrmDfn (               -- DROP TABLE dbo.zzz_PrmDfn; CREATE TABLE dbo.zzz_PrmDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ObjNAM                                            sysname              NOT NULL,
        PrmLvl                                            tinyint              NOT NULL,
        ProTxt                                            varchar(20)          NOT NULL,
        UsrNam                                            sysname              NOT NULL,
        ActLst                                            varchar(300)         NOT NULL,
        ProDir                                            varchar(6)           NOT NULL,
        UsrDsc                                            varchar(80)          NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_PrmDfn
    ------------------------------------------------------------------------------------------------
    IF @PopPrm = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfPrm; EXEC ut_zzRSJ PRM,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ObjNam,@PrmLvl,@ProTxt,@UsrNam,@ActLst,@ProDir,@UsrDsc
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #PrmDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ObjNam, PrmLvl, ProTxt, UsrNam, ActLst, ProDir, UsrDsc
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ObjNam,@PrmLvl,@ProTxt,@UsrNam,@ActLst,@ProDir,@UsrDsc
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_PrmDefs CURSOR LOCAL FOR SELECT *                              FROM #PrmDfn ORDER BY ObjNam  -- Cursor
    DECLARE @PrmCnt smallint; SET @PrmCnt = (SELECT COUNT(*)                   FROM #PrmDfn)                 -- Record Count
    DECLARE @PrmFln smallint; SET @PrmFln = (SELECT ISNULL(MAX(LEN(UsrNam)),0) FROM #PrmDfn)                 -- Max FieldName Length
    DECLARE @PrmTln smallint; SET @PrmTln = (SELECT ISNULL(MAX(LEN(ObjNam)),0) FROM #PrmDfn)                 -- Max TableName Length
    DECLARE @PrmVln smallint; SET @PrmVln = CASE WHEN @PrmFln < @MinLenVAR THEN @MinLenVAR ELSE @PrmFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    DECLARE @PrmPln smallint; SET @PrmPln = (SELECT ISNULL(MAX(LEN(ProTxt)),0) FROM #PrmDfn)                 -- Max ProtectText Length
    DECLARE @PrmAln smallint; SET @PrmAln = (SELECT ISNULL(MAX(LEN(ActLst)),0) FROM #PrmDfn)                 -- Max ActionList Length
    SET @PrmPln = CASE WHEN @PrmPln < 6 THEN 6 ELSE @PrmPln END                                              -- Min ProtectText Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_PrmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PrmDfn' = '#PrmDfn', PrmCnt = @PrmCnt
            ,DfnID    = dfn.DfnID
            ,DfnCls   = dfn.DfnCls
            ,DfnExs   = dfn.DfnExs
            ,DfnFmt   = dfn.DfnFmt
            ,DfnFmx   = dfn.DfnFmx
            ,DfnFix   = dfn.DfnFix
            ,ObjNam   = LEFT(dfn.ObjNam,30)
            ,PrmLvl   = dfn.PrmLvl
            ,ProTxt   = dfn.ProTxt
            ,UsrNam   = LEFT(dfn.UsrNam,30)
            ,ActLst   = LEFT(dfn.ActLst,30)
            ,ProDir   = dfn.ProDir
            ,UsrDsc   = LEFT(dfn.UsrDsc,30)
        FROM
            #PrmDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Object Dependencies
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #JdpDfn temp table
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Column Definitions
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #ClmDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #ClmDfn (               -- DROP TABLE dbo.zzz_ClmDfn; CREATE TABLE dbo.zzz_ClmDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        --------------------------------------------------------------------------------------------
        DfnCls                                            char(3)              NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        --------------------------------------------------------------------------------------------
        ClmLvl                                            tinyint                  NULL,
        ClmObj                                            sysname                  NULL,
        ClmOrd                                            smallint                 NULL,
        ClmNam                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmUtp                                            int                      NULL,
        ClmUtx                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmStp                                            tinyint                  NULL,
        ClmStx                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmDtp                                            int                      NULL,
        ClmDTX                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmLen                                            smallint                 NULL,
        ClmWid                                            smallint                 NULL,
        ClmPrc                                            tinyint                  NULL,
        ClmScl                                            tinyint                  NULL,
        ClmDsp                                            smallint                 NULL,
        --------------------------------------------------------------------------------------------
        ClmDct                                            varchar(3)               NULL,
        ClmQot                                            bit                      NULL,
        ClmNul                                            bit                      NULL,
        ClmIdn                                            bit                      NULL,
        ClmCmp                                            bit                      NULL,
        ClmOup                                            bit                      NULL,
        ClmVwd                                            bit                      NULL,
        ClmMax                                            bit                      NULL,
        ClmAud                                            bit                      NULL,
        --------------------------------------------------------------------------------------------
        ClmPky                                            tinyint                  NULL,
        ClmUky                                            tinyint                  NULL,
        ClmXky                                            tinyint                  NULL,
        ClmFky                                            bit                      NULL,
        --------------------------------------------------------------------------------------------
        ClmDef                                            bit                      NULL,
        ClmDfv                                            nvarchar(max)            NULL,
        ClmEmv                                            varchar(100)             NULL,
        ClmCpx                                            nvarchar(max)            NULL,
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_ClmDfn
    ------------------------------------------------------------------------------------------------
    IF @PopClm = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfClm; EXEC ut_zzRSJ CLM,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnFix,@ClmLvl,@ClmObj,@ClmOrd,@ClmNam,@ClmUtp,@ClmUtx,@ClmStp,@ClmStx,@ClmDtp,@ClmDTX,@ClmLen,@ClmWid,@ClmPrc,@ClmScl,@ClmDsp,@ClmDct,@ClmQot,@ClmNul,@ClmIdn,@ClmCmp,@ClmOup,@ClmVwd,@ClmMax,@ClmAud,@ClmPky,@ClmUky,@ClmXky,@ClmFky,@ClmDef,@ClmDfv,@ClmEmv,@ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #ClmDfn (  -- EXEC ut_zzSQL INX,zzz_ClmDfn
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnFix, ClmLvl, ClmObj, ClmOrd, ClmNam, ClmUtp, ClmUtx, ClmStp, ClmStx, ClmDtp, ClmDTX, ClmLen, ClmWid, ClmPrc, ClmScl, ClmDsp, ClmDct, ClmQot, ClmNul, ClmIdn, ClmCmp, ClmOup, ClmVwd, ClmMax, ClmAud, ClmPky, ClmUky, ClmXky, ClmFky, ClmDef, ClmDfv, ClmEmv, ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnFix,@ClmLvl,@ClmObj,@ClmOrd,@ClmNam,@ClmUtp,@ClmUtx,@ClmStp,@ClmStx,@ClmDtp,@ClmDTX,@ClmLen,@ClmWid,@ClmPrc,@ClmScl,@ClmDsp,@ClmDct,@ClmQot,@ClmNul,@ClmIdn,@ClmCmp,@ClmOup,@ClmVwd,@ClmMax,@ClmAud,@ClmPky,@ClmUky,@ClmXky,@ClmFky,@ClmDef,@ClmDfv,@ClmEmv,@ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_ClmDefs CURSOR LOCAL FAST_FORWARD FOR SELECT *                 FROM #ClmDfn ORDER BY ClmOrd  -- Cursor
    DECLARE @ClmCnt smallint; SET @ClmCnt = (SELECT COUNT(*)                   FROM #ClmDfn)                 -- Record Count
    DECLARE @ClmFln smallint; SET @ClmFln = (SELECT ISNULL(MAX(LEN(ClmNam)),0) FROM #ClmDfn)                 -- Max FieldName Length
    DECLARE @ClmTln smallint; SET @ClmTln = (SELECT ISNULL(MAX(LEN(ClmObj)),0) FROM #ClmDfn)                 -- Max TableName Length
    DECLARE @ClmVln smallint; SET @ClmVln = CASE WHEN @ClmFln < @MinLenVAR THEN @MinLenVAR ELSE @ClmFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#ClmDfn' = '#ClmDfn', ClmCnt = @ClmCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnFix = dfn.DfnFix
            ,ClmLvl = dfn.ClmLvl
            ,ClmObj = CONVERT(varchar(50),dfn.ClmObj)
            ,ClmOrd = dfn.ClmOrd
            ,ClmNam = CONVERT(varchar(40),dfn.ClmNam)
            ,ClmUtp = dfn.ClmUtp
            ,ClmUtx = CONVERT(varchar(20),dfn.ClmUtx)
            ,ClmStp = dfn.ClmStp
            ,ClmStx = CONVERT(varchar(20),dfn.ClmStx)
            ,ClmDtp = dfn.ClmDtp
            ,ClmDTX = CONVERT(varchar(20),dfn.ClmDTX)
            ,ClmLen = dfn.ClmLen
            ,ClmWid = dfn.ClmWid
            ,ClmPrc = dfn.ClmPrc
            ,ClmScl = dfn.ClmScl
            ,ClmDsp = dfn.ClmDsp
            ,ClmDct = dfn.ClmDct
            ,ClmQot = dfn.ClmQot
            ,ClmNul = dfn.ClmNul
            ,ClmIdn = dfn.ClmIdn
            ,ClmCmp = dfn.ClmCmp
            ,ClmOup = dfn.ClmOup
            ,ClmVwd = dfn.ClmVwd
            ,ClmMax = dfn.ClmMax
            ,ClmAud = dfn.ClmAud
            ,ClmPky = dfn.ClmPky
            ,ClmUky = dfn.ClmUky
            ,ClmXky = dfn.ClmXky
            ,ClmFky = dfn.ClmFky
            ,ClmDef = dfn.ClmDef
            ,ClmDfv = CONVERT(varchar(20),dfn.ClmDfv)
            ,ClmEmv = CONVERT(varchar(20),dfn.ClmEmv)
            ,ClmCpx = dfn.ClmCpx
        FROM
            #ClmDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Default Columns
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #DcvDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #DcvDfn (               -- DROP TABLE dbo.zzz_DcvDfn; CREATE TABLE dbo.zzz_DcvDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        --------------------------------------------------------------------------------------------
        DfnCls                                            char(3)              NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        --------------------------------------------------------------------------------------------
        ClmLvl                                            tinyint                  NULL,
        ClmObj                                            sysname                  NULL,
        ClmOrd                                            smallint                 NULL,
        ClmNam                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmUtp                                            int                      NULL,
        ClmUtx                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmStp                                            tinyint                  NULL,
        ClmStx                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmDtp                                            int                      NULL,
        ClmDTX                                            sysname                  NULL,
        --------------------------------------------------------------------------------------------
        ClmLen                                            smallint                 NULL,
        ClmWid                                            smallint                 NULL,
        ClmPrc                                            tinyint                  NULL,
        ClmScl                                            tinyint                  NULL,
        ClmDsp                                            smallint                 NULL,
        --------------------------------------------------------------------------------------------
        ClmDct                                            varchar(3)               NULL,
        ClmQot                                            bit                      NULL,
        ClmNul                                            bit                      NULL,
        ClmIdn                                            bit                      NULL,
        ClmCmp                                            bit                      NULL,
        ClmOup                                            bit                      NULL,
        ClmVwd                                            bit                      NULL,
        ClmMax                                            bit                      NULL,
        ClmAud                                            bit                      NULL,
        --------------------------------------------------------------------------------------------
        ClmPky                                            tinyint                  NULL,
        ClmUky                                            tinyint                  NULL,
        ClmXky                                            tinyint                  NULL,
        ClmFky                                            bit                      NULL,
        --------------------------------------------------------------------------------------------
        ClmDef                                            bit                      NULL,
        ClmDfv                                            nvarchar(max)            NULL,
        ClmEmv                                            varchar(100)             NULL,
        ClmCpx                                            nvarchar(max)            NULL,
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_DcvDfn
    ------------------------------------------------------------------------------------------------
    IF @PopDcv = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfDcv; EXEC ut_zzRSJ DCV,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnFix,@ClmLvl,@ClmObj,@ClmOrd,@ClmNam,@ClmUtp,@ClmUtx,@ClmStp,@ClmStx,@ClmDtp,@ClmDTX,@ClmLen,@ClmWid,@ClmPrc,@ClmScl,@ClmDsp,@ClmDct,@ClmQot,@ClmNul,@ClmIdn,@ClmCmp,@ClmOup,@ClmVwd,@ClmMax,@ClmAud,@ClmPky,@ClmUky,@ClmXky,@ClmFky,@ClmDef,@ClmDfv,@ClmEmv,@ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #DcvDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnFix, ClmLvl, ClmObj, ClmOrd, ClmNam, ClmUtp, ClmUtx, ClmStp, ClmStx, ClmDtp, ClmDTX, ClmLen, ClmWid, ClmPrc, ClmScl, ClmDsp, ClmDct, ClmQot, ClmNul, ClmIdn, ClmCmp, ClmOup, ClmVwd, ClmMax, ClmAud, ClmPky, ClmUky, ClmXky, ClmFky, ClmDef, ClmDfv, ClmEmv, ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnFix,@ClmLvl,@ClmObj,@ClmOrd,@ClmNam,@ClmUtp,@ClmUtx,@ClmStp,@ClmStx,@ClmDtp,@ClmDTX,@ClmLen,@ClmWid,@ClmPrc,@ClmScl,@ClmDsp,@ClmDct,@ClmQot,@ClmNul,@ClmIdn,@ClmCmp,@ClmOup,@ClmVwd,@ClmMax,@ClmAud,@ClmPky,@ClmUky,@ClmXky,@ClmFky,@ClmDef,@ClmDfv,@ClmEmv,@ClmCpx
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_DcvDefs CURSOR LOCAL FAST_FORWARD FOR SELECT *                 FROM #ClmDfn ORDER BY ClmOrd  -- Cursor
    DECLARE @DcvCnt smallint; SET @DcvCnt = (SELECT COUNT(*)                   FROM #DcvDfn)                 -- Record Count
    DECLARE @DcvFln smallint; SET @DcvFln = (SELECT ISNULL(MAX(LEN(ClmNam)),0) FROM #DcvDfn)                 -- Max FieldName Length
    DECLARE @DcvTln smallint; SET @DcvTln = (SELECT ISNULL(MAX(LEN(ClmObj)),0) FROM #DcvDfn)                 -- Max TableName Length
    DECLARE @DcvVln smallint; SET @DcvVln = CASE WHEN @DcvFln < @MinLenVAR THEN @MinLenVAR ELSE @DcvFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_DcvDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#DcvDfn' = '#DcvDfn', DcvCnt = @DcvCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnFix = dfn.DfnFix
            ,ClmLvl = dfn.ClmLvl
            ,ClmObj = CONVERT(varchar(50),dfn.ClmObj)
            ,ClmOrd = dfn.ClmOrd
            ,ClmNam = CONVERT(varchar(40),dfn.ClmNam)
            ,ClmUtp = dfn.ClmUtp
            ,ClmUtx = CONVERT(varchar(20),dfn.ClmUtx)
            ,ClmStp = dfn.ClmStp
            ,ClmStx = CONVERT(varchar(20),dfn.ClmStx)
            ,ClmDtp = dfn.ClmDtp
            ,ClmDTX = CONVERT(varchar(20),dfn.ClmDTX)
            ,ClmLen = dfn.ClmLen
            ,ClmWid = dfn.ClmWid
            ,ClmPrc = dfn.ClmPrc
            ,ClmScl = dfn.ClmScl
            ,ClmDsp = dfn.ClmDsp
            ,ClmDct = dfn.ClmDct
            ,ClmQot = dfn.ClmQot
            ,ClmNul = dfn.ClmNul
            ,ClmIdn = dfn.ClmIdn
            ,ClmCmp = dfn.ClmCmp
            ,ClmOup = dfn.ClmOup
            ,ClmVwd = dfn.ClmVwd
            ,ClmMax = dfn.ClmMax
            ,ClmAud = dfn.ClmAud
            ,ClmPky = dfn.ClmPky
            ,ClmUky = dfn.ClmUky
            ,ClmXky = dfn.ClmXky
            ,ClmFky = dfn.ClmFky
            ,ClmDef = dfn.ClmDef
            ,ClmDfv = CONVERT(varchar(20),dfn.ClmDfv)
            ,ClmEmv = CONVERT(varchar(20),dfn.ClmEmv)
            ,ClmCpx = dfn.ClmCpx
        FROM
            #DcvDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Parameters
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #ParDfn temp table
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Primary Keys
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #PkyDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #PkyDfn (               -- DROP TABLE dbo.zzz_PkyDfn; CREATE TABLE dbo.zzz_PkyDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FilFct                                            smallint             NOT NULL,
        IsClus                                            bit                  NOT NULL,
        IsUniq                                            bit                  NOT NULL,
        ConKys                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_PkyDfn
    ------------------------------------------------------------------------------------------------
    IF @PopPky = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfPky; EXEC ut_zzRSJ PKY,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #PkyDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FilFct, IsClus, IsUniq, ConKys, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_PkyDefs CURSOR LOCAL FOR SELECT *                              FROM #PkyDfn ORDER BY ConTbl  -- Cursor
    DECLARE @PkyCnt smallint; SET @PkyCnt = (SELECT COUNT(*)                   FROM #PkyDfn)                 -- Record Count
    DECLARE @PkyFln smallint; SET @PkyFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #PkyDfn)                 -- Max FieldName Length
    DECLARE @PkyTln smallint; SET @PkyTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #PkyDfn)                 -- Max TableName Length
    DECLARE @PkyVln smallint; SET @PkyVln = CASE WHEN @PkyFln < @MinLenVAR THEN @MinLenVAR ELSE @PkyFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_PkyDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PkyDfn' = '#PkyDfn', PkyCnt = @PkyCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FilFct = dfn.FilFct
            ,IsClus = dfn.IsClus
            ,IsUniq = dfn.IsUniq
            ,ConKys = LEFT(dfn.ConKys,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #PkyDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  PrimaryKey Key List
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #PklDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #PklDfn (               -- DROP TABLE dbo.zzz_PklDfn; CREATE TABLE dbo.zzz_PklDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FilFct                                            smallint             NOT NULL,
        IsClus                                            bit                  NOT NULL,
        IsUniq                                            bit                  NOT NULL,
        KeyOrd                                            smallint             NOT NULL,
        KeyClm                                            sysname              NOT NULL,
        IndDir                                            varchar(5)           NOT NULL,
        ConID                                             int                  NOT NULL,
        TblID                                             int                  NOT NULL,
        ClmID                                             int                  NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_PklDfn
    ------------------------------------------------------------------------------------------------
    IF @PopPky = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfPky; EXEC ut_zzRSJ PKL,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@KeyOrd,@KeyClm,@IndDir,@ConID,@TblID,@ClmID,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #PklDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FilFct, IsClus, IsUniq, KeyOrd, KeyClm, IndDir, ConID, TblID, ClmID, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@KeyOrd,@KeyClm,@IndDir,@ConID,@TblID,@ClmID,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_PklDefs CURSOR LOCAL FOR SELECT *                              FROM #PklDfn ORDER BY ConTbl  -- Cursor
    DECLARE @PklCnt smallint; SET @PklCnt = (SELECT COUNT(*)                   FROM #PklDfn)                 -- Record Count
    DECLARE @PklFln smallint; SET @PklFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #PklDfn)                 -- Max FieldName Length
    DECLARE @PklTln smallint; SET @PklTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #PklDfn)                 -- Max TableName Length
    DECLARE @PklVln smallint; SET @PklVln = CASE WHEN @PklFln < @MinLenVAR THEN @MinLenVAR ELSE @PklFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_PklDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR @DbgRsj = 1 OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PklDfn' = '#PklDfn', PklCnt = @PklCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,40)
            ,ConNam = LEFT(dfn.ConNam,60)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FilFct = dfn.FilFct
            ,IsClus = dfn.IsClus
            ,IsUniq = dfn.IsUniq
            ,KeyOrd = dfn.KeyOrd
            ,KeyClm = LEFT(dfn.KeyClm,30)
            ,IndDir = dfn.IndDir
            ,ConID  = dfn.ConID
            ,TblID  = dfn.TblID
            ,ClmID  = dfn.ClmID
            ,DfnStd = dfn.DfnStd
        FROM
            #PklDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Unique Keys
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #UkyDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #UkyDfn (               -- DROP TABLE dbo.zzz_UkyDfn; CREATE TABLE dbo.zzz_UkyDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FilFct                                            smallint             NOT NULL,
        IsClus                                            bit                  NOT NULL,
        IsUniq                                            bit                  NOT NULL,
        ConKys                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_UkyDfn
    ------------------------------------------------------------------------------------------------
    IF @PopUky = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfUky; EXEC ut_zzRSJ UKY,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #UkyDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FilFct, IsClus, IsUniq, ConKys, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_UkyDefs CURSOR LOCAL FOR SELECT *                              FROM #UkyDfn ORDER BY ConTbl  -- Cursor
    DECLARE @UkyCnt smallint; SET @UkyCnt = (SELECT COUNT(*)                   FROM #UkyDfn)                 -- Record Count
    DECLARE @UkyFln smallint; SET @UkyFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #UkyDfn)                 -- Max FieldName Length
    DECLARE @UkyTln smallint; SET @UkyTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #UkyDfn)                 -- Max TableName Length
    DECLARE @UkyVln smallint; SET @UkyVln = CASE WHEN @UkyFln < @MinLenVAR THEN @MinLenVAR ELSE @UkyFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_UkyDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR @DbgRsj = 1 OR 0=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#UkyDfn' = '#UkyDfn', UkyCnt = @UkyCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FilFct = dfn.FilFct
            ,IsClus = dfn.IsClus
            ,IsUniq = dfn.IsUniq
            ,ConKys = LEFT(dfn.ConKys,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #UkyDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Standard Indexes
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #IndDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #IndDfn (               -- DROP TABLE dbo.zzz_IndDfn; CREATE TABLE dbo.zzz_IndDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FilFct                                            smallint             NOT NULL,
        IsClus                                            bit                  NOT NULL,
        IsUniq                                            bit                  NOT NULL,
        ConKys                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_IndDfn
    ------------------------------------------------------------------------------------------------
    IF @PopInd = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfInd; EXEC ut_zzRSJ IND,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #IndDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FilFct, IsClus, IsUniq, ConKys, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FilFct,@IsClus,@IsUniq,@ConKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_IndDefs CURSOR LOCAL FOR SELECT *                              FROM #IndDfn ORDER BY ConTbl  -- Cursor
    DECLARE @IndCnt smallint; SET @IndCnt = (SELECT COUNT(*)                   FROM #IndDfn)                 -- Record Count
    DECLARE @IndFln smallint; SET @IndFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #IndDfn)                 -- Max FieldName Length
    DECLARE @IndTln smallint; SET @IndTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #IndDfn)                 -- Max TableName Length
    DECLARE @IndVln smallint; SET @IndVln = CASE WHEN @IndFln < @MinLenVAR THEN @MinLenVAR ELSE @IndFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#IndDfn' = '#IndDfn', IndCnt = @IndCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FilFct = dfn.FilFct
            ,IsClus = dfn.IsClus
            ,IsUniq = dfn.IsUniq
            ,ConKys = LEFT(dfn.ConKys,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #IndDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Dynamic Statistics
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #SttDfn temp table
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Foreign Keys
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #FkyDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #FkyDfn (               -- DROP TABLE dbo.zzz_FkyDfn; CREATE TABLE dbo.zzz_FkyDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FkyTbl                                            sysname              NOT NULL,
        FkyKys                                            varchar(1000)        NOT NULL,
        RkyTbl                                            sysname              NOT NULL,
        RkyKys                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_FkyDfn
    ------------------------------------------------------------------------------------------------
    IF @PopFky = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfFky; EXEC ut_zzRSJ FKY,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #FkyDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FkyTbl, FkyKys, RkyTbl, RkyKys, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_FkyDefs CURSOR LOCAL FOR SELECT *                              FROM #FkyDfn ORDER BY FkyTbl  -- Cursor
    DECLARE @FkyCnt smallint; SET @FkyCnt = (SELECT COUNT(*)                   FROM #FkyDfn)                 -- Record Count
    DECLARE @FkyFln smallint; SET @FkyFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #FkyDfn)                 -- Max FieldName Length
    DECLARE @FkyTln smallint; SET @FkyTln = (SELECT ISNULL(MAX(LEN(FkyTbl)),0) FROM #FkyDfn)                 -- Max TableName Length
    DECLARE @FkyVln smallint; SET @FkyVln = CASE WHEN @FkyFln < @MinLenVAR THEN @MinLenVAR ELSE @FkyFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FkyDfn' = '#FkyDfn', FkyCnt = @FkyCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FkyTbl = LEFT(dfn.FkyTbl,30)
            ,FkyKys = LEFT(dfn.FkyKys,30)
            ,RkyTbl = LEFT(dfn.RkyTbl,30)
            ,RkyKys = LEFT(dfn.RkyKys,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #FkyDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Foreign References
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #RkyDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #RkyDfn (               -- DROP TABLE dbo.zzz_RkyDfn; CREATE TABLE dbo.zzz_RkyDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        FkyTbl                                            sysname              NOT NULL,
        FkyKys                                            varchar(1000)        NOT NULL,
        RkyTbl                                            sysname              NOT NULL,
        RkyKys                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_RkyDfn
    ------------------------------------------------------------------------------------------------
    IF @PopRky = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfRky; EXEC ut_zzRSJ RKY,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #RkyDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, FkyTbl, FkyKys, RkyTbl, RkyKys, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_RkyDefs CURSOR LOCAL FOR SELECT *                              FROM #RkyDfn ORDER BY RkyTbl  -- Cursor
    DECLARE @RkyCnt smallint; SET @RkyCnt = (SELECT COUNT(*)                   FROM #RkyDfn)                 -- Record Count
    DECLARE @RkyFln smallint; SET @RkyFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #RkyDfn)                 -- Max FieldName Length
    DECLARE @RkyTln smallint; SET @RkyTln = (SELECT ISNULL(MAX(LEN(RkyTbl)),0) FROM #RkyDfn)                 -- Max TableName Length
    DECLARE @RkyVln smallint; SET @RkyVln = CASE WHEN @RkyFln < @MinLenVAR THEN @MinLenVAR ELSE @RkyFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#RkyDfn' = '#RkyDfn', RkyCnt = @RkyCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,FkyTbl = LEFT(dfn.FkyTbl,30)
            ,FkyKys = LEFT(dfn.FkyKys,30)
            ,RkyTbl = LEFT(dfn.RkyTbl,30)
            ,RkyKys = LEFT(dfn.RkyKys,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #RkyDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Default Constraints
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #DefDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #DefDfn (               -- DROP TABLE dbo.zzz_DefDfn; CREATE TABLE dbo.zzz_DefDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        ClmNam                                            sysname              NOT NULL,
        ConTxt                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_DefDfn
    ------------------------------------------------------------------------------------------------
    IF @PopDef = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfDef; EXEC ut_zzRSJ DEF,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@ClmNam,@ConTxt,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #DefDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, ClmNam, ConTxt, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@ClmNam,@ConTxt,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_DefDefs CURSOR LOCAL FOR SELECT *                              FROM #DefDfn ORDER BY ConTbl  -- Cursor
    DECLARE @DefCnt smallint; SET @DefCnt = (SELECT COUNT(*)                   FROM #DefDfn)                 -- Record Count
    DECLARE @DefFln smallint; SET @DefFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #DefDfn)                 -- Max FieldName Length
    DECLARE @DefTln smallint; SET @DefTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #DefDfn)                 -- Max TableName Length
    DECLARE @DefVln smallint; SET @DefVln = CASE WHEN @DefFln < @MinLenVAR THEN @MinLenVAR ELSE @DefFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#DefDfn' = '#DefDfn', DefCnt = @DefCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,ClmNam = LEFT(dfn.ClmNam,30)
            ,ConTxt = LEFT(dfn.ConTxt,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #DefDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Check Constraints
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #ChkDfn temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #ChkDfn (               -- DROP TABLE dbo.zzz_ChkDfn; CREATE TABLE dbo.zzz_ChkDfn (
    ------------------------------------------------------------------------------------------------
        DfnID                                             smallint             NOT NULL,
        DfnCls                                            char(3)              NOT NULL,
        DfnExs                                            bit                  NOT NULL,
        DfnFmt                                            smallint             NOT NULL,
        DfnFmx                                            varchar(3)           NOT NULL,
        DfnFix                                            bit                  NOT NULL,
        ConTbl                                            sysname              NOT NULL,
        ConNam                                            sysname              NOT NULL,
        ConDsc                                            varchar(200)         NOT NULL,
        ClmNam                                            sysname              NOT NULL,
        ConTxt                                            varchar(1000)        NOT NULL,
        DfnStd                                            sysname              NOT NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate temp table from a dynamic cursor                        EXEC ut_zzSQL INX,zzz_ChkDfn
    ------------------------------------------------------------------------------------------------
    IF @PopChk = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @AsnDef = @AdfChk; EXEC ut_zzRSJ CHK,@TblLst,@ClmLst,@DtpLst,@BlkLst,@StpLst,@SetFmt,@AsnDef,@DfnCur OUTPUT; WHILE 1=1 BEGIN FETCH NEXT FROM @DfnCur INTO
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@ClmNam,@ConTxt,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            IF @@FETCH_STATUS <> 0 BREAK; INSERT INTO #ChkDfn (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                 DfnID, DfnCls, DfnExs, DfnFmt, DfnFmx, DfnFix, ConTbl, ConNam, ConDsc, ClmNam, ConTxt, DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            ) VALUES (
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
                @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@ClmNam,@ConTxt,@DfnStd
            ---------------------------------------------------------------------------------------------------------------------------------------------------------
            )
        -------------------------------------------------------------------------------------------------------------------------------------------------------------
        END; DEALLOCATE @DfnCur
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_ChkDefs CURSOR LOCAL FOR SELECT *                              FROM #ChkDfn ORDER BY ConTbl  -- Cursor
    DECLARE @ChkCnt smallint; SET @ChkCnt = (SELECT COUNT(*)                   FROM #ChkDfn)                 -- Record Count
    DECLARE @ChkFln smallint; SET @ChkFln = (SELECT ISNULL(MAX(LEN(ConNam)),0) FROM #ChkDfn)                 -- Max FieldName Length
    DECLARE @ChkTln smallint; SET @ChkTln = (SELECT ISNULL(MAX(LEN(ConTbl)),0) FROM #ChkDfn)                 -- Max TableName Length
    DECLARE @ChkVln smallint; SET @ChkVln = CASE WHEN @ChkFln < @MinLenVAR THEN @MinLenVAR ELSE @ChkFln END  -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display temp table values                                    EXEC ut_zzSQL SEL,zzz_ClmDfn,dfn
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR @DbgRsj = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#ChkDfn' = '#ChkDfn', ChkCnt = @ChkCnt
            ,DfnID  = dfn.DfnID
            ,DfnCls = dfn.DfnCls
            ,DfnExs = dfn.DfnExs
            ,DfnFmt = dfn.DfnFmt
            ,DfnFmx = dfn.DfnFmx
            ,DfnFix = dfn.DfnFix
            ,ConTbl = LEFT(dfn.ConTbl,30)
            ,ConNam = LEFT(dfn.ConNam,30)
            ,ConDsc = LEFT(dfn.ConDsc,30)
            ,ClmNam = LEFT(dfn.ClmNam,30)
            ,ConTxt = LEFT(dfn.ConTxt,30)
            ,DfnStd = LEFT(dfn.DfnStd,30)
        FROM
            #ChkDfn dfn
        ORDER BY
            dfn.DfnID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################
 
 
    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  Index Fragmentation
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    -- Create Table:  #FrgDfn temp table
    ------------------------------------------------------------------------------------------------


    --LST###########################################################################################
 

    ------------------------------------------------------------------------------------------------
    -- Create Primary Key List temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #PkyList (            -- DROP TABLE dbo.zzz_PkyList; CREATE TABLE dbo.zzz_PkyList (
    ------------------------------------------------------------------------------------------------
        TblID                                             int                      NULL,
        TblNam                                            sysname                  NULL,
        PkyNam                                            sysname                  NULL,
        PkyOrd                                            tinyint                  NULL,
        ClmID                                             int                      NULL,
        ClmNam                                            sysname                  NULL,
        IsClus                                            bit                      NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate Primary Key List temp table
    ------------------------------------------------------------------------------------------------
    INSERT INTO
        #PkyList
    SELECT
        pkl.TblID  AS TblID,
        pkl.ConTbl AS TblNam,
        pkl.ConNam AS PkyNam,
        pkl.KeyOrd AS PkyOrd,
        pkl.ClmID  AS ClmID,
        pkl.KeyClm AS ClmNam,
        pkl.IsClus AS IsClus
    FROM
        #PklDfn pkl
    ------------------------------------------------------------------------------------------------
    -- Delete PKeys which are also FKeys, i.e., are not the original PKey but borrow from another PKey
    ------------------------------------------------------------------------------------------------
    DELETE FROM #PkyList WHERE TblID IN (SELECT ref.fkeyid FROM SysReferences ref WHERE ref.fkeyid = TblID AND ClmID IN (
        ref.fkey1,ref.fkey2 ,ref.fkey3 ,ref.fkey4 ,ref.fkey5 ,ref.fkey6 ,ref.fkey7 ,ref.fkey8,
        ref.fkey9,ref.fkey10,ref.fkey11,ref.fkey12,ref.fkey13,ref.fkey14,ref.fkey15,ref.fkey16
    ))
    ------------------------------------------------------------------------------------------------
    -- Display Primary Key List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PkyList (CRT)' = '#PkyList (CRT)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #PkyList))
            ,TblID  = pkl.TblID
            ,TblNam = LEFT(pkl.TblNam,40)
            ,PkyNam = LEFT(pkl.PkyNam,60)
            ,PkyOrd = pkl.PkyOrd
            ,ClmID  = pkl.ClmID
            ,ClmNam = pkl.ClmNam
        FROM
            #PkyList pkl
        ORDER BY
            pkl.PkyOrd
        --RETURN
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Populate the Field lists from Column lists
    /*----------------------------------------------------------------------------------------------
        @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx;
         FldLvl, FldObj, FldOrd, FldNam, FldUtp, FldDtx, FldLen, FldDct, FldQot, FldNul, FldIdn, FldOup, FldPko, FldFko, FldLko, FldAud, FldVal, FldDfv, FldVfx
         FldID,FldLvl,FldObj,FldOrd,FldNam,FldUtp,FldDtx,FldLen,FldDct,FldQot,FldNul,FldIdn,FldOup,FldPko,FldFko,FldLko,FldAud,FldVal,FldDfv,FldVfx
        ---------------------------------------------------------------------------------------------
        @DfnID,@DfnCls,@DfnFix,@ClmLvl,@ClmObj,@ClmOrd,@ClmNam,@ClmUtp,@ClmDTX,@ClmSiz,@ClmLen,@ClmDct,@ClmQot,@ClmNul,@ClmIdn,@ClmCmp,@ClmOup,@ClmPky,@ClmUky,@ClmFky,@ClmXky,@ClmDef,@ClmVwd,@ClmAud,@ClmEmv,@ClmDfv,@ClmCpx
         DfnID, DfnCls, DfnFix, ClmLvl, ClmObj, ClmOrd, ClmNam, ClmUtp, ClmDTX, ClmSiz, ClmLen, ClmDct, ClmQot, ClmNul, ClmIdn, ClmCmp, ClmOup, ClmPky, ClmUky, ClmFky, ClmXky, ClmDef, ClmVwd, ClmAud, ClmEmv, ClmDfv, ClmCpx
        ---------------------------------------------------------------------------------------------
        ClmLvl                                            tinyint              NOT NULL,
        ClmObj                                            sysname              NOT NULL,
        ClmOrd                                            smallint             NOT NULL,
        ClmNam                                            sysname              NOT NULL,
        ClmUtp                                            int                  NOT NULL,
        ClmDTX                                            varchar(21)          NOT NULL,
        ClmSiz                                            smallint             NOT NULL, ****
        ClmLen                                            smallint             NOT NULL,
        ClmDct                                            varchar(3)           NOT NULL,
        ClmQot                                            bit                  NOT NULL,
        ClmNul                                            bit                  NOT NULL,
        ClmIdn                                            bit                  NOT NULL,
        ClmCmp                                            bit                  NOT NULL, ****
        ClmOup                                            bit                  NOT NULL,
        ClmPky                                            tinyint              NOT NULL, ****
        ClmUky                                            tinyint              NOT NULL, ****
        ClmFky                                            bit                  NOT NULL, ****
        ClmXky                                            bit                  NOT NULL, ****
        ClmDef                                            bit                  NOT NULL, ****
        ClmVwd                                            bit                  NOT NULL, ****
        ClmAud                                            bit                  NOT NULL, ****
        ClmEmv                                            varchar(100)         NOT NULL,
        ClmDfv                                            varchar(100)         NOT NULL,
        ClmCpx                                            varchar(4000)        NOT NULL
    ----------------------------------------------------------------------------------------------*/


    ------------------------------------------------------------------------------------------------
    -- Declare FieldList object variables
    ------------------------------------------------------------------------------------------------
    DECLARE @FldID     smallint        ; SET @FldID     = 0
    DECLARE @FldLvl    tinyint         ; SET @FldLvl    = 0
    DECLARE @FldObj    sysname         ; SET @FldObj    = ''
    DECLARE @FldOrd    smallint        ; SET @FldOrd    = 0
    DECLARE @FldNam    sysname         ; SET @FldNam    = ''
    DECLARE @FldUtp    smallint        ; SET @FldUtp    = 0
    DECLARE @FldDtx    varchar(21)     ; SET @FldDtx    = ''
    DECLARE @FldSiz    smallint        ; SET @FldSiz    = 0
    DECLARE @FldLen    smallint        ; SET @FldLen    = 0
    DECLARE @FldDct    varchar(3)      ; SET @FldDct    = ''
    DECLARE @FldQot    bit             ; SET @FldQot    = 0
    DECLARE @FldNul    char(9)         ; SET @FldNul    = ''
    DECLARE @FldIdn    varchar(9)      ; SET @FldIdn    = ''
    DECLARE @FldCmp    bit             ; SET @FldCmp    = 0
    DECLARE @FldOup    varchar(7)      ; SET @FldOup    = ''
    DECLARE @FldPko    tinyint         ; SET @FldPko    = 0
    DECLARE @FldUko    tinyint         ; SET @FldUko    = 0
    DECLARE @FldFko    tinyint         ; SET @FldFko    = 0
    DECLARE @FldLko    tinyint         ; SET @FldLko    = 0
    DECLARE @FldInd    bit             ; SET @FldInd    = 0
    DECLARE @FldDef    bit             ; SET @FldDef    = 0
    DECLARE @FldChk    bit             ; SET @FldChk    = 0
    DECLARE @FldVwd    bit             ; SET @FldVwd    = 0
    DECLARE @FldAud    bit             ; SET @FldAud    = 0
    DECLARE @FldVal    varchar(100)    ; SET @FldVal    = ''
    DECLARE @FldDfv    varchar(100)    ; SET @FldDfv    = ''
    DECLARE @FldVfx    varchar(4000)   ; SET @FldVfx    = ''
    DECLARE @FldFst    sysname         ; SET @FldFst    = ''
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Create FieldList temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #FldList (            -- DROP TABLE dbo.zzz_FldList; CREATE TABLE dbo.zzz_FldList (
    ------------------------------------------------------------------------------------------------
        FldID                                             smallint                 NULL,
        FldLvl                                            tinyint                  NULL,
        FldObj                                            sysname                  NULL,
        FldOrd                                            smallint                 NULL,
        FldNam                                            sysname                  NULL,
        FldUtp                                            int                      NULL,
        FldDtx                                            varchar(21)              NULL,
        FldLen                                            smallint                 NULL,
        FldDct                                            varchar(3)               NULL,
        FldQot                                            tinyint                  NULL,
        FldNul                                            char(9)                  NULL,
        FldIdn                                            varchar(9)               NULL,
        FldOup                                            varchar(7)               NULL,
        FldPko                                            tinyint                  NULL,
        FldFko                                            tinyint                  NULL,
        FldLko                                            tinyint                  NULL,
        FldAud                                            bit                      NULL,
        FldVal                                            varchar(100)             NULL,
        FldDfv                                            varchar(100)             NULL,
        FldVfx                                            varchar(4000)            NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate the FieldList temp table                           EXEC ut_zzSQL INZ,zzz_FldList,clm
    ------------------------------------------------------------------------------------------------
    INSERT INTO #FldList (
    ----------------------------------------------------------------------------------------------------------------------------------------------
        FldID,FldLvl,FldObj,FldOrd,FldNam,FldUtp,FldDtx,FldLen,FldDct,FldQot,FldNul,FldIdn,FldOup,FldPko,FldFko,FldLko,FldAud,FldVal,FldDfv,FldVfx
    ----------------------------------------------------------------------------------------------------------------------------------------------
    ) SELECT
    ------------------------------------------------------------------------------------------------
         FldID  = clm.DfnID                                                                       -- smallint
        ,FldLvl = clm.ClmLvl                                                                      -- tinyint
        ,FldObj = clm.ClmObj                                                                      -- sysname
        ,FldOrd = clm.ClmOrd                                                                      -- smallint
        ,FldNam = clm.ClmNam                                                                      -- sysname
        ,FldUtp = clm.ClmUtp                                                                      -- int
        ,FldDtx = clm.ClmDtx                                                                      -- varchar(21)
        ,FldLen = clm.ClmLen                                                                      -- smallint
        ,FldDct = clm.ClmDct                                                                      -- varchar(3)
        ,FldQot = clm.ClmQot                                                                      -- tinyint
        ,FldNul = CASE WHEN clm.ClmNul = 1 THEN @ClmNulALN ELSE @ClmNulNNL END                    -- char(9)
        ,FldIdn = CASE WHEN clm.ClmIdn = 1 THEN @ClmIdtYID ELSE @ClmIdtNID END                    -- varchar(9)
        ,FldOup = CASE WHEN clm.ClmOup = 1 THEN @PrmOupTXT ELSE ''         END                    -- varchar(7)
        ,FldPko = clm.ClmPky                                                                      -- tinyint
        ,FldFko = clm.ClmFky                                                                      -- tinyint
        ,FldLko = 0                                                                               -- tinyint
        ,FldAud = clm.ClmAud                                                                      -- bit
        ,FldVal = clm.ClmEmv                                                                      -- varchar(100)
        ,FldDfv = clm.ClmDfv                                                                      -- varchar(100)
        ,FldVfx = clm.ClmCpx                                                                      -- varchar(4000)
    ------------------------------------------------------------------------------------------------
    FROM
        #ClmDfn clm
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_FldList CURSOR LOCAL FOR SELECT *                              FROM #FldList ORDER BY FldLvl,FldOrd  -- Cursor
    DECLARE @FldCnt smallint; SET @FldCnt = (SELECT COUNT(*)                   FROM #FldList)                        -- Record Count
    DECLARE @FldFln smallint; SET @FldFln = (SELECT ISNULL(MAX(LEN(FldNam)),0) FROM #FldList)                        -- Max FieldName Length
    DECLARE @FldTln smallint; SET @FldTln = (SELECT ISNULL(MAX(LEN(FldDtx)),0) FROM #FldList)                        -- Max TableName Length
    DECLARE @FldVln smallint; SET @FldVln = CASE WHEN @FldFln < @MinLenVAR THEN @MinLenVAR ELSE @FldFln END          -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table                      EXEC ut_zzSQL SEL,zzz_FldList,fld
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (CRT)' = '#FldList (CRT)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Create Default Fields temp table
    ------------------------------------------------------------------------------------------------
    CREATE TABLE #DflList (            -- DROP TABLE dbo.zzz_DflList; CREATE TABLE dbo.zzz_DflList (
    ------------------------------------------------------------------------------------------------
        FldID                                             smallint                 NULL,
        FldLvl                                            tinyint                  NULL,
        FldObj                                            sysname                  NULL,
        FldOrd                                            smallint                 NULL,
        FldNam                                            sysname                  NULL,
        FldUtp                                            int                      NULL,
        FldDtx                                            varchar(21)              NULL,
        FldLen                                            smallint                 NULL,
        FldDct                                            varchar(3)               NULL,
        FldQot                                            tinyint                  NULL,
        FldNul                                            char(9)                  NULL,
        FldIdn                                            varchar(9)               NULL,
        FldOup                                            varchar(7)               NULL,
        FldPko                                            tinyint                  NULL,
        FldFko                                            tinyint                  NULL,
        FldLko                                            tinyint                  NULL,
        FldAud                                            bit                      NULL,
        FldVal                                            varchar(100)             NULL,
        FldDfv                                            varchar(100)             NULL,
        FldVfx                                            varchar(4000)            NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate the temp table
    ------------------------------------------------------------------------------------------------
    INSERT INTO #DflList (
    ----------------------------------------------------------------------------------------------------------------------------------------------
        FldID,FldLvl,FldObj,FldOrd,FldNam,FldUtp,FldDtx,FldLen,FldDct,FldQot,FldNul,FldIdn,FldOup,FldPko,FldFko,FldLko,FldAud,FldVal,FldDfv,FldVfx
    ----------------------------------------------------------------------------------------------------------------------------------------------
    ) SELECT
    ------------------------------------------------------------------------------------------------
         FldID  = clm.DfnID                                                                       -- smallint
        ,FldLvl = clm.ClmLvl                                                                      -- tinyint
        ,FldObj = clm.ClmObj                                                                      -- sysname
        ,FldOrd = clm.ClmOrd                                                                      -- smallint
        ,FldNam = clm.ClmNam                                                                      -- sysname
        ,FldUtp = clm.ClmUtp                                                                      -- int
        ,FldDtx = clm.ClmDtx                                                                      -- varchar(21)
        ,FldLen = clm.ClmLen                                                                      -- smallint
        ,FldDct = clm.ClmDct                                                                      -- varchar(3)
        ,FldQot = clm.ClmQot                                                                      -- tinyint
        ,FldNul = CASE WHEN clm.ClmNul = 1 THEN @ClmNulALN ELSE @ClmNulNNL END                    -- char(9)
        ,FldIdn = CASE WHEN clm.ClmIdn = 1 THEN @ClmIdtYID ELSE @ClmIdtNID END                    -- varchar(9)
        ,FldOup = CASE WHEN clm.ClmOup = 1 THEN @PrmOupTXT ELSE ''         END                    -- varchar(7)
        ,FldPko = clm.ClmPky                                                                      -- tinyint
        ,FldFko = clm.ClmFky                                                                      -- tinyint
        ,FldLko = 0                                                                               -- tinyint
        ,FldAud = clm.ClmAud                                                                      -- bit
        ,FldVal = clm.ClmEmv                                                                      -- varchar(100)
        ,FldDfv = clm.ClmDfv                                                                      -- varchar(100)
        ,FldVfx = clm.ClmCpx                                                                      -- varchar(4000)
    FROM
        #DcvDfn clm
    ------------------------------------------------------------------------------------------------
    -- Initialize cursor tracking variables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_DflList CURSOR LOCAL FOR SELECT *                              FROM #DflList ORDER BY FldLvl,FldOrd  -- Cursor
    DECLARE @DflCnt smallint; SET @DflCnt = (SELECT COUNT(*)                   FROM #DflList)                        -- Record Count
    DECLARE @DflFln smallint; SET @DflFln = (SELECT ISNULL(MAX(LEN(FldNam)),0) FROM #DflList)                        -- Max FieldName Length
    DECLARE @DflTln smallint; SET @DflTln = (SELECT ISNULL(MAX(LEN(FldDtx)),0) FROM #DflList)                        -- Max TableName Length
    DECLARE @DflVln smallint; SET @DflVln = CASE WHEN @DflFln < @MinLenVAR THEN @MinLenVAR ELSE @DflFln END          -- Min VariableName Length
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#DflList (CRT)' = '#DflList (CRT)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #DflList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #DflList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Object Definitions:  VBA Values
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ TBXTBL,zzz_TEST01
    ----------------------------------------------------------------------------------------------*/
    CREATE TABLE #VbaDfn (               -- DROP TABLE dbo.zzz_VbaDfn; CREATE TABLE dbo.zzz_VbaDfn (
    ------------------------------------------------------------------------------------------------
        VbaID                                             smallint                 NULL,
        OutFlg                                            bit                      NULL,
        IdnFlg                                            bit                      NULL,
        UpcFlg                                            bit                      NULL,
        CboFlg                                            bit                      NULL,
        ChkFlg                                            bit                      NULL,
        DirPrp                                            tinyint                  NULL,
        VarCat                                            varchar(10)              NULL,
        VarPfx                                            varchar(10)              NULL,
        VarNam                                            sysname                  NULL,
        VarDtp                                            varchar(30)              NULL,
        VarDfc                                            varchar(30)              NULL,
        VarVal                                            varchar(30)              NULL,
        VarNul                                            varchar(30)              NULL,
        CtlPfx                                            varchar(10)              NULL,
        CtlNam                                            sysname                  NULL,
        PrmDtp                                            varchar(20)              NULL,
        PrmDir                                            varchar(20)              NULL,
        PrmLen                                            varchar(10)              NULL
    ------------------------------------------------------------------------------------------------
    )
    ------------------------------------------------------------------------------------------------
    -- Populate VBA column table                                       EXEC ut_zzSQL INX,zzz_FldList
    ------------------------------------------------------------------------------------------------
    DECLARE cur_VbaDef CURSOR LOCAL FAST_FORWARD FOR SELECT * FROM #FldList ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    OPEN cur_VbaDef; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaDef INTO 
    ------------------------------------------------------------------------------------------------------------------------------------------------------------------
        @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx
    ------------------------------------------------------------------------------------------------------------------------------------------------------------------
    IF @@FETCH_STATUS <> 0 BREAK
    ------------------------------------------------------------------------------------------------
        SET @VbaID  = @FldID
        SET @DtpNam = @FldDtx
        --------------------------------------------------------------------------------------------
        -- Assign first field
        --------------------------------------------------------------------------------------------
        SET @FstFld = CASE WHEN LEN(@FstFld) = 0 THEN @FldNam ELSE @FstFld END
        --------------------------------------------------------------------------------------------
        -- Resolve variable category
        --------------------------------------------------------------------------------------------
        SET @VarCat = CASE 
            WHEN @DtpNam LIKE 'bit'           THEN @VbaCatBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'int'           THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'integer'       THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'money'         THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'real%'         THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'float%'        THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaCatNUM
            WHEN @DtpNam LIKE 'char%'         THEN @VbaCatTXT
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaCatTXT
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaCatTXT
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaCatTXT
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaCatTXT
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaCatDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaCatDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaCatVRN
            ELSE                                   @VbaCatTXT
        END
        --------------------------------------------------------------------------------------------
        -- Resolve variable prefix
        --------------------------------------------------------------------------------------------
        SET @VarPfx = CASE 
            WHEN @DtpNam LIKE 'bit'           THEN @VbaPfxBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaPfxINT
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaPfxINT
            WHEN @DtpNam LIKE 'int'           THEN @VbaPfxLNG
            WHEN @DtpNam LIKE 'integer'       THEN @VbaPfxLNG
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaPfxLNG
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaPfxCUR
            WHEN @DtpNam LIKE 'money'         THEN @VbaPfxCUR
            WHEN @DtpNam LIKE 'real%'         THEN @VbaPfxSGL
            WHEN @DtpNam LIKE 'float%'        THEN @VbaPfxDBL
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaPfxDBL
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaPfxDBL
            WHEN @DtpNam LIKE 'char%'         THEN @VbaPfxSTR
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaPfxSTR
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaPfxSTR
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaPfxSTR
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaPfxSTR
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaPfxDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaPfxDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaPfxVRN
            ELSE                                   @VbaPfxSTR
        END
        --------------------------------------------------------------------------------------------
        -- Assign variable datatype
        --------------------------------------------------------------------------------------------
        SET @VarDtp = CASE
            WHEN @DtpNam LIKE 'bit'           THEN @VbaDtpBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaDtpINT
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaDtpINT
            WHEN @DtpNam LIKE 'int'           THEN @VbaDtpLNG
            WHEN @DtpNam LIKE 'integer'       THEN @VbaDtpLNG
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaDtpLNG
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaDtpCUR
            WHEN @DtpNam LIKE 'money'         THEN @VbaDtpCUR
            WHEN @DtpNam LIKE 'real%'         THEN @VbaDtpSGL
            WHEN @DtpNam LIKE 'float%'        THEN @VbaDtpDBL
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaDtpDBL
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaDtpDBL
            WHEN @DtpNam LIKE 'char%'         THEN @VbaDtpSTR
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaDtpSTR
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaDtpSTR
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaDtpSTR
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaDtpSTR
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaDtpDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaDtpDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaDtpVRN
            ELSE                                   @VbaDtpSTR
        END
        --------------------------------------------------------------------------------------------
        -- Assign default constant
        --------------------------------------------------------------------------------------------
        SET @VarDfc = CASE
            WHEN @DtpNam LIKE 'bit'           THEN @VbaDfcBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'int'           THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'integer'       THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'money'         THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'real%'         THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'float%'        THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaDfcNUM
            WHEN @DtpNam LIKE 'char%'         THEN @VbaDfcTXT
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaDfcTXT
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaDfcTXT
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaDfcTXT
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaDfcTXT
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaDfcDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaDfcDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaDfcVRN
            ELSE                                   @VbaDfcTXT
        END
        --------------------------------------------------------------------------------------------
        -- Assign default value
        --------------------------------------------------------------------------------------------
        SET @VarVal = CASE
            WHEN @DtpNam LIKE 'bit'           THEN @VbaDfvBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'int'           THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'integer'       THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'money'         THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'real%'         THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'float%'        THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaDfvNUM
            WHEN @DtpNam LIKE 'char%'         THEN @VbaDfvTXT
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaDfvTXT
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaDfvTXT
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaDfvTXT
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaDfvTXT
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaDfvDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaDfvDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaDfvVRN
            ELSE                                   @VbaDfvTXT
        END
        --------------------------------------------------------------------------------------------
        -- Assign null value
        --------------------------------------------------------------------------------------------
        SET @VarNul = CASE
            WHEN @DtpNam LIKE 'bit'           THEN @VbaNulBLN
            WHEN @DtpNam LIKE 'tinyint'       THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'smallint'      THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'int'           THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'integer'       THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'bigint'        THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'smallmoney'    THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'money'         THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'real%'         THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'float%'        THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'decimal%'      THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'numeric%'      THEN @VbaNulNUM
            WHEN @DtpNam LIKE 'char%'         THEN @VbaNulTXT
            WHEN @DtpNam LIKE 'varchar%'      THEN @VbaNulTXT
            WHEN @DtpNam LIKE 'nchar%'        THEN @VbaNulTXT
            WHEN @DtpNam LIKE 'nvarchar%'     THEN @VbaNulTXT
            WHEN @DtpNam LIKE 'sysname'       THEN @VbaNulTXT
            WHEN @DtpNam LIKE 'smalldatetime' THEN @VbaNulDAT
            WHEN @DtpNam LIKE 'datetime'      THEN @VbaNulDAT
            WHEN @DtpNam LIKE 'sql_variant'   THEN @VbaNulVRN
            ELSE                                   @VbaNulTXT
        END
        --------------------------------------------------------------------------------------------
        -- Assign variable name
        --------------------------------------------------------------------------------------------
        SET @VarNam = @VarPfx+@FldNam 
        --------------------------------------------------------------------------------------------
        -- Resolve control prefix
        --------------------------------------------------------------------------------------------
        SET @CtlPfx = CASE 
            WHEN RIGHT(@FldNam,2) = 'ID'   THEN @CtlPfxCBO
            WHEN RIGHT(@FldNam,3) = 'Flg'  THEN @CtlPfxCHK
            WHEN RIGHT(@FldNam,4) = 'Flag' THEN @CtlPfxCHK
            ELSE                                @CtlPfxTXT
        END
        --------------------------------------------------------------------------------------------
        -- Assign Output flag
        --------------------------------------------------------------------------------------------
        SET @OutFlg = CASE 
            WHEN LEN(@FldOup) > 0                             THEN 1
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Assign Identity flag
        --------------------------------------------------------------------------------------------
        SET @IdnFlg = CASE 
            WHEN LEN(@FldOup) > 0 AND RIGHT(@FldNam,2) = "ID" THEN 1
            WHEN LEN(@FldIdn) > 0                             THEN 1
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Assign ComboBox flag
        --------------------------------------------------------------------------------------------
        SET @CboFlg = CASE 
            WHEN @CtlPfx IN (@CtlPfxCBO)                 THEN 1
            WHEN @FldNam IN ('TaxYer','BegYer','EndYer') THEN 1
            WHEN @FldNam IN ('TaxPrd','BegPrd','EndPrd') THEN 1
            WHEN @FldNam IN ('TaxMon','BegMon','EndMon') THEN 1
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Assign CheckBox flag
        --------------------------------------------------------------------------------------------
        SET @ChkFlg = CASE 
            WHEN @CtlPfx IN (@CtlPfxCHK) THEN 1
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Resolve control prefix
        --------------------------------------------------------------------------------------------
        SET @CtlPfx = CASE 
            WHEN @CboFlg = 1 THEN @CtlPfxCBO
            WHEN @ChkFlg = 1 THEN @CtlPfxCHK
            ELSE                  @CtlPfx
        END
        --------------------------------------------------------------------------------------------
        -- Assign direction flag
        --------------------------------------------------------------------------------------------
        SET @DirPrp = CASE
            WHEN LEN(@FldOup) > 0 THEN 3
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Assign control name
        --------------------------------------------------------------------------------------------
        SET @CtlNam = @CtlPfx+@FldNam 
        --------------------------------------------------------------------------------------------
        -- Assign upper case flag
        --------------------------------------------------------------------------------------------
        SET @UpcFlg = CASE
            WHEN @CtlPfx = @CtlPfxTXT AND RIGHT(@FldNam,3) = 'Cod'  THEN 1
            WHEN @CtlPfx = @CtlPfxTXT AND RIGHT(@FldNam,4) = 'Code' THEN 1
            WHEN @CtlPfx = @CtlPfxTXT AND @FldLen <= 5              THEN 1
            ELSE 0
        END
        --------------------------------------------------------------------------------------------
        -- Assign parameter datatype
        --------------------------------------------------------------------------------------------
        SET @PrmDtp = CASE 
            WHEN @DtpNam LIKE 'bit'           THEN "adBoolean"
            WHEN @DtpNam LIKE 'tinyint'       THEN "adTinyInt"
            WHEN @DtpNam LIKE 'smallint'      THEN "adSmallInt"
            WHEN @DtpNam LIKE 'int'           THEN "adInteger"
            WHEN @DtpNam LIKE 'integer'       THEN "adInteger"
            WHEN @DtpNam LIKE 'bigint'        THEN "adInteger"
            WHEN @DtpNam LIKE 'smallmoney'    THEN "adCurrency"
            WHEN @DtpNam LIKE 'money'         THEN "adCurrency"
            WHEN @DtpNam LIKE 'real%'         THEN "adSingle"
            WHEN @DtpNam LIKE 'float%'        THEN "adDouble"
            WHEN @DtpNam LIKE 'decimal%'      THEN "adDecimal"
            WHEN @DtpNam LIKE 'numeric%'      THEN "adNumeric"
            WHEN @DtpNam LIKE 'char%'         THEN "adChar"
            WHEN @DtpNam LIKE 'varchar%'      THEN "adVarChar"
            WHEN @DtpNam LIKE 'nchar%'        THEN "adWChar"
            WHEN @DtpNam LIKE 'nvarchar%'     THEN "adWVarChar"
            WHEN @DtpNam LIKE 'sysname'       THEN "adVarChar"
            WHEN @DtpNam LIKE 'smalldatetime' THEN "adDate"
            WHEN @DtpNam LIKE 'datetime'      THEN "adDate"
            WHEN @DtpNam LIKE 'sql_variant'   THEN "adPropVariant"
            ELSE                                   "adVarChar"
        END
        --------------------------------------------------------------------------------------------
        -- Assign parameter direction
        --------------------------------------------------------------------------------------------
        SET @PrmDir = CASE
            WHEN @DirPrp = 1 THEN @PrmDirINP
            WHEN @DirPrp = 2 THEN @PrmDirIOP
            WHEN @DirPrp = 3 THEN @PrmDirOUP
            WHEN @DirPrp = 4 THEN @PrmDirRTN
            WHEN @DirPrp = 5 THEN @PrmDirUNK
            ELSE                  @PrmDirINP
        END
        --------------------------------------------------------------------------------------------
        -- Assign parameter length
        --------------------------------------------------------------------------------------------
        SET @PrmLen = CAST(CASE 
            WHEN @FldDtx LIKE 'sysname'       THEN 128
            ELSE @FldLen
        END AS varchar(10))
        --------------------------------------------------------------------------------------------
        -- Insert the record                                            EXEC ut_zzSQL INX,zzz_VbaDfn
        --------------------------------------------------------------------------------------------
        INSERT INTO #VbaDfn (
        ----------------------------------------------------------------------------------------------------------------------------------------------------------
             VbaID, OutFlg, IdnFlg, UpcFlg, CboFlg, ChkFlg, DirPrp, VarCat, VarPfx, VarNam, VarDtp, VarDfc, VarVal, VarNul, CtlPfx, CtlNam, PrmDtp, PrmDir, PrmLen
        ----------------------------------------------------------------------------------------------------------------------------------------------------------
        ) VALUES (
        ----------------------------------------------------------------------------------------------------------------------------------------------------------
            @VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen
        ----------------------------------------------------------------------------------------------------------------------------------------------------------
        )
    ------------------------------------------------------------------------------------------------
    END; DEALLOCATE cur_VbaDef
    ------------------------------------------------------------------------------------------------
    DECLARE cur_VbaLst CURSOR LOCAL FOR SELECT * FROM #VbaDfn ORDER BY VbaID  -- Cursor
    SELECT @VbaCnt = COUNT(*)                    FROM #VbaDfn                 -- Record Count
    SELECT @VbaVln = ISNULL(MAX(LEN(VarNam)),0)  FROM #VbaDfn                 -- Max Length
    SELECT @VbaTln = ISNULL(MAX(LEN(VarDtp)),0)  FROM #VbaDfn                 -- Max Length
    SELECT @VbaPln = ISNULL(MAX(LEN(PrmDtp)),0)  FROM #VbaDfn                 -- Max Length
    SELECT @VbaDln = ISNULL(MAX(LEN(PrmDir)),0)  FROM #VbaDfn                 -- Max Length
    SELECT @VbaLln = ISNULL(MAX(LEN(PrmLen)),0)  FROM #VbaDfn                 -- Max Length
    ------------------------------------------------------------------------------------------------
    -- Display VBA Definition List temp table                       EXEC ut_zzSQL SEL,zzz_VbaDfn,vba
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT @N+@LinDbl+@N+'#VbaDfn: '+@InpObj+' = '+@InpSIX
        SELECT
             VbaID  = vba.VbaID
            ,OutFlg = vba.OutFlg
            ,IdnFlg = vba.IdnFlg
            ,UpcFlg = vba.UpcFlg
            ,CboFlg = vba.CboFlg
            ,ChkFlg = vba.ChkFlg
            ,DirPrp = vba.DirPrp
            ,VarCat = vba.VarCat
            ,VarPfx = vba.VarPfx
            ,VarNam = LEFT(vba.VarNam,30)
            ,VarDtp = vba.VarDtp
            ,VarDfc = vba.VarDfc
            ,VarVal = vba.VarVal
            ,VarNul = vba.VarNul
            ,CtlPfx = vba.CtlPfx
            ,CtlNam = LEFT(vba.CtlNam,30)
            ,PrmDtp = vba.PrmDtp
            ,PrmDir = vba.PrmDir
            ,PrmLen = vba.PrmLen
        FROM
            #VbaDfn vba
        ORDER BY
            VbaID
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Declare dimension/fact flags
    ------------------------------------------------------------------------------------------------
    DECLARE @IncDmm    tinyint      ; SET @IncDmm    = 0                   -- Include dimension master columns
    DECLARE @IncFtk    tinyint      ; SET @IncFtk    = 0                   -- Include fact keymap columns
    DECLARE @IncFtm    tinyint      ; SET @IncFtm    = 0                   -- Include fact master columns
    DECLARE @IncFtx    tinyint      ; SET @IncFtx    = 0                   -- Include fact exception columns
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Set dimension/fact flags
    ------------------------------------------------------------------------------------------------
    IF (@DefTyp IN ('DMM',            'DIM','DM0','DM1','DM2','DML')) SET @IncDim = 1
    IF (@DefTyp IN ('DMM'                  ,'DM0','DM1','DM2'      )) SET @IncDmm = 1
    ------------------------------------------------------------------------------------------------
    IF (@DefTyp IN ('FTK','FTM','FTX','FCT','FT0',            'FTL')) SET @IncFct = 1
    IF (@DefTyp IN ('FTK'                                          )) SET @IncFtk = 1
    IF (@DefTyp IN (      'FTM'            ,'FT0'                  )) SET @IncFtm = 1
    IF (@DefTyp IN (            'FTX'                              )) SET @IncFtx = 1
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Update Field List parameter values (UDF)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ dmm_TblDfn,CKC
    ----------------------------------------------------------------------------------------------*/
    UPDATE fld SET
        FldLvl = CASE
            WHEN dfl.FldLvl IN (@FlvDMM,@FlvFTM,@FlvSRC) THEN dfl.FldLvl
            WHEN dfl.FldLvl IN (@FlvSTD,@FlvLKP,@FlvLNK) THEN fld.FldLvl
            ELSE                                              dfl.FldLvl
        END,
        FldObj = dfl.FldObj,
        FldOrd = CASE
            WHEN dfl.FldLvl IN (@FlvSTD,@FlvLKP,@FlvLNK) THEN fld.FldOrd
            ELSE                                              dfl.FldOrd
        END,
        FldLen = dfl.FldLen,
        FldDct = dfl.FldDct,
        FldQot = dfl.FldQot,
        FldNul = dfl.FldNul,
        FldAud = dfl.FldAud,
        FldVal = dfl.FldVal,
        FldDfv = dfl.FldDfv,
        FldVfx = dfl.FldVfx
    FROM
        #FldList fld
    INNER JOIN
        #DflList dfl
            ON dfl.FldNam = fld.FldNam
    ------------------------------------------------------------------------------------------------
    IF @IncDim = 1 OR @TctDIM = 1 BEGIN
        UPDATE
            #FldList
        SET
            FldLvl = @FlvSRC
        WHERE
            LEFT(FldNam,3) IN ('Src') AND RIGHT(FldNam,2) IN ('ID')
    END
    ------------------------------------------------------------------------------------------------
    -- UPDATE
    --     fld
    -- SET
    --     FldLvl = dfl.FldLvl
    -- FROM
    --     #FldList fld
    -- INNER JOIN
    --     #DflList dfl
    --         ON dfl.FldNam = fld.FldNam
    --        AND dfl.FldLvl NOT IN (@FlvSTD,@FlvLKP,@FlvLNK)
    ------------------------------------------------------------------------------------------------
    -- Update Field List default constraint values
    ------------------------------------------------------------------------------------------------
    IF @OupFmt IN ('SCM') BEGIN
        UPDATE fld SET
            FldDfv = def.ConTxt
        FROM
            #FldList fld
        INNER JOIN
            #DefDfn  def
                ON fld.FldObj = def.ConTbl
               AND fld.FldNam = def.ClmNam
    END
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (UDF)' = '#FldList (UDF)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Include template default fields (TPL)
    ------------------------------------------------------------------------------------------------
    IF @FldCnt = 0 BEGIN
    ------------------------------------------------------------------------------------------------
        IF @StpTBL = 1 BEGIN
            IF          @TblCat IN (@TblCatSEC) BEGIN
                INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvSEC,@FlvCRT,@FlvUPD,@FlvEXP)
            END ELSE IF @TblCat IN (@TblCatLKP) BEGIN
                INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvLKP,@FlvCRT,@FlvUPD,@FlvEXP)
            END ELSE IF @TblCat IN (@TblCatLNK) BEGIN
                INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvLNK,@FlvCRT,@FlvUPD)
            END ELSE BEGIN
                INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvSTD)
                IF @IncAud IS NULL BEGIN
                    EXEC ut_zzNAM TBL,AUD,XXX,@InpObj,@IncAud OUTPUT
                END
            END
            SET @IdnClm = @TblCpx+'ID'
        END ELSE IF @StpUFN = 1 BEGIN
            INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvRTN)
        END
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (TPL)' = '#FldList (TPL)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
 

    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Resolve dimension fields (DIM)
    ------------------------------------------------------------------------------------------------
    IF @IncDim = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        -- ANY:  Force PK Identity
        --------------------------------------------------------------------------------------------
        IF @IncDmm = 1 BEGIN
            UPDATE #FldList SET FldIdn = @ClmIdtYID WHERE FldPko = 1 AND RIGHT(FldNam,2) IN ('ID')
        END ELSE BEGIN
            UPDATE #FldList SET FldIdn = ''         WHERE FldPko = 1 AND RIGHT(FldNam,2) IN ('ID')
        END
        --------------------------------------------------------------------------------------------
        -- ANY:  Exclude audit columns
        --------------------------------------------------------------------------------------------
        DELETE FROM #FldList WHERE FldLvl IN (@FlvFTM,@FlvDSB,@FlvDLT,@FlvLOK,@FlvCRT,@FlvUPD,@FlvEXP,@FlvDEL)
        --DELETE FROM #FldList WHERE FldNam IN ('Discontinued')
        --------------------------------------------------------------------------------------------
        -- ANY:  Flag SourceID's
        --------------------------------------------------------------------------------------------
        UPDATE #FldList SET FldLvl = @FlvSRC WHERE LEFT(FldNam,3) IN ('Src') AND RIGHT(FldNam,2) IN ('ID')
        --------------------------------------------------------------------------------------------
        -- ANY:  Align support objects
        --------------------------------------------------------------------------------------------
        IF @DspObj <> @InpObj BEGIN
        --------------------------------------------------------------------------------------------
            UPDATE #PkyDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #UkyDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #IndDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #FkyDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #RkyDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #DefDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
            UPDATE #ChkDfn SET ConTbl = REPLACE(ConTbl,@InpObj,@DspObj),ConNam = REPLACE(ConNam,@InpObj,@DspObj),ConDsc = REPLACE(ConDsc,@InpObj,@DspObj)
        --------------------------------------------------------------------------------------------
        END
        --------------------------------------------------------------------------------------------
        -- DIM:  Exclude dimension source IDs
        --------------------------------------------------------------------------------------------
        IF @IncDmm = 0 BEGIN
            DELETE FROM #FldList WHERE FldLvl IN (@FlvSRC)
        END
        --------------------------------------------------------------------------------------------
        -- DMM:  Dimension source IDs
        --------------------------------------------------------------------------------------------
        IF @IncDmm = 1 BEGIN
        --------------------------------------------------------------------------------------------
            UPDATE #FldList SET FldLvl = @FlvSRC WHERE LEFT(FldNam,3) IN ('Src') AND RIGHT(FldNam,2) IN ('ID')
            IF @DspExs = 0 BEGIN
                -- Exclude support objects
                DELETE FROM #UkyDfn; SET @UkyCnt = 0
                DELETE FROM #IndDfn; SET @IndCnt = 0
                DELETE FROM #FkyDfn; SET @FkyCnt = 0
                DELETE FROM #RkyDfn; SET @RkyCnt = 0
                DELETE FROM #DefDfn; SET @DefCnt = 0
                DELETE FROM #ChkDfn; SET @ChkCnt = 0
                -- Include default dimension source IDs
                DECLARE cur_NewSrc CURSOR LOCAL FOR
                    SELECT
                        FldNam
                    FROM
                        #FldList
                    WHERE
                        FldPko > 0
                    ORDER BY
                        FldLvl,
                        FldOrd
                OPEN cur_NewSrc
                WHILE 1=1 BEGIN FETCH NEXT FROM cur_NewSrc INTO @FldNam; IF @@FETCH_STATUS <> 0 BREAK
                    SET @TX1 = 'Src'+@FldNam
                    IF EXISTS (SELECT * FROM #FldList WHERE FldNam = @TX1) BEGIN
                        UPDATE #FldList SET FldLvl = @FlvSRC WHERE FldNam = @TX1
                    END ELSE BEGIN
                    --------------------------------------------------------------------------------
                        INSERT INTO #FldList (
                        -----------------------------------------------------------------------------------------------------------------------------------------------------
                            FldLvl,FldObj,FldOrd,FldNam,FldUtp,FldDtx,FldLen,FldDct,FldQot,FldNul,FldIdn,FldOup,FldPko,FldFko,FldLko,FldAud,FldVal,FldDfv,FldVfx
                        -----------------------------------------------------------------------------------------------------------------------------------------------------
                        ) SELECT
                        ----------------------------------------------------------------------------
                            FldLvl = @FlvSRC,
                            FldObj = fld.FldObj,
                            FldOrd = fld.FldOrd,
                            FldNam = @TX1,
                            FldUtp = fld.FldUtp,
                            FldDtx = fld.FldDtx,
                            FldLen = fld.FldLen,
                            FldDct = fld.FldDct,
                            FldQot = fld.FldQot,
                            FldNul = @ClmNulNNL,
                            FldIdn = '',
                            FldOup = '',
                            FldPko = 0,
                            FldFko = 0,
                            FldLko = 0,
                            FldAud = 0,
                            FldVal = fld.FldVal,
                            FldDfv = fld.FldDfv,
                            FldVfx = fld.FldVfx
                        FROM
                            #FldList fld
                        WHERE
                            fld.FldNam = @FldNam
                    --------------------------------------------------------------------------------
                    END
                END
                DEALLOCATE cur_NewSrc
            END
        END
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (DIM)' = '#FldList (DIM)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Resolve fact fields (FCT)
    ------------------------------------------------------------------------------------------------
    IF @IncFct = 1 BEGIN
        -- Clear audit columns
        DELETE FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvDSB,@FlvDLT,@FlvLOK,@FlvCRT,@FlvUPD,@FlvEXP,@FlvDEL)
        -- Set PKey identity as bigint
        UPDATE #FldList SET FldDtx = 'bigint',FldLen = 16 WHERE FldPko = 1 AND LEN(FldIdn) > 0
    END
    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (FCT)' = '#FldList (FCT)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------

    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Reset default parameter values (DEF)
    ------------------------------------------------------------------------------------------------
    SET @HasHtk = 0
    SET @HasDim = 0
    SET @HasFct = 0
    SET @HasDsb = 0
    SET @HasDlt = 0
    SET @HasLok = 0
    SET @HasCrt = 0
    SET @HasUpd = 0
    SET @HasExp = 0
    SET @HasAud = 0
    SET @HasHst = 0
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Process standard columns
    -- @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@UspNam,@DepUsp,@DepFnc,@DepTrg,@UseTbl,@UseVew,@UseUsp,@UseFnc,@UseTrg,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
    -- @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@UspNam,@DepUsp,@DepFnc,@DepTrg,@UseTbl,@UseVew,@UseUsp,@UseFnc,@UseTrg,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd
    ------------------------------------------------------------------------------------------------
    IF @StpUSP = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        -- Assign SProc Object values
        OPEN cur_UspDefs; WHILE 1=1 BEGIN FETCH NEXT FROM cur_UspDefs INTO @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@UspNam,@DepUsp,@DepFnc,@DepTrg,@UseTbl,@UseVew,@UseUsp,@UseFnc,@UseTrg,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd; IF @@FETCH_STATUS <> 0 BREAK;
            BREAK
        END; CLOSE cur_UspDefs
    ------------------------------------------------------------------------------------------------
    END ELSE IF @StpVEW = 1 BEGIN
    ------------------------------------------------------------------------------------------------
        -- Assign View Object values
        OPEN cur_VewDefs; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VewDefs INTO @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@VewNam,@RecQty,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd; IF @@FETCH_STATUS <> 0 BREAK;
            BREAK
        END; CLOSE cur_VewDefs
    ------------------------------------------------------------------------------------------------
    END ELSE BEGIN
    ------------------------------------------------------------------------------------------------
        -- Assign Table Object values
        OPEN cur_TblDefs; WHILE 1=1 BEGIN FETCH NEXT FROM cur_TblDefs INTO @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@TblNam,@TblSiz,@RecQty,@RecSiz,@HasPky,@UkyQty,@FkyQty,@RkyQty,@IndQty,@DefQty,@ChkQty,@ClxNam,@HasHtk,@HasDim,@HasFct,@HasDsb,@HasDlt,@HasLok,@HasCrt,@HasUpd,@HasExp,@HasDel,@HasHst,@HasAud,@DfnStd; IF @@FETCH_STATUS <> 0 BREAK;
            BREAK
        END; CLOSE cur_TblDefs
    ------------------------------------------------------------------------------------------------
    END
 
    ------------------------------------------------------------------------------------------------
    -- Align standard parameter values
    ------------------------------------------------------------------------------------------------
    SET @IncDim = ISNULL(@IncDim,@HasDim)
    SET @IncFct = ISNULL(@IncFct,@HasFct)
    SET @IncDsb = ISNULL(@IncDsb,@HasDsb)
    SET @IncDlt = ISNULL(@IncDlt,@HasDlt)
    SET @IncLok = ISNULL(@IncLok,@HasLok)
    SET @IncCrt = ISNULL(@IncCrt,@HasCrt)
    SET @IncUpd = ISNULL(@IncUpd,@HasUpd)
    SET @IncExp = ISNULL(@IncExp,@HasExp)
    SET @IncAud = ISNULL(@IncAud,@HasAud)
    SET @IncHst = ISNULL(@IncHst,@HasHst)
 
    ------------------------------------------------------------------------------------------------
    -- Include standard columns (delete any stragglers first)
    ------------------------------------------------------------------------------------------------
    IF (@IncDim = 1               ) AND @HasDim = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvDMM)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvDMM) ORDER BY FldLvl,FldOrd
    END
    IF (@IncFct = 1               ) AND @HasFct = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvFTM)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvFTM) ORDER BY FldLvl,FldOrd
    END
    IF (@IncDsb = 1               ) AND @HasDsb = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvDSB)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvDSB) ORDER BY FldLvl,FldOrd
    END
    IF (@IncDlt = 1               ) AND @HasDlt = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvDLT)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvDLT) ORDER BY FldLvl,FldOrd
    END
    IF (@IncLok = 1               ) AND @HasLok = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvLOK)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvLOK) ORDER BY FldLvl,FldOrd
    END
    IF (@IncAud = 1 OR @IncCrt = 1) AND @HasCrt = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvCRT)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvCRT) ORDER BY FldLvl,FldOrd
    END
    IF (@IncAud = 1 OR @IncUpd = 1) AND @HasUpd = 0 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvUPD)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD) ORDER BY FldLvl,FldOrd
    END
    IF (@IncExp = 1               )                 BEGIN
        DELETE FROM #FldList                        WHERE FldLvl IN (@FlvEXP)
        INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvEXP) ORDER BY FldLvl,FldOrd
    END
    -- (@IncHst = 1               ) AND @HasHtk = 0 INSERT INTO #FldList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK) ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (DEF)' = '#FldList (DEF)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    -- Process Empty PKey fields (substitute UKey)
    IF @HasPky = 0 AND @UkyQty > 0  BEGIN
        --------------------------------------------------------------------------------------------
        -- Create Primary Key Column temp table
        --------------------------------------------------------------------------------------------
        CREATE TABLE #PkyColm ( -- DROP TABLE dbo.zzz_PkyColm; CREATE TABLE dbo.zzz_PkyColm (
            TblID                                             int                      NULL,
            TblNam                                            sysname                  NULL,
            PkyNam                                            sysname                  NULL,
            PkyOrd                                            tinyint                  NULL,
            ClmID                                             int                      NULL,
            ClmNam                                            sysname                  NULL,
            IsClus                                            bit                      NULL
        )
        --------------------------------------------------------------------------------------------
        INSERT INTO
            #PkyColm
        SELECT
            TblID  = obj.id,
            TblNam = obj.name,
            PkyNam = con.name,
            PkyOrd = spt.number,
            ClmID  = clm.colid,
            ClmNam = clm.name,
            IsClus = CASE WHEN ind.indid = 1 THEN 1 ELSE 0 END
        FROM
            SysObjects obj
        INNER JOIN
            SysObjects con
                ON con.parent_obj = obj.id
        INNER JOIN
            SysIndexes ind
                ON ind.id   = obj.id
               AND ind.name = con.name
        INNER JOIN
            SysColumns clm
                ON clm.id = obj.id
        INNER JOIN
            master.dbo.spt_values spt
                ON spt.number >= 1
               AND spt.number <= ind.keycnt
               AND spt.type    = 'P'
        WHERE
            obj.xtype IN ('U')
        AND con.xtype IN ('UQ','UK')
        AND clm.name = index_col(obj.name,ind.indid,spt.number)
        AND obj.id = @SrcID
        ORDER BY
            spt.number
        --------------------------------------------------------------------------------------------
        -- Can only reference one UKey
        SET @TXT = ISNULL((SELECT TOP 1 PkyNam FROM #PkyColm),"")
        DELETE FROM #PkyColm WHERE PkyNam <> @TXT
        --------------------------------------------------------------------------------------------
        UPDATE
            fld
        SET
            FldPko = 1
        FROM
            #FldList fld
        INNER JOIN
            #PkyColm pky
                ON pky.ClmNam = fld.FldNam
    END


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Process Lookup fields
    ------------------------------------------------------------------------------------------------
    IF LEN(@LkpLst) > 0 BEGIN
        SET @IDN = 0
        SET @LST = @LkpLst
        WHILE LEN(@LST) > 0 BEGIN
            SET @POS = CHARINDEX(",",@LST)
            IF @POS > 0 BEGIN
                SET @ITM = RTRIM(LEFT(@LST,@POS - 1))
                SET @LST = LTRIM(RIGHT(@LST,LEN(@LST) - @POS))
            END ELSE BEGIN
                SET @ITM = @LST
                SET @LST = ""
            END
            SET @IDN += 1
            UPDATE #FldList SET FldLko = @IDN WHERE FldNam = @ITM
        END
    END
 

    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Assign identity text
    ------------------------------------------------------------------------------------------------
    UPDATE #FldList SET FldIdn = @ClmIdtYID WHERE FldNam = @IdnClm

    ------------------------------------------------------------------------------------------------
    -- Set audit flag
    ------------------------------------------------------------------------------------------------
    --UPDATE #FldList SET FldAud = 1 WHERE FldLvl IN (@FlvLOK,@FlvLOD,@FlvMOD,@FlvEXP,@FlvEXP,@FlvCRT,@FlvUPD,@FlvHST)


    --##############################################################################################

 
    ------------------------------------------------------------------------------------------------
    -- Assign unique IDs to FldList
    ------------------------------------------------------------------------------------------------
    DECLARE cur_FldNumb CURSOR LOCAL FOR
        SELECT
            FldObj,
            FldNam
        FROM
            #FldList
        ORDER BY
            FldLvl,
            FldOrd
    SET @IDN = 0
    OPEN cur_FldNumb
    WHILE 1=1 BEGIN
        FETCH NEXT FROM cur_FldNumb INTO @FldObj,@FldNam
        IF @@FETCH_STATUS <> 0 BREAK
        SET @IDN += 1
        UPDATE #FldList SET FldID = @IDN WHERE FldObj = @FldObj AND FldNam = @FldNam
    END
    DEALLOCATE cur_FldNumb
 
    ------------------------------------------------------------------------------------------------
    -- Standardize all regular fields
    ------------------------------------------------------------------------------------------------
    UPDATE #FldList SET FldLvl = @FlvSTD WHERE FldLvl IN (@FlvLKP,@FlvLNK,@FlvPRS,@FlvSEC)


    --##############################################################################################

 
    ------------------------------------------------------------------------------------------------
    -- Create working temp tables from Field List
    ------------------------------------------------------------------------------------------------
    SELECT * INTO #PkfList FROM #FldList WHERE FldPko > 0               ORDER BY FldLvl,FldOrd
    SELECT * INTO #LkpList FROM #FldList WHERE FldPko > 0 OR FldLko > 0 ORDER BY FldLvl,FldOrd
    SELECT * INTO #ColList FROM #FldList WHERE 1=0
    SELECT * INTO #RtnList FROM #FldList WHERE 1=0
    SELECT * INTO #SigList FROM #FldList WHERE 1=0
    SELECT * INTO #StxList FROM #FldList WHERE 1=0
    SELECT * INTO #InsList FROM #FldList WHERE 1=0
    SELECT * INTO #UpdList FROM #FldList WHERE 1=0
    SELECT * INTO #HstList FROM #FldList WHERE 1=0
    SELECT * INTO #AudList FROM #FldList WHERE 1=0
    SELECT * INTO #PopList FROM #FldList WHERE 1=0
    SELECT * INTO #SetList FROM #FldList WHERE 1=0
    SELECT * INTO #CrtList FROM #FldList WHERE 1=0
    SELECT * INTO #SrcList FROM #FldList WHERE 1=0
    SELECT * INTO #StdList FROM #FldList WHERE 1=0
    SELECT * INTO #MfkList FROM #FldList WHERE 1=0

    ------------------------------------------------------------------------------------------------
    -- Update temp table values based on utility 
    ------------------------------------------------------------------------------------------------
    IF @OupFmt IN ('INS') BEGIN
        UPDATE #FldList SET FldOup = @PrmOupTXT WHERE FldIdn = @ClmIdtYID
    END ELSE IF @OupFmt IN ('USP') AND @DefTyp IN ('INS','ADN','UPS') BEGIN
        UPDATE #FldList SET FldOup = @PrmOupTXT WHERE FldIdn = @ClmIdtYID
    END

    ------------------------------------------------------------------------------------------------
    -- Build template lists
    ------------------------------------------------------------------------------------------------
    -- Select:
    --   > Sig: Include PKs and custom Lookups as filter criteria
    --   > Stm: Include all standard columns (no audit, etc)
    SELECT * INTO #SelSig FROM #FldList WHERE  FldPko > 0 OR FldLko > 0                                                                                                               ORDER BY FldLvl,FldOrd
    SELECT * INTO #SelStm FROM #FldList WHERE  FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD)                                                                                            ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Insert:
    --   > Sig: Return identity column value as OUTPUT
    --   > Stm: Cannot insert identity column
    --   > Stm: Cannot insert computed columns
    --   > Stm: Disabled,Deleted,Locked are inserted intitially
    ------------------------------------------------------------------------------------------------
    SELECT * INTO #InsSig FROM #FldList WHERE  FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,@FlvDSB,@FlvDLT,@FlvLOK,@FlvCRT) AND LEN(       FldVfx) = 0                                 ORDER BY FldLvl,FldOrd
    SELECT * INTO #InsStm FROM #FldList WHERE  FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,@FlvDSB,@FlvDLT,@FlvLOK,@FlvCRT) AND LEN(FldIdn+FldVfx) = 0                                 ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Update:
    --   > Sig: Include PK as filter criteria (assumes an Identity column is also the PK)
    --   > Stm: Cannot update identity column
    --   > Stm: Cannot update computed columns
    --   > Stm: No need to update PKs since they should not change
    --   > Stm: Disabled,Deleted,Locked are handled by separate mechanisms
    ------------------------------------------------------------------------------------------------
    SELECT * INTO #UpdSig FROM #FldList WHERE (FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,                        @FlvUPD) OR  FldLko > 0) AND LEN(       FldVfx) = 0         ORDER BY FldLvl,FldOrd
    SELECT * INTO #UpdStm FROM #FldList WHERE  FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,                        @FlvUPD) AND FldPko = 0  AND LEN(FldIdn+FldVfx) = 0         ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Delete:
    --   > Sig: Include PKs and custom Lookups as filter criteria
    --   > Stm: None
    SELECT * INTO #DelSig FROM #FldList WHERE  FldPko > 0 OR FldLko > 0                                                                                                               ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Upsert:
    --   > Sig: Return identity column value as OUTPUT
    --   > Stm: Cannot insert/update identity column
    --   > Stm: Cannot insert/update computed columns
    --   > Stm: No need to update PKs since they should not change
    --   > Stm: Disabled,Deleted,Locked are inserted intitially
    --   > Stm: Disabled,Deleted,Locked are handled by separate mechanisms
    ------------------------------------------------------------------------------------------------
    SELECT * INTO #UpsSig FROM #FldList WHERE (FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,@FlvDSB,@FlvDLT,@FlvLOK,@FlvUPD,@FlvCRT) OR  FldLko > 0) AND LEN(       FldVfx) = 0 ORDER BY FldLvl,FldOrd
    SELECT * INTO #UpsStm FROM #FldList WHERE  FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM,@FlvSTD,@FlvDSB,@FlvDLT,@FlvLOK,@FlvUPD,@FlvCRT) AND FldPko = 0  AND LEN(FldIdn+FldVfx) = 0 ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- Build fixed lists
    ------------------------------------------------------------------------------------------------
    -- Insert/Update:
    --   > Always include all columns except identity
    ------------------------------------------------------------------------------------------------
    INSERT INTO #InsList SELECT * FROM #InsStm                            ORDER BY FldLvl,FldOrd
    INSERT INTO #UpdList SELECT * FROM #UpdStm                            ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Column:
    --   > Always include all columns
    ------------------------------------------------------------------------------------------------
    INSERT INTO #ColList SELECT * FROM #FldList WHERE 1=1                 ORDER BY FldLvl,FldOrd
    IF @IncHst = 1 BEGIN
    INSERT INTO #ColList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK) ORDER BY FldLvl,FldOrd
    UPDATE      #ColList SET FldIdn = ""
    DELETE FROM #ColList WHERE LEN(FldVfx) = 1
    END
    ------------------------------------------------------------------------------------------------
    -- Populate:
    --   > Can    insert identity column!
    --   > CanNOT insert computed columns
    ------------------------------------------------------------------------------------------------
    INSERT INTO #PopList SELECT * FROM #FldList WHERE LEN(FldVfx) = 0     ORDER BY FldLvl,FldOrd
    IF @IncHst = 1 BEGIN
    INSERT INTO #PopList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK) ORDER BY FldLvl,FldOrd
    UPDATE      #PopList SET FldIdn = ""
    END
    ------------------------------------------------------------------------------------------------
    -- Dimension source ID records:
    --   > Include dimension source ID columns
    ------------------------------------------------------------------------------------------------
    INSERT INTO #SrcList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC) ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Missing ForeignKeys:
    --   > Include all fields which join to PKeys from other tables
    --   > Exclude current table columns
    --   > Exclude already created FKeys
    ------------------------------------------------------------------------------------------------
    INSERT INTO #MfkList SELECT fld.* FROM
        #FldList fld
    INNER JOIN
        #PkyList pky
            ON pky.ClmNam = fld.FldNam
    WHERE
        pky.TblNam <> fld.FldObj
    ORDER BY
        fld.FldLvl,
        fld.FldOrd
    /*----------------------------------------------------------------------------------------------
        Delete FKeys which already exist
        --------------------------------------------------------------------------------------------
        SELECT
            TblNam = LEFT(obj.name,40),
            ClmNam = LEFT(clm.name,40)
        FROM
            SysObjects    obj
        INNER JOIN
            SysColumns    clm
                ON clm.id = obj.id
        INNER JOIN
            SysReferences ref
                ON ref.fkeyid = obj.id
        WHERE
            clm.colid IN (
                ref.fkey1,ref.fkey2 ,ref.fkey3 ,ref.fkey4 ,ref.fkey5 ,ref.fkey6 ,ref.fkey7 ,ref.fkey8,
                ref.fkey9,ref.fkey10,ref.fkey11,ref.fkey12,ref.fkey13,ref.fkey14,ref.fkey15,ref.fkey16
            )
        AND obj.name = 'lnk_Test01'
        AND clm.name = 'AppDfnID'
        --------------------------------------------------------------------------------------------
        EXEC ut_zzSQJ lnk_Test01,TBA
        SELECT * FROM #PkyList
        SELECT * FROM #FldList
        SELECT * FROM #MfkList
    ----------------------------------------------------------------------------------------------*/
    DELETE FROM #MfkList WHERE EXISTS (
        SELECT
            *
        FROM
            SysObjects    obj
        INNER JOIN
            SysColumns    clm
                ON clm.id = obj.id
        INNER JOIN
            SysReferences ref
                ON ref.fkeyid = obj.id
        WHERE
            clm.colid IN (
                ref.fkey1,ref.fkey2 ,ref.fkey3 ,ref.fkey4 ,ref.fkey5 ,ref.fkey6 ,ref.fkey7 ,ref.fkey8,
                ref.fkey9,ref.fkey10,ref.fkey11,ref.fkey12,ref.fkey13,ref.fkey14,ref.fkey15,ref.fkey16
            )
        AND obj.name = FldObj
        AND clm.name = FldNam
    )
    /*----------------------------------------------------------------------------------------------
        SELECT * FROM #MfkList
    ----------------------------------------------------------------------------------------------*/
 
    ------------------------------------------------------------------------------------------------
    -- Build SET recordset from SET list
    ------------------------------------------------------------------------------------------------
    SET @TXT = ','+REPLACE(@SetLst," ","")+','
    INSERT INTO #SetList SELECT * FROM #FldList WHERE @TXT LIKE '%,'+FldNam+',%' ORDER BY FldLvl,FldOrd
 

    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Fields for:  Dimensions
    ------------------------------------------------------------------------------------------------
    IF @DefTyp IN ('DMS') BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvELD,@FlvDBG)                 ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #InsStm                                                    ORDER BY FldLvl,FldOrd
    END ELSE IF @DefTyp IN ('DML') BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvELD,@FlvDBG)                 ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #FldList WHERE FldLvl IN (@FlvPKY,@FlvDMM,@FlvSTD)         ORDER BY FldLvl,FldOrd
    END ELSE IF @DefTyp IN ('DM0','DM1','DM2') BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvELD,@FlvDBG)                 ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #InsStm                                                    ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  FactTable Current Import
    ------------------------------------------------------------------------------------------------
    END ELSE IF @DefTyp IN ('FTI','FTL') BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvELD)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #InsStm                                                    ORDER BY FldLvl,FldOrd
    ------------------------------------------------------------------------------------------------
    -- Fields for:  FactTable Current Builders
    ------------------------------------------------------------------------------------------------
    END ELSE IF @DefTyp IN ('FT0') BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvELD)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #InsStm                                                    ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Temp Table
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('TPL') BEGIN
        INSERT INTO #SigList SELECT * FROM #SelSig                                                    ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #ColList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK)                         ORDER BY FldLvl,FldOrd
        UPDATE      #ColList SET FldIdn = ""
        DELETE FROM #ColList WHERE LEN(FldVfx) = 1
        END
        --------------------------------------------------------------------------------------------
        INSERT INTO #StxList SELECT * FROM #FldList WHERE 1=1                                         ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE 1=1                                         ORDER BY FldLvl,FldOrd
        END
        --------------------------------------------------------------------------------------------
        DELETE FROM #PopList WHERE LEN(FldIdn) > 0

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Schema
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('SCM') BEGIN
        INSERT INTO #StxList SELECT * FROM #ColList                                                   ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Table
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('TBL') BEGIN
        IF          @DefTyp IN ('DIM') BEGIN
            INSERT INTO #StxList SELECT * FROM #ColList WHERE FldLvl NOT IN (@FlvSRC)                 ORDER BY FldLvl,FldOrd
        END ELSE IF @DefTyp IN ('FCT') BEGIN
            INSERT INTO #StxList SELECT * FROM #ColList WHERE FldLvl NOT IN (@FlvSRC)                 ORDER BY FldLvl,FldOrd
        END ELSE BEGIN
            INSERT INTO #StxList SELECT * FROM #ColList                                               ORDER BY FldLvl,FldOrd
        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  View
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('VEW') BEGIN
        INSERT INTO #StxList SELECT * FROM #SelStm                                                    ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Function
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('UFN','XUF') BEGIN
        INSERT INTO #RtnList SELECT * FROM #FldList WHERE FldOrd = 0                                  ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldOrd > 0                                  ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #FldList WHERE FldOrd > 0                                  ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Stored Procedure
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('USP','XSP') BEGIN
        IF @StpTBL = 1 BEGIN
            IF          @DefTyp IN ('SEL') BEGIN
                INSERT INTO #SigList SELECT * FROM #SelSig                                            ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #SelStm                                            ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('INS','ADN') BEGIN
                INSERT INTO #SigList SELECT * FROM #InsSig                                            ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #InsStm                                            ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('UPS') BEGIN
                INSERT INTO #SigList SELECT * FROM #UpsSig                                            ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #UpsStm                                            ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('UPD','MOD') BEGIN
                INSERT INTO #SigList SELECT * FROM #UpdSig                                            ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #UpdStm                                            ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('DEL') BEGIN
                INSERT INTO #SigList SELECT * FROM #SelSig                                            ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('PRK') BEGIN
                INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvPRS)                 ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #DflList WHERE FldLvl IN (@FlvPRS)                 ORDER BY FldLvl,FldOrd
            END ELSE IF @DefTyp IN ('BLD','BLX','TFR') BEGIN 
                INSERT INTO #StxList SELECT * FROM #ColList                                           ORDER BY FldLvl,FldOrd
            END ELSE BEGIN
                INSERT INTO #SigList SELECT * FROM #SelSig                                            ORDER BY FldLvl,FldOrd
                INSERT INTO #StxList SELECT * FROM #SelStm                                            ORDER BY FldLvl,FldOrd
            END

        END ELSE IF @StpUSP = 1 BEGIN
            INSERT INTO #SigList SELECT * FROM #ColList                                               ORDER BY FldLvl,FldOrd

        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Populate Data
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('POP') BEGIN
        INSERT INTO #StxList SELECT * FROM #PopList                                                   ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Select
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('SEL') BEGIN
        INSERT INTO #SigList SELECT * FROM #SelSig                                                    ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #SelStm                                                    ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Insert
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('INS','APN') BEGIN
        INSERT INTO #SigList SELECT * FROM #InsSig                                                    ORDER BY FldLvl,FldOrd
        IF @IncHst = 1 BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@TkhVar)                         ORDER BY FldLvl,FldOrd
        END
        --------------------------------------------------------------------------------------------
        INSERT INTO #StxList SELECT * FROM #InsStm                                                    ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM)         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSTD,@FlvDSB,@FlvDLT)         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvCRT)                         ORDER BY FldLvl,FldOrd
        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Update Record
    ------------------------------------------------------------------------------------------------
    --   > Disabled,Deleted,Locked are handled by separate mechanisms
    --   > Cannot update identity column
    --   > No need to update pkeys since they are the search criteria
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('UPD') BEGIN
        INSERT INTO #SigList SELECT * FROM #UpdSig                                                    ORDER BY FldLvl,FldOrd
        IF @IncHst = 1 BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@TkhVar)                         ORDER BY FldLvl,FldOrd
        END
        --------------------------------------------------------------------------------------------
        INSERT INTO #StxList SELECT * FROM #UpdStm                                                    ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM)         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSTD,@FlvDSB,@FlvDLT)         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Delete Record
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('DEL') BEGIN
        INSERT INTO #SigList SELECT * FROM #DelSig                                                    ORDER BY FldLvl,FldOrd
        IF @IncHst = 1 BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@TkhVar)                         ORDER BY FldLvl,FldOrd
        END
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE FldPko > 0 OR FldLko > 0                    ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Manage History
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('HST') BEGIN
        IF          @BldLST IN ('HSTAPN') BEGIN 
            INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                                        ORDER BY FldLvl,FldOrd
            INSERT INTO #SigList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM)                        ORDER BY FldLvl,FldOrd
            INSERT INTO #SigList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSTD,@FlvDSB,@FlvDLT)                        ORDER BY FldLvl,FldOrd
            INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHST)                                        ORDER BY FldLvl,FldOrd
            ----------------------------------------------------------------------------------------
            INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK)                                        ORDER BY FldLvl,FldOrd
            INSERT INTO #HstList SELECT * FROM #FldList WHERE 1=1                                                        ORDER BY FldLvl,FldOrd

        END ELSE IF @BldLST IN ('HSTCPY') BEGIN 
            INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0 OR FldLko > 0                                   ORDER BY FldLvl,FldOrd
            INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                                        ORDER BY FldLvl,FldOrd
            INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHST)                                        ORDER BY FldLvl,FldOrd
            --------------------------------------------------------------------------------------------
            INSERT INTO #StxList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM)                        ORDER BY FldLvl,FldOrd
            INSERT INTO #StxList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSTD,@FlvDSB,@FlvDLT) AND FldPko = 0 AND LEN(FldIdn+FldVfx) = 0 ORDER BY FldLvl,FldOrd
            ----------------------------------------------------------------------------------------
            INSERT INTO #HstList SELECT * FROM #DflList WHERE FldNam IN (@HttVar)                                        ORDER BY FldLvl,FldOrd
            INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSRC,@FlvDMM,@FlvFTM)                        ORDER BY FldLvl,FldOrd
            INSERT INTO #HstList SELECT * FROM #FldList WHERE FldLvl IN (@FlvSTD,@FlvDSB,@FlvDLT)                        ORDER BY FldLvl,FldOrd
            INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHST)                                        ORDER BY FldLvl,FldOrd

        END ELSE IF @BldLST IN ('HSTSYN') BEGIN 
            INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@CldVar)                                        ORDER BY FldLvl,FldOrd
            ----------------------------------------------------------------------------------------

        END ELSE  BEGIN 
            INSERT INTO #SigList SELECT * FROM #SelSig                                                                   ORDER BY FldLvl,FldOrd

        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Copy Record
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('CPY') BEGIN
        INSERT INTO #SigList SELECT * FROM #SelSig                                                    ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #ColList                                                   ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Set Column Values
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('SET') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0                                  ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #SetList WHERE 1=1                                         ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        INSERT INTO #StxList SELECT * FROM #SigList WHERE FldPko = 0                                  ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  DataSet
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('DST') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0                                  ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        IF @IncHst = 1 BEGIN
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@TkhVar)                         ORDER BY FldLvl,FldOrd
        END

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Manage Links
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('LNK') BEGIN
        SET @TXT = ','+REPLACE(@StdTx1," ","")+','
        INSERT INTO #SigList SELECT * FROM #FldList WHERE @TXT LIKE '%,'+FldNam+',%'                  ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@DlsVar,@AlsVar)                 ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldLvl IN (@FlvUPD)                         ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        INSERT INTO #StxList SELECT * FROM #DflList WHERE FldNam IN (@DlsVar,@AlsVar)                 ORDER BY FldLvl,FldOrd


    -- Fields for:  GetXxx
    END ELSE IF @OupFmt IN ('GET') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0                                  ORDER BY FldLvl,FldOrd

    -- Primary Key+TestMode criteria
    END ELSE IF @OupFmt IN ('DYN') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0                                  ORDER BY FldLvl,FldOrd
        INSERT INTO #SigList SELECT * FROM #DflList WHERE FldNam IN (@TmdVar)                         ORDER BY FldLvl,FldOrd

    -- Primary Key+Lookup List criteria
    END ELSE IF @OupFmt IN ('GKY') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldPko > 0 OR FldLko > 0                    ORDER BY FldLvl,FldOrd

    -- Lookup List criteria only
    END ELSE IF @OupFmt IN ('EXS') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE FldLko > 0                                  ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Maintain empty lists
    ------------------------------------------------------------------------------------------------
    END ELSE IF @OupFmt IN ('BLD','GAL','LOD') BEGIN
        INSERT INTO #SigList SELECT * FROM #FldList WHERE 1=0                                         ORDER BY FldLvl,FldOrd

    ------------------------------------------------------------------------------------------------
    -- Fields for:  Default
    ------------------------------------------------------------------------------------------------
    END ELSE BEGIN
        INSERT INTO #SigList SELECT * FROM #SelSig                                                    ORDER BY FldLvl,FldOrd
        INSERT INTO #StxList SELECT * FROM #ColList                                                   ORDER BY FldLvl,FldOrd
        --------------------------------------------------------------------------------------------
        IF @IncHst = 1 BEGIN
        INSERT INTO #HstList SELECT * FROM #DflList WHERE FldLvl IN (@FlvHTK)                         ORDER BY FldLvl,FldOrd
        INSERT INTO #HstList SELECT * FROM #FldList WHERE 1=1                                         ORDER BY FldLvl,FldOrd
        END

    END
    ------------------------------------------------------------------------------------------------

    ------------------------------------------------------------------------------------------------
    -- History does not allow identity columns
    ------------------------------------------------------------------------------------------------
    UPDATE #HstList SET FldIdn = ""

    ------------------------------------------------------------------------------------------------
    -- Reset field levels - DO THIS LAST!
    ------------------------------------------------------------------------------------------------
    UPDATE #FldList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #PkfList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #LkpList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #ColList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #RtnList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #SigList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #StxList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #InsList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #UpdList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #HstList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #AudList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #PopList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #SetList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #CrtList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #SrcList SET FldLvl = @FlvPKY WHERE FldPko > 0
    UPDATE #MfkList SET FldLvl = @FlvPKY WHERE FldPko > 0

    ------------------------------------------------------------------------------------------------
    -- Display Default Column List temp table
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#FldList (INS)' = '#FldList (INS)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #FldList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #FldList fld
        ORDER BY
            fld.FldLvl,
            fld.FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- OrderBy List:  Object definitions
    ------------------------------------------------------------------------------------------------
    IF @BldCOD IN ('TPLSQJ','TPLTSJ','TPLRSJ') BEGIN
        SET @ObyLst = "DfnID"
    END


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Reassign FieldList values
    ------------------------------------------------------------------------------------------------
    SELECT @FldCnt = COUNT(*)                   FROM #FldList
    SELECT @FldFln = ISNULL(MAX(LEN(FldNam)),0) FROM #FldList
    SELECT @FldTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #FldList
    SET @FldVln = CASE WHEN @FldFln < @MinLenVAR THEN @MinLenVAR ELSE @FldFln END

    ------------------------------------------------------------------------------------------------
    -- Create cursors and set column properties for working tables
    ------------------------------------------------------------------------------------------------
    DECLARE cur_PkfList CURSOR LOCAL FOR SELECT *                         FROM #PkfList ORDER BY FldLvl,FldOrd
    DECLARE @PkfCnt smallint; SELECT @PkfCnt = COUNT(*)                   FROM #PkfList
    DECLARE @PkfFln smallint; SELECT @PkfFln = ISNULL(MAX(LEN(FldNam)),0) FROM #PkfList
    DECLARE @PkfTln smallint; SELECT @PkfTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #PkfList
    DECLARE @PkfVln smallint; SET @PkfVln = CASE WHEN @PkfFln < @MinLenVAR THEN @MinLenVAR ELSE @PkfFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_LkpList CURSOR LOCAL FOR SELECT *                         FROM #LkpList ORDER BY FldLvl,FldOrd
    DECLARE @LkpCnt smallint; SELECT @LkpCnt = COUNT(*)                   FROM #LkpList
    DECLARE @LkpFln smallint; SELECT @LkpFln = ISNULL(MAX(LEN(FldNam)),0) FROM #LkpList
    DECLARE @LkpTln smallint; SELECT @LkpTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #LkpList
    DECLARE @LkpVln smallint; SET @LkpVln = CASE WHEN @LkpFln < @MinLenVAR THEN @MinLenVAR ELSE @LkpFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_ColList CURSOR LOCAL FOR SELECT *                         FROM #ColList ORDER BY FldLvl,FldOrd
    DECLARE @ColCnt smallint; SELECT @ColCnt = COUNT(*)                   FROM #ColList
    DECLARE @ColFln smallint; SELECT @ColFln = ISNULL(MAX(LEN(FldNam)),0) FROM #ColList
    DECLARE @ColTln smallint; SELECT @ColTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #ColList
    DECLARE @ColVln smallint; SET @ColVln = CASE WHEN @ColFln < @MinLenVAR THEN @MinLenVAR ELSE @ColFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_RtnList CURSOR LOCAL FOR SELECT *                         FROM #RtnList ORDER BY FldLvl,FldOrd
    DECLARE @RtnCnt smallint; SELECT @RtnCnt = COUNT(*)                   FROM #RtnList
    DECLARE @RtnFln smallint; SELECT @RtnFln = ISNULL(MAX(LEN(FldNam)),0) FROM #RtnList
    DECLARE @RtnTln smallint; SELECT @RtnTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #RtnList
                              SET @RtnFln = CASE WHEN @RtnFln < @UfnFln THEN @UfnFln ELSE @RtnFln END
    DECLARE @RtnVln smallint; SET @RtnVln = CASE WHEN @RtnFln < @MinLenVAR THEN @MinLenVAR ELSE @RtnFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_SigList CURSOR LOCAL FOR SELECT *                         FROM #SigList ORDER BY FldLvl,FldOrd
    DECLARE @SigCnt smallint; SELECT @SigCnt = COUNT(*)                   FROM #SigList
    DECLARE @SigFln smallint; SELECT @SigFln = ISNULL(MAX(LEN(FldNam)),0) FROM #SigList
    DECLARE @SigTln smallint; SELECT @SigTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #SigList
    DECLARE @SigVln smallint; SET @SigVln = CASE WHEN @SigFln < @MinLenVAR THEN @MinLenVAR ELSE @SigFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_StmList CURSOR LOCAL FOR SELECT *                         FROM #StxList ORDER BY FldLvl,FldOrd
    DECLARE @StxCnt smallint; SELECT @StxCnt = COUNT(*)                   FROM #StxList
    DECLARE @StxFln smallint; SELECT @StxFln = ISNULL(MAX(LEN(FldNam)),0) FROM #StxList
    DECLARE @StxTln smallint; SELECT @StxTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #StxList
    DECLARE @StxVln smallint; SET @StxVln = CASE WHEN @StxFln < @MinLenVAR THEN @MinLenVAR ELSE @StxFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_InsList CURSOR LOCAL FOR SELECT *                         FROM #InsList ORDER BY FldLvl,FldOrd
    DECLARE @InsCnt smallint; SELECT @InsCnt = COUNT(*)                   FROM #InsList
    DECLARE @InsFln smallint; SELECT @InsFln = ISNULL(MAX(LEN(FldNam)),0) FROM #InsList
    DECLARE @InsTln smallint; SELECT @InsTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #InsList
    DECLARE @InsVln smallint; SET @InsVln = CASE WHEN @InsFln < @MinLenVAR THEN @MinLenVAR ELSE @InsFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_UpdList CURSOR LOCAL FOR SELECT *                         FROM #UpdList ORDER BY FldLvl,FldOrd
    DECLARE @UpdCnt smallint; SELECT @UpdCnt = COUNT(*)                   FROM #UpdList
    DECLARE @UpdFln smallint; SELECT @UpdFln = ISNULL(MAX(LEN(FldNam)),0) FROM #UpdList
    DECLARE @UpdTln smallint; SELECT @UpdTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #UpdList
    DECLARE @UpdVln smallint; SET @UpdVln = CASE WHEN @UpdFln < @MinLenVAR THEN @MinLenVAR ELSE @UpdFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_HstList CURSOR LOCAL FOR SELECT *                         FROM #HstList ORDER BY FldLvl,FldOrd
    DECLARE @HstCnt smallint; SELECT @HstCnt = COUNT(*)                   FROM #HstList
    DECLARE @HstFln smallint; SELECT @HstFln = ISNULL(MAX(LEN(FldNam)),0) FROM #HstList
    DECLARE @HstTln smallint; SELECT @HstTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #HstList
    DECLARE @HstVln smallint; SET @HstVln = CASE WHEN @HstFln < @MinLenVAR THEN @MinLenVAR ELSE @HstFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_AudList CURSOR LOCAL FOR SELECT *                         FROM #DflList WHERE FldLvl IN (@FlvCRT,@FlvUPD) ORDER BY FldLvl,FldOrd
    DECLARE @AudCnt smallint; SELECT @AudCnt = COUNT(*)                   FROM #DflList WHERE FldLvl IN (@FlvCRT,@FlvUPD)
    DECLARE @AudFln smallint; SELECT @AudFln = ISNULL(MAX(LEN(FldNam)),0) FROM #DflList WHERE FldLvl IN (@FlvCRT,@FlvUPD)
    DECLARE @AudTln smallint; SELECT @AudTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #DflList WHERE FldLvl IN (@FlvCRT,@FlvUPD)
    DECLARE @AudVln smallint; SET @AudVln = CASE WHEN @AudFln < @MinLenVAR THEN @MinLenVAR ELSE @AudFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_PopList CURSOR LOCAL FOR SELECT *                         FROM #PopList ORDER BY FldLvl,FldOrd
    DECLARE @PopCnt smallint; SELECT @PopCnt = COUNT(*)                   FROM #PopList
    DECLARE @PopFln smallint; SELECT @PopFln = ISNULL(MAX(LEN(FldNam)),0) FROM #PopList
    DECLARE @PopTln smallint; SELECT @PopTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #PopList
    DECLARE @PopVln smallint; SET @PopVln = CASE WHEN @PopFln < @MinLenVAR THEN @MinLenVAR ELSE @PopFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_SetList CURSOR LOCAL FOR SELECT *                         FROM #SetList ORDER BY FldLvl,FldOrd
    DECLARE @SetCnt smallint; SELECT @SetCnt = COUNT(*)                   FROM #SetList
    DECLARE @SetFln smallint; SELECT @SetFln = ISNULL(MAX(LEN(FldNam)),0) FROM #SetList
    DECLARE @SetTln smallint; SELECT @SetTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #SetList
    DECLARE @SetVln smallint; SET @SetVln = CASE WHEN @SetFln < @MinLenVAR THEN @MinLenVAR ELSE @SetFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_CrtList CURSOR LOCAL FOR SELECT *                         FROM #CrtList ORDER BY FldLvl,FldOrd
    DECLARE @CrtCnt smallint; SELECT @CrtCnt = COUNT(*)                   FROM #CrtList
    DECLARE @CrtFln smallint; SELECT @CrtFln = ISNULL(MAX(LEN(FldNam)),0) FROM #CrtList
    DECLARE @CrtTln smallint; SELECT @CrtTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #CrtList
    DECLARE @CrtVln smallint; SET @CrtVln = CASE WHEN @CrtFln < @MinLenVAR THEN @MinLenVAR ELSE @CrtFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_SrcList CURSOR LOCAL FOR SELECT *                         FROM #SrcList ORDER BY FldLvl,FldOrd
    DECLARE @SrcCnt smallint; SELECT @SrcCnt = COUNT(*)                   FROM #SrcList
    DECLARE @SrcFln smallint; SELECT @SrcFln = ISNULL(MAX(LEN(FldNam)),0) FROM #SrcList
    DECLARE @SrcTln smallint; SELECT @SrcTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #SrcList
    DECLARE @SrcVln smallint; SET @SrcVln = CASE WHEN @SrcFln < @MinLenVAR THEN @MinLenVAR ELSE @SrcFln END
    ------------------------------------------------------------------------------------------------
    DECLARE cur_MfkList CURSOR LOCAL FOR SELECT *                         FROM #MfkList ORDER BY FldLvl,FldOrd
    DECLARE @MfkCnt smallint; SELECT @MfkCnt = COUNT(*)                   FROM #MfkList
    DECLARE @MfkFln smallint; SELECT @MfkFln = ISNULL(MAX(LEN(FldNam)),0) FROM #MfkList
    DECLARE @MfkTln smallint; SELECT @MfkTln = ISNULL(MAX(LEN(FldDtx)),0) FROM #MfkList
    DECLARE @MfkVln smallint; SET @MfkVln = CASE WHEN @MfkFln < @MinLenVAR THEN @MinLenVAR ELSE @MfkFln END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    -- Assign Field List values
    SELECT @IdnClm = ISNULL(FldNam,"")          FROM #FldList WHERE LEN(FldIdn) > 0
    SELECT @OupClm = ISNULL(FldNam,"")          FROM #FldList WHERE LEN(FldOup) > 0
    SELECT TOP 1 @FldFst = ISNULL(FldNam,"")    FROM #FldList WHERE FldLvl >= @FlvSTD ORDER BY FldLvl,FldOrd

    SET @IncIdn = CASE WHEN LEN(@IdnClm) > 0 THEN @IncIdn ELSE 0       END
    SET @HasIdn = CASE WHEN LEN(@IdnClm) > 0 THEN @IncIdn ELSE 0       END
    SET @HasOup = CASE WHEN LEN(@OupClm) > 0 THEN 1       ELSE 0       END

    -- Set minimum column length
    SET @StmCln = CASE WHEN @StmCln < @FldFln THEN @FldFln ELSE @StmCln END


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Include RetVal variable in parameter sizes
    ------------------------------------------------------------------------------------------------
    SET @RetFln = CASE WHEN LEN(@RetNam) > @SigFln THEN LEN(@RetNam) ELSE @SigFln END
    SET @RetTln = CASE WHEN LEN(@RetDtp) > @SigTln THEN LEN(@RetDtp) ELSE @SigTln END

    ------------------------------------------------------------------------------------------------
    -- Build PKY strings
    ------------------------------------------------------------------------------------------------
    IF @PkfCnt > 0 SET @PkyNam = ISNULL((SELECT TOP 1 FldNam FROM #PkfList ORDER BY FldLvl,FldOrd),"")
    SET @IDN = 0; SET @CNT = @PkfCnt; OPEN cur_PkfList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_PkfList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
        IF @IDN > 1 SET @CMA = "," ELSE SET @CMA = ""
        IF @IDN > 1 SET @ANX = " AND " ELSE SET @ANX = ""
        SET @PkyFch = @PkyFch+@CMA+" @"+@FldNam
        SET @PkyWhr = @PkyWhr+@ANX+@FldNam+" = @"+@FldNam
    END; CLOSE cur_PkfList
    SET @PkyFch = LTRIM(RTRIM(@PkyFch))
    SET @PkyWhr = LTRIM(RTRIM(@PkyWhr))

    ------------------------------------------------------------------------------------------------
    -- Build Field strings
    ------------------------------------------------------------------------------------------------
    SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
        IF @IDN > 1 SET @CMA = "," ELSE SET @CMA = ""
        SET @FldLin = @FldLin+@CMA+"" +@FldNam
        SET @VarLin = @VarLin+@CMA+"@"+@FldNam
    END; CLOSE cur_ColList
 
    ------------------------------------------------------------------------------------------------
    -- Check for history tracking and history variable declaration
    ------------------------------------------------------------------------------------------------
    SELECT @TrkHst = COUNT(*) FROM #SigList WHERE FldNam = @TkhVar
    SET @TkhDec = @TrkHst
 
    ------------------------------------------------------------------------------------------------
    -- Check for audit columns in signature column list
    ------------------------------------------------------------------------------------------------
    DECLARE @SigAud bit; SET @SigAud = (SELECT COUNT(*) FROM #SigList WHERE FldAud = 1 AND FldDct = @DtpCatNBR)


    --##############################################################################################

    ------------------------------------------------------------------------------------------------
    -- Display Default List templates
    ------------------------------------------------------------------------------------------------
    -- EXEC ut_zzSQJ zzz_TEST01,ZZZ
    ------------------------------------------------------------------------------------------------
    IF ((@DbgFlg = 1) AND 0=1) OR 0=9 BEGIN
        PRINT @N+@LinDbl+@N+'#DefTmpl: '+@InpObj+' = '+@InpSIX
        SELECT
            "INSERT INTO #DflList (FldLvl,FldObj,FldNam,FldDtx,FldLen,FldNul,FldIdn,FldOup,FldOrd,FldDct,FldQot,FldVal,FldDfv,FldPko,FldLko,FldFko,FldAud,FldVfx) VALUES (",
            ""+ RIGHT(@ITX+CAST(FldLvl AS varchar(10))  ,07) AS FldLvl,
            LEFT(",@InpObj"       +@ITX                 ,08) AS FldObj,
            LEFT(",'" +FldNam+ "'"+@ITX                 ,17) AS FldNam,
            LEFT(","""+FldDtx+""""+@ITX                 ,16) AS FldDtx,
            ","+RIGHT(@ITX+CAST(FldLen AS varchar(10))  ,07) AS FldLen,
            LEFT(","+REPLACE(REPLACE(FldNul,'    NULL','@ClmNulALN'),'NOT NULL','@ClmNulNNL')+@ITX,08) AS FldNul,
            LEFT(","+CASE WHEN FldNul = ' IDENTITY' THEN '@ClmIdtYID' ELSE """""" END     +@ITX,08) AS FldIdn,
            LEFT(","+CASE WHEN FldNul = ' OUTPUT'   THEN '@ClmIdtYID' ELSE """""" END     +@ITX,08) AS FldOup,
            ","+RIGHT(@ITX+CAST(FldOrd AS varchar(10))  ,07) AS FldOrd,
            LEFT(","""+FldQot+""""+@ITX                 ,07) AS FldQot,
            LEFT(","""+FldVal+""""+@ITX                 ,10) AS FldVal,
            LEFT(","""+FldDfv+""""+@ITX                 ,10) AS FldDfv,
            ","+RIGHT(@ITX+CAST(FldPko AS varchar(10))  ,07) AS FldPko,
            ","+RIGHT(@ITX+CAST(FldLko AS varchar(10))  ,07) AS FldLko,
            ","+RIGHT(@ITX+CAST(FldFko AS varchar(10))  ,07) AS FldFko,
            LEFT(","""+FldVfx+""""+@ITX                 ,07) AS FldVfx,
            ")"
        FROM
            #FldList
        ORDER BY
            FldLvl,
            FldOrd
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#ColList (CHK)' = '#ColList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #ColList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #ColList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PkfList (CHK)' = '#PkfList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #PkfList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #PkfList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#LkpList (CHK)' = '#LkpList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #LkpList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #LkpList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#SigList (CHK)' = '#SigList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #SigList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #SigList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#StxList (CHK)' = '#StxList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #StxList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #StxList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#InsList (CHK)' = '#InsList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #InsList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #InsList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#UpdList (CHK)' = '#UpdList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #UpdList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #UpdList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#HstList (CHK)' = '#HstList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #HstList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #HstList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#AudList (CHK)' = '#AudList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #AudList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #AudList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#PopList (CHK)' = '#PopList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #PopList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #PopList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#SetList (CHK)' = '#SetList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #SetList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #SetList fld
        ORDER BY
            FldLvl,
            FldOrd
    END
    ------------------------------------------------------------------------------------------------
    IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
    ------------------------------------------------------------------------------------------------
        SELECT '#CrtList (CHK)' = '#CrtList (CHK)', RowCnt = CONVERT(SMALLINT,(SELECT COUNT(*) FROM #CrtList))
            ,FldID  = fld.FldID
            ,FldLvl = fld.FldLvl
            ,FldObj = LEFT(fld.FldObj,40)
            ,FldOrd = fld.FldOrd
            ,FldNam = LEFT(fld.FldNam,40)
            ,FldUtp = fld.FldUtp
            ,FldDtx = fld.FldDtx
            ,FldLen = fld.FldLen
            ,FldDct = fld.FldDct
            ,FldQot = fld.FldQot
            ,FldNul = fld.FldNul
            ,FldIdn = fld.FldIdn
            ,FldOup = fld.FldOup
            ,FldPko = fld.FldPko
            ,FldFko = fld.FldFko
            ,FldLko = fld.FldLko
            ,FldAud = fld.FldAud
            ,FldVal = LEFT(fld.FldVal,10)
            ,FldDfv = LEFT(fld.FldDfv,10)
            ,FldVfx = LEFT(fld.FldVfx,10)
        FROM
            #CrtList fld
        ORDER BY
            FldLvl,
            FldOrd
    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --##############################################################################################


    --SQJ@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@SQJ


    --##############################################################################################


    -- Statement construction
    DECLARE @StmVar    sysname      ; SET @StmVar    = ""
    DECLARE @StmCtl    sysname      ; SET @StmCtl    = ""
    DECLARE @StmSqt    varchar(1)   ; SET @StmSqt    = ""


    --##############################################################################################


    DECLARE @LstVba    varchar(100) ; SET @LstVba    = ""
    DECLARE @LstVbx    varchar(100) ; SET @LstVbx    = ""
    DECLARE @LstVbj    varchar(100) ; SET @LstVbj    = ""


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Initialize Module working variables
    ------------------------------------------------------------------------------------------------
    DECLARE @ModNam    varchar(100) ; SET @ModNam    = ""
    DECLARE @ModTtl    varchar(100) ; SET @ModTtl    = ""
    DECLARE @ModAls    varchar(10)  ; SET @ModAls    = ""


    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Join VBA values to field values
    ------------------------------------------------------------------------------------------------
    DEALLOCATE cur_ColList; DECLARE cur_ColList CURSOR LOCAL FOR SELECT * FROM #ColList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_PkfList; DECLARE cur_PkfList CURSOR LOCAL FOR SELECT * FROM #PkfList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_LkpList; DECLARE cur_LkpList CURSOR LOCAL FOR SELECT * FROM #LkpList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_RtnList; DECLARE cur_RtnList CURSOR LOCAL FOR SELECT * FROM #RtnList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_SigList; DECLARE cur_SigList CURSOR LOCAL FOR SELECT * FROM #SigList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_StmList; DECLARE cur_StmList CURSOR LOCAL FOR SELECT * FROM #StxList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_InsList; DECLARE cur_InsList CURSOR LOCAL FOR SELECT * FROM #InsList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_UpdList; DECLARE cur_UpdList CURSOR LOCAL FOR SELECT * FROM #UpdList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_HstList; DECLARE cur_HstList CURSOR LOCAL FOR SELECT * FROM #HstList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_AudList; DECLARE cur_AudList CURSOR LOCAL FOR SELECT * FROM #DflList lst INNER JOIN #VbaDfn vba ON FldID = VbaID WHERE FldLvl IN (@FlvCRT,@FlvUPD) ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_PopList; DECLARE cur_PopList CURSOR LOCAL FOR SELECT * FROM #PopList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_MfkList; DECLARE cur_MfkList CURSOR LOCAL FOR SELECT * FROM #MfkList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_SetList; DECLARE cur_SetList CURSOR LOCAL FOR SELECT * FROM #SetList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd
    DEALLOCATE cur_CrtList; DECLARE cur_CrtList CURSOR LOCAL FOR SELECT * FROM #CrtList lst INNER JOIN #VbaDfn vba ON FldID = VbaID                                   ORDER BY FldLvl,FldOrd


    --##############################################################################################
    -- Output sections setup
    --##############################################################################################

 
    ------------------------------------------------------------------------------------------------
    -- Initialize package options
    ------------------------------------------------------------------------------------------------
    DECLARE @OupSep    varchar(2)   ; SET @OupSep    = ","
    DECLARE @CusSep    varchar(2)   ; SET @CusSep    = ":"
    DECLARE @CusTyp    varchar(10)  ; SET @CusTyp    = ""
    ------------------------------------------------------------------------------------------------
 

    ------------------------------------------------------------------------------------------------
    -- Initialize standard section constants
    ------------------------------------------------------------------------------------------------
    DECLARE @SecALL    varchar(11)  ; SET @SecALL    = 'ALL'     -- Include all code sections
    DECLARE @SecXXX    varchar(11)  ; SET @SecXXX    = 'XXX'     -- New code section
    DECLARE @SecZZZ    varchar(11)  ; SET @SecZZZ    = 'ZZZ'     -- Invalid code error message
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- Initialize output section constants
    ------------------------------------------------------------------------------------------------
    DECLARE @SecPSH    varchar(11)  ; SET @SecPSH    = 'PSH'               -- Push margin 4 spaces right
    DECLARE @SecPUL    varchar(11)  ; SET @SecPUL    = 'PUL'               -- Pull margin 4 spaces left
    DECLARE @SecLM0    varchar(11)  ; SET @SecLM0    = 'LM0'               -- Set left margin to zero
    DECLARE @SecLM1    varchar(11)  ; SET @SecLM1    = 'LM1'               -- Set left margin to one
    DECLARE @SecLM2    varchar(11)  ; SET @SecLM2    = 'LM2'               -- Set left margin to two
    DECLARE @SecRWP    varchar(11)  ; SET @SecRWP    = 'RWP'               -- Set report width Portrait
    DECLARE @SecRWL    varchar(11)  ; SET @SecRWL    = 'RWL'               -- Set report width Landscape

    DECLARE @SecLSG    varchar(11)  ; SET @SecLSG    = 'LSG'               -- Set lines for single lines
    DECLARE @SecLDB    varchar(11)  ; SET @SecLDB    = 'LDB'               -- Set lines for double lines
    DECLARE @SecLPD    varchar(11)  ; SET @SecLPD    = 'LPD'               -- Set lines for pound  lines
    DECLARE @SecHSG    varchar(11)  ; SET @SecHSG    = 'HSG'               -- Set header for single lines
    DECLARE @SecHDB    varchar(11)  ; SET @SecHDB    = 'HDB'               -- Set header for double lines
    DECLARE @SecHPD    varchar(11)  ; SET @SecHPD    = 'HPD'               -- Set header for pound  lines

    DECLARE @SecSLN    varchar(11)  ; SET @SecSLN    = 'SLN'               -- Print single line
    DECLARE @SecDLN    varchar(11)  ; SET @SecDLN    = 'DLN'               -- Print double line
    DECLARE @SecALN    varchar(11)  ; SET @SecALN    = 'ALN'               -- Print asterick line
    DECLARE @SecPLN    varchar(11)  ; SET @SecPLN    = 'PLN'               -- Print pound line
    DECLARE @SecMLN    varchar(11)  ; SET @SecMLN    = 'MLN'               -- Print ampersand line
    DECLARE @SecTLN    varchar(11)  ; SET @SecTLN    = 'TLN'               -- Print tilde line

    DECLARE @SecLN0    varchar(11)  ; SET @SecLN0    = 'LN0'               -- Print current LinSpc
    DECLARE @SecLN1    varchar(11)  ; SET @SecLN1    = 'LN1'               -- Print empty lines (1)
    DECLARE @SecLN2    varchar(11)  ; SET @SecLN2    = 'LN2'               -- Print empty lines (2)
    DECLARE @SecPLP    varchar(11)  ; SET @SecPLP    = 'PLP'               -- Pound line (prefixed)
    DECLARE @SecALP    varchar(11)  ; SET @SecALP    = 'ALP'               -- AtSign line (prefixed)
    DECLARE @SecPHB    varchar(11)  ; SET @SecPHB    = 'PHB'               -- Print begin header (previously set)
    DECLARE @SecPHE    varchar(11)  ; SET @SecPHE    = 'PHE'               -- Print end header (previously set)

    DECLARE @SecT0N    varchar(11)  ; SET @SecT0N    = 'T0N'               -- Set Title: Space 0 Lines N
    DECLARE @SecT0Y    varchar(11)  ; SET @SecT0Y    = 'T0Y'               -- Set Title: Space 0 Lines Y
    DECLARE @SecT1N    varchar(11)  ; SET @SecT1N    = 'T1N'               -- Set Title: Space 1 Lines N
    DECLARE @SecT1Y    varchar(11)  ; SET @SecT1Y    = 'T1Y'               -- Set Title: Space 1 Lines Y
    DECLARE @SecT2N    varchar(11)  ; SET @SecT2N    = 'T2N'               -- Set Title: Space 2 Lines N
    DECLARE @SecT2Y    varchar(11)  ; SET @SecT2Y    = 'T2Y'               -- Set Title: Space 2 Lines Y

    DECLARE @SecIST    varchar(11)  ; SET @SecIST    = 'IST'               -- Toggle   Include space(s) before the header
    DECLARE @SecISR    varchar(11)  ; SET @SecISR    = 'ISR'               -- Reset    Include space(s) before the header
    DECLARE @SecIS0    varchar(11)  ; SET @SecIS0    = 'IS0'               -- Set OFF  Include space(s) before the header
    DECLARE @SecIS1    varchar(11)  ; SET @SecIS1    = 'IS1'               -- Set Opt1 Include space(s) before the header
    DECLARE @SecIS2    varchar(11)  ; SET @SecIS2    = 'IS2'               -- Set Opt2 Include space(s) before the header

    DECLARE @SecITT    varchar(11)  ; SET @SecITT    = 'ITT'               -- Toggle   Include code segment titles
    DECLARE @SecITR    varchar(11)  ; SET @SecITR    = 'ITR'               -- Reset    Include code segment titles
    DECLARE @SecIT0    varchar(11)  ; SET @SecIT0    = 'IT0'               -- Set OFF  Include code segment titles
    DECLARE @SecIT1    varchar(11)  ; SET @SecIT1    = 'IT1'               -- Set Opt1 Include code segment titles
    DECLARE @SecIT2    varchar(11)  ; SET @SecIT2    = 'IT2'               -- Set Opt2 Include code segment titles
    DECLARE @SecIT3    varchar(11)  ; SET @SecIT3    = 'IT3'               -- Set Opt3 Include code segment titles

    DECLARE @SecIPT    varchar(11)  ; SET @SecIPT    = 'IPT'               -- Toggle   Include templates
    DECLARE @SecIPR    varchar(11)  ; SET @SecIPR    = 'IPR'               -- Reset    Include templates
    DECLARE @SecIP0    varchar(11)  ; SET @SecIP0    = 'IP0'               -- Set OFF  Include templates
    DECLARE @SecIP1    varchar(11)  ; SET @SecIP1    = 'IP1'               -- Set Opt1 Include templates
    DECLARE @SecIP2    varchar(11)  ; SET @SecIP2    = 'IP2'               -- Set Opt2 Include templates

    DECLARE @SecIGT    varchar(11)  ; SET @SecIGT    = 'IGT'               -- Toggle   Include debug logic
    DECLARE @SecIGR    varchar(11)  ; SET @SecIGR    = 'IGR'               -- Reset    Include debug logic
    DECLARE @SecIG0    varchar(11)  ; SET @SecIG0    = 'IG0'               -- Disable  Include debug logic
    DECLARE @SecIG1    varchar(11)  ; SET @SecIG1    = 'IG1'               -- Enable   Include debug logic

    DECLARE @SecIFT    varchar(11)  ; SET @SecIFT    = 'IFT'               -- Toggle   Include information message
    DECLARE @SecIFR    varchar(11)  ; SET @SecIFR    = 'IFR'               -- Reset    Include information message
    DECLARE @SecIF0    varchar(11)  ; SET @SecIF0    = 'IF0'               -- Disable  Include information message
    DECLARE @SecIF1    varchar(11)  ; SET @SecIF1    = 'IF1'               -- Enable   Include information message

    DECLARE @SecIQT    varchar(11)  ; SET @SecIQT    = 'IQT'               -- Toggle   Include error message
    DECLARE @SecIQR    varchar(11)  ; SET @SecIQR    = 'IQR'               -- Reset    Include error message
    DECLARE @SecIQ0    varchar(11)  ; SET @SecIQ0    = 'IQ0'               -- Disable  Include error message
    DECLARE @SecIQ1    varchar(11)  ; SET @SecIQ1    = 'IQ1'               -- Enable   Include error message

    DECLARE @SecIVT    varchar(11)  ; SET @SecIVT    = 'IVT'               -- Toggle   Include separator line between objects
    DECLARE @SecIVR    varchar(11)  ; SET @SecIVR    = 'IVR'               -- Reset    Include separator line between objects
    DECLARE @SecIV0    varchar(11)  ; SET @SecIV0    = 'IV0'               -- Disable  Include separator line between objects
    DECLARE @SecIV1    varchar(11)  ; SET @SecIV1    = 'IV1'               -- Enable   Include separator line between objects

    DECLARE @SecIDT    varchar(11)  ; SET @SecIDT    = 'IDT'               -- Toggle   Include drop statement
    DECLARE @SecIDR    varchar(11)  ; SET @SecIDR    = 'IDR'               -- Reset    Include drop statement
    DECLARE @SecID0    varchar(11)  ; SET @SecID0    = 'ID0'               -- Disable  Include drop statement
    DECLARE @SecID1    varchar(11)  ; SET @SecID1    = 'ID1'               -- Enable   Include drop statement

    DECLARE @SecIBT    varchar(11)  ; SET @SecIBT    = 'IBT'               -- Toggle   Include batch GO statement
    DECLARE @SecIBR    varchar(11)  ; SET @SecIBR    = 'IBR'               -- Reset    Include batch GO statement
    DECLARE @SecIB0    varchar(11)  ; SET @SecIB0    = 'IB0'               -- Disable  Include batch GO statement
    DECLARE @SecIB1    varchar(11)  ; SET @SecIB1    = 'IB1'               -- Enable   Include batch GO statement

    DECLARE @SecINT    varchar(11)  ; SET @SecINT    = 'INT'               -- Toggle   Include batch GO statement
    DECLARE @SecINR    varchar(11)  ; SET @SecINR    = 'INR'               -- Reset    Include batch GO statement
    DECLARE @SecIN0    varchar(11)  ; SET @SecIN0    = 'IN0'               -- Disable  Include batch GO statement
    DECLARE @SecIN1    varchar(11)  ; SET @SecIN1    = 'IN1'               -- Enable   Include batch GO statement

    DECLARE @SecIMT    varchar(11)  ; SET @SecIMT    = 'IMT'               -- Toggle   Include permissions statements
    DECLARE @SecIMR    varchar(11)  ; SET @SecIMR    = 'IMR'               -- Reset    Include permissions statements
    DECLARE @SecIM0    varchar(11)  ; SET @SecIM0    = 'IM0'               -- Disable  Include permissions statements
    DECLARE @SecIM1    varchar(11)  ; SET @SecIM1    = 'IM1'               -- Enable   Include permissions statements

    DECLARE @SecIIT    varchar(11)  ; SET @SecIIT    = 'IIT'               -- Toggle   Include identity column logic
    DECLARE @SecIIR    varchar(11)  ; SET @SecIIR    = 'IIR'               -- Reset    Include identity column logic
    DECLARE @SecII0    varchar(11)  ; SET @SecII0    = 'II0'               -- Disable  Include identity column logic
    DECLARE @SecII1    varchar(11)  ; SET @SecII1    = 'II1'               -- Enable   Include identity column logic

    DECLARE @SecIXT    varchar(11)  ; SET @SecIXT    = 'IXT'               -- Toggle   Include record expires columns
    DECLARE @SecIXR    varchar(11)  ; SET @SecIXR    = 'IXR'               -- Reset    Include record expires columns
    DECLARE @SecIX0    varchar(11)  ; SET @SecIX0    = 'IX0'               -- Disable  Include record expires columns
    DECLARE @SecIX1    varchar(11)  ; SET @SecIX1    = 'IX1'               -- Enable   Include record expires columns

    DECLARE @SecOTX    varchar(11)  ; SET @SecOTX    = 'OTX'               -- Object text (SysComments)

    DECLARE @SecDVA    varchar(11)  ; SET @SecDVA    = 'DVA'               -- Developer action history
    DECLARE @SecVLN    varchar(11)  ; SET @SecVLN    = 'VLN'               -- Set VbaVln length
    DECLARE @SecTMB    varchar(11)  ; SET @SecTMB    = 'TMB'               -- Initialize text management objects (Begin)
    DECLARE @SecTME    varchar(11)  ; SET @SecTME    = 'TME'               -- Initialize text management objects (End)

    DECLARE @SecDTV    varchar(11)  ; SET @SecDTV    = 'DTV'               -- Declare field column variables
    DECLARE @SecITV    varchar(11)  ; SET @SecITV    = 'ITV'               -- Initialize field column variables
    DECLARE @SecATV    varchar(11)  ; SET @SecATV    = 'ATV'               -- Assign standard field variables

    DECLARE @SecDSV    varchar(11)  ; SET @SecDSV    = 'DSV'               -- Declare statement column variables
    DECLARE @SecISV    varchar(11)  ; SET @SecISV    = 'ISV'               -- Initialize statement column variables
    DECLARE @SecASV    varchar(11)  ; SET @SecASV    = 'ASV'               -- Assign statement column variables

    DECLARE @SecDKV    varchar(11)  ; SET @SecDKV    = 'DKV'               -- Declare primary key column variables
    DECLARE @SecIKV    varchar(11)  ; SET @SecIKV    = 'IKV'               -- Initialize primary key column variables
    DECLARE @SecAKV    varchar(11)  ; SET @SecAKV    = 'AKV'               -- Assign primary key column variables

    DECLARE @SecDGV    varchar(11)  ; SET @SecDGV    = 'DGV'               -- Declare function parameter variables
    DECLARE @SecIGV    varchar(11)  ; SET @SecIGV    = 'IGV'               -- Initialize parameter variables
    DECLARE @SecAGV    varchar(11)  ; SET @SecAGV    = 'AGV'               -- Assign parameter variables

    DECLARE @SecIRV    varchar(11)  ; SET @SecIRV    = 'IRV'               -- Initialize recordset variables

    DECLARE @SecMLV    varchar(11)  ; SET @SecMLV    = 'MLV'               -- Module Level Variables
    DECLARE @SecMLP    varchar(11)  ; SET @SecMLP    = 'MLP'               -- Module Level Properties - LET/GET
    DECLARE @SecMLL    varchar(11)  ; SET @SecMLL    = 'MLL'               -- Module Level Properties - LET
    DECLARE @SecMLG    varchar(11)  ; SET @SecMLG    = 'MLG'               -- Module Level Properties - GET
    DECLARE @SecMLC    varchar(11)  ; SET @SecMLC    = 'MLC'               -- Module Level Properties - CLEAR

    DECLARE @SecMAV    varchar(11)  ; SET @SecMAV    = 'MAV'               -- Module Level AssignSQL (from variables)
    DECLARE @SecMAC    varchar(11)  ; SET @SecMAC    = 'MAC'               -- Module Level AssignSQL (from controls)
    DECLARE @SecMAN    varchar(11)  ; SET @SecMAN    = 'MAN'               -- Module Level AssignSQL (from nulls)

    DECLARE @SecMRV    varchar(11)  ; SET @SecMRV    = 'MRV'               -- Module Level ReadSQL (into variables)
    DECLARE @SecMRC    varchar(11)  ; SET @SecMRC    = 'MRC'               -- Module Level ReadSQL (into controls)

    DECLARE @SecMCV    varchar(11)  ; SET @SecMCV    = 'MCV'               -- Module Level ClearSQL (variables)
    DECLARE @SecMCC    varchar(11)  ; SET @SecMCC    = 'MCC'               -- Module Level ClearSQL (controls)

    DECLARE @SecMAD    varchar(11)  ; SET @SecMAD    = 'MAD'               -- Module IF Criteria
    DECLARE @SecMUP    varchar(11)  ; SET @SecMUP    = 'MUP'               -- Module IF Criteria
    DECLARE @SecMIF    varchar(11)  ; SET @SecMIF    = 'MIF'               -- Module IF Criteria
    DECLARE @SecMTP    varchar(11)  ; SET @SecMTP    = 'MTP'               -- Module IF Criteria
    DECLARE @SecMTF    varchar(11)  ; SET @SecMTF    = 'MTF'               -- Module IF Criteria

    DECLARE @SecRUNUSX varchar(11)  ; SET @SecRUNUSX = 'RUNUSX'            -- Run SProc commands

    DECLARE @SecSQLSTM varchar(11)  ; SET @SecSQLSTM = 'SQLSTM'            -- Build the full SQL statement

    DECLARE @SecFRMHDR varchar(11)  ; SET @SecFRMHDR = 'FRMHDR'            -- Header form
    DECLARE @SecFRMDTL varchar(11)  ; SET @SecFRMDTL = 'FRMDTL'            -- Detail form
    DECLARE @SecFRMLST varchar(11)  ; SET @SecFRMLST = 'FRMLST'            -- List form

    DECLARE @SecPOPADD varchar(11)  ; SET @SecPOPADD = 'POPADD'            -- PopAdd form
    DECLARE @SecPOPUPD varchar(11)  ; SET @SecPOPUPD = 'POPUPD'            -- PopUpd form

    DECLARE @SecRECINS varchar(11)  ; SET @SecRECINS = 'RECINS'            -- Insert record function
    DECLARE @SecRECUPD varchar(11)  ; SET @SecRECUPD = 'RECUPD'            -- Update record function
    ------------------------------------------------------------------------------------------------

 
    ------------------------------------------------------------------------------------------------
    -- Create comment test code packages
    ------------------------------------------------------------------------------------------------
    --SELECT @OupTyp AS OupCat,@OupObj AS OupObj,@BldLST AS OupLst
    DECLARE @PkxCMT    varchar(100)
    IF @OupTyp IN (@ObjTypTBL,@ObjTypDDL) BEGIN
        SET @PkxCMT = 'CMT,CFT,'+CASE 
            WHEN @OupObj LIKE "usp_Lookup_%" THEN 'LN1,PSH,SVS,POP,'
            WHEN @BldLST IN   ("USPLKP")     THEN 'LN1,PSH,SVS,POP,'
            ELSE ""
        END
    END ELSE IF @OupTyp IN (@ObjTypUSP) BEGIN
        SET @PkxCMT = 'CMT,CFP,'+CASE 
            WHEN @OupObj LIKE "usp_Lookup_%" THEN 'LN1,PSH,SVG,POP,'
            WHEN @BldLST IN   ("USPLKP")     THEN 'LN1,PSH,SVG,POP,'
            ELSE ""
        END
    END ELSE IF @OupTyp IN (@ObjTypVEW) BEGIN
        SET @PkxCMT = 'CMT,CFV,'
    END ELSE BEGIN
        SET @PkxCMT = 'CMT,'
    END
    ------------------------------------------------------------------------------------------------

 
    ------------------------------------------------------------------------------------------------
    -- Create header packages
    ------------------------------------------------------------------------------------------------
    -- LEC - List empty comment
    DECLARE @PkxLEC    varchar(100) ; SET @PkxLEC    = ''
    ------------------------------------------------------------------------------------------------

 
    /*----------------------------------------------------------------------------------------------
        Build standard packages
        EXEC ut_zzVBJ vba_TblDfn,WRTRST                     -- Write text management objects

        EXEC ut_zzVBJ vba_TblDfn,RUNUSX                     -- Run SProc process
        EXEC ut_zzVBJ vba_TblDfn,SQLSTM                     -- Build the full SQL statement
        EXEC ut_zzVBJ vba_TblDfn,SQLWHR                     -- Build the SQL WHERE clause
        EXEC ut_zzVBJ vba_TblDfn,SQLOBY                     -- Build the SQL ORDER BY clause
        EXEC ut_zzVBJ vba_TblDfn,BLDSQL                     -- Build SQL Text from SelectOn controls

        EXEC ut_zzVBJ vba_TblDfn,FRMHDR                     -- Header form
        EXEC ut_zzVBJ vba_TblDfn,FRMDTL                     -- Detail form
        EXEC ut_zzVBJ vba_TblDfn,FRMLST                     -- List form
        EXEC ut_zzVBJ vba_TblDfn,POPADD                     -- PopAdd form
        EXEC ut_zzVBJ vba_TblDfn,POPUPD                     -- PopUpd form
        EXEC ut_zzVBJ vba_TblDfn,RECINS                     -- Insert record function
        EXEC ut_zzVBJ vba_TblDfn,RECUPD                     -- Update record function
    ----------------------------------------------------------------------------------------------*/
    SET @BldLST = CASE @BldLST
        WHEN 'LUPRST' THEN @PkxLEC+'IT1,VLN=20,DSV,TMB,PSH,IT0,IRV,PUL,TME,'
        WHEN 'WRTRST' THEN @PkxLEC+'TMB,PSH,IT0,IRV,PUL,TME,'
        --------------------------------------------------------------------------------------------
        WHEN 'RUNUSX' THEN @PkxLEC+'RUNUSX,'               -- Run SProc process
        WHEN 'SQLSTM' THEN @PkxLEC+'SQLSTM,'               -- Build the full SQL statement
        WHEN 'SQLWHR' THEN @PkxLEC+'SQLWHR,'               -- Build the SQL WHERE clause
        WHEN 'SQLOBY' THEN @PkxLEC+'SQLOBY,'               -- Build the SQL ORDER BY clause
        WHEN 'BLDSQL' THEN @PkxLEC+'BLDSQL,'               -- Build SQL Text from SelectOn controls
        --------------------------------------------------------------------------------------------
        WHEN 'FRMHDR' THEN @PkxLEC+'FRMHDR,'               -- Header form
        WHEN 'FRMDTL' THEN @PkxLEC+'FRMDTL,'               -- Detail form
        WHEN 'FRMLST' THEN @PkxLEC+'FRMLST,'               -- List form
        WHEN 'POPADD' THEN @PkxLEC+'POPADD,'               -- PopAdd form
        WHEN 'POPUPD' THEN @PkxLEC+'POPUPD,'               -- PopUpd form
        WHEN 'RECINS' THEN @PkxLEC+'RECINS,'               -- Insert record function
        WHEN 'RECUPD' THEN @PkxLEC+'RECUPD,'               -- Update record function
        --------------------------------------------------------------------------------------------
        ELSE @BldLST
    END
    ------------------------------------------------------------------------------------------------

 
    --##############################################################################################
    WHILE LEFT (@BldLST,LEN(@OupSep)) = @OupSep SET @BldLST = RIGHT(@BldLST,LEN(@BldLST)-LEN(@OupSep))
    WHILE RIGHT(@BldLST,LEN(@OupSep)) = @OupSep SET @BldLST = LEFT (@BldLST,LEN(@BldLST)-LEN(@OupSep))
    WHILE @BldLST LIKE "%"+@OupSep+@OupSep+"%"  SET @BldLST = REPLACE(@BldLST,@OupSep+@OupSep,@OupSep)
    ------------------------------------------------------------------------------------------------
    IF @DbgFlg = 1 OR 0=9 SELECT LUP='LUP',@BldSfx AS BldSfx,@BldCOD AS BldCOD,@BldLST AS BldLST
    ------------------------------------------------------------------------------------------------
    DECLARE @BldSqn varchar(2000); SET @BldSqn = @BldLST
    ------------------------------------------------------------------------------------------------
    WHILE LEN(@BldLST) > 0 BEGIN
    ------------------------------------------------------------------------------------------------
        SET @PrvBld = @BldCOD
        SET @POS = CHARINDEX(@OupSep,@BldLST)
        IF @POS > 0 BEGIN
            SET @BldCOD = LTRIM(RTRIM(LEFT(@BldLST,@POS-1))); SET @BldLST = LTRIM(RIGHT(@BldLST,LEN(@BldLST)-@POS-(LEN(@OupSep)-1)))
        END ELSE BEGIN
            SET @BldCOD = LTRIM(RTRIM(@BldLST)); SET @BldLST = ''
        END
        SET @POS = CHARINDEX(@CusSep,@BldCOD)
        IF @POS > 0 BEGIN
            SET @CusTyp = LTRIM(RTRIM(RIGHT(@BldCOD,LEN(@BldCOD)-@POS)))
            SET @BldCOD = LTRIM(RTRIM(LEFT(@BldCOD, @POS - 1)))
        END ELSE BEGIN
            SET @CusTyp = ''
        END
       -------------------------------------------------------------------------------------------------
       -- Extract output value
       -------------------------------------------------------------------------------------------------
        SET @BldVAL = ''; IF CHARINDEX('=',@BldCOD) > 0 BEGIN
            SET @POS    = CHARINDEX('=',@BldCOD)
            SET @BldVAL = SUBSTRING(@BldCOD,@POS+1,999)
            SET @BldCOD = LEFT(@BldCOD,@POS-1)
        END
       -------------------------------------------------------------------------------------------------
        -- Assign utility signatures
       -------------------------------------------------------------------------------------------------
        SET @LstVbj = '  -  ut_zzVBJ '+@DspTbl+','+@BldCOD
       -------------------------------------------------------------------------------------------------
        -- Display output codes/values
       -------------------------------------------------------------------------------------------------
        IF (@DbgFlg = 1 OR 0=9) AND 9=9 BEGIN
            SELECT @BldCOD AS BldCOD,@BldSfx AS BldSfx,@BldVAL AS BldVAL,@LinCnt AS LinCnt,@IncSpc AS IncSpc,@IncTtl AS IncTtl,@IncCmt AS IncCmt,@IncHdr AS IncHdr,@IncTpl AS IncTpl,@IncDbg AS IncDbg,@IncMsg AS IncMsg,@IncErm AS IncErm,@IncSep AS IncSep,@IncDrp AS IncDrp,@IncAdd AS IncAdd,@IncBat AS IncBat,@IncTcd AS IncTcd,@IncDat AS IncDat,@IncPrm AS IncPrm,@IncTrn AS IncTrn,@IncIdn AS IncIdn,@IncDsb AS IncDsb,@IncDlt AS IncDlt,@IncLok AS IncLok,@IncAud AS IncAud,@IncHst AS IncHst,@IncMod AS IncMod
        END
    --##############################################################################################


    ------------------------------------------------------------------------------------------------
    -- Output code based values
    ------------------------------------------------------------------------------------------------
    -- Utility signature
    SET @LstVba = "  -  ut_zzVBA "+@DspTbl+","+@BldCOD
    SET @LstVbx = "  -  ut_zzVBX "+@DspTbl+","+@BldCOD
    SET @LstVbj = "  -  ut_zzVBJ "+@DspTbl+","+@BldCOD
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- PSH = Push margin 4 spaces right
    -- PUL = Pull margin 4 spaces left
    -- LM0 = Set left margin to zero
    -- LM1 = Set left margin to one
    -- LM2 = Set left margin to two
    -- RWP = Set report width = Portrait
    -- RWL = Set report width = Landscape
    ------------------------------------------------------------------------------------------------
    IF @BldCOD IN (@SecPSH,@SecPUL,@SecLM0,@SecLM1,@SecLM2,@SecRWP,@SecRWL) BEGIN  -- (OPM)
        SET @RptWid = CASE @BldCOD
            WHEN @SecRWP THEN @811WidPOR
            WHEN @SecRWL THEN @811WidLND
            ELSE              @RptWid
        END
        SET @WidSLT = CASE @RptWid WHEN @811WidLND THEN "RWL,SLT" ELSE "SLT" END
        SET @WidDLT = CASE @RptWid WHEN @811WidLND THEN "RWL,DLT" ELSE "SDT" END
        SET @WidMn0 = @RptWid - 0
        SET @WidMn1 = @RptWid - 1
        SET @WidMn2 = @RptWid - 2
        SET @WidMn4 = @RptWid - 4
        SET @LftMrg = CASE
            WHEN @BldCOD IN (@SecPSH) THEN @LftMrg+1
            WHEN @BldCOD IN (@SecPUL) THEN @LftMrg - 1
            WHEN @BldCOD IN (@SecLM1) THEN 1
            WHEN @BldCOD IN (@SecLM2) THEN 2
            ELSE 0
        END
        IF @LftMrg < 0 SET @LftMrg = 0
        SET @StmMrg = @LftMrg+1
        SET @LEN    = @LftMrg * @MrgInc
        SET @M      = REPLICATE(" ", @LEN)
        SET @T      = REPLICATE(" ", @LEN+@MrgInc)
        SET @LinSgl = "'"+REPLICATE("-", @WidMn1 - @LEN)
        SET @LinDbl = "'"+REPLICATE("=", @WidMn1 - @LEN)
        SET @LinAst = "'"+REPLICATE("*", @WidMn1 - @LEN)
        SET @LinPnd = "'"+REPLICATE("#", @WidMn1 - @LEN)
        SET @LinAts = "'"+REPLICATE("@", @WidMn1 - @LEN)
        SET @LinTld = "'"+REPLICATE("~", @WidMn1 - @LEN)
        SET @LinBng = "'"+REPLICATE("!", @WidMn1 - @LEN)
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- LSG = Set lines for single lines
    -- LDB = Set lines for double lines
    -- LPD = Set lines for pound  lines
    -- HSG = Set header for single lines
    -- HDB = Set header for double lines
    -- HPD = Set header for pound  lines
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecLSG,@SecLDB,@SecLPD,@SecHSG,@SecHDB,@SecHPD) BEGIN
        SET @TX1 = CASE @BldCOD
            WHEN @SecLDB THEN @LinDbl
            WHEN @SecLPD THEN @LinPnd
            WHEN @SecHDB THEN @LinDbl
            WHEN @SecHPD THEN @LinDbl
            ELSE @LinSgl
        END
        SET @TX2 = CASE @BldCOD
            WHEN @SecLSG THEN @LinCmt
            WHEN @SecHSG THEN @LinCmt
            ELSE ""
        END
        SET @TX3 = CASE @BldCOD
            WHEN @SecHDB THEN @LinDbl
            WHEN @SecHPD THEN @LinPnd
            ELSE @TX1
        END
        SET @HdrBeg = @TX1
        SET @HdrCmt = @TX2
        SET @HdrEnd = @TX3
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- SLN = Print single line
    -- DLN = Print double line
    -- ALN = Print asterick line
    -- PLN = Print pound line
    -- MLN = Print ampersand line
    -- TLN = Print tilde line
    /*----------------------------------------------------------------------------------------------
        --   ut_zzSQX Oup Stx Lft Spc
        EXEC ut_zzSQX SLN,'' ,1  ,0  
        EXEC ut_zzSQX DLN,'' ,1  ,0  
        EXEC ut_zzSQX ALN,'' ,1  ,0  
        EXEC ut_zzSQX PLN,'' ,1  ,0  
        EXEC ut_zzSQX MLN,'' ,1  ,0  
        EXEC ut_zzSQX TLN,'' ,1  ,0  
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecSLN,@SecDLN,@SecALN,@SecPLN,@SecMLN,@SecTLN) BEGIN
        SET @TX1 = CASE @BldCOD
            WHEN @SecSLN THEN @LinSgl
            WHEN @SecDLN THEN @LinDbl
            WHEN @SecALN THEN @LinAst
            WHEN @SecPLN THEN @LinPnd
            WHEN @SecMLN THEN @LinAts
            WHEN @SecTLN THEN @LinTld
            ELSE              @LinSgl
        END
        IF @PrnSpc = 1 PRINT @LinSpc
        PRINT @M+@TX1
        CONTINUE
    -- --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- -- RWP = Set report width = Portrait
    -- -- RWL = Set report width = Landscape
    -- ------------------------------------------------------------------------------------------------
    -- END ELSE IF @BldCOD IN (@SecRWP,@SecRWL) BEGIN
    --     SET @RptWid = CASE @BldCOD 
    --         WHEN @SecRWP THEN @811Por
    --         WHEN @SecRWL THEN @811Lnd
    --         ELSE              @RptWid
    --     END
    --     SET @WidMn0 = @RptWid - 0
    --     SET @WidMn1 = @RptWid - 1
    --     SET @WidMn2 = @RptWid - 2
    --     SET @WidMn4 = @RptWid - 4
    --     CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- LN0 = Print current @LinSpc
    -- LN1 = Print empty lines (1)
    -- LN2 = Print empty lines (2)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzXXX LN0
        EXEC ut_zzXXX LN1
        EXEC ut_zzXXX LN2
    ----------------------------------------------------------------------------------------------*/
    IF @BldCOD IN (@SecLN0,@SecLN1,@SecLN2) BEGIN
        IF @BldCOD IN (@SecLN0) AND @PrnSpc = 1 PRINT @LinSpc
        IF @BldCOD IN (@SecLN1)                 PRINT @N
        IF @BldCOD IN (@SecLN2)                 PRINT @N+@N
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- PLP = Pound line (prefixed)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzXXX PLP
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecPLP) BEGIN
        IF @PrnSpc = 1 PRINT @LinSpc
        PRINT @M+LEFT(RTRIM(@LinCmt)+@LinPnd,100 - (@LftMrg * @MrgInc))
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- ALP = AtSign line (prefixed)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzXXX ALP
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecALP) BEGIN
        IF @PrnSpc = 1 PRINT @LinSpc
        PRINT @M+LEFT(RTRIM(@LinCmt)+@LinAts,100 - (@LftMrg * @MrgInc))
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- PHB = Print begin header (previously set)
    -- PHE = Print end header (previously set)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzXXX PHB
        EXEC ut_zzXXX PHE
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecPHB) BEGIN
        PRINT @HdrBeg; CONTINUE
    END ELSE IF @BldCOD IN (@SecPHE) BEGIN
        PRINT @HdrEnd; CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- T0N = Set Title: Space=0 Lines=N
    -- T0Y = Set Title: Space=0 Lines=Y
    -- T1N = Set Title: Space=1 Lines=N
    -- T1Y = Set Title: Space=1 Lines=Y
    -- T2N = Set Title: Space=2 Lines=N
    -- T2Y = Set Title: Space=2 Lines=Y
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecT0N,@SecT0Y,@SecT1N,@SecT1Y,@SecT2N,@SecT2Y) BEGIN
        SET @IncSpc = CASE @BldCOD
            WHEN @SecT0N THEN 0
            WHEN @SecT0Y THEN 0
            WHEN @SecT1N THEN 1
            WHEN @SecT1Y THEN 1
            WHEN @SecT2N THEN 2
            WHEN @SecT2Y THEN 2
            ELSE @IncSpc
        END
        SET @LinSpc = ""; SET @PrnSpc = 0; SET @CNT = @IncSpc; WHILE @CNT > 0 BEGIN
            SET @LinSpc = @LinSpc+@N; SET @PrnSpc = 1; SET @CNT = @CNT - 1
        END
        SET @IncTtl = CASE @BldCOD
            WHEN @SecT0N THEN 1
            WHEN @SecT0Y THEN 2
            WHEN @SecT1N THEN 1
            WHEN @SecT1Y THEN 2
            WHEN @SecT2N THEN 1
            WHEN @SecT2Y THEN 2
            ELSE @IncTtl
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- Include list:  ux_zzUTL INC
    ------------------------------------------------------------------------------------------------
    -- A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z
    -- Add Bat Fct Drp Dlt Msg Dbg Hdr Idn     Lok Dsb Prm Dat Dim Tpl Erm Trn Spc Ttl Aud Sep Cmt Exp Hst Tcd
    ------------------------------------------------------------------------------------------------
    --   PSH/PUL/LM0/LM1/LM2     @LftMrg tinyint       = 1,              -- Increase left margin (4x)
    -- S IST/ISR/IS0/IS1/IS2     @IncSpc tinyint       = 2,              -- Include space(s) before the header
    -- T ITT/ITR/IT0/IT1/IT2/IT3 @IncTtl tinyint       = 1,              -- Include code segment titles
    -- W                         @IncCmt tinyint       = 1,              -- Include comment text
    -- H                         @IncHdr tinyint       = 1,              -- Include header lines/text
    -- P IPT/IPR/IP0/IP1/IP2     @IncTpl tinyint       = 0,              -- Include templates
    -- G IGT/IGR/IG0/IG1         @IncDbg tinyint       = 0,              -- Include debug logic
    -- F IFT/IFR/IF0/IF1         @IncMsg tinyint       = 0,              -- Include information message
    -- Q IQT/IQR/IQ0/IQ1         @IncErm tinyint       = 0,              -- Include error message
    -- V IVT/IVR/IV0/IV1         @IncSep tinyint       = 0,              -- Include separator line between objects
    -- D IDT/IDR/ID0/ID1         @IncDrp tinyint       = 0,              -- Include drop statement
    -- A                         @IncAdd tinyint       = 1,              -- Include add statement
    -- B IBT/IBR/IB0/IB1         @IncBat tinyint       = 1,              -- Include batch GO statement
    -- Z                         @IncTcd tinyint       = 0,              -- Include test code
    -- N INT/INR/IN0/IN1         @IncDat tinyint       = 0,              -- Include data insert statements
    -- M                         @IncPrm tinyint       = 0,              -- Include permissions statements
    -- R                         @IncTrn tinyint       = 0,              -- Include transaction logic
    -- I IIT/IIR/II0/II1         @IncIdn tinyint       = 0,              -- Include identity column logic
    -- O                         @IncDim tinyint       = NULL,           -- Include record dimension columns
    -- C                         @IncFct tinyint       = NULL,           -- Include record fact columns
    -- L                         @IncDsb tinyint       = NULL,           -- Include record disabled columns
    -- E                         @IncDlt tinyint       = NULL,           -- Include record delflag columns
    -- K                         @IncLok tinyint       = NULL,           -- Include record locking columns
    -- X IXT/IXR/IX0/IX1         @IncExp tinyint       = NULL,           -- Include record expired columns
    -- U                         @IncAud tinyint       = NULL,           -- Include record auditing columns
    -- Y                         @IncHst tinyint       = NULL,           -- Include record history columns
    ------------------------------------------------------------------------------------------------
    -- IST = Toggle   Include space(s) before the header
    -- ISR = Reset    Include space(s) before the header
    -- IS0 = Set OFF  Include space(s) before the header
    -- IS1 = Set Opt1 Include space(s) before the header
    -- IS2 = Set Opt2 Include space(s) before the header
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIST,@SecISR,@SecIS0,@SecIS1,@SecIS2) BEGIN
        SET @IncSpc = CASE 
            WHEN @BldCOD IN (@SecIST) THEN CASE WHEN @IncSpc = 0 AND @OrgSpc > 0 THEN @OrgSpc ELSE 0 END
            WHEN @BldCOD IN (@SecISR) THEN @OrgSpc
            WHEN @BldCOD IN (@SecIS0) THEN 0
            WHEN @BldCOD IN (@SecIS1) THEN 1
            WHEN @BldCOD IN (@SecIS2) THEN 2
            ELSE 0
        END
        SET @LinSpc = ""; SET @PrnSpc = 0; SET @CNT = @IncSpc; WHILE @CNT > 0 BEGIN
            SET @LinSpc = @LinSpc+@N; SET @PrnSpc = 1; SET @CNT = @CNT - 1
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- ITT = Toggle   Include code segment titles
    -- ITR = Reset    Include code segment titles
    -- IT0 = Set OFF  Include code segment titles
    -- IT1 = Set Opt1 Include code segment titles
    -- IT2 = Set Opt2 Include code segment titles
    -- IT3 = Set Opt3 Include code segment titles
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecITT,@SecITR,@SecIT0,@SecIT1,@SecIT2,@SecIT3) BEGIN
        SET @IncTtl = CASE 
            WHEN @BldCOD IN (@SecITT) THEN CASE WHEN @IncTtl = 0 AND @OrgTtl > 0 THEN @OrgTtl ELSE 0 END
            WHEN @BldCOD IN (@SecITR) THEN @OrgTtl
            WHEN @BldCOD IN (@SecIT0) THEN 0
            WHEN @BldCOD IN (@SecIT1) THEN 1
            WHEN @BldCOD IN (@SecIT2) THEN 2
            WHEN @BldCOD IN (@SecIT3) THEN 3
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IPT = Toggle   Include templates
    -- IPR = Reset    Include templates
    -- IP0 = Set OFF  Include templates
    -- IP1 = Set Opt1 Include templates
    -- IP2 = Set Opt2 Include templates
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIPT,@SecIPR,@SecIP0,@SecIP1,@SecIP2) BEGIN
        SET @IncTpl = CASE 
            WHEN @BldCOD IN (@SecIPT) THEN CASE WHEN @IncTpl = 0 AND @OrgTpl > 0 THEN @OrgTpl ELSE 0 END
            WHEN @BldCOD IN (@SecIPR) THEN @OrgTpl
            WHEN @BldCOD IN (@SecIP0) THEN 0
            WHEN @BldCOD IN (@SecIP1) THEN 1
            WHEN @BldCOD IN (@SecIP2) THEN 2
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IGT = Toggle   Include debug logic
    -- IGR = Reset    Include debug logic
    -- IG0 = Disable  Include debug logic
    -- IG1 = Enable   Include debug logic
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIGT,@SecIGR,@SecIG0,@SecIG1) BEGIN
        SET @IncDbg = CASE 
            WHEN @BldCOD IN (@SecIGT) THEN CASE WHEN @IncDbg = 0 AND @OrgDbg > 0 THEN @OrgDbg ELSE 0 END
            WHEN @BldCOD IN (@SecIGR) THEN @OrgDbg
            WHEN @BldCOD IN (@SecIG0) THEN 0
            WHEN @BldCOD IN (@SecIG1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IFT = Toggle   Include information message
    -- IFR = Reset    Include information message
    -- IF0 = Disable  Include information message
    -- IF1 = Enable   Include information message
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIFT,@SecIFR,@SecIF0,@SecIF1) BEGIN
        SET @IncMsg = CASE 
            WHEN @BldCOD IN (@SecIFT) THEN CASE WHEN @IncMsg = 0 AND @OrgMsg > 0 THEN @OrgMsg ELSE 0 END
            WHEN @BldCOD IN (@SecIFR) THEN @OrgMsg
            WHEN @BldCOD IN (@SecIF0) THEN 0
            WHEN @BldCOD IN (@SecIF1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IQT = Toggle   Include error message
    -- IQR = Reset    Include error message
    -- IQ0 = Disable  Include error message
    -- IQ1 = Enable   Include error message
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIQT,@SecIQR,@SecIQ0,@SecIQ1) BEGIN
        SET @IncErm = CASE 
            WHEN @BldCOD IN (@SecIQT) THEN CASE WHEN @IncErm = 0 AND @OrgErm > 0 THEN @OrgErm ELSE 0 END
            WHEN @BldCOD IN (@SecIQR) THEN @OrgErm
            WHEN @BldCOD IN (@SecIQ0) THEN 0
            WHEN @BldCOD IN (@SecIQ1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IVT = Toggle   Include separator line between objects
    -- IVR = Reset    Include separator line between objects
    -- IV0 = Disable  Include separator line between objects
    -- IV1 = Enable   Include separator line between objects
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIVT,@SecIVR,@SecIV0,@SecIV1) BEGIN
        SET @IncSep = CASE
            WHEN @BldCOD IN (@SecIVT) THEN CASE WHEN @IncSep = 0 AND @OrgSep > 0 THEN @OrgSep ELSE 0 END
            WHEN @BldCOD IN (@SecIVR) THEN @OrgSep
            WHEN @BldCOD IN (@SecIV0) THEN 0
            WHEN @BldCOD IN (@SecIV1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IDT = Toggle   Include drop statement
    -- IDR = Reset    Include drop statement
    -- ID0 = Disable  Include drop statement
    -- ID1 = Enable   Include drop statement
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIDT,@SecIDR,@SecID0,@SecID1) BEGIN
        SET @IncDrp = CASE
            WHEN @BldCOD IN (@SecIDT) THEN CASE WHEN @IncDrp = 0 AND @OrgDrp > 0 THEN @OrgDrp ELSE 0 END
            WHEN @BldCOD IN (@SecIDR) THEN @OrgDrp
            WHEN @BldCOD IN (@SecID0) THEN 0
            WHEN @BldCOD IN (@SecID1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IBT = Toggle   Include batch GO statement
    -- IBR = Reset    Include batch GO statement
    -- IB0 = Disable  Include batch GO statement
    -- IB1 = Enable   Include batch GO statement
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIBT,@SecIBR,@SecIB0,@SecIB1) BEGIN
        SET @IncBat = CASE 
            WHEN @BldCOD IN (@SecIBT) THEN CASE WHEN @IncBat = 0 AND @OrgBat > 0 THEN @OrgBat ELSE 0 END
            WHEN @BldCOD IN (@SecIBR) THEN @OrgBat
            WHEN @BldCOD IN (@SecIB0) THEN 0
            WHEN @BldCOD IN (@SecIB1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- INT = Toggle   Include batch GO statement
    -- INR = Reset    Include batch GO statement
    -- IN0 = Disable  Include batch GO statement
    -- IN1 = Enable   Include batch GO statement
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecINT,@SecINR,@SecIN0,@SecIN1) BEGIN
        SET @IncDat = CASE 
            WHEN @BldCOD IN (@SecINT) THEN CASE WHEN @IncDat = 0 AND @OrgDat > 0 THEN @OrgDat ELSE 0 END
            WHEN @BldCOD IN (@SecINR) THEN @OrgDat
            WHEN @BldCOD IN (@SecIN0) THEN 0
            WHEN @BldCOD IN (@SecIN1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IMT = Toggle   Include permissions statements
    -- IMR = Reset    Include permissions statements
    -- IM0 = Disable  Include permissions statements
    -- IM1 = Enable   Include permissions statements
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIMT,@SecIMR,@SecIM0,@SecIM1) BEGIN
        SET @IncPrm = CASE 
            WHEN @BldCOD IN (@SecIMT) THEN CASE WHEN @IncPrm = 0 AND @OrgPrm > 0 THEN @OrgPrm ELSE 0 END
            WHEN @BldCOD IN (@SecIMR) THEN @OrgPrm
            WHEN @BldCOD IN (@SecIM0) THEN 0
            WHEN @BldCOD IN (@SecIM1) THEN 1
            ELSE 0
        END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IIT = Toggle   Include identity column logic
    -- IIR = Reset    Include identity column logic
    -- II0 = Disable  Include identity column logic
    -- II1 = Enable   Include identity column logic
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIIT,@SecIIR,@SecII0,@SecII1) BEGIN
        SET @IncIdn = CASE 
            WHEN @BldCOD IN (@SecIIT) THEN CASE WHEN @IncIdn = 0 AND @OrgIdn > 0 THEN @OrgIdn ELSE 0 END
            WHEN @BldCOD IN (@SecIIR) THEN @OrgIdn
            WHEN @BldCOD IN (@SecII0) THEN 0
            WHEN @BldCOD IN (@SecII1) THEN 1
            ELSE 0
        END
        SET @HasIdn = CASE WHEN LEN(@IdnClm) > 0 THEN @IncIdn ELSE 0 END
        CONTINUE
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- IXT = Toggle   Include record expires columns
    -- IXR = Reset    Include record expires columns
    -- IX0 = Disable  Include record expires columns
    -- IX1 = Enable   Include record expires columns
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecIXT,@SecIXR,@SecIX0,@SecIX1) BEGIN
        SET @IncExp = CASE
            WHEN @BldCOD IN (@SecIXT) THEN CASE WHEN @IncExp = 0 AND @OrgExp > 0 THEN @OrgExp ELSE 0 END
            WHEN @BldCOD IN (@SecIXR) THEN @OrgExp
            WHEN @BldCOD IN (@SecIX0) THEN 0
            WHEN @BldCOD IN (@SecIX1) THEN 1
            ELSE 0
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- OTX = Object text (SysComments)
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup Stx     Lft Spc Ttl Bat Tx1 Tx2 Tx3 Trn Idn Erm
        EXEC ut_zzVBX OTX,@OupObj,0  ,0  ,0  ,0  ,'' ,'' ,'' ,0  ,0  ,0
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecOTX) BEGIN
        --   ut_zzVBX Oup     Stx     Lft Spc Ttl Bat Tx1 Tx2 Tx3 Trn Idn Erm
        EXEC ut_zzVBX @BldCOD,@OupObj,0  ,0  ,0  ,0  ,'' ,'' ,'' ,0  ,0  ,0
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DVA = Developer action history
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,DVA,USP,usp_Module,Module_Description
    ----------------------------------------------------------------------------------------------*/
    IF @BldCOD IN (@SecDVA) BEGIN
        SET @ModNam = CASE WHEN LEN(@OupObj) > 0 THEN       @OupObj  ELSE "pfx_ModuleName"       END
        SET @ModTtl = CASE WHEN LEN(@OupDsc) > 0 THEN       @OupDsc  ELSE "ModuleTitle"          END
        SET @ModAls = CASE WHEN LEN(@OupAls) > 0 THEN LOWER(@OupAls) ELSE LOWER(LEFT(@ModNam,3)) END
        PRINT @M+"'###################################################################################################"
        PRINT @M+"' Name:"+REPLICATE(" ",63)+"("+@OupFmt+":"+@DefTyp+")  "+CONVERT(char(19),GETDATE(),120)
        PRINT @M+"'   "+@ModNam+""
        PRINT @M+"'###################################################################################################"
        PRINT @M+"' Purpose:"
        PRINT @M+"'   "+@ModTtl
        PRINT @M+"'###################################################################################################"
        PRINT @M+"' Developer    Date     Action"
        PRINT @M+"' ------------ -------- ----------------------------------------------------------------------------"
        PRINT @M+"' "+@DvpTxt+" "+CONVERT(char(8),GETDATE(),1)+" Created the script"
        PRINT @M+"'###################################################################################################"
        PRINT @M+"'# Templates:"
        PRINT @M+"'"
        PRINT @M+"'    ' Initialize working objects"
        PRINT @M+"'    Dim "+@ModAls+"    As "+@ModNam 
        PRINT @M+"'    Set "+@ModAls+" = New "+@ModNam
        PRINT @M+"'"
        PRINT @M+"'###################################################################################################"
        PRINT @M+"Option Compare Database"
        PRINT @M+"Option Explicit"
        PRINT @M+"Option Base 0"
        PRINT @M+"'***************************************************************************************************"
        PRINT @M+"' Initialize module message constants"
        PRINT @M+"'***************************************************************************************************"
        PRINT @M+"Private Const mcModNam              As String = """+@ModNam+""""
        PRINT @M+"Private Const mcModTtl              As String = """+@ModTtl+""""
        PRINT @M+"Private Const mcModErr              As String = mcModNam"
        PRINT @M+"Private Const mcModMsg              As String = mcModTtl & "" - """
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- VLN = Set VbaVln length
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecVLN) BEGIN
        IF ISNUMERIC(@BldVAL) = 1 BEGIN
            SET @LEN = CAST(@BldVAL AS int)
            IF @LEN > @VbaVln SET @VbaVln = @LEN
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- TMB = Initialize text management objects (Begin)
    -- TME = Initialize text management objects (End)
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,TMB
        EXEC ut_zzVBJ zzz_TEST01,TME
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecTMB,@SecTME) BEGIN
        IF @BldCOD IN (@SecTMB) BEGIN
        PRINT @M+"' Initialize text management objects"+@LstVbj
        PRINT @M+"Dim wx    As clsWrtTxt"
        PRINT @M+"Set wx = New clsWrtTxt"
        PRINT @M+"Call wx.AX_Clear(0): With wx"
        PRINT @M+"'***********************************************************************************************"
        PRINT @M+"Do While Not rst.EOF"
        END ELSE BEGIN
        PRINT @M+"    rst.MoveNext"
        PRINT @M+"Loop"
        PRINT @M+"'***********************************************************************************************"
        PRINT @M+"End With: Call wx.WX_Write"
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DTV = Declare field column variables
    -- ITV = Initialize field column variables
    -- ATV = Assign standard field variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,DTV
        EXEC ut_zzVBJ zzz_TEST01,ITV
        EXEC ut_zzVBJ zzz_TEST01,ATV
    ----------------------------------------------------------------------------------------------*/
    IF @BldCOD IN (@SecDTV,@SecITV,@SecATV) BEGIN
        IF @IncTtl > 0 BEGIN
            IF @BldCOD IN (@SecDTV) BEGIN
                SET @TXT = "Declare "+@RefDsc+" column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE IF @BldCOD IN (@SecITV) BEGIN
                SET @TXT = "Initialize "+@RefDsc+" column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE BEGIN
                SET @TXT = "Assign "+@RefDsc+" column values"+"  ("+@BldCOD+")"+@LstVbj
            END
            EXEC ut_zzVBX SLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
            SET @TypNam = @VarDtp+';'; SET @TypTxt = LEFT(@TypNam,@VbaTln+1)
            SET @VarVal = CASE 
                WHEN @FldDct = @DtpCatNBR AND @FldAud = 1 THEN 'gcNull'+@VarCat
                WHEN @FldDct = @DtpCatDAT AND @FldAud = 1 THEN 'Date()'
                WHEN @FldDct = @DtpCatDAT                 THEN 'Date()'
                WHEN @FldQot = 0                       THEN 'gcNull'+@VarCat
                WHEN LEN(@FldVal) = 0                  THEN 'gcNull'+@VarCat
                ELSE @VarVal
            END
            IF @BldCOD IN (@SecDTV) BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@VarDtp
            END ELSE IF @BldCOD IN (@SecITV) BEGIN
                IF @FldAud = 1 BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END ELSE BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END
            END ELSE BEGIN
                PRINT @M+@StmFld+" = "+@FldVal
            END
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DSV = Declare statement column variables
    -- ISV = Initialize statement column variables
    -- ASV = Assign statement column variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,DSV
        EXEC ut_zzVBJ zzz_TEST01,ISV
        EXEC ut_zzVBJ zzz_TEST01,ASV
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecDSV,@SecISV,@SecASV) BEGIN
        IF @IncTtl > 0 BEGIN
            IF @BldCOD IN (@SecDSV) BEGIN
                SET @TXT = "Declare "+@RefDsc+" statement column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE IF @BldCOD IN (@SecISV) BEGIN
                SET @TXT = "Initialize "+@RefDsc+" statement column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE BEGIN
                SET @TXT = "Assign "+@RefDsc+" statement column values"+"  ("+@BldCOD+")"+@LstVbj
            END
            EXEC ut_zzVBX SLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        SET @IDN = 0; SET @CNT = @StxCnt; OPEN cur_StmList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_StmList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
            SET @TypNam = @VarDtp+';'; SET @TypTxt = LEFT(@TypNam,@VbaTln+1)
            SET @VarVal = CASE 
                WHEN @FldDct = @DtpCatNBR AND @FldAud = 1 THEN 'gcNull'+@VarCat
                WHEN @FldDct = @DtpCatDAT AND @FldAud = 1 THEN 'Date()'
                WHEN @FldDct = @DtpCatDAT                 THEN 'Date()'
                WHEN @FldQot = 0                       THEN 'gcNull'+@VarCat
                WHEN LEN(@FldVal) = 0                  THEN 'gcNull'+@VarCat
                ELSE @VarVal
            END
            IF @BldCOD IN (@SecDSV) BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@VarDtp
            END ELSE IF @BldCOD IN (@SecISV) BEGIN
                IF @FldAud = 1 BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END ELSE BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END
            END ELSE BEGIN
                PRINT @M+@StmFld+" = "+@FldVal
            END
        END; CLOSE cur_StmList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DKV = Declare primary key column variables
    -- IKV = Initialize primary key column variables
    -- AKV = Assign primary key column variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,DKV
        EXEC ut_zzVBJ zzz_TEST01,IKV
        EXEC ut_zzVBJ zzz_TEST01,AKV
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecDKV,@SecIKV,@SecAKV) BEGIN
        IF @IncTtl > 0 BEGIN
            IF @BldCOD IN (@SecDKV) BEGIN
                SET @TXT = "Declare "+@RefDsc+" primary key column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE IF @BldCOD IN (@SecIKV) BEGIN
                SET @TXT = "Initialize "+@RefDsc+" primary key column variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE BEGIN
                SET @TXT = "Assign "+@RefDsc+" primary key column values"+"  ("+@BldCOD+")"+@LstVbj
            END
            EXEC ut_zzVBX SLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        SET @IDN = 0; SET @CNT = @PkfCnt; OPEN cur_PkfList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_PkfList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
            SET @TypNam = @VarDtp+';'; SET @TypTxt = LEFT(@TypNam,@VbaTln+1)
            SET @VarVal = CASE 
                WHEN @FldDct = @DtpCatNBR AND @FldAud = 1 THEN 'gcNull'+@VarCat
                WHEN @FldDct = @DtpCatDAT AND @FldAud = 1 THEN 'Date()'
                WHEN @FldDct = @DtpCatDAT                 THEN 'Date()'
                WHEN @FldQot = 0                       THEN 'gcNull'+@VarCat
                WHEN LEN(@FldVal) = 0                  THEN 'gcNull'+@VarCat
                ELSE @VarVal
            END
            IF @BldCOD IN (@SecDKV) BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@VarDtp
            END ELSE IF @BldCOD IN (@SecIKV) BEGIN
                IF @FldAud = 1 BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END ELSE BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END
            END ELSE BEGIN
                PRINT @M+@StmFld+" = "+@FldVal
            END
        END; CLOSE cur_PkfList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- DGV = Declare function parameter variables
    -- IGV = Initialize parameter variables
    -- AGV = Assign parameter variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,DGV
        EXEC ut_zzVBJ zzz_TEST01,IGV
        EXEC ut_zzVBJ zzz_TEST01,AGV
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecDGV,@SecIGV,@SecAGV) BEGIN
        IF @IncTtl > 0 BEGIN
            IF @BldCOD IN (@SecDGV) BEGIN
                SET @TXT = "Declare "+@RefDsc+" parameter variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE IF @BldCOD IN (@SecIGV) BEGIN
                SET @TXT = "Initialize "+@RefDsc+" parameter variables"+"  ("+@BldCOD+")"+@LstVbj
            END ELSE BEGIN
                SET @TXT = "Assign "+@RefDsc+" parameter values"+"  ("+@BldCOD+")"+@LstVbj
            END
            EXEC ut_zzVBX SLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        SET @IDN = 0; SET @CNT = @SigCnt; OPEN cur_SigList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_SigList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
            SET @TypNam = @VarDtp+';'; SET @TypTxt = LEFT(@TypNam,@VbaTln+1)
            SET @VarVal = CASE 
                WHEN @FldDct = @DtpCatNBR AND @FldAud = 1 THEN 'gcNull'+@VarCat
                WHEN @FldDct = @DtpCatDAT AND @FldAud = 1 THEN 'Date()'
                WHEN @FldDct = @DtpCatDAT                 THEN 'Date()'
                WHEN @FldQot = 0                       THEN 'gcNull'+@VarCat
                WHEN LEN(@FldVal) = 0                  THEN 'gcNull'+@VarCat
                ELSE @VarVal
            END
            IF @BldCOD IN (@SecDGV) BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@VarDtp
            END ELSE IF @BldCOD IN (@SecIGV) BEGIN
                IF @FldAud = 1 BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END ELSE BEGIN
                PRINT @M+"Dim "+@StmFld+" As "+@TypTxt+" "+@StmFld+" = "+@VarVal
                END
            END ELSE BEGIN
                PRINT @M+@StmFld+" = "+@FldVal
            END
        END; CLOSE cur_SigList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- IRV = Initialize recordset variables
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ zzz_TEST01,IRV
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecIRV) BEGIN
        IF @IncTtl > 0 BEGIN
            SET @TXT = "Initialize "+@RefDsc+" recordset values"+"  ("+@BldCOD+")"+@LstVbj
            EXEC ut_zzVBX SLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        --------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
            PRINT @M+""+@StmFld+" = rst.Fields("""+@FldNam+""")"
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MLV = Module Level Variables
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MLV,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ MLV,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMLV) BEGIN
    ------------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            SET @TXT = 'Module Level Variables'+'  ('+@BldCOD+')'+@LstVbj
            EXEC ut_zzVBX ALT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        --------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = ',' ELSE SET @CMA = ''
            SET @TXT = @M+'Private m'+@VarNam; SET @LEN = LEN(@TXT); IF @LEN < @StdRmgDEC SET @LEN = @StdRmgDEC
            PRINT LEFT(LEFT(@TXT+@ITX,@LEN)+' As '+@VarDtp+@ITX,@StdRmgSTM)+@SP2+@SQT+@SP1+@FldDtx
        END; CLOSE cur_ColList
        --------------------------------------------------------------------------------------------
        IF @IncTtl > 0 PRINT @M+@LinAst
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- MLP = Module Level Properties - LET/GET
    -- MLL = Module Level Properties - LET
    -- MLG = Module Level Properties - GET
    -- MLC = Module Level Properties - CLEAR
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MLP,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ MLL,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ MLG,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ MLC,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMLP,@SecMLL,@SecMLG) BEGIN
    ------------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            SET @TXT = "Module Level Properties"+"  ("+@BldCOD+")"+@LstVbj
            EXEC ut_zzVBX PLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        --------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            IF @BldCOD IN (@SecMLP,@SecMLL) BEGIN
                PRINT @M+"Public Property Let "+@FldNam+"(ByVal vNewVal As "+@VarDtp+") ' "+@FldDtx
                PRINT @M+"    m"+@VarNam+" = vNewVal"
                PRINT @M+"End Property"
            END
            IF @BldCOD IN (@SecMLP) AND 9=9 BEGIN
                PRINT LEFT(@M+@LinSgl,100)
            END
            IF @BldCOD IN (@SecMLP,@SecMLG) BEGIN
                PRINT @M+"Public Property Get "+@FldNam+"() As "+@VarDtp
                PRINT @M+@MX1+@FldNam+" = m"+@VarNam
                PRINT @M+"End Property"
            END
            PRINT @M+@LinDbl
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- MLC = Module Level Properties - CLEAR
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MLC,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMLC) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                PRINT @M+"m"+@VarNam+" = mcNul"+@FldDct
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MAV = Module Level AssignSQL (from variables)
    -- MAC = Module Level AssignSQL (from controls)
    -- MAN = Module Level AssignSQL (from nulls)
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ zzz_TEST01,MAV,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ zzz_TEST01,MAC,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ zzz_TEST01,MAN,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMAV,@SecMAC,@SecMAN) BEGIN
    ------------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            SET @TXT = "Assign SQL Values"+"  ("+@BldCOD+")"+@LstVbj
            EXEC ut_zzVBX PLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
            PRINT @M+"Private Sub AssignSQL()"
            PRINT @M+"    With RunSQL"
        END
        --------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@FldVln)
            SET @StmTmp = @VarNam; SET @StmVar = LEFT(@StmTmp,@VbaVln)
            SET @StmTmp = @CtlNam; SET @StmCtl = LEFT(@StmTmp,@VbaVln)
            IF @BldCOD IN (@SecMAV) BEGIN
                PRINT @M+"        ."+@StmFld+" = m"+@VarNam
            END ELSE IF @BldCOD IN (@SecMAC) BEGIN
                PRINT @M+"        ."+@StmFld+" = Nz(Me."+@StmCtl+", gcNull"+@FldDct+")"
            END ELSE BEGIN
                PRINT @M+"        ."+@StmFld+" = gcNull"+@FldDct
            END
        END; CLOSE cur_ColList
        --------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            PRINT @M+"    End With"
            PRINT @M+"End Sub"
            PRINT @M+@LinDbl
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MRV = Module Level ReadSQL (into variables)
    -- MRC = Module Level ReadSQL (into controls)
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MRV,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ MRC,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMRV,@SecMRC) BEGIN SET @MXX = @MX0
    ------------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            SET @TXT = "Read SQL Values"+"  ("+@BldCOD+")"+@LstVbj
            EXEC ut_zzVBX PLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
            PRINT @M+"Private Sub ReadSQL()"
            PRINT @M+"    With RunSQL"
            SET @MXX = @MX2
        END
        --------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@FldVln)
            SET @StmTmp = @VarNam; SET @StmVar = LEFT(@StmTmp,@VbaVln)
            SET @StmTmp = @CtlNam; SET @StmCtl = LEFT(@StmTmp,@VbaVln)
            IF @BldCOD IN (@SecMRV) BEGIN
                PRINT @M+@MXX+"m"+@StmVar+" = Nz(.Fields("""+@FldNam+"""), mcNul"+@FldDct+")"
            END ELSE BEGIN
                PRINT @M+@MXX+"Me."+@StmCtl+" = ."+@FldNam
            END
        END; CLOSE cur_ColList
        --------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            PRINT @M+"    End With"
            PRINT @M+"End Sub"
            PRINT @M+@LinDbl
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MCV = Module Level ClearSQL (variables)
    -- MCC = Module Level ClearSQL (controls)
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Soj        Oup Fmt Obj Dsc Dsp Oup Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ zzz_TEST01,MCV,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
        EXEC ut_zzVBJ zzz_TEST01,MCC,'' ,'' ,'' ,'' ,'' ,0  ,0  ,0  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMCV,@SecMCC) BEGIN
    ------------------------------------------------------------------------------------------------
        IF @IncTtl > 0 BEGIN
            SET @TXT = "Clear SQL Values"+"  ("+@BldCOD+")"+@LstVbj
            EXEC ut_zzVBX PLT,@TXT,@LftMrg,@IncSpc,@IncTtl,@IncBat
        END
        PRINT @M+"Public Sub ClearSQL()"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@FldVln)
            SET @StmTmp = @VarNam; SET @StmVar = LEFT(@StmTmp,@VbaVln)
            SET @StmTmp = @CtlNam; SET @StmCtl = LEFT(@StmTmp,@VbaVln)
            IF @BldCOD IN (@SecMCV) BEGIN
                PRINT @M+"    m"+@StmVar+" = gcNull"+@FldDct
            END ELSE BEGIN
                PRINT @M+"    Me."+@StmCtl+" = gcNull"+@FldDct
            END
        END; CLOSE cur_ColList
        PRINT @M+"End Sub"
        PRINT @M+@LinDbl
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MAD = Module Level Properties - ADDNEW
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MAD,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMAD) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
            PRINT @M+"AX ""    "+@CMX+"["+@FldNam+"]"""
        END; CLOSE cur_ColList
        PRINT @M+"AX "") VALUES ("""
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
            SET @QOT = CASE @FldQot WHEN 1          THEN @SQT  ELSE @SP1 END
            SET @OPN = CASE @FldDct WHEN @DtpCatTXT THEN "AQ(" ELSE @MTY END
            SET @CPN = CASE @FldDct WHEN @DtpCatTXT THEN ")"   ELSE @MTY END
            PRINT @M+"AX ""    "+@CMX+@QOT+""" & "+@OPN+"m"+@VarNam+@CPN+" & """+LTRIM(@QOT)+""""
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- MUP = Module Level Properties - UPDATE
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MUP,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMUP) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
            SET @StmTmp = '['+@FldNam+']'; SET @StmFld = LEFT(@StmTmp,@FldVln+2)
            SET @QOT = CASE @FldQot WHEN 1          THEN @SQT  ELSE @SP1 END
            SET @OPN = CASE @FldDct WHEN @DtpCatTXT THEN "AQ(" ELSE @MTY END
            SET @CPN = CASE @FldDct WHEN @DtpCatTXT THEN ")"   ELSE @MTY END
            PRINT @M+"AX ""    "+@CMX+@StmFld+" = "+@QOT+""" & "+@OPN+"m"+@VarNam+@CPN+" & """+LTRIM(@QOT)+""""
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- MIF = Module IF Criteria
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Soj        Oup Fmt Obj Dsc Dsp Oup Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ zzz_TEST01,MIF,'' ,'' ,'' ,'' ,'' ,0  ,1  ,0  ,0  ,0  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMIF) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@FldVln)
            SET @StmTmp = @VarNam; SET @StmVar = LEFT(@StmTmp,@VbaVln)
            SET @StmTmp = @CtlNam; SET @StmCtl = LEFT(@StmTmp,@VbaVln)
            SET @StmSqt = CASE @FldQot WHEN 1 THEN "'" ELSE "" END
            IF @FldDct IN (@DtpCatNBR) BEGIN
                PRINT @M+"If m"+@VarNam+" > 0 Then"
            END ELSE BEGIN
                PRINT @M+"If Len(m"+@VarNam+") > 0 Then"
            END
            PRINT @M+"    If strWHR = mWHR Then"
            PRINT @M+"        .AX strWHR: strWHR = mcMTY"
            PRINT @M+"    End If"
            PRINT @M+"    .AX strAND & pALS & ""."+@FldNam+" = "+@StmSqt+""" & m"+@VarNam+" & """+@StmSqt+""""
            PRINT @M+"    strAND = mAND"
            PRINT @M+"End If"
            PRINT @M+@LinSgl
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- MTP = Module Level Properties - Test Set Prop Values
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MTP,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMTP) BEGIN
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
            PRINT @M+"Public Property Get Test_"+@FldNam+"() As String: Test_"+@FldNam+" = """": End Property"
        END; CLOSE cur_ColList
        CONTINUE
    ------------------------------------------------------------------------------------------------
    -- MTF = Module Level Properties - Test Set Prop Functions
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBJ Bld Inp        Fmt Oup Dsc Dsp Def Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat
        EXEC ut_zzVBJ MTF,zzz_TEST01,'' ,'' ,'' ,'' ,'' ,0  ,0  ,2  ,2  ,1  ,0  ,0  ,0  ,1  ,1
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecMTF) BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT @LinSpc
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' SKY: Include SurrogateKey Only"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_SKY(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_SKY"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' MIN: Include only Minimal Properties so that one or more HasRequiredData properties will fail"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Prop_MIN(ByRef cls As "+@DspObj+"): Const pcMsgTtl As String = mcModObj & "".Set_Prop_MIN"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    With cls: Call .Clear"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        ."+@FldNam+" = .Test_"+@FldNam
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_PMN(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_PMN"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' RQD: Include HasRequiredData properties only"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Prop_RQD(ByRef cls As "+@DspObj+"): Const pcMsgTtl As String = mcModObj & "".Set_Prop_RQD"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    With cls: Call .Clear"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        ."+@FldNam+" = .Test_"+@FldNam
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_PRQ(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_PRQ"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' ANW: Assign New values; Must Include all HasRequiredData properties; Must Exclude Update properties"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Prop_ANW(ByRef cls As "+@DspObj+"): Const pcMsgTtl As String = mcModObj & "".Set_Prop_ANW"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    With cls: Call .Clear"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        ."+@FldNam+" = .Test_"+@FldNam
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_PAN(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_PAN"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' UPD: Assign several Update properties; Must Exclude any HasRequiredData properties"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Prop_UPD(ByRef cls As "+@DspObj+"): Const pcMsgTtl As String = mcModObj & "".Set_Prop_UPD"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    With cls: 'Call .Clear"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        ."+@FldNam+" = .Test_"+@FldNam
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_PUP(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_PUP"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT @M+"'==================================================================================================="
        PRINT @M+"' ALL: Must Include ALL properties"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Prop_ALL(ByRef cls As "+@DspObj+"): Const pcMsgTtl As String = mcModObj & "".Set_Prop_ALL"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    With cls: Call .Clear"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        ."+@FldNam+" = .Test_"+@FldNam
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT @M+"Private Sub Set_Mesg_PAL(ByRef cls As "+@DspObj+", ByRef strMSG As String): Const pcMsgTtl As String = mcModObj & "".Set_Mesg_PAL"":"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"    Dim strLFD As String: strLFD = IIf(Len(strMSG) > 0, vbCrLf, """"): With cls:"
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        SET @IDN = 0; SET @CNT = @ColCnt; OPEN cur_ColList; WHILE 1=1 BEGIN FETCH NEXT FROM cur_ColList INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN > 1 SET @CMX = @CMA ELSE SET @CMX = @SP1
        PRINT @M+"        strMSG = strMSG & strLFD & """+@FldNam+" = |"" & ."+@FldNam+" & ""|"": strLFD = vbCrLf"
        END; CLOSE cur_ColList
        PRINT @M+"    '-----------------------------------------------------------------------------------------------"
        PRINT @M+"    End With"
        PRINT @M+"'---------------------------------------------------------------------------------------------------"
        PRINT @M+"End Sub"
        PRINT @M+"'==================================================================================================="
        PRINT ""
        PRINT ""
        PRINT "'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- RUNUSX = Run SProc commands
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ usp_AddNew_trx_SysClm_Sync,RUNUSX
        EXEC ut_zzVBJ usp_Run_SyncCNT,RUNUSX
        EXEC ut_zzVBA RUNUSX,usp_Run_SyncCNT
    ----------------------------------------------------------------------------------------------*/
    IF @BldCOD IN (@SecRUNUSX) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @MthNam = "Execute_"+@OupObj
        SET @RtnVar = "lngPkyVal"
        SET @RtnDtp = "Long"
        SET @TmpFln = CASE WHEN LEN("RetVal"            ) > @FldFln THEN LEN("RetVal")             ELSE @FldFln END+2  -- Includes quotes
        SET @TmpPln = CASE WHEN LEN("adInteger"         ) > @VbaPln THEN LEN("adInteger")          ELSE @VbaPln END
        SET @TmpDln = CASE WHEN LEN("adParamReturnValue") > @VbaDln THEN LEN("adParamReturnValue") ELSE @VbaDln END
        SET @TmpLln = CASE WHEN LEN("4"                 ) > @VbaLln THEN LEN("4")                  ELSE @VbaLln END
        SET @TmpVln = CASE WHEN LEN("mlngRetVal"        ) > @FldFln THEN LEN("mlngRetVal")         ELSE @VbaVln END      -- Parameters use GET Fields
        PRINT ""
        PRINT ""
        PRINT "'###################################################################################################"
        PRINT "' Execute Procedure:  "+@OupObj+"   - ut_zzVBA "+@BldCOD+","+@OupObj
        PRINT "'###################################################################################################"
        PRINT "Public Function "+@MthNam+"() As "+@RtnDtp
        PRINT "    Const pcMsgTtl = mcModNam & ""."+@MthNam+""""
        PRINT "    On Error GoTo Error_Handler"
        PRINT ""
        PRINT "    ' Initialize a command object"
        PRINT "    Dim cmd    As ADODB.Command"
        PRINT "    Set cmd = New ADODB.Command"
        PRINT ""
        PRINT "    ' Assign command property values"
        PRINT "    cmd.CommandText = """+@OupObj+""""
        PRINT "    cmd.CommandType = adCmdStoredProc"
        PRINT ""
        PRINT "    ' Create command parameters"
        SET @ITX = '"RetVal"'                 ; SET @FldNam =  LEFT(@ITX,@TmpFln)
        SET @ITX = "adInteger"                ; SET @PrmDtp =  LEFT(@ITX,@TmpPln)
        SET @ITX = "adParamReturnValue"       ; SET @PrmDir =  LEFT(@ITX,@TmpDln)
        SET @ITM = LEFT(@LXX,20)+"4"        ; SET @PrmLen = RIGHT(@ITM,@TmpLln)
        SET @ITX = "mlngRetVal"               ; SET @VarNam =  LEFT(@ITX,@TmpVln)
        PRINT "    cmd.Parameters.Append cmd.CreateParameter("+@FldNam+", "+@PrmDtp+", "+@PrmDir+", "+@PrmLen+", "+@VarNam+")"
        SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            SET @TX1 = @FldNam
            SET @ITX = '"'+@FldNam+'"'        ; SET @FldNam =  LEFT(@ITX,@TmpFln)
            SET @ITX =     @PrmDtp            ; SET @PrmDtp =  LEFT(@ITX,@TmpPln)
            SET @ITX =     @PrmDir            ; SET @PrmDir =  LEFT(@ITX,@TmpDln)
            SET @ITM = LEFT(@LXX,20)+@PrmLen; SET @PrmLen = RIGHT(@ITM,@TmpLln)
            SET @ITX =     @TX1               ; SET @VarNam =  LEFT(@ITX,@TmpVln)
            PRINT "    cmd.Parameters.Append cmd.CreateParameter("+@FldNam+", "+@PrmDtp+", "+@PrmDir+", "+@PrmLen+", "+@VarNam+")"
        END; CLOSE cur_VbaLst
        PRINT ""
        PRINT "    ' Execute the command"
        PRINT "    Set cmd.ActiveConnection = Application.CurrentProject.Connection"
        PRINT "    cmd.Execute"
        SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK;
            IF @OutFlg = 1 OR @IdnFlg = 1 BEGIN
                SET @IDN += 1
                PRINT ""
                PRINT "    ' Return the identity value"
                PRINT "    "+@MthNam+" = cmd.Parameters("""+@FldNam+""").Value"
                BREAK
            END
        END; CLOSE cur_VbaLst
        IF @IDN = 0 BEGIN
            PRINT ""
            PRINT "    ' Return success"
            PRINT "    "+@MthNam+" = cmd.Parameters(""RetVal"").Value"
        END
        PRINT ""
        PRINT "Exit_Procedure:"
        PRINT "    Exit Function"
        PRINT "Error_Handler:"
        PRINT "    MsgBox pcMsgTtl & "" ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
        PRINT "    Resume Exit_Procedure"
        PRINT "End Function"
        PRINT "'==================================================================================================="
        SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            PRINT "'    usp."+@FldNam+" = "+@VarVal
        END; CLOSE cur_VbaLst
        IF @IDN > 0 BEGIN
        PRINT "'==================================================================================================="
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- SQLSTM = Build the full SQL statement
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ trx_SysObj,SQLSTM
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecSQLSTM) BEGIN
    ------------------------------------------------------------------------------------------------
        SET @MthNam = "GetSql_"+@TblBas
        --EXEC ut_zzNAM TBL,ALS,RB3,@TblNam,@TblAls OUTPUT
        PRINT ""
        PRINT ""
        PRINT "'###################################################################################################"
        PRINT "' SQL Statement:  "+@TblNam
        PRINT "'###################################################################################################"
        PRINT "Public Function "+@MthNam+"( _"
        PRINT "    Optional strWHR As String = """", _"
        PRINT "    Optional strGBY As String = """", _"
        PRINT "    Optional strHAV As String = """", _"
        PRINT "    Optional strOBY As String = """" _"
        PRINT ") As String"
        PRINT ""
        PRINT "    ' Declare alias variables"
        PRINT "    Dim lstALS() As String"
        PRINT ""
        PRINT "    ' Initialize alias values"
        PRINT "    ReDim Preserve lstALS(1): lstALS(1) = rt.GetAls("""+@TblNam+""")"
        SET @IDN = 0; SET @CNT = @FkyCnt; OPEN cur_FkyDefs; WHILE 1=1 BEGIN FETCH NEXT FROM cur_FkyDefs INTO @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
        SET @TXT = CAST(@IDN+1 AS varchar(10))
        PRINT "    ReDim Preserve lstALS("+@TXT+"): lstALS("+@TXT+") = rt.GetAls("""+@RkyTbl+""")"
        END; CLOSE cur_FkyDefs
        PRINT ""
        PRINT "    ' Initialize text concatenation variables"
        PRINT "    Call AX_Clear"
        PRINT ""
        PRINT "    ' Build base select statement"
        PRINT "    AX ""SELECT"""
        SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
        PRINT "    AX mcMG1 & lstALS(1) & ""."+@FldNam+@CMA+""""
        END; CLOSE cur_VbaLst
        PRINT "    AX ""FROM"""
        PRINT "    AX ""    "+@TblNam+" "" & lstALS(1)"
        SET @IDN = 0; SET @CNT = @FkyCnt; OPEN cur_FkyDefs; WHILE 1=1 BEGIN FETCH NEXT FROM cur_FkyDefs INTO @DfnID,@DfnCls,@DfnExs,@DfnFmt,@DfnFmx,@DfnFix,@ConTbl,@ConNam,@ConDsc,@FkyTbl,@FkyKys,@RkyTbl,@RkyKys,@DfnStd; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
        SET @TXT = CAST(@IDN+1 AS varchar(10))
        PRINT "    AX ""INNER JOIN"""
        PRINT "    AX ""    pfx_TblNam "" & lstALS("+@TXT+")"
        PRINT "    AX ""        ON "" & lstALS("+@TXT+") & ""."+@RkyKys+" = "" & lstALS(1) & ""."+@FkyKys+""""
        END; CLOSE cur_FkyDefs
        IF @IDN = 0 BEGIN
        PRINT "    'X ""INNER JOIN"""
        PRINT "    'X ""    pfx_TblNam "" & lstALS(2)"
        PRINT "    'X ""        ON "" & lstALS(2) & ""."+@FldFst+" = "" & lstALS(1) & ""."+@FldFst+""""
        END
        PRINT ""
        PRINT "    ' Assign default clauses"
        PRINT "    If Len(strOBY) = 0 Then"
        PRINT "        'strOBY = strOBY & mcOB1 & lstALS(X) & "".FldNam"""
        PRINT "        'strOBY = strOBY & mcRT2 & lstALS(X) & "".FldNam"""
        PRINT "    End If"
        PRINT ""
        PRINT "    ' Append optional clauses"
        PRINT "    BX strWHR    ' WHERE "
        PRINT "    BX strGBY    ' GROUP BY"
        PRINT "    BX strHAV    ' HAVING"
        PRINT "    BX strOBY    ' ORDER BY"
        PRINT ""
        PRINT "    ' Review SQL statement"
        PRINT "    If False Then Debug.Print mX  ' True False"
        PRINT ""
        PRINT "    ' Return SQL string"
        PRINT "    "+@MthNam+" = mX"
        PRINT ""
        PRINT "End Function"
        PRINT "'==================================================================================================="
        CONTINUE
    ------------------------------------------------------------------------------------------------


    -- --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- 
    -- 
    -- --!!!  THE REST OF THESE ARE NOT USED !!!
    -- 
    -- ------------------------------------------------------------------------------------------------
    -- -- SQLWHR = Build the SQL WHERE clause
    -- -- ut_zzVBJ trx_SysObj,SQLWHR
    -- ------------------------------------------------------------------------------------------------
    -- END ELSE IF @BldCOD IN (@SecSQLWHR) BEGIN
    -- ------------------------------------------------------------------------------------------------
    --     SET @MthNam = "GetWhr_"+@TblBas
    --     PRINT ""
    --     PRINT ""
    --     PRINT "'###################################################################################################"
    --     PRINT "Public Function "+@MthNam+"() As String"
    --     PRINT "'###################################################################################################"
    --     PRINT ""
    --     PRINT "    ' Declare SQL construction variables"
    --     PRINT "    Dim strAND As String"
    --     PRINT "    Const pcMG1 = ""    """
    --     PRINT "    Const pcWHR = ""WHERE"" & vbCrLf & pcMG1"
    --     PRINT "    Const pcAND = ""AND """
    --     PRINT ""
    --     PRINT "    ' Clear text concatenation variables"
    --     PRINT "    Call AX_Clear"
    --     PRINT ""
    --     PRINT "    ' Prepare concatenation values"
    --     PRINT "    strAND = pcWHR"
    --     PRINT ""
    --     PRINT "    ' Assign the SELECT clause"
    --     SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
    --     PRINT ""
    --     PRINT "    ' Process "+@FldNam
    --     PRINT "    If Len(mstr"+@FldNam+", """") > 0 Then"
    --     PRINT "        .AX strAND & """+@SrcAls+@DOT+@FldNam+" = "" & mstr"+@FldNam
    --     PRINT "        strAND = pcAND"
    --     PRINT "    End If"
    --     END; CLOSE cur_VbaLst
    --     PRINT ""
    --     PRINT "    ' Assign the FROM clause"
    --     PRINT "    .AX ""FROM"""
    --     PRINT "    .AX ""    "+@TblNam+" "+@SrcAls+""""
    --     PRINT "    '.AX ""INNER JOIN"""
    --     PRINT "    '.AX ""    pfx_TblNam pfx"""
    --     PRINT "    '.AX ""        ON pfx."+@FldFst+" = "+@SrcAls+@DOT+@FldFst+""""
    --     PRINT ""
    --     PRINT "    ' Value of SelectOn field"
    --     PRINT "    If Len(Trim(Nz(strSelectOn01, """"))) > 0 Then"
    --     PRINT "        .AX strAND & """+@SrcAls+@DOT+@FldFst+" = "" & Trim(strSelectOn01) & """""
    --     PRINT "        strAND = pcAND"
    --     PRINT "    End If"
    --     PRINT ""
    --     PRINT "    ' Value of SelectOn field"
    --     PRINT "    If Len(Trim(Nz(strSelectOn02, """"))) > 0 Then"
    --     PRINT "        .AX strAND & """+@SrcAls+".TaxYer = "" & Trim(strSelectOn02) & """""
    --     PRINT "        strAND = pcAND"
    --     PRINT "    End If"
    --     PRINT ""
    --     PRINT "    ' Assign the ORDER BY clause based on 'Sort By' option"
    --     PRINT "    strOBY = ""ORDER BY"" & vbCrLf"
    --     PRINT "    Select Case intSortBy"
    --     PRINT "        Case mcSortBy001"
    --     PRINT "            .AX strOBY & vbCrLf & pcMG1 & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy002"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy003"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy004"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy005"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case Else"
    --     PRINT "            strOBY = strOBY & vbCrLf & """+@SrcAls+".TaxYer, "+@SrcAls+@DOT+@FldFst+""""
    --     PRINT "    End Select"
    --     PRINT "    .AX ""FROM"""
    --     PRINT ""
    --     PRINT "    ' Record SQL statement"
    --     PRINT "    If False Then Debug.Print mX  ' True False"
    --     PRINT ""
    --     PRINT "    ' Return SQL string"
    --     PRINT "    "+@MthNam+" = mX"
    --     PRINT ""
    --     PRINT "End Function"
    --     PRINT "'==================================================================================================="
    --     CONTINUE
    -- ------------------------------------------------------------------------------------------------
    -- -- SQLOBY = Build the SQL ORDER BY clause
    -- -- ut_zzVBJ trx_SysObj,SQLOBY
    -- ------------------------------------------------------------------------------------------------
    -- END ELSE IF @BldCOD IN (@SecSQLOBY) BEGIN
    -- ------------------------------------------------------------------------------------------------
    --     SET @MthNam = "GetOby_"+@TblBas
    --     PRINT ""
    --     PRINT ""
    --     PRINT "'###################################################################################################"
    --     PRINT "Public Function "+@MthNam+"() As String"
    --     PRINT "'###################################################################################################"
    --     PRINT ""
    --     PRINT "    ' Declare SQL construction variables"
    --     PRINT "    Dim strOBY As String"
    --     PRINT "    Const pcMG1 = ""    """
    --     PRINT "    Const pcOBY = ""ORDER BY"" & vbCrLf & pcMG1"
    --     PRINT ""
    --     PRINT "    ' Clear text concatenation variables"
    --     PRINT "    Call AX_Clear"
    --     PRINT ""
    --     PRINT "    ' Prepare concatenation values"
    --     PRINT "    strOBY = pcOBY"
    --     PRINT ""
    --     PRINT "    ' Assign the SELECT clause"
    --     SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
    --     PRINT ""
    --     PRINT "    ' Process "+@FldNam
    --     PRINT "    If Len(mstr"+@FldNam+", """") > 0 Then"
    --     PRINT "        .AX strOBY & """+@SrcAls+@DOT+@FldNam+""","""
    --     PRINT "        strAND = pcMG1"
    --     PRINT "    End If"
    --     END; CLOSE cur_VbaLst
    --     PRINT ""
    --     PRINT "    ' Record SQL statement"
    --     PRINT "    If False Then Debug.Print mX  ' True False"
    --     PRINT ""
    --     PRINT "    ' Return SQL string"
    --     PRINT "    "+@MthNam+" = mX"
    --     PRINT ""
    --     PRINT "End Function"
    --     PRINT "'==================================================================================================="
    --     CONTINUE
    -- ------------------------------------------------------------------------------------------------
    -- 
    -- 
    -- ------------------------------------------------------------------------------------------------
    -- -- BLDSQL = Build SQL Text from SelectOn controls
    -- -- ut_zzVBJ trx_SysObj,BLDSQL
    -- ------------------------------------------------------------------------------------------------
    -- END ELSE IF @BldCOD IN (@SecBLDSQL) BEGIN
    -- ------------------------------------------------------------------------------------------------
    --     PRINT ""
    --     PRINT ""
    --     PRINT "'###################################################################################################"
    --     PRINT "Public Function BuildSQLSelect( _"
    --     PRINT "    strSelectOn01 As String, _"
    --     PRINT "    strSelectOn02 As String, _"
    --     PRINT "    intSortBy As Integer _"
    --     PRINT ") As String"
    --     PRINT "'###################################################################################################"
    --     PRINT ""
    --     PRINT "    ' Assign tracking values"
    --     PRINT "    mstrPKey = Trim$(strSelectOn01)"
    --     PRINT "    mstrYear = Trim$(strSelectOn02)"
    --     PRINT ""
    --     PRINT "    ' Declare SQL construction variables"
    --     PRINT "    Dim strSQL As String"
    --     PRINT "    Dim strSEL As String"
    --     PRINT "    Dim strFRM As String"
    --     PRINT "    Dim strWHR As String"
    --     PRINT "    Dim strGBY As String"
    --     PRINT "    Dim strHAV As String"
    --     PRINT "    Dim strOBY As String"
    --     PRINT "    Dim strAND As String"
    --     PRINT "    Dim strCMA As String"
    --     PRINT "    Dim strSPC As String"
    --     PRINT "    Const pcSEL = ""SELECT """
    --     PRINT "    Const pcFRM = ""FROM """
    --     PRINT "    Const pcWHR = ""WHERE"""
    --     PRINT "    Const pcGBY = ""GROUP BY"""
    --     PRINT "    Const pcHAV = ""HAVING"""
    --     PRINT "    Const pcOBY = ""ORDER BY"""
    --     PRINT "    Const pcAND = ""AND """
    --     PRINT "    Const pcCMA = "", """
    --     PRINT "    Const pcSPC = "" """
    --     PRINT ""
    --     PRINT "    ' Assign the SELECT clause"
    --     PRINT "    strSEL = ""SELECT"""
    --     SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
    --     PRINT "    strSEL = strSEL & vbCrLf & ""    "+@SrcAls+@DOT+@FldNam+@CMA
    --     END; CLOSE cur_VbaLst
    --     PRINT ""
    --     PRINT "    ' Assign the FROM clause"
    --     PRINT "    strFRM = ""FROM"""
    --     PRINT "    strFRM = strFRM & vbCrLf & ""    "+@TblNam+" "+@SrcAls+""""
    --     PRINT "    'trFRM = strFRM & vbCrLf & ""INNER JOIN"""
    --     PRINT "    'trFRM = strFRM & vbCrLf & ""    pfx_TblNam pfx"""
    --     PRINT "    'trFRM = strFRM & vbCrLf & ""        ON pfx."+@FldFst+" = "+@SrcAls+@DOT+@FldFst+""""
    --     PRINT ""
    --     PRINT "    ' Assign the WHERE clause based on 'Select On' criteria"
    --     PRINT "    strWHR = pcWHR"
    --     PRINT "    strAND = ""    """
    --     PRINT ""
    --     PRINT "    ' Value of SelectOn field"
    --     PRINT "    If Len(Trim(Nz(strSelectOn01, """"))) > 0 Then"
    --     PRINT "        strWHR = strWHR & vbCrLf & strAND & """+@SrcAls+@DOT+@FldFst+" = "" & Trim(strSelectOn01) & """""
    --     PRINT "        strAND = pcAND"
    --     PRINT "    End If"
    --     PRINT ""
    --     PRINT "    ' Value of SelectOn field"
    --     PRINT "    If Len(Trim(Nz(strSelectOn02, """"))) > 0 Then"
    --     PRINT "        strWHR = strWHR & vbCrLf & strAND & """+@SrcAls+".TaxYer = "" & Trim(strSelectOn02) & """""
    --     PRINT "        strAND = pcAND"
    --     PRINT "    End If"
    --     PRINT ""
    --     PRINT "    strWHR = IIf(strWHR = pcWHR, """", Trim(strWHR))"
    --     PRINT ""
    --     PRINT "    ' Assign the ORDER BY clause based on 'Sort By' option"
    --     PRINT "    strOBY = pcOBY"
    --     PRINT "    Select Case intSortBy"
    --     PRINT "        Case mcSortBy001"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy002"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy003"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy004"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case mcSortBy005"
    --     PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
    --     PRINT "        Case Else"
    --     PRINT "            strOBY = strOBY & vbCrLf & """+@SrcAls+".TaxYer, "+@SrcAls+@DOT+@FldFst+""""
    --     PRINT "    End Select"
    --     PRINT "    strOBY = IIf(strOBY = pcOBY, """", strOBY)"
    --     PRINT ""
    --     PRINT "    ' Build the SELECT SQL string"
    --     PRINT "    strSQL = strSEL"
    --     PRINT "    strSQL = strSQL & IIf(Len(strFRM) > 0, vbCrLf, """") & strFRM"
    --     PRINT "    strSQL = strSQL & IIf(Len(strWHR) > 0, vbCrLf, """") & strWHR"
    --     PRINT "    strSQL = strSQL & IIf(Len(strGBY) > 0, vbCrLf, """") & strGBY"
    --     PRINT "    strSQL = strSQL & IIf(Len(strHAV) > 0, vbCrLf, """") & strHAV"
    --     PRINT "    strSQL = strSQL & IIf(Len(strOBY) > 0, vbCrLf, """") & strOBY"
    --     PRINT ""
    --     PRINT "    ' Record SQL statement  True False"
    --     PRINT "    If False Then Debug.Print strSQL"
    --     PRINT ""
    --     PRINT "    ' Return SQL string"
    --     PRINT "    BuildSQLSelect = strSQL"
    --     PRINT ""
    --     PRINT "End Function"
    --     PRINT "'==================================================================================================="
    --     CONTINUE
    -- ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    END
    ------------------------------------------------------------------------------------------------


    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


    ------------------------------------------------------------------------------------------------
    -- FRMHDR = Header form
    -- FRMDTL = Detail form
    -- FRMLST = List form
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- POPADD = PopAdd form
    -- POPUPD = PopUpd form
    --@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    -- RECINS = Insert record function
    -- RECUPD = Update record function
    /*----------------------------------------------------------------------------------------------
        EXEC ut_zzVBJ ref_UtlDef,RECINS
        EXEC ut_zzVBJ usp_AddNew_ref_UtlDef,RECINS
    ----------------------------------------------------------------------------------------------*/
    IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST,@SecPOPADD,@SecPOPUPD,@SecRECINS,@SecRECUPD) BEGIN
    ------------------------------------------------------------------------------------------------
        --SET @DefTyp = "VBA"
        SET @FncNam = "InsertRecord"
        SET @MthNam = "Process_"+@OupObj
        SET @SrcAls = CASE
            WHEN LEN(@SrcAls) > 0 THEN @SrcAls
            ELSE LOWER(RIGHT(@TblNam,3))
        END
        --SELECT * FROM #VbaDfn
        --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
        -- ut_zzVBX FRMSDC,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form module declarations
        -- ut_zzVBX ADFFOP,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form open function
        -- ut_zzVBX ADPFOP,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Popup form open functions
        -- ut_zzVBX ADFFUN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form module functions
        --------------------------------------------------------------------------------------------
        -- ut_zzVBX ADETXX,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Text Box (text)
        -- ut_zzVBX ADETXC,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Text Box (code)
        -- ut_zzVBX ADETXN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Text Box (numeric)
        -- ut_zzVBX ADETXD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Text Box (date)
        -- ut_zzVBX ADECBO,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Combo Box
        -- ut_zzVBX ADECHK,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Event - Check Box
        --------------------------------------------------------------------------------------------
        -- ut_zzVBX ADPCMD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Popup form commands
        -- ut_zzVBX ADFCMD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form commands
        -- ut_zzVBX ADFPRN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Print default report
        -- ut_zzVBX ADFXTD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Extend object for active tax years
        -- ut_zzVBX ADFSYN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Synchronize related objects
        -- ut_zzVBX ADFOPN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Open external form
        -- ut_zzVBX ADIVFY,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Verify Insert function
        --------------------------------------------------------------------------------------------
        IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST,@SecPOPADD,@SecPOPUPD) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX FRMSDC,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form module declarations
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADFFOP,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form open function
            EXEC ut_zzVBX ADFFUN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form module functions
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecPOPADD,@SecPOPUPD) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADPFOP,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Popup form open functions
            EXEC ut_zzVBX ADFFUN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form module functions
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADFSQ1,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- SQL - Manage SelectOn synchronization
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "Private Function BuildSQLSelect( _"
            PRINT "    strSelectOn01 As String, _"
            PRINT "    strSelectOn02 As String, _"
            PRINT "    intSortBy As Integer _"
            PRINT ") As String"
            PRINT "'###################################################################################################"
            PRINT ""
            PRINT "    ' Assign tracking values"
            PRINT "    mstrPKey = Trim$(strSelectOn01)"
            PRINT "    mstrYear = Trim$(strSelectOn02)"
            PRINT ""
            PRINT "    ' Declare SQL construction variables"
            PRINT "    Dim strSQL As String"
            PRINT "    Dim strSEL As String"
            PRINT "    Dim strFRM As String"
            PRINT "    Dim strWHR As String"
            PRINT "    Dim strGBY As String"
            PRINT "    Dim strHAV As String"
            PRINT "    Dim strOBY As String"
            PRINT "    Dim strAND As String"
            PRINT "    Dim strCMA As String"
            PRINT "    Dim strSPC As String"
            PRINT "    Const pcSEL = ""SELECT """
            PRINT "    Const pcFRM = ""FROM """
            PRINT "    Const pcWHR = ""WHERE"""
            PRINT "    Const pcGBY = ""GROUP BY"""
            PRINT "    Const pcHAV = ""HAVING"""
            PRINT "    Const pcOBY = ""ORDER BY"""
            PRINT "    Const pcAND = ""AND """
            PRINT "    Const pcCMA = "", """
            PRINT "    Const pcSPC = "" """
            PRINT ""
            PRINT "    ' Assign the SELECT clause"
            PRINT "    strSEL = ""SELECT"""
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            PRINT "    strSEL = strSEL & vbCrLf & ""    "+@SrcAls+@DOT+@FldNam+@CMA
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT "    ' Assign the FROM clause"
            PRINT "    strFRM = ""FROM"""
            PRINT "    strFRM = strFRM & vbCrLf & ""    "+@TblNam+" "+@SrcAls+""""
            PRINT "    'trFRM = strFRM & vbCrLf & ""INNER JOIN"""
            PRINT "    'trFRM = strFRM & vbCrLf & ""    pfx_TblNam pfx"""
            PRINT "    'trFRM = strFRM & vbCrLf & ""        ON pfx."+@FldFst+" = "+@SrcAls+@DOT+@FldFst+""""
            PRINT ""
            PRINT "    ' Assign the WHERE clause based on 'Select On' criteria"
            PRINT "    strWHR = pcWHR"
            PRINT "    strAND = ""    """
            PRINT ""
            PRINT "    ' Value of SelectOn field"
            PRINT "    If Len(Trim(Nz(strSelectOn01, """"))) > 0 Then"
            PRINT "        strWHR = strWHR & vbCrLf & strAND & """+@SrcAls+@DOT+@FldFst+" = "" & Trim(strSelectOn01) & """""
            PRINT "        strAND = pcAND"
            PRINT "    End If"
            PRINT ""
            PRINT "    ' Value of SelectOn field"
            PRINT "    If Len(Trim(Nz(strSelectOn02, """"))) > 0 Then"
            PRINT "        strWHR = strWHR & vbCrLf & strAND & """+@SrcAls+".TaxYer = "" & Trim(strSelectOn02) & """""
            PRINT "        strAND = pcAND"
            PRINT "    End If"
            PRINT ""
            PRINT "    strWHR = IIf(strWHR = pcWHR, """", Trim(strWHR))"
            PRINT ""
            PRINT "    ' Assign the ORDER BY clause based on 'Sort By' option"
            PRINT "    strOBY = pcOBY"
            PRINT "    Select Case intSortBy"
            PRINT "        Case mcSortBy001"
            PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
            PRINT "        Case mcSortBy002"
            PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
            PRINT "        Case mcSortBy003"
            PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
            PRINT "        Case mcSortBy004"
            PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
            PRINT "        Case mcSortBy005"
            PRINT "            strOBY = strOBY & vbCrLf & ""pfx.FldNam"""
            PRINT "        Case Else"
            PRINT "            strOBY = strOBY & vbCrLf & """+@SrcAls+".TaxYer, "+@SrcAls+@DOT+@FldFst+""""
            PRINT "    End Select"
            PRINT "    strOBY = IIf(strOBY = pcOBY, """", strOBY)"
            PRINT ""
            PRINT "    ' Build the SELECT SQL string"
            PRINT "    strSQL = strSEL"
            PRINT "    strSQL = strSQL & IIf(Len(strFRM) > 0, vbCrLf, """") & strFRM"
            PRINT "    strSQL = strSQL & IIf(Len(strWHR) > 0, vbCrLf, """") & strWHR"
            PRINT "    strSQL = strSQL & IIf(Len(strGBY) > 0, vbCrLf, """") & strGBY"
            PRINT "    strSQL = strSQL & IIf(Len(strHAV) > 0, vbCrLf, """") & strHAV"
            PRINT "    strSQL = strSQL & IIf(Len(strOBY) > 0, vbCrLf, """") & strOBY"
            PRINT ""
            PRINT "    ' Record SQL statement  True False"
            PRINT "    If False Then Debug.Print strSQL"
            PRINT ""
            PRINT "    ' Return SQL string"
            PRINT "    BuildSQLSelect = strSQL"
            PRINT ""
            PRINT "End Function"
            PRINT "'==================================================================================================="
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADFSQ2,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- SQL - Manage SelectOn events
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST,@SecPOPADD,@SecPOPUPD) BEGIN
        --------------------------------------------------------------------------------------------
            SET @TXT = CASE WHEN @BldCOD IN (@SecPOPADD,@SecPOPUPD) THEN "PopUp" ELSE "" END
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
                IF          @CtlPfx = @CtlPfxCBO      BEGIN
                    EXEC ut_zzVBX ADECBO,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Combo Box
                END ELSE IF @CtlPfx = @CtlPfxCHK      BEGIN
                    EXEC ut_zzVBX ADECHK,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Check Box
                END ELSE IF @VarCat = @VbaCatBLN      BEGIN
                    EXEC ut_zzVBX ADECHK,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Check Box
                END ELSE IF @VarCat = @VbaCatNUM      BEGIN
                    EXEC ut_zzVBX ADETXN,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Text Box (numeric)
                END ELSE IF @VarCat = @VbaCatDAT      BEGIN
                    EXEC ut_zzVBX ADETXD,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Text Box (date)
                END ELSE IF @UpcFlg = 1               BEGIN
                    EXEC ut_zzVBX ADETXC,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Text Box (code)
                END ELSE BEGIN
                    EXEC ut_zzVBX ADETXX,@FldNam               ,0  ,0  ,0  ,0  ,'' ,'' ,@TXT    ,0  ,0  ,0  -- Event - Text Box (text)
                END
            END; CLOSE cur_VbaLst
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecFRMHDR,@SecFRMDTL,@SecFRMLST) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADFCMD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Standard form commands
            EXEC ut_zzVBX ADFPRN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Print default report
            EXEC ut_zzVBX ADFXTD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Extend object for active tax years
            EXEC ut_zzVBX ADFSYN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Synchronize related objects
            EXEC ut_zzVBX ADFOPN,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Open external form
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecPOPADD,@SecPOPUPD) BEGIN
        --------------------------------------------------------------------------------------------
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADPCMD,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0  -- Form - Popup form commands
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "Private Function DirtyFields() As Boolean"
            PRINT "'###################################################################################################"
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Initialize the dirty flag"
            PRINT "    Dim blnDirty As Boolean"
            PRINT ""
            PRINT "    ' Check each text, number or date field for dirt"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                IF          @CtlPfx IN (@CtlPfxCHK) BEGIN
                    PRINT "    blnDirty = IIf(Me."+@CtlNam+" = True, True, blnDirty)"
                END ELSE IF @CtlPfx IN (@CtlPfxCBO) BEGIN
                    PRINT "    blnDirty = IIf(Nz(Me."+@CtlNam+", 0) > 0, True, blnDirty)"
                END ELSE IF @VarCat  = @VbaCatNUM   BEGIN
                    PRINT "    blnDirty = IIf(Len(Trim(Nz(Me."+@CtlNam+", """"))) > 0, True, blnDirty)  ' Me.txtStxPct > 0"
                END ELSE BEGIN
                    PRINT "    blnDirty = IIf(Len(Trim(Nz(Me."+@CtlNam+", """"))) > 0, True, blnDirty)"
                END
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT "    ' Return result"
            PRINT "    DirtyFields = blnDirty"
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Function"
            PRINT "Error_Handler:"
            PRINT "    MsgBox ""DirtyFields ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End Function"
            PRINT "'==================================================================================================="
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "Private Sub EnableFields()"
            PRINT "'###################################################################################################"
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Enable/Disable the form fields"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
            PRINT "    Me."+@CtlNam+".Enabled = True"
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT "    ' Enable/Disable the form buttons"
            PRINT "    Me.cmdAccept.Enabled = True"
            PRINT "    Me.cmdCancel.Enabled = True"
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Sub"
            PRINT "Error_Handler:"
            PRINT "    MsgBox ""EnableFields ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End Sub"
            PRINT "'==================================================================================================="
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "Private Sub EnableAcceptButton()"
            PRINT "'###################################################################################################"
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Enable/disable the Accept button"
            SET @TXT = "IF     "
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                IF          @CtlPfx IN (@CtlPfxCHK) BEGIN
                    PRINT "    "+@TXT+" Me."+@CtlNam+" = True Then"
                    PRINT "        Me.cmdAccept.Enabled = False"
                END ELSE IF @CtlPfx IN (@CtlPfxCBO) BEGIN
                    PRINT "    "+@TXT+" Nz(Me."+@CtlNam+", 0) > 0 Then"
                    PRINT "        Me.cmdAccept.Enabled = False"
                END ELSE IF @VarCat  = @VbaCatNUM   BEGIN
                    PRINT "    "+@TXT+" Len(Trim(Nz(Me."+@CtlNam+", """"))) = 0 Then  ' Me.txtStxPct > 0"
                    PRINT "        Me.cmdAccept.Enabled = False"
                END ELSE BEGIN
                    PRINT "    "+@TXT+" Len(Trim(Nz(Me."+@CtlNam+", """"))) = 0 Then"
                    PRINT "        Me.cmdAccept.Enabled = False"
                END
                SET @TXT = "ElseIf "
            END; CLOSE cur_VbaLst
            PRINT "    Else"
            PRINT "        Me.cmdAccept.Enabled = True"
            PRINT "    End If"
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Sub"
            PRINT "Error_Handler:"
            PRINT "    MsgBox ""EnableAcceptButton ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End Sub"
            PRINT "'==================================================================================================="
            --   ut_zzVBX Oup    Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3      Trn Idn Erm
            EXEC ut_zzVBX ADIVFY,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,''      ,0  ,0  ,0
        --------------------------------------------------------------------------------------------
        END
        IF @BldCOD IN (@SecPOPADD,@SecRECINS,@SecRECUPD) BEGIN
        --------------------------------------------------------------------------------------------
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            IF          @BldCOD IN (@SecPOPADD,@SecRECINS) BEGIN
                SET @FncTyp = "Function"
                SET @FncNam = "InsertRecord"
                SET @RunNam = "Run_"+REPLACE(REPLACE(@OupObj,"usp_",""),"AddNew","Update")
                SET @MthNam = "Process_"+REPLACE(@OupObj,"Update","AddNew")
                SET @RtnVar = "lngPkyVal"
                SET @RtnDtp = "Long"
                PRINT "' Insert the record - EXEC ut_zzVBA "+@BldCOD+","+@OupObj
                PRINT "'###################################################################################################"
                PRINT "Private "+@FncTyp+" "+@FncNam+"() As Long"
            END ELSE IF @BldCOD IN (@SecRECUPD) BEGIN
                SET @FncTyp = "Function"
                SET @FncNam = "UpdateRecord"
                SET @RunNam = "Run_"+REPLACE(REPLACE(@OupObj,"usp_",""),"AddNew","Update")
                SET @MthNam = "Process_"+REPLACE(@OupObj,"AddNew","Update")
                SET @RtnVar = "blnRetVal"
                SET @RtnDtp = "Boolean"
                PRINT "' Update the record - EXEC ut_zzVBA "+@BldCOD+","+@OupObj
                PRINT "'###################################################################################################"
                PRINT "Private "+@FncTyp+" "+@FncNam+"() As Long"
            END
            PRINT "    Const pcMsgTtl = mcModNam & ""."+@FncNam+""""
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Hide the cursor"
            PRINT "    Me.cmdHideCursor.SetFocus"
            IF @VbaCnt > 0 BEGIN
                PRINT ""
                PRINT "    ' Declare parameter variables"
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                    PRINT "    Dim "+@StmFld+" As "+@VarDtp
                END; CLOSE cur_VbaLst
                PRINT ""
                PRINT "    ' Assign parameter values"
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @StmTmp = @VarNam; SET @VarNam = LEFT(@StmTmp,@VbaVln)
                    SET @StmTmp = @CtlNam; SET @CtlNam = LEFT(@StmTmp,@VbaVln)
                    IF @IdnFlg = 0 BEGIN
                        PRINT "    "+@VarNam+" = Nz(Me."+@CtlNam+", "+@VarVal+")"
                    END ELSE BEGIN
                        PRINT "    "+@VarNam+" = 0"
                    END
                END; CLOSE cur_VbaLst
            END
            PRINT ""
            PRINT "    ' Initialize the process class object"
            PRINT "    Dim run    As clsProcess"
            PRINT "    Set run = New clsProcess"
            PRINT ""
            PRINT "    ' Run the process and return the result"
            IF @VbaCnt > 0 BEGIN
                IF @FncTyp = "Function" BEGIN
                    PRINT "    "+@FncNam+" = run."+@MthNam+"( _"
                END ELSE BEGIN
                    PRINT "    Call run."+@MthNam+"( _"
                END
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    PRINT "        "+@VarNam+@CMA+" _"
                END; CLOSE cur_VbaLst
                PRINT "    )"
            END ELSE BEGIN
                IF @FncTyp = "Function" BEGIN
                    PRINT "    "+@FncNam+" = run."+@MthNam+"()"
                END ELSE BEGIN
                    PRINT "    Call run."+@MthNam+"()"
                END
            END
            PRINT ""
            PRINT "    ' Refresh the screen"
            PRINT "    Call Form_Requery"
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Function"
            PRINT "Error_Handler:"
            PRINT "    MsgBox pcMsgTtl & "" ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End "+@FncTyp+""
            PRINT "'==================================================================================================="
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "' Process "+@OupObj+" - ut_zzVBA "+@BldCOD+","+@OupObj
            PRINT "'###################################################################################################"
            IF @VbaCnt > 0 BEGIN
                PRINT "Public Function "+@MthNam+"( _"
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                    PRINT "    "+@StmFld+" As "+@VarDtp+@CMA+" _"
                END; CLOSE cur_VbaLst
                PRINT ") As "+@RtnDtp
            END ELSE BEGIN
                PRINT "Public Function "+@MthNam+"() As "+@RtnDtp
            END
            PRINT "    Const pcMsgTtl = mcModNam & ""."+@MthNam+""""
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Initialize a command object"
            PRINT "    Dim cmd    As ADODB.Command"
            PRINT "    Set cmd = New ADODB.Command"
            PRINT ""
            PRINT "    ' Assign command property values"
            PRINT "    cmd.CommandText = """+@OupObj+""""
            PRINT "    cmd.CommandType = adCmdStoredProc"
            IF @VbaCnt > 0 BEGIN
                PRINT ""
                PRINT "    ' Create command parameters"
                PRINT "    cmd.Parameters.Append cmd.CreateParameter(""RetVal"", adInteger, adParamReturnValue, 4, mlngRetVal)"
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @ClmNam = '"'+@FldNam+'"'; SET @FldNam = LEFT(@StmTmp,@FldFln+2)
                    SET @ClmNam =     @PrmDtp    ; SET @PrmDtp = LEFT(@StmTmp,@VbaPln)
                    SET @ClmNam =     @PrmDir    ; SET @PrmDir = LEFT(@StmTmp,@VbaDln)
                    SET @PrmLen = RIGHT(REPLICATE(" ",20)+@PrmLen,@VbaLln)
                    SET @ClmNam =     @VarNam    ; SET @VarNam = LEFT(@StmTmp,@VbaVln)
                    PRINT "    cmd.Parameters.Append cmd.CreateParameter("+@FldNam+", "+@PrmDtp+", "+@PrmDir+", "+@PrmLen+", "+@VarNam+")"
                END; CLOSE cur_VbaLst
            END
            PRINT ""
            PRINT "    ' Execute the command"
            PRINT "    Set cmd.ActiveConnection = Application.CurrentProject.Connection"
            PRINT "    cmd.Execute"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK;
                IF @IdnFlg = 1 BEGIN
                    IF @RtnVar = "Long" BEGIN
                        SET @IDN += 1
                        PRINT ""
                        PRINT "    ' Return the identity value"
                        PRINT "    "+@MthNam+" = cmd.Parameters("""+@FldNam+""").Value"
                    END
                    BREAK
                END
            END; CLOSE cur_VbaLst
            IF @IDN = 0 BEGIN
                PRINT ""
                PRINT "    ' Return success"
                PRINT "    "+@MthNam+" = cmd.Parameters(""RetVal"").Value"
            END
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Function"
            PRINT "Error_Handler:"
            PRINT "    MsgBox pcMsgTtl & "" ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End Function"
            PRINT "'==================================================================================================="
            PRINT ""
            PRINT ""
            PRINT "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            PRINT ""
            PRINT ""
            PRINT "    ' Run function working variables"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln+3)
                PRINT "    Dim "+@StmFld+" As "+@VarDtp
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT ""
            PRINT "    ' Select recordset fields"
            PRINT "    strSQL = ""SELECT"""
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                PRINT "    strSQL = strSQL & vbCrLf & ""    xxx."+@FldNam+@CMA+""""
            END; CLOSE cur_VbaLst
            PRINT "    strSQL = strSQL & vbCrLf & ""FROM"""
            PRINT "    strSQL = strSQL & vbCrLf & ""    "+@TblNam+" xxx"""
            PRINT "    strSQL = strSQL & vbCrLf & ""WHERE"""
            PRINT "    strSQL = strSQL & vbCrLf & ""    xxx."+@FstFld+" = ''"""
            PRINT "    strSQL = strSQL & vbCrLf & ""ORDER BY"""
            PRINT "    strSQL = strSQL & vbCrLf & ""    xxx."+@FstFld+""""
            PRINT ""
            PRINT ""
            PRINT "    ' Assign recordset field values to Run variables"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                PRINT "    "+@VarNam+" = .Fields("""+@FldNam+""")"
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT ""
            PRINT "    ' Call the Run function and return the result"
            IF @VbaCnt > 0 BEGIN
                PRINT "    "+@RtnVar+" = "+@RunNam+"( _"
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    PRINT "        "+@VarNam+@CMA+" _"
                END; CLOSE cur_VbaLst
                PRINT "    )"
            END ELSE BEGIN
                PRINT "    "+@RtnVar+" = "+@RunNam+"()"
            END
            PRINT ""
            PRINT ""
            PRINT "    ' Run the process using working variables"
            PRINT "    Dim run    As clsProcess"
            PRINT "    Set run = New clsProcess"
            PRINT "    With run"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                PRINT "        ."+@StmFld+" = "+@VarNam
            END; CLOSE cur_VbaLst
            PRINT "       "+@RtnVar+" = ."+@MthNam
            PRINT "    End With"
            PRINT ""
            PRINT ""
            PRINT "    ' Run the process using recordset fields"
            PRINT "    Dim run    As clsProcess"
            PRINT "    Set run = New clsProcess"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                PRINT "    run."+@StmFld+" = .Fields("""+@FldNam+""")"
            END; CLOSE cur_VbaLst
            PRINT "    "+@RtnVar+" = run."+@MthNam
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "' Run "+@OupObj+" - ut_zzVBA "+@BldCOD+","+@OupObj
            PRINT "'###################################################################################################"
            IF @VbaCnt > 0 BEGIN
                PRINT "Public Function "+@RunNam+"("
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @StmTmp = @VarNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                    PRINT "    "+@StmFld+" As "+@VarDtp+@CMA+" _"
                END; CLOSE cur_VbaLst
                PRINT ") As "+@RtnDtp
            END ELSE BEGIN
                PRINT "Public Function "+@RunNam+"() As "+@RtnDtp
            END
            PRINT "    Const pcMsgTtl = mcModNam & ""."+@MthNam+""""
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Initialize the process class object"
            PRINT "    Dim run    As clsProcess"
            PRINT "    Set run = New clsProcess"
            PRINT "    With run"
            PRINT ""
            PRINT "        ' Assign parameter values"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                SET @StmTmp = @FldNam; SET @StmFld = LEFT(@StmTmp,@VbaVln)
                PRINT "        ."+@StmFld+" = "+@VarNam
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT "       ' Run the process and return the result"
            PRINT "       "+@RunNam+" = ."+@MthNam
            PRINT ""
            PRINT "    End With"
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Function"
            PRINT "Error_Handler:"
            PRINT "    MsgBox pcMsgTtl & "" ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End "+@FncTyp+""
            PRINT "'==================================================================================================="
            PRINT ""
            PRINT ""
            PRINT "' Parameter Properties:  "+@TblNam
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                PRINT "Private m"+LEFT(@VarNam+REPLICATE(" ",26),26)+" As "+LEFT(@VarDtp+REPLICATE(" ",17),17)+" ' "
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "' Parameter Properties:  "+@TblNam
            PRINT "'###################################################################################################"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                PRINT "Public Property Let "+@FldNam+"(ByVal vNewVal As "+@VarDtp+"): m"+@VarNam+" = vNewVal: End Property"
            END; CLOSE cur_VbaLst
            PRINT ""
            PRINT ""
            PRINT "'###################################################################################################"
            PRINT "' Process "+@OupObj+" - ut_zzVBA "+@BldCOD+","+@OupObj
            PRINT "'###################################################################################################"
            PRINT "Public Function "+@MthNam+"() As "+@RtnDtp
            PRINT "    Const pcMsgTtl = mcModNam & ""."+@MthNam+""""
            PRINT "    On Error GoTo Error_Handler"
            PRINT ""
            PRINT "    ' Initialize a command object"
            PRINT "    Dim cmd    As ADODB.Command"
            PRINT "    Set cmd = New ADODB.Command"
            PRINT ""
            PRINT "    ' Assign command property values"
            PRINT "    cmd.CommandText = """+@OupObj+""""
            PRINT "    cmd.CommandType = adCmdStoredProc"
            PRINT ""
            PRINT "    ' Create command parameters"
            PRINT "    cmd.Parameters.Append cmd.CreateParameter(""RetVal"", adInteger, adParamReturnValue, 4, mlngRetVal)"
            IF @VbaCnt > 0 BEGIN
                SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK; SET @IDN += 1; IF @IDN < @CNT SET @CMA = "," ELSE SET @CMA = ""
                    SET @ClmNam = '"'+@FldNam+'"'; SET @FldNam = LEFT(@StmTmp,@FldFln+2)
                    SET @ClmNam =     @PrmDtp    ; SET @PrmDtp = LEFT(@StmTmp,@VbaPln)
                    SET @ClmNam =     @PrmDir    ; SET @PrmDir = LEFT(@StmTmp,@VbaDln)
                    SET @PrmLen = RIGHT(REPLICATE(" ",20)+@PrmLen,@VbaLln)
                    SET @ClmNam =     @VarNam    ; SET @VarNam = LEFT(@StmTmp,@VbaVln)
                    PRINT "    cmd.Parameters.Append cmd.CreateParameter("+@FldNam+", "+@PrmDtp+", "+@PrmDir+", "+@PrmLen+", m"+@VarNam+")"
                END; CLOSE cur_VbaLst
            END
            PRINT ""
            PRINT "    ' Execute the command"
            PRINT "    Set cmd.ActiveConnection = Application.CurrentProject.Connection"
            PRINT "    cmd.Execute"
            SET @IDN = 0; SET @CNT = @VbaCnt; OPEN cur_VbaLst; WHILE 1=1 BEGIN FETCH NEXT FROM cur_VbaLst INTO @FldID,@FldLvl,@FldObj,@FldOrd,@FldNam,@FldUtp,@FldDtx,@FldLen,@FldDct,@FldQot,@FldNul,@FldIdn,@FldOup,@FldPko,@FldFko,@FldLko,@FldAud,@FldVal,@FldDfv,@FldVfx,@VbaID,@OutFlg,@IdnFlg,@UpcFlg,@CboFlg,@ChkFlg,@DirPrp,@VarCat,@VarPfx,@VarNam,@VarDtp,@VarDfc,@VarVal,@VarNul,@CtlPfx,@CtlNam,@PrmDtp,@PrmDir,@PrmLen; IF @@FETCH_STATUS <> 0 BREAK;
                IF @IdnFlg = 1 BEGIN
                    IF @RtnVar = "Long" BEGIN
                        SET @IDN += 1
                        PRINT ""
                        PRINT "    ' Return the identity value"
                        PRINT "    "+@MthNam+" = cmd.Parameters("""+@FldNam+""").Value"
                    END
                    BREAK
                END
            END; CLOSE cur_VbaLst
            IF @IDN = 0 BEGIN
                PRINT ""
                PRINT "    ' Return success"
                PRINT "    "+@MthNam+" = (cmd.Parameters(""RetVal"").Value > 0)"
            END
            PRINT ""
            PRINT "Exit_Procedure:"
            PRINT "    Exit Function"
            PRINT "Error_Handler:"
            PRINT "    MsgBox pcMsgTtl & "" ERROR:"" & Err.Number & "" "" & Err.Description, vbCritical, ""Error Messages"""
            PRINT "    Resume Exit_Procedure"
            PRINT "End Function"
            PRINT "'==================================================================================================="
        --------------------------------------------------------------------------------------------
        END
        CONTINUE
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- XXX = Output code description
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3 Trn Idn Erm
        EXEC ut_zzVBX XXX,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,'' ,0  ,0  ,0
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecXXX) BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT @LinSgl
        PRINT "Strip Time From Date:  CAST(CONVERT(varchar(10),GETDATE(),101) AS datetime)"
        PRINT @LinSgl
    ------------------------------------------------------------------------------------------------


    ------------------------------------------------------------------------------------------------
    -- XXX = Output code description
    /*----------------------------------------------------------------------------------------------
        --   ut_zzVBX Oup Stx                    Lft Spc Ttl Bat Tx1 Tx2 Tx3 Trn Idn Erm
        EXEC ut_zzVBX XXX,''                    ,0  ,0  ,0  ,0  ,'' ,'' ,'' ,0  ,0  ,0
    ----------------------------------------------------------------------------------------------*/
    END ELSE IF @BldCOD IN (@SecXXX) BEGIN
    ------------------------------------------------------------------------------------------------
        PRINT @LinSgl
        PRINT "-- XXX = Output code description - IS UNDER CONSTRUCTION!"
        PRINT @LinSgl
    ------------------------------------------------------------------------------------------------
 
 
    ------------------------------------------------------------------------------------------------
    -- Test output code
    ------------------------------------------------------------------------------------------------
    END ELSE IF @BldCOD IN (@SecZZZ) BEGIN
        SET @SecZZZ = @SecZZZ
    ------------------------------------------------------------------------------------------------
    -- Invalid output code
    ------------------------------------------------------------------------------------------------
    END ELSE BEGIN
        PRINT @LinBng
        PRINT @CurUSP+":  Invalid @BldCOD ("+@BldCOD+")"
        PRINT @LinBng
    END
    ------------------------------------------------------------------------------------------------
 
 
    --##############################################################################################
    END  -- Output sections loop
    --##############################################################################################
 
END
GO
 
/*--(LSP)-------------------------------------------------------------------------------------------
 
    --  (Oup: PHL SIG UTL URP LSP PVL)
 
    --   ut_zzUTL Soj        Oup Dbg Obj Dsc                          Dsp Par Cod Exm Tcd
    EXEC ut_zzUTL ut_zzVBJ,UTL,1  ,'' ,'Build core VBA code logic statements','' ,"
    @BldLST varchar(2000) = '',             -- Build code list (comma delimited; see below)
    @InpObj sysname       = '',             -- Input object name
    @OupFmt varchar(3)    = '',             -- Output format (SEL,INS,etc)
    @OupObj sysname       = '',             -- Output object name
    @OupDsc sysname       = '',             -- Output object description
    @DspObj sysname       = '',             -- Display object (replaces @InpObj)
    @DefTyp varchar(11)   = '',             -- Default object type code
    @SqlExc tinyint       = 0,              -- Execute the dynamic SQL statement
    @LftMrg smallint      = 1,              -- Increase left margin (4x)
    @IncSpc tinyint       = 2,              -- Include space(s) before the header
    @IncTtl tinyint       = 1,              -- Include code segment titles
    @IncHdr tinyint       = 1,              -- Include header lines/text
    @IncTpl tinyint       = 0,              -- Include templates
    @IncMsg tinyint       = 0,              -- Include information message
    @IncDrp tinyint       = 0,              -- Include drop statement
    @IncAdd tinyint       = 1,              -- Include add statement
    @IncBat tinyint       = 1,              -- Include batch GO statement
    @IncDat tinyint       = 1,              -- Include data insert statements
    @SelStm varchar(100)  = '',             -- SELECT statement (DISTINCT, TOP, etc)
    @SetLst varchar(2000) = '',             -- SET Column = Value list (colon delimited)
    @JnnLst varchar(2000) = '',             -- JOIN list (colon delimited)
    @WhrLst varchar(2000) = '',             -- WHERE list (colon delimited)
    @GbyLst varchar(2000) = '',             -- GROUP BY list (colon delimited)
    @HavLst varchar(2000) = '',             -- HAVING list (colon delimited)
    @ObyLst varchar(2000) = '',             -- ORDER BY list (comma delimited)
    @LkpLst varchar(2000) = '',             -- Lookup parameters (comma delimited list)
    @StdTx1 varchar(8000) = '',             -- Miscellaneous text value
    @StdTx2 varchar(8000) = '',             -- Miscellaneous text value
    @StdTx3 varchar(8000) = '',             -- Miscellaneous text value
    @IncTrn tinyint       = 0,              -- Include transaction logic
    @IncIdn tinyint       = 1,              -- Include identity column logic
    @IncDsb tinyint       = NULL,           -- Include record disabled columns
    @IncDlt tinyint       = NULL,           -- Include record delflag columns
    @IncLok tinyint       = NULL,           -- Include record locking columns
    @IncAud tinyint       = NULL,           -- Include record auditing columns
    @IncHst tinyint       = NULL,           -- Include record history columns
    @IncMod tinyint       = NULL            -- Include record modified columns
    ","
        PRINT '    PSH    = Push margin 4 spaces right'
        PRINT '    PUL    = Pull margin 4 spaces left'
        PRINT '    LM0    = Set left margin to zero'
        PRINT '    LM1    = Set left margin to one'
        PRINT '    LM2    = Set left margin to two'
        PRINT '    RWP    = Set report width Portrait'
        PRINT '    RWL    = Set report width Landscape'
        PRINT ''
        PRINT '    LSG    = Set lines for single lines'
        PRINT '    LDB    = Set lines for double lines'
        PRINT '    LPD    = Set lines for pound  lines'
        PRINT '    HSG    = Set header for single lines'
        PRINT '    HDB    = Set header for double lines'
        PRINT '    HPD    = Set header for pound  lines'
        PRINT ''
        PRINT '    SLN    = Print single line'
        PRINT '    DLN    = Print double line'
        PRINT '    ALN    = Print asterick line'
        PRINT '    PLN    = Print pound line'
        PRINT '    MLN    = Print ampersand line'
        PRINT '    TLN    = Print tilde line'
        PRINT ''
        PRINT '    LN0    = Print current LinSpc'
        PRINT '    LN1    = Print empty lines (1)'
        PRINT '    LN2    = Print empty lines (2)'
        PRINT '    PLP    = Pound line (prefixed)'
        PRINT '    ALP    = AtSign line (prefixed)'
        PRINT '    PHB    = Print begin header (previously set)'
        PRINT '    PHE    = Print end header (previously set)'
        PRINT ''
        PRINT '    T0N    = Set Title: Space 0 Lines N'
        PRINT '    T0Y    = Set Title: Space 0 Lines Y'
        PRINT '    T1N    = Set Title: Space 1 Lines N'
        PRINT '    T1Y    = Set Title: Space 1 Lines Y'
        PRINT '    T2N    = Set Title: Space 2 Lines N'
        PRINT '    T2Y    = Set Title: Space 2 Lines Y'
        PRINT ''
        PRINT '    IST    = Toggle   Include space(s) before the header'
        PRINT '    ISR    = Reset    Include space(s) before the header'
        PRINT '    IS0    = Set OFF  Include space(s) before the header'
        PRINT '    IS1    = Set Opt1 Include space(s) before the header'
        PRINT '    IS2    = Set Opt2 Include space(s) before the header'
        PRINT ''
        PRINT '    ITT    = Toggle   Include code segment titles'
        PRINT '    ITR    = Reset    Include code segment titles'
        PRINT '    IT0    = Set OFF  Include code segment titles'
        PRINT '    IT1    = Set Opt1 Include code segment titles'
        PRINT '    IT2    = Set Opt2 Include code segment titles'
        PRINT '    IT3    = Set Opt3 Include code segment titles'
        PRINT ''
        PRINT '    IPT    = Toggle   Include templates'
        PRINT '    IPR    = Reset    Include templates'
        PRINT '    IP0    = Set OFF  Include templates'
        PRINT '    IP1    = Set Opt1 Include templates'
        PRINT '    IP2    = Set Opt2 Include templates'
        PRINT ''
        PRINT '    IGT    = Toggle   Include debug logic'
        PRINT '    IGR    = Reset    Include debug logic'
        PRINT '    IG0    = Disable  Include debug logic'
        PRINT '    IG1    = Enable   Include debug logic'
        PRINT ''
        PRINT '    IFT    = Toggle   Include information message'
        PRINT '    IFR    = Reset    Include information message'
        PRINT '    IF0    = Disable  Include information message'
        PRINT '    IF1    = Enable   Include information message'
        PRINT ''
        PRINT '    IQT    = Toggle   Include error message'
        PRINT '    IQR    = Reset    Include error message'
        PRINT '    IQ0    = Disable  Include error message'
        PRINT '    IQ1    = Enable   Include error message'
        PRINT ''
        PRINT '    IVT    = Toggle   Include separator line between objects'
        PRINT '    IVR    = Reset    Include separator line between objects'
        PRINT '    IV0    = Disable  Include separator line between objects'
        PRINT '    IV1    = Enable   Include separator line between objects'
        PRINT ''
        PRINT '    IDT    = Toggle   Include drop statement'
        PRINT '    IDR    = Reset    Include drop statement'
        PRINT '    ID0    = Disable  Include drop statement'
        PRINT '    ID1    = Enable   Include drop statement'
        PRINT ''
        PRINT '    IBT    = Toggle   Include batch GO statement'
        PRINT '    IBR    = Reset    Include batch GO statement'
        PRINT '    IB0    = Disable  Include batch GO statement'
        PRINT '    IB1    = Enable   Include batch GO statement'
        PRINT ''
        PRINT '    INT    = Toggle   Include batch GO statement'
        PRINT '    INR    = Reset    Include batch GO statement'
        PRINT '    IN0    = Disable  Include batch GO statement'
        PRINT '    IN1    = Enable   Include batch GO statement'
        PRINT ''
        PRINT '    IMT    = Toggle   Include permissions statements'
        PRINT '    IMR    = Reset    Include permissions statements'
        PRINT '    IM0    = Disable  Include permissions statements'
        PRINT '    IM1    = Enable   Include permissions statements'
        PRINT ''
        PRINT '    IIT    = Toggle   Include identity column logic'
        PRINT '    IIR    = Reset    Include identity column logic'
        PRINT '    II0    = Disable  Include identity column logic'
        PRINT '    II1    = Enable   Include identity column logic'
        PRINT ''
        PRINT '    IXT    = Toggle   Include record expires columns'
        PRINT '    IXR    = Reset    Include record expires columns'
        PRINT '    IX0    = Disable  Include record expires columns'
        PRINT '    IX1    = Enable   Include record expires columns'
        PRINT ''
        PRINT '    OTX    = Object text (SysComments)'
        PRINT ''
        PRINT '    DVA    = Developer action history'
        PRINT '    VLN    = Set VbaVln length'
        PRINT '    TMB    = Initialize text management objects (Begin)'
        PRINT '    TME    = Initialize text management objects (End)'
        PRINT ''
        PRINT '    DTV    = Declare field column variables'
        PRINT '    ITV    = Initialize field column variables'
        PRINT '    ATV    = Assign standard field variables'
        PRINT ''
        PRINT '    DSV    = Declare statement column variables'
        PRINT '    ISV    = Initialize statement column variables'
        PRINT '    ASV    = Assign statement column variables'
        PRINT ''
        PRINT '    DKV    = Declare primary key column variables'
        PRINT '    IKV    = Initialize primary key column variables'
        PRINT '    AKV    = Assign primary key column variables'
        PRINT ''
        PRINT '    DGV    = Declare function parameter variables'
        PRINT '    IGV    = Initialize parameter variables'
        PRINT '    AGV    = Assign parameter variables'
        PRINT ''
        PRINT '    IRV    = Initialize recordset variables'
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
        PRINT '    RUNUSX = Run SProc commands'
        PRINT ''
        PRINT '    SQLSTM = Build the full SQL statement'
        PRINT ''
        PRINT '    FRMHDR = Header form'
        PRINT '    FRMDTL = Detail form'
        PRINT '    FRMLST = List form'
        PRINT ''
        PRINT '    POPADD = PopAdd form'
        PRINT '    POPUPD = PopUpd form'
        PRINT ''
        PRINT '    RECINS = Insert record function'
        PRINT '    RECUPD = Update record function'
    ","
        PRINT '    --   ut_zzVBJ Soj     Oup Fmt Obj Dsc Dsp Oup Sqx Lft Spc Ttl Hdr Tpl Msg Drp Add Bat Dat Stm Set Jnn Whr Gby Hav Oby Lkp Tx1 Tx2 Tx3 Trn Idn Dsb  Dlt  Lok  Aud  Hst  Mod'
        PRINT '    EXEC ut_zzVBJ ''''     ,ZZZ,   ,'''' ,'''' ,'''' ,'''' ,0  ,1  ,2  ,1  ,1  ,0  ,0  ,0  ,1  ,1  ,0  ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,'''' ,0  ,1  ,NULL,NULL,NULL,NULL,NULL,NULL'
        PRINT '    '
        PRINT '    --   ut_zzVBJ Soj     Oup     Fmt     Obj     Dsc     Dsp     Oup     Sqx     Lft     Spc     Ttl     Hdr     Tpl     Msg     Drp     Add     Bat     Dat     Stm     Set     Jnn     Whr     Gby     Hav     Oby     Lkp     Tx1     Tx2     Tx3     Trn     Idn     Dsb     Dlt     Lok     Aud     Hst     Mod'
        PRINT '    EXEC ut_zzVBJ @InpObj,@BldLST,@OupFmt,@OupObj,@OupDsc,@DspObj,@DefTyp,@SqlExc,@LftMrg,@IncSpc,@IncTtl,@IncHdr,@IncTpl,@IncMsg,@IncDrp,@IncAdd,@IncBat,@IncDat,@SelStm,@SetLst,@JnnLst,@WhrLst,@GbyLst,@HavLst,@ObyLst,@LkpLst,@StdTx1,@StdTx2,@StdTx3,@IncTrn,@IncIdn,@IncDsb,@IncDlt,@IncLok,@IncAud,@IncHst,@IncMod'
    ","
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzVBJ vba_TblDfn,RUNUSX    -- Run SProc process
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzVBJ vba_TblDfn,SQLSTM    -- Build the full SQL statement
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzVBJ vba_TblDfn,FRMHDR    -- Header form
    EXEC ut_zzVBJ vba_TblDfn,FRMDTL    -- Detail form
    EXEC ut_zzVBJ vba_TblDfn,FRMLST    -- List form
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzVBJ vba_TblDfn,POPADD    -- PopAdd form
    EXEC ut_zzVBJ vba_TblDfn,POPUPD    -- PopUpd form
    ------------------------------------------------------------------------------------------------
    EXEC ut_zzVBJ vba_TblDfn,RECINS    -- Insert record function
    EXEC ut_zzVBJ vba_TblDfn,RECUPD    -- Update record function
    ------------------------------------------------------------------------------------------------
    "
 
--------------------------------------------------------------------------------------------------*/
