namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public enum FtabValues : ushort
    {
        COUNT = 0x0000,
        IF = 0x0001,
        ISNA = 0x0002,
        ISERROR = 0x0003,
        SUM  = 0x0004,
        AVERAGE = 0x0005,
        MIN = 0x0006,
        MAX = 0x0007 ,
        ROW = 0x0008,
        COLUMN = 0x0009,
        NA = 0x000A,
        NPV = 0x000B,
        STDEV = 0x000C,
        DOLLAR = 0x000D,
        FIXED = 0x000E,
        SIN = 0x000F,
        COS = 0x0010,

        TAN = 0x0011,
        ATAN = 0x0012, 
        PI = 0x0013, 
        SQRT = 0x0014,
        EXP = 0x0015,
        LN = 0x0016,
        LOG10 = 0x0017,
        ABS = 0x0018,
        INT = 0x0019,
        SIGN = 0x001A,
        ROUND = 0x001B,
        LOOKUP = 0x001C,
        INDEX = 0x001D,
        REPT = 0x001E,
        MID = 0x001F,
        LEN =0x0020,
        VALUE = 0x0021,
        TRUE = 0x0022,
        FALSE = 0x0023,
        AND = 0x0024,
        OR = 0x0025,
        NOT = 0x0026,

        MOD = 0x0027,
        DCOUNT=0x0028,
        DSUM=0x0029,
        DAVERAGE=0x002A,
        DMIN=0x002B,
        DMAX=0x002C,
        DSTDEV=0x002D,
        VAR=0x002E,
        DVAR=0x002F,
        TEXT=0x0030,
        LINEST=0x0031,
        TREND=0x0032,
        LOGEST=0x0033,
        GROWTH=0x0034,
        GOTO=0x0035,
        HALT=0x0036,
        RETURN=0x0037,
        PV=0x0038,
        FV=0x0039,
        NPER=0x003A,
        PMT=0x003B,

        RATE=0x003C,
        MIRR=0x003D,
        IRR=0x003E,
        RAND=0x003F,
        MATCH=0x0040,
        DATE=0x0041,
        TIME=0x0042,
        DAY=0x0043,
        MONTH=0x0044,
        YEAR=0x0045,
        WEEKDAY=0x0046,
        HOUR=0x0047,
        MINUTE=0x0048,
        SECOND=0x0049,
        NOW=0x004A,
        AREAS=0x004B,
        ROWS=0x004C,
        COLUMNS=0x004D,
        OFFSET=0x004E,
        ABSREF=0x004F,
        RELREF=0x0050,
        ARGUMENT=0x0051,

        SEARCH =0x0052,
        TRANSPOSE=0x0053,
        ERROR=0x0054,
        STEP=0x0055,
        TYPE=0x0056,
        ECHO=0x0057,
        SET_NAME=0x0058,
        CALLER=0x0059,
        DEREF=0x005A,
        WINDOWS=0x005B,
        SERIES=0x005C,
        DOCUMENTS=0x005D,
        ACTIVE_CELL=0x005E,
        SELECTION=0x005F,
        RESULT=0x0060,
        ATAN2=0x0061,
        ASIN=0x0062,
        ACOS=0x0063,
        CHOOSE=0x0064,
        HLOOKUP=0x0065,
        VLOOKUP=0x0066,

        LINKS=0x0067,       
        INPUT=0x0068,      
        ISREF=0x0069,       
        GET_FORMULA=0x006A ,
        GET_NAME=0x006B,    
        SET_VALUE=0x006C,   
        LOG=0x006D,   
        EXEC=0x006E,  
        CHAR=0x006F, 
        LOWER=0x0070,
        UPPER=0x0071,       
        PROPER=0x0072,      
        LEFT=0x0073,
        RIGHT=0x0074,
        EXACT=0x0075,
        TRIM=0x0076,
        REPLACE=0x0077, 
        SUBSTITUTE=0x0078,
        CODE=0x0079,  
        NAMES=0x007A, 
        DIRECTORY=0x007B,
        FIND=0x007C, 

        CELL=0x007D,             
        ISERR=0x007E,            
        ISTEXT=0x007F,           
        ISNUMBER=0x0080,         
        ISBLANK=0x0081,          
        T=0x0082,                
        N=0x0083,                
        FOPEN=0x0084,            
        FCLOSE=0x0085,           
        FSIZE=0x0086,            
        FREADLN=0x0087,          
        FREAD=0x0088,            
        FWRITELN=0x0089,         
        FWRITE=0x008A,           
        FPOS=0x008B,             
        DATEVALUE=0x008C,        
        TIMEVALUE=0x008D,        
        SLN=0x008E,              
        SYD=0x008F,              
        DDB=0x0090,              
        GET_DEF=0x0091,          

         REFTEXT=0x0092,          
         TEXTREF=0x0093,          
         INDIRECT=0x0094,         
         REGISTER=0x0095,         
         CALL=0x0096 ,            
         ADD_BAR=0x0097,          
         ADD_MENU=0x0098,         
         ADD_COMMAND=0x0099,      
         ENABLE_COMMAND=0x009A,   
         CHECK_COMMAND=0x009B,    
         RENAME_COMMAND=0x009C,   
         SHOW_BAR=0x009D,         
         DELETE_MENU=0x009E,      
         DELETE_COMMAND=0x009F,   
         GET_CHART_ITEM=0x00A0,   
         DIALOG_BOX=0x00A1,       
         CLEAN=0x00A2,            
         MDETERM=0x00A3,          
         MINVERSE=0x00A4,         
         MMULT=0x00A5,            
         FILES=0x00A6,            
         IPMT=0x00A7,             
         
        PPMT=0x00A8,             
        COUNTA=0x00A9,           
        CANCEL_KEY=0x00AA,       
        FOR=0x00AB,              
        WHILE=0x00AC,            
        BREAK=0x00AD,            
        NEXT=0x00AE,             
        INITIATE=0x00AF,         
        REQUEST=0x00B0,          
        POKE=0x00B1,             
        EXECUTE=0x00B2,          
        TERMINATE=0x00B3,        
        RESTART=0x00B4,          
        HELP=0x00B5,             
        GET_BAR=0x00B6,          
        PRODUCT=0x00B7,          
        FACT=0x00B8,             
        GET_CELL=0x00B9,         
        GET_WORKSPACE=0x00BA,    
        GET_WINDOW=0x00BB,       
        GET_DOCUMENT =0x00BC,    

        DPRODUCT=0x00BD,   
        ISNONTEXT=0x00BE,  
        GET_NOTE=0x00BF,   
        NOTE=0x00C0,       
        STDEVP=0x00C1,     
        VARP=0x00C2,       
        DSTDEVP=0x00C3,    
        DVARP=0x00C4,      
        TRUNC=0x00C5 ,     
        ISLOGICAL=0x00C6,  
        DCOUNTA=0x00C7,    
        DELETE_BAR=0x00C8, 
        UNREGISTER=0x00C9, 
        USDOLLAR=0x00CC,   
        FINDB=0x00CD,      
        SEARCHB=0x00CE,    
        REPLACEB=0x00CF,   
        LEFTB=0x00D0,      
        RIGHTB=0x00D1,     
        MIDB=0x00D2,       
        LENB=0x00D3,       
        ROUNDUP=0x00D4,    

        ROUNDDOWN=0x00D5, 
        ASC=0x00D6,       
        DBCS=0x00D7,      
        RANK=0x00D8,      
        ADDRESS=0x00DB,   
        DAYS360=0x00DC,   
        TODAY=0x00DD,     
        VDB=0x00DE,       
        ELSE=0x00DF,      
        ELSE_IF=0x00E0,   
        END_IF=0x00E1,    
        FOR_CELL=0x00E2,  
        MEDIAN=0x00E3,    
        SUMPRODUCT=0x00E4,
        SINH=0x00E5,      
        COSH=0x00E6,      
        TANH=0x00E7,      
        ASINH=0x00E8,     
        ACOSH=0x00E9 ,    
        ATANH=0x00EA,     
        DGET=0x00EB,

        CREATE_OBJECT = 0x00EC,
        VOLATILE = 0x00ED,
        LAST_ERROR = 0x00EE,
        CUSTOM_UNDO = 0x00EF,
        CUSTOM_REPEAT = 0x00F0,
        FORMULA_CONVERT = 0x00F1,
        GET_LINK_INFO = 0x00F2,
        TEXT_BOX = 0x00F3,
        INFO = 0x00F4,
        GROUP = 0x00F5,
        GET_OBJECT = 0x00F6,
        DB = 0x00F7,
        PAUSE = 0x00F8,
        RESUME = 0x00FB,
        FREQUENCY = 0x00FC,
        ADD_TOOLBAR = 0x00FD,
        DELETE_TOOLBAR = 0x00FE,
        RESET_TOOLBAR = 0x0100,
        EVALUATE = 0x0101,
        GET_TOOLBAR = 0x0102,

        GET_TOOL = 0x0103,
        SPELLING_CHECK = 0x0104,
        ERROR_TYPE = 0x0105,
        APP_TITLE = 0x0106,
        WINDOW_TITLE = 0x0107,
        SAVE_TOOLBAR = 0x0108,
        ENABLE_TOOL = 0x0109,
        PRESS_TOOL = 0x010A,
        REGISTER_ID = 0x010B,
        GET_WORKBOOK = 0x010C,
        AVEDEV = 0x010D,
        BETADIST = 0x010E,
        GAMMALN = 0x010F,
        BETAINV = 0x0110,
        BINOMDIST = 0x0111,
        CHIDIST = 0x0112,
        CHIINV = 0x0113,
        COMBIN = 0x0114,
        CONFIDENCE = 0x0115,
        CRITBINOM = 0x0116,
        EVEN = 0x0117,
        EXPONDIST = 0x0118,

        FDIST = 0x0119,
        FINV = 0x011A,
        FISHER = 0x011B,
        FISHERINV = 0x011C,
        FLOOR = 0x011D,
        GAMMADIST = 0x011E,
        GAMMAINV = 0x011F,
        CEILING = 0x0120,
        HYPGEOMDIST = 0x0121,
        LOGNORMDIST = 0x0122,
        LOGINV = 0x0123,
        NEGBINOMDIST = 0x0124,
        NORMDIST = 0x0125,
        NORMSDIST = 0x0126,
        NORMINV = 0x0127,
        NORMSINV = 0x0128,
        STANDARDIZE = 0x0129,
        ODD = 0x012A,
        PERMUT = 0x012B,
        POISSON = 0x012C,
        TDIST = 0x012D,

        WEIBULL = 0x012E,
        SUMXMY2 = 0x012F,
        SUMX2MY2 = 0x0130,
        SUMX2PY2 = 0x0131,
        CHITEST = 0x0132,
        CORREL = 0x0133,
        COVAR = 0x0134,
        FORECAST = 0x0135,
        FTEST = 0x0136,
        INTERCEPT = 0x0137,
        PEARSON = 0x0138,
        RSQ = 0x0139,
        STEYX = 0x013A,
        SLOPE = 0x013B,
        TTEST = 0x013C,
        PROB = 0x013D,
        DEVSQ = 0x013E,
        GEOMEAN = 0x013F,
        HARMEAN = 0x0140,
        SUMSQ = 0x0141,
        KURT = 0x0142,
        SKEW = 0x0143,  

        ZTEST=0x0144,             
        LARGE=0x0145,             
        SMALL=0x0146,
        QUARTILE=0x0147,             
        PERCENTILE=0x0148,             
        PERCENTRANK=0x0149,             
        MODE=0x014A,             
        TRIMMEAN=0x014B,             
        TINV=0x014C,             
        MOVIE_COMMAND=0x014E,            
        GET_MOVIE=0x014F,           
        CONCATENATE=0x0150,          
        POWER=0x0151,         
        PIVOT_ADD_DATA=0x0152,        
        GET_PIVOT_TABLE=0x0153,       
        GET_PIVOT_FIELD=0x0154,      
        GET_PIVOT_ITEM=0x0155,     
        RADIANS=0x0156,    
        DEGREES=0x0157,   
        SUBTOTAL=0x0158,  
        SUMIF=0x0159,

        COUNTIF = 0x015A,
        COUNTBLANK = 0x015B,
        SCENARIO_GET = 0x015C,
        OPTIONS_LISTS_GET = 0x015D,
        ISPMT = 0x015E,
        DATEDIF = 0x015F,
        DATESTRING = 0x0160,
        NUMBERSTRING = 0x0161,
        ROMAN = 0x0162,
        OPEN_DIALOG = 0x0163,
        SAVE_DIALOG = 0x0164,
        VIEW_GET = 0x0165,
        GETPIVOTDATA = 0x0166,
        HYPERLINK = 0x0167,
        PHONETIC = 0x0168,
        AVERAGEA = 0x0169,
        MAXA = 0x016A,
        MINA = 0x016B,
        STDEVPA = 0x016C,
        VARPA = 0x016D,
        STDEVA = 0x016E,
        VARA = 0x016F,      

   }

}