# utl-importing-multiple-excel-worksheets-without-access-to-pc-files
Importing multiple excel worksheets without access to pc files  
    %let pgm=utl-importing-multiple-excel-worksheets-without-access-to-pc-files ;

    Importing multiple excel worksheets without access to pc files

    github
    http://tinyurl.com/49drmj7z
    https://github.com/rogerjdeangelis/utl-importing-multiple-excel-worksheets-without-access-to-pc-files

    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|
    */

    /**************************************************************************************************************************/
    /*   _                   _                       _                                                                        */
    /*  (_)_ __  _ __  _   _| |_    _____  _____ ___| |                                                                       */
    /*  | | `_ \| `_ \| | | | __|  / _ \ \/ / __/ _ \ |                                                                       */
    /*  | | | | | |_) | |_| | |_  |  __/>  < (_|  __/ |                                                                       */
    /*  |_|_| |_| .__/ \__,_|\__|  \___/_/\_\___\___|_|                                                                       */
    /*          |_|                                                                                                           */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /*                                              d:/xls/threesheets.xlsx                                                   */
    /*                                                                                                                        */
    /*        Sheet All Sexes                |            Sheet Males               |               Sheet Females             */
    /*                                       |                                      |                                         */
    /*    +-------------------------------+  |   +-------------------------------+  |    +-------------------------------+    */
    /*    |     A  |  B | C |  D   |  E   |  |   |     A  |  B | C |  D   |  E   |  |    |     A  |  B | C |  D   |  E   |    */
    /*    +-------------------------------+  |   +-------------------------------+  |    +-------------------------------+    */
    /*  1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|  | 1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|  |  1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*  2 | ALICE  |  F |14 |  69  |112.5 |  | 2 | ALFRED |  M |14 |  69  |112.5 |  |  2 | ALICE  |  F |14 |  69  |112.5 |    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*    ...                                |    ...                               |     ...                                 */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*  19| WILLIAM|  M |15 | 66.5 |112   |  | 10| WILLIAM|  M |15 | 66.5 |112   |  |  9 | WILMA  |  F |15 | 66.5 |112   |    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*                                       |                                      |                                         */
    /*     [ALLSEX}                          |   [MALES]                            |     [FEMALES]                           */
    /*                                                                              |                                         */
    /*               _               _                  _____ _         _       _                                             */
    /*    ___  _   _| |_ _ __  _   _| |_   ___  __ _ __|___  | |__   __| | __ _| |_ ___                                       */
    /*   / _ \| | | | __| `_ \| | | | __| / __|/ _` / __| / /| `_ \ / _` |/ _` | __/ __|                                      */
    /*  | (_) | |_| | |_| |_) | |_| | |_  \__ \ (_| \__ \/ / | |_) | (_| | (_| | |_\__ \                                      */
    /*   \___/ \__,_|\__| .__/ \__,_|\__| |___/\__,_|___/_/  |_.__/ \__,_|\__,_|\__|___/                                      */
    /*                  |_|                                                                                                   */
    /*                                      |                                        |                                        */
    /*                                      |                                        |        SD1.MALES.SAS7BDAT              */
    /*       SD1.ALLSEXES.SAS7BDAT          |      SD1.MALES.SAS7BDAT                |                                        */
    /*                                      |                                        |                                        */
    /*   NAME       SEX AGE HEIGHT WEIGHT   |  NAME     AGE HEIGHT WEIGHT            |    NAME      AGE HEIGHT WEIGHT         */
    /*                                      |                                        |                                        */
    /*   Alfred      M   14  69.0   112.5   |  Alfred    14  69.0   112.5            |    Alice      13  56.5    84.0         */
    /*   Alice       F   13  56.5    84.0   |  Henry     14  63.5   102.5            |    Barbara    13  65.3    98.0         */
    /*   Barbara     F   13  65.3    98.0   |  James     12  57.3    83.0            |    Carol      14  62.8   102.5         */
    /*   Carol       F   14  62.8   102.5   |  Jeffrey   13  62.5    84.0            |    Jane       12  59.8    84.5         */
    /*   ...                                |  ...                                   |                                        */
    /*                                      |                                        |                                        */
    /*  _ __  _ __ ___   ___ ___  ___ ___                                                                                     */
    /* | `_ \| `__/ _ \ / __/ _ \/ __/ __|                                                                                    */
    /* | |_) | | | (_) | (_|  __/\__ \__ \                                                                                    */
    /* | .__/|_|  \___/ \___\___||___/___/                                                                                    */
    /* |_|                                                                                                                    */
    /*                                                                                                                        */
    /*  library("openxlsx");                                                                                                  */
    /*  xlsxFile="d:/xls/dsns.xlsx";                                                                                          */
    /*  allsexes <- read.xlsx(xlsxFile = xlsxFile, sheet = "ALLSEXES", skipEmptyRows = FALSE);                                */
    /*  males    <- read.xlsx(xlsxFile = xlsxFile, sheet = "MALES", skipEmptyRows = FALSE);                                   */
    /*  females  <- read.xlsx(xlsxFile = xlsxFile, sheet = "FEMALES", skipEmptyRows = FALSE);                                 */
    /*  endsubmit;                                                                                                            */
    /*  import data=sd1.allsexes     r=allsexes;                                                                              */
    /*  import data=sd1.males        r=males;                                                                                 */
    /*  import data=sd1.females      r=females;                                                                               */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    %utlfkil(d:/xls/dsns.xlsx);

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
      set sashelp.class;
    run;quit;

    %utl_submit_wps64x('
    libname sd1 "d:/sd1";
    ods excel file="d:/xls/dsns.xlsx";
    ods excel options(sheet_name="ALLSEXES");
    proc print data=sd1.have;
    run;quit;
    ods excel options(sheet_name="MALES" );
    proc print data=sd1.have(where=(sex="M"));
    run;quit;
    ods excel options(sheet_name="FEMALES" );
    proc print data=sd1.have(where=(sex="F"));
    run;quit;
    ods excel close;
    ');

    /**************************************************************************************************************************/
    /*                                                                                                                        */
    /*                                              d:/xls/threesheets.xlsx                                                   */
    /*                                                                                                                        */
    /*        Sheet All Sexes                |            Sheet Males               |               Sheet Females             */
    /*                                       |                                      |                                         */
    /*    +-------------------------------+  |   +-------------------------------+  |    +-------------------------------+    */
    /*    |     A  |  B | C |  D   |  E   |  |   |     A  |  B | C |  D   |  E   |  |    |     A  |  B | C |  D   |  E   |    */
    /*    +-------------------------------+  |   +-------------------------------+  |    +-------------------------------+    */
    /*  1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|  | 1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|  |  1 |  NAME  | SEX|AGE|HEIGHT|WEIGHT|    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*  2 | ALICE  |  F |14 |  69  |112.5 |  | 2 | ALFRED |  M |14 |  69  |112.5 |  |  2 | ALICE  |  F |14 |  69  |112.5 |    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*    ...                                |    ...                               |     ...                                 */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*  19| WILLIAM|  M |15 | 66.5 |112   |  | 10| WILLIAM|  M |15 | 66.5 |112   |  |  9 | WILMA  |  F |15 | 66.5 |112   |    */
    /*    +--------+----+---+------+------+  |   +--------+----+---+------+------+  |    +--------+----+---+------+------+    */
    /*                                       |                                      |                                         */
    /*     [ALLSEX}                          |   [MALES]                            |     [FEMALES]                           */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|
    ;

    proc datasets lib=sd1 nolist mt=data mt=view nodetails;delete allsexes male female; run;quit;

    proc datasets lib=work kill;
    run;quit;

    %symdel imports / nowarn;

    * geet sheetnames and create 100 dataframes;
    %utl_submit_wps64x('
       libname sd1 "d:/sd1";
       proc r;
       submit;
       library("openxlsx");
       xlsxFile="d:/xls/dsns.xlsx";
       allsexes <- read.xlsx(xlsxFile = xlsxFile, sheet = "ALLSEXES", skipEmptyRows = FALSE);
       males    <- read.xlsx(xlsxFile = xlsxFile, sheet = "MALES", skipEmptyRows = FALSE);
       females  <- read.xlsx(xlsxFile = xlsxFile, sheet = "FEMALES", skipEmptyRows = FALSE);
       endsubmit;
       import data=sd1.allsexes     r=allsexes;
       import data=sd1.males        r=males;
       import data=sd1.females      r=females;
       run;quit;
       proc print data=sd1.allsexes;run;quit;
       proc print data=sd1.males   ;run;quit;
       proc print data=sd1.females ;run;quit;
    ');

    /**************************************************************************************************************************/
    /*                                      |                                        |                                        */
    /*                                      |                                        |                                        */
    /*                                      |                                        |        SD1.MALES.SAS7BDAT              */
    /*       SD1.ALLSEXES.SAS7BDAT          |      SD1.MALES.SAS7BDAT                |                                        */
    /*                                      |                                        |                                        */
    /*   NAME       SEX AGE HEIGHT WEIGHT   |  NAME     AGE HEIGHT WEIGHT            |    NAME      AGE HEIGHT WEIGHT         */
    /*                                      |                                        |                                        */
    /*   Alfred      M   14  69.0   112.5   |  Alfred    14  69.0   112.5            |    Alice      13  56.5    84.0         */
    /*   Alice       F   13  56.5    84.0   |  Henry     14  63.5   102.5            |    Barbara    13  65.3    98.0         */
    /*   Barbara     F   13  65.3    98.0   |  James     12  57.3    83.0            |    Carol      14  62.8   102.5         */
    /*   Carol       F   14  62.8   102.5   |  Jeffrey   13  62.5    84.0            |    Jane       12  59.8    84.5         */
    /*   ...                                |  ...                                   |                                        */
    /*                                      |                                        |                                        */
    /**************************************************************************************************************************/

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
