# utl-altair-slc-reading-and-writing-OpenDocument-spreadsheets-using-r-and-python
Altair slc reading and writing OpenDocument spreadsheets using r and python
    %let pgm=utl-altair-slc-reading-and-writing-OpenDocument-spreadsheets-using-r-and-python;

    %stop_submission;

    Altair slc reading and writing OpenDocument spreadsheets using r and python

    Too long to post here, see github

     CONTENTS

      1 r create an OpenDocument spreadsheet  
      2 r opendocument spreadsheet subset for males  
      3 python read OpenDocument
      4 python write OpenDocument spreadsheet

    github
    https://tinyurl.com/mwp5yutj
    https://github.com/rogerjdeangelis/utl-altair-slc-reading-and-writing-OpenDocument-spreadsheets-using-r-and-python

    community.altair
    https://tinyurl.com/a42546kf
    https://community.altair.com/discussion/60717/altair-slc-can-read-excel-spreadsheets-but-can-it-read-opendocument-spreadsheets-ods?tab=accepted&utm_source=community-search&utm_medium=organic-search&utm_term=slc

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    &_init_;
    options
     validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.have;
      input
        name$
        sex$ age;
    cards4;
    Alfred  M 14
    Alice   F 13
    Barbara F 13
    Carol   F 14
    Henry   M 14
    James   M 12
    ;;;;
    run;quit;

    /*                              _                               ____                                        _
    / |  _ __    ___ _ __ ___  __ _| |_ ___   ___  _ __   ___ _ __ |  _ \  ___   ___ _   _ _ __ ___   ___ _ __ | |_
    | | | `__|  / __| `__/ _ \/ _` | __/ _ \ / _ \| `_ \ / _ \ `_ \| | | |/ _ \ / __| | | | `_ ` _ \ / _ \ `_ \| __|
    | | | |    | (__| | |  __/ (_| | ||  __/| (_) | |_) |  __/ | | | |_| | (_) | (__| |_| | | | | | |  __/ | | | |_
    |_| |_|     \___|_|  \___|\__,_|\__\___| \___/| .__/ \___|_| |_|____/ \___/ \___|\__,_|_| |_| |_|\___|_| |_|\__|
                                                  |_|
    */

    &_init_;
    proc r;
    export data=sd1.class r=class;
    submit;
    library(readODS)
    write_ods(class
     ,"d:/ods/class.ods")
    endsubmit;
    import data=
    run;quit;

    OUTPUT
    ------

    class.ods -- LibreOffice Calc

    -----------------------+
    | A1| fx    |DAYNUM    |
    ------------------------------------
    [_] |    A     |    B    |    C    |
    ------------------------------------
     1  | NAME     |   SEX   |   AGE   |
     -- |----------+---------+---------+
     2  |  Alfred  | M       | 14      |
     -- |----------+---------+---------+
     3  |  Alice   | F       | 13      |
     -- |----------+---------+---------+
     4  |  Barbara | F       | 13      |
     -- |----------+---------+---------+
     5  |  Carol   | F       | 14      |
     -- |----------+---------+---------+
     6  |  Henry   | M       | 14      |
     -- |----------+---------+---------+
     7  |  James   | M       | 12      |
     -- |----------+---------+---------+
    [CLASS]

    /*___                              _             _              _
    |___ \   _ __   _ __ ___  __ _  __| |  ___ _   _| |__  ___  ___| |_
      __) | | `__| | `__/ _ \/ _` |/ _` | / __| | | | `_ \/ __|/ _ \ __|
     / __/  | |    | | |  __/ (_| | (_| | \__ \ |_| | |_) \__ \  __/ |_
    |_____| |_|    |_|  \___|\__,_|\__,_| |___/\__,_|_.__/|___/\___|\__|

     read the LibreOffice Calc and subset using sql
    */

    &_init_;
    proc r;
    export data=sd1.class r=class;
    submit;
    library(readODS)
    library(sqldf)
    options(sqldf.dll = "d:/dll/sqlean.dll")
    class<-read_ods("d:/ods/class.ods")
    males<-sqldf('
      select
        *
      from
        class
      where
        sex="M"
    ')
    males
    endsubmit;
    export data=males r=males;
    run;quit;

    proc print data=males;
    run;quit;

    OUTPUT
    ======

    Altair SLC

        NAME SEX AGE
    1 Alfred   M  14
    2  Henry   M  14
    3  James   M  12

    /*____               _   _                                     _                            ____                                        _
    |___ /   _ __  _   _| |_| |__   ___  _ _    _ __ ___  __ _  __| | ___  _ __   ___ _ __ ___ |  _ \  ___   ___ _   _ _ __ ___   ___ _ __ | |_
      |_ \  | `_ \| | | | __| `_ \ / _ \| `_ \ | `__/ _ \/ _` |/ _` |/ _ \| `_ \ / _ \ `_ ` _ \| | | |/ _ \ / __| | | | `_ ` _ \ / _ \ `_ \| __|
     ___) | | |_) | |_| | |_| | | | (_) | | | || | |  __/ (_| | (_| | (_) | |_) |  __/ | | | | | |_| | (_) | (__| |_| | | | | | |  __/ | | | |_
    |____/  | .__/ \__, |\__|_| |_|\___/|_| |_||_|  \___|\__,_|\__,_|\___/| .__/ \___|_| |_| |_|____/ \___/ \___|\__,_|_| |_| |_|\___|_| |_|\__|
            |_|    |___/                                                  |_|
    3 python write opemDocument spreadsheet
    */

    &_init_;
    OPTIONS NOERRORABEND;
    options set=PYTHONHOME "D:\python310";
    proc python;
    submit;
    exec(open('c:/wpsoto/fn_pythonx.py').read());
    class = pd.read_excel('d:/ods/class.ods', engine='odf')
    print(class)
    males=pdsql('''
      select
        *
      from
        class
      where
        sex="M"
    ''')
    print(males)
    pr.write_rds('d:/rds/males.rds',males)
    endsubmit;
    ;quit;

    /*---- CONVERT RDS FILE TO A SAS TABLE ----*/

    &_init_;
    options noerrorabend;
    options set=RHOME "D:\d451";
    %wps_py2sastable(
       inp=d:/rds/males.rds
      ,out=males );

    proc print data=work.males;
    run;quit;

    OUTPUT
    ======

    Altair SLC

    Obs     NAME     SEX    AGE

     1     Alfred     M      14
     2     Henry      M      14
     3     James      M      12


    /*  _                 _   _                                 _ _                                   ____                                        _
    | || |    _ __  _   _| |_| |__   ___  _ __   __      ___ __(_) |_ ___   ___  _ __   ___ _ __ ___ |  _ \  ___   ___ _   _ _ __ ___   ___ _ __ | |_
    | || |_  | `_ \| | | | __| `_ \ / _ \| `_ \  \ \ /\ / / `__| | __/ _ \ / _ \| `_ \ / _ \ `_ ` _ \| | | |/ _ \ / __| | | | `_ ` _ \ / _ \ `_ \| __|
    |__   _| | |_) | |_| | |_| | | | (_) | | | |  \ V  V /| |  | | ||  __/| (_) | |_) |  __/ | | | | | |_| | (_) | (__| |_| | | | | | |  __/ | | | |_
       |_|   | .__/ \__, |\__|_| |_|\___/|_| |_|   \_/\_/ |_|  |_|\__\___| \___/| .__/ \___|_| |_| |_|____/ \___/ \___|\__,_|_| |_| |_|\___|_| |_|\__|
             |_|    |___/                                                       |_|
    wrte openDocument LibreOffice spreadsheet
    */

    &_init_;
    OPTIONS NOERRORABEND;
    options set=PYTHONHOME "D:\python310";
    proc python;
    submit;
    exec(open('c:/wpsoto/fn_pythonx.py').read());
    class,meta = ps.read_sas7bdat('d:/sd1/have.sas7bdat')
    import install odfpy
    with pd.ExcelWriter('d:/osd/class.ods', engine='odf') as writer:
        df.to_excel(writer, index=False, sheet_name='class')
    endsubmit;
    run;quit;

    OUTPUT
    ------

    class.ods -- LibreOffice Calc

    -----------------------+
    | A1| fx    |DAYNUM    |
    ------------------------------------
    [_] |    A     |    B    |    C    |
    ------------------------------------
     1  | NAME     |   SEX   |   AGE   |
     -- |----------+---------+---------+
     2  |  Alfred  | M       | 14      |
     -- |----------+---------+---------+
     3  |  Alice   | F       | 13      |
     -- |----------+---------+---------+
     4  |  Barbara | F       | 13      |
     -- |----------+---------+---------+
     5  |  Carol   | F       | 14      |
     -- |----------+---------+---------+
     6  |  Henry   | M       | 14      |
     -- |----------+---------+---------+
     7  |  James   | M       | 12      |
     -- |----------+---------+---------+
    [CLASS]

    /*___        _
    | ___|   ___| | ___   _ __ ___ _ __   ___  ___
    |___ \  / __| |/ __| | `__/ _ \ `_ \ / _ \/ __|
     ___) | \__ \ | (__  | | |  __/ |_) | (_) \__ \
    |____/  |___/_|\___| |_|  \___| .__/ \___/|___/
                                  |_|
    */

    sLC repos
    -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    https://github.com/rogerjdeangelis/setup-personal-edition-altair-slc-eclipse-workspace-config-sasautos-sasuser-saswork-autoexec
    https://github.com/rogerjdeangelis/utl-altair-slc-to-fill-gaps-in-proc-sql-select-third-place-in-the-daily-double-r-python-solutions
    https://github.com/rogerjdeangelis/utl-calling-python-from-personal-altair-slc-and-integrating-python-with-sql
    https://github.com/rogerjdeangelis/utl-calling-r-from-personal-altair-slc-and-integrating-r-with-sql
    https://github.com/rogerjdeangelis/utl-dropping-down-to-powershell-from-personal-altair-slc
    https://github.com/rogerjdeangelis/utl-how-to-create-a-sas-dataset-from-python-panda-dataframe-using-the-personal-altair-slc
    https://github.com/rogerjdeangelis/utl-how-to-create-a-sas-dataset-from-r-dataframe-using-the-personal-altair-slc

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
