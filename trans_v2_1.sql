create or replace PACKAGE trans
IS
  /* TRANS stands for TRANSformation and TRANSfer 
  ------------------------------------------------------------------------------------------------

  This package is a set of handy functions and procedures for data transformation and transfer, 
  mainly associated with the data preparation for sending with email.


  Igor Perkovic
  Created: June, 2017.
  Last modified: 20.12.2018. 8:56:19

  CHANGE LOG:
  ===================================================================
    20.12.2018. 8:56:19
    -------------------
    ~ Changed default color in predefined table css

    15.6.2018. 8:15:35
    ------------------
    + Added to_char(systimestamp, 'Dy, DD Mon YYYY HH24:MI:SS TZHTZM', 'NLS_DATE_LANGUAGE=ENGLISH') in MAIL for correct Sent date creation

    8.6.2018. 14:11:15
    -----------------------------------------------------------------
    + Add v_pass for password input in MAIL procedure

    8.3.2018. 10:19:59
    -----------------------------------------------------------------
    + Custom headers and footers in CSV functions
    
    ~ Changed DBMS_SQL.DESCRIBE_COLUMNS > DBMS_SQL.DESCRIBE_COLUMNS2
    ~ Changed DBMS_SQL.DESC_TAB > DBMS_SQL.DESC_TAB2

    04.10.2017.
    -----------------------------------------------------------------
    + Added CRLF or CR choice in "query2html" and "ref2html"


    22.09.2017.
    -----------------------------------------------------------------    
    + Added procedure blob2file


    14.09.2017.
    -----------------------------------------------------------------
    For query2csv_file and ref2csv_file
    + added default values for parameters

    For MAIL
    + added 4th CLOB attachment as default


    12.09.2017.
    -----------------------------------------------------------------
    Added 2 procedures
    + query2csv_file
    + ref2csv_file
    for creating csv files. 

    It was already implemented with query2csv and clob2file, but these are a few miliseconds faster in average...
  ---------------------------------------------------------------------------------------------------------------
  */


FUNCTION query2csv(
    /* 
   Igor Perkovic, Created: 11.07.2017. | Last updated: 21.07.2017.

   THE PURPOSE:
   This function translates the SQL query to the CSV format and return the result as a CLOB variable.
   This way I can pass the CLOB to the CLOB2FILE procedure - if I want to save the result as a CSV file, or
   I can send the CLOB result via email with MAIL procedure.
   
   IN SHORT: From SQL query to file or as a mail attachment (in cooperation with P_CLOB2FILE and MAIL procedures)

   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:

    */
    p_sql           IN VARCHAR2,                -- SQL query
    p_type          IN NUMBER   := 0,           -- '1' for mail, '0' for file is default
    p_delimiter     IN VARCHAR2 :=';',          -- Main delimiter
    p_header        IN NUMBER   := 2,           -- 1 = Header ON | 2 = CUSTOM HEADRE ON | 0 = HEADER OFF
    p_footer        IN NUMBER   := 0,           -- 1 = Footer ON  | 0 = Footer OFF
    p_cust_header   IN VARCHAR2 := NULL,        -- Custom header
    p_cust_footer   IN VARCHAR2 := NULL,        -- Custom footer
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD.MM.RRRR') -- Date format
    RETURN CLOB;

FUNCTION ref2csv(
    /* 
   Igor Perkovic, Created: 11.07.2017. | Last modified: 21.07.2017.
   
   THE PURPOSE:
   This function catches the SYS_REFCURSOR from some outer function and transform the function result to the CSV format 
   and return the result as a CLOB variable.
   
   This way I can pass the CLOB to the CLOB2FILE procedure - if I want to save the result as a CSV file, or
   I can send the CLOB result via email using procedure MAIL.
   
   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:    
    */

    p_refcursor IN OUT SYS_REFCURSOR,           -- get the sys_refcursor
    p_type          IN NUMBER   := 0,           -- '1' for mail, '0' for file is default
    p_delimiter     IN VARCHAR2 :=';',          -- Main delimiter
    p_header        IN NUMBER   := 2,           -- 1 = Header ON | 2 = CUSTOM HEADRE ON | 0 = HEADER OFF
    p_footer        IN NUMBER   := 0,           -- 1 = Footer ON  | 0 = Footer OFF
    p_cust_header   IN VARCHAR2 := NULL,        -- Custom header
    p_cust_footer   IN VARCHAR2 := NULL,        -- Custom footer    
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD.MM.RRRR') -- Date format
    RETURN CLOB;

FUNCTION query2html(
    /* 
   Igor Perkovic, Created: 12.07.2017. | Updated: 21.07.2017.

   THE PURPOSE:
   This function translates the SQL query to HTML table and return the result as a CLOB variable.
   This way I can pass the CLOB to the CLOB2FILE procedure - if I want to save the result as a html file, or
   I can send the CLOB result via email with MAIL procedure.

   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:

    */
    p_sql           IN VARCHAR2,                -- SQL query
    p_css           IN NUMBER   :=0,            -- No CSS by default - this is for in-mail sending (clob in mail body embedded). 
                                                -- as a side option, this can be used for CSS predefined styles, like 1,2,3, ...
                                                -- but for now, there is only one CSS in code. 
    p_crlf          IN NUMBER   :=1,            -- "crlf" is set to 1 by default.  0 is "cr" only.
    p_charset       IN VARCHAR2 := 'UTF-8',     -- META CHARSET encoding string                   
    p_tagl          IN VARCHAR2 := '<td>',      -- left tag (universal)
    p_tagr          IN VARCHAR2 := '</td>',     -- right tag (universal)
    p_tagl_v        IN VARCHAR2 := '<td>',      -- left tag for text data fields
    p_tagl_n        IN VARCHAR2 := '<td>',      -- left tag for number data fields
    p_tagl_d        IN VARCHAR2 := '<td>',      -- left tag for dates fields
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD-MM-YYYY',-- Date format
    p_cust_header   IN VARCHAR2 := '',
    p_cust_footer   IN VARCHAR2 := 'Powered by: <b>your_site.com</b>'
    )
    RETURN CLOB;

FUNCTION ref2html(
    /* 
   Igor Perkovic, Created: 12.07.2017. | Updated: 21.07.2017.

   THE PURPOSE:
   This function translates the SYS_REF cursor to HTML table and return the result as a CLOB variable.
   This way I can pass the CLOB to the CLOB2FILE procedure - if I want to save the result as a html file, or
   I can send the CLOB result via email with MAIL procedure.

   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:

    */
    p_refcursor IN OUT SYS_REFCURSOR,           -- get the sys_refcursor
    p_css           IN NUMBER   :=0,            -- No CSS by default - this is for in-mail sending (clob in mail body embedded). 
                                                -- as a side option, this can be used for CSS predefined styles, like 1,2,3, ...
                                                -- but for now, there is only one CSS in code. 
    p_crlf          IN NUMBER   :=1,            -- "crlf" is set to 1 by default.  0 is "cr" only.
    p_charset       IN VARCHAR2 := 'UTF-8',     -- META CHARSET encoding string
    p_tagl          IN VARCHAR2 := '<td>',      -- left tag (universal)
    p_tagr          IN VARCHAR2 := '</td>',     -- right tag (universal)
    p_tagl_v        IN VARCHAR2 := '<td>',      -- left tag for text data fields
    p_tagl_n        IN VARCHAR2 := '<td>',      -- left tag for number data fields
    p_tagl_d        IN VARCHAR2 := '<td>',      -- left tag for dates fields
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD-MM-YYYY',-- Date format
    p_cust_header   IN VARCHAR2 := '',
    p_cust_footer   IN VARCHAR2 := 'Powered by: <b>your_site.com</b>'
    )
    RETURN CLOB;

FUNCTION file2blob(
    p_dir VARCHAR2, 
    p_file_name VARCHAR2) 
    RETURN BLOB;

FUNCTION clob2blob (
    /*  Igor Perkovic, Created 18.7.2017 | Last modified: 21.07.2017
    
        If you want to send any CLOB as attachment through an email function, which does not use the raw format for writting data (and it doesn't hahaha)
        then all specific characters like čćžšđ ČĆŽŠĐ would be transfered as cczsd CCZSD...
    
        Now, with this function we will convert this bad, bad CLOB to good and shiny BLOB with DESIRED ENCODING (mic drop!)
    
        Thanks to Udo,  https://community.oracle.com/thread/2318795 for this idea...IT WORKS!
    */
        in_clob     IN CLOB, 
        in_charset  IN VARCHAR2 DEFAULT 'EE8MSWIN1250') 
    RETURN BLOB;

PROCEDURE clob2file (
    v_clob    IN CLOB,
    v_dir     IN VARCHAR2,
    v_fname   IN VARCHAR2
    );

PROCEDURE blob2file( 
    p_blob  IN BLOB, 
    p_dir   IN VARCHAR2, 
    p_file  IN VARCHAR2
    );

PROCEDURE query2csv_file(
  /* 
     Igor Perkovic, 10.07.2017.
  
     THE PURPOSE:
     This function translates the SQL query to the CSV format and write the result into the CSV file.
  
     OPTIONS:
     You can customize the output a bit through the procedure INput parameters:
  */
    p_sql           IN VARCHAR2,
    p_delimiter     IN VARCHAR2 :=',',
    p_text_quote    IN VARCHAR2 :='"',
    p_number_quote  IN VARCHAR2 := NULL,
    p_date_quote    IN VARCHAR2 := NULL,
    p_date_format   IN VARCHAR2 :='RRRR.MM.DD HH24:MI:SS',
    p_dir           IN VARCHAR2,
    p_header_file   IN VARCHAR2,
    p_data_file     IN VARCHAR2 := NULL
    );

PROCEDURE ref2csv_file(
  /* 
     Igor Perkovic, 10.07.2017.
  
     THE PURPOSE:
     This procedure catches the SYS_REFCURSOR from some outer function and write the result into the CSV file format.
  
     OPTIONS:
     You can customize the output a bit through the procedure INput parameters:
  */
    p_refcursor     IN OUT SYS_REFCURSOR,                     -- get the sys_refcursor
    p_delimiter     IN VARCHAR2 :=',',                        -- Main delimiter
    p_text_quote    IN VARCHAR2 :='"',                        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,                      -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,                      -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='RRRR.MM.DD HH24:MI:SS',    -- Date format
    p_dir           IN VARCHAR2,                              -- DIR alias
    p_header_file   IN VARCHAR2,                              -- file name
    p_data_file     IN VARCHAR2 := NULL
    ); 

PROCEDURE MAIL (
    /*
        Igor Perkovic, 21. July, 2017.
    
        THE PURPOSE: To send emails with multiple attachments to multiple recipients
    
        SPECIFICATION: This procedure can send up to 3 CLOB + 3 BLOB (overall 6) attachments.
                       The number of attachment can be easily extended if needed...
        --------------------------------------------------------------------------------------------
    
        Change log: 14.08.2017. - Added v_from, smtpHost and smtpPort as input variables and if-then switch,
                                  so I can add another smtp servers and ports
    
    */
    v_from        IN VARCHAR2 := 'mail@address.com',
    v_pass        IN VARCHAR2 := '<password>',
    smtpHost      IN VARCHAR2 := 'smtp.mailgun.org',
    smtpPort      IN NUMBER   :=  587,

    p_to          IN VARCHAR2,              -- recipient(s)
    p_cc          IN VARCHAR2 DEFAULT NULL, -- recipient(s)
    p_bcc         IN VARCHAR2 DEFAULT NULL, -- recipient(s)
    p_subject     IN VARCHAR2,              -- First part of the subject

    p_text_msg    IN VARCHAR2 DEFAULT NULL, -- Plain text message
    p_html_msg    IN VARCHAR2 DEFAULT NULL, -- HTML message
    p_clob_in     IN CLOB DEFAULT NULL,     -- Any CLOB you want. Either defined as parameter, or passed as a result of some function.
    p_charset     IN VARCHAR2 :='UTF-8',    -- You can change the character set encoding

    att_c       IN CLOB      DEFAULT NULL,
    att_c_file  IN VARCHAR2  DEFAULT NULL,
    att_c_mime  IN VARCHAR2  := 'text/html',

    att_c2      IN CLOB      DEFAULT NULL,
    att_c2_file IN VARCHAR2  DEFAULT NULL,
    att_c2_mime IN VARCHAR2  := 'text/html',

    att_c3      IN CLOB      DEFAULT NULL,
    att_c3_file IN VARCHAR2  DEFAULT NULL,
    att_c3_mime IN VARCHAR2  := 'text/html',

    att_c4      IN CLOB      DEFAULT NULL,
    att_c4_file IN VARCHAR2  DEFAULT NULL,
    att_c4_mime IN VARCHAR2  := 'text/html',    

    att_b       IN BLOB      DEFAULT NULL,
    att_b_file  IN VARCHAR2  DEFAULT NULL,
    att_b_mime  IN VARCHAR2  := 'application/excel',

    att_b2      IN BLOB      DEFAULT NULL,
    att_b2_file IN VARCHAR2  DEFAULT NULL,
    att_b2_mime IN VARCHAR2  := 'application/excel',

    att_b3      IN BLOB      DEFAULT NULL,
    att_b3_file IN VARCHAR2  DEFAULT NULL,
    att_b3_mime IN VARCHAR2  := 'application/excel'
    );


/* 

oooooooooooo                                                    oooo
`888'     `8                                                    `888
 888         oooo    ooo  .oooo.   ooo. .oo.  .oo.   oo.ooooo.   888   .ooooo.   .oooo.o
 888oooo8     `88b..8P'  `P  )88b  `888P"Y88bP"Y88b   888' `88b  888  d88' `88b d88(  "8
 888    "       Y888'     .oP"888   888   888   888   888   888  888  888ooo888 `"Y88b.
 888       o  .o8"'88b   d8(  888   888   888   888   888   888  888  888    .o o.  )88b
o888ooooood8 o88'   888o `Y888""8o o888o o888o o888o  888bod8P' o888o `Y8bod8P' 8""888P'
                                                      888
                                                     o888o


EXAMPLE 1: Send a picture from folder as BLOB attachment through Metronet smtp server
--------------------------------------------------------------------------------------

declare
    b blob;

begin
    b := trans.file2blob('DIR','Igor.png');
    
    trans.MAIL(
        v_from     => 'Woofy@mail.hr',
        smtpHost   => 'smtp.provider.com',
        smtpPort   =>  25,
        p_to       => 'john_doe@mail.com',
        p_subject  => 'Try and do not cry',
        p_html_msg => '<HTML><head><meta charset="UTF-8"></head> <BODY><P><font face="Segoe UI"> ... your masterpiece... </font></P> </BODY></HTML>',
        att_b      =>  b,
        att_b_file => 'Igor.png',
        att_b_mime => 'image/png'
    );
end;


EXAMPLE 2: Mail with HTML report in mail boday
--------------------------------------------------------------------
Prerequisites: as_xlsx

declare 
    m_clob clob;

begin
    m_clob := trans.query2html(
        p_sql       => 'select * from dept',
        p_css       => 1,
        p_charset   => 'Windows-1250');


    trans.MAIL(
        v_from      =>  'monitoring@mail.com',
        smtpHost    =>  'smtp.provider.com',
        smtpPort    =>   25,

        p_to        =>  'john_doe@mail.com',
        p_subject   =>  'TRY',

        p_clob_in   => m_clob,
        p_charset   => 'Windows-1250'
        );
end;




EXAMPLE 3: Multi-attachment and multi-recipients
--------------------------------------------------------------------
Prerequisites: as_xlsx

DECLARE
    l_clob  CLOB;
    m_clob  CLOB;
    b       BLOB;
    c       BLOB;

    catch   SYS_REFCURSOR;
    fn_cur  SYS_REFCURSOR;
    fm_cur  SYS_REFCURSOR;

BEGIN 
    catch := monitoring.fn_report_data('heta');
    l_clob := trans.ref2csv(
        p_refcursor     => catch,
        p_delimiter     => ';',
        p_text_quote    => '"',
        p_date_format   => 'YYYY-MM-DD'
        );

    m_clob := trans.query2html(
        p_sql           => 'select * from emp',
        p_charset       => 'Windows-1250',
        p_css           => 1,
        p_date_format   => 'DD/MM/RRRR'
    );

    -- If I want to generate xlsx file
        as_xlsx.new_sheet('EMP');
        as_xlsx.query2sheet( 'select * from emp', p_sheet => 1 );
        as_xlsx.new_sheet('DEPT');
        as_xlsx.query2sheet( 'select * from dept', p_sheet => 2 );

        -- Get the result blob into variable b for sending through the MAIL procedure
        b := as_xlsx.finish;

        fn_cur := fn_1(arg);
        fm_cur := fn_2(arg);

        -- Process the query result

        as_xlsx.clear_workbook;
        as_xlsx.query2sheet( p_cur => fn_cur, p_sheetname => 'one' );
        as_xlsx.query2sheet( p_cur => fm_cur, p_sheetname => 'two' );

        c := as_xlsx.finish;

    trans.MAIL (
            p_to        =>'john_doe@mail.com',
            p_bcc       =>'john_doe_bcc@mail.com',
            p_subject   => 'Monitoring',
            p_clob_in   => m_clob,

            att_c       => l_clob, 
            att_c_file  => 'file1.csv',
            att_b       =>  b,
            att_b_file  =>  'file2.xlsx',
            att_b2      =>  c,
            att_b2_file =>  'file3.xlsx'
            );
END;



EXAMPLE 4: Create a CSV file from query
--------------------------------------------------------------------

BEGIN 
    trans.query2csv_file(
        p_sql           => 'select * from employee',
        p_delimiter     => '|',
        p_text_quote    => '"',
        p_number_quote  => '',
        p_date_quote    => '',
        p_date_format   => 'DD/MM/RRRR',
        p_dir           => 'DIR',
        p_header_file   => 'trynow.csv');
END;

EXAMPLE 5: Catch the SYS_REFCURSOR from outer function names_for('New York') 
           and write the results into the file 'TRY.csv'

DECLARE
    catch  SYS_REFCURSOR;

BEGIN 
    catch := names_for('New York');

    trans.sysref2csv_file (
        p_refcursor     => catch,
        p_delimiter     => ';',
        p_text_quote    => '"',
        p_date_format   => 'YYYY-MM-DD',
        p_dir           => 'DIR',
        p_header_file   => 'TRY.csv'
    );
END;
*/

END;
/

  /*
  oooooooooo.    .oooooo.   oooooooooo.   oooooo   oooo
  `888'   `Y8b  d8P'  `Y8b  `888'   `Y8b   `888.   .8'
   888     888 888      888  888      888   `888. .8'
   888oooo888' 888      888  888      888    `888.8'
   888    `88b 888      888  888      888     `888'
   888    .88P `88b    d88'  888     d88'      888
  o888bood8P'   `Y8bood8P'  o888bood8P'       o888o
  */


create or replace PACKAGE BODY trans
IS

  /*
                                                          .oooo.
                                                        .dP""Y88b
   .ooooo oo oooo  oooo   .ooooo.  oooo d8b oooo    ooo       ]8P'  .ooooo.   .oooo.o oooo    ooo
  d88' `888  `888  `888  d88' `88b `888""8P  `88.  .8'      .d8P'  d88' `"Y8 d88(  "8  `88.  .8'
  888   888   888   888  888ooo888  888       `88..8'     .dP'     888       `"Y88b.    `88..8'
  888   888   888   888  888    .o  888        `888'    .oP     .o 888   .o8 o.  )88b    `888'
  `V8bod888   `V88V"V8P' `Y8bod8P' d888b        .8'     8888888888 `Y8bod8P' 8""888P'     `8'
        888.                                .o..P'
        8P'                                 `Y8P'
        "
  */
FUNCTION query2csv(
    p_sql           IN VARCHAR2,                -- SQL query
    p_type          IN NUMBER   := 0,           -- '1' for mail, '0' for file is default
    p_delimiter     IN VARCHAR2 :=';',          -- Main delimiter
    p_header        IN NUMBER   := 2,
    p_footer        IN NUMBER   := 0,
    p_cust_header   IN VARCHAR2 := NULL,
    p_cust_footer   IN VARCHAR2 := NULL,
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD.MM.RRRR') -- Date format
    RETURN CLOB
    AS
    v_finaltxt  VARCHAR2(4000);
    v_v_val     VARCHAR2(4000);
    v_n_val     NUMBER;
    v_d_val     DATE;
    v_ret       NUMBER;
    c           NUMBER;
    d           NUMBER;
    col_cnt     INTEGER;
    rec_tab     DBMS_SQL.DESC_TAB2;
    col_num     NUMBER;
    crlf        VARCHAR2(5)   := CHR(13) || CHR(10);
    cr          VARCHAR2(5)   := CHR(13);
    v_clob      CLOB := EMPTY_CLOB();

    BEGIN
    /* Query processing */ 
    c := DBMS_SQL.OPEN_CURSOR;
    DBMS_SQL.PARSE(c, p_sql, DBMS_SQL.NATIVE);
    d := DBMS_SQL.EXECUTE(c);
    
    /* Let's read and describe the columns automatically */
    DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
    FOR j in 1..col_cnt
    LOOP
      CASE rec_tab(j).col_type
        /* Text or varchar type */
        WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        /* Number type */
        WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
        /* DAte type */
        WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
      ELSE
        /* Other types */
        DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
      END CASE;
    END LOOP;
    

    IF p_header = 1 THEN
        /* This part outputs the HEADER */
        DBMS_LOB.CreateTemporary( v_clob, true );
        FOR j in 1..col_cnt
        LOOP
          v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||rec_tab(j).col_name||p_text_quote,p_delimiter);
        END LOOP;
        IF p_type = 0 THEN
          v_clob := v_finaltxt || cr;
        ELSE
          v_clob := v_finaltxt || UTL_TCP.crlf;
        END IF;
    END IF;

    IF p_header = 2 THEN
        /* This part outputs the CUSTOM HEADER */
        IF p_type = 0 THEN
          v_clob := p_cust_header || cr;
        ELSE
          v_clob := p_cust_header || UTL_TCP.crlf;
        END IF;
    END IF;
  

    /* This part outputs the DATA */
    LOOP
      v_ret := DBMS_SQL.FETCH_ROWS(c);
      EXIT WHEN v_ret = 0;
      v_finaltxt := NULL;
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                      v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||replace(v_v_val,p_delimiter,'')||p_text_quote,p_delimiter);
          WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                      v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_number_quote||replace(v_n_val,p_delimiter,'.')||p_number_quote,p_delimiter);
          WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                      v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote,p_delimiter);
        ELSE
          v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||v_v_val||p_text_quote,p_delimiter);
        END CASE;
      END LOOP;
      
    /* If we will transform this CLOB into the file, then we don't need extra blank line with CR + LF */
    IF p_type = 0 THEN
      v_clob := v_clob || v_finaltxt || cr;
    /* But if we send this CLOB through the mail, we could use it */
    ELSE
      v_clob := v_clob || v_finaltxt || UTL_TCP.crlf;
    END IF;
    END LOOP;

    IF p_footer = 1 THEN
        /* This part outputs the CUSTOM HEADER */
        IF p_type = 0 THEN
          v_clob := v_clob || p_cust_footer || cr;
        ELSE
          v_clob := v_clob || p_cust_footer || UTL_TCP.crlf;
        END IF;
    END IF;    
  
    /* Cleaning */
    DBMS_SQL.CLOSE_CURSOR(c);
    
    /* The result! */
    return v_clob;
  
    /* Some more cleaning */
    if DBMS_LOB.IsOpen( v_clob ) = 1 then
       DBMS_LOB.FreeTemporary( v_clob );
    end if;
    END;

  /*
                      .o88o.   .oooo.
                      888 `" .dP""Y88b
  oooo d8b  .ooooo.  o888oo        ]8P'  .ooooo.   .oooo.o oooo    ooo
  `888""8P d88' `88b  888        .d8P'  d88' `"Y8 d88(  "8  `88.  .8'
   888     888ooo888  888      .dP'     888       `"Y88b.    `88..8'
   888     888    .o  888    .oP     .o 888   .o8 o.  )88b    `888'
  d888b    `Y8bod8P' o888o   8888888888 `Y8bod8P' 8""888P'     `8'
  */
FUNCTION ref2csv(
  /* 
     Igor Perkovic, Created: 11.07.2017. | Last modified: 21.07.2017.
     
     THE PURPOSE:
     This function catches the SYS_REFCURSOR from some outer function and transform the function result to the CSV format 
     and return the result as a CLOB variable.
     
     This way I can pass the CLOB to the P_CLOB2FILE procedure - if I want to save the result as a CSV file, or
     I can send the CLOB result via email using procedure MAIL.
     
     OPTIONS:
     You can customize the output a bit through the procedure INput parameters:    
  */

    p_refcursor IN OUT SYS_REFCURSOR,       -- get the sys_refcursor
    p_type          IN NUMBER   := 0,           -- '1' for mail, '0' for file is default
    p_delimiter     IN VARCHAR2 :=';',          -- Main delimiter
    p_header        IN NUMBER   := 2,
    p_footer        IN NUMBER   := 0,
    p_cust_header   IN VARCHAR2 := NULL,
    p_cust_footer   IN VARCHAR2 := NULL,    
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD.MM.RRRR') -- Date format
    RETURN CLOB
    AS
    v_finaltxt  VARCHAR2(4000);
    v_v_val     VARCHAR2(4000);
    v_n_val     NUMBER;
    v_d_val     DATE;
    v_ret       NUMBER;
    c           BINARY_INTEGER;
    col_cnt     BINARY_INTEGER;
    rec_tab     DBMS_SQL.DESC_TAB2;
    col_num     NUMBER;
    crlf        VARCHAR2(5)   := CHR(13) || CHR(10);
    cr          VARCHAR2(5)   := CHR(13);
    v_clob      CLOB := EMPTY_CLOB();

    BEGIN
      /* Query processing */ 
      c := DBMS_SQL.TO_CURSOR_NUMBER(p_refcursor);
      
     /* Let's read and describe the columns automatically */
      DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          /* Text or varchar type */
          WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
          /* Number type */
          WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
          /* DAte type */
          WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
        ELSE
          /* Other types */
          DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        END CASE;
      END LOOP;

    IF p_header = 1 THEN
        /* This part outputs the HEADER */
        DBMS_LOB.CreateTemporary( v_clob, true );
        FOR j in 1..col_cnt
        LOOP
          v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||rec_tab(j).col_name||p_text_quote,p_delimiter);
        END LOOP;
        IF p_type = 0 THEN
          v_clob := v_finaltxt || cr;
        ELSE
          v_clob := v_finaltxt || UTL_TCP.crlf;
        END IF;
    END IF;

    IF p_header = 2 THEN
        /* This part outputs the CUSTOM HEADER */
        IF p_type = 0 THEN
          v_clob := p_cust_header || cr;
        ELSE
          v_clob := p_cust_header || UTL_TCP.crlf;
        END IF;
    END IF;

    
      /* This part outputs the DATA */
      LOOP
        v_ret := DBMS_SQL.FETCH_ROWS(c);
        EXIT WHEN v_ret = 0;
        v_finaltxt := NULL;
        FOR j in 1..col_cnt
        LOOP
          CASE rec_tab(j).col_type
            WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||replace(v_v_val,p_delimiter,'')||p_text_quote,p_delimiter);
            WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_number_quote||replace(v_n_val,p_delimiter,'.')||p_number_quote,p_delimiter);
            WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote,p_delimiter);
          ELSE
            v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||v_v_val||p_text_quote,p_delimiter);
          END CASE;
        END LOOP;
        
      /* If we will transform this CLOB into the file, then we don't need extra blank line with CR + LF */
      IF p_type = 0 THEN
        v_clob := v_clob || v_finaltxt || cr;
      /* But if we send this CLOB through the mail, we could use it */
      ELSE
        v_clob := v_clob || v_finaltxt || UTL_TCP.crlf;
      END IF;
      END LOOP;
    
    IF p_footer = 1 THEN
        /* This part outputs the CUSTOM HEADER */
        IF p_type = 0 THEN
          v_clob := v_clob || p_cust_footer || cr;
        ELSE
          v_clob := v_clob || p_cust_footer || UTL_TCP.crlf;
        END IF;
    END IF;    

      /* Cleaning */
      DBMS_SQL.CLOSE_CURSOR(c);
      
      /* The result! */
      return v_clob;
    
      /* Some more cleaning */
      if DBMS_LOB.IsOpen( v_clob ) = 1 then
         DBMS_LOB.FreeTemporary( v_clob );
      end if;
    END;


  /*
                                                          .oooo.   oooo            .                     oooo
                                                        .dP""Y88b  `888          .o8                     `888
   .ooooo oo oooo  oooo   .ooooo.  oooo d8b oooo    ooo       ]8P'  888 .oo.   .o888oo ooo. .oo.  .oo.    888
  d88' `888  `888  `888  d88' `88b `888""8P  `88.  .8'      .d8P'   888P"Y88b    888   `888P"Y88bP"Y88b   888
  888   888   888   888  888ooo888  888       `88..8'     .dP'      888   888    888    888   888   888   888
  888   888   888   888  888    .o  888        `888'    .oP     .o  888   888    888 .  888   888   888   888
  `V8bod888   `V88V"V8P' `Y8bod8P' d888b        .8'     8888888888 o888o o888o   "888" o888o o888o o888o o888o
        888.                                .o..P'
        8P'                                 `Y8P'
        "
  */
FUNCTION query2html(
    /* 
   Igor Perkovic, Created: 12.07.2017. | Updated: 21.07.2017.

   THE PURPOSE:
   This function translates the SQL query to HTML table and return the result as a CLOB variable.
   This way I can pass the CLOB to the P_CLOB2FILE procedure - if I want to save the result as a html file, or
   I can send the CLOB result via email with the MAIL procedure.

   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:

    */
    p_sql           IN VARCHAR2,                -- SQL query
    p_css           IN NUMBER   :=0,            -- No CSS by default - this is for in-mail sending (clob in mail body embedded). 
                                                -- as a side option, this can be used for CSS predefined styles, like 1,2,3, ...
                                                -- but for now, there is only one CSS in code. 
    p_crlf          IN NUMBER   :=1,            -- "crlf" is set to 1 by default.  0 is "cr" only.
    p_charset       IN VARCHAR2 := 'UTF-8',     -- META CHARSET encoding string
    p_tagl          IN VARCHAR2 := '<td>',      -- left tag (universal)
    p_tagr          IN VARCHAR2 := '</td>',     -- right tag (universal)
    p_tagl_v        IN VARCHAR2 := '<td>',      -- left tag for text data fields
    p_tagl_n        IN VARCHAR2 := '<td>',      -- left tag for number data fields
    p_tagl_d        IN VARCHAR2 := '<td>',      -- left tag for dates fields
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD-MM-YYYY',-- Date format
    p_cust_header   IN VARCHAR2 := '',
    p_cust_footer   IN VARCHAR2 := 'Powered by: <b>your_site.com</b>'
    )
    
    RETURN CLOB
    AS
    v_finaltxt      VARCHAR2(4000);
    v_v_val         VARCHAR2(4000);
    v_n_val         NUMBER;
    v_d_val         DATE;
    v_ret           NUMBER;
    c               NUMBER;
    d               NUMBER;
    col_cnt         INTEGER;
    rec_tab         DBMS_SQL.DESC_TAB2;
    col_num         NUMBER;
    crlf            VARCHAR2(5)   := CHR(13) || CHR(10);
    cr              VARCHAR2(5)   := CHR(13);
    v_date          VARCHAR2(15)  := TO_CHAR(TRUNC(SYSDATE),'YYYY-MM-DD');
    v_header        CLOB := EMPTY_CLOB();
    v_footer        CLOB := EMPTY_CLOB();
    v_css           CLOB := EMPTY_CLOB();
    v_clob          CLOB := EMPTY_CLOB();

    BEGIN
    /* Query processing */ 
    c := DBMS_SQL.OPEN_CURSOR;
    DBMS_SQL.PARSE(c, p_sql, DBMS_SQL.NATIVE);
    d := DBMS_SQL.EXECUTE(c);

    /* Let's read and describe the columns automatically */
    DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
    FOR j in 1..col_cnt
    LOOP
      CASE rec_tab(j).col_type
        WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
        WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
      ELSE
        DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
      END CASE;
    END LOOP;
    
    -- This part outputs the HEADER
    DBMS_LOB.CreateTemporary( v_clob,   true );
    DBMS_LOB.CreateTemporary( v_header, true );
    DBMS_LOB.CreateTemporary( v_footer, true );
    
    FOR j in 1..col_cnt
    LOOP
      v_finaltxt := ltrim(v_finaltxt||'<th>'||p_text_quote||rec_tab(j).col_name||p_text_quote||'</th>');
    END LOOP;
    
    /* If we want some style...
       this can be changed to desired style or even change to function argument - if you want to change your style often,
       OR you can set some predefined set (presets) here in the function and then call those presets by number from this IN argumnet...*/
    IF p_css <> 0 THEN
        DBMS_LOB.CreateTemporary( v_css, true );
        v_css := '<style type="text/css"> 
                       body { font-family: Arial, Helvetica, sans-serif; 
                              font-size:10pt;} 
                      table { empty-cells:show; 
                              border-collapse: collapse; 
                              border:solid 2px #777777;} 
                         td { border:solid 1px #333333; 
                              font-size:10pt; padding:5px;} 
                         th { background:#f9a61a; 
                              border:solid 1px #333333; 
                              font-size:11pt; 
                              padding:5px; 
                              vertical-align:center; 
                              text-align:center;} 
                         dt { font-weight: bold; }     
               </style>';

        SELECT TO_CHAR(TRUNC(SYSDATE),'DD.MM.YYYY') INTO v_date FROM dual;
     
        v_header := '<html> <head><meta charset="'||p_charset||'">'|| v_css ||'</head><body><P><font face="Segoe UI"> Datum: <b>'|| v_date || '</b>'|| p_cust_header || '<br><br>';
        v_footer := '</table><br>'||p_cust_footer||'</body></html>';
    /* Or style is just not our thing... give me just the vanilla table */
    ELSE
        v_header := '<table>';
        v_footer := '</table>';
    END IF;
  
    v_clob := v_header || '<table cellspacing="0" cellpadding="3"><tr>'|| v_finaltxt || '</tr>';

    /* This part outputs the DATA */
    LOOP
      v_ret := DBMS_SQL.FETCH_ROWS(c);
      EXIT WHEN v_ret = 0;
      v_finaltxt := NULL;
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_v||p_text_quote||v_v_val||p_text_quote||p_tagr);
          WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_n||p_number_quote||v_n_val||p_number_quote||p_tagr);
          WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_d||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote||p_tagr);
        ELSE
          v_finaltxt := ltrim(v_finaltxt||p_tagl||p_text_quote||v_v_val||p_text_quote||p_tagr);
        END CASE;
      END LOOP;

      if p_crlf = 1 THEN
        v_clob := v_clob || '<tr>' || v_finaltxt || '</tr>'|| crlf;
      else
        v_clob := v_clob || '<tr>' || v_finaltxt || '</tr>'|| cr;
      end if;
    END LOOP;
  
    /* Add some footer data if you have it... */
    v_clob := v_clob || v_footer;

    /* Cleaning...*/
    DBMS_SQL.CLOSE_CURSOR(c);

    /* THE RESULT!*/
    return v_clob;  

    /*Some more cleaning...*/
    IF DBMS_LOB.IsOpen( v_clob )   = 1 THEN  DBMS_LOB.FreeTemporary( v_clob );   END IF;
    IF DBMS_LOB.IsOpen( v_header ) = 1 THEN  DBMS_LOB.FreeTemporary( v_header ); END IF;
    IF DBMS_LOB.IsOpen( v_footer ) = 1 THEN  DBMS_LOB.FreeTemporary( v_header ); END IF;
    IF DBMS_LOB.IsOpen( v_css )    = 1 THEN  DBMS_LOB.FreeTemporary( v_css );    END IF;
  
    END;



  /*
                      .o88o.   .oooo.   oooo            .                     oooo
                      888 `" .dP""Y88b  `888          .o8                     `888
  oooo d8b  .ooooo.  o888oo        ]8P'  888 .oo.   .o888oo ooo. .oo.  .oo.    888
  `888""8P d88' `88b  888        .d8P'   888P"Y88b    888   `888P"Y88bP"Y88b   888
   888     888ooo888  888      .dP'      888   888    888    888   888   888   888
   888     888    .o  888    .oP     .o  888   888    888 .  888   888   888   888
  d888b    `Y8bod8P' o888o   8888888888 o888o o888o   "888" o888o o888o o888o o888o
  */

FUNCTION ref2html(
    /* 
   Igor Perkovic, Created: 12.07.2017. | Updated: 21.07.2017.

   THE PURPOSE:
   This function translates the SYS_REF cursor to HTML table and return the result as a CLOB variable.
   This way I can pass the CLOB to the P_CLOB2FILE procedure - if I want to save the result as a html file, or
   I can send the CLOB result via email..

   OPTIONS:
   You can customize the output a bit through the procedure INput parameters:

    */
    p_refcursor IN OUT SYS_REFCURSOR,           -- get the sys_refcursor
    p_css           IN NUMBER   :=0,            -- No CSS by default - this is for in-mail sending (clob in mail body embedded). 
                                                -- as a side option, this can be used for CSS predefined styles, like 1,2,3, ...
                                                -- but for now, there is only one CSS in code.
    p_crlf          IN NUMBER   :=1,            -- "crlf" is set to 1 by default.  0 is "cr" only.
    p_charset       IN VARCHAR2 := 'UTF-8',     -- META CHARSET encoding string                   
    p_tagl          IN VARCHAR2 := '<td>',      -- left tag (universal)
    p_tagr          IN VARCHAR2 := '</td>',     -- right tag (universal)
    p_tagl_v        IN VARCHAR2 := '<td>',      -- left tag for text data fields
    p_tagl_n        IN VARCHAR2 := '<td>',      -- left tag for number data fields
    p_tagl_d        IN VARCHAR2 := '<td>',      -- left tag for dates fields
    p_text_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,        -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,        -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='DD-MM-YYYY',-- Date format
    p_cust_header   IN VARCHAR2 := '',
    p_cust_footer   IN VARCHAR2 := 'Powered by: <b>your_site.com</b>'
    )

    RETURN CLOB
    AS
    v_finaltxt      VARCHAR2(4000);
    v_v_val         VARCHAR2(4000);
    v_n_val         NUMBER;
    v_d_val         DATE;
    v_ret           NUMBER;
    c               NUMBER;
    col_cnt         INTEGER;
    rec_tab         DBMS_SQL.DESC_TAB2;
    col_num         NUMBER;
    crlf            VARCHAR2(5)   := CHR(13) || CHR(10);
    cr              VARCHAR2(5)   := CHR(13);
    v_date          VARCHAR2(15)  := TO_CHAR(TRUNC(SYSDATE),'YYYY-MM-DD');
    v_header        CLOB := EMPTY_CLOB();
    v_footer        CLOB := EMPTY_CLOB();
    v_css           CLOB := EMPTY_CLOB();
    v_clob          CLOB := EMPTY_CLOB();

    BEGIN
    /* Query processing */ 
    c := DBMS_SQL.TO_CURSOR_NUMBER(p_refcursor);

    /* Let's read and describe the columns automatically */
    DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
    FOR j in 1..col_cnt
    LOOP
      CASE rec_tab(j).col_type
        WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
        WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
      ELSE
        DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
      END CASE;
    END LOOP;
    
    -- This part outputs the HEADER
    DBMS_LOB.CreateTemporary( v_clob,   true );
    DBMS_LOB.CreateTemporary( v_header, true );
    DBMS_LOB.CreateTemporary( v_footer, true );
    
    FOR j in 1..col_cnt
    LOOP
      v_finaltxt := ltrim(v_finaltxt||'<th>'||p_text_quote||rec_tab(j).col_name||p_text_quote||'</th>');
    END LOOP;
    
    /* If we want some style...
       this can be changed to desired style or even change to function argument - if you want to change your style often,
       OR you can set some predefined set (presets) here in the function and then call those presets by number from this IN argumnet...*/
    IF p_css <> 0 THEN
        DBMS_LOB.CreateTemporary( v_css, true );
        v_css := '<style type="text/css"> 
                       body { font-family: Arial, Helvetica, sans-serif; 
                              font-size:10pt;} 
                      table { empty-cells:show; 
                              border-collapse: collapse; 
                              border:solid 2px #777777;} 
                         td { border:solid 1px #333333; 
                              font-size:10pt; padding:5px;} 
                         th { background:#f9a61a; 
                              border:solid 1px #333333; 
                              font-size:11pt; 
                              padding:5px; 
                              vertical-align:center; 
                              text-align:center;} 
                         dt { font-weight: bold; }     
               </style>';

        SELECT TO_CHAR(TRUNC(SYSDATE),'DD.MM.YYYY') INTO v_date FROM dual;
     
        v_header := '<html> <head><meta charset="'||p_charset||'">'|| v_css ||'</head><body><P><font face="Segoe UI"> Datum: <b>'|| v_date || '</b>'|| p_cust_header || '<br><br>';
        v_footer := '</table><br>'||p_cust_footer||'</body></html>';
    /* Or style is just not our thing... give me just the vanilla table */
    ELSE
        v_header := '<table>';
        v_footer := '</table>';
    END IF;
  
    v_clob := v_header || '<table cellspacing="0" cellpadding="3"><tr>'|| v_finaltxt || '</tr>';

    /* This part outputs the DATA */
    LOOP
      v_ret := DBMS_SQL.FETCH_ROWS(c);
      EXIT WHEN v_ret = 0;
      v_finaltxt := NULL;
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_v||p_text_quote||v_v_val||p_text_quote||p_tagr);
          WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_n||p_number_quote||v_n_val||p_number_quote||p_tagr);
          WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                      v_finaltxt := ltrim(v_finaltxt||p_tagl_d||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote||p_tagr);
        ELSE
          v_finaltxt := ltrim(v_finaltxt||p_tagl||p_text_quote||v_v_val||p_text_quote||p_tagr);
        END CASE;
      END LOOP;

      if p_crlf = 1 THEN
        v_clob := v_clob || '<tr>' || v_finaltxt || '</tr>'|| crlf;
      else
        v_clob := v_clob || '<tr>' || v_finaltxt || '</tr>'|| cr;
      end if;
    END LOOP;
  
    /* Add some footer data if you have it... */
    v_clob := v_clob || v_footer;

    /* Cleaning...*/
    DBMS_SQL.CLOSE_CURSOR(c);

    /* THE RESULT!*/
    return v_clob;  

    /*Some more cleaning...*/
    IF DBMS_LOB.IsOpen( v_clob )   = 1 THEN  DBMS_LOB.FreeTemporary( v_clob );   END IF;
    IF DBMS_LOB.IsOpen( v_header ) = 1 THEN  DBMS_LOB.FreeTemporary( v_header ); END IF;
    IF DBMS_LOB.IsOpen( v_footer ) = 1 THEN  DBMS_LOB.FreeTemporary( v_header ); END IF;
    IF DBMS_LOB.IsOpen( v_css )    = 1 THEN  DBMS_LOB.FreeTemporary( v_css );    END IF;
    END;


  /*
   .o88o.  o8o  oooo              .oooo.    .o8       oooo             .o8
   888 `"  `"'  `888            .dP""Y88b  "888       `888            "888
  o888oo  oooo   888   .ooooo.        ]8P'  888oooo.   888   .ooooo.   888oooo.
   888    `888   888  d88' `88b     .d8P'   d88' `88b  888  d88' `88b  d88' `88b
   888     888   888  888ooo888   .dP'      888   888  888  888   888  888   888
   888     888   888  888    .o .oP     .o  888   888  888  888   888  888   888
  o888o   o888o o888o `Y8bod8P' 8888888888  `Y8bod8P' o888o `Y8bod8P'  `Y8bod8P'
  */
FUNCTION file2blob(
  /* 
    Igor Perkovic, Created: June, 2017 | Last modified: 21:07.2017.
   -----------------------------------------------------------------
    Read file into a BLOB
   
  */

    p_dir VARCHAR2, 
    p_file_name VARCHAR2) 
  RETURN BLOB 
  AS
    dest_loc  BLOB    := empty_blob();
    src_loc   BFILE   := BFILENAME(p_dir, p_file_name);
  
  BEGIN
    -- Open source binary file from OS
    DBMS_LOB.OPEN(src_loc, DBMS_LOB.LOB_READONLY);
  
    -- Create temporary LOB object
    DBMS_LOB.CREATETEMPORARY(
          lob_loc => dest_loc
        , cache   => true
        , dur     => dbms_lob.session
    );
  
    -- Open temporary lob
    DBMS_LOB.OPEN(dest_loc, DBMS_LOB.LOB_READWRITE);
  
    -- Load binary file into temporary LOB
    DBMS_LOB.LOADFROMFILE(
          dest_lob => dest_loc
        , src_lob  => src_loc
        , amount   => DBMS_LOB.getLength(src_loc));
  
    -- Close lob objects
    DBMS_LOB.CLOSE(dest_loc);
    DBMS_LOB.CLOSE(src_loc);
  
    -- Return temporary LOB object
    RETURN dest_loc;
  END file2blob;



  /*
            oooo             .o8         .oooo.    .o8       oooo             .o8
            `888            "888       .dP""Y88b  "888       `888            "888
   .ooooo.   888   .ooooo.   888oooo.        ]8P'  888oooo.   888   .ooooo.   888oooo.
  d88' `"Y8  888  d88' `88b  d88' `88b     .d8P'   d88' `88b  888  d88' `88b  d88' `88b
  888        888  888   888  888   888   .dP'      888   888  888  888   888  888   888
  888   .o8  888  888   888  888   888 .oP     .o  888   888  888  888   888  888   888
  `Y8bod8P' o888o `Y8bod8P'  `Y8bod8P' 8888888888  `Y8bod8P' o888o `Y8bod8P'  `Y8bod8P'
  */
FUNCTION clob2blob (
    /*  Igor Perkovic, Created 18.7.2017 | Last modified: 21.07.2017  */
        in_clob     IN CLOB, 
        in_charset  IN VARCHAR2 DEFAULT 'EE8MSWIN1250') 
    RETURN BLOB 
    AS 
        l_blob          BLOB;
        l_length        INTEGER;
        l_dest_offset   INTEGER := 1;
        l_src_offset    INTEGER := 1;
        l_lang_context  INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
        l_warning       INTEGER;
    
    BEGIN -- create new temporary BLOB
        DBMS_LOB.createtemporary(l_blob, FALSE);
        DBMS_LOB.convertToBlob(
            dest_lob        => l_blob, 
            src_clob        => in_clob, 
            amount          => DBMS_LOB.LOBMAXSIZE, 
            dest_offset     => l_dest_offset, 
            src_offset      => l_src_offset, 
            blob_csid       => nls_charset_id(in_charset), 
            lang_context    => l_lang_context, 
            warning         => l_warning);
        RETURN l_blob;
    
        DBMS_LOB.freetemporary(l_blob);
    
        EXCEPTION 
            WHEN OTHERS THEN 
                DBMS_LOB.freetemporary(l_blob);
            RAISE;
    END;



  /*
            oooo             .o8         .oooo.    .o88o.  o8o  oooo
            `888            "888       .dP""Y88b   888 `"  `"'  `888
   .ooooo.   888   .ooooo.   888oooo.        ]8P' o888oo  oooo   888   .ooooo.
  d88' `"Y8  888  d88' `88b  d88' `88b     .d8P'   888    `888   888  d88' `88b
  888        888  888   888  888   888   .dP'      888     888   888  888ooo888
  888   .o8  888  888   888  888   888 .oP     .o  888     888   888  888    .o
  `Y8bod8P' o888o `Y8bod8P'  `Y8bod8P' 8888888888 o888o   o888o o888o `Y8bod8P'
  */

PROCEDURE clob2file (
    v_clob    IN CLOB,
    v_dir     IN VARCHAR2,
    v_fname   IN VARCHAR2
    )
    IS
      l_file        UTL_FILE.FILE_TYPE;
      l_buffer      VARCHAR2(32767);
      l_amount      BINARY_INTEGER := 32767;
      l_pos         INTEGER := 1;
      l_clob_len    INTEGER;
    
    BEGIN
      l_clob_len := DBMS_LOB.GetLength(v_clob);
      l_file := UTL_FILE.fopen(v_dir, v_fname, 'w', 32767);
    
      WHILE l_pos < l_clob_len 
      LOOP
        DBMS_LOB.Read(v_clob, l_amount, l_pos, l_buffer);
        UTL_FILE.Put(l_file, l_buffer);
        UTL_FILE.fFlush(l_file);
        l_pos := l_pos + l_amount;
      END LOOP;
    
      UTL_FILE.fclose(l_file);
    
    EXCEPTION
       WHEN OTHERS THEN
          IF(UTL_FILE.Is_Open(l_file))THEN
             UTL_FILE.fClose(l_file);
          END IF;
          RAISE;
    END;



  /*
   .o8       oooo             .o8         .oooo.    .o88o.  o8o  oooo
  "888       `888            "888       .dP""Y88b   888 `"  `"'  `888
   888oooo.   888   .ooooo.   888oooo.        ]8P' o888oo  oooo   888   .ooooo.
   d88' `88b  888  d88' `88b  d88' `88b     .d8P'   888    `888   888  d88' `88b
   888   888  888  888   888  888   888   .dP'      888     888   888  888ooo888
   888   888  888  888   888  888   888 .oP     .o  888     888   888  888    .o
   `Y8bod8P' o888o `Y8bod8P'  `Y8bod8P' 8888888888 o888o   o888o o888o `Y8bod8P'
  */
PROCEDURE blob2file( 
    p_blob  IN BLOB, 
    p_dir   IN VARCHAR2, 
    p_file  IN VARCHAR2)

    IS
    
    vstart  NUMBER := 1;
    bytelen NUMBER := 32000;
    len     NUMBER;
    my_vr   RAW(32000);
    x       NUMBER;
    l_output utl_file.file_type; 

    BEGIN
        -- define output file with path
        l_output := utl_file.fopen(p_dir, p_file, 'WB', 32760);
        
        -- get the length of the blob
        len := dbms_lob.getlength(p_blob);
        -- save blob length
        x := len;
        
        -- if small enough for a single write
        IF len < 32760 THEN
            utl_file.put_raw(l_output,p_blob);
            utl_file.fflush(l_output);
        ELSE -- write in pieces
            vstart := 1;
            WHILE vstart < len
            LOOP
                dbms_lob.read(p_blob,bytelen,vstart,my_vr);
        
                utl_file.put_raw(l_output,my_vr);
                utl_file.fflush(l_output);
        
                -- set the start position for the next cut
                vstart := vstart + bytelen;
        
                -- set the end position if less than 32000 bytes
                x := x - bytelen;
                IF x < 32000 THEN
                    bytelen := x;
                END IF;
            END LOOP;
        END IF;
        utl_file.fclose(l_output);
    END blob2file;


  /*
                                                          .oooo.                                               .o88o.  o8o  oooo
                                                        .dP""Y88b                                              888 `"  `"'  `888
   .ooooo oo oooo  oooo   .ooooo.  oooo d8b oooo    ooo       ]8P'  .ooooo.   .oooo.o oooo    ooo             o888oo  oooo   888   .ooooo.
  d88' `888  `888  `888  d88' `88b `888""8P  `88.  .8'      .d8P'  d88' `"Y8 d88(  "8  `88.  .8'               888    `888   888  d88' `88b
  888   888   888   888  888ooo888  888       `88..8'     .dP'     888       `"Y88b.    `88..8'                888     888   888  888ooo888
  888   888   888   888  888    .o  888        `888'    .oP     .o 888   .o8 o.  )88b    `888'                 888     888   888  888    .o
  `V8bod888   `V88V"V8P' `Y8bod8P' d888b        .8'     8888888888 `Y8bod8P' 8""888P'     `8'     ooooooooooo o888o   o888o o888o `Y8bod8P'
        888.                                .o..P'
        8P'                                 `Y8P'
        "
  */
PROCEDURE query2csv_file(
    /* 
       Igor Perkovic, 10.07.2017.
    
       THE PURPOSE:
       This function translates the SQL query to the CSV format and write the result into the CSV file.
    
       OPTIONS:
       You can customize the output a bit through the procedure INput parameters:
    */
    p_sql           IN VARCHAR2,
    p_delimiter     IN VARCHAR2 :=',',
    p_text_quote    IN VARCHAR2 :='"',
    p_number_quote  IN VARCHAR2 := NULL,
    p_date_quote    IN VARCHAR2 := NULL,
    p_date_format   IN VARCHAR2 :='RRRR.MM.DD HH24:MI:SS',
    p_dir           IN VARCHAR2,
    p_header_file   IN VARCHAR2,
    p_data_file     IN VARCHAR2 := NULL)

    IS

    v_finaltxt         VARCHAR2(4000);
    v_v_val            VARCHAR2(4000);
    v_n_val            NUMBER;
    v_d_val            DATE;
    v_ret              NUMBER;
    c                  NUMBER;
    d                  NUMBER;
    col_cnt            INTEGER;
    rec_tab            DBMS_SQL.DESC_TAB2;
    col_num            NUMBER;
    v_fh               UTL_FILE.FILE_TYPE;
    v_samefile         BOOLEAN := (NVL(p_data_file,p_header_file) = p_header_file);

    BEGIN
      c := DBMS_SQL.OPEN_CURSOR;
      DBMS_SQL.PARSE(c, p_sql, DBMS_SQL.NATIVE);
      d := DBMS_SQL.EXECUTE(c);
      DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
    
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
          WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
          WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
        ELSE
          DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        END CASE;
      END LOOP;
    
      -- This part outputs the HEADER
      v_fh := UTL_FILE.FOPEN(upper(p_dir),p_header_file,'w',32767);
      FOR j in 1..col_cnt
      LOOP
        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||rec_tab(j).col_name||p_text_quote,p_delimiter);
      END LOOP;
    
      UTL_FILE.PUT_LINE(v_fh, v_finaltxt);
      IF NOT v_samefile THEN
        UTL_FILE.FCLOSE(v_fh);
      END IF;
      
      -- This part outputs the DATA
      IF NOT v_samefile THEN
        v_fh := UTL_FILE.FOPEN(upper(p_dir),p_data_file,'w',32767);
      END IF;
    
      LOOP
        v_ret := DBMS_SQL.FETCH_ROWS(c);
        EXIT WHEN v_ret = 0;
        v_finaltxt := NULL;
        FOR j in 1..col_cnt
        LOOP
          CASE rec_tab(j).col_type
            WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||replace(v_v_val,p_delimiter,'')||p_text_quote,p_delimiter);
            WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_number_quote||replace(v_n_val,p_delimiter,'.')||p_number_quote,p_delimiter);
            WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote,p_delimiter);
          ELSE
            v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||v_v_val||p_text_quote,p_delimiter);
          END CASE;
        END LOOP;
        UTL_FILE.PUT_LINE(v_fh, v_finaltxt);
      END LOOP;
      UTL_FILE.FCLOSE(v_fh);
      DBMS_SQL.CLOSE_CURSOR(c);
    END;


  /*
                      .o88o.   .oooo.                                               .o88o.  o8o  oooo
                      888 `" .dP""Y88b                                              888 `"  `"'  `888
  oooo d8b  .ooooo.  o888oo        ]8P'  .ooooo.   .oooo.o oooo    ooo             o888oo  oooo   888   .ooooo.
  `888""8P d88' `88b  888        .d8P'  d88' `"Y8 d88(  "8  `88.  .8'               888    `888   888  d88' `88b
   888     888ooo888  888      .dP'     888       `"Y88b.    `88..8'                888     888   888  888ooo888
   888     888    .o  888    .oP     .o 888   .o8 o.  )88b    `888'                 888     888   888  888    .o
  d888b    `Y8bod8P' o888o   8888888888 `Y8bod8P' 8""888P'     `8'     ooooooooooo o888o   o888o o888o `Y8bod8P'

  */
PROCEDURE ref2csv_file(
    /* 
       Igor Perkovic, 10.07.2017.
    
       THE PURPOSE:
       This procedure catches the SYS_REFCURSOR from some outer function and write the result into the CSV file format.
    
       OPTIONS:
       You can customize the output a bit through the procedure INput parameters:
    */
    p_refcursor     IN OUT SYS_REFCURSOR,                     -- get the sys_refcursor
    p_delimiter     IN VARCHAR2 :=',',                        -- Main delimiter
    p_text_quote    IN VARCHAR2 :='"',                        -- Quote symbol for wrapping text data
    p_number_quote  IN VARCHAR2 := NULL,                      -- Quote symbol for wrapping number data
    p_date_quote    IN VARCHAR2 := NULL,                      -- Quote symbol for wrapping dates
    p_date_format   IN VARCHAR2 :='RRRR.MM.DD HH24:MI:SS',    -- Date format
    p_dir           IN VARCHAR2,                              -- DIR alias
    p_header_file   IN VARCHAR2,                              -- file name
    p_data_file     IN VARCHAR2 := NULL) 

    IS

    v_finaltxt         VARCHAR2(4000);
    v_v_val            VARCHAR2(4000);
    v_n_val            NUMBER;
    v_d_val            DATE;
    v_ret              NUMBER;
    c                  BINARY_INTEGER;
    col_cnt            BINARY_INTEGER;
    rec_tab            DBMS_SQL.DESC_TAB2;
    col_num            NUMBER;
         
    v_fh               UTL_FILE.FILE_TYPE;
    v_samefile         BOOLEAN := (NVL(p_data_file,p_header_file) = p_header_file);

    BEGIN
    
      c := DBMS_SQL.TO_CURSOR_NUMBER(p_refcursor);
      DBMS_SQL.DESCRIBE_COLUMNS2(c, col_cnt, rec_tab);
    
      FOR j in 1..col_cnt
      LOOP
        CASE rec_tab(j).col_type
          WHEN 1 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
          WHEN 2 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_n_val);
          WHEN 12 THEN DBMS_SQL.DEFINE_COLUMN(c,j,v_d_val);
        ELSE
          DBMS_SQL.DEFINE_COLUMN(c,j,v_v_val,2000);
        END CASE;
      END LOOP;
    
      -- This part outputs the HEADER
      v_fh := UTL_FILE.FOPEN(upper(p_dir),p_header_file,'w',32767);
      FOR j in 1..col_cnt
      LOOP
        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||rec_tab(j).col_name||p_text_quote,p_delimiter);
      END LOOP;
    
      UTL_FILE.PUT_LINE(v_fh, v_finaltxt);
      IF NOT v_samefile THEN
        UTL_FILE.FCLOSE(v_fh);
      END IF;
      --
      -- This part outputs the DATA
      IF NOT v_samefile THEN
        v_fh := UTL_FILE.FOPEN(upper(p_dir),p_data_file,'w',32767);
      END IF;
      LOOP
        v_ret := DBMS_SQL.FETCH_ROWS(c);
        EXIT WHEN v_ret = 0;
        v_finaltxt := NULL;
        FOR j in 1..col_cnt
        LOOP
          CASE rec_tab(j).col_type
            WHEN 1 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_v_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||replace(v_v_val,p_delimiter,'')||p_text_quote,p_delimiter);
            WHEN 2 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_n_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_number_quote||replace(v_n_val,p_delimiter,'.')||p_number_quote,p_delimiter);
            WHEN 12 THEN DBMS_SQL.COLUMN_VALUE(c,j,v_d_val);
                        v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_date_quote||to_char(v_d_val,p_date_format)||p_date_quote,p_delimiter);
          ELSE
            v_finaltxt := ltrim(v_finaltxt||p_delimiter||p_text_quote||v_v_val||p_text_quote,p_delimiter);
          END CASE;
        END LOOP;
        UTL_FILE.PUT_LINE(v_fh, v_finaltxt);
      END LOOP;
      UTL_FILE.FCLOSE(v_fh);
      DBMS_SQL.CLOSE_CURSOR(c);
    END;



  /*
  ooo        ooooo       .o.       ooooo ooooo
  `88.       .888'      .888.      `888' `888'
   888b     d'888      .8"888.      888   888
   8 Y88. .P  888     .8' `888.     888   888
   8  `888'   888    .88ooo8888.    888   888
   8    Y     888   .8'     `888.   888   888       o
  o8o        o888o o88o     o8888o o888o o888ooooood8
  */
PROCEDURE MAIL (
    v_from        IN VARCHAR2 := 'mail@address.com',
    v_pass        IN VARCHAR2 := '<password>',
    smtpHost      IN VARCHAR2 := 'smtp.mailgun.org',
    smtpPort      IN NUMBER   :=  587,

    p_to          IN VARCHAR2,              -- recipient(s)
    p_cc          IN VARCHAR2 DEFAULT NULL, -- recipient(s)
    p_bcc         IN VARCHAR2 DEFAULT NULL, -- recipient(s)
    p_subject     IN VARCHAR2,              -- First part of the subject

    p_text_msg    IN VARCHAR2 DEFAULT NULL, -- Plain text message
    p_html_msg    IN VARCHAR2 DEFAULT NULL, -- HTML message
    p_clob_in     IN CLOB DEFAULT NULL,     -- Any CLOB you want. Either defined as parameter, or passed as a result of some function.
    p_charset     IN VARCHAR2 :='UTF-8',    -- You can change the character set encoding

    att_c       IN CLOB      DEFAULT NULL,
    att_c_file  IN VARCHAR2  DEFAULT NULL,
    att_c_mime  IN VARCHAR2  := 'text/html',

    att_c2      IN CLOB      DEFAULT NULL,
    att_c2_file IN VARCHAR2  DEFAULT NULL,
    att_c2_mime IN VARCHAR2  := 'text/html',

    att_c3      IN CLOB      DEFAULT NULL,
    att_c3_file IN VARCHAR2  DEFAULT NULL,
    att_c3_mime IN VARCHAR2  := 'text/html',

    att_c4      IN CLOB      DEFAULT NULL,
    att_c4_file IN VARCHAR2  DEFAULT NULL,
    att_c4_mime IN VARCHAR2  := 'text/html',

    att_b       IN BLOB      DEFAULT NULL,
    att_b_file  IN VARCHAR2  DEFAULT NULL,
    att_b_mime  IN VARCHAR2  := 'application/excel',

    att_b2      IN BLOB      DEFAULT NULL,
    att_b2_file IN VARCHAR2  DEFAULT NULL,
    att_b2_mime IN VARCHAR2  := 'application/excel',

    att_b3      IN BLOB      DEFAULT NULL,
    att_b3_file IN VARCHAR2  DEFAULT NULL,
    att_b3_mime IN VARCHAR2  := 'application/excel'
    )  
    
    AS
    v_subject           VARCHAR2(100);
    v_message_type      VARCHAR2(100)    := 'text/html';
    v_file_type         VARCHAR2(100)    := 'text/plain';
    v_message_body      CLOB;
    v_css               CLOB;

    conn                utl_smtp.connection;

    l_encoded_username  VARCHAR2(50 BYTE);
    l_encoded_password  VARCHAR2(50 BYTE);

    v_date              VARCHAR2(15) := TO_CHAR(TRUNC(SYSDATE),'YYYY-MM-DD');

    TYPE attach_info    IS RECORD (
         attach_name    VARCHAR2(40),
         data_type      VARCHAR2(40) DEFAULT 'text/plain',
         attach_content CLOB DEFAULT ''
    );
    TYPE array_attachments IS TABLE OF attach_info;
    attachments         array_attachments := array_attachments();


    TYPE attach_info_b  IS RECORD (
         attach_name    VARCHAR2(40),
         data_type      VARCHAR2(40) DEFAULT 'text/plain',
         attach_content BLOB DEFAULT NULL
    );
    TYPE array_attachments_b IS TABLE OF attach_info_b;
    attachments_b         array_attachments_b := array_attachments_b();

    /* For CLOB attachments */

    a1                  NUMBER :=0;
    a2                  NUMBER :=0;
    a3                  NUMBER :=0;
    a4                  NUMBER :=0;
    counter             NUMBER :=0 ;

    /* For BLOB attachments */
    b1                  NUMBER :=0;
    b2                  NUMBER :=0;
    b3                  NUMBER :=0;
    counter_b           NUMBER :=0 ;

    crlf                VARCHAR2(5)   := CHR(13) || CHR(10);
    n_offset            NUMBER;
    n_amount            NUMBER        := 1900;
    c_mime_boundary     CONSTANT VARCHAR2(256) := 'the boundary can be almost anything';
    l_step              PLS_INTEGER             := 12000; -- make sure you set a multiple of 3 not higher than 24573


    /*  Processing multiple recipients
       --------------------------------
    */
    FUNCTION fn_email_array( p_type  IN VARCHAR2, 
                             p_email IN VARCHAR2) 
        RETURN VARCHAR2 IS
        v_txt VARCHAR2(4000) := ''; v_id NUMBER; v_dop VARCHAR2(4000) := '';

        BEGIN
            IF p_email IS NULL 
                THEN RETURN NULL; 
            END IF;
            v_txt := replace(p_email,';',',');
            LOOP
                v_id := Instr(v_txt, ',');
                IF v_id <= 0 THEN 
                    IF Length(v_dop) > 0 THEN
                        v_dop := v_dop||Substr(v_txt, v_id + 1);
                        utl_smtp.rcpt(conn, replace(Substr(v_txt, v_id + 1),',',''));
                    ELSE
                        v_dop := v_txt;
                        utl_smtp.rcpt(conn, replace(v_txt,',',''));
                    END IF;
                    EXIT;
                END IF;
                v_dop := v_dop||Substr(v_txt, 1, v_id - 1)||';';
                utl_smtp.rcpt(conn, replace(Substr(v_txt, 1, v_id - 1),',',''));
                v_txt := Substr(v_txt, v_id + 1);
            END LOOP;
        RETURN v_dop;
    RETURN p_type||p_email;
    END;

    PROCEDURE proc_email_array( p_type  IN VARCHAR2, 
                                p_email IN VARCHAR2) 
        IS
            v_txt VARCHAR2(4000);
            BEGIN
                v_txt := fn_email_array(p_type, p_email);
            END;

    PROCEDURE writeData(p_text IN VARCHAR2) 
        AS
        BEGIN
            IF p_text IS NOT NULL THEN 
                utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(p_text||crlf)); 
            END IF;
        END;

    /* THIS IS THE BEGINING OF A GREAT ADVENTURE*/
    BEGIN
    /* Let's start in style */
    v_css := '<style type="text/css"> 
                     body { font-family: Arial, Helvetica, sans-serif; 
                            font-size:10pt;} 
                    table { empty-cells:show; 
                            border-collapse: collapse; 
                            border:solid 2px #777777;} 
                       td { border:solid 1px #333333; 
                            font-size:10pt; padding:5px;} 
                       th { background:#f9a61a; 
                            border:solid 1px #333333; 
                            font-size:11pt; 
                            padding:5px; 
                            vertical-align:center; 
                            text-align:center;} 
                       dt { font-weight: bold; }
             </style>';

    /* What is the messaage you'd like to send to the world ? */

    /* If you have something declared in p_text_message argument, then this will override other message body content */
    IF p_text_msg IS NOT NULL THEN
        v_message_body:= p_text_msg;
        v_message_type:='text/plain';
    ELSE 
        /* if you declare something in p_html_message, then this will override p_clob_in */
        if  p_html_msg IS NOT NULL THEN
            v_message_body:= '<html> <head><meta charset="'|| p_charset||'">' || v_css || '</head><body>'
                    || '<P><font face="Segoe UI">' 
                    || p_html_msg 
                    || '</body></html>';
        /* Finally, if you have nothing declared before, and HAVE something for p_clob_in, then your message body will be your outer CLOB*/
        else
            v_message_body:= p_clob_in;
        end if;            
    END IF;

    /*  ATTACHMENTS PART 
    -------------------------------------------------------------------------------
    This part of code is dedicated to process attachment part of the mail message  */

    /* CLOBs */

    /* Let's see how many attachments is declared... */
    if att_c  IS NOT NULL THEN a1 :=1; END IF;
    if att_c2 IS NOT NULL THEN a2 :=1; END IF;
    if att_c3 IS NOT NULL THEN a3 :=1; END IF;
    if att_c4 IS NOT NULL THEN a4 :=1; END IF;    

    /*... so we can precisely state the record size (and avoid to ask this info in the procedure aregumnets) */
    attachments.extend(a1+a2+a3+a4);

    if a1 = 1 THEN
        counter := counter + 1;
        SELECT att_c_file,att_c_mime,att_c INTO attachments(counter) FROM dual;
    end if;

    if a2 = 1 THEN
        counter := counter + 1;
        SELECT att_c2_file,att_c2_mime,att_c2 INTO attachments(counter) FROM dual;
    end if;

    if a3 = 1 THEN
        counter := counter + 1;
        SELECT att_c3_file,att_c3_mime,att_c3 INTO attachments(counter) FROM dual;
    end if;

    if a4 = 1 THEN
        counter := counter + 1;
        SELECT att_c4_file,att_c4_mime,att_c4 INTO attachments(counter) FROM dual;
    end if;

    /* BLOBs */

    /* Let's see how many attachments is declared... */
    if att_b  IS NOT NULL THEN b1 :=1; END IF;
    if att_b2 IS NOT NULL THEN b2 :=1; END IF;
    if att_b3 IS NOT NULL THEN b3 :=1; END IF;

    /*... so we can precisely state the record size (and avoid to ask this info in the procedure aregumnets) */
    attachments_b.extend(b1+b2+b3);

    if b1 = 1 THEN
        counter_b := counter_b + 1;
        SELECT att_b_file,att_b_mime,att_b INTO attachments_b(counter_b) FROM dual;
    end if;

    if b2 = 1 THEN
        counter_b := counter_b + 1;
        SELECT att_b2_file,att_b2_mime,att_b2 INTO attachments_b(counter_b) FROM dual;
    end if;

    if b3 = 1 THEN
        counter_b := counter_b + 1;
        SELECT att_b3_file,att_b3_mime,att_b3 INTO attachments_b(counter_b) FROM dual;
    end if;

    /* Give me a fresh and crispy date */
    SELECT TO_CHAR(TRUNC(SYSDATE),'YYYY-MM-DD') INTO v_date FROM dual;

  -- Open the SMTP connection ...

    conn := UTL_SMTP.open_connection(smtpHost, smtpPort);  

    if smtpPort = 587 THEN
        l_encoded_username := UTL_RAW.cast_to_varchar2(UTL_ENCODE.base64_encode(UTL_RAW.cast_to_raw(v_from)));  
        l_encoded_password := UTL_RAW.cast_to_varchar2(UTL_ENCODE.base64_encode(UTL_RAW.cast_to_raw(v_pass)));  

        UTL_SMTP.ehlo(conn, smtpHost);--DO NOT USE HELO  
        UTL_SMTP.command(conn, 'AUTH', 'LOGIN');  
        UTL_SMTP.command(conn, l_encoded_username);  
        UTL_SMTP.command(conn, l_encoded_password);
    else 
        utl_smtp.helo(conn, smtpHost);
    end if;

    utl_smtp.mail(conn, v_from);

    IF p_to  IS NOT NULL THEN proc_email_array('To: '  ,p_to);  END IF;
    IF p_cc  IS NOT NULL THEN proc_email_array('CC: '  ,p_cc);  END IF;
    IF p_bcc IS NOT NULL THEN proc_email_array('BCC: ' ,p_bcc); END IF;

  -- Open data
    utl_smtp.open_data(conn);

  -- Message info
    IF p_to  IS NOT NULL THEN writeData( 'To: '||p_to);   END IF;
    IF p_cc  IS NOT NULL THEN writeData( 'CC: '||p_cc);   END IF;
    IF p_bcc IS NOT NULL THEN writeData( 'BCC: '||p_bcc); END IF;

    /* We need to use "write_raw_data" to avoid encoding issues -- been there, done that, live to regret it... So, use this if you are from Europe or Asia
       The other way is to use just "write data", AND transform CLOBs to BLOBs with a proper encoding (been there, done that...)  */
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Date: '    || to_char(systimestamp, 'Dy, DD Mon YYYY HH24:MI:SS TZHTZM', 'NLS_DATE_LANGUAGE=ENGLISH') || crlf));
    --utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Date: '    || to_char(sysdate, 'Dy, DD Mon YYYY hh24:mi:ss') || crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('From: '    || v_from || crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Subject: ' || p_subject||' ['||v_date||']'||crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('MIME-Version: 1.0' || crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: multipart/mixed; boundary="SECBOUND"' || crlf || crlf));

    -- Message body
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND'     || crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: ' || v_message_type || crlf || crlf));
    utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(v_message_body   || crlf));

    -- Attachment Part
    /* CLOBs */
    /* If there is any CLOB attachments, roll the loop */
    IF counter > 0 THEN
        FOR i IN attachments.FIRST .. attachments.LAST
        LOOP
        -- Attach info
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND'     || crlf));
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: ' || attachments(i).data_type || ' name="'|| attachments(i).attach_name || '"' || crlf));
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Disposition: attachment; filename="'|| attachments(i).attach_name || '"' || crlf || crlf));
        -- Attach body
            n_offset := 1;
            WHILE n_offset < dbms_lob.getlength(attachments(i).attach_content)
            LOOP
                utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(dbms_lob.substr(attachments(i).attach_content, n_amount, n_offset)));
                n_offset := n_offset + n_amount;
            END LOOP;
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('' || crlf));
        END LOOP;
    END IF;

    /* BLOBs */
    /* If there is any BLOB attachments, roll the loop */
    IF counter_b > 0 THEN
        FOR i IN attachments_b.FIRST .. attachments_b.LAST
        LOOP
        -- Attach info
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('--SECBOUND'      || UTL_TCP.crlf));
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Type: '  || attachments_b(i).data_type || '; name="' || attachments_b(i).attach_name || '"' || UTL_TCP.crlf));
            UTL_SMTP.write_raw_data(conn, utl_raw.cast_to_raw('Content-Transfer-Encoding: base64' || UTL_TCP.crlf));
            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw('Content-Disposition: attachment; filename="' || attachments_b(i).attach_name || '"' || UTL_TCP.crlf || UTL_TCP.crlf));

            FOR j IN 0 .. TRUNC((DBMS_LOB.getlength(attachments_b(i).attach_content) - 1 )/l_step) 
            LOOP
                 UTL_SMTP.write_raw_data(conn, UTL_ENCODE.base64_encode(DBMS_LOB.substr(attachments_b(i).attach_content, l_step, j * l_step + 1)));
            END LOOP;

            utl_smtp.write_raw_data(conn, utl_raw.cast_to_raw(UTL_TCP.crlf || UTL_TCP.crlf));
            END LOOP;
    END IF;

    -- Close data
    utl_smtp.close_data(conn);
    utl_smtp.quit(conn);
    END;

END;
/
