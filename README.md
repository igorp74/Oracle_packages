# Oracle_packages
Some useful packages for Oracle's ecosystem

This is from time when I did not use Python for data processing and ETL. It is very usefull for reporting in Oracle environment.

### 💡 EXAMPLE 1: Send a picture from folder as BLOB attachment through smtp server

```sql
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
```

### 💡 EXAMPLE 2: Mail with HTML report in mail boday


Prerequisites: as_xlsx

```sql
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
```



### 💡 EXAMPLE 3: Multi-attachment and multi-recipients

Prerequisites: as_xlsx

```sql
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
```


### 💡 EXAMPLE 4: Create a CSV file from query

```sql
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
```

### 💡 EXAMPLE 5: Catch the SYS_REFCURSOR from outer function names_for('New York') and write the results into the file 'TRY.csv'

```sql
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
```
