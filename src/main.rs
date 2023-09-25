use std::{fs::File, io::Write};
use calamine::{Reader, open_workbook, Xlsx, DataType};
use chrono::{Utc, Datelike, Duration, NaiveDate};
use oracle::{Connection, Statement};

struct Row {
    isin: String,
    suspension: String,
    level: String,
    start_date: String,
    end_date: String,
    as_of_date: String
}


async fn download_esma_file(filename: &str) -> Result<(), Box<dyn std::error::Error>> {

    let url = "https://www.esma.europa.eu/sites/default/files/dvc_suspensions.xlsx";

    let resp = reqwest::get(url)
        .await?;

    println!("Response return code: {}", resp.status().as_str());

    let buf = resp.bytes().await?;

    println!("Response body size: {} bytes", buf.len());

    let mut file = File::create(filename)?;

    file.write_all(&buf)?;

    Ok(())
}

fn handle_row(stm: &mut Statement, row: &Row) -> Result<(), oracle::Error>
{
    println!("Inserting: isin='{}' susp='{}' level='{}' start={} end={} as_of={}",
        row.isin, row.suspension, row.level, row.start_date, row.end_date, row.as_of_date);

    return stm.execute(&[&row.isin, &row.suspension, &row.level,
                       &row.start_date, &row.end_date, &row.as_of_date]);
}


fn to_date_str(dt: f64) -> String
{
    let epoch = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
    (epoch + Duration::days(dt as i64 - 2)).to_string()
}

fn read_spreadsheet(stm: &mut Statement, filename: &str) {

    let mut workbook: Xlsx<_> = open_workbook(filename).expect("Cannot open file");

    if let Some(Ok(range)) = workbook.worksheet_range("dvc_suspensions") {

        println!("Total entries: {}", range.get_size().0);

        let mut i = 0;
        for row_id in 1..range.get_size().0 {
            // let cell = range.get((0,0)).unwrap().to_owned();

            let isin        = range.get((row_id, 0)).unwrap().to_string();
            let suspension  = range.get((row_id, 1)).unwrap().to_string();
            let level       = range.get((row_id, 2)).unwrap().to_string();
            let DataType::DateTime(start_date_f64)  = range.get((row_id, 3)).unwrap().to_owned()
                else { panic!("Unexpected data type"); };
            let DataType::DateTime(end_date_f64)    = range.get((row_id, 4)).unwrap().to_owned()
                else { panic!("Unexpected data type"); };
            let DataType::DateTime(as_of_date_f64)  = range.get((row_id, 5)).unwrap().to_owned()
                else { panic!("Unexpected data type"); };

            let start_date  = to_date_str(start_date_f64);
            let end_date    = to_date_str(end_date_f64);
            let as_of_date  = to_date_str(as_of_date_f64);

            // println!("row: {} {} {} {:?} {:?} {:?}", isin, suspension, level,
            //     to_date_str(start_date), to_date_str(end_date), to_date_str(as_of_date));

            let row = Row {
                isin        ,
                suspension  ,
                level       ,
                start_date,
                end_date  ,
                as_of_date
            };

            let rc = handle_row(stm, &row);
            match rc {
                Err(e) => {
                    println!("Failed: {}", e);
                    return;
                },
                _ => {}
            }

            i = i + 1;

            if i % 100 == 0 {
                println!("Inserted {} rows", i);
            }
        }
    }
    else {
        println!("Nothing???");
    }
}


#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("ESMA File downloader - Equiduct 2023");
    std::env::set_var("NLS_LANG", "AMERICAN_AMERICA.AL32UTF8");

    let now = Utc::now();
    let filename = format!("dvc_suspensions_{}{}{}.xlsx", now.year(), now.month(), now.day());

    // download_esma_file(filename.as_str()).await?;

    let conn = Connection::connect("testaf", "testaf", "//devdb001/mifex3")?;
    let mut stmt = conn.statement(
"insert into rd_esma_dvcsuspension(isin, dvcstatus, dvclevel, dvcstartdate, dvcenddate, dvcasofdate, importdate)
values (:1, :2, :3, to_date(:4, 'yyyy-mm-dd'), to_date(:5, 'yyyy-mm-dd'), to_date(:6, 'yyyy-mm-dd'), trunc(sysdate)) ").build()?;

    read_spreadsheet(&mut stmt, filename.as_str());

    let _ = stmt.close();
    let _ = conn.commit();

    Ok(())
}
