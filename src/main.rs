/*
    https://github.com/netvl/xml-rs
 */

use std::{fs::File, io::Write};
use calamine::{Reader, open_workbook, Xlsx};
use chrono::{Utc, Datelike};
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

fn handle_row(stm: &mut Statement, row: &Row)
{
    println!("isin='{}' susp='{}' level='{}' start={} end={} as_of={}",
        row.isin, row.suspension, row.level, row.start_date, row.end_date, row.as_of_date);

    let _ = stm.execute(&[&row.isin, &row.suspension, &row.level, &row.start_date, &row.end_date, &row.as_of_date]);
}

fn read_spreadsheet(stm: &mut Statement, filename: &str) {

    let mut workbook: Xlsx<_> = open_workbook(filename).expect("Cannot open file");

    if let Some(Ok(range)) = workbook.worksheet_range("dvc_suspensions") {

        println!("Total entries: {}", range.get_size().0);

        for row_id in 1..range.get_size().0 {
            // let cell = range.get((0,0)).unwrap();
            let row = Row {
                isin        : range.get((row_id, 0)).unwrap().to_string(),
                suspension  : range.get((row_id, 1)).unwrap().to_string(),
                level       : range.get((row_id, 2)).unwrap().to_string(),
                start_date  : range.get((row_id, 3)).unwrap().to_string(),
                end_date    : range.get((row_id, 4)).unwrap().to_string(),
                as_of_date  : range.get((row_id, 5)).unwrap().to_string()
            };

            handle_row(stm, &row);
        }
    }
    else {
        println!("Nothing???");
    }
}



#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("ESMA File downloader - Equiduct 2023");

    let now = Utc::now();
    let filename = format!("dvc_suspensions_{}{}{}.xlsx", now.year(), now.month(), now.day());
    download_esma_file(filename.as_str()).await?;

    let conn = Connection::connect("testaf", "testaf", "//devdb001/mifex3")?;
    let mut stmt = conn.statement(
"insert into rd_esma_dvcsuspension(isin, dvcstatus, dvclevel, dvcstartdate, dvcenddate, dvcasofdate, importdate)
values (:1, :2, :3, :4, :5, :6, trunc(sysdate)) ").build()?;

    read_spreadsheet(&mut stmt, filename.as_str());

    let _ = stmt.close();
    let _ = conn.commit();

    Ok(())
}
