/*
    https://github.com/netvl/xml-rs
 */

use std::{fs::File, io::Write};
use calamine::{Reader, open_workbook, Xlsx};
use chrono::{Utc, Datelike};

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

fn handle_row(isin: &String, suspension: &String, level: &String, start_date: &String, end_date: &String,as_of_date: &String)
{
    println!("isin='{}' susp='{}' level='{}' start={} end={} as_of={}",
        isin, suspension, level, start_date, end_date, as_of_date );
}

fn read_spreadsheet(filename: &str) {

    let mut workbook: Xlsx<_> = open_workbook(filename).expect("Cannot open file");

    if let Some(Ok(range)) = workbook.worksheet_range("dvc_suspensions") {

        println!("Total entries: {}", range.get_size().0);

        for row_id in 1..range.get_size().0 {
            // let cell = range.get((0,0)).unwrap();
            let isin        = range.get((row_id, 0)).unwrap().to_string();
            let suspension  = range.get((row_id, 1)).unwrap().to_string();
            let level       = range.get((row_id, 2)).unwrap().to_string();
            let start_date  = range.get((row_id, 3)).unwrap().to_string();
            let end_date    = range.get((row_id, 4)).unwrap().to_string();
            let as_of_date  = range.get((row_id, 5)).unwrap().to_string();
            handle_row(&isin, &suspension, &level, &start_date, &end_date, &as_of_date);
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
    read_spreadsheet(filename.as_str());

    Ok(())
}
