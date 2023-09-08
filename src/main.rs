use std::{fs::File, io::Write};



#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("ESMA File downloader!");

    let filename = "temp.xlsx";

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
