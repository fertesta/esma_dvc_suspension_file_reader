# Downloader ESMA uploader

## Resources

## Resources

* Handling XML: https://github.com/netvl/xml-rs
* Oracle: https://blogs.oracle.com/timesten/post/using-rust-with-oracle-databases


# Notes

On Ubuntu, you need to install libaio:

```bash
sudo apt install libaio-dev libaio1
```

Example of connection to database for query:

```rust
    let conn = Connection::connect("testaf", "testaf", "//devdb001/mifex3")?;
    let mut stmt = conn.statement("select to_char(sysdate,'YYYY_MM_DD') from dual").build()?;

    let result_set = stmt.query(&[])?; // , &[&30])?;
    for result in result_set {
        match result {
            Err(e) => println!("Shit happened on err: {}", e),
            Ok(row) => {
                // let (empno, ename) = row.get_as::<(i32, String)>()?;
                let sdate = row.get_as::<String>()?;
                println!("Database date: {}", sdate);
            }
        }
    }
```

