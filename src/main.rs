extern crate calamine;

use std::error::Error;
use std::env;
use std::path::Path;
use calamine::{Sheets, Range, Rows, DataType};

fn main() {
    let args: Vec<String> = env::args().collect();
    let file_path: &Path = Path::new(args.get(1).unwrap());
    let display = file_path.display();
    let mut workbook: Sheets = match Sheets::open(file_path) {
        Err(why) => panic!("couldn't open {}: {}", display, Error::description(&why)),
        Ok(file) => file,
    };
    let sheet_names: Vec<String> = workbook.sheet_names().unwrap();
    for name in sheet_names {
        let range: Range<DataType> = workbook.worksheet_range(&name).unwrap();
        let rows: Rows<DataType> = range.rows();
        let mut output: Vec<String> = Vec::new();
        for row in rows {
            for (_index, cell) in row.iter().enumerate() {
                let value = match *cell {
                    DataType::Empty => "".to_string(),
                    DataType::String(ref string) => string.to_string(),
                    DataType::Float(ref float) => float.to_string(),
                    DataType::Int(ref int) => int.to_string(),
                    DataType::Bool(ref boolean) => boolean.to_string(),
                    DataType::Error(ref err) => err.to_string(),
                };
                output.push(value);
                output.push("\t".to_string());
            }
        }
        let out = output.as_mut_slice().concat();
        println!("{}", out);
    }
}
