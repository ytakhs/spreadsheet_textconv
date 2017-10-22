extern crate calamine;
extern crate regex;
#[macro_use]
extern crate lazy_static;

use std::error::Error;
use std::env;
use std::path::Path;
use calamine::{Sheets, Range, Rows, DataType};
use regex::Regex;

lazy_static! {
    static ref REGEX_N: Regex = Regex::new(r"\n").unwrap();
    static ref REGEX_R: Regex = Regex::new(r"\r").unwrap();
    static ref REGEX_T: Regex = Regex::new(r"\t").unwrap();
    static ref REGEX_BS: Regex = Regex::new(r"\\").unwrap();
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 2 {
        panic!("Usage: spreadsheet_textconv file.xlsx");
    }
    let file_path: &Path = Path::new(args.get(1).unwrap());
    let mut workbook: Sheets = match Sheets::open(file_path) {
        Err(why) => {
            let display = file_path.display();
            panic!("couldn't open {}: {}", display, Error::description(&why));
        }
        Ok(file) => file,
    };
    let sheet_names: Vec<String> = workbook.sheet_names().unwrap();
    for name in sheet_names {
        let range: Range<DataType> = workbook.worksheet_range(&name).unwrap();
        let rows: Rows<DataType> = range.rows();
        let mut output: Vec<String> = Vec::new();

        for row in rows {
            let sheet_name: String = "[".to_string() + &name.to_string() + "]\t";
            output.push(sheet_name);
            for cell in row.iter() {
                let value = match *cell {
                    DataType::Empty => "".to_string(),
                    DataType::String(ref string) => string.to_string(),
                    DataType::Float(ref float) => float.to_string(),
                    DataType::Int(ref int) => int.to_string(),
                    DataType::Bool(ref boolean) => boolean.to_string(),
                    DataType::Error(ref err) => err.to_string(),
                };
                let replaced_value: String = replace_special_chars(value);
                output.push(replaced_value);
                output.push("\t".to_string());
            }
            output.push("\r\n".to_string());
        }
        let out = output.as_mut_slice().concat();
        print!("{}", out);
    }
}

fn replace_special_chars(cell: String) -> String {
    let mut value = cell;
    value = REGEX_BS.replace_all(value.as_str(), "\\\\").to_string();
    value = REGEX_N.replace_all(value.as_str(), "\\n").to_string();
    value = REGEX_R.replace_all(value.as_str(), "\\r").to_string();
    value = REGEX_T.replace_all(value.as_str(), "\\t").to_string();
    value
}
