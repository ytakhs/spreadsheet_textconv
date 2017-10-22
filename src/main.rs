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
            let sheet_name = "[".to_string() + &name.to_string() + "]\t";
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
                let replaced_value: String = replace_special_chars(&value);
                output.push(replaced_value);
                output.push("\t".to_string());
            }
            output.push("\r\n".to_string());
        }
        let out = output.as_mut_slice().concat();
        print!("{}", out);
    }
}

fn replace_special_chars(cell: &String) -> String {
    let bs_str = REGEX_BS.replace_all(cell.as_str(), "\\\\").to_string();
    let n_str = REGEX_N.replace_all(bs_str.as_str(), "\\n").to_string();
    let r_str = REGEX_R.replace_all(n_str.as_str(), "\\r").to_string();
    let t_str = REGEX_T.replace_all(r_str.as_str(), "\\t").to_string();
    t_str
}
