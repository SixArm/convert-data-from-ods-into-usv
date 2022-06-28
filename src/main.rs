use std::env;
use std::path::Path;
use std::io::{Read, Seek};
use calamine::{Cell, DataType, Range, Reader, Xlsx}; // Also: CellType

type Workbook<T> = Xlsx<T>;
type Worksheet = (String, Range<DataType>);
type Row = [DataType];
type CellPosition = (u32, u32);

fn main() {
    let args: Vec<String> = env::args().collect();
    for arg in args { do_arg(arg) }
}

fn do_arg(arg: String) {
    let path = Path::new(&arg);
    do_path(path);
}

fn do_path(path: &Path) {
    let workbook = calamine::open_workbook(path).expect("Cannot open workbook");
    do_workbook(workbook);
}

fn do_workbook<T: Read + Seek>(mut workbook: Workbook<T>) {    
    let worksheets: Vec<Worksheet> = workbook.worksheets();
    for worksheet in worksheets { do_worksheet(worksheet) }
}

fn do_worksheet(worksheet: Worksheet) {
    let (name, range) = worksheet;
    let (height, width) = range.get_size();
    println!("worksheet name: {name}, range: {range:?}, height: {height}, width: {width}");
    let rows = range.rows();
    for row in rows { do_row(row) }
}

#[allow(dead_code)]
fn do_range(range: Range<DataType>) {
    let (size_height, size_width) = range.get_size();
    let (start_height, start_width) = range.start().expect("range.start");
    let (end_height, end_width) = range.end().expect("range.end");
    println!("size_height: {size_height}, size_width: {size_width}, start_height: {start_height}, start_width: {start_width}, end_height: {end_height}, end_width: {end_width}");
}

fn do_row(row: &Row) {
    let len = row.len();
    println!("row len: {len}");
    for data in row { do_data(data) }
}

#[allow(dead_code)]
fn do_cell(cell: Cell<DataType>) {
    let position: CellPosition = cell.get_position();
    let (position_0, position_1) = position;
    let value = cell.get_value();
    println!("position: {position:?}, position_0: {position_0}, position_1: {position_1}, value: {value:?}");
    do_data(value)
}

fn do_data(data: &DataType) {
    println!("data: {data:?}");
}
