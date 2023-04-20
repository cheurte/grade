use calamine::{Reader, open_workbook, Xlsx, DataType};

fn main() {
    let path = "sources/BIOTEC.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path).expect("cannot open file");
    
    if let Some(Ok(range)) = workbook.worksheet_range("Master - Flexible Overview") {
        // let total_cells = range.get_size().0 * range
        println!("{:?}", range.get_size());
    }
}
