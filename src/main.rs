use office::{CellErrorType, DataType, Excel};
// use ansi_term::Colour::Green;
// use ansi_term::Colour::Red;
// use std::{thread, time::Duration};

fn main() {
    // opens a new workbook
    let path = format!("{}/Chat.xlsm", env!("CARGO_MANIFEST_DIR"));
    // let path = std::path::Path::new("./Chat.xlsm");
    let mut workbook = Excel::open(path).unwrap();
    // Read whole worksheet data and provide some statistics
    // if let Ok(range) = workbook.worksheet_range("Chat") {
    //     let total_cells = range.get_size().0 * range.get_size().1;
    //     let non_empty_cells: usize = range
    //         .rows()
    //         .map(|r| r.iter().filter(|cell| cell != &&DataType::Empty).count())
    //         .sum();
    //     println!(
    //         "Found {} cells in 'Chat', including {} non empty cells",
    //         total_cells, non_empty_cells
    //     );
    // }

    let sheets = workbook.sheet_names().expect("Empty workbook");
    for sheet_name in sheets {
        if sheet_name.starts_with('_') {
            continue;
        }
        println!("{}", &sheet_name);
        let range = workbook.worksheet_range(&sheet_name).unwrap();
        // let total_cells = range.get_size().0 * range.get_size().1;
        // let non_empty_cells: usize = range
            // .rows()
            // .map(|r| r.iter().filter(|cell| cell != &&DataType::Empty).count())
            // .sum();
        if &sheet_name == "Chat" {
            for c in range.rows() {
                if c != DataType::Empty {
                    println!("{:?}", &c);
                }
            }
        }
    }
}
