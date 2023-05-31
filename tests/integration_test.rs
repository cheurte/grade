use grade::{ConfigXlsx, TabParameters};
use strum::IntoEnumIterator;

/// Test file
/// Need to add test for:
///     - Rendering
///     - Compilation
///     - empty config
///

#[test]
fn test_search_cells_coordinates() {
    let config_xlsx = ConfigXlsx::default();
    for pdf_file in config_xlsx.pdf_file.iter() {
        for (_, param) in TabParameters::iter().enumerate() {
            let res = pdf_file.search_cells_coordinates(param);
            assert_ne!(res, None);
        }
    }
}

// #[test]
// fn test_search_cells_coordinates_empty() {
//     let config_xlsx = ConfigXlsx::new();
//     for pdf_file in config_xlsx.pdf_file.iter() {
//         // println!("{pdf_file:?}");
//         // for param in TabParameters::iter() {
//         //     assert_eq!(pdf_file.search_cells_coordinates(param), None);
//         // }
//     }
// }
