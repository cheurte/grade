// use grade::{create_pdf, create_tex_file, page_blue_print};
use grade::{AlignTab, TabularTex};
// use grade::ConfigXlsx;

// use std::process::Command;

fn main() {
    let tab: TabularTex = TabularTex::from(AlignTab::C, 3);
    tab.define_column();
    // let content = page_blue_print();
    // println!("{content:?}");
    // match create_tex_file(content) {
    //     Ok(_) => println!("ok"),
    //     Err(e) => println!("{e:?}"),
    // }
    // let exit_status = Command::new("latexmk")
    //     .arg("output/report.tex")
    //     .arg("--output-directory=output/")
    //     .status()
    //     .unwrap();
    // assert!(exit_status.success());

    // create_pdf().expect("error creating pdf");
    // let configs = ConfigXlsx::from("config/config_source.json");
    // for pdf_file in configs.pdf_file.iter() {
    //     let (begin_categories, prod_coords): (Vec<(usize, usize)>, Vec<(usize, usize)>) =
    //         pdf_file.get_prod_categ_coordinates();
    //     // println!("{begin_categories:?}, {prod_coords:?}");
    //     let end_categories = pdf_file.get_parameters_range(&begin_categories);
    //     // println!("{:?}", end_categories);
    //     let _param = pdf_file.get_parameters_name(&begin_categories, &end_categories);
    //     // println!("{param:?}");
    //
    //     for (i, prod_coord) in prod_coords.iter().enumerate() {
    //         let data = pdf_file.get_values_from_parameters(
    //             *prod_coord,
    //             &begin_categories,
    //             &end_categories,
    //         );
    //         println!("{data:?}");
    //     }
    //     println!("{_param:?}");
    //     break;
    // }
}
