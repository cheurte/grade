// use grade::{create_pdf, create_tex_file, page_blue_print};
use grade::{page_blue_print, render_tex_file};

use grade::ConfigXlsx;

// use std::process::Command;

fn main() {
    let configs = ConfigXlsx::from("config/config_source.json");
    let content = page_blue_print(configs);
    println!("{content:?}");
    match render_tex_file(content) {
        Ok(_) => println!("ok"),
        Err(e) => println!("{e:?}"),
    }
    let exit_status = std::process::Command::new("latexmk")
        .arg("output/report.tex")
        .arg("--output-directory=output/")
        .status()
        .unwrap();
    assert!(exit_status.success());

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
