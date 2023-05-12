// use grade::{create_pdf, create_tex_file, page_blue_print};
use grade::{page_blue_print, render_tex_file};

use grade::ConfigXlsx;

// use std::process::Command;

fn main() {
    let configs = ConfigXlsx::from("config/config_source.json");

    for pdf_file in configs.pdf_file.iter() {
        let (begin_categories, prod_coords): (Vec<(usize, usize)>, Vec<(usize, usize)>) =
            pdf_file.get_prod_categ_coordinates();
        // println!("{begin_categories:?}, {prod_coords:?}");
        let end_categories = pdf_file.get_parameters_range(&begin_categories);

        let _titles = pdf_file.get_title_names(&begin_categories);
        let _params = pdf_file.get_parameters_name(&begin_categories, &end_categories);

        let mut content: Vec<Vec<Vec<String>>> = Vec::new();

        for (_, prod_coord) in prod_coords.iter().enumerate() {
            let cont_buff = pdf_file.get_values_from_parameters(
                *prod_coord,
                &begin_categories,
                &end_categories,
            );
            content.push(cont_buff.clone());
        }

        let content = page_blue_print(configs, _titles, _params, content);
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
        break;
    }
}
