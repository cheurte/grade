use grade::{page_blue_print, render_tex_file, starting_pdf};

use grade::{ConfigXlsx, TabParameters};
use latex::{print, Document};

// use std::process::Command;

fn main() {
    let configs = ConfigXlsx::from("config/config_source.json");

    // Iteration over pdf files
    // only one for now & for testing
    for pdf_file in configs.pdf_file.iter() {
        let begin_categories_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Category);
        let parameters_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Parameter);
        let products_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Product);
        //
        let end_categories_coord = pdf_file.get_parameters_range(&begin_categories_coord);
        let titles = pdf_file.get_title_names(&begin_categories_coord);

        let t: Vec<usize> = parameters_coord.iter().map(|v| v.0).collect();
        let parameters =
            pdf_file.get_parameters_by_id(&begin_categories_coord, &end_categories_coord, t);

        let mut values: Vec<Vec<Vec<String>>> = Vec::new();

        // finding the actual content
        for (_, prod_coord) in products_coord.iter().enumerate() {
            let cont_buff = pdf_file.get_values_from_parameters(
                *prod_coord,
                &begin_categories_coord,
                &end_categories_coord,
            );
            values.push(cont_buff.clone());
        }

        let mut page = Document::new(latex::DocumentClass::Article);
        starting_pdf(&mut page, &configs);

        // Page creation
        // We iterate over the PRODUCT
        for _ in 0..values.len() {
            let content = values.iter().next().unwrap();
            page_blue_print(
                &mut page,
                &titles,
                &parameters,
                content,
                parameters_coord.len(),
            );
            break;
        }
        let render = print(&page).unwrap();
        // match render_tex_file(render) {
        //     Ok(_) => println!("rendered completed"),
        //     Err(e) => println!("{e:?}"),
        // }
        // let exit_status = std::process::Command::new("latexmk")
        //     .arg("output/report.tex")
        //     .arg("--output-directory=output/")
        //     .status()
        //     .unwrap();
        // assert!(exit_status.success());
        break;
    }
}
