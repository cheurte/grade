use grade::render_tex_file;

use grade::{ConfigXlsx, TabParameters};
use latex::{print, Document};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let configs = ConfigXlsx::from("config/config_source.json")?;
    // Iteration over pdf files
    // only one for now & for testing
    for pdf_file in configs.pdf_file.iter() {
        let begin_categories_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Category);
        let parameters_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Parameter);
        let products_coord: Vec<(usize, usize)> =
            pdf_file.search_cells_coordinates(TabParameters::Product);
        let end_categories_coord = pdf_file.get_parameters_range(&begin_categories_coord);
        let titles = pdf_file
            .get_values_at(&begin_categories_coord)
            .ok_or("Error at getting titles")?;
        let parameters = pdf_file
            .get_values_at(&parameters_coord)
            .ok_or("Error at getting parameters")?;
        let product_names = pdf_file
            .get_values_at(&products_coord)
            .ok_or("Error at getting product names")?;
        // println!("{parameters:?}");

        let t: Vec<usize> = parameters_coord.iter().map(|v| v.0).collect();
        let general_content =
            pdf_file.get_parameters_by_id(&begin_categories_coord, &end_categories_coord, t);

        let mut product_values: Vec<Vec<Vec<String>>> = Vec::new();

        // finding the actual content
        for (_, prod_coord) in products_coord.iter().enumerate() {
            let cont_buff = pdf_file.get_values_from_parameters(
                *prod_coord,
                &begin_categories_coord,
                &end_categories_coord,
            );
            product_values.push(cont_buff.clone());
        }

        let mut page = Document::new(latex::DocumentClass::Article);
        configs.starting_pdf(&mut page);
        configs.first_page(&mut page, &product_names);
        // Page creation
        // We iterate over the PRODUCT
        let mut product_value = product_values.iter();
        let mut product_names = product_names.iter();
        for _ in 0..product_values.len() {
            let values = product_value.next().ok_or("no value next")?;
            let product_name = product_names.next().ok_or("no value next")?;
            configs
                .page_blue_print(
                    &mut page,
                    product_name.to_string(),
                    &titles,
                    parameters.clone(),
                    &general_content,
                    values,
                    parameters_coord.len(),
                )
                .ok_or("")?;
            // break;
        }

        let render = print(&page)?;
        match render_tex_file(render) {
            Ok(_) => println!("rendered completed"),
            Err(e) => println!("{e:?}"),
        }

        std::process::Command::new("latexmk")
            .arg("output/report.tex")
            .arg("-pdf")
            .arg("--output-directory=output/")
            .status()?;
        break;
    }
    Ok(())
}
