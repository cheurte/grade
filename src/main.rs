use grade::{Config, ConfigXlsx, TabParameters};
use latex::Document;
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let mut config_file: Config = Config::default();
    match Config::new(&args) {
        Ok(v) => {
            config_file = v;
        }
        Err(e) => println!("WARNING {}, the default is used instead", e),
    };
    let configs = ConfigXlsx::from(config_file.get_config_path())?;
    // Iteration over pdf files
    // only one for now & for testing
    // configs.pdf_file.push(PdfFile::new());
    for pdf_file in configs.pdf_file.iter() {
        let begin_categories_coord: Option<Vec<(usize, usize)>> =
            pdf_file.search_cells_coordinates(TabParameters::Category);
        let parameters_coord: Option<Vec<(usize, usize)>> =
            pdf_file.search_cells_coordinates(TabParameters::Parameter);
        let products_coord: Option<Vec<(usize, usize)>> =
            pdf_file.search_cells_coordinates(TabParameters::Product);
        let end_categories_coord = pdf_file.get_parameters_range(&begin_categories_coord);

        let titles = pdf_file.get_values_at(&begin_categories_coord);
        let parameters = pdf_file.get_values_at(&parameters_coord);
        let product_names = pdf_file.get_values_at(&products_coord);

        let general_content = pdf_file.get_parameters_by_id(
            &begin_categories_coord,
            &end_categories_coord,
            &parameters_coord,
        );

        let mut product_values: Vec<Vec<Vec<String>>> = Vec::new();

        // finding the actual content
        if let Some(product_coord) = products_coord {
            for (_, prod_coord) in product_coord.iter().enumerate() {
                let cont_buff = pdf_file.get_values_from_parameters(
                    *prod_coord,
                    &begin_categories_coord,
                    &end_categories_coord,
                );
                product_values.push(cont_buff.ok_or("Content Not Found")?.clone());
            }
        }

        //
        let mut page = Document::new(latex::DocumentClass::Article);
        configs.preamble(&mut page);
        configs.first_page(&mut page, &product_names);
        //     // Page creation
        //     // We iterate over the PRODUCT
        if !product_values.is_empty() {
            let mut product_value = product_values.iter();
            let mut product_name = product_names.as_ref().ok_or("")?.iter();
            for _ in 0..product_values.len() {
                let values = product_value.next().ok_or("no value next")?;
                let product_name = product_name.next();
                configs
                    .page_blue_print(
                        &mut page,
                        product_name.ok_or("Erreur product names")?.to_string(),
                        &titles,
                        &parameters,
                        &general_content,
                        values,
                        match parameters_coord {
                            Some(ref v) => v.len(),
                            None => 0,
                        },
                    )
                    .ok_or("")?;
            }
        }

        match pdf_file.create_and_render(page) {
            Ok(_) => println!("PDF CREATED WITH SUCCESS"),
            Err(e) => println!("ERROR IN CREATION {:?}", e),
        }
    }
    Ok(())
}
