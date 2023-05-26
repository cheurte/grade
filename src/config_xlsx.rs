mod config_xlsx {
    use calamine::open_workbook;
    use calamine::Xlsx;
    use serde::Deserialize;
    use std::fs::File;
    use std::io::BufReader;
    use std::path::Path;

    #[derive(Debug)]
    pub enum TabParameters {
        Parameter,
        Category,
        Product,
    }

    #[derive(Debug, Deserialize, Clone)]
    #[serde(rename_all = "camelCase")]
    pub struct PdfFile {
        pdf_name: String,
        sources: Vec<String>,
        sheets_name: Vec<String>,
        products: Vec<String>,
        categories: Vec<String>,
        parameters: Vec<String>,
    }

    impl PdfFile {
        /// Function to find and return a workbook by id
        pub fn get_workbook(&self, id_woorkbook: usize) -> Xlsx<BufReader<File>> {
            let path = self.sources.iter().nth(id_woorkbook).unwrap();
            let workbook: Xlsx<_> = open_workbook(path).expect(
                "Cannot open workbook. It is probably an error in the path of the workbook.",
            );
            workbook
        }

        /// Function to find and return a worksheet by id
        pub fn get_worksheet(&self, id_worksheet: usize) -> &String {
            self.sheets_name.iter().nth(id_worksheet).expect(
                "Please make sur the name between the worksheet and
                the name given in the configuration file are the same.",
            )
        }
        /// Function to search the key elements of the file.
        /// Suppose to depend on a json file to know wich line or row can be read.
        pub fn search_cells_coordinates(&self, field: TabParameters) -> Vec<(usize, usize)> {
            let mut workbook = self.get_workbook(0);

            let mut output: Vec<(usize, usize)> = Vec::new();
            let field = match field {
                TabParameters::Product => &self.products,
                TabParameters::Parameter => &self.parameters,
                TabParameters::Category => &self.categories,
            };
            match workbook.worksheet_range(self.get_worksheet(0)) {
                Some(Ok(range)) => {
                    for row in 0..range.get_size().0 {
                        for col in 0..range.get_size().1 {
                            let value = range.get_value((row as u32, col as u32));
                            if value != Some(&DataType::Empty) {
                                for category in field {
                                    if &value.unwrap().to_string() == category {
                                        output.push((row, col));
                                    }
                                }
                            }
                        }
                    }
                }
                Some(Err(e)) => println!("{e:?}"),
                None => println!("Sheets name unknown. Maybe check the name in the config file"),
            }
            assert_eq!(output.len(), field.len());
            output
        }

        /// Search the range of categories, to know where to stop
        /// Take the beginning corrdinates of categories, and return the end coordinates
        pub fn get_parameters_range(
            &self,
            categories_coord: &Vec<(usize, usize)>,
        ) -> Vec<(usize, usize)> {
            // let sheet = self.sheets_name.iter().nth(0).unwrap();
            let mut end_categories: Vec<(usize, usize)> = vec![];
            let mut workbook = self.get_workbook(0);

            match workbook.worksheet_range(&self.get_worksheet(0)) {
                Some(Ok(range)) => {
                    for (category_row, category_col) in categories_coord.into_iter() {
                        let mut col = category_col + 1;
                        loop {
                            if range.get_value((*category_row as u32, col as u32))
                                != Some(&DataType::Empty)
                                || col > range.get_size().1
                            {
                                break;
                            }
                            col += 1;
                        }
                        end_categories.push((*category_row, col - 1));
                    }
                }
                Some(Err(e)) => println!("{e:?}"),
                None => println!("Sheets name unknown. Maybe check the name in the config file"),
            }
            end_categories
        }
        /// Return the tab titles
        pub fn get_values_at(&self, begin_categories: &Vec<(usize, usize)>) -> Vec<String> {
            let mut workbook = self.get_workbook(0);
            let mut output: Vec<String> = vec![];

            match workbook.worksheet_range(self.get_worksheet(0)) {
                Some(Ok(range)) => {
                    for category in begin_categories {
                        let (a, b) = category;
                        // println!("{:?}", range.get_value((*a as u32, *b as u32)));
                        output.push(range.get_value((*a as u32, *b as u32)).unwrap().to_string())
                    }
                }
                Some(Err(e)) => println!("{e:?}"),
                None => println!("Sheets name unknown. Maybe check the name in the config file"),
            }
            output
        }

        pub fn get_parameters_by_id(
            &self,
            start_categ_coord: &Vec<(usize, usize)>,
            end_categ_coord: &Vec<(usize, usize)>,
            id_line: Vec<usize>,
        ) -> Vec<Vec<String>> {
            assert_eq!(start_categ_coord.len(), end_categ_coord.len());

            let mut workbook = self.get_workbook(0);
            let mut output: Vec<Vec<String>> = vec![];

            match workbook.worksheet_range(self.get_worksheet(0)) {
                Some(Ok(range)) => {
                    let it = start_categ_coord.iter().zip(end_categ_coord.iter());
                    for (_, (start_coord, end_coord)) in it.enumerate() {
                        let mut parameters: Vec<String> = vec![];
                        for col in start_coord.1..end_coord.1 + 1 {
                            for line in id_line.iter() {
                                parameters.push(
                                    range
                                        .get_value(((start_coord.0 + line) as u32, col as u32))
                                        .unwrap()
                                        .to_string(),
                                )
                            }
                        }
                        output.push(parameters);
                    }
                }
                Some(Err(e)) => println!("{e:?}"),
                None => println!("Sheets name unknown. Maybe check the name in the config file"),
            }
            output
        }

        pub fn get_values_from_parameters(
            &self,
            product_coordinates: (usize, usize),
            start_categ_coord: &Vec<(usize, usize)>,
            end_categ_coord: &Vec<(usize, usize)>,
        ) -> Vec<Vec<String>> {
            let mut workbook = self.get_workbook(0);
            let mut out: Vec<Vec<String>> = Vec::new();
            match workbook.worksheet_range(self.get_worksheet(0)) {
                Some(Ok(range)) => {
                    for param in 0..start_categ_coord.len() {
                        let mut parameters: Vec<String> = vec![];
                        for y in start_categ_coord.iter().nth(param).unwrap().1
                            ..end_categ_coord.iter().nth(param).unwrap().1 + 1
                        {
                            let x = product_coordinates.0;
                            parameters
                                .push(range.get_value((x as u32, y as u32)).unwrap().to_string());
                        }
                        out.push(parameters);
                    }
                }
                Some(Err(e)) => println!("{e:?}"),
                None => println!("Sheets name unknown. Maybe check the name in the config file"),
            }
            out
        }
    }

    #[derive(Debug, Deserialize, Clone)]
    #[serde(rename_all = "camelCase")]
    pub struct ConfigXlsx {
        pub pdf_file: Vec<PdfFile>,
        pub color_text: Vec<i32>,
        pub color_tab_title: Vec<i32>,
        pub color_tab_line: Vec<i32>,
        pub margin_size: f32,
        pub alignment_tabular: String,
    }

    impl ConfigXlsx {
        pub fn new() -> Self {
            Self {
                pdf_file: vec![],
                color_text: vec![13, 64, 47],
                color_tab_title: vec![237, 233, 230],
                color_tab_line: vec![215, 212, 210],
                margin_size: 0.84,
                alignment_tabular: String::from("left"),
            }
        }

        pub fn from(path: &str) -> Self {
            let path = Path::new(path);
            let file = File::open(path).expect("Problem with the file");
            let config: ConfigXlsx = serde_json::from_reader(file).expect("error in the reading");
            config
        }
    }
}
