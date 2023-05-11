/// ALignment defined here
#[derive(Debug, Clone)]
pub enum AlignTab {
    C,
    R,
    L,
}

/// Structure to help creating a tab
#[derive(Debug, Clone)]
pub struct TabularTex {
    align: AlignTab,
    nb_col: usize,
    title_color: String,
    line_color: String,
    title: String,
    sub_title: String,
    content: String,
}

impl TabularTex {
    pub fn new() -> Self {
        Self {
            align: AlignTab::L,
            nb_col: 6,
            title_color: String::new(),
            line_color: String::new(),
            title: String::new(),
            sub_title: String::new(),
            content: String::new(),
        }
    }
    pub fn empty_default(
        align: AlignTab,
        nb_col: usize,
        title_color: String,
        line_color: String,
    ) -> Self {
        Self {
            align,
            nb_col,
            title_color,
            line_color,
            title: String::new(),
            sub_title: String::new(),
            content: String::new(),
        }
    }

    pub fn from(title: String, sub_title: Vec<String>, content: Vec<Vec<String>>) -> Self {
        let mut nb_c: usize = sub_title.len();
        if nb_c < 6 {
            nb_c = 6;
        }
        let mut t: TabularTex = TabularTex {
            align: AlignTab::L,
            nb_col: nb_c,
            title_color: String::from("cyan"),
            line_color: String::from("cyan"),
            title: String::new(),
            sub_title: String::new(),
            content: String::new(),
        };
        t.define_title(title);
        t.define_sub_title(sub_title);
        t.define_content(content);
        t
    }

    // pub fn from(align: AlignTab, )
    /// Define the number of column needed
    fn define_column(&self) -> String {
        let mut column_definition: String = String::new();
        let align = match self.align {
            AlignTab::C => String::from("c"),
            AlignTab::L => String::from("l"),
            AlignTab::R => String::from("r"),
        };
        for _ in 0..self.nb_col {
            column_definition.push_str(&format!("X {} ", align)[..]);
        }
        column_definition
    }
    /// Define and return latex title for a
    fn define_title(&mut self, content: String) {
        // let mut title = String::new();
        self.title
            .push_str(&format!("\\rowcolor{{{}}}{{{}}} ", self.title_color, content)[..]);
        for _ in 0..self.nb_col - 1 {
            self.title.push_str(" & ");
        }
        self.content.push_str("\\\\");
    }

    fn define_sub_title(&mut self, content: Vec<String>) {
        for value in content.iter() {
            self.sub_title
                .push_str(&format!("\\textbf{{{}}} & ", value)[..]);
        }
        if content.len() < self.nb_col {
            for _ in 0..self.nb_col - content.len() {
                self.sub_title.push_str(" & ");
            }
        }
        self.content.push_str("\\\\");
    }
    fn define_content(&mut self, content: Vec<Vec<String>>) {
        for line in content.iter() {
            self.content.push_str(&line.join(" & "));
            if line.len() < self.nb_col {
                self.content
                    .push_str(&" & ".repeat(self.nb_col - line.len()));
            }
            self.content.push_str("\\\\");
            self.content.push_str(&format!(
                "\n\\arrayrulecolor{{{}}}\\hline\n",
                self.line_color
            ));
        }
    }
    pub fn create_tabularx(&mut self) -> String {
        let mut tabularx = String::from(&format!(
            "\\begin{{tabularx}}{{\\textwidth}}{{{}}}\n",
            self.define_column()
        ));
        tabularx.push_str(&format!(
            "{}\n{}\n{}",
            self.title, self.sub_title, self.content
        ));
        tabularx.push_str("\\end{tabularx}");
        tabularx
        // tabularx =
    }
}
pub fn begin_section_wo_param(name: String, content: String) -> String {
    let section = &format!(
        "\\begin{{{}}}\n\
            {}\
        \\end{{{}}}",
        name, content, name
    );
    section.to_string()
}
