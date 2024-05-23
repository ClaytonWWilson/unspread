use calamine::{open_workbook, Reader, Xlsx};
use clap::{command, Parser};
use std::{
    error::Error,
    fs, io,
    path::{Component, Path, PathBuf},
    process::exit,
    vec,
};
use {directories::UserDirs, lazy_regex::*};

fn handle_exit(exit_code: i32, show_exit_message: bool) -> ! {
    if show_exit_message {
        println!("Press enter to exit");
        let mut buffer = String::new();
        let _ = io::stdin().read_line(&mut buffer);
    }
    exit(exit_code);
}

fn handle_message_output(message: &str, print_to_stdout: bool) {
    if print_to_stdout {
        println!("{}", message);
    }
}

#[derive(Debug)]
enum HeadersMode {
    Combine,
    Remove,
    Ignore,
}

impl From<u8> for HeadersMode {
    fn from(value: u8) -> Self {
        match value {
            0 => HeadersMode::Combine,
            1 => HeadersMode::Remove,
            2 => HeadersMode::Ignore,
            _ => HeadersMode::Combine,
        }
    }
}

/// Combine data spread across multiple spreadsheets into one
#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// The folder that contains the spreadsheet files
    /// Default './input'
    #[arg()]
    folder: Option<String>,

    /// Combined spreadsheet output file
    /// Default 'output.csv'
    #[arg(short = 'o', default_value_t = String::from("./output.csv"))]
    out_file: String,

    /// 0 = Combine headers, 1 = Remove headers, 2 = Ignore headers
    #[arg(short = 'm', long, default_value_t = 0)]
    headers_mode: u8,

    /// Suppress all output
    #[arg(short = 'q', long)]
    quiet: bool,

    /// Skip the `press enter to exit` prompt
    #[arg(short = 's', long)]
    skip_waiting: bool,
}

fn data_from_csv(path: PathBuf) -> Result<Vec<Vec<String>>, Box<dyn Error>> {
    // let reader = csv::Reader::from_path(path)?;
    let reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_path(path)?;
    let cells = reader
        .into_records()
        // for c in cells {
        //     println!("{:?}", c);
        // }
        .filter_map(|res| match res {
            Ok(row) => Some(row),
            Err(e) => {
                println!("{}", e.to_string());
                None
            }
        })
        .map(|row| {
            row.iter()
                .map(|cell| cell.to_string())
                .collect::<Vec<String>>()
        })
        .collect::<Vec<Vec<String>>>();

    Ok(cells)
    // Ok(vec![vec!["".to_string()]])
}

fn data_from_excel(path: PathBuf) -> Result<Vec<Vec<String>>, Box<dyn Error>> {
    let mut workbook: Xlsx<_> = match open_workbook(path) {
        Ok(s) => s,
        Err(e) => return Err(Box::new(e)),
    };

    let binding = workbook.sheet_names();
    let sheet_one_name = match binding.get(0) {
        Some(s) => s,
        None => return Err("Xslx file does not have any sheets".into()),
    };

    if workbook.sheet_names().len() > 1 {
        println!("Warning: A spreadsheet contains more than one sheet. This tool will only read the first sheet and ignore the rest.");
    }

    let sheet_one = workbook.worksheet_range(sheet_one_name)?;

    // println!("{:?}", sheet_one);
    let data = sheet_one
        .rows()
        .map(|r| {
            let cells = r.iter().map(|c| c.to_string());
            cells.collect::<Vec<String>>()
        })
        .collect::<Vec<Vec<String>>>();

    Ok(data)
}

/// https://stackoverflow.com/questions/68231306/stdfscanonicalize-for-files-that-dont-exist
/// build a usable path from a user input which may be absolute
/// (if it starts with / or ~) or relative to the supplied base_dir.
/// (we might want to try detect windows drives in the future, too)
pub fn path_from<P: AsRef<Path>>(base_dir: P, input: &str) -> PathBuf {
    let tilde = regex!(r"^~(/|$)");
    if input.starts_with('/') {
        // if the input starts with a `/`, we use it as is
        input.into()
    } else if tilde.is_match(input) {
        // if the input starts with `~` as first token, we replace
        // this `~` with the user home directory
        PathBuf::from(&*tilde.replace(input, |c: &Captures| {
            if let Some(user_dirs) = UserDirs::new() {
                format!("{}{}", user_dirs.home_dir().to_string_lossy(), &c[1],)
            } else {
                println!("no user dirs found, no expansion of ~");
                c[0].to_string()
            }
        }))
    } else {
        // we put the input behind the source (the selected directory
        // or its parent) and we normalize so that the user can type
        // paths with `../`
        normalize_path(base_dir.as_ref().join(input))
    }
}

/// Improve the path to try remove and solve .. token.
///
/// This assumes that `a/b/../c` is `a/c` which might be different from
/// what the OS would have chosen when b is a link. This is OK
/// for broot verb arguments but can't be generally used elsewhere
///
/// This function ensures a given path ending with '/' still
/// ends with '/' after normalization.
pub fn normalize_path<P: AsRef<Path>>(path: P) -> PathBuf {
    let ends_with_slash = path.as_ref().to_str().map_or(false, |s| s.ends_with('/'));
    let mut normalized = PathBuf::new();
    for component in path.as_ref().components() {
        match &component {
            Component::ParentDir => {
                if !normalized.pop() {
                    normalized.push(component);
                }
            }
            _ => {
                normalized.push(component);
            }
        }
    }
    if ends_with_slash {
        normalized.push("");
    }
    normalized
}

fn save_to_csv(data: &Vec<Vec<String>>, destination: &PathBuf) -> Result<(), Box<dyn Error>> {
    let mut writer = csv::Writer::from_path(destination)?;

    data.iter().for_each(|row| {
        match writer.write_record(row) {
            Ok(_) => {}
            Err(e) => {
                println!("{}", e.to_string());
            }
        };
    });

    Ok(())
}

fn main() {
    let args = Args::parse();
    let headers_mode = HeadersMode::from(args.headers_mode);

    let folder_name = match args.folder {
        Some(folder) => folder,
        None => {
            handle_message_output(
                "No folder provided, using default folder 'input'",
                !args.quiet,
            );
            "./input".to_string()
        }
    };

    let spreadsheet_folder = match Path::new(&folder_name).canonicalize() {
        Ok(path) => path,
        Err(e) => {
            handle_message_output(
                &format!(" '{}': {}", &folder_name, e.to_string()),
                !args.quiet,
            );
            // println!(" '{}': {}", &folder_name, e.to_string());
            handle_exit(1, !args.skip_waiting);
        }
    };

    if !spreadsheet_folder.is_dir() {
        handle_message_output("Argument for folder must be a valid folder", !args.quiet);
        // println!("Argument for folder must be a valid folder");
        handle_exit(1, !args.skip_waiting);
    }

    let output_file = path_from(".", &args.out_file);
    if output_file.is_dir() {
        handle_message_output("Output file cannot be a directory", !args.quiet);
        // println!("Output file cannot be a directory");
        handle_exit(1, !args.skip_waiting);
    }

    let dir = match fs::read_dir(spreadsheet_folder) {
        Ok(d) => d,
        Err(e) => {
            println!("Error opening the inputs folder: {}", e.to_string());
            handle_exit(1, !args.skip_waiting);
        }
    };

    let mut total_file_count = 0;

    let dir_entries = dir.filter_map(|x| {
        total_file_count += 1;
        x.ok()
    });

    let mut headers: Vec<String> = vec![];
    let mut combined_spreadsheet_data: Vec<Vec<String>> = vec![];

    for (index, entry) in dir_entries.enumerate() {
        // Filter out folders
        match entry.file_type() {
            Ok(f) => {
                if f.is_dir() {
                    println!(
                        "Skipping directory {}",
                        entry
                            .file_name()
                            .into_string()
                            .unwrap_or("`error`".to_string())
                    );
                    continue;
                }
            }
            Err(_) => {
                println!(
                    "Skipping directory {}",
                    entry
                        .file_name()
                        .into_string()
                        .unwrap_or("`error`".to_string())
                );
            }
        }

        let name_binding = entry.file_name();
        let file_name = name_binding.as_os_str().to_string_lossy();

        let spreadsheet_data = match {
            if file_name.ends_with(".csv") {
                data_from_csv(entry.path())
            } else if file_name.ends_with(".ods")
                | file_name.ends_with("xls")
                | file_name.ends_with("xlsx")
                | file_name.ends_with("xlsm")
                | file_name.ends_with("xlsb")
                | file_name.ends_with("xla")
                | file_name.ends_with("xlam")
            {
                data_from_excel(entry.path())
            } else {
                println!("Unsupported file type: {}", file_name);
                continue;
            }
        } {
            Ok(data) => data,
            Err(e) => {
                println!(
                    "Error occurred while reading file {} : {}",
                    file_name,
                    e.to_string()
                );
                continue;
            }
        };

        // println!("{:?}", spreadsheet_data);

        // Save headers on first spreadsheet
        if index == 0 {
            headers = match spreadsheet_data.get(0) {
                Some(h) => h.to_owned(),
                None => {
                    println!("Skipping {} since it's empty", file_name);
                    continue;
                }
            }
        }

        let first_row = match spreadsheet_data.get(0) {
            Some(data) => data,
            None => {
                println!("Skipping {} since it's empty", file_name);
                continue;
            }
        };

        let mut temp = vec![];
        let mut final_spreadsheet_data = match headers_mode {
            HeadersMode::Ignore => spreadsheet_data,
            HeadersMode::Remove => {
                if first_row.len() != headers.len() {
                    temp = spreadsheet_data;
                } else {
                    for header_pair in first_row.iter().zip(headers.iter()) {
                        if header_pair.0 != header_pair.1 {
                            temp = spreadsheet_data.clone();
                            break;
                        }
                    }

                    if temp.len() == 0 {
                        temp = spreadsheet_data.clone();
                        temp.remove(0);
                    }
                }
                temp
            }
            HeadersMode::Combine => {
                if index == 0 {
                    spreadsheet_data
                } else {
                    temp = spreadsheet_data.clone();
                    temp.remove(0);
                    temp
                }
            }
        };

        combined_spreadsheet_data.append(&mut final_spreadsheet_data)
    }

    match save_to_csv(&combined_spreadsheet_data, &output_file) {
        Ok(_) => {
            println!("Success: {} lines written", combined_spreadsheet_data.len());
            handle_exit(0, !args.skip_waiting)
        }
        Err(e) => {
            println!("Failure {}", e.to_string());
            handle_exit(1, !args.skip_waiting)
        }
    }
}

// TODO: Add option to ignore press enter to continue
