use clap::Parser;

mod cli;
mod com;
mod commands;
mod config;
mod error;
mod office;
mod pipeline;
mod shapes;
mod utils;
mod zip_ops;

use cli::{Cli, Commands};

fn main() {
    let cli = Cli::parse();

    let result: Result<(), String> = match cli.command {
        Commands::Update(ref args) => {
            commands::update::run_update(args).map_err(|e| e.to_string())
        }
        Commands::Run(args) => {
            commands::run::run_runfile(
                &args.runfile,
                args.check,
                args.dry_run,
                args.verbose,
                args.quiet,
            )
            .map_err(|e| e.to_string())
        }
        Commands::Check(ref args) => {
            let mut config = config::Config::default();
            if let Err(e) = config.apply_overrides(&args.set) {
                return eprintln!("Error: {e}");
            }
            match commands::check::run_check(&args.file, args.excel.as_deref(), &config) {
                Ok(result) => {
                    if !result.passed() {
                        std::process::exit(1);
                    }
                    Ok(())
                }
                Err(e) => Err(e.to_string()),
            }
        }
        Commands::Diff(args) => {
            commands::diff::run_diff(&args.file_a, &args.file_b)
                .map(|_| ())
                .map_err(|e| e.to_string())
        }
        Commands::Info(args) => {
            commands::info::run_info(&args.file).map_err(|e| e.to_string())
        }
        Commands::Clean(args) => {
            commands::clean::run_clean(args.force).map_err(|e| e.to_string())
        }
        Commands::Config => {
            commands::config_cmd::run_config();
            Ok(())
        }
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        std::process::exit(2);
    }
}
