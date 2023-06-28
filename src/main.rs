use calamine::{open_workbook, Reader, Xls};
use clap::{arg, Parser, command};
use serde::Serialize;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Caminho até o arquivo excel
    #[arg(short, long)]
    file_path: String,

    /// Nome da página no excel
    #[arg(short, long, default_value = "Imoveis")]
    page_name: String,

    /// Mostrar resposta reduzida
    #[arg(long)]
    short_output: bool,
}

#[derive(Default, Serialize, Debug)]
struct ParcialStats {
    #[serde(skip_serializing)]
    ids: Vec<u32>,
    name: String,
    total_raw: u32,
    total_unique: u32,
    total_duplicates: u32,
    total_ativo: u32,
    total_cancelado: u32,
    total_ficha: u32,
    total_locado: u32,
    total_provisorio: u32,
    total_suspenso: u32,
    total_vendido: u32,
}

fn sum_stats(items: &Vec<&ParcialStats>) -> ParcialStats {
    let mut output = ParcialStats::default();
    for item in items {
        let mut myids = item.ids.clone();
        output.ids.append(&mut myids);

        if output.name == "" {
            output.name = format!("{}", item.name);
        } else {
            output.name = format!("{}_{}", output.name, item.name);
        }

        output.total_duplicates += item.total_duplicates;
        output.total_raw += item.total_raw;
        output.total_unique += item.total_unique;

        output.total_ativo += item.total_ativo;
        output.total_cancelado += item.total_cancelado;
        output.total_ficha += item.total_ficha;
        output.total_locado += item.total_locado;
        output.total_provisorio += item.total_provisorio;
        output.total_suspenso += item.total_suspenso;
        output.total_vendido += item.total_vendido;
    }
    output
}

#[derive(Debug)]
enum PropertyType {
    Sale,
    Rent,
    SaleRent,
    Error,
}

impl ParcialStats {
    fn add_id(&mut self, value: u32) {
        self.ids.push(value)
    }

    fn add_status_computation(&mut self, value: String) {
        match value.as_str() {
            "Ativo" => self.total_ativo += 1,
            "Cancelado" => self.total_cancelado += 1,
            "Ficha" => self.total_ficha += 1,
            "Locado" => self.total_locado += 1,
            "Provisório" => self.total_provisorio += 1,
            "Suspenso" => self.total_suspenso += 1,
            "Vendido" => self.total_vendido += 1,
            _ => {}
        }
    }

    fn compute_data(&mut self) {
        let mut new_vec = self.ids.to_vec();
        new_vec.sort();
        new_vec.dedup();

        self.total_raw = self.ids.len() as u32;
        self.total_unique = new_vec.len() as u32;
        self.total_duplicates = (self.ids.len() - new_vec.len()) as u32;
    }
}

fn property_type(sale_price: f64, rent_price: f64) -> PropertyType {
    if sale_price <= 0.0 && rent_price > 0.0 {
        return PropertyType::Rent;
    }

    if sale_price > 0.0 && rent_price <= 0.0 {
        return PropertyType::Sale;
    }

    if sale_price > 0.0 && rent_price > 0.0 {
        return PropertyType::SaleRent;
    }

    PropertyType::Error
}

fn main() {
    let args = Args::parse();
    let file_path = args.file_path;
    let page_name = args.page_name;
    let is_short_output = args.short_output;

    let mut venda = ParcialStats {
        name: "sale".to_string(),
        ..Default::default()
    };
    let mut aluguel = ParcialStats {
        name: "rent".to_string(),
        ..Default::default()
    };
    let mut venda_aluguel = ParcialStats {
        name: "rent-sale".to_string(),
        ..Default::default()
    };
    let mut error = ParcialStats {
        name: "error".to_string(),
        ..Default::default()
    };

    let excel: Result<Xls<_>, _> = open_workbook(file_path);
    if let Err(value) = excel {
      panic!("{}", value)
    }

    if let Some(Ok(r)) = excel.unwrap().worksheet_range(&page_name) {
        for row in r.rows().skip(1) {
            let property_id = match row[0].as_i64() {
                Some(value) => value as u32,
                None => 0,
            };
            let valor_venda = match row[11].as_f64() {
                Some(value) => value,
                None => 0.0,
            };
            let valor_aluguel = match row[12].as_f64() {
                Some(value) => value,
                None => 0.0,
            };
            let valor_status = match row[31].as_string() {
                Some(value) => value,
                None => "n/a".to_string(),
            };

            match property_type(valor_venda, valor_aluguel) {
                PropertyType::Rent => {
                    aluguel.add_id(property_id);
                    aluguel.add_status_computation(valor_status);
                }
                PropertyType::Sale => {
                    venda.add_id(property_id);
                    venda.add_status_computation(valor_status);
                }
                PropertyType::Error => {
                    error.add_id(property_id);
                    error.add_status_computation(valor_status);
                }
                PropertyType::SaleRent => {
                    venda_aluguel.add_id(property_id);
                    venda_aluguel.add_status_computation(valor_status);
                }
            }
        }
        venda.compute_data();
        aluguel.compute_data();
        venda_aluguel.compute_data();
        error.compute_data();

        let total = sum_stats(&vec![&venda, &aluguel, &venda_aluguel, &error]);

     
        if is_short_output {
          println!(
              "\"{}\t{}\t{}\t{}\t{}\t{}\t{}\"",
              venda.total_raw,
              aluguel.total_raw,
              error.total_raw,
              venda_aluguel.total_raw,
              total.total_duplicates,
              "",
              total.total_ativo
            )
        } else {
          let output =
          serde_json::to_string_pretty(&vec![&total, &venda, &aluguel, &error, &venda_aluguel]);

          print!("{}\n", output.unwrap());
        }
    } else {
        panic!("Página não encontrada com o nome Imoveis")
    }
}
