use calamine::{Data, DataType, Reader, Xlsx, open_workbook};
use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};
#[derive(Debug)]
// Estructura que representa un pago
struct Pago {
    beneficiario: String,
    clabe: String,
    importe: f64,
    concepto: String,
    estado: Estado,
    motivo_error: Option<String>,
}
#[derive(Debug)]
// Enumeración que representa el estado de un pago
enum Estado {
    Valido,
    Invalido,
}
impl Pago {
    // constructor que crea un pago a partir de una fila de datos
    fn desde_fila(fila: &[Data]) -> Self {
        Self {
            beneficiario: fila.get(0).map(|c| c.to_string()).unwrap_or_default(),
            clabe: fila
                .get(1)
                .map(|c| {
                    // Eliminar espacios en blanco de la CLABE
                    c.to_string()
                        .chars()
                        .filter(|letra| !letra.is_whitespace())
                        .collect()
                })
                .unwrap_or_default(),
            importe: fila.get(2).and_then(|c| c.as_f64()).unwrap_or(f64::NAN),
            concepto: fila.get(3).map(|c| c.to_string()).unwrap_or_default(),
            estado: Estado::Valido,
            motivo_error: None,
        }
    }
    fn verificar_beneficiario(&mut self) {
        if self.beneficiario.is_empty() {
            self.motivo_error = Some("Beneficiario vacío".to_string());
        }
    }
    fn verificar_clabe(&mut self) {
        if self.clabe.is_empty() {
            self.motivo_error = Some("CLABE vacía".to_string());
            return;
        }
        if !self.clabe.chars().all(|c| c.is_digit(10)) {
            self.motivo_error = Some("CLABE inválida contiene caracteres no numéricos".to_string());
            return;
        }
        if self.clabe.len() != 18 {
            self.motivo_error = Some(format!(
                "CLABE inválida tiene {} caracteres",
                self.clabe.len()
            ));
            return;
        }
    }
    fn verificar_importe(&mut self) {
        if self.importe.is_nan() {
            self.motivo_error = Some("Importe no es un número".to_string());
            return;
        }
        if self.importe <= 0.0 {
            self.motivo_error = Some("Importe inválido es menor que 0".to_string());
            return;
        }
    }
    fn validar(&mut self) {
        self.verificar_beneficiario();
        self.verificar_clabe();
        self.verificar_importe();
        self.estado = if self.motivo_error.is_some() {
            Estado::Invalido
        } else {
            Estado::Valido
        };
    }
}
fn escribir_reporte(pagos: &[Pago], ruta: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let hoja = workbook.add_worksheet();
    hoja.set_name("Reporte")?;

    // Formatos
    let encabezado = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xD9E1F2));
    let moneda = Format::new().set_num_format("$#,##0.00");
    let invalido = Format::new().set_background_color(Color::RGB(0xFFC7CE));
    let moneda_invalido = Format::new()
        .set_num_format("$#,##0.00")
        .set_background_color(Color::RGB(0xFFC7CE));

    // Encabezados
    let titulos = [
        "Beneficiario",
        "CLABE",
        "Importe",
        "Concepto",
        "Estado",
        "Motivo de error",
    ];
    for (col, titulo) in titulos.iter().enumerate() {
        hoja.write_string_with_format(0, col as u16, *titulo, &encabezado)?;
    }

    // Datos
    for (i, pago) in pagos.iter().enumerate() {
        let fila = (i + 1) as u32; // +1 porque la fila 0 es el encabezado

        let es_invalido = matches!(pago.estado, Estado::Invalido);
        let fmt_texto = if es_invalido {
            &invalido
        } else {
            &Format::new()
        };
        let fmt_moneda = if es_invalido {
            &moneda_invalido
        } else {
            &moneda
        };

        hoja.write_string_with_format(fila, 0, &pago.beneficiario, fmt_texto)?;
        hoja.write_string_with_format(fila, 1, &pago.clabe, fmt_texto)?; // como texto para no perder dígitos

        if pago.importe.is_nan() {
            hoja.write_string_with_format(fila, 2, "", fmt_texto)?;
        } else {
            hoja.write_number_with_format(fila, 2, pago.importe, fmt_moneda)?;
        }

        hoja.write_string_with_format(fila, 3, &pago.concepto, fmt_texto)?;

        // Escribir el estado como texto
        let estado = match pago.estado {
            Estado::Valido => "Válido",
            Estado::Invalido => "Inválido",
        };
        hoja.write_string_with_format(fila, 4, estado, fmt_texto)?;

        let motivo = pago.motivo_error.as_deref().unwrap_or("");
        hoja.write_string_with_format(fila, 5, motivo, fmt_texto)?;
    }

    hoja.autofit(); // ajusta el ancho de las columnas al contenido
    workbook.save(ruta)?;
    Ok(())
}
fn main() {
    // Abrir el archivo Excel y leer la hoja "Pagos"
    let mut workbook: Xlsx<_> = open_workbook("./datos.xlsx").expect("No se puede leer el archivo");
    let range = workbook
        .worksheet_range("Pagos")
        .expect("No existe la hoja Pagos");
    //crear un vector(array) de pagos a partir de las filas del rango
    let mut pagos: Vec<Pago> = Vec::new();
    // Iterar sobre las filas del rango, omitiendo la primera fila (encabezados)
    for (indice, fila) in range.rows().enumerate().skip(1) {
        let mut pago = Pago::desde_fila(fila);
        pago.validar();
        println!("Fila num {}: {:?}", indice, pago);
        pagos.push(pago);
    }
    escribir_reporte(&pagos, "./reporte.xlsx").expect("No se pudo escribir el reporte");
    println!("Reporte generado con {} pagos", pagos.len());
}
