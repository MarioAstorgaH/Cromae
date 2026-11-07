# Prueba tecnica Cromae

###  Descripcion del proyecto
Este proyecto es una prueba tecnica para la empresa Cromae, que consiste en leer un archivo Excel con una lista de pagos, validar la información y generar un nuevo archivo Excel con los resultados de la validación.

### Requisitos
- Rust 1.70 o superior
- Cargo
- Dependencias del proyecto (se especifican en el archivo Cargo.toml):
  - calamine = "0.18.0"
  - xlsxwriter = "0.6.0"
- archivo Excel de entrada con la hoja "Pagos" y las columnas: Beneficiario, CLABE, Monto, Concepto.

### Uso
1. Clonar el repositorio.
2. Navegar al directorio del proyecto.
3. Ejecutar el comando `cargo run` para compilar y ejecutar el proyecto.
4. El programa leerá el archivo Excel de entrada, validará los pagos y generará un nuevo archivo Excel con los resultados de la validación.

### Uso opcional
- Ejecutar el archivo prueba_tecnica.exe para generar el reporte sin necesidad de compilar el proyecto.

### ¿Qué fue lo más difícil del reto y cómo lo resolviste?
Lo mas dificil del reto fue la validación de los datos, ya que se requería verificar que los datos cumplieran con ciertos criterios y generar un reporte con los resultados. Para resolverlo, utilicé la biblioteca calamine para leer el archivo Excel y la biblioteca xlsxwriter para generar el nuevo archivo Excel con los resultados de la validación. Además, implementé funciones de validación para cada campo. Tambien utilice una estructura de datos para almacenar los resultados y apegarme un poco a la programacion orientada a objetos para mantener el codigo organizado y facil de entender. Tambien utilice algunos conceptos de programación funcional para manejar los datos de manera eficiente ya que rust es un lenguaje que se presta para ello.
