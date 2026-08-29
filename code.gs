function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard de Finanzas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nombreArchivo) {
  return HtmlService
    .createHtmlOutputFromFile(nombreArchivo)
    .getContent();
}

function registrarTransaccion(datos) {

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName('REGISTRO');

  if (!hoja) {
    throw new Error("No se encontró la hoja 'REGISTRO'");
  }

  const montoNumerico = parseFloat(datos.monto);

  if (isNaN(montoNumerico)) {
    throw new Error("El monto no es válido");
  }

  hoja.appendRow([
    datos.fecha,
    datos.tipo,
    datos.categoria,
    datos.concepto,
    datos.vivienda,
    datos.cuenta,
    montoNumerico
  ]);

  return "Transacción registrada con éxito";
}

function obtenerEstructuraDesplegables() {

  const libro = SpreadsheetApp.getActiveSpreadsheet();

  const hojaDesplegables =
    libro.getSheetByName('DESPLEGABLES');

  if (!hojaDesplegables) {
    throw new Error(
      "No se encontró la hoja 'DESPLEGABLES'."
    );
  }

  const datos =
    hojaDesplegables
      .getDataRange()
      .getValues();


  const estructura = {

    "INGRESOS": {
      categorias: []
    },

    "AHORROS E INVERSIONES": {

      categorias: [],

      conceptosPorCategoria: {
        "AHORRO": [],
        "INVERSION": []
      }

    },

    "GASTOS": {

      categorias: [
        "FIJOS",
        "VARIABLES"
      ],

      conceptosPorCategoria: {
        "FIJOS": [],
        "VARIABLES": []
      },

      subConceptos: {

        "SERVICIOS PUBLICOS": [],

        "DEUDA": []

      }

    },

    "VIVIENDAS": [],

    "CUENTAS": []

  };


  for (let i = 1; i < datos.length; i++) {

    const fila = datos[i];


    // INGRESOS
    if (fila[2]) {
      estructura["INGRESOS"]
        .categorias
        .push(fila[2]);
    }


    // VIVIENDAS
    if (fila[6]) {
      estructura["VIVIENDAS"]
        .push(fila[6]);
    }


    // GASTOS FIJOS
    if (fila[8]) {
      estructura["GASTOS"]
        .conceptosPorCategoria["FIJOS"]
        .push(fila[8]);
    }


    // SERVICIOS PÚBLICOS
    if (fila[10]) {
      estructura["GASTOS"]
        .subConceptos["SERVICIOS PUBLICOS"]
        .push(fila[10]);
    }


    // GASTOS VARIABLES
    if (fila[12]) {
      estructura["GASTOS"]
        .conceptosPorCategoria["VARIABLES"]
        .push(fila[12]);
    }


    // CONCEPTOS DE AHORRO
    if (fila[14]) {
      estructura["AHORROS E INVERSIONES"]
        .conceptosPorCategoria["AHORRO"]
        .push(fila[14]);
    }


    // CUENTAS
    if (fila[16]) {
      estructura["CUENTAS"]
        .push(fila[16]);
    }


    // AHORRO / INVERSIÓN
    if (fila[18]) {
      estructura["AHORROS E INVERSIONES"]
        .categorias
        .push(fila[18]);
    }


    // DEUDAS
    if (fila[20]) {
      estructura["GASTOS"]
        .subConceptos["DEUDA"]
        .push(fila[20]);
    }

  }


  const limpiar = (array) => {

    return [
      ...new Set(array)
    ].filter(Boolean);

  };


  estructura["INGRESOS"].categorias =
    limpiar(
      estructura["INGRESOS"].categorias
    );


  estructura["VIVIENDAS"] =
    limpiar(
      estructura["VIVIENDAS"]
    );


  estructura["GASTOS"]
    .conceptosPorCategoria["FIJOS"] =
    limpiar(
      estructura["GASTOS"]
        .conceptosPorCategoria["FIJOS"]
    );


  estructura["GASTOS"]
    .conceptosPorCategoria["VARIABLES"] =
    limpiar(
      estructura["GASTOS"]
        .conceptosPorCategoria["VARIABLES"]
    );


  estructura["GASTOS"]
    .subConceptos["SERVICIOS PUBLICOS"] =
    limpiar(
      estructura["GASTOS"]
        .subConceptos["SERVICIOS PUBLICOS"]
    );


  estructura["GASTOS"]
    .subConceptos["DEUDA"] =
    limpiar(
      estructura["GASTOS"]
        .subConceptos["DEUDA"]
    );


  estructura["AHORROS E INVERSIONES"]
    .categorias =
    limpiar(
      estructura["AHORROS E INVERSIONES"]
        .categorias
    );


  estructura["AHORROS E INVERSIONES"]
    .conceptosPorCategoria["AHORRO"] =
    limpiar(
      estructura["AHORROS E INVERSIONES"]
        .conceptosPorCategoria["AHORRO"]
    );


  estructura["CUENTAS"] =
    limpiar(
      estructura["CUENTAS"]
    );


  return estructura;
}

