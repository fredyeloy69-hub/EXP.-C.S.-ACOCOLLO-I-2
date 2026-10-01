async function generarListaGeneralExcel() {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Lista General");

    // ============================================================
    // CONFIGURACIÓN DE COLUMNAS
    // ============================================================
    sheet.columns = [
      { width: 14 }, // N°
      { width: 75 }, // DESCRIPCIÓN
      { width: 22 }, // N° DE ARCHIVADOR
    ];

    // ============================================================
    // ENCABEZADO INSTITUCIONAL
    // ============================================================
    sheet.mergeCells("A1:C1");
    sheet.mergeCells("A2:C2");
    sheet.mergeCells("A3:C3");
    sheet.mergeCells("A4:C4");

    sheet.getCell("A1").value = "LISTA GENERAL DE DOCUMENTOS";
    sheet.getCell("A2").value = "EXPEDIENTE TÉCNICO";
    sheet.getCell("A3").value =
      `ÁREA: ${areaSeleccionada?.nombre || ""}`;

    sheet.getCell("A4").value =
      "La columna N° DE ARCHIVADOR queda en blanco para ser llenada manualmente.";

    ["A1", "A2", "A3", "A4"].forEach((direccion) => {
      const celda = sheet.getCell(direccion);

      celda.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
    });

    sheet.getCell("A1").font = {
      bold: true,
      size: 14,
    };

    sheet.getCell("A2").font = {
      bold: true,
      size: 12,
    };

    sheet.getCell("A3").font = {
      bold: true,
      size: 11,
    };

    sheet.getCell("A4").font = {
      italic: true,
      size: 10,
    };

    // ============================================================
    // ENCABEZADO DE LA TABLA
    // ============================================================
    const FILA_ENCABEZADO = 6;
    const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);

    filaEncabezado.values = [
      "N°",
      "DESCRIPCIÓN",
      "N° DE ARCHIVADOR",
    ];

    filaEncabezado.height = 28;

    filaEncabezado.eachCell((celda) => {
      celda.font = {
        bold: true,
        color: { argb: "FFFFFFFF" },
      };

      celda.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };

      celda.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1F4E78" },
      };

      celda.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // ============================================================
    // OBTENER LA MISMA JERARQUÍA DE LA LISTA PARA SEPARADORES
    // ============================================================
    const carpetasDelArea =
      obtenerCarpetasDelArea(areaSeleccionada);

    const arbol =
      construirArbolSeparadores(carpetasDelArea);

    const items = [];

    generarFilasSeparadores(
      arbol,
      1,
      items
    );

    // ============================================================
    // GENERAR FILAS
    // ============================================================
    for (const item of items) {
      const fila = sheet.addRow([
        item.numero,
        item.nombre,
        "",
      ]);

      // ----------------------------------------------------------
      // ALINEACIÓN
      // ----------------------------------------------------------
      fila.getCell(1).alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };

      fila.getCell(2).alignment = {
        horizontal: "left",
        vertical: "middle",
        wrapText: true,
      };

      fila.getCell(3).alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };

      // ----------------------------------------------------------
      // COLOR SEGÚN NIVEL JERÁRQUICO
      // ----------------------------------------------------------
      if (item.nivel <= 2) {
        const fondo =
          colorNivelSeparador(item.nivel);

        [1, 2].forEach((col) => {
          fila.getCell(col).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: {
              argb: fondo,
            },
          };
        });
      }

      // ----------------------------------------------------------
      // COLUMNA N° DE ARCHIVADOR
      // Queda vacía y resaltada para llenado manual
      // ----------------------------------------------------------
      fila.getCell(3).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FFFFF2CC",
        },
      };

      // ----------------------------------------------------------
      // NIVEL PRINCIPAL EN NEGRITA
      // ----------------------------------------------------------
      if (item.nivel === 1) {
        [1, 2].forEach((col) => {
          fila.getCell(col).font = {
            bold: true,
          };
        });
      }

      // ----------------------------------------------------------
      // BORDES
      // ----------------------------------------------------------
      [1, 2, 3].forEach((col) => {
        fila.getCell(col).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    }

    // ============================================================
    // CONFIGURACIÓN DE VISTA
    // ============================================================
    sheet.views = [
      {
        state: "frozen",
        ySplit: FILA_ENCABEZADO,
      },
    ];

    // ============================================================
    // CONFIGURACIÓN DE IMPRESIÓN
    // ============================================================
    sheet.pageSetup = {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9,
    };

    sheet.pageMargins = {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2,
    };

    // ============================================================
    // PIE DE PÁGINA
    // ============================================================
    sheet.headerFooter.oddFooter.center.text =
      "Sistema ACASO I-2";

    // ============================================================
    // NOMBRE DEL ARCHIVO
    // ============================================================
    const nombreArea =
      (areaSeleccionada?.nombre || "GENERAL")
        .replace(/[\\/:*?"<>|]/g, "_")
        .trim();

    const nombreArchivo =
      `ListaGeneral_${nombreArea}.xlsx`;

    // ============================================================
    // GENERAR Y DESCARGAR EXCEL
    // ============================================================
    const buffer =
      await workbook.xlsx.writeBuffer();

    const blob = new Blob(
      [buffer],
      {
        type:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    );

    const url =
      URL.createObjectURL(blob);

    const enlace =
      document.createElement("a");

    enlace.href = url;
    enlace.download = nombreArchivo;

    document.body.appendChild(enlace);

    enlace.click();

    enlace.remove();

    URL.revokeObjectURL(url);

  } catch (error) {
    console.error(
      "Error al generar Lista General:",
      error
    );

    alert(
      `No se pudo generar la Lista General: ${
        error?.message || error
      }`
    );
  }
}
