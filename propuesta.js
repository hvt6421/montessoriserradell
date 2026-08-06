const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, WidthType, BorderStyle, ShadingType, HeadingLevel } = require('docx');
const fs = require('fs');

const azul = "1F5C99";
const azulClaro = "D6E4F0";
const gris = "595959";
const blanco = "FFFFFF";

const borde = { style: BorderStyle.SINGLE, size: 6, color: azul };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const bordeHeader = { style: BorderStyle.SINGLE, size: 8, color: azul };
const bordesHeader = { top: bordeHeader, bottom: bordeHeader, left: bordeHeader, right: bordeHeader };

function celHeader(texto, ancho) {
  return new TableCell({
    borders: bordesHeader,
    width: { size: ancho, type: WidthType.DXA },
    shading: { fill: azul, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text: texto, bold: true, color: blanco, size: 22 })]
    })]
  });
}

function celNormal(texto, ancho, shading = false, bold = false) {
  return new TableCell({
    borders: bordes,
    width: { size: ancho, type: WidthType.DXA },
    shading: shading ? { fill: azulClaro, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text: texto, bold, size: 22 })]
    })]
  });
}

function celDerecha(texto, ancho, shading = false, bold = false) {
  return new TableCell({
    borders: bordes,
    width: { size: ancho, type: WidthType.DXA },
    shading: shading ? { fill: azulClaro, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: texto, bold, size: 22 })]
    })]
  });
}

function bullet(texto) {
  return new Paragraph({
    spacing: { after: 80 },
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text: texto, size: 22 })]
  });
}

function seccion(texto) {
  return new Paragraph({
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: azul } },
    children: [new TextRun({ text: texto, bold: true, size: 28, color: azul })]
  });
}

function separador() {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun("")] });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: "bullet",
        text: "\u2713",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 560, hanging: 320 } } }
      }]
    }, {
      reference: "numeros",
      levels: [{
        level: 0,
        format: "decimal",
        text: "%1.",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 560, hanging: 320 } } }
      }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1440, bottom: 1080, left: 1440 }
      }
    },
    children: [

      // ENCABEZADO
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "PROPUESTA DE ENTREGA", bold: true, size: 40, color: azul })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "Landing Page — Montessori Serradell", bold: true, size: 28, color: "1A1A1A" })]
      }),
      new Paragraph({
        spacing: { after: 320 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: azul } },
        children: [new TextRun({ text: "Navojoa, Sonora — 2026", size: 22, color: gris })]
      }),

      // PROYECTO EN VIVO
      new Paragraph({
        spacing: { before: 240, after: 80 },
        shading: { fill: "EBF3FB", type: ShadingType.CLEAR },
        children: [
          new TextRun({ text: "PROYECTO COMPLETADO Y EN VIVO  ", bold: true, size: 24, color: azul }),
          new TextRun({ text: "https://montessori-serradell.vercel.app", size: 22, color: "2980B9" })
        ]
      }),
      separador(),

      // SECCIÓN 1
      seccion("1. Lo que Incluye el Sitio"),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: "7 secciones desarrolladas:", bold: true, size: 22 })]
      }),
      bullet("Hero — Frase de impacto, 2 CTAs, rating y datos clave del negocio"),
      bullet("Metodo Montessori — Explicacion del metodo con 4 tarjetas de beneficios"),
      bullet("Servicios (4 niveles) — Maternal (1.5-3 anos), Preescolar (3-6 anos), Vida Practica, Primaria Taller (6 anos en adelante)"),
      bullet("Por que elegirnos — 6 puntos diferenciadores del negocio"),
      bullet("Testimonios — 2 resenas reales de familias, calificacion 4.5 estrellas"),
      bullet("Contacto — WhatsApp directo, direccion, horario, telefono, redes sociales"),
      bullet("Feed de Facebook en vivo — Publicaciones recientes actualizadas automaticamente"),
      separador(),

      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: "Integrado:", bold: true, size: 22 })]
      }),
      bullet("Logo oficial en navbar"),
      bullet("WhatsApp flotante con mensaje predefinido al numero 642 147 8457"),
      bullet("Facebook, Instagram y correo con iconos en footer"),
      bullet("Link directo a Google Maps"),
      bullet("Meta tags para compartir el link en WhatsApp/redes con vista previa"),
      separador(),

      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: "Tecnico:", bold: true, size: 22 })]
      }),
      bullet("100% responsive (desktop, tablet, movil)"),
      bullet("Animaciones optimizadas (GPU, sin lag en movil)"),
      bullet("Carga < 2 segundos"),
      bullet("HTTPS con certificado SSL incluido"),
      bullet("SEO basico listo"),
      bullet("Hosting en Vercel (gratuito, sin expiracion)"),
      bullet("Respaldo del codigo en GitHub"),
      separador(),

      // SECCIÓN 2 — INVERSIÓN
      seccion("2. Inversion"),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [6552, 2808],
        rows: [
          new TableRow({ children: [
            celHeader("Concepto", 6552),
            celHeader("Costo", 2808)
          ]}),
          new TableRow({ children: [
            celNormal("Paquete Entregado", 6552, false, true),
            celDerecha("$25,000 MXN", 2808, false, true)
          ]}),
          new TableRow({ children: [
            celNormal("   Sitio completo desarrollado y en produccion", 6552, true),
            celDerecha("", 2808, true)
          ]}),
          new TableRow({ children: [
            celNormal("   Hosting Vercel (gratis, sin limite de tiempo)", 6552, true),
            celDerecha("", 2808, true)
          ]}),
          new TableRow({ children: [
            celNormal("   Codigo en GitHub (respaldo privado)", 6552, true),
            celDerecha("", 2808, true)
          ]}),
          new TableRow({ children: [
            celNormal("   Soporte tecnico por 3 meses (incluido)", 6552, true),
            celDerecha("", 2808, true)
          ]}),
          new TableRow({ children: [
            celNormal("Dominio .com.mx (opcional)", 6552, false, true),
            celDerecha("$200 MXN/ano", 2808, false, true)
          ]}),
          new TableRow({ children: [
            celNormal("   montessoriserradell.com.mx — mejora imagen y SEO", 6552, true),
            celDerecha("", 2808, true)
          ]}),
          new TableRow({ children: [
            celNormal("Cambios por fuera del paquete", 6552, false, true),
            celDerecha("A consultar", 2808, false, true)
          ]}),
          new TableRow({ children: [
            celNormal("   Cambios de contenido, actualizaciones, nuevas secciones", 6552, true),
            celDerecha("", 2808, true)
          ]}),
        ]
      }),
      separador(),

      // SECCIÓN 3 — FORMA DE PAGO
      seccion("3. Forma de Pago"),
      new Paragraph({
        spacing: { after: 80 },
        numbering: { reference: "numeros", level: 0 },
        children: [new TextRun({ text: "Opcion A — Pago unico: $25,000 MXN", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        numbering: { reference: "numeros", level: 0 },
        children: [new TextRun({ text: "Opcion B — En dos pagos: $12,500 MXN al confirmar + $12,500 MXN al lanzamiento", size: 22 })]
      }),
      separador(),

      // SECCIÓN 4 — CONTACTO
      seccion("4. Contacto"),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "Email: hiramverduzco12@gmail.com", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "Demo en vivo: https://montessori-serradell.vercel.app", size: 22, color: "2980B9" })]
      }),
      separador(),

      // FOOTER
      new Paragraph({
        spacing: { before: 200 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: azul } },
        children: [new TextRun({ text: "Propuesta valida por 15 dias.", italics: true, size: 20, color: gris })]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Propuesta_Montessori_Serradell.docx", buffer);
  console.log("OK");
}).catch(e => console.error(e));
