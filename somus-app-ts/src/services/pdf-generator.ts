import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
import type { TDocumentDefinitions, Content } from 'pdfmake/interfaces';

// @ts-ignore
pdfMake.vfs = pdfFonts.pdfMake ? pdfFonts.pdfMake.vfs : pdfFonts.vfs;

const SOMUS_GREEN = '#004D33';
const HEADER_BG = '#F0F0E0';

function formatCurrencyPDF(value: number): string {
  return `R$ ${value.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function formatDatePDF(dateStr: string): string {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('pt-BR');
}

/**
 * Gera PDF do Fluxo RF para um assessor com seus eventos
 */
export async function generateFluxoRFPDF(
  assessor: string,
  events: Array<{
    ativo: string;
    tipoEvento: string;
    data: string;
    valor: number;
    descricao?: string;
  }>
): Promise<Blob> {
  const tableBody: any[][] = [
    [
      { text: 'Ativo', style: 'tableHeader' },
      { text: 'Tipo Evento', style: 'tableHeader' },
      { text: 'Data', style: 'tableHeader' },
      { text: 'Valor', style: 'tableHeader', alignment: 'right' },
    ],
  ];

  events.forEach((evt) => {
    tableBody.push([
      evt.ativo,
      evt.tipoEvento,
      formatDatePDF(evt.data),
      { text: formatCurrencyPDF(evt.valor), alignment: 'right' },
    ]);
  });

  const totalValor = events.reduce((sum, e) => sum + e.valor, 0);
  tableBody.push([
    { text: 'TOTAL', bold: true, colSpan: 3 },
    {},
    {},
    {
      text: formatCurrencyPDF(totalValor),
      bold: true,
      alignment: 'right',
    },
  ]);

  const docDefinition: TDocumentDefinitions = {
    pageSize: 'A4',
    pageMargins: [40, 80, 40, 60],
    header: {
      columns: [
        {
          text: 'SOMUS CAPITAL',
          style: 'headerTitle',
          margin: [40, 20, 0, 0],
        },
        {
          text: `Fluxo Renda Fixa - ${assessor}`,
          alignment: 'right',
          style: 'headerSubtitle',
          margin: [0, 25, 40, 0],
        },
      ],
    },
    footer: (currentPage: number, pageCount: number) => ({
      text: `Pagina ${currentPage} de ${pageCount}`,
      alignment: 'center',
      margin: [0, 20, 0, 0],
      fontSize: 8,
      color: '#888888',
    }),
    content: [
      {
        text: `Agenda de Eventos - ${assessor}`,
        style: 'title',
        margin: [0, 0, 0, 15],
      },
      {
        text: `Data de geracao: ${new Date().toLocaleDateString('pt-BR')}`,
        style: 'subtitle',
        margin: [0, 0, 0, 20],
      },
      {
        table: {
          headerRows: 1,
          widths: ['*', 'auto', 'auto', 'auto'],
          body: tableBody,
        },
        layout: {
          hLineWidth: () => 0.5,
          vLineWidth: () => 0.5,
          hLineColor: () => '#E5E7EB',
          vLineColor: () => '#E5E7EB',
          fillColor: (rowIndex: number) =>
            rowIndex === 0 ? HEADER_BG : null,
          paddingLeft: () => 8,
          paddingRight: () => 8,
          paddingTop: () => 6,
          paddingBottom: () => 6,
        },
      },
    ],
    styles: {
      headerTitle: {
        fontSize: 14,
        bold: true,
        color: SOMUS_GREEN,
      },
      headerSubtitle: {
        fontSize: 10,
        color: '#888888',
      },
      title: {
        fontSize: 18,
        bold: true,
        color: SOMUS_GREEN,
      },
      subtitle: {
        fontSize: 10,
        color: '#888888',
      },
      tableHeader: {
        bold: true,
        fontSize: 9,
        color: SOMUS_GREEN,
      },
    },
    defaultStyle: {
      font: 'Roboto',
      fontSize: 9,
      color: '#111111',
    },
  };

  return new Promise((resolve) => {
    const pdfDoc = pdfMake.createPdf(docDefinition);
    pdfDoc.getBlob((blob: Blob) => resolve(blob));
  });
}

/**
 * Gera um relatorio PDF generico com titulo e conteudo
 */
export async function generateReportPDF(
  title: string,
  content: Content
): Promise<Blob> {
  const docDefinition: TDocumentDefinitions = {
    pageSize: 'A4',
    pageMargins: [40, 80, 40, 60],
    header: {
      text: 'SOMUS CAPITAL',
      style: 'headerTitle',
      margin: [40, 20, 0, 0],
    },
    footer: (currentPage: number, pageCount: number) => ({
      text: `Pagina ${currentPage} de ${pageCount}`,
      alignment: 'center',
      margin: [0, 20, 0, 0],
      fontSize: 8,
      color: '#888888',
    }),
    content: [
      { text: title, style: 'title', margin: [0, 0, 0, 20] },
      content,
    ],
    styles: {
      headerTitle: { fontSize: 14, bold: true, color: SOMUS_GREEN },
      title: { fontSize: 18, bold: true, color: SOMUS_GREEN },
    },
    defaultStyle: {
      font: 'Roboto',
      fontSize: 10,
      color: '#111111',
    },
  };

  return new Promise((resolve) => {
    const pdfDoc = pdfMake.createPdf(docDefinition);
    pdfDoc.getBlob((blob: Blob) => resolve(blob));
  });
}

/**
 * Gera PDF de consorcio
 */
export async function generateConsorcioPDF(data: {
  clienteNome: string;
  valorCarta: number;
  prazo: number;
  parcelas: Array<{ mes: number; valor: number }>;
}): Promise<Blob> {
  const tableBody: any[][] = [
    [
      { text: 'Mes', style: 'tableHeader' },
      { text: 'Valor Parcela', style: 'tableHeader', alignment: 'right' },
    ],
  ];

  data.parcelas.forEach((p) => {
    tableBody.push([
      p.mes.toString(),
      { text: formatCurrencyPDF(p.valor), alignment: 'right' },
    ]);
  });

  const content: Content = [
    {
      text: `Proposta de Consorcio - ${data.clienteNome}`,
      style: 'title',
      margin: [0, 0, 0, 15],
    },
    {
      columns: [
        {
          text: `Valor da Carta: ${formatCurrencyPDF(data.valorCarta)}`,
          width: '*',
        },
        { text: `Prazo: ${data.prazo} meses`, width: '*' },
      ],
      margin: [0, 0, 0, 20],
    },
    {
      table: {
        headerRows: 1,
        widths: ['auto', '*'],
        body: tableBody,
      },
      layout: {
        hLineWidth: () => 0.5,
        vLineWidth: () => 0.5,
        hLineColor: () => '#E5E7EB',
        vLineColor: () => '#E5E7EB',
        fillColor: (rowIndex: number) =>
          rowIndex === 0 ? HEADER_BG : null,
      },
    },
  ];

  return generateReportPDF('Proposta de Consorcio', content);
}
