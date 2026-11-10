import { Fragment, type ReactNode, useEffect, useMemo, useState } from 'react';
import {
  Download,
  FileText,
  Info,
  GripVertical,
  ArrowDown,
  ArrowUp,
  Printer,
  RotateCcw,
  Save,
} from 'lucide-react';
import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  type IParagraphOptions,
  PageBreak,
  Paragraph,
  Packer,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

type TableElement = { id: string; title: string; columns: string[]; rows: string[][] };
type HighlightElement = { id: string; title: string; text: string };
type TextSection = { id: string; kicker: string; title: string; lead: string; body: string; tables: TableElement[]; highlights: HighlightElement[] };

type PublicationCopy = {
  title: string;
  subtitle: string;
  coverKicker: string;
  author: string;
  position: string;
  organization: string;
  reportingPeriod: string;
  publicationDate: string;
  intro: string;
  introLabel: string;
  contextHeading: string;
  contextLabel: string;
  contextLead: string;
  contextBody: string;
  comparisonTitle: string;
  marginHeading: string;
  marginNote: string;
  implicationsHeading: string;
  implicationsLabel: string;
  implicationsLead: string;
  implicationsBody: string;
  reviewPoints: string[];
  callout: string;
  calloutTitle: string;
  reviewTitle: string;
  chartTitle: string;
  chartLabel: string;
  chartSummary: string;
  chartUnit: string;
  chartData: { label: string; value: string }[];
  dataTableTitle: string;
  dataTableLabel: string;
  dataTableSummary: string;
  dataTableColumns: string[];
  dataTableRows: string[][];
  textSections: TextSection[];
  hiddenPages: string[];
  pageOrder: string[];
  sectionOrder: ('context' | 'implications')[];
};

const STORAGE_KEY = 'ifrs-metrics-publication-studio-v1';

function styledParagraph(
  text: string,
  run: Record<string, unknown>,
  options: Omit<IParagraphOptions, 'children'> = {},
) {
  return new Paragraph({ ...options, children: [new TextRun({ text, ...run })] });
}

const SAMPLE_COPY: PublicationCopy = {
  title: 'Interpreting\nchange',
  subtitle: 'A practical reporting note for finance teams',
  coverKicker: 'Perspective · Finance teams',
  author: '',
  position: '',
  organization: '',
  reportingPeriod: '',
  publicationDate: '',
  intro:
    'A concise framework for reading change, understanding its context, and keeping review conversations clear.',
  introLabel: 'In this note',
  contextHeading: 'What is changing',
  contextLabel: 'Context',
  contextLead:
    'A useful starting point is to describe the movement before assigning it meaning.',
  contextBody:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.\n\nLorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.',
  comparisonTitle: 'A simple comparison',
  marginHeading: 'Reading the movement',
  marginNote:
    'Clear explanation begins with a shared view of what moved.',
  implicationsHeading: 'What it means',
  implicationsLabel: 'Application',
  implicationsLead:
    'Keep the narrative proportionate: distinguish the observed change from the explanation around it.',
  implicationsBody:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.\n\nLorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.',
  reviewPoints: [
    'Describe the movement. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.',
    'Make the context visible. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.',
    'Keep the trail reviewable. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Vivamus non augue at erat posuere commodo. Donec feugiat, lectus sed facilisis dictum, nibh eros tincidunt mi, in luctus neque risus sed erat.',
  ],
  callout:
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer vitae justo sed urna posuere volutpat. Donec feugiat, lectus sed facilisis dictum.',
  calloutTitle: 'Working principle',
  reviewTitle: 'Review points',
  chartTitle: 'Change across reporting areas',
  chartLabel: 'Data snapshot',
  chartSummary: 'A sample view of how selected indicators moved during the reporting period.',
  chartUnit: '%',
  chartData: [
    { label: 'Revenue', value: '12' },
    { label: 'Operating costs', value: '-4' },
    { label: 'Total assets', value: '8' },
    { label: 'Cash flow', value: '5' },
  ],
  dataTableTitle: 'Selected indicators',
  dataTableLabel: 'Data table',
  dataTableSummary: 'Illustrative values for a reporting-period comparison.',
  dataTableColumns: ['Indicator', 'Current', 'Prior', 'Movement'],
  dataTableRows: [
    ['Revenue', '1,240', '1,108', '+11.9%'],
    ['Operating costs', '820', '854', '−4.0%'],
    ['Total assets', '3,680', '3,405', '+8.1%'],
    ['Operating cash flow', '412', '392', '+5.1%'],
  ],
  textSections: [{
    id: 'sample-text-1',
    kicker: 'Further reading',
    title: 'A clear narrative',
    lead: 'Connect each reported movement to its source and explain what it means for readers.',
    body: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. A useful report gives readers enough context to understand the movement, the assumptions behind it, and the decisions that follow.\n\nUse this page as a reusable text section. Add more pages from the editor, change each heading and paragraph, and remove sections that do not belong in your publication.',
    tables: [],
    highlights: [],
  }],
  hiddenPages: [],
  sectionOrder: ['context', 'implications'],
  pageOrder: ['cover', 'context', 'implications', 'chart', 'table', 'sample-text-1'],
};

function readSavedCopy(): PublicationCopy {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return SAMPLE_COPY;
    const parsed = JSON.parse(saved) as Partial<PublicationCopy>;
    const dataTableColumns = Array.isArray(parsed.dataTableColumns) ? parsed.dataTableColumns : SAMPLE_COPY.dataTableColumns;
    const savedRows = parsed.dataTableRows as unknown[] | undefined;
    const dataTableRows = savedRows?.map((row) => Array.isArray(row)
      ? row.map((cell) => String(cell ?? ''))
      : row && typeof row === 'object'
        ? Object.values(row as Record<string, unknown>).map((cell) => String(cell ?? ''))
        : []) ?? SAMPLE_COPY.dataTableRows;
    const validOrder = Array.isArray(parsed.sectionOrder)
      ? parsed.sectionOrder.filter((section): section is 'context' | 'implications' =>
          section === 'context' || section === 'implications',
        )
      : [];
    const sectionOrder = validOrder.length === 2 && new Set(validOrder).size === 2
      ? validOrder
      : [...SAMPLE_COPY.sectionOrder];
    const textSections = Array.isArray(parsed.textSections) ? parsed.textSections.map((section) => ({
      ...section,
      tables: Array.isArray(section.tables) ? section.tables : [],
      highlights: Array.isArray(section.highlights) ? section.highlights : [],
    })) : SAMPLE_COPY.textSections.map((section) => ({ ...section, tables: [], highlights: [] }));
    const allPageIds = ['cover', ...sectionOrder, 'chart', 'table', ...textSections.map((section) => section.id)];
    const savedPageOrder = Array.isArray(parsed.pageOrder)
      ? [...new Set(parsed.pageOrder.filter((id): id is string => typeof id === 'string' && allPageIds.includes(id)))]
      : [];
    return {
      ...SAMPLE_COPY,
      ...parsed,
      dataTableColumns,
      dataTableRows: dataTableRows.map((row) => dataTableColumns.map((_, index) => row[index] ?? '')),
      textSections,
      hiddenPages: Array.isArray(parsed.hiddenPages) ? parsed.hiddenPages.filter((page): page is string => typeof page === 'string') : [],
      sectionOrder,
      pageOrder: [...savedPageOrder, ...allPageIds.filter((id) => !savedPageOrder.includes(id))],
    };
  } catch {
    return SAMPLE_COPY;
  }
}

function splitParagraphs(value: string) {
  return value.split(/\n\s*\n/).filter((paragraph) => paragraph.trim());
}

function escapeXml(value: string) {
  return value.replace(/[<>&'"]/g, (character) => ({
    '<': '&lt;',
    '>': '&gt;',
    '&': '&amp;',
    "'": '&apos;',
    '"': '&quot;',
  })[character] ?? character);
}

function coverMetadata(copy: PublicationCopy) {
  const metadata = [
    { label: 'Prepared by', value: copy.author },
    { label: 'Position', value: copy.position },
    { label: 'Organization', value: copy.organization },
    { label: 'Reporting period', value: copy.reportingPeriod },
    { label: 'Published', value: formatPublicationDate(copy.publicationDate) },
  ];
  return metadata.filter((item) => item.value.trim());
}

function formatPublicationDate(value: string) {
  if (!value) return '';
  const [year, month, day] = value.split('-').map(Number);
  if (!year || !month || !day) return value;
  return new Date(year, month - 1, day).toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
}

async function makePngFallback(svgBytes: Uint8Array<ArrayBuffer>, width: number, height: number) {
  const objectUrl = URL.createObjectURL(new Blob([svgBytes], { type: 'image/svg+xml' }));
  try {
    const image = new Image();
    image.src = objectUrl;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Could not prepare cover artwork for Word.');
    context.drawImage(image, 0, 0, width, height);
    const png = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('Could not prepare cover artwork for Word.'))), 'image/png');
    });
    return new Uint8Array(await png.arrayBuffer());
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function BrandMark() {
  return <span className="studio-mark" aria-hidden="true">IM</span>;
}

function PublicationStudio() {
  const [copy, setCopy] = useState<PublicationCopy>(readSavedCopy);
  const [saveState, setSaveState] = useState<'saved' | 'saving' | 'error'>('saved');
  const [mobilePreview, setMobilePreview] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [draggedSection, setDraggedSection] = useState<'context' | 'implications' | null>(null);

  useEffect(() => {
    setSaveState('saving');
    const timer = window.setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(copy));
        setSaveState('saved');
      } catch {
        setSaveState('error');
      }
    }, 450);
    return () => window.clearTimeout(timer);
  }, [copy]);

  const titleForRunningHead = useMemo(
    () => copy.title.replace(/\n/g, ' ').trim() || 'Untitled publication',
    [copy.title],
  );
  const pageIds = ['cover', ...copy.sectionOrder, 'chart', 'table', ...copy.textSections.map((section) => section.id)];
  const fullPageOrder = [...copy.pageOrder.filter((id) => pageIds.includes(id)), ...pageIds.filter((id) => !copy.pageOrder.includes(id))];
  const visiblePageIds = fullPageOrder.filter((id) => !copy.hiddenPages.includes(id));
  const pageCount = visiblePageIds.length;
  const pageNumber = (id: string) => Math.max(visiblePageIds.indexOf(id) + 1, 1);
  const coverDetails = coverMetadata(copy);
  const comparisonRows = useMemo(() => {
    const excerpts = splitParagraphs(copy.contextBody);
    const compact = (value: string, fallback: string) => {
      const normalized = value?.replace(/\s+/g, ' ').trim() || fallback;
      return normalized.length > 88 ? `${normalized.slice(0, 85).trimEnd()}…` : normalized;
    };
    return [
      { view: 'Earlier view', observation: compact(excerpts[0] ?? '', 'Lorem ipsum dolor sit amet'), context: compact(copy.contextLead, 'Consectetur adipiscing') },
      { view: 'Current view', observation: compact(excerpts[1] ?? excerpts[0] ?? '', 'Integer vitae justo sed urna'), context: compact(copy.subtitle, 'Posuere volutpat') },
      { view: 'Movement', observation: compact(excerpts[2] ?? excerpts[1] ?? '', 'Vivamus non augue at erat'), context: compact(copy.marginHeading, 'Commodo feugiat') },
    ];
  }, [copy.contextBody, copy.contextLead, copy.marginHeading, copy.subtitle]);
  const chartValues = copy.chartData.map((point) => {
    const parsed = Number.parseFloat(point.value);
    return Number.isFinite(parsed) ? parsed : 0;
  });
  const chartScale = Math.max(1, ...chartValues.map(Math.abs));

  function updateField(field: keyof PublicationCopy, value: string) {
    setCopy((current) => ({ ...current, [field]: value }));
  }

  function updateReviewPoint(index: number, value: string) {
    setCopy((current) => ({
      ...current,
      reviewPoints: current.reviewPoints.map((point, pointIndex) =>
        pointIndex === index ? value : point,
      ),
    }));
  }

  function updateChartPoint(index: number, key: 'label' | 'value', value: string) {
    setCopy((current) => ({
      ...current,
      chartData: current.chartData.map((point, pointIndex) =>
        pointIndex === index ? { ...point, [key]: value } : point,
      ),
    }));
  }

  function updateDataTableCell(
    index: number,
    columnIndex: number,
    value: string,
  ) {
    setCopy((current) => ({
      ...current,
      dataTableRows: current.dataTableRows.map((row, rowIndex) =>
        rowIndex === index ? row.map((cell, cellIndex) => cellIndex === columnIndex ? value : cell) : row,
      ),
    }));
  }

  function updateDataTableColumn(index: number, value: string) {
    setCopy((current) => ({
      ...current,
      dataTableColumns: current.dataTableColumns.map((column, columnIndex) => columnIndex === index ? value : column),
    }));
  }

  function addDataTableColumn() {
    setCopy((current) => current.dataTableColumns.length >= 8 ? current : ({
      ...current,
      dataTableColumns: [...current.dataTableColumns, `Column ${current.dataTableColumns.length + 1}`],
      dataTableRows: current.dataTableRows.map((row) => [...row, '']),
    }));
  }

  function removeDataTableColumn(index: number) {
    setCopy((current) => current.dataTableColumns.length <= 1 ? current : ({
      ...current,
      dataTableColumns: current.dataTableColumns.filter((_, columnIndex) => columnIndex !== index),
      dataTableRows: current.dataTableRows.map((row) => row.filter((_, columnIndex) => columnIndex !== index)),
    }));
  }

  function updateTextSection(index: number, field: 'kicker' | 'title' | 'lead' | 'body', value: string) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, sectionIndex) =>
        sectionIndex === index ? { ...section, [field]: value } : section,
      ),
    }));
  }

  function addTextSection() {
    setCopy((current) => {
      const section = {
        id: crypto.randomUUID(),
        kicker: 'New section',
        title: 'Section heading',
        lead: '',
        body: 'Write your text here.',
        tables: [],
        highlights: [],
      };
      return {
        ...current,
        textSections: [...current.textSections, section],
        pageOrder: [...current.pageOrder, section.id],
      };
    });
  }

  function removeTextSection(index: number) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.filter((_, sectionIndex) => sectionIndex !== index),
      hiddenPages: current.hiddenPages.filter((page) => page !== current.textSections[index]?.id),
      pageOrder: current.pageOrder.filter((page) => page !== current.textSections[index]?.id),
    }));
  }

  function togglePage(id: string) {
    setCopy((current) => {
      if (!current.hiddenPages.includes(id) && visiblePageIds.length <= 1) return current;
      return {
        ...current,
        hiddenPages: current.hiddenPages.includes(id)
          ? current.hiddenPages.filter((page) => page !== id)
          : [...current.hiddenPages, id],
      };
    });
  }

  function pageVisibilityButton(id: string, label: string) {
    const hidden = copy.hiddenPages.includes(id);
    return <button type="button" className="row-remove-button" onClick={() => togglePage(id)} aria-label={`${hidden ? 'Show' : 'Hide'} ${label} page`}>{hidden ? 'Show page' : 'Hide page'}</button>;
  }

  function addTextTable(sectionIndex: number) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        tables: [...section.tables, {
          id: crypto.randomUUID(),
          title: 'New table',
          columns: ['Column 1', 'Column 2'],
          rows: [['', '']],
        }],
      })),
    }));
  }

  function updateTextTable(sectionIndex: number, tableIndex: number, transform: (table: TableElement) => TableElement) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        tables: section.tables.map((table, index) => index === tableIndex ? transform(table) : table),
      })),
    }));
  }

  function removeTextTable(sectionIndex: number, tableIndex: number) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        tables: section.tables.filter((_, index) => index !== tableIndex),
      })),
    }));
  }

  function addTextHighlight(sectionIndex: number) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        highlights: [...section.highlights, { id: crypto.randomUUID(), title: 'Key point', text: 'Add highlighted text.' }],
      })),
    }));
  }

  function updateTextHighlight(sectionIndex: number, highlightIndex: number, field: 'title' | 'text', value: string) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        highlights: section.highlights.map((item, itemIndex) => itemIndex === highlightIndex ? { ...item, [field]: value } : item),
      })),
    }));
  }

  function removeTextHighlight(sectionIndex: number, highlightIndex: number) {
    setCopy((current) => ({
      ...current,
      textSections: current.textSections.map((section, index) => index !== sectionIndex ? section : ({
        ...section,
        highlights: section.highlights.filter((_, itemIndex) => itemIndex !== highlightIndex),
      })),
    }));
  }

  function addChartPoint() {
    setCopy((current) => current.chartData.length >= 8 ? current : ({
      ...current,
      chartData: [...current.chartData, { label: `Indicator ${current.chartData.length + 1}`, value: '0' }],
    }));
  }

  function removeChartPoint(index: number) {
    setCopy((current) => ({
      ...current,
      chartData: current.chartData.filter((_, pointIndex) => pointIndex !== index),
    }));
  }

  function addDataTableRow() {
    setCopy((current) => current.dataTableRows.length >= 12 ? current : ({
      ...current,
      dataTableRows: [...current.dataTableRows, current.dataTableColumns.map((_, index) => index === 0 ? `Metric ${current.dataTableRows.length + 1}` : '')],
    }));
  }

  function removeDataTableRow(index: number) {
    setCopy((current) => ({
      ...current,
      dataTableRows: current.dataTableRows.filter((_, rowIndex) => rowIndex !== index),
    }));
  }

  function moveSection(source: 'context' | 'implications', target: 'context' | 'implications') {
    if (source === target) return;
    setCopy((current) => {
      const ids = ['cover', ...current.sectionOrder, 'chart', 'table', ...current.textSections.map((section) => section.id)];
      const order = [...current.pageOrder.filter((id) => ids.includes(id)), ...ids.filter((id) => !current.pageOrder.includes(id))];
      const sourceIndex = order.indexOf(source);
      const targetIndex = order.indexOf(target);
      [order[sourceIndex], order[targetIndex]] = [order[targetIndex], order[sourceIndex]];
      const sectionOrder = order.filter((id): id is 'context' | 'implications' => id === 'context' || id === 'implications');
      return { ...current, pageOrder: order, sectionOrder };
    });
  }

  function movePage(id: string, direction: -1 | 1) {
    setCopy((current) => {
      const ids = ['cover', ...current.sectionOrder, 'chart', 'table', ...current.textSections.map((section) => section.id)];
      const order = [...current.pageOrder.filter((page) => ids.includes(page)), ...ids.filter((page) => !current.pageOrder.includes(page))];
      const index = order.indexOf(id);
      const targetIndex = index + direction;
      if (index < 0 || targetIndex < 0 || targetIndex >= order.length) return current;
      [order[index], order[targetIndex]] = [order[targetIndex], order[index]];
      const sectionOrder = order.filter((page): page is 'context' | 'implications' => page === 'context' || page === 'implications');
      return { ...current, pageOrder: order, sectionOrder };
    });
  }

  function pageMoveButtons(id: string) {
    const index = fullPageOrder.indexOf(id);
    return (
      <span className="section-order-controls">
        <button type="button" className="reorder-button" aria-label={`Move ${id} page up`} disabled={index <= 0} onClick={() => movePage(id, -1)}><ArrowUp size={12} /></button>
        <button type="button" className="reorder-button" aria-label={`Move ${id} page down`} disabled={index < 0 || index >= fullPageOrder.length - 1} onClick={() => movePage(id, 1)}><ArrowDown size={12} /></button>
      </span>
    );
  }

  function printPublication() {
    // Persist the current controlled copy before opening the browser print dialog.
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(copy));
      setSaveState('saved');
    } catch {
      setSaveState('error');
    }
    void (document.fonts?.ready ?? Promise.resolve()).then(() => window.print());
  }

  function sectionMoveButtons(section: 'context' | 'implications') {
    const index = copy.sectionOrder.indexOf(section);
    return (
      <span className="section-order-controls">
        <button
          type="button"
          className="reorder-button"
          aria-label={`Move ${section} section up`}
          disabled={index === 0}
          onClick={() => moveSection('context', 'implications')}
        ><ArrowUp size={12} /></button>
        <button
          type="button"
          className="reorder-button"
          aria-label={`Move ${section} section down`}
          disabled={index === 1}
          onClick={() => moveSection('context', 'implications')}
        ><ArrowDown size={12} /></button>
      </span>
    );
  }

  function restoreSample() {
    setCopy({
      ...SAMPLE_COPY,
      reviewPoints: [...SAMPLE_COPY.reviewPoints],
      chartData: SAMPLE_COPY.chartData.map((point) => ({ ...point })),
      dataTableRows: SAMPLE_COPY.dataTableRows.map((row) => [...row]),
      dataTableColumns: [...SAMPLE_COPY.dataTableColumns],
      textSections: SAMPLE_COPY.textSections.map((section) => ({ ...section, tables: [] })),
    });
  }

  async function exportDocx() {
    setExporting(true);
    try {
      const coverArt = new TextEncoder().encode(
        `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="300" viewBox="0 0 900 300">
          <defs>
            <linearGradient id="paper" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#d9e7e5"/><stop offset="1" stop-color="#c7d9dc"/></linearGradient>
          </defs>
          <rect width="900" height="300" fill="url(#paper)"/>
          <ellipse cx="590" cy="185" rx="280" ry="88" fill="none" stroke="#24686e" stroke-opacity=".42" transform="rotate(-24 590 185)"/>
          <circle cx="650" cy="135" r="150" fill="none" stroke="#24686e" stroke-opacity=".42"/>
          <circle cx="650" cy="135" r="170" fill="none" stroke="#f5f7f5" stroke-opacity=".16" stroke-width="20"/>
          <circle cx="650" cy="135" r="192" fill="none" stroke="#f5f7f5" stroke-opacity=".1" stroke-width="22"/>
          <line x1="70" y1="225" x2="470" y2="52" stroke="#24686e" stroke-opacity=".45"/>
        </svg>`,
      );
      const coverArtPng = await makePngFallback(coverArt, 900, 300);
      const maximumChartValue = Math.max(1, ...chartValues.map(Math.abs));
      const chartBars = copy.chartData.map((point, index) => {
        const value = chartValues[index] ?? 0;
        const slot = 800 / Math.max(copy.chartData.length, 1);
        const x = 50 + slot * index + slot * 0.23;
        const width = slot * 0.54;
        const height = Math.abs(value) / maximumChartValue * 115;
        const y = value >= 0 ? 190 - height : 190;
        return `<rect x="${x}" y="${y}" width="${width}" height="${Math.max(height, 1)}" rx="4" fill="${value < 0 ? '#ba756f' : '#287b7e'}"/>
          <text x="${x + width / 2}" y="${value >= 0 ? y - 10 : y + height + 23}" text-anchor="middle" font-family="Arial" font-size="18" fill="#193644">${escapeXml(point.value)}${escapeXml(copy.chartUnit)}</text>
          <text x="${x + width / 2}" y="285" text-anchor="middle" font-family="Arial" font-size="15" fill="#526a75">${escapeXml(point.label)}</text>`;
      }).join('');
      const chartSvg = new TextEncoder().encode(
        `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="330" viewBox="0 0 900 330">
          <rect width="900" height="330" fill="#f5f7f5"/>
          <g stroke="#dce4e3" stroke-width="1"><line x1="45" x2="860" y1="70" y2="70"/><line x1="45" x2="860" y1="130" y2="130"/><line x1="45" x2="860" y1="190" y2="190"/><line x1="45" x2="860" y1="250" y2="250"/></g>
          <line x1="45" x2="860" y1="190" y2="190" stroke="#84979d" stroke-width="2"/>
          ${chartBars}
        </svg>`,
      );
      const chartPng = await makePngFallback(chartSvg, 900, 330);
      const paragraphs = (text: string) =>
        splitParagraphs(text).map(
          (paragraph) =>
            styledParagraph(paragraph, { size: 20, color: '435D68' }, { spacing: { after: 160, line: 285 } }),
        );
      const makeCell = (text: string, isHeader = false, width: number) =>
        new TableCell({
          width: { size: width, type: WidthType.DXA },
          shading: isHeader ? { fill: 'E3EEEC' } : undefined,
          children: [
            styledParagraph(
              text,
              { size: isHeader ? 15 : 17, color: isHeader ? '1B5E66' : '526975', bold: isHeader },
              { spacing: { before: 90, after: 90 } },
            ),
          ],
        });
      const comparisonTable = new Table({
        width: { size: 9900, type: WidthType.DXA },
        columnWidths: [2000, 5000, 2900],
        rows: [
          new TableRow({ children: [makeCell('View', true, 2000), makeCell('Observation', true, 5000), makeCell('Context', true, 2900)] }),
          ...comparisonRows.map((row) => new TableRow({
            children: [makeCell(row.view, false, 2000), makeCell(row.observation, false, 5000), makeCell(row.context, false, 2900)],
          })),
        ],
      });
      const makeEditableTable = (columns: string[], rows: string[][]) => {
        const safeColumns = columns.length ? columns : ['Column 1'];
        const width = Math.floor(9900 / safeColumns.length);
        return new Table({
          width: { size: 9900, type: WidthType.DXA },
          columnWidths: safeColumns.map(() => width),
          rows: [
            new TableRow({ children: safeColumns.map((label) => makeCell(label, true, width)) }),
            ...rows.map((row) => new TableRow({
              children: safeColumns.map((_, index) => makeCell(row[index] ?? '', false, width)),
            })),
          ],
        });
      };
      const dataTable = makeEditableTable(copy.dataTableColumns, copy.dataTableRows);
      const contextDocumentContent = [
        styledParagraph(`${String(copy.sectionOrder.indexOf('context') + 1).padStart(2, '0')}  /  ${copy.contextLabel.toUpperCase()}`, { bold: true, color: '287B7E', size: 17 }, { spacing: { after: 420 } }),
        new Paragraph({ text: copy.contextHeading, heading: HeadingLevel.HEADING_1, spacing: { after: 180 } }),
        styledParagraph(copy.contextLead, { font: 'Georgia', size: 24, color: '476171' }, { spacing: { after: 260 } }),
        ...paragraphs(copy.contextBody),
        styledParagraph(`“${copy.marginNote}”`, { font: 'Georgia', italics: true, size: 27, color: '1B5E66' }, { spacing: { before: 160, after: 80 } }),
        styledParagraph(copy.marginHeading.toUpperCase(), { bold: true, color: '607783', size: 15 }, { spacing: { after: 260 } }),
        styledParagraph(copy.comparisonTitle.toUpperCase(), { bold: true, color: '287B7E', size: 16 }, { spacing: { before: 80, after: 100 } }),
        comparisonTable,
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber('context')).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT, spacing: { before: 260 } }),
      ];
      const implicationsDocumentContent = [
        styledParagraph(`${String(copy.sectionOrder.indexOf('implications') + 1).padStart(2, '0')}  /  ${copy.implicationsLabel.toUpperCase()}`, { bold: true, color: '287B7E', size: 17 }, { spacing: { after: 420 } }),
        new Paragraph({ text: copy.implicationsHeading, heading: HeadingLevel.HEADING_1, spacing: { after: 180 } }),
        styledParagraph(copy.implicationsLead, { font: 'Georgia', size: 24, color: '476171' }, { spacing: { after: 260 } }),
        ...paragraphs(copy.implicationsBody),
        styledParagraph(copy.calloutTitle.toUpperCase(), { bold: true, color: '1B5E66', size: 15 }, { spacing: { before: 180, after: 70 } }),
        styledParagraph(copy.callout, { color: '39545E', size: 19 }, { shading: { fill: 'E3EEEC' }, spacing: { after: 200 } }),
        styledParagraph(copy.reviewTitle.toUpperCase(), { bold: true, color: '287B7E', size: 16 }, { spacing: { before: 180, after: 100 } }),
        ...copy.reviewPoints.map((point, index) =>
          new Paragraph({
            children: [
              new TextRun({ text: `0${index + 1}   `, font: 'Georgia', size: 22, color: '287B7E' }),
              new TextRun({ text: point, size: 18, color: '526A77' }),
            ],
            spacing: { after: 130, line: 280 },
          }),
        ),
        styledParagraph(`${copy.marginHeading.toUpperCase()}  ·  ${copy.marginNote}`, { color: '526A75', size: 16 }, { shading: { fill: 'E5EEEF' }, spacing: { before: 140, after: 180 } }),
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber('implications')).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT, spacing: { before: 180 } }),
      ];
      const chartDocumentContent = [
        styledParagraph(`03  /  ${copy.chartLabel.toUpperCase()}`, { bold: true, color: '287B7E', size: 17 }, { spacing: { after: 260 } }),
        new Paragraph({ text: copy.chartTitle, heading: HeadingLevel.HEADING_1, spacing: { after: 180 } }),
        styledParagraph(copy.chartSummary, { font: 'Georgia', size: 22, color: '476171' }, { spacing: { after: 180 } }),
        new Paragraph({
          children: [new ImageRun({
            data: chartSvg,
            type: 'svg',
            fallback: { data: chartPng, type: 'png' },
            transformation: { width: 610, height: 224 },
            altText: { title: copy.chartTitle, description: copy.chartSummary, name: 'Editable sample bar chart' },
          })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 180 },
        }),
        styledParagraph(`Illustrative sample data · Units: ${copy.chartUnit}`, { size: 16, color: '607783' }, { spacing: { after: 220 } }),
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber('chart')).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT }),
      ];
      const tableDocumentContent = [
        styledParagraph(`04  /  ${copy.dataTableLabel.toUpperCase()}`, { bold: true, color: '287B7E', size: 17 }, { spacing: { after: 260 } }),
        new Paragraph({ text: copy.dataTableTitle, heading: HeadingLevel.HEADING_1, spacing: { after: 180 } }),
        styledParagraph(copy.dataTableSummary, { font: 'Georgia', size: 22, color: '476171' }, { spacing: { after: 220 } }),
        dataTable,
        styledParagraph('Sample figures are illustrative. Verify source values and calculations before publication.', { size: 16, color: '607783' }, { spacing: { before: 220, after: 220 } }),
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber('table')).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT }),
      ];
      const extraTextPageContents = copy.textSections.map((section, index) => [
        styledParagraph(`${section.kicker.toUpperCase()}`, { bold: true, color: '287B7E', size: 17 }, { spacing: { after: 420 } }),
        new Paragraph({ text: section.title, heading: HeadingLevel.HEADING_1, spacing: { after: 180 } }),
        ...(section.lead.trim() ? [styledParagraph(section.lead, { font: 'Georgia', size: 24, color: '476171' }, { spacing: { after: 260 } })] : []),
        ...paragraphs(section.body),
        ...section.highlights.flatMap((highlight) => [
          styledParagraph(highlight.title.toUpperCase(), { bold: true, size: 15, color: '1B5E66' }, { shading: { fill: 'E3EEEC' }, spacing: { before: 180, after: 60 } }),
          styledParagraph(highlight.text, { size: 19, color: '39545E' }, { shading: { fill: 'E3EEEC' }, spacing: { after: 180 } }),
        ]),
        ...section.tables.flatMap((table) => [
          new Paragraph({ text: table.title, heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 120 } }),
          makeEditableTable(table.columns, table.rows),
        ]),
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber(section.id)).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT, spacing: { before: 260 } }),
      ]);
      const coverDocumentContent = [
        styledParagraph('IFRS METRICS  /  REPORTING NOTE', { bold: true, size: 16, color: '1B5E66', allCaps: true }, { spacing: { after: 480 } }),
        styledParagraph(copy.coverKicker.toUpperCase(), { bold: true, color: '287B7E', size: 18 }, { spacing: { after: 260 } }),
        new Paragraph({ children: [new TextRun({ text: copy.title, font: 'Georgia', size: 72, color: '193644' })], spacing: { after: 180 } }),
        styledParagraph(copy.subtitle, { color: '526A75', size: 24 }, { spacing: { after: 220 } }),
        ...coverDetails.map(({ label, value }) => styledParagraph(`${label}  ${value}`, { size: 17, color: '607783' }, { spacing: { after: 70 } })),
        ...(coverDetails.length ? [new Paragraph({ spacing: { after: 130 } })] : []),
        styledParagraph(copy.introLabel.toUpperCase(), { bold: true, color: '1B5E66', size: 17 }, { spacing: { after: 90 } }),
        styledParagraph(copy.intro, { font: 'Aptos', size: 20, color: '607783' }, { spacing: { after: 260 } }),
        new Paragraph({
          children: [new ImageRun({
            data: coverArt, type: 'svg', fallback: { data: coverArtPng, type: 'png' },
            transformation: { width: 620, height: 205 },
            altText: { title: 'Insight Magazine abstract cover art', description: 'Abstract teal line art on a pale blue-green field', name: 'Insight Magazine cover art' },
          })],
          alignment: AlignmentType.CENTER, spacing: { after: 240 },
        }),
        styledParagraph(`${titleForRunningHead}  ·  ${String(pageNumber('cover')).padStart(2, '0')} / ${pageCount}`, { size: 15, color: '607783', bold: true }, { alignment: AlignmentType.RIGHT, spacing: { before: 260 } }),
      ];
      const docPageContent: Record<string, (Paragraph | Table)[]> = {
        cover: coverDocumentContent,
        context: contextDocumentContent,
        implications: implicationsDocumentContent,
        chart: chartDocumentContent,
        table: tableDocumentContent,
      };
      copy.textSections.forEach((section, index) => {
        docPageContent[section.id] = extraTextPageContents[index];
      });
      const documentChildren = visiblePageIds.flatMap((id, index) => [
        ...(index === 0 ? [] : [new Paragraph({ children: [new PageBreak()] })]),
        ...(docPageContent[id] ?? []),
      ]);
      const document = new Document({
        styles: {
          default: {
            document: {
              run: { font: 'Aptos', size: 22, color: '17364A' },
              paragraph: { spacing: { line: 300 } },
            },
            heading1: { run: { font: 'Georgia', size: 48, color: '17364A', bold: false } },
            heading2: { run: { font: 'Georgia', size: 36, color: '17364A', bold: false } },
          },
        },
        sections: [
          {
            properties: {
              page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 900, right: 1000, bottom: 850, left: 1000 },
              },
            },
            children: documentChildren,
          },
        ],
      });
      const blob = await Packer.toBlob(document);
      const url = URL.createObjectURL(blob);
      const link = window.document.createElement('a');
      link.href = url;
      link.download = `${titleForRunningHead.toLowerCase().replace(/[^a-z0-9]+/g, '-') || 'publication'}.docx`;
      link.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'The Word document could not be created.');
    } finally {
      setExporting(false);
    }
  }

  const contextEditor = (
    <section
      className={`editor-group reorderable-group ${draggedSection === 'context' ? 'is-dragging' : ''}`}
      key="context-editor"
      draggable
      onDragStart={(event) => {
        setDraggedSection('context');
        event.dataTransfer.setData('text/plain', 'context');
        event.dataTransfer.effectAllowed = 'move';
      }}
      onDragEnd={() => setDraggedSection(null)}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const source = event.dataTransfer.getData('text/plain');
        if (source === 'context' || source === 'implications') moveSection(source, 'context');
        setDraggedSection(null);
      }}
    >
      <div className="group-label"><GripVertical size={14} className="drag-grip" aria-hidden="true" /><span className="group-index">{String(copy.sectionOrder.indexOf('context') + 2).padStart(2, '0')}</span> Context <span className="drag-hint">Drag to reorder</span>{sectionMoveButtons('context')}{pageMoveButtons('context')}{pageVisibilityButton('context', 'Context')}</div>
      <Field label="Section label" id="context-label" value={copy.contextLabel} onChange={(value) => updateField('contextLabel', value)} placeholder="e.g. Context" />
      <Field label="Section heading" id="context-heading" value={copy.contextHeading} onChange={(value) => updateField('contextHeading', value)} />
      <Field label="Lead paragraph" id="context-lead" value={copy.contextLead} onChange={(value) => updateField('contextLead', value)} multiline />
      <Field label="Body copy" id="context-body" value={copy.contextBody} onChange={(value) => updateField('contextBody', value)} multiline tall hint="Separate paragraphs with a blank line." />
      <Field label="Comparison table title" id="comparison-title" value={copy.comparisonTitle} onChange={(value) => updateField('comparisonTitle', value)} />
      <Field label="Margin note heading" id="margin-heading" value={copy.marginHeading} onChange={(value) => updateField('marginHeading', value)} />
      <Field label="Margin note" id="margin-note" value={copy.marginNote} onChange={(value) => updateField('marginNote', value)} multiline />
    </section>
  );

  const implicationsEditor = (
    <section
      className={`editor-group reorderable-group ${draggedSection === 'implications' ? 'is-dragging' : ''}`}
      key="implications-editor"
      draggable
      onDragStart={(event) => {
        setDraggedSection('implications');
        event.dataTransfer.setData('text/plain', 'implications');
        event.dataTransfer.effectAllowed = 'move';
      }}
      onDragEnd={() => setDraggedSection(null)}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        const source = event.dataTransfer.getData('text/plain');
        if (source === 'context' || source === 'implications') moveSection(source, 'implications');
        setDraggedSection(null);
      }}
    >
      <div className="group-label"><GripVertical size={14} className="drag-grip" aria-hidden="true" /><span className="group-index">{String(copy.sectionOrder.indexOf('implications') + 2).padStart(2, '0')}</span> Implications &amp; review <span className="drag-hint">Drag to reorder</span>{sectionMoveButtons('implications')}{pageMoveButtons('implications')}{pageVisibilityButton('implications', 'Application')}</div>
      <Field label="Section label" id="implications-label" value={copy.implicationsLabel} onChange={(value) => updateField('implicationsLabel', value)} />
      <Field label="Section heading" id="implications-heading" value={copy.implicationsHeading} onChange={(value) => updateField('implicationsHeading', value)} />
      <Field label="Callout title" id="callout-title" value={copy.calloutTitle} onChange={(value) => updateField('calloutTitle', value)} />
      <Field label="Lead paragraph" id="implications-lead" value={copy.implicationsLead} onChange={(value) => updateField('implicationsLead', value)} multiline />
      <Field label="Body copy" id="implications-body" value={copy.implicationsBody} onChange={(value) => updateField('implicationsBody', value)} multiline />
      {copy.reviewPoints.map((point, index) => (
        <Field key={index} label={`Review point 0${index + 1}`} id={`review-${index + 1}`} value={point} onChange={(value) => updateReviewPoint(index, value)} multiline />
      ))}
      <Field label="Review list title" id="review-title" value={copy.reviewTitle} onChange={(value) => updateField('reviewTitle', value)} />
      <Field label="For discussion callout" id="callout" value={copy.callout} onChange={(value) => updateField('callout', value)} multiline />
    </section>
  );

  const chartEditor = (
    <section className="editor-group">
      <div className="group-label"><span className="group-index">04</span> Chart sample page {pageMoveButtons('chart')}{pageVisibilityButton('chart', 'Chart')}</div>
      <Field label="Section label" id="chart-label" value={copy.chartLabel} onChange={(value) => updateField('chartLabel', value)} />
      <Field label="Chart title" id="chart-title" value={copy.chartTitle} onChange={(value) => updateField('chartTitle', value)} />
      <Field label="Chart explanation" id="chart-summary" value={copy.chartSummary} onChange={(value) => updateField('chartSummary', value)} multiline />
      <Field label="Value unit" id="chart-unit" value={copy.chartUnit} onChange={(value) => updateField('chartUnit', value)} placeholder="e.g. %, USD millions" />
      <div className="data-editor-list">
        {copy.chartData.map((point, index) => (
          <div className="data-editor-row" key={`chart-${index}`}>
            <Field label={`Label ${index + 1}`} id={`chart-label-${index}`} value={point.label} onChange={(value) => updateChartPoint(index, 'label', value)} />
            <Field label="Value" id={`chart-value-${index}`} value={point.value} onChange={(value) => updateChartPoint(index, 'value', value)} />
            <button type="button" className="row-remove-button" onClick={() => removeChartPoint(index)} disabled={copy.chartData.length <= 1} aria-label={`Remove chart item ${index + 1}`}>Remove</button>
          </div>
        ))}
      </div>
      <button type="button" className="studio-button add-row-button" onClick={addChartPoint} disabled={copy.chartData.length >= 8}>Add chart item</button>
    </section>
  );

  const tableEditor = (
    <section className="editor-group">
      <div className="group-label"><span className="group-index">05</span> Table sample page {pageMoveButtons('table')}{pageVisibilityButton('table', 'Table')}</div>
      <Field label="Section label" id="table-label" value={copy.dataTableLabel} onChange={(value) => updateField('dataTableLabel', value)} />
      <Field label="Table title" id="table-title" value={copy.dataTableTitle} onChange={(value) => updateField('dataTableTitle', value)} />
      <Field label="Table introduction" id="table-summary" value={copy.dataTableSummary} onChange={(value) => updateField('dataTableSummary', value)} multiline />
      <p className="data-table-help">Choose your own columns and rows.</p>
      <div className="data-editor-list">
        {copy.dataTableColumns.map((column, columnIndex) => (
          <div className="data-editor-row" key={`table-column-${columnIndex}`}>
            <Field label={`Column ${columnIndex + 1} heading`} id={`table-column-${columnIndex}`} value={column} onChange={(value) => updateDataTableColumn(columnIndex, value)} />
            <button type="button" className="row-remove-button" onClick={() => removeDataTableColumn(columnIndex)} disabled={copy.dataTableColumns.length <= 1}>Remove column</button>
          </div>
        ))}
      </div>
      <button type="button" className="studio-button add-row-button" onClick={addDataTableColumn} disabled={copy.dataTableColumns.length >= 8}>Add column</button>
      <div className="data-editor-list">
        {copy.dataTableRows.map((row, index) => (
          <div className="data-editor-row table-editor-row" key={`table-${index}`}>
            {copy.dataTableColumns.map((column, columnIndex) => (
              <Field key={columnIndex} label={`${column} · row ${index + 1}`} id={`table-cell-${index}-${columnIndex}`} value={row[columnIndex] ?? ''} onChange={(value) => updateDataTableCell(index, columnIndex, value)} />
            ))}
            <button type="button" className="row-remove-button" onClick={() => removeDataTableRow(index)} disabled={copy.dataTableRows.length <= 1} aria-label={`Remove table row ${index + 1}`}>Remove row</button>
          </div>
        ))}
      </div>
      <button type="button" className="studio-button add-row-button" onClick={addDataTableRow} disabled={copy.dataTableRows.length >= 12}>Add table row</button>
    </section>
  );

  const textSectionEditors = copy.textSections.map((section, index) => (
    <section className="editor-group" key={section.id}>
      <div className="group-label">
        <span className="group-index">{String(index + 6).padStart(2, '0')}</span>
        Text page {index + 1}
        {pageMoveButtons(section.id)}
        {pageVisibilityButton(section.id, `Text page ${index + 1}`)}
        <button type="button" className="row-remove-button" onClick={() => removeTextSection(index)} aria-label={`Remove text page ${index + 1}`}>Remove page</button>
      </div>
      <Field label="Small section label" id={`text-kicker-${section.id}`} value={section.kicker} onChange={(value) => updateTextSection(index, 'kicker', value)} />
      <Field label="Page heading" id={`text-title-${section.id}`} value={section.title} onChange={(value) => updateTextSection(index, 'title', value)} />
      <Field label="Introduction (optional)" id={`text-lead-${section.id}`} value={section.lead} onChange={(value) => updateTextSection(index, 'lead', value)} multiline />
      <Field label="Body text" id={`text-body-${section.id}`} value={section.body} onChange={(value) => updateTextSection(index, 'body', value)} multiline tall hint="Separate paragraphs with a blank line." />
      {section.highlights.map((highlight, highlightIndex) => (
        <div className="embedded-table-editor" key={highlight.id}>
          <div className="field-section-label">Green highlight {highlightIndex + 1}</div>
          <Field label="Highlight title" id={`highlight-title-${highlight.id}`} value={highlight.title} onChange={(value) => updateTextHighlight(index, highlightIndex, 'title', value)} />
          <Field label="Highlight text" id={`highlight-text-${highlight.id}`} value={highlight.text} onChange={(value) => updateTextHighlight(index, highlightIndex, 'text', value)} multiline />
          <button type="button" className="row-remove-button" onClick={() => removeTextHighlight(index, highlightIndex)}>Remove highlight</button>
        </div>
      ))}
      <button type="button" className="studio-button add-row-button" onClick={() => addTextHighlight(index)}>Add green highlight</button>
      {section.tables.map((table, tableIndex) => (
        <div className="embedded-table-editor" key={table.id}>
          <div className="field-section-label">Table element {tableIndex + 1}</div>
          <Field label="Element title" id={`text-table-title-${table.id}`} value={table.title} onChange={(value) => updateTextTable(index, tableIndex, (current) => ({ ...current, title: value }))} />
          {table.columns.map((column, columnIndex) => (
            <div className="data-editor-row" key={`custom-column-${columnIndex}`}>
              <Field label={`Column ${columnIndex + 1}`} id={`text-table-column-${table.id}-${columnIndex}`} value={column} onChange={(value) => updateTextTable(index, tableIndex, (current) => ({ ...current, columns: current.columns.map((item, i) => i === columnIndex ? value : item) }))} />
              <button type="button" className="row-remove-button" disabled={table.columns.length <= 1} onClick={() => updateTextTable(index, tableIndex, (current) => ({ ...current, columns: current.columns.filter((_, i) => i !== columnIndex), rows: current.rows.map((row) => row.filter((_, i) => i !== columnIndex)) }))}>Remove column</button>
            </div>
          ))}
          <button type="button" className="studio-button add-row-button" disabled={table.columns.length >= 8} onClick={() => updateTextTable(index, tableIndex, (current) => ({ ...current, columns: [...current.columns, `Column ${current.columns.length + 1}`], rows: current.rows.map((row) => [...row, '']) }))}>Add column</button>
          {table.rows.map((row, rowIndex) => (
            <div className="data-editor-row table-editor-row" key={`custom-row-${rowIndex}`}>
              {table.columns.map((column, columnIndex) => (
                <Field key={columnIndex} label={`${column} · row ${rowIndex + 1}`} id={`text-table-cell-${table.id}-${rowIndex}-${columnIndex}`} value={row[columnIndex] ?? ''} onChange={(value) => updateTextTable(index, tableIndex, (current) => ({ ...current, rows: current.rows.map((cells, i) => i === rowIndex ? cells.map((cell, j) => j === columnIndex ? value : cell) : cells) }))} />
              ))}
              <button type="button" className="row-remove-button" disabled={table.rows.length <= 1} onClick={() => updateTextTable(index, tableIndex, (current) => ({ ...current, rows: current.rows.filter((_, i) => i !== rowIndex) }))}>Remove row</button>
            </div>
          ))}
          <button type="button" className="studio-button add-row-button" onClick={() => updateTextTable(index, tableIndex, (current) => ({ ...current, rows: [...current.rows, current.columns.map(() => '')] }))}>Add row</button>
          <button type="button" className="row-remove-button" onClick={() => removeTextTable(index, tableIndex)}>Remove table element</button>
        </div>
      ))}
      <button type="button" className="studio-button add-row-button" onClick={() => addTextTable(index)}>Add table element to this page</button>
    </section>
  ));

  const contextPage = (
    <article className="mag-page" data-testid="preview-page-context" key="context-page">
      <MagazineHeader title={titleForRunningHead} />
      <header className="mag-section-head">
        <div><p className="mag-kicker">{String(copy.sectionOrder.indexOf('context') + 1).padStart(2, '0')} / {copy.contextLabel}</p><h2>{copy.contextHeading}</h2></div>
        <p className="mag-deck">{copy.contextLead}</p>
      </header>
      <div className="mag-columns">
        <div className="mag-copy">
          {splitParagraphs(copy.contextBody).map((paragraph, index) => <p key={index}>{paragraph}</p>)}
        </div>
        <blockquote className="mag-quote">
          “{copy.marginNote}”
          <small>{copy.marginHeading} · IFRS Metrics</small>
        </blockquote>
      </div>
      <p className="mag-mini-label">{copy.comparisonTitle}</p>
      <div className="mag-table-wrap">
        <table className="mag-table">
          <thead><tr><th>View</th><th>Observation</th><th>Context</th></tr></thead>
          <tbody>
            {comparisonRows.map((row) => (
              <tr key={row.view}><td>{row.view}</td><td>{row.observation}</td><td>{row.context}</td></tr>
            ))}
          </tbody>
        </table>
      </div>
      <MagazineFooter title={titleForRunningHead} page={`${String(pageNumber('context')).padStart(2, '0')} / ${pageCount}`} />
    </article>
  );

  const implicationsPage = (
    <article className="mag-page mag-page-three" data-testid="preview-page-implications" key="implications-page">
      <MagazineHeader title={titleForRunningHead} />
      <header className="mag-section-head">
        <div><p className="mag-kicker">{String(copy.sectionOrder.indexOf('implications') + 1).padStart(2, '0')} / {copy.implicationsLabel}</p><h2>{copy.implicationsHeading}</h2></div>
        <p className="mag-deck">{copy.implicationsLead}</p>
      </header>
      <div className="mag-page-three-layout">
        <div>
          <div className="mag-copy">
            {splitParagraphs(copy.implicationsBody).map((paragraph, index) => <p key={index}>{paragraph}</p>)}
          </div>
          <div className="mag-callout">
            <div className="mag-callout-label">{copy.calloutTitle}</div>
            <p>{copy.callout}</p>
          </div>
        </div>
        <aside className="mag-side-note">
          <strong>{copy.marginHeading}</strong>
          {copy.marginNote}
        </aside>
      </div>
      <p className="mag-mini-label">{copy.reviewTitle}</p>
      <ol className="mag-review-list">
        {copy.reviewPoints.map((point, index) => (
          <li key={index}><div><strong>Review point 0{index + 1}</strong><span>{point}</span></div></li>
        ))}
      </ol>
      <MagazineFooter title={`${titleForRunningHead} · End of note`} page={`${String(pageNumber('implications')).padStart(2, '0')} / ${pageCount}`} />
    </article>
  );

  const chartPage = (
    <article className="mag-page mag-data-page" data-testid="preview-page-chart">
      <MagazineHeader title={titleForRunningHead} />
      <header className="mag-section-head">
        <div><p className="mag-kicker">03 / {copy.chartLabel}</p><h2>{copy.chartTitle}</h2></div>
        <p className="mag-deck">{copy.chartSummary}</p>
      </header>
      <div className="mag-chart-panel">
        <div className="mag-chart-unit">{copy.chartUnit}</div>
        <svg className="mag-bar-chart" viewBox="0 0 800 300" role="img" aria-label={`${copy.chartTitle}: ${copy.chartData.map((point) => `${point.label} ${point.value} ${copy.chartUnit}`).join(', ')}`}>
          {[70, 120, 170, 220].map((y) => <line key={y} x1="42" x2="780" y1={y} y2={y} className="chart-gridline" />)}
          <line x1="42" x2="780" y1="170" y2="170" className="chart-baseline" />
          {copy.chartData.map((point, index) => {
            const value = chartValues[index] ?? 0;
            const slot = 730 / Math.max(copy.chartData.length, 1);
            const x = 42 + slot * index + slot * 0.24;
            const width = slot * 0.52;
            const height = Math.abs(value) / chartScale * 100;
            const y = value >= 0 ? 170 - height : 170;
            return (
              <g key={`bar-${index}`}>
                <rect x={x} y={y} width={width} height={Math.max(height, 1)} rx="2" className={value < 0 ? 'chart-bar negative' : 'chart-bar'} />
                <text x={x + width / 2} y={value >= 0 ? y - 9 : y + height + 18} textAnchor="middle" className="chart-value-label">{point.value}{copy.chartUnit}</text>
                <text x={x + width / 2} y="255" textAnchor="middle" className="chart-category-label">{point.label}</text>
              </g>
            );
          })}
        </svg>
      </div>
      <p className="mag-chart-note">Illustrative sample data. Replace the labels and values with verified source figures before publication.</p>
      <MagazineFooter title={titleForRunningHead} page={`${String(pageNumber('chart')).padStart(2, '0')} / ${pageCount}`} />
    </article>
  );

  const dataTablePage = (
    <article className="mag-page mag-data-page" data-testid="preview-page-table">
      <MagazineHeader title={titleForRunningHead} />
      <header className="mag-section-head">
        <div><p className="mag-kicker">04 / {copy.dataTableLabel}</p><h2>{copy.dataTableTitle}</h2></div>
        <p className="mag-deck">{copy.dataTableSummary}</p>
      </header>
      <div className="mag-table-wrap mag-editable-table-wrap">
        <table className="mag-table mag-data-table">
          <thead><tr>{copy.dataTableColumns.map((column, index) => <th key={`column-${index}`}>{column}</th>)}</tr></thead>
          <tbody>
            {copy.dataTableRows.map((row, index) => (
              <tr key={`data-row-${index}`}>{copy.dataTableColumns.map((_, columnIndex) => <td key={columnIndex}>{row[columnIndex] ?? ''}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="mag-chart-note">Sample figures are illustrative. Check units, period labels, and calculations against your approved source data.</p>
      <MagazineFooter title={titleForRunningHead} page={`${String(pageNumber('table')).padStart(2, '0')} / ${pageCount}`} />
    </article>
  );
  const extraTextPages = copy.textSections.map((section, index) => copy.hiddenPages.includes(section.id) ? null : (
    <article className="mag-page mag-data-page mag-text-page" data-testid={`preview-text-page-${index + 1}`} key={section.id}>
      <MagazineHeader title={titleForRunningHead} />
      <header className="mag-section-head">
        <div><p className="mag-kicker">{String(index + 5).padStart(2, '0')} / {section.kicker}</p><h2>{section.title}</h2></div>
        {section.lead.trim() && <p className="mag-deck">{section.lead}</p>}
      </header>
      <div className="mag-copy mag-text-body">
        {splitParagraphs(section.body).map((paragraph, paragraphIndex) => <p key={paragraphIndex}>{paragraph}</p>)}
      </div>
      {section.highlights.map((highlight) => (
        <div className="mag-callout embedded-highlight" key={highlight.id}>
          <div className="mag-callout-label">{highlight.title}</div>
          <p>{highlight.text}</p>
        </div>
      ))}
      {section.tables.map((table) => (
        <div className="embedded-table-preview" key={table.id}>
          <p className="mag-mini-label">{table.title}</p>
          <div className="mag-table-wrap">
            <table className="mag-table mag-data-table">
              <thead><tr>{table.columns.map((column, columnIndex) => <th key={columnIndex}>{column}</th>)}</tr></thead>
              <tbody>{table.rows.map((row, rowIndex) => (
                <tr key={rowIndex}>{table.columns.map((_, columnIndex) => <td key={columnIndex}>{row[columnIndex] ?? ''}</td>)}</tr>
              ))}</tbody>
            </table>
          </div>
        </div>
      ))}
      <MagazineFooter title={titleForRunningHead} page={`${String(pageNumber(section.id)).padStart(2, '0')} / ${pageCount}`} />
    </article>
  ));
  const coverPage = !copy.hiddenPages.includes('cover') ? (
    <article className="mag-page mag-cover-page" data-testid="preview-page-1">
      <header className="mag-topline">
        <span className="mag-brandmark"><span className="mag-brand-symbol">i</span>IFRS Metrics</span>
        <span>Reporting note&nbsp; / &nbsp;01—{String(pageCount).padStart(2, '0')}</span>
      </header>
      <div className="mag-cover-grid">
        <div>
          <p className="mag-kicker">{copy.coverKicker}</p>
          <h1 className="mag-title">{copy.title}</h1>
          <p className="mag-subtitle">{copy.subtitle}</p>
          {coverDetails.length > 0 && (
            <div className="mag-byline" data-testid="cover-metadata">
              {coverDetails.map(({ label, value }) => <p key={label}><span>{label}</span>{value}</p>)}
            </div>
          )}
        </div>
        <div className="mag-cover-note"><strong>{copy.introLabel}</strong>{copy.intro}</div>
      </div>
      <div className="mag-cover-art" aria-hidden="true"><span className="mag-art-line" aria-hidden="true" /></div>
      <MagazineFooter title={titleForRunningHead} page={`${String(pageNumber('cover')).padStart(2, '0')} / ${pageCount}`} />
    </article>
  ) : null;
  const previewPageById: Record<string, ReactNode> = {
    cover: coverPage,
    context: copy.hiddenPages.includes('context') ? null : contextPage,
    implications: copy.hiddenPages.includes('implications') ? null : implicationsPage,
    chart: copy.hiddenPages.includes('chart') ? null : chartPage,
    table: copy.hiddenPages.includes('table') ? null : dataTablePage,
  };
  copy.textSections.forEach((section, index) => { previewPageById[section.id] = extraTextPages[index]; });

  return (
    <div className={`studio ${mobilePreview ? 'is-mobile-preview' : ''}`}>
      <header className="studio-header">
        <div className="studio-brand">
          <BrandMark />
          <span className="brand-name">IFRS Metrics</span>
          <span className="brand-divider" />
          <span className="workspace-name">Publication Studio</span>
        </div>
        <div className="header-actions">
          <div className="save-status" data-testid="status-save">
            <span className={`save-dot ${saveState === 'saving' ? 'saving' : ''} ${saveState === 'error' ? 'error' : ''}`} />
            {saveState === 'saving' ? 'Saving changes' : saveState === 'error' ? 'Save failed' : 'All changes saved'}
          </div>
          <button className="studio-button" type="button" onClick={restoreSample} data-testid="button-restore-sample" aria-label="Restore sample content">
            <RotateCcw size={14} /> <span>Restore sample</span>
          </button>
          <button className="studio-button header-mobile-preview" type="button" onClick={() => setMobilePreview(true)} data-testid="button-header-preview">
            <FileText size={14} /> <span>Preview</span>
          </button>
          <button className="studio-button" type="button" onClick={printPublication} data-testid="button-print">
            <Printer size={14} /> <span>Print / PDF</span>
          </button>
          <button className="studio-button primary" type="button" onClick={exportDocx} disabled={exporting} data-testid="button-export-docx">
            {exporting ? <Save size={14} /> : <Download size={14} />}
            <span>{exporting ? 'Preparing…' : 'Export DOCX'}</span>
          </button>
        </div>
      </header>

      <main className={`studio-main ${mobilePreview ? 'mobile-preview' : ''}`}>
        <aside className="editor-panel" aria-label="Publication editor">
          <div className="editor-head">
            <div className="eyebrow">Publication · Working draft</div>
            <h1>Shape the story.</h1>
            <p>Edit the controlled copy below. Your changes are saved in this browser as you work.</p>
          </div>
          <div className="editor-form">
            <section className="editor-group">
          <div className="group-label"><span className="group-index">01</span> Cover {pageMoveButtons('cover')}{pageVisibilityButton('cover', 'Cover')}</div>
              <Field label="Publication title" id="title" value={copy.title} onChange={(value) => updateField('title', value)} multiline />
              <Field label="Subtitle" id="subtitle" value={copy.subtitle} onChange={(value) => updateField('subtitle', value)} />
      <Field label="Cover category line" id="cover-kicker" value={copy.coverKicker} onChange={(value) => updateField('coverKicker', value)} placeholder="e.g. Perspective · Finance teams" />
              <div className="field-section-label">Author and publication details</div>
              <p className="cover-details-help">Filled details appear on the cover and in the Word export.</p>
              <Field label="Author" id="author" value={copy.author} onChange={(value) => updateField('author', value)} placeholder="e.g. Tendai Moyo" />
              <Field label="Position / role" id="position" value={copy.position} onChange={(value) => updateField('position', value)} placeholder="e.g. Chief Financial Officer" />
              <Field label="Organization" id="organization" value={copy.organization} onChange={(value) => updateField('organization', value)} placeholder="e.g. Vainona Finance" />
              <Field label="Reporting period" id="reporting-period" value={copy.reportingPeriod} onChange={(value) => updateField('reportingPeriod', value)} placeholder="e.g. Year ended 31 December 2026" />
              <Field label="Publication date" id="publication-date" value={copy.publicationDate} onChange={(value) => updateField('publicationDate', value)} inputType="date" />
              <Field label="Cover introduction label" id="intro-label" value={copy.introLabel} onChange={(value) => updateField('introLabel', value)} />
              <Field label="Cover introduction" id="intro" value={copy.intro} onChange={(value) => updateField('intro', value)} multiline />
            </section>
            {copy.sectionOrder.map((section) => section === 'context' ? contextEditor : implicationsEditor)}
            {chartEditor}
            {tableEditor}
            <div className="text-sections-label">Reusable text pages</div>
            <p className="data-table-help">Add, edit, or remove plain-text pages without changing code.</p>
            {textSectionEditors}
            <button type="button" className="studio-button add-row-button" onClick={addTextSection}>Add text page</button>
            <div className="editor-footnote">
              <Info size={14} />
              <span>Review every figure and statement against your approved source data before publication.</span>
            </div>
          </div>
        </aside>

        <section className="preview-area" aria-label="Live publication preview" data-testid="publication-preview">
          <div className="preview-toolbar">
            <h2>Document preview <span className="preview-count">{pageCount} pages · A4</span></h2>
            <div className="view-toggle" aria-label="Workspace view">
              <button type="button" className={!mobilePreview ? 'active' : ''} onClick={() => setMobilePreview(false)} data-testid="button-show-editor">Edit</button>
              <button type="button" className={mobilePreview ? 'active' : ''} onClick={() => setMobilePreview(true)} data-testid="button-show-preview">Preview</button>
            </div>
            <span className="eyebrow">Live preview</span>
          </div>
          <div className="report-stack">
            {visiblePageIds.map((id) => <Fragment key={id}>{previewPageById[id]}</Fragment>)}
          </div>
        </section>
      </main>
    </div>
  );
}

function Field({
  label,
  id,
  value,
  onChange,
  multiline = false,
  tall = false,
  compact = false,
  hint,
  placeholder,
  inputType = 'text',
}: {
  label: string;
  id: string;
  value: string;
  onChange: (value: string) => void;
  multiline?: boolean;
  tall?: boolean;
  compact?: boolean;
  hint?: string;
  placeholder?: string;
  inputType?: 'text' | 'date';
}) {
  const fieldId = `field-${id}`;
  return (
    <div className="field">
      <label htmlFor={fieldId}>{label}</label>
      {multiline ? (
        <textarea id={fieldId} className={`${tall ? 'tall' : ''} ${compact ? 'compact' : ''}`} value={value} placeholder={placeholder} onChange={(event) => onChange(event.target.value)} data-testid={`input-${id}`} aria-describedby={hint ? `${fieldId}-hint` : undefined} />
      ) : (
        <input id={fieldId} type={inputType} value={value} placeholder={placeholder} onChange={(event) => onChange(event.target.value)} data-testid={`input-${id}`} />
      )}
      {hint && <small id={`${fieldId}-hint`}>{hint}</small>}
    </div>
  );
}

function MagazineHeader({ title }: { title: string }) {
  return (
    <header className="mag-topline">
      <span className="mag-brandmark"><span className="mag-brand-symbol">i</span>IFRS Metrics</span>
      <span>{title}</span>
    </header>
  );
}

function MagazineFooter({ title, page }: { title: string; page: string }) {
  return (
    <footer className="mag-footer">
      <span>{title}</span>
      <span className="mag-page-number">{page}</span>
    </footer>
  );
}

export default PublicationStudio;
