import * as XLSX from 'xlsx';
import { FittingResult, DuctShape } from '../types';

export const generateScheduleExcel = (
  projectRows: FittingResult[],
  totalProjectLoss: number,
  showSiUnits: boolean,
  systemNo: string
) => {
  const unitSystem = showSiUnits ? 'SI' : 'IP';
  const flowUnit = showSiUnits ? 'CMH' : 'CFM';
  const lenUnit = showSiUnits ? 'mm' : 'in';
  const velUnit = showSiUnits ? 'm/s' : 'FPM';
  const pressUnit = showSiUnits ? 'Pa' : 'in. wg';

  const formatValNumber = (val: number, type: 'flow' | 'vel' | 'press', isSi: boolean): number => {
    if (isSi) {
      switch (type) {
        case 'flow': return Math.round(val * 1.699);
        case 'vel': return parseFloat((val * 0.00508).toFixed(2));
        case 'press': return parseFloat((val * 249.089).toFixed(2));
      }
    }
    if (type === 'flow') return Math.round(val);
    return parseFloat(val.toFixed(2));
  };

  const formatDimensions = (res: FittingResult, isSi: boolean): string => {
    if (res.customSize) return res.customSize;
    const f = (v: number | undefined) => {
      const num = v || 0;
      return isSi ? Math.round(num * 25.4).toString() : parseFloat(num.toFixed(2)).toString();
    };
    if (res.shape === DuctShape.ROUND) return `D${f(res.dimensions.diameter)}`;
    return `${f(res.dimensions.width)}x${f(res.dimensions.height)}`;
  };

  const getCleanedDescription = (row: FittingResult): string => {
    if (row.ashraeCode === 'Custom' || row.fittingType === 'Custom' || row.ashraeCode === 'User Defined' || row.fittingType === 'User Defined') {
      return '';
    }
    let text = row.aiReasoning || row.fittingDescription || '';

    // 1. Remove Fitting Type and ASHRAE No. (Prefix)
    if (row.ashraeCode) {
      const escapedCode = row.ashraeCode.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const prefixRegex = new RegExp(`^.*?${escapedCode}.*?:\\s*`, 'i');
      text = text.replace(prefixRegex, '');
    }

    if (text.startsWith(row.fittingType)) {
      text = text.substring(row.fittingType.length).replace(/^[:\s]+/, '');
    }

    // 2. Remove Loss Coefficient info (Suffix)
    const separators = [
      ' ->',
      '. Base Cp=',
      '. Re=',
      ': Loss coefficient',
      ' Loss coefficient',
      ' based on'
    ];

    for (const sep of separators) {
      const idx = text.indexOf(sep);
      if (idx !== -1) {
        text = text.substring(0, idx);
      }
    }

    text = text.replace(/^[,\s]+/, '');
    text = text.replace(/[.,\s]+$/, '');
    return text;
  };

  // Build sheet rows
  const dataRows: (string | number)[][] = [
    [`Total Pressure Loss Calculation (${unitSystem})`],
    [`System Tag: ${systemNo || 'N/A'}`],
    [], // Blank line
    [
      'No.',
      'ASHRAE No.',
      'Fitting Type',
      'Description / Specification',
      `Airflow (${flowUnit})`,
      `Size (${lenUnit})`,
      `Velocity (${velUnit})`,
      `Vel. Press. (${pressUnit})`,
      'Coeff.',
      `Loss (${pressUnit})`
    ]
  ];

  projectRows.forEach((row, index) => {
    const secondary = getCleanedDescription(row);
    dataRows.push([
      index + 1,
      row.ashraeCode || '-',
      row.fittingType,
      secondary || '-',
      formatValNumber(row.airflow, 'flow', showSiUnits),
      formatDimensions(row, showSiUnits),
      formatValNumber(row.velocity, 'vel', showSiUnits),
      formatValNumber(row.velocityPressure, 'press', showSiUnits),
      parseFloat(row.coefficient.toFixed(2)),
      formatValNumber(row.totalPressureLoss, 'press', showSiUnits)
    ]);
  });

  // Total summary row
  const totalLoss = formatValNumber(totalProjectLoss, 'press', showSiUnits);
  dataRows.push([
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    'Total Pressure Loss',
    totalLoss
  ]);

  const worksheet = XLSX.utils.aoa_to_sheet(dataRows);

  // Set column widths for clean readability
  worksheet['!cols'] = [
    { wch: 6 },  // No.
    { wch: 14 }, // ASHRAE No.
    { wch: 32 }, // Fitting Type
    { wch: 28 }, // Description
    { wch: 14 }, // Airflow
    { wch: 14 }, // Size
    { wch: 14 }, // Velocity
    { wch: 16 }, // Vel. Press.
    { wch: 10 }, // Coeff.
    { wch: 16 }  // Loss
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Duct Loss Schedule');

  const filename = `duct-loss-schedule-${systemNo || 'draft'}.xlsx`;
  XLSX.writeFile(workbook, filename);
};

export interface ImportScheduleResult {
  success: boolean;
  rows: FittingResult[];
  systemTag?: string;
  detectedUnitSystem?: 'SI' | 'IP';
  message: string;
}

export const parseScheduleExcel = async (
  file: File,
  currentShowSiUnits: boolean
): Promise<ImportScheduleResult> => {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      return { success: false, rows: [], message: 'The Excel file contains no worksheets.' };
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      return { success: false, rows: [], message: 'Could not read worksheet data from the file.' };
    }

    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    if (!rawRows || rawRows.length === 0) {
      return { success: false, rows: [], message: 'The selected spreadsheet is empty.' };
    }

    // 1. Detect System Tag from top rows
    let detectedSystemTag: string | undefined;
    for (let r = 0; r < Math.min(rawRows.length, 6); r++) {
      const rowText = rawRows[r].map(c => String(c)).join(' ');
      const match = rowText.match(/System\s*Tag\s*[:=]\s*(.+)/i);
      if (match) {
        const candidate = match[1].trim();
        if (candidate && candidate.toUpperCase() !== 'N/A') {
          detectedSystemTag = candidate;
        }
      }
    }

    // 2. Detect Unit System (SI vs IP)
    let isSi = currentShowSiUnits;
    let unitDetected: 'SI' | 'IP' | undefined;
    const headerSample = rawRows.slice(0, 10).map(r => r.join(' ')).join(' ');
    if (/\(SI\)|CMH|\(mm\)|\(Pa\)|m\/s/i.test(headerSample)) {
      isSi = true;
      unitDetected = 'SI';
    } else if (/\(IP\)|CFM|\(in\)|\(in\.?\s*wg\)|FPM/i.test(headerSample)) {
      isSi = false;
      unitDetected = 'IP';
    }

    // 3. Locate Header Row
    let headerRowIndex = -1;
    const colMap = {
      no: 0,
      ashrae: 1,
      type: 2,
      desc: 3,
      airflow: 4,
      size: 5,
      vel: 6,
      vp: 7,
      coeff: 8,
      loss: 9
    };

    for (let r = 0; r < Math.min(rawRows.length, 12); r++) {
      const row = rawRows[r].map(c => String(c).toLowerCase().trim());
      const hasAshrae = row.some(c => c.includes('ashrae'));
      const hasType = row.some(c => c.includes('fitting') || c.includes('type'));
      const hasLoss = row.some(c => c.includes('loss') || c.includes('press'));
      const hasFlow = row.some(c => c.includes('flow') || c.includes('cfm') || c.includes('cmh'));

      if ((hasAshrae || hasType) && (hasLoss || hasFlow)) {
        headerRowIndex = r;
        row.forEach((colName, idx) => {
          if (colName.includes('ashrae')) colMap.ashrae = idx;
          else if (colName.includes('fitting') || colName.includes('type')) colMap.type = idx;
          else if (colName.includes('desc') || colName.includes('spec')) colMap.desc = idx;
          else if (colName.includes('air') || colName.includes('flow') || colName.includes('cfm') || colName.includes('cmh')) colMap.airflow = idx;
          else if (colName.includes('size') || colName.includes('dim')) colMap.size = idx;
          else if (colName.includes('vel.') || colName.includes('vp') || colName.includes('velocity press')) colMap.vp = idx;
          else if (colName.includes('vel')) colMap.vel = idx;
          else if (colName.includes('coeff') || colName === 'c') colMap.coeff = idx;
          else if (colName.includes('loss') || colName.includes('dp')) colMap.loss = idx;
          else if (colName === 'no.' || colName === 'no') colMap.no = idx;
        });
        break;
      }
    }

    if (headerRowIndex === -1) {
      if (rawRows.length > 3 && rawRows[3].length >= 6) {
        headerRowIndex = 3;
      } else if (rawRows.length > 0 && rawRows[0].length >= 6) {
        headerRowIndex = 0;
      } else {
        return { success: false, rows: [], message: 'Could not identify table columns in the spreadsheet.' };
      }
    }

    // 4. Parse Rows
    const importedRows: FittingResult[] = [];
    const parseNum = (val: any): number => {
      if (typeof val === 'number') return isNaN(val) ? 0 : val;
      if (!val) return 0;
      const cleaned = String(val).replace(/,/g, '').trim();
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    };

    for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || row.length === 0) continue;

      const rowJoined = row.map(c => String(c)).join(' ');
      // Skip Total row
      if (/Total\s*Pressure\s*Loss/i.test(rowJoined) || /^Total/i.test(String(row[colMap.coeff] || ''))) {
        continue;
      }

      // Check if row has meaningful data
      const hasValues = row.some((c: any) => c !== '' && c !== null && c !== undefined);
      if (!hasValues) continue;

      const rawAshrae = String(row[colMap.ashrae] ?? '').trim();
      const rawType = String(row[colMap.type] ?? '').trim();
      const rawDesc = String(row[colMap.desc] ?? '').trim();
      const rawAirflow = row[colMap.airflow];
      const rawSize = String(row[colMap.size] ?? '').trim();
      const rawVel = row[colMap.vel];
      const rawVp = row[colMap.vp];
      const rawCoeff = row[colMap.coeff];
      const rawLoss = row[colMap.loss];

      if (!rawAshrae && !rawType && rawLoss === '') continue;

      const airflowVal = parseNum(rawAirflow);
      const velVal = parseNum(rawVel);
      const vpVal = parseNum(rawVp);
      const coeffVal = parseNum(rawCoeff);
      const lossVal = parseNum(rawLoss);

      // Convert to internal IP units
      const airflow = isSi ? (airflowVal / 1.699) : airflowVal;
      const velocity = isSi ? (velVal / 0.00508) : velVal;
      const velocityPressure = isSi ? (vpVal / 249.089) : vpVal;
      const totalPressureLoss = isSi ? (lossVal / 249.089) : lossVal;
      const coefficient = coeffVal;

      const ashraeCode = (rawAshrae === '-' || !rawAshrae) ? 'User Defined' : rawAshrae;
      const fittingType = rawType || 'User Defined';

      let width = 0;
      let height = 0;
      let diameter = 0;
      let shape = DuctShape.RECTANGULAR;

      const rectMatch = rawSize.match(/^(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)$/);
      const roundMatch = rawSize.match(/^[dDøØ]?\s*(\d+(?:\.\d+)?)$/);

      if (rectMatch) {
        const w = parseFloat(rectMatch[1]);
        const h = parseFloat(rectMatch[2]);
        width = isSi ? w / 25.4 : w;
        height = isSi ? h / 25.4 : h;
        shape = DuctShape.RECTANGULAR;
      } else if (roundMatch && (rawSize.toLowerCase().startsWith('d') || rawSize.startsWith('ø') || rawSize.startsWith('Ø') || ashraeCode === 'CD8-5' || ashraeCode === 'ED4-2' || ashraeCode === 'SR4-3')) {
        const d = parseFloat(roundMatch[1]);
        diameter = isSi ? d / 25.4 : d;
        shape = DuctShape.ROUND;
      }

      const cleanSecondary = (rawDesc === '-' ? '' : rawDesc);
      const fittingDescription = cleanSecondary 
        ? `${fittingType}, ${cleanSecondary}` 
        : `${fittingType}${rawSize ? `, ${rawSize}` : ''}`;

      const generatedId = (typeof crypto !== 'undefined' && crypto.randomUUID) 
        ? crypto.randomUUID() 
        : `import-${Date.now()}-${r}-${Math.random().toString(36).substring(2, 7)}`;

      importedRows.push({
        id: generatedId,
        ashraeCode,
        fittingType,
        shape,
        dimensions: { width, height, diameter },
        customSize: rawSize || undefined,
        airflow,
        velocity,
        velocityPressure,
        coefficient,
        totalPressureLoss,
        fittingDescription,
        aiReasoning: cleanSecondary,
        timestamp: Date.now() + r
      });
    }

    if (importedRows.length === 0) {
      return { success: false, rows: [], message: 'No valid fitting rows were found in the spreadsheet.' };
    }

    return {
      success: true,
      rows: importedRows,
      systemTag: detectedSystemTag,
      detectedUnitSystem: unitDetected,
      message: `Successfully imported ${importedRows.length} fitting${importedRows.length === 1 ? '' : 's'}.`
    };
  } catch (err: any) {
    return {
      success: false,
      rows: [],
      message: err?.message || 'An error occurred while parsing the Excel file.'
    };
  }
};
