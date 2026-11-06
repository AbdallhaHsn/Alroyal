import ExcelJS from 'exceljs';

export interface ExportParams {
  tbGroupedData: any[];
  processedTrialBalanceRows: any[];
  tbGrandTotals: any;
  tbStartDate: string;
  tbEndDate: string;
  tbCurrency: 'EGP' | 'USD' | 'BOTH';
  tbMode: 'detailed' | 'summary';
  tbGroupBy: 'mainParent' | 'root' | 'subParent' | 'none';
}

export async function exportStyledTrialBalanceExcel(params: ExportParams) {
  const {
    tbGroupedData,
    processedTrialBalanceRows,
    tbGrandTotals,
    tbStartDate,
    tbEndDate,
    tbCurrency,
    tbMode,
    tbGroupBy
  } = params;

  const isDolar = tbCurrency === 'USD';
  const isBoth = tbCurrency === 'BOTH';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'الرويال للتطوير العقاري - ERP System';
  workbook.lastModifiedBy = 'النظام المالي الموحد';
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet('ميزان المراجعة بالأرصدة', {
    views: [{ rightToLeft: true, showGridLines: true }]
  });

  // Calculate dynamic columns
  // EGP or USD: 4 info + 1 Open + 2 Movement + 1 Ending = 8 cols
  // BOTH: 4 info + 2 Open + 4 Movement + 2 Ending = 12 cols
  const totalCols = isBoth ? 12 : 8;

  const fontFamily = 'Segoe UI';

  // Fonts
  const fontTitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontSubtitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontMeta: Partial<ExcelJS.Font> = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF1E293B' } };
  const fontHeader: Partial<ExcelJS.Font> = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontGrandTotal: Partial<ExcelJS.Font> = { name: fontFamily, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };

  // Fills
  const fillPrimaryDark: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  const fillPrimaryMedium: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
  const fillMetaBar: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  const fillGroupHeader: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
  const fillSubtotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  const fillGrandTotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  const fillZebraEven: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  const fillZebraOdd: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  // Tinted Fills for Debit & Credit
  const fillDebitTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F2' } };
  const fillCreditTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };

  // Borders
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const doubleBottomBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF0F172A' } },
    bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF0F172A' } },
    right: { style: 'thin', color: { argb: 'FF0F172A' } }
  };

  // Row 1: Top spacing
  worksheet.addRow([]);
  worksheet.getRow(1).height = 10;

  // Row 2: Company Header Banner
  const titleRow = worksheet.addRow([]);
  titleRow.height = 36;
  worksheet.mergeCells(2, 1, 2, totalCols);
  const titleCell = worksheet.getCell(2, 1);
  titleCell.value = 'الرويال للتطوير العقاري - الإدارة المالية والحسابات العامة';
  titleCell.font = fontTitle;
  titleCell.fill = fillPrimaryDark;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 3: Subtitle
  const subRow = worksheet.addRow([]);
  subRow.height = 28;
  worksheet.mergeCells(3, 1, 3, totalCols);
  const subCell = worksheet.getCell(3, 1);
  const groupLabel = tbGroupBy === 'mainParent' ? 'مجمع حسب الحساب الرئيسي الكبير' :
                     tbGroupBy === 'root' ? 'مجمع حسب التصنيف الجذري' :
                     tbGroupBy === 'subParent' ? 'مجمع حسب المجموعة الفرعية' : 'عرض مسطح مباشر';
  subCell.value = `تقرير ميزان المراجعة بالأرصدة الصافية وحركات الفترة (${groupLabel} - ${tbMode === 'detailed' ? 'تفصيلي' : 'تجميعي'})`;
  subCell.font = fontSubtitle;
  subCell.fill = fillPrimaryMedium;
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 4: Metadata Bar
  const metaRow = worksheet.addRow([]);
  metaRow.height = 24;
  worksheet.mergeCells(4, 1, 4, totalCols);
  const metaCell = worksheet.getCell(4, 1);
  const currLabel = tbCurrency === 'EGP' ? 'الجنيه السوداني (SDG)' : tbCurrency === 'USD' ? 'الدولار الأمريكي ($ USD)' : 'الجنيه والدولار معاً';
  metaCell.value = `الفترة المالية: من ${tbStartDate} إلى ${tbEndDate}   |   العملة: ${currLabel}   |   الرصيد: صافي مدين (أحمر) أو دائن (أخضر)   |   تاريخ الاستخراج: ${new Date().toLocaleString('ar-EG')}`;
  metaCell.font = fontMeta;
  metaCell.fill = fillMetaBar;
  metaCell.alignment = { horizontal: 'center', vertical: 'middle' };
  metaCell.border = thinBorder;

  // Row 5: Spacing
  worksheet.addRow([]);
  worksheet.getRow(5).height = 10;

  // Row 6 & 7: Headers
  const headerRow1 = worksheet.addRow([]);
  headerRow1.height = 26;
  const headerRow2 = worksheet.addRow([]);
  headerRow2.height = 24;

  // Info Columns 1..4 (Merged rows 6 & 7)
  const metaCols = [
    { text: 'رقم الحساب', col: 1 },
    { text: 'اسم الحساب المالي', col: 2 },
    { text: 'المجموعة التحليلية', col: 3 },
    { text: 'التصنيف', col: 4 }
  ];
  metaCols.forEach(m => {
    worksheet.mergeCells(6, m.col, 7, m.col);
    const c = worksheet.getCell(6, m.col);
    c.value = m.text;
    c.font = fontHeader;
    c.fill = fillPrimaryDark;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = thinBorder;
  });

  if (!isBoth) {
    const curSymbol = isDolar ? '$' : 'ج.س';

    // 1. Opening Net Balance (Col 5)
    worksheet.mergeCells(6, 5, 7, 5);
    const cOpen = worksheet.getCell(6, 5);
    cOpen.value = `الرصيد الافتتاحي (${curSymbol})`;
    cOpen.font = fontHeader;
    cOpen.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
    cOpen.alignment = { horizontal: 'center', vertical: 'middle' };
    cOpen.border = thinBorder;

    // 2. Period Movements (Cols 6 & 7)
    worksheet.mergeCells(6, 6, 6, 7);
    const cPeriod = worksheet.getCell(6, 6);
    cPeriod.value = `حركات الفترة (${tbStartDate} إلى ${tbEndDate})`;
    cPeriod.font = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
    cPeriod.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
    cPeriod.alignment = { horizontal: 'center', vertical: 'middle' };
    cPeriod.border = thinBorder;

    const subDr = worksheet.getCell(7, 6);
    subDr.value = `حركات مدين (${curSymbol})`;
    subDr.font = fontHeader;
    subDr.fill = fillPrimaryMedium;
    subDr.alignment = { horizontal: 'center', vertical: 'middle' };
    subDr.border = thinBorder;

    const subCr = worksheet.getCell(7, 7);
    subCr.value = `حركات دائن (${curSymbol})`;
    subCr.font = fontHeader;
    subCr.fill = fillPrimaryMedium;
    subCr.alignment = { horizontal: 'center', vertical: 'middle' };
    subCr.border = thinBorder;

    // 3. Ending Net Balance (Col 8)
    worksheet.mergeCells(6, 8, 7, 8);
    const cEnd = worksheet.getCell(6, 8);
    cEnd.value = `الرصيد الختامي (${curSymbol})`;
    cEnd.font = fontHeader;
    cEnd.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
    cEnd.alignment = { horizontal: 'center', vertical: 'middle' };
    cEnd.border = thinBorder;
  } else {
    // DUAL CURRENCY (BOTH)
    // 1. Opening (Cols 5 & 6)
    worksheet.mergeCells(6, 5, 6, 6);
    const cOpen = worksheet.getCell(6, 5);
    cOpen.value = 'الرصيد الافتتاحي';
    cOpen.font = fontHeader;
    cOpen.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
    cOpen.alignment = { horizontal: 'center', vertical: 'middle' };
    cOpen.border = thinBorder;

    const o1 = worksheet.getCell(7, 5);
    o1.value = 'الرصيد (ج.س)';
    o1.font = fontHeader;
    o1.fill = fillPrimaryMedium;
    o1.alignment = { horizontal: 'center', vertical: 'middle' };
    o1.border = thinBorder;

    const o2 = worksheet.getCell(7, 6);
    o2.value = 'الرصيد ($)';
    o2.font = fontHeader;
    o2.fill = fillPrimaryDark;
    o2.alignment = { horizontal: 'center', vertical: 'middle' };
    o2.border = thinBorder;

    // 2. Period Movements (Cols 7..10)
    worksheet.mergeCells(6, 7, 6, 10);
    const cPeriod = worksheet.getCell(6, 7);
    cPeriod.value = `حركات الفترة (${tbStartDate} إلى ${tbEndDate})`;
    cPeriod.font = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
    cPeriod.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
    cPeriod.alignment = { horizontal: 'center', vertical: 'middle' };
    cPeriod.border = thinBorder;

    const pHeaders = ['مدين (ج.س)', 'دائن (ج.س)', 'مدين ($)', 'دائن ($)'];
    pHeaders.forEach((ph, i) => {
      const pc = worksheet.getCell(7, 7 + i);
      pc.value = ph;
      pc.font = fontHeader;
      pc.fill = (i < 2) ? fillPrimaryMedium : fillPrimaryDark;
      pc.alignment = { horizontal: 'center', vertical: 'middle' };
      pc.border = thinBorder;
    });

    // 3. Ending (Cols 11 & 12)
    worksheet.mergeCells(6, 11, 6, 12);
    const cEnd = worksheet.getCell(6, 11);
    cEnd.value = 'الرصيد الختامي';
    cEnd.font = fontHeader;
    cEnd.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
    cEnd.alignment = { horizontal: 'center', vertical: 'middle' };
    cEnd.border = thinBorder;

    const e1 = worksheet.getCell(7, 11);
    e1.value = 'الرصيد (ج.س)';
    e1.font = fontHeader;
    e1.fill = fillPrimaryMedium;
    e1.alignment = { horizontal: 'center', vertical: 'middle' };
    e1.border = thinBorder;

    const e2 = worksheet.getCell(7, 12);
    e2.value = 'الرصيد ($)';
    e2.font = fontHeader;
    e2.fill = fillPrimaryDark;
    e2.alignment = { horizontal: 'center', vertical: 'middle' };
    e2.border = thinBorder;
  }

  let currentExcelRow = 8;

  // Helper to write Net Balance Cell with distinctive Color
  const writeNetBalanceCell = (
    rowObj: ExcelJS.Row,
    colNumber: number,
    netVal: number,
    isBold = true,
    customFill?: ExcelJS.Fill
  ) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;

    if (Math.abs(netVal) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
      if (customFill) cell.fill = customFill;
      return;
    }

    if (netVal > 0) {
      // DEBTOR (مدين) - RED / ROSE
      cell.value = Number(netVal);
      cell.numFmt = '#,##0.00 " (مدين)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FFE11D48' } }; // Rose 600
      cell.fill = customFill || fillDebitTint;
    } else {
      // CREDITOR (دائن) - GREEN / EMERALD
      cell.value = Number(Math.abs(netVal));
      cell.numFmt = '#,##0.00 " (دائن)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FF059669' } }; // Emerald 600
      cell.fill = customFill || fillCreditTint;
    }
  };

  // Helper to write period movement amounts
  const writeMovementCell = (
    rowObj: ExcelJS.Row,
    colNumber: number,
    val: number,
    colorARGB: string,
    isBold: boolean,
    customFill?: ExcelJS.Fill
  ) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;
    cell.numFmt = '#,##0.00';
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: colorARGB } };
    if (customFill) cell.fill = customFill;

    if (Math.abs(val) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
    } else {
      cell.value = Number(val);
    }
  };

  // Helper to write full row amounts
  const writeRowData = (
    rowObj: ExcelJS.Row,
    dataObj: any,
    isBold: boolean,
    bgFill?: ExcelJS.Fill
  ) => {
    // 1. OPENING
    const openVal = isDolar 
      ? ((dataObj.openDebtorDolar || 0) - (dataObj.openCreditorDolar || 0))
      : ((dataObj.openDebtor || 0) - (dataObj.openCreditor || 0));

    const openValDolar = (dataObj.openDebtorDolar || 0) - (dataObj.openCreditorDolar || 0);

    // 2. PERIOD
    const pDr = isDolar ? (dataObj.periodDebtorDolar || 0) : (dataObj.periodDebtor || 0);
    const pCr = isDolar ? (dataObj.periodCreditorDolar || 0) : (dataObj.periodCreditor || 0);
    const pDrD = dataObj.periodDebtorDolar || 0;
    const pCrD = dataObj.periodCreditorDolar || 0;

    // 3. ENDING
    const endVal = isDolar
      ? (((dataObj.openDebtorDolar || 0) + (dataObj.periodDebtorDolar || 0)) - ((dataObj.openCreditorDolar || 0) + (dataObj.periodCreditorDolar || 0)))
      : (((dataObj.openDebtor || 0) + (dataObj.periodDebtor || 0)) - ((dataObj.openCreditor || 0) + (dataObj.periodCreditor || 0)));

    const endValDolar = (((dataObj.openDebtorDolar || 0) + (dataObj.periodDebtorDolar || 0)) - ((dataObj.openCreditorDolar || 0) + (dataObj.periodCreditorDolar || 0)));

    if (!isBoth) {
      writeNetBalanceCell(rowObj, 5, openVal, isBold, bgFill);
      writeMovementCell(rowObj, 6, pDr, isBold ? 'FF0F172A' : 'FFE11D48', isBold, bgFill);
      writeMovementCell(rowObj, 7, pCr, isBold ? 'FF0F172A' : 'FF059669', isBold, bgFill);
      writeNetBalanceCell(rowObj, 8, endVal, isBold, bgFill);
    } else {
      writeNetBalanceCell(rowObj, 5, (dataObj.openDebtor || 0) - (dataObj.openCreditor || 0), isBold, bgFill);
      writeNetBalanceCell(rowObj, 6, openValDolar, isBold, bgFill);
      writeMovementCell(rowObj, 7, dataObj.periodDebtor || 0, isBold ? 'FF0F172A' : 'FFE11D48', isBold, bgFill);
      writeMovementCell(rowObj, 8, dataObj.periodCreditor || 0, isBold ? 'FF0F172A' : 'FF059669', isBold, bgFill);
      writeMovementCell(rowObj, 9, pDrD, isBold ? 'FF0F172A' : 'FF2563EB', isBold, bgFill);
      writeMovementCell(rowObj, 10, pCrD, isBold ? 'FF0F172A' : 'FF2563EB', isBold, bgFill);
      writeNetBalanceCell(rowObj, 11, (((dataObj.openDebtor || 0) + (dataObj.periodDebtor || 0)) - ((dataObj.openCreditor || 0) + (dataObj.periodCreditor || 0))), isBold, bgFill);
      writeNetBalanceCell(rowObj, 12, endValDolar, isBold, bgFill);
    }
  };

  // Populate Groups and Child Rows
  tbGroupedData.forEach((grp) => {
    // 1. Group Header Row
    const grpRow = worksheet.addRow([]);
    grpRow.height = 24;

    const gCode = grpRow.getCell(1);
    gCode.value = `[ ${grp.groupCode} ]`;
    gCode.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF1E3A8A' } };
    gCode.alignment = { horizontal: 'center', vertical: 'middle' };
    gCode.fill = fillGroupHeader;
    gCode.border = thinBorder;

    const gName = grpRow.getCell(2);
    gName.value = `مجموعة: ${grp.groupName} (${grp.rows.length} حساب فرعي)`;
    gName.font = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FF0F172A' } };
    gName.alignment = { horizontal: 'right', vertical: 'middle' };
    gName.fill = fillGroupHeader;
    gName.border = thinBorder;

    const gSub = grpRow.getCell(3);
    gSub.value = grp.groupName;
    gSub.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF475569' } };
    gSub.alignment = { horizontal: 'right', vertical: 'middle' };
    gSub.fill = fillGroupHeader;
    gSub.border = thinBorder;

    const gCat = grpRow.getCell(4);
    gCat.value = grp.category;
    gCat.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF1E3A8A' } };
    gCat.alignment = { horizontal: 'center', vertical: 'middle' };
    gCat.fill = fillGroupHeader;
    gCat.border = thinBorder;

    writeRowData(grpRow, grp.totals, true, fillGroupHeader);
    currentExcelRow++;

    // 2. Child Leaf Accounts (if detailed)
    if (tbMode === 'detailed') {
      grp.rows.forEach((row: any, rIdx: number) => {
        const itemRow = worksheet.addRow([]);
        itemRow.height = 20;
        const isEven = rIdx % 2 === 0;
        const rowFill = isEven ? fillZebraEven : fillZebraOdd;

        const cNum = itemRow.getCell(1);
        cNum.value = `   ${row.AccountNumber}`;
        cNum.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF2563EB' } };
        cNum.alignment = { horizontal: 'right', vertical: 'middle' };
        cNum.fill = rowFill;
        cNum.border = thinBorder;

        const cName = itemRow.getCell(2);
        cName.value = row.AccountName;
        cName.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
        cName.alignment = { horizontal: 'right', vertical: 'middle' };
        cName.fill = rowFill;
        cName.border = thinBorder;

        const cSub = itemRow.getCell(3);
        cSub.value = row.SubParentName || row.ParentGroupName || '-';
        cSub.font = { name: fontFamily, size: 8.5, color: { argb: 'FF64748B' } };
        cSub.alignment = { horizontal: 'right', vertical: 'middle' };
        cSub.fill = rowFill;
        cSub.border = thinBorder;

        const cCat = itemRow.getCell(4);
        cCat.value = row.category;
        cCat.font = { name: fontFamily, size: 8.5, color: { argb: 'FF64748B' } };
        cCat.alignment = { horizontal: 'center', vertical: 'middle' };
        cCat.fill = rowFill;
        cCat.border = thinBorder;

        writeRowData(itemRow, row, false, undefined);
        currentExcelRow++;
      });
    }

    // 3. Group Subtotal Row
    const subtotalRow = worksheet.addRow([]);
    subtotalRow.height = 22;

    const sCode = subtotalRow.getCell(1);
    sCode.value = `إجمالي [ ${grp.groupCode} ]`;
    sCode.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E3A8A' } };
    sCode.alignment = { horizontal: 'center', vertical: 'middle' };
    sCode.fill = fillSubtotal;
    sCode.border = thinBorder;

    const sName = subtotalRow.getCell(2);
    sName.value = `إجمالي صافي مجموعة: ${grp.groupName}`;
    sName.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E3A8A' } };
    sName.alignment = { horizontal: 'right', vertical: 'middle' };
    sName.fill = fillSubtotal;
    sName.border = thinBorder;

    const sSub = subtotalRow.getCell(3);
    sSub.value = '';
    sSub.fill = fillSubtotal;
    sSub.border = thinBorder;

    const sCat = subtotalRow.getCell(4);
    sCat.value = grp.category;
    sCat.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF64748B' } };
    sCat.alignment = { horizontal: 'center', vertical: 'middle' };
    sCat.fill = fillSubtotal;
    sCat.border = thinBorder;

    writeRowData(subtotalRow, grp.totals, true, fillSubtotal);
    currentExcelRow++;

    // Blank spacing row
    const spRow = worksheet.addRow([]);
    spRow.height = 8;
    currentExcelRow++;
  });

  // GRAND TOTAL ROW (Executive Double Border)
  const grandRow = worksheet.addRow([]);
  grandRow.height = 30;

  worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, 4);
  const grandTitle = worksheet.getCell(currentExcelRow, 1);
  grandTitle.value = 'المجموع العام لميزان المراجعة (صافي الأرصدة وحركات الفترة)';
  grandTitle.font = fontGrandTotal;
  grandTitle.fill = fillGrandTotal;
  grandTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  grandTitle.border = doubleBottomBorder;

  writeRowData(grandRow, tbGrandTotals, true, fillGrandTotal);
  for (let c = 5; c <= totalCols; c++) {
    grandRow.getCell(c).border = doubleBottomBorder;
  }
  currentExcelRow++;

  // SIGNATURES & AUDIT APPROVAL BLOCK
  worksheet.addRow([]);
  currentExcelRow++;
  const signRow1 = worksheet.addRow([]);
  signRow1.height = 26;
  const signRow2 = worksheet.addRow([]);
  signRow2.height = 34;

  const signCols = [
    { title: 'إعداد / المحاسب المسؤول', col: 1 },
    { title: 'مراجعة / المراجع الداخلي', col: 3 },
    { title: 'اعتماد / المدير المالي (CFO)', col: 5 },
    { title: 'مصادقة / الإدارة العامة', col: totalCols - 1 }
  ];

  signCols.forEach((s) => {
    if (s.col <= totalCols) {
      const cell1 = signRow1.getCell(s.col);
      cell1.value = s.title;
      cell1.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF334155' } };
      cell1.alignment = { horizontal: 'center', vertical: 'middle' };

      const cell2 = signRow2.getCell(s.col);
      cell2.value = 'التوقيع: ........................................';
      cell2.font = { name: fontFamily, size: 9, color: { argb: 'FF94A3B8' } };
      cell2.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  });

  // Auto-fit Column Widths (Generous padding so zero text/number is clipped!)
  worksheet.getColumn(1).width = 24; // رقم الحساب
  worksheet.getColumn(2).width = 40; // اسم الحساب المالي
  worksheet.getColumn(3).width = 28; // المجموعة التحليلية
  worksheet.getColumn(4).width = 16; // التصنيف

  for (let c = 5; c <= totalCols; c++) {
    worksheet.getColumn(c).width = 24; // Wide enough for amount + " (مدين)" / " (دائن)"
  }

  // Trigger browser download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `ميزان_المراجعة_${tbGroupBy}_${tbStartDate}_${tbEndDate}.xlsx`;
  a.click();
  window.URL.revokeObjectURL(url);
}

export async function exportStyledTransactionsExcel(entries: any[], filename = 'سجل_القيود_المحاسبية') {
  if (!entries || !entries.length) return;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'الرويال للتطوير العقاري - ERP System';
  const worksheet = workbook.addWorksheet('سجل القيود', {
    views: [{ rightToLeft: true, showGridLines: true }]
  });

  const fontFamily = 'Segoe UI';
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  // Header Banner
  worksheet.addRow([]);
  worksheet.mergeCells('A2:I2');
  const tCell = worksheet.getCell('A2');
  tCell.value = 'الرويال للتطوير العقاري - سجل دفتر القيود والسندات المجمعة';
  tCell.font = { name: fontFamily, size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
  tCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  tCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(2).height = 32;

  worksheet.mergeCells('A3:I3');
  const mCell = worksheet.getCell('A3');
  mCell.value = `تاريخ التصدير: ${new Date().toLocaleString('ar-EG')}   |   إجمالي القيود المصدرة: ${entries.length} قيد`;
  mCell.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
  mCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  mCell.alignment = { horizontal: 'center', vertical: 'middle' };
  mCell.border = thinBorder;
  worksheet.getRow(3).height = 22;

  worksheet.addRow([]);

  // Column Headers
  const hRow = worksheet.addRow([
    'رقم القيد / السند',
    'نوع المستند',
    'تاريخ القيد',
    'عدد البنود',
    'إجمالي مدين (ج.س)',
    'إجمالي دائن (ج.س)',
    'إجمالي مدين ($)',
    'إجمالي دائن ($)',
    'البيان التوضيحي'
  ]);
  hRow.height = 24;
  hRow.eachCell((cell) => {
    cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder;
  });

  const numFormat = '#,##0.00';
  let totalDeb = 0;
  let totalCred = 0;
  let totalDebD = 0;
  let totalCredD = 0;

  entries.forEach((e, idx) => {
    const isEven = idx % 2 === 0;
    const bgFill: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isEven ? 'FFF8FAFC' : 'FFFFFFFF' }
    };

    const row = worksheet.addRow([
      e.ReqNum || e.Id,
      e.docType || e.DocType || 'سند',
      e.theDate || e.TheDate || '',
      e.itemsCount || 1,
      Number(e.totalDebtor || e.Debtor || 0),
      Number(e.totalCreditor || e.Creditor || 0),
      Number(e.totalDebtorDolar || e.DebtorDolar || 0),
      Number(e.totalCreditorDolar || e.CreditorDolar || 0),
      e.sampleDetail || e.Detail || ''
    ]);
    row.height = 20;

    totalDeb += Number(e.totalDebtor || e.Debtor || 0);
    totalCred += Number(e.totalCreditor || e.Creditor || 0);
    totalDebD += Number(e.totalDebtorDolar || e.DebtorDolar || 0);
    totalCredD += Number(e.totalCreditorDolar || e.CreditorDolar || 0);

    row.eachCell((cell, colNumber) => {
      cell.border = thinBorder;
      cell.fill = bgFill;
      cell.font = { name: fontFamily, size: 9 };
      if (colNumber === 1) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF1E40AF' } };
      } else if (colNumber >= 5 && colNumber <= 8) {
        cell.numFmt = numFormat;
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      }
    });
  });

  // Grand Total Row
  const gRow = worksheet.addRow([
    'المجموع العام',
    '',
    '',
    `${entries.length} قيد`,
    totalDeb,
    totalCred,
    totalDebD,
    totalCredD,
    ''
  ]);
  gRow.height = 26;
  gRow.eachCell((cell, colNumber) => {
    cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.border = thinBorder;
    if (colNumber >= 5 && colNumber <= 8) {
      cell.numFmt = numFormat;
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    } else {
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  });

  worksheet.getColumn(1).width = 18;
  worksheet.getColumn(2).width = 16;
  worksheet.getColumn(3).width = 16;
  worksheet.getColumn(4).width = 12;
  worksheet.getColumn(5).width = 20;
  worksheet.getColumn(6).width = 20;
  worksheet.getColumn(7).width = 20;
  worksheet.getColumn(8).width = 20;
  worksheet.getColumn(9).width = 35;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}.xlsx`;
  a.click();
  window.URL.revokeObjectURL(url);
}

export interface GeneralLedgerExportParams {
  accounts: any[];
  grandTotals: any;
  startDate: string;
  endDate: string;
  currency: 'EGP' | 'USD' | 'BOTH';
  mode: 'detailed' | 'summary';
}

export async function exportStyledGeneralLedgerExcel(params: GeneralLedgerExportParams) {
  const {
    accounts,
    grandTotals,
    startDate,
    endDate,
    currency,
    mode
  } = params;

  if (!accounts || !accounts.length) return;

  const isDolar = currency === 'USD';
  const isBoth = currency === 'BOTH';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'الرويال للتطوير العقاري - ERP System';
  workbook.lastModifiedBy = 'نظام كشف الحسابات العامة';
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet('كشف الحسابات العامة', {
    views: [{ rightToLeft: true, showGridLines: true }]
  });

  const fontFamily = 'Segoe UI';

  // Fonts
  const fontTitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontSubtitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontMeta: Partial<ExcelJS.Font> = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF1E293B' } };
  const fontHeader: Partial<ExcelJS.Font> = { name: fontFamily, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontGrandTotal: Partial<ExcelJS.Font> = { name: fontFamily, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };

  // Fills
  const fillPrimaryDark: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  const fillPrimaryMedium: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
  const fillMetaBar: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  const fillGroupHeader: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
  const fillSubtotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  const fillGrandTotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  const fillZebraEven: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  const fillZebraOdd: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  const fillDebitTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F2' } };
  const fillCreditTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const doubleBottomBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF0F172A' } },
    bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF0F172A' } },
    right: { style: 'thin', color: { argb: 'FF0F172A' } }
  };

  // Determine total columns
  // Detailed:
  // Non-both: Date, ReqNum, DocType, AccountName, Detail, Debtor, Creditor, RunningBalance = 8 cols
  // Both: Date, ReqNum, DocType, AccountName, Detail, Dr(EGP), Cr(EGP), Dr($), Cr($), Running(EGP), Running($) = 11 cols
  // Summary:
  // Non-both: Code, AccountName, OpenNet, PeriodDr, PeriodCr, EndingNet = 6 cols
  // Both: Code, AccountName, Open(EGP), Open($), Dr(EGP), Cr(EGP), Dr($), Cr($), End(EGP), End($) = 10 cols
  const totalCols = mode === 'detailed' ? (isBoth ? 11 : 8) : (isBoth ? 10 : 6);

  // Row 1: Spacing
  worksheet.addRow([]);
  worksheet.getRow(1).height = 10;

  // Row 2: Company Header Banner
  const titleRow = worksheet.addRow([]);
  titleRow.height = 36;
  worksheet.mergeCells(2, 1, 2, totalCols);
  const titleCell = worksheet.getCell(2, 1);
  titleCell.value = 'الرويال للتطوير العقاري - الإدارة المالية والحسابات العامة';
  titleCell.font = fontTitle;
  titleCell.fill = fillPrimaryDark;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 3: Subtitle
  const subRow = worksheet.addRow([]);
  subRow.height = 28;
  worksheet.mergeCells(3, 1, 3, totalCols);
  const subCell = worksheet.getCell(3, 1);
  subCell.value = `تقرير كشف الحسابات العامة (General Ledger - ${mode === 'detailed' ? 'تفصيلي بالقيود والحركات' : 'تجميعي بالحسابات الفرعية والأرصدة'})`;
  subCell.font = fontSubtitle;
  subCell.fill = fillPrimaryMedium;
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 4: Metadata Bar
  const metaRow = worksheet.addRow([]);
  metaRow.height = 24;
  worksheet.mergeCells(4, 1, 4, totalCols);
  const metaCell = worksheet.getCell(4, 1);
  const currLabel = currency === 'EGP' ? 'الجنيه السوداني (SDG)' : currency === 'USD' ? 'الدولار الأمريكي ($ USD)' : 'الجنيه والدولار معاً';
  metaCell.value = `الفترة المالية: من ${startDate} إلى ${endDate}   |   العملة: ${currLabel}   |   عدد الحسابات العامة: ${accounts.length}   |   صافي الرصيد: أحمر (مدين) / أخضر (دائن)`;
  metaCell.font = fontMeta;
  metaCell.fill = fillMetaBar;
  metaCell.alignment = { horizontal: 'center', vertical: 'middle' };
  metaCell.border = thinBorder;

  // Row 5: Spacing
  worksheet.addRow([]);
  worksheet.getRow(5).height = 10;

  let currentExcelRow = 6;

  // Helper to format net badge cell in Excel
  const writeNetCell = (rowObj: ExcelJS.Row, colNumber: number, val: number, isBold = true, customFill?: ExcelJS.Fill) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;

    if (Math.abs(val) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
      if (customFill) cell.fill = customFill;
      return;
    }

    if (val > 0) {
      cell.value = Number(val);
      cell.numFmt = '#,##0.00 " (مدين)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FFE11D48' } };
      cell.fill = customFill || fillDebitTint;
    } else {
      cell.value = Number(Math.abs(val));
      cell.numFmt = '#,##0.00 " (دائن)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FF059669' } };
      cell.fill = customFill || fillCreditTint;
    }
  };

  const writeNumCell = (rowObj: ExcelJS.Row, colNumber: number, val: number, colorARGB: string, isBold = false, customFill?: ExcelJS.Fill) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;
    cell.numFmt = '#,##0.00';
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: colorARGB } };
    if (customFill) cell.fill = customFill;

    if (Math.abs(val) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
    } else {
      cell.value = Number(val);
    }
  };

  // Loop through selected general accounts
  accounts.forEach((acc) => {
    // 1. Account Section Banner
    const bannerRow = worksheet.addRow([]);
    bannerRow.height = 26;
    worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, totalCols);
    const bannerCell = worksheet.getCell(currentExcelRow, 1);
    
    const openLabel = isDolar 
      ? (acc.netOpenDolar > 0 ? `صافي رصيد سابق: $${Number(acc.netOpenDolar).toLocaleString()} (مدين)` : acc.netOpenDolar < 0 ? `صافي رصيد سابق: $${Number(Math.abs(acc.netOpenDolar)).toLocaleString()} (دائن)` : 'صافي رصيد سابق: متزن')
      : (acc.netOpen > 0 ? `صافي رصيد سابق: ${Number(acc.netOpen).toLocaleString()} ج.س (مدين)` : acc.netOpen < 0 ? `صافي رصيد سابق: ${Number(Math.abs(acc.netOpen)).toLocaleString()} ج.س (دائن)` : 'صافي رصيد سابق: متزن');

    bannerCell.value = `[ ${acc.code} ]  ${acc.name}   |   التصنيف: ${acc.category}   |   ${openLabel}`;
    bannerCell.font = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
    bannerCell.fill = fillGroupHeader;
    bannerCell.alignment = { horizontal: 'right', vertical: 'middle' };
    bannerCell.border = thinBorder;
    currentExcelRow++;

    if (mode === 'detailed') {
      const headers = isBoth
        ? ['التاريخ', 'رقم القيد', 'نوع السند', 'الحساب التحليلي', 'البيان والشرح', 'مدين (ج.س)', 'دائن (ج.س)', 'مدين ($)', 'دائن ($)', 'الرصيد (ج.س)', 'الرصيد ($)']
        : ['التاريخ', 'رقم القيد', 'نوع السند', 'الحساب التحليلي', 'البيان والشرح', `حركة مدين (${isDolar ? '$' : 'ج.س'})`, `حركة دائن (${isDolar ? '$' : 'ج.س'})`, `الرصيد التراكمي (${isDolar ? '$' : 'ج.س'})`];

      const subList = (acc.subAccounts && acc.subAccounts.length > 0) ? acc.subAccounts : [];

      if (subList.length === 0) {
        const noDataRow = worksheet.addRow([]);
        noDataRow.height = 24;
        worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, totalCols);
        const ndCell = worksheet.getCell(currentExcelRow, 1);
        ndCell.value = 'لا توجد حركات أو قيود مسجلة لهذا الحساب خلال الفترة المحددة';
        ndCell.alignment = { horizontal: 'center', vertical: 'middle' };
        ndCell.font = { name: fontFamily, size: 9.5, italic: true, color: { argb: 'FF94A3B8' } };
        for (let i = 1; i <= totalCols; i++) noDataRow.getCell(i).border = thinBorder;
        currentExcelRow++;
      } else {
        subList.forEach((sub: any) => {
          // A. Sub-Account Banner Row
          const subBannerRow = worksheet.addRow([]);
          subBannerRow.height = 24;
          worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, totalCols);
          const subBannerCell = worksheet.getCell(currentExcelRow, 1);

          const subOpenStr = isDolar
            ? (sub.netOpenDolar > 0 ? `رصيد سابق: $${Number(sub.netOpenDolar).toLocaleString()} (مدين)` : sub.netOpenDolar < 0 ? `رصيد سابق: $${Number(Math.abs(sub.netOpenDolar)).toLocaleString()} (دائن)` : 'رصيد سابق: متزن')
            : (sub.netOpen > 0 ? `رصيد سابق: ${Number(sub.netOpen).toLocaleString()} ج.س (مدين)` : sub.netOpen < 0 ? `رصيد سابق: ${Number(Math.abs(sub.netOpen)).toLocaleString()} ج.س (دائن)` : 'رصيد سابق: متزن');

          const subEndingStr = isDolar
            ? (sub.netEndingDolar > 0 ? `صافي ختامي: $${Number(sub.netEndingDolar).toLocaleString()} (مدين)` : sub.netEndingDolar < 0 ? `صافي ختامي: $${Number(Math.abs(sub.netEndingDolar)).toLocaleString()} (دائن)` : 'صافي ختامي: متزن')
            : (sub.netEnding > 0 ? `صافي ختامي: ${Number(sub.netEnding).toLocaleString()} ج.س (مدين)` : sub.netEnding < 0 ? `صافي ختامي: ${Number(Math.abs(sub.netEnding)).toLocaleString()} ج.س (دائن)` : 'صافي ختامي: متزن');

          subBannerCell.value = `◀  [ ${sub.accountNumber} ]  ${sub.accountName}   |   ${subOpenStr}   |   ${subEndingStr}   |   (${sub.transactions?.length || 0} قيد)`;
          subBannerCell.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E40AF' } };
          subBannerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FF' } };
          subBannerCell.alignment = { horizontal: 'right', vertical: 'middle' };
          for (let i = 1; i <= totalCols; i++) subBannerRow.getCell(i).border = thinBorder;
          currentExcelRow++;

          // B. Sub-Account Column Headers
          const shRow = worksheet.addRow([]);
          shRow.height = 20;
          headers.forEach((h, i) => {
            const c = shRow.getCell(i + 1);
            c.value = h;
            c.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
            c.fill = fillPrimaryDark;
            c.alignment = { horizontal: 'center', vertical: 'middle' };
            c.border = thinBorder;
          });
          currentExcelRow++;

          // C. Opening Balance Row for Sub-Account
          const initRow = worksheet.addRow([]);
          initRow.height = 20;
          initRow.getCell(1).value = startDate;
          initRow.getCell(2).value = '-';
          initRow.getCell(3).value = 'رصيد سابق';
          initRow.getCell(4).value = `${sub.accountNumber} - ${sub.accountName}`;
          initRow.getCell(5).value = `الرصيد الافتتاحي السابق المدور لحساب [ ${sub.accountName} ]`;

          for (let i = 1; i <= 5; i++) {
            const c = initRow.getCell(i);
            c.font = { name: fontFamily, size: 8.5, bold: true, color: { argb: 'FF475569' } };
            c.fill = fillSubtotal;
            c.border = thinBorder;
            c.alignment = (i === 4 || i === 5) ? { horizontal: 'right', vertical: 'middle' } : { horizontal: 'center', vertical: 'middle' };
          }

          if (!isBoth) {
            initRow.getCell(6).value = '-';
            initRow.getCell(6).fill = fillSubtotal;
            initRow.getCell(6).border = thinBorder;
            initRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

            initRow.getCell(7).value = '-';
            initRow.getCell(7).fill = fillSubtotal;
            initRow.getCell(7).border = thinBorder;
            initRow.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };

            writeNetCell(initRow, 8, isDolar ? sub.netOpenDolar : sub.netOpen, true, fillSubtotal);
          } else {
            for (let i = 6; i <= 9; i++) {
              const c = initRow.getCell(i);
              c.value = '-';
              c.fill = fillSubtotal;
              c.border = thinBorder;
              c.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            writeNetCell(initRow, 10, sub.netOpen, true, fillSubtotal);
            writeNetCell(initRow, 11, sub.netOpenDolar, true, fillSubtotal);
          }
          currentExcelRow++;

          // D. Transactions lines for Sub-Account
          (sub.transactions || []).forEach((t: any, tIdx: number) => {
            const tRow = worksheet.addRow([]);
            tRow.height = 20;
            const rowBg = tIdx % 2 === 0 ? fillZebraEven : fillZebraOdd;

            const dateStr = t.TheDate ? new Date(t.TheDate).toISOString().split('T')[0] : '';
            tRow.getCell(1).value = dateStr;
            tRow.getCell(2).value = t.ReqNum || t.Id;
            tRow.getCell(3).value = t.DocType || 'قيد';
            tRow.getCell(4).value = `${t.AccountNumber} - ${t.AccountName || ''}`;
            tRow.getCell(5).value = t.Detail || '';

            for (let i = 1; i <= 5; i++) {
              const c = tRow.getCell(i);
              c.font = { name: fontFamily, size: 8.5, color: { argb: 'FF1E293B' } };
              c.fill = rowBg;
              c.border = thinBorder;
              if (i === 1 || i === 2 || i === 3) c.alignment = { horizontal: 'center', vertical: 'middle' };
              else c.alignment = { horizontal: 'right', vertical: 'middle' };
            }

            if (!isBoth) {
              const dr = isDolar ? t.DebtorDolar : t.Debtor;
              const cr = isDolar ? t.CreditorDolar : t.Creditor;
              const run = isDolar ? t.runningNetDolar : t.runningNet;

              writeNumCell(tRow, 6, dr, 'FFE11D48', false, rowBg);
              writeNumCell(tRow, 7, cr, 'FF059669', false, rowBg);
              writeNetCell(tRow, 8, run, true, rowBg);
            } else {
              writeNumCell(tRow, 6, t.Debtor, 'FFE11D48', false, rowBg);
              writeNumCell(tRow, 7, t.Creditor, 'FF059669', false, rowBg);
              writeNumCell(tRow, 8, t.DebtorDolar, 'FF2563EB', false, rowBg);
              writeNumCell(tRow, 9, t.CreditorDolar, 'FF2563EB', false, rowBg);
              writeNetCell(tRow, 10, t.runningNet, true, rowBg);
              writeNetCell(tRow, 11, t.runningNetDolar, true, rowBg);
            }
            currentExcelRow++;
          });

          // E. Sub-Account Subtotal Row
          const subTotRow = worksheet.addRow([]);
          subTotRow.height = 22;
          worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, 5);
          const subTotLabel = worksheet.getCell(currentExcelRow, 1);
          subTotLabel.value = `مجموع حركات وصافي رصيد [ ${sub.accountName} ]:`;
          subTotLabel.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF1E293B' } };
          subTotLabel.alignment = { horizontal: 'right', vertical: 'middle' };
          for (let i = 1; i <= 5; i++) {
            subTotRow.getCell(i).fill = fillSubtotal;
            subTotRow.getCell(i).border = thinBorder;
          }

          if (!isBoth) {
            const pDr = isDolar ? sub.periodDebtorDolar : sub.periodDebtor;
            const pCr = isDolar ? sub.periodCreditorDolar : sub.periodCreditor;
            const endNet = isDolar ? sub.netEndingDolar : sub.netEnding;

            writeNumCell(subTotRow, 6, pDr, 'FFE11D48', true, fillSubtotal);
            writeNumCell(subTotRow, 7, pCr, 'FF059669', true, fillSubtotal);
            writeNetCell(subTotRow, 8, endNet, true, fillSubtotal);
          } else {
            writeNumCell(subTotRow, 6, sub.periodDebtor, 'FFE11D48', true, fillSubtotal);
            writeNumCell(subTotRow, 7, sub.periodCreditor, 'FF059669', true, fillSubtotal);
            writeNumCell(subTotRow, 8, sub.periodDebtorDolar, 'FF2563EB', true, fillSubtotal);
            writeNumCell(subTotRow, 9, sub.periodCreditorDolar, 'FF2563EB', true, fillSubtotal);
            writeNetCell(subTotRow, 10, sub.netEnding, true, fillSubtotal);
            writeNetCell(subTotRow, 11, sub.netEndingDolar, true, fillSubtotal);
          }
          currentExcelRow++;

          // Spacer between sub-accounts
          worksheet.addRow([]);
          worksheet.getRow(currentExcelRow).height = 6;
          currentExcelRow++;
        });
      }
    } else {
      // Summary Mode
      const hRow = worksheet.addRow([]);
      hRow.height = 22;
      const headers = isBoth
        ? ['كود الحساب', 'اسم الحساب الفرعي', 'الرصيد الافتتاحي (ج.س)', 'الرصيد الافتتاحي ($)', 'مدين (ج.س)', 'دائن (ج.س)', 'مدين ($)', 'دائن ($)', 'الرصيد الختامي (ج.س)', 'الرصيد الختامي ($)']
        : ['كود الحساب الفرعي', 'اسم الحساب المالي', `الرصيد الافتتاحي (${isDolar ? '$' : 'ج.س'})`, `حركة مدين (${isDolar ? '$' : 'ج.س'})`, `حركة دائن (${isDolar ? '$' : 'ج.س'})`, `الرصيد الختامي (${isDolar ? '$' : 'ج.س'})`];

      headers.forEach((h, i) => {
        const c = hRow.getCell(i + 1);
        c.value = h;
        c.font = fontHeader;
        c.fill = fillPrimaryDark;
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });
      currentExcelRow++;

      (acc.childAccounts || []).forEach((cItem: any, cIdx: number) => {
        const cRow = worksheet.addRow([]);
        cRow.height = 20;
        const rowBg = cIdx % 2 === 0 ? fillZebraEven : fillZebraOdd;

        cRow.getCell(1).value = cItem.AccountNumber;
        cRow.getCell(1).font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF2563EB' } };
        cRow.getCell(1).fill = rowBg;
        cRow.getCell(1).border = thinBorder;
        cRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

        cRow.getCell(2).value = cItem.AccountName;
        cRow.getCell(2).font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF1E293B' } };
        cRow.getCell(2).fill = rowBg;
        cRow.getCell(2).border = thinBorder;
        cRow.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };

        if (!isBoth) {
          const openVal = isDolar ? cItem.netOpenDolar : cItem.netOpen;
          const dr = isDolar ? cItem.periodDebtorDolar : cItem.periodDebtor;
          const cr = isDolar ? cItem.periodCreditorDolar : cItem.periodCreditor;
          const endVal = isDolar ? cItem.netEndingDolar : cItem.netEnding;

          writeNetCell(cRow, 3, openVal, false, rowBg);
          writeNumCell(cRow, 4, dr, 'FFE11D48', false, rowBg);
          writeNumCell(cRow, 5, cr, 'FF059669', false, rowBg);
          writeNetCell(cRow, 6, endVal, true, rowBg);
        } else {
          writeNetCell(cRow, 3, cItem.netOpen, false, rowBg);
          writeNetCell(cRow, 4, cItem.netOpenDolar, false, rowBg);
          writeNumCell(cRow, 5, cItem.periodDebtor, 'FFE11D48', false, rowBg);
          writeNumCell(cRow, 6, cItem.periodCreditor, 'FF059669', false, rowBg);
          writeNumCell(cRow, 7, cItem.periodDebtorDolar, 'FF2563EB', false, rowBg);
          writeNumCell(cRow, 8, cItem.periodCreditorDolar, 'FF2563EB', false, rowBg);
          writeNetCell(cRow, 9, cItem.netEnding, true, rowBg);
          writeNetCell(cRow, 10, cItem.netEndingDolar, true, rowBg);
        }
        currentExcelRow++;
      });
    }

    // 3. Subtotal Row for this Account
    const subtotalRow = worksheet.addRow([]);
    subtotalRow.height = 24;

    const labelColCount = mode === 'detailed' ? 5 : 2;
    worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, labelColCount);
    const sLabel = worksheet.getCell(currentExcelRow, 1);
    sLabel.value = mode === 'detailed' 
      ? `إجمالي الحساب العام بالكامل [ ${acc.code} - ${acc.name} ] (${(acc.subAccounts || []).length} حساب تحليلي):` 
      : `إجمالي حركات وصافي رصيد حساب [ ${acc.code} - ${acc.name} ]`;
    sLabel.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E3A8A' } };
    sLabel.fill = mode === 'detailed' ? fillGroupHeader : fillSubtotal;
    sLabel.border = thinBorder;
    sLabel.alignment = { horizontal: 'right', vertical: 'middle' };

    if (mode === 'detailed') {
      if (!isBoth) {
        writeNumCell(subtotalRow, 6, isDolar ? acc.periodDebtorDolar : acc.periodDebtor, 'FFE11D48', true, fillSubtotal);
        writeNumCell(subtotalRow, 7, isDolar ? acc.periodCreditorDolar : acc.periodCreditor, 'FF059669', true, fillSubtotal);
        writeNetCell(subtotalRow, 8, isDolar ? acc.netEndingDolar : acc.netEnding, true, fillSubtotal);
      } else {
        writeNumCell(subtotalRow, 6, acc.periodDebtor, 'FFE11D48', true, fillSubtotal);
        writeNumCell(subtotalRow, 7, acc.periodCreditor, 'FF059669', true, fillSubtotal);
        writeNumCell(subtotalRow, 8, acc.periodDebtorDolar, 'FF2563EB', true, fillSubtotal);
        writeNumCell(subtotalRow, 9, acc.periodCreditorDolar, 'FF2563EB', true, fillSubtotal);
        writeNetCell(subtotalRow, 10, acc.netEnding, true, fillSubtotal);
        writeNetCell(subtotalRow, 11, acc.netEndingDolar, true, fillSubtotal);
      }
    } else {
      if (!isBoth) {
        writeNetCell(subtotalRow, 3, isDolar ? acc.netOpenDolar : acc.netOpen, true, fillSubtotal);
        writeNumCell(subtotalRow, 4, isDolar ? acc.periodDebtorDolar : acc.periodDebtor, 'FFE11D48', true, fillSubtotal);
        writeNumCell(subtotalRow, 5, isDolar ? acc.periodCreditorDolar : acc.periodCreditor, 'FF059669', true, fillSubtotal);
        writeNetCell(subtotalRow, 6, isDolar ? acc.netEndingDolar : acc.netEnding, true, fillSubtotal);
      } else {
        writeNetCell(subtotalRow, 3, acc.netOpen, true, fillSubtotal);
        writeNetCell(subtotalRow, 4, acc.netOpenDolar, true, fillSubtotal);
        writeNumCell(subtotalRow, 5, acc.periodDebtor, 'FFE11D48', true, fillSubtotal);
        writeNumCell(subtotalRow, 6, acc.periodCreditor, 'FF059669', true, fillSubtotal);
        writeNumCell(subtotalRow, 7, acc.periodDebtorDolar, 'FF2563EB', true, fillSubtotal);
        writeNumCell(subtotalRow, 8, acc.periodCreditorDolar, 'FF2563EB', true, fillSubtotal);
        writeNetCell(subtotalRow, 9, acc.netEnding, true, fillSubtotal);
        writeNetCell(subtotalRow, 10, acc.netEndingDolar, true, fillSubtotal);
      }
    }
    currentExcelRow++;

    // Spacing
    worksheet.addRow([]);
    currentExcelRow++;
  });

  // GRAND TOTAL SUMMARY ROW
  const grandRow = worksheet.addRow([]);
  grandRow.height = 30;

  const gLabelCols = mode === 'detailed' ? 5 : 2;
  worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, gLabelCols);
  const grandTitle = worksheet.getCell(currentExcelRow, 1);
  grandTitle.value = 'المجموع العام الإجمالي لكافة الحسابات العامة المختارة';
  grandTitle.font = fontGrandTotal;
  grandTitle.fill = fillGrandTotal;
  grandTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  grandTitle.border = doubleBottomBorder;

  if (mode === 'detailed') {
    if (!isBoth) {
      writeNumCell(grandRow, 6, isDolar ? grandTotals.periodDebtorDolar : grandTotals.periodDebtor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 7, isDolar ? grandTotals.periodCreditorDolar : grandTotals.periodCreditor, 'FFFFFFFF', true, fillGrandTotal);
      writeNetCell(grandRow, 8, isDolar ? (grandTotals.netEndingDebitDolar - grandTotals.netEndingCreditDolar) : (grandTotals.netEndingDebit - grandTotals.netEndingCredit), true, fillGrandTotal);
    } else {
      writeNumCell(grandRow, 6, grandTotals.periodDebtor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 7, grandTotals.periodCreditor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 8, grandTotals.periodDebtorDolar, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 9, grandTotals.periodCreditorDolar, 'FFFFFFFF', true, fillGrandTotal);
      writeNetCell(grandRow, 10, grandTotals.netEndingDebit - grandTotals.netEndingCredit, true, fillGrandTotal);
      writeNetCell(grandRow, 11, grandTotals.netEndingDebitDolar - grandTotals.netEndingCreditDolar, true, fillGrandTotal);
    }
  } else {
    if (!isBoth) {
      writeNetCell(grandRow, 3, isDolar ? (grandTotals.netOpenDebitDolar - grandTotals.netOpenCreditDolar) : (grandTotals.netOpenDebit - grandTotals.netOpenCredit), true, fillGrandTotal);
      writeNumCell(grandRow, 4, isDolar ? grandTotals.periodDebtorDolar : grandTotals.periodDebtor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 5, isDolar ? grandTotals.periodCreditorDolar : grandTotals.periodCreditor, 'FFFFFFFF', true, fillGrandTotal);
      writeNetCell(grandRow, 6, isDolar ? (grandTotals.netEndingDebitDolar - grandTotals.netEndingCreditDolar) : (grandTotals.netEndingDebit - grandTotals.netEndingCredit), true, fillGrandTotal);
    } else {
      writeNetCell(grandRow, 3, grandTotals.netOpenDebit - grandTotals.netOpenCredit, true, fillGrandTotal);
      writeNetCell(grandRow, 4, grandTotals.netOpenDebitDolar - grandTotals.netOpenCreditDolar, true, fillGrandTotal);
      writeNumCell(grandRow, 5, grandTotals.periodDebtor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 6, grandTotals.periodCreditor, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 7, grandTotals.periodDebtorDolar, 'FFFFFFFF', true, fillGrandTotal);
      writeNumCell(grandRow, 8, grandTotals.periodCreditorDolar, 'FFFFFFFF', true, fillGrandTotal);
      writeNetCell(grandRow, 9, grandTotals.netEndingDebit - grandTotals.netEndingCredit, true, fillGrandTotal);
      writeNetCell(grandRow, 10, grandTotals.netEndingDebitDolar - grandTotals.netEndingCreditDolar, true, fillGrandTotal);
    }
  }

  for (let c = gLabelCols + 1; c <= totalCols; c++) {
    grandRow.getCell(c).border = doubleBottomBorder;
  }
  currentExcelRow++;

  // SIGNATURES BLOCK
  worksheet.addRow([]);
  currentExcelRow++;
  const signRow1 = worksheet.addRow([]);
  signRow1.height = 26;
  const signRow2 = worksheet.addRow([]);
  signRow2.height = 34;

  const signCols = [
    { title: 'إعداد / المحاسب المسؤول', col: 1 },
    { title: 'مراجعة / المراجع المالي', col: Math.floor(totalCols / 2) },
    { title: 'اعتماد / المدير المالي (CFO)', col: totalCols }
  ];

  signCols.forEach((s) => {
    if (s.col <= totalCols) {
      const cell1 = signRow1.getCell(s.col);
      cell1.value = s.title;
      cell1.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF334155' } };
      cell1.alignment = { horizontal: 'center', vertical: 'middle' };

      const cell2 = signRow2.getCell(s.col);
      cell2.value = 'التوقيع: ........................................';
      cell2.font = { name: fontFamily, size: 9, color: { argb: 'FF94A3B8' } };
      cell2.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  });

  // Adjust widths
  if (mode === 'detailed') {
    worksheet.getColumn(1).width = 15; // التاريخ
    worksheet.getColumn(2).width = 15; // رقم القيد
    worksheet.getColumn(3).width = 14; // نوع السند
    worksheet.getColumn(4).width = 28; // الحساب التحليلي
    worksheet.getColumn(5).width = 40; // البيان والشرح
    for (let c = 6; c <= totalCols; c++) {
      worksheet.getColumn(c).width = 22; // المبالغ والرصيد التراكمي
    }
  } else {
    worksheet.getColumn(1).width = 20; // كود الحساب
    worksheet.getColumn(2).width = 38; // اسم الحساب
    for (let c = 3; c <= totalCols; c++) {
      worksheet.getColumn(c).width = 22;
    }
  }

  // Trigger download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `كشف_الحسابات_العامة_${mode}_${startDate}_${endDate}.xlsx`;
  a.click();
  window.URL.revokeObjectURL(url);
}

export interface SubAccountLedgerExportParams {
  account: {
    code: string;
    name: string;
    parentCode?: string;
    parentName?: string;
    category?: string;
  };
  openDebtor: number;
  openCreditor: number;
  openDebtorDolar: number;
  openCreditorDolar: number;
  netOpen: number;
  netOpenDolar: number;
  periodDebtor: number;
  periodCreditor: number;
  periodDebtorDolar: number;
  periodCreditorDolar: number;
  totalDebtor: number;
  totalCreditor: number;
  netEnding: number;
  netEndingDolar: number;
  transactions: any[];
  startDate: string;
  endDate: string;
  currency: 'EGP' | 'USD' | 'BOTH';
}

export async function exportStyledSubAccountLedgerExcel(params: SubAccountLedgerExportParams) {
  const {
    account,
    netOpen,
    netOpenDolar,
    periodDebtor,
    periodCreditor,
    periodDebtorDolar,
    periodCreditorDolar,
    netEnding,
    netEndingDolar,
    transactions,
    startDate,
    endDate,
    currency
  } = params;

  if (!account) return;

  const isDolar = currency === 'USD';
  const isBoth = currency === 'BOTH';

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'الرويال للتطوير العقاري - ERP System';
  workbook.lastModifiedBy = 'نظام كشف الحسابات الفرعية';
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheetName = (account.name || 'كشف حساب').slice(0, 30).replace(/[:\\\/\?\*\[\]]/g, '');
  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ rightToLeft: true, showGridLines: true }]
  });

  const totalCols = isBoth ? 10 : 7;
  const fontFamily = 'Segoe UI';

  // Fonts
  const fontTitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontSubtitle: Partial<ExcelJS.Font> = { name: fontFamily, size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontMeta: Partial<ExcelJS.Font> = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
  const fontHeader: Partial<ExcelJS.Font> = { name: fontFamily, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontGrandTotal: Partial<ExcelJS.Font> = { name: fontFamily, size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };

  // Fills
  const fillPrimaryDark: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  const fillPrimaryMedium: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
  const fillMetaBar: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
  const fillGroupHeader: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
  const fillSubtotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  const fillGrandTotal: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  const fillZebraEven: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  const fillZebraOdd: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  const fillDebitTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F2' } };
  const fillCreditTint: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  const doubleBottomBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF0F172A' } },
    bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF0F172A' } },
    right: { style: 'thin', color: { argb: 'FF0F172A' } }
  };

  // Row 1: Spacing
  worksheet.addRow([]);
  worksheet.getRow(1).height = 10;

  // Row 2: Company Header Banner
  const titleRow = worksheet.addRow([]);
  titleRow.height = 36;
  worksheet.mergeCells(2, 1, 2, totalCols);
  const titleCell = worksheet.getCell(2, 1);
  titleCell.value = 'الرويال للتطوير العقاري - الإدارة المالية والحسابات العامة';
  titleCell.font = fontTitle;
  titleCell.fill = fillPrimaryDark;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 3: Subtitle
  const subRow = worksheet.addRow([]);
  subRow.height = 28;
  worksheet.mergeCells(3, 1, 3, totalCols);
  const subCell = worksheet.getCell(3, 1);
  subCell.value = `كشف حساب تفصيلي: [ ${account.code} ]  ${account.name}   ${account.parentName ? `(الحساب العام: ${account.parentName})` : ''}`;
  subCell.font = fontSubtitle;
  subCell.fill = fillPrimaryMedium;
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 4: Metadata Bar
  const metaRow = worksheet.addRow([]);
  metaRow.height = 24;
  worksheet.mergeCells(4, 1, 4, totalCols);
  const metaCell = worksheet.getCell(4, 1);
  const currLabel = currency === 'EGP' ? 'الجنيه السوداني (SDG)' : currency === 'USD' ? 'الدولار الأمريكي ($ USD)' : 'الجنيه والدولار معاً';
  metaCell.value = `الفترة المالية: من ${startDate} إلى ${endDate}   |   العملة: ${currLabel}   |   التصنيف: ${account.category || 'الأصول'}   |   عدد الحركات: ${transactions.length} قيد   |   الرصيد: أحمر (مدين) / أخضر (دائن)`;
  metaCell.font = fontMeta;
  metaCell.fill = fillMetaBar;
  metaCell.alignment = { horizontal: 'center', vertical: 'middle' };
  metaCell.border = thinBorder;

  // Row 5: Spacing
  worksheet.addRow([]);
  worksheet.getRow(5).height = 10;

  let currentExcelRow = 6;

  // Helpers for cells
  const writeNetCell = (rowObj: ExcelJS.Row, colNumber: number, val: number, isBold = true, customFill?: ExcelJS.Fill) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;

    if (Math.abs(val) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
      if (customFill) cell.fill = customFill;
      return;
    }

    if (val > 0) {
      cell.value = Number(val);
      cell.numFmt = '#,##0.00 " (مدين)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FFE11D48' } };
      cell.fill = customFill || fillDebitTint;
    } else {
      cell.value = Number(Math.abs(val));
      cell.numFmt = '#,##0.00 " (دائن)"';
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: 'FF059669' } };
      cell.fill = customFill || fillCreditTint;
    }
  };

  const writeNumCell = (rowObj: ExcelJS.Row, colNumber: number, val: number, colorARGB: string, isBold = false, customFill?: ExcelJS.Fill) => {
    const cell = rowObj.getCell(colNumber);
    cell.border = thinBorder;
    cell.numFmt = '#,##0.00';
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.font = { name: fontFamily, size: 9.5, bold: isBold, color: { argb: colorARGB } };
    if (customFill) cell.fill = customFill;

    if (Math.abs(val) < 0.001) {
      cell.value = '-';
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { name: fontFamily, size: 9.5, color: { argb: 'FF94A3B8' } };
    } else {
      cell.value = Number(val);
    }
  };

  // Table Headers
  const headers = isBoth
    ? ['التاريخ', 'رقم القيد / السند', 'نوع السند', 'البيان والشرح والتفاصيل', 'مدين (ج.س)', 'دائن (ج.س)', 'مدين ($)', 'دائن ($)', 'الرصيد التراكمي (ج.س)', 'الرصيد التراكمي ($)']
    : ['التاريخ', 'رقم القيد / السند', 'نوع السند', 'البيان والشرح والتفاصيل', `حركة مدين (${isDolar ? '$' : 'ج.س'})`, `حركة دائن (${isDolar ? '$' : 'ج.س'})`, `الرصيد التراكمي (${isDolar ? '$' : 'ج.س'})`];

  const hRow = worksheet.addRow([]);
  hRow.height = 22;
  headers.forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value = h;
    c.font = fontHeader;
    c.fill = fillPrimaryDark;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = thinBorder;
  });
  currentExcelRow++;

  // Opening Balance Row
  const initRow = worksheet.addRow([]);
  initRow.height = 22;
  initRow.getCell(1).value = startDate;
  initRow.getCell(2).value = '-';
  initRow.getCell(3).value = 'رصيد سابق';
  initRow.getCell(4).value = `الرصيد الافتتاحي السابق المدور لحساب [ ${account.name} ] عن فترة ما قبل ${startDate}`;

  for (let i = 1; i <= 4; i++) {
    const c = initRow.getCell(i);
    c.font = { name: fontFamily, size: 9, bold: true, color: { argb: 'FF475569' } };
    c.fill = fillSubtotal;
    c.border = thinBorder;
    c.alignment = i === 4 ? { horizontal: 'right', vertical: 'middle' } : { horizontal: 'center', vertical: 'middle' };
  }

  if (!isBoth) {
    initRow.getCell(5).value = '-';
    initRow.getCell(5).fill = fillSubtotal;
    initRow.getCell(5).border = thinBorder;
    initRow.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    initRow.getCell(6).value = '-';
    initRow.getCell(6).fill = fillSubtotal;
    initRow.getCell(6).border = thinBorder;
    initRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

    writeNetCell(initRow, 7, isDolar ? netOpenDolar : netOpen, true, fillSubtotal);
  } else {
    for (let i = 5; i <= 8; i++) {
      const c = initRow.getCell(i);
      c.value = '-';
      c.fill = fillSubtotal;
      c.border = thinBorder;
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    writeNetCell(initRow, 9, netOpen, true, fillSubtotal);
    writeNetCell(initRow, 10, netOpenDolar, true, fillSubtotal);
  }
  currentExcelRow++;

  // Transactions Rows
  if (!transactions || transactions.length === 0) {
    const noDataRow = worksheet.addRow([]);
    noDataRow.height = 24;
    worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, totalCols);
    const ndCell = worksheet.getCell(currentExcelRow, 1);
    ndCell.value = 'لا توجد حركات مسجلة لهذا الحساب خلال الفترة المحددة';
    ndCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ndCell.font = { name: fontFamily, size: 9.5, italic: true, color: { argb: 'FF94A3B8' } };
    for (let i = 1; i <= totalCols; i++) noDataRow.getCell(i).border = thinBorder;
    currentExcelRow++;
  } else {
    transactions.forEach((t: any, tIdx: number) => {
      const tRow = worksheet.addRow([]);
      tRow.height = 20;
      const rowBg = tIdx % 2 === 0 ? fillZebraEven : fillZebraOdd;

      const dateStr = t.TheDate ? new Date(t.TheDate).toISOString().split('T')[0] : '';
      tRow.getCell(1).value = dateStr;
      tRow.getCell(2).value = t.ReqNum || t.Id;
      tRow.getCell(3).value = t.DocType || 'قيد';
      tRow.getCell(4).value = t.Detail || '-';

      for (let i = 1; i <= 4; i++) {
        const c = tRow.getCell(i);
        c.font = { name: fontFamily, size: 9, color: { argb: 'FF1E293B' } };
        c.fill = rowBg;
        c.border = thinBorder;
        if (i === 1 || i === 2 || i === 3) c.alignment = { horizontal: 'center', vertical: 'middle' };
        else c.alignment = { horizontal: 'right', vertical: 'middle' };
      }

      if (!isBoth) {
        const dr = isDolar ? t.DebtorDolar : t.Debtor;
        const cr = isDolar ? t.CreditorDolar : t.Creditor;
        const run = isDolar ? t.runningNetDolar : t.runningNet;

        writeNumCell(tRow, 5, dr, 'FFE11D48', false, rowBg);
        writeNumCell(tRow, 6, cr, 'FF059669', false, rowBg);
        writeNetCell(tRow, 7, run, true, rowBg);
      } else {
        writeNumCell(tRow, 5, t.Debtor, 'FFE11D48', false, rowBg);
        writeNumCell(tRow, 6, t.Creditor, 'FF059669', false, rowBg);
        writeNumCell(tRow, 7, t.DebtorDolar, 'FF2563EB', false, rowBg);
        writeNumCell(tRow, 8, t.CreditorDolar, 'FF2563EB', false, rowBg);
        writeNetCell(tRow, 9, t.runningNet, true, rowBg);
        writeNetCell(tRow, 10, t.runningNetDolar, true, rowBg);
      }
      currentExcelRow++;
    });
  }

  // Summary / Subtotal Row
  const totRow = worksheet.addRow([]);
  totRow.height = 24;
  worksheet.mergeCells(currentExcelRow, 1, currentExcelRow, 4);
  const totLabel = worksheet.getCell(currentExcelRow, 1);
  totLabel.value = `إجمالي حركات وصافي رصيد حساب [ ${account.name} ]:`;
  totLabel.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF0F172A' } };
  totLabel.alignment = { horizontal: 'right', vertical: 'middle' };
  for (let i = 1; i <= 4; i++) {
    totRow.getCell(i).fill = fillGroupHeader;
    totRow.getCell(i).border = thinBorder;
  }

  if (!isBoth) {
    const pDr = isDolar ? periodDebtorDolar : periodDebtor;
    const pCr = isDolar ? periodCreditorDolar : periodCreditor;
    const endNet = isDolar ? netEndingDolar : netEnding;

    writeNumCell(totRow, 5, pDr, 'FFE11D48', true, fillGroupHeader);
    writeNumCell(totRow, 6, pCr, 'FF059669', true, fillGroupHeader);
    writeNetCell(totRow, 7, endNet, true, fillGroupHeader);
  } else {
    writeNumCell(totRow, 5, periodDebtor, 'FFE11D48', true, fillGroupHeader);
    writeNumCell(totRow, 6, periodCreditor, 'FF059669', true, fillGroupHeader);
    writeNumCell(totRow, 7, periodDebtorDolar, 'FF2563EB', true, fillGroupHeader);
    writeNumCell(totRow, 8, periodCreditorDolar, 'FF2563EB', true, fillGroupHeader);
    writeNetCell(totRow, 9, netEnding, true, fillGroupHeader);
    writeNetCell(totRow, 10, netEndingDolar, true, fillGroupHeader);
  }
  currentExcelRow++;

  // Spacing before signatures
  worksheet.addRow([]);
  worksheet.getRow(currentExcelRow).height = 14;
  currentExcelRow++;

  // Signatures Row
  const signRow1 = worksheet.addRow([]);
  signRow1.height = 20;
  const signRow2 = worksheet.addRow([]);
  signRow2.height = 22;

  const signCols = [
    { title: 'إعداد / المحاسب المسؤول', col: 2 },
    { title: 'مراجعة / المراجع المالي', col: Math.floor(totalCols / 2) },
    { title: 'اعتماد / المدير المالي (CFO)', col: totalCols }
  ];

  signCols.forEach((s) => {
    if (s.col <= totalCols) {
      const cell1 = signRow1.getCell(s.col);
      cell1.value = s.title;
      cell1.font = { name: fontFamily, size: 9.5, bold: true, color: { argb: 'FF334155' } };
      cell1.alignment = { horizontal: 'center', vertical: 'middle' };

      const cell2 = signRow2.getCell(s.col);
      cell2.value = 'التوقيع: ........................................';
      cell2.font = { name: fontFamily, size: 8.5, color: { argb: 'FF94A3B8' } };
      cell2.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  });

  // Adjust Column Widths
  worksheet.getColumn(1).width = 14; // التاريخ
  worksheet.getColumn(2).width = 16; // رقم القيد
  worksheet.getColumn(3).width = 14; // نوع السند
  worksheet.getColumn(4).width = 44; // البيان والشرح
  for (let c = 5; c <= totalCols; c++) {
    worksheet.getColumn(c).width = 20; // المبالغ والرصيد
  }

  // Trigger Download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const safeName = (account.name || account.code).replace(/[\s\/\\?%*:|"<>]/g, '_');
  a.download = `كشف_حساب_${safeName}_${startDate}_${endDate}.xlsx`;
  a.click();
  window.URL.revokeObjectURL(url);
}
