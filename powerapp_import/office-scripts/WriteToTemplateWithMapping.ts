function main(
  workbook: ExcelScript.Workbook,
  resultJson: string
): void {
  const results: Array<{
    氏名: string;
    日付: string;
    出勤時刻_修正: string;
    退勤時刻_修正: string;
    勤務時間: number;
    残業時間: number;
    作業番号: string;
    案件: string;
    自由入力１: string;
    自由入力２: string;
    単価: number;
  }> = JSON.parse(resultJson);

  const columnDefinitions: Array<{ header: string; field: string }> = [
    { header: "日付", field: "日付" },
    { header: "出勤時刻", field: "出勤時刻_修正" },
    { header: "退勤時刻", field: "退勤時刻_修正" },
    { header: "勤務時間", field: "勤務時間" },
    { header: "残業時間", field: "残業時間" },
    { header: "作業番号", field: "作業番号" },
    { header: "案件", field: "案件" },
    { header: "自由入力１", field: "自由入力１" },
    { header: "自由入力２", field: "自由入力２" },
    { header: "単価", field: "単価" }
  ];

  const grouped: { [name: string]: typeof results } = {};
  for (const r of results) {
    if (!grouped[r.氏名]) {
      grouped[r.氏名] = [];
    }
    grouped[r.氏名].push(r);
  }

  for (const [name, records] of Object.entries(grouped)) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) {
      console.log(`シート "${name}" が見つかりません。スキップします。`);
      continue;
    }

    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
      console.log(`シート "${name}" にデータがありません。`);
      continue;
    }

    const rowCount = usedRange.getRowCount();
    const colCount = usedRange.getColumnCount();
    let headerRowIndex = -1;
    const columnMap: { [field: string]: number } = {};

    for (let row = 0; row < rowCount; row++) {
      const rowValues = usedRange.getRow(row).getValues()[0];
      let matchCount = 0;
      const tempMap: { [field: string]: number } = {};

      for (let col = 0; col < colCount; col++) {
        const cellValue = String(rowValues[col] || "").trim();
        for (const def of columnDefinitions) {
          if (cellValue === def.header) {
            tempMap[def.field] = col;
            matchCount++;
            break;
          }
        }
      }

      if (matchCount >= columnDefinitions.length * 0.5) {
        headerRowIndex = row;
        for (const [field, colIdx] of Object.entries(tempMap)) {
          columnMap[field] = colIdx;
        }
        break;
      }
    }

    if (headerRowIndex === -1) {
      console.log(`シート "${name}" で見出し行が見つかりませんでした。`);
      continue;
    }

    const dataStartRow = headerRowIndex + 1;
    records.sort((a, b) => a.日付.localeCompare(b.日付));

    const dataRows: (string | number)[][] = [];
    for (const r of records) {
      const row: (string | number)[] = new Array(colCount).fill("");
      for (const def of columnDefinitions) {
        const colIdx = columnMap[def.field];
        if (colIdx !== undefined) {
          row[colIdx] = (r as any)[def.field] !== undefined ? (r as any)[def.field] : "";
        }
      }
      dataRows.push(row);
    }

    if (dataRows.length > 0) {
      const targetRange = sheet.getRangeByIndexes(
        dataStartRow,
        0,
        dataRows.length,
        colCount
      );
      targetRange.setValues(dataRows);
    }
  }
}
