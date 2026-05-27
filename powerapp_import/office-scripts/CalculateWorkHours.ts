function main(
  workbook: ExcelScript.Workbook,
  attendanceJson: string,
  lunchBreakMinutes: number = 60,
  baseWorkHours: number = 8,
  overtimeThreshold: number = 8.5,
  priceTableJson: string = "[]"
): string {
  const records: Array<{
    氏名: string;
    日付: string;
    出勤時刻: string;
    退勤時刻: string;
    作業番号: string;
  }> = JSON.parse(attendanceJson);

  const priceRaw: Array<{
    氏名: string;
    単価: number;
    退社日: string;
  }> = JSON.parse(priceTableJson);

  const priceMapTemp: { [name: string]: { 単価: number; 退社日: string } } = {};
  for (const p of priceRaw) {
    const existing = priceMapTemp[p.氏名];
    if (!existing) {
      priceMapTemp[p.氏名] = { 単価: p.単価, 退社日: p.退社日 || "" };
    } else if (!p.退社日 || p.退社日 === "") {
      priceMapTemp[p.氏名] = { 単価: p.単価, 退社日: "" };
    }
  }

  const priceMap: { [name: string]: number } = {};
  for (const [name, info] of Object.entries(priceMapTemp)) {
    priceMap[name] = info.単価;
  }

  const results: Array<{
    氏名: string;
    日付: string;
    出勤時刻: string;
    退勤時刻: string;
    出勤時刻_修正: string;
    退勤時刻_修正: string;
    勤務時間: number;
    残業時間: number;
    作業番号: string;
    単価: number;
  }> = [];

  for (const record of records) {
    const [inH, inM] = record.出勤時刻.split(':').map(Number);
    const [outH, outM] = record.退勤時刻.split(':').map(Number);

    const inRounded = inH + Math.floor(inM / 30) * 0.5;
    const outRounded = outH + Math.floor(outM / 30) * 0.5;

    const inH2 = Math.floor(inRounded);
    const inM2 = (inRounded - inH2) * 60;
    const outH2 = Math.floor(outRounded);
    const outM2 = (outRounded - outH2) * 60;

    const 出勤時刻_修正 = `${String(inH2).padStart(2, '0')}:${String(inM2).padStart(2, '0')}`;
    const 退勤時刻_修正 = `${String(outH2).padStart(2, '0')}:${String(outM2).padStart(2, '0')}`;

    const workMinutes = (outRounded - inRounded) * 60 - lunchBreakMinutes;
    const 勤務時間 = Math.max(0, workMinutes / 60);

    let 残業時間 = 0;
    if (勤務時間 > overtimeThreshold) {
      const overtime = 勤務時間 - overtimeThreshold;
      残業時間 = Math.ceil(overtime * 2) / 2;
    }

    // baseWorkHours〜overtimeThreshold の範囲は残業0（業務ルール）
    if (勤務時間 <= baseWorkHours) {
      残業時間 = 0;
    }

    const 単価 = priceMap[record.氏名] || 0;

    results.push({
      氏名: record.氏名,
      日付: record.日付,
      出勤時刻: record.出勤時刻,
      退勤時刻: record.退勤時刻,
      出勤時刻_修正,
      退勤時刻_修正,
      勤務時間: Math.round(勤務時間 * 10) / 10,
      残業時間: Math.round(残業時間 * 10) / 10,
      作業番号: record.作業番号,
      単価: 単価
    });
  }

  return JSON.stringify(results);
}
