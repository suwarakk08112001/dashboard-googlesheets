import { Injectable, OnModuleInit, Logger } from '@nestjs/common';
import { SearchDto } from './dto/search-dashboard.dto';
import { ConfigService } from '@nestjs/config';
import { google, sheets_v4 } from 'googleapis';

@Injectable()
export class DashboardService implements OnModuleInit {
  private sheets: sheets_v4.Sheets;
  private spreadsheetId: string;
  private sheetName: string;
  private sheetName1: string;
  private sheetName2: string;

  private readonly logger = new Logger(DashboardService.name);

  constructor(private configService: ConfigService) {}

  async onModuleInit() {
    const auth = new google.auth.GoogleAuth({
      keyFile: this.configService.get<string>(
        'GOOGLE_SERVICE_ACCOUNT_KEY_PATH',
      ),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const authClient = await auth.getClient();
    this.sheets = google.sheets({
      version: 'v4',
      auth: authClient as InstanceType<typeof google.auth.JWT>,
    });
    this.spreadsheetId = this.configService.get<string>(
      'GOOGLE_SPREADSHEET_ID',
    )!;

    // ดึงชื่อ Sheet จริงจาก Google Sheet
    const spreadsheet = await this.sheets.spreadsheets.get({
      spreadsheetId: this.spreadsheetId,
    });
    const allSheets =
      spreadsheet.data.sheets?.map((s) => s.properties?.title || '') || [];
    this.logger.log(`📋 Sheet ทั้งหมด: ${JSON.stringify(allSheets)}`);

    this.sheetName = allSheets[1] || 'Sheet1';
    this.sheetName1 = allSheets[2] || 'Sheet2';
    this.sheetName2 = allSheets[3] || 'Sheet3';
    // this.sheetName1 = allSheets.find((n) => n.includes('รับเข้า')) || '';
    // this.sheetName2 = allSheets.find((n) => n.includes('จ่ายออก')) || '';

    this.logger.log(`✅ เชื่อมต่อสำเร็จ`);
    this.logger.log(`📌 Sheet หลัก: "${this.sheetName}"`);
    this.logger.log(`📌 Sheet รับเข้า: "${this.sheetName1}"`);
    this.logger.log(`📌 Sheet จ่ายออก: "${this.sheetName2}"`);
  }

  async findTotalSKU(): Promise<{ total_of_SKU: number }> {
    const res = await this.sheets.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range: this.sheetName,
    });

    const rows = res.data.values;
    if (!rows || rows.length <= 1) return { total_of_SKU: 0 };

    return { total_of_SKU: rows.length - 1 };
  }

  async findTotalStockValue(): Promise<{
    total_stock_in: number;
    total_stock_out: number;
    total_stock_value: number;
  }> {
    const resIn = await this.sheets.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range: this.sheetName1,
    });

    const resOut = await this.sheets.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range: this.sheetName2,
    });

    const sumColumn = (rows: any[][] | null | undefined): number => {
      if (!rows || rows.length <= 1) return 0;
      const headers = rows[0];
      const colIndex = headers.findIndex((h: string) =>
        h.includes('มูลค่ารวม'),
      );
      if (colIndex === -1) return 0;

      return rows
        .slice(1)
        .filter((row) => {
          // ตัดแถวที่เป็นแถวรวม/สรุป ออก
          const firstCell = String(row[0] || '').trim();
          return firstCell !== '' && !firstCell.includes('รวม');
        })
        .reduce((sum, row) => {
          const val = parseFloat(
            String(row[colIndex] || '0').replace(/,/g, ''),
          );
          return sum + (isNaN(val) ? 0 : val);
        }, 0);
    };

    const totalIn = sumColumn(resIn.data.values);
    const totalOut = sumColumn(resOut.data.values);

    return {
      total_stock_in: totalIn,
      total_stock_out: totalOut,
      total_stock_value: totalIn - totalOut,
    };
  }

  async findMonthlyStockValue(dto: SearchDto): Promise<{
    fiscalYear: number;
    months: {
      month: number;
      totalIn: number;
      totalOut: number;
      medSupply: number;
    }[];
  }> {
    const [resIn, resOut] = await Promise.all([
      this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: this.sheetName1,
      }),
      this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: this.sheetName2,
      }),
    ]);

    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentFiscalYear =
      currentMonth >= 10 ? now.getFullYear() + 1 : now.getFullYear();

    const fiscalYear = Number(dto.financialYear) || currentFiscalYear;

    const startYear = fiscalYear - 1;
    const isInFiscalYear = (rowYear: number, rowMonth: number): boolean => {
      if (rowYear === startYear && rowMonth >= 10) return true;
      if (rowYear === fiscalYear && rowMonth <= 9) return true;
      return false;
    };

    const aggregateByMonth = (
      rows: any[][] | null | undefined,
    ): Map<number, number> => {
      const map = new Map<number, number>();
      if (!rows || rows.length <= 1) return map;

      const headers = rows[0];
      const valueColIndex = headers.findIndex((h: string) =>
        h.includes('มูลค่ารวม'),
      );
      const ymColIndex = headers.findIndex((h: string) =>
        h.includes('ปีเดือน'),
      );

      if (valueColIndex === -1 || ymColIndex === -1) return map;

      for (const row of rows.slice(1)) {
        const ymStr = String(row[ymColIndex] || '').trim();
        if (!ymStr) continue;

        const [yearStr, monthStr] = ymStr.split('-');
        const rowYear = parseInt(yearStr, 10);
        const rowMonth = parseInt(monthStr, 10);

        if (!isInFiscalYear(rowYear, rowMonth)) continue;

        const val = parseFloat(
          String(row[valueColIndex] || '0').replace(/,/g, ''),
        );
        if (isNaN(val)) continue;

        map.set(rowMonth, (map.get(rowMonth) || 0) + val);
      }

      return map;
    };

    const stockInByMonth = aggregateByMonth(resIn.data.values);
    const stockOutByMonth = aggregateByMonth(resOut.data.values);

    // เรียงตามลำดับปีงบ: ต.ค.(10) → ก.ย.(9)
    const fiscalMonthOrder = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9];
    const months = fiscalMonthOrder.map((month) => {
      const totalIn = stockInByMonth.get(month) || 0;
      const totalOut = stockOutByMonth.get(month) || 0;
      return {
        month,
        totalIn,
        totalOut,
        medSupply: totalOut > 0 ? totalIn / totalOut : 0,
      };
    });

    return { fiscalYear, months };
  }

  async findTopTenStock(
    dto: SearchDto,
  ): Promise<{ name: string; total: number }[]> {
    const [resIn, resOut] = await Promise.all([
      this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: this.sheetName1,
      }),
      this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: this.sheetName2,
      }),
    ]);

    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentFiscalYear =
      currentMonth >= 10 ? now.getFullYear() + 1 : now.getFullYear();

    const fiscalYear = Number(dto.financialYear) || currentFiscalYear;
    const targetMonth = dto.month ? Number(dto.month) : currentMonth;

    const startYear = fiscalYear - 1;
    const isInFiscalYear = (rowYear: number, rowMonth: number): boolean => {
      if (rowYear === startYear && rowMonth >= 10) return true;
      if (rowYear === fiscalYear && rowMonth <= 9) return true;
      return false;
    };

    const aggregate = (
      rows: any[][] | null | undefined,
    ): Map<string, number> => {
      const map = new Map<string, number>();
      if (!rows || rows.length <= 1) return map;

      const headers = rows[0];
      const nameColIndex = headers.findIndex((h: string) =>
        h.includes('ชื่อสินค้า'),
      );
      const valueColIndex = headers.findIndex((h: string) =>
        h.includes('มูลค่ารวม'),
      );
      const ymColIndex = headers.findIndex((h: string) =>
        h.includes('ปีเดือน'),
      );

      if (nameColIndex === -1 || valueColIndex === -1 || ymColIndex === -1)
        return map;

      for (const row of rows.slice(1)) {
        const ymStr = String(row[ymColIndex] || '').trim();
        if (!ymStr) continue;

        const [yearStr, monthStr] = ymStr.split('-');
        const rowYear = parseInt(yearStr, 10);
        const rowMonth = parseInt(monthStr, 10);

        if (!isInFiscalYear(rowYear, rowMonth)) continue;
        if (rowMonth !== targetMonth) continue;

        const name = String(row[nameColIndex] || '').trim();
        if (!name) continue;

        const val = parseFloat(
          String(row[valueColIndex] || '0').replace(/,/g, ''),
        );
        if (isNaN(val)) continue;

        map.set(name, (map.get(name) || 0) + val);
      }

      return map;
    };

    const stockIn = aggregate(resIn.data.values);
    const stockOut = aggregate(resOut.data.values);

    const allNames = new Set([...stockIn.keys(), ...stockOut.keys()]);

    const result: { name: string; total: number }[] = [];

    for (const name of allNames) {
      const inVal = stockIn.get(name) || 0;
      const outVal = stockOut.get(name) || 0;
      const total = inVal - outVal;
      if (total > 0) result.push({ name, total });
    }

    return result.sort((a, b) => b.total - a.total).slice(0, 10);
  }

  async findTopTenTransOut(
    dto: SearchDto,
  ): Promise<{ name: string; total: number }[]> {
    const res = await this.sheets.spreadsheets.values.get({
      spreadsheetId: this.spreadsheetId,
      range: this.sheetName2,
    });

    const rows = res.data.values;
    if (!rows || rows.length <= 1) return [];

    const headers = rows[0];
    const nameColIndex = headers.findIndex((h: string) =>
      h.includes('ชื่อสินค้า'),
    );
    const valueColIndex = headers.findIndex((h: string) =>
      h.includes('มูลค่ารวม'),
    );
    const ymColIndex = headers.findIndex((h: string) => h.includes('ปีเดือน'));

    if (nameColIndex === -1 || valueColIndex === -1 || ymColIndex === -1)
      return [];

    // ✅ คำนวณปีงบปัจจุบันให้ถูก
    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentFiscalYear =
      currentMonth >= 10 ? now.getFullYear() + 1 : now.getFullYear();

    const fiscalYear = Number(dto.financialYear) || currentFiscalYear;
    const targetMonth = dto.month ? Number(dto.month) : undefined;

    const startYear = fiscalYear - 1;
    const isInFiscalYear = (rowYear: number, rowMonth: number): boolean => {
      if (rowYear === startYear && rowMonth >= 10) return true;
      if (rowYear === fiscalYear && rowMonth <= 9) return true;
      return false;
    };

    const aggregated = new Map<string, number>();

    for (const row of rows.slice(1)) {
      const ymStr = String(row[ymColIndex] || '').trim();
      if (!ymStr) continue;

      const [yearStr, monthStr] = ymStr.split('-');
      const rowYear = parseInt(yearStr, 10);
      const rowMonth = parseInt(monthStr, 10);

      if (!isInFiscalYear(rowYear, rowMonth)) continue;
      if (targetMonth && rowMonth !== targetMonth) continue;

      const name = String(row[nameColIndex] || '').trim();
      if (!name) continue;

      const val = parseFloat(
        String(row[valueColIndex] || '0').replace(/,/g, ''),
      );
      if (isNaN(val)) continue;

      aggregated.set(name, (aggregated.get(name) || 0) + val);
    }

    return [...aggregated.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, total]) => ({ name, total }));
  }
}
