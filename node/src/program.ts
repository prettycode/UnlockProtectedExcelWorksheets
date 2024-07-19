import { Worksheet } from './Worksheet';
import { Workbook } from './Workbook';

const DEBUG = true;
const args = ['../data/Unprotected.xlsx', '../data/Protected.xlsx'];

(async (): Promise<Array<void>> => await Promise.all(args.map(processFile)))();

export async function processFile(xlsxFilePath: string): Promise<void> {
    const workbook = await Workbook.fromFile(xlsxFilePath);
    const worksheets: Array<Worksheet> = workbook.worksheets;

    console.log(`File: ${xlsxFilePath}`);
    console.log(`Found ${worksheets.length} worksheets.`);

    worksheets.forEach((sheet) => removeSheetProtection(sheet, DEBUG));

    if (!DEBUG) {
        await workbook.save(xlsxFilePath);
    }

    console.log();
}

export function removeSheetProtection(sheet: Worksheet, reportProtectionOnly: boolean): void {
    const entryName = sheet.zipEntry.entryName;

    if (!sheet.isProtected()) {
        console.log(`${entryName}: No protection found.`);
        return;
    }

    if (reportProtectionOnly) {
        console.log(`${entryName}: Protection found.`);
        return;
    }

    sheet.removeProtection();

    console.log(`${entryName}: Protection removed.`);
}
