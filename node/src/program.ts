import { Worksheet } from './Worksheet';
import { Workbook } from './Workbook';
import { parse } from 'ts-command-line-args';

const { apply, files } = parse<{
    apply: boolean;
    files: Array<string>;
}>({
    apply: { type: Boolean, optional: true },
    files: { type: String, multiple: true, optional: true }
});

if (!files?.length) {
    console.error('No files specified.');
    process.exit(1);
}

await Promise.all(files.map((filePath: string) => processFile(filePath, !apply)));

export async function processFile(xlsxFilePath: string, isDebug: boolean): Promise<void> {
    const workbook = await Workbook.fromFile(xlsxFilePath);
    const { worksheets } = workbook;

    console.log(`File: ${xlsxFilePath}`);
    console.log(`Found ${worksheets.length} worksheets.`);

    worksheets.forEach((sheet) => removeSheetProtection(sheet, isDebug));

    if (!isDebug) {
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
