import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';
import AdmZip from 'adm-zip';
import { JSDOM } from 'jsdom';

const WORKSHEET_REGEX_PATTERN = /xl\/worksheets\/.*\.xml$/;
const DEBUG = true;

const args = ['../data/Unprotected.xlsx', '../data/Protected.xlsx'];

args.forEach((xlsx) => {
    const zip = new AdmZip(readFileSync(xlsx));
    const zipEntries = zip.getEntries();

    const worksheets = zipEntries.filter((entry) => WORKSHEET_REGEX_PATTERN.test(entry.entryName));
    const worksheetsCount = worksheets.length;

    console.log(`File: ${resolve(xlsx)}`);
    console.log(`Found ${worksheetsCount} worksheets.`);

    worksheets.forEach((sheet) => removeSheetProtection(sheet, DEBUG));

    if (!DEBUG) {
        writeFileSync(xlsx, zip.toBuffer());
    }

    console.log();
});

function removeSheetProtection(entry: AdmZip.IZipEntry, reportProtectionOnly: boolean): void {
    const content = entry.getData().toString('utf8');
    const dom = new JSDOM(content, { contentType: 'text/xml' });
    const doc = dom.window.document;

    const sheetProtectionElements = doc.querySelectorAll('sheetProtection');
    const elementCount = sheetProtectionElements.length;

    if (elementCount < 1) {
        console.log(`${entry.entryName}: No protection found.`);
        return;
    }

    if (!reportProtectionOnly) {
        sheetProtectionElements.forEach((element) => element.remove());
        const modifiedContent = dom.serialize();
        entry.setData(Buffer.from(modifiedContent, 'utf8'));

        console.log(`${entry.entryName}: Protection removed.`);
    } else {
        console.log(`${entry.entryName}: Protection found.`);
    }
}
