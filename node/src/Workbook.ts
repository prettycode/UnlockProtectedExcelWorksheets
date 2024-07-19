import AdmZip, { IZipEntry } from 'adm-zip';
import { Worksheet } from './Worksheet';
import { promises as fs } from 'fs';

export class Workbook {
    public readonly worksheets: Array<Worksheet>;
    public readonly zip: AdmZip;

    constructor(zip: AdmZip) {
        this.zip = zip;
        this.worksheets = this.getSheetEntries(zip);
    }

    public static async fromFile(xlsxFilePath: string): Promise<Workbook> {
        const fileBuffer = await fs.readFile(xlsxFilePath);
        const zip = new AdmZip(fileBuffer);

        return new Workbook(zip);
    }

    public async save(xlsxFilePath: string): Promise<void> {
        await fs.writeFile(xlsxFilePath, this.zip.toBuffer());
    }

    private getSheetEntries(zip: AdmZip): Array<Worksheet> {
        const SHEET_FILEPATH_REGEX = /xl\/worksheets\/.*\.xml$/;
        const zipFileList = zip.getEntries();
        const sheetFilePathFilter = (entry: IZipEntry): boolean => SHEET_FILEPATH_REGEX.test(entry.entryName);
        const sheetEntries = zipFileList.filter(sheetFilePathFilter);

        return sheetEntries.map((sheet) => new Worksheet(sheet));
    }
}
