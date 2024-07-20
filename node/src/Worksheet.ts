import { IZipEntry } from 'adm-zip';
import { JSDOM } from 'jsdom';

export class Worksheet {
    private readonly zipEntry: IZipEntry;
    private readonly dom: JSDOM;
    private readonly protectionElements: NodeListOf<Element>;
    private protected: boolean;
    public readonly zipEntryName: string;

    constructor(zipEntry: IZipEntry) {
        const sheetXml = zipEntry.getData().toString('utf8');

        this.zipEntry = zipEntry;
        this.zipEntryName = zipEntry.entryName;
        this.dom = new JSDOM(sheetXml, { contentType: 'text/xml' });
        this.protectionElements = this.dom.window.document.querySelectorAll('sheetProtection');
        this.protected = !!this.protectionElements.length;
    }

    public isProtected(): boolean {
        return this.protected;
    }

    public removeProtection(): void {
        if (!this.isProtected) {
            throw new Error('Worksheet is not protected');
        }

        this.protectionElements.forEach((element) => element.remove());
        this.zipEntry.setData(Buffer.from(this.dom.serialize(), 'utf8'));

        this.protected = false;
    }
}
