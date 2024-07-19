import { IZipEntry } from 'adm-zip';
import { JSDOM } from 'jsdom';

export class Worksheet {
    private dom: JSDOM;
    private protected: boolean;
    private protectionElements: NodeListOf<Element>;

    public readonly zipEntry: IZipEntry;

    constructor(zipEntry: IZipEntry) {
        const sheetXml = zipEntry.getData().toString('utf8');

        this.dom = new JSDOM(sheetXml, { contentType: 'text/xml' });
        this.protectionElements = this.dom.window.document.querySelectorAll('sheetProtection');
        this.protected = !!this.protectionElements.length;
        this.zipEntry = zipEntry;
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
