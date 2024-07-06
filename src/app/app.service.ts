import { Injectable } from '@angular/core';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root',
})
export class AppService {
  constructor() {}

  async loadTemplate(url: string): Promise<PizZip> {
    const response = await fetch(url);
    if (response.ok) {
      console.log('response: ', response);
    } else {
      throw new Error(`Erro ao baixar o template: ${response.statusText}`);
    }
    const contentArrayBuffer = await response.arrayBuffer();
    return new PizZip(contentArrayBuffer);
  }

  async generateDocx(data: any, templateUrl: string) {
    const zip = await this.loadTemplate(templateUrl);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(data);
    const docBlob = doc.getZip().generate({
      type: 'blob',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    saveAs(docBlob, 'output.docx');
  }
}
