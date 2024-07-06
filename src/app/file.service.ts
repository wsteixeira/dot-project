import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { PDFDocument, rgb } from 'pdf-lib';
import JSZip from 'jszip';

@Injectable({
  providedIn: 'root',
})
export class FileService {
  constructor() {}

  convertDotxToPdf(file: File): Observable<Blob> {
    return new Observable((observer) => {
      const reader = new FileReader();

      reader.onload = async (event) => {
        try {
          const arrayBuffer = reader.result as ArrayBuffer;

          // Usando JSZip para descompactar o arquivo .dotx
          const zip = new JSZip();
          const content = await zip.loadAsync(arrayBuffer);

          // Acessa o documento XML principal dentro do .dotx
          const documentXml = await content
            .file('word/document.xml')
            ?.async('string');

          if (!documentXml) {
            throw new Error('Erro ao ler o documento .dotx');
          }

          // Extraindo texto simples do XML (simplificação)
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(documentXml, 'text/xml');
          const paragraphs = xmlDoc.getElementsByTagName('w:p');

          let extractedText = '';
          for (let i = 0; i < paragraphs.length; i++) {
            const texts = paragraphs[i].getElementsByTagName('w:t');
            for (let j = 0; j < texts.length; j++) {
              extractedText += texts[j].textContent + ' ';
            }
            extractedText += '\n';
          }

          // Criando um novo PDF usando pdf-lib
          const pdfDoc = await PDFDocument.create();
          const page = pdfDoc.addPage([600, 400]);
          const { width, height } = page.getSize();
          const fontSize = 12;

          page.drawText(extractedText, {
            x: 50,
            y: height - 4 * fontSize,
            size: fontSize,
            color: rgb(0, 0, 0),
          });

          const pdfBytes = await pdfDoc.save();
          const blob = new Blob([pdfBytes], { type: 'application/pdf' });
          observer.next(blob);
          observer.complete();
        } catch (error) {
          observer.error(error);
        }
      };

      reader.readAsArrayBuffer(file);
    });
  }
}
