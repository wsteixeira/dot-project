import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { HttpClient } from '@angular/common/http';

import JSZip from 'jszip';
import { jsPDF } from 'jspdf';
import { DOMParser, XMLSerializer } from 'xmldom';

@Injectable({
  providedIn: 'root',
})
export class FileService {
  constructor(private http: HttpClient) {}

  convertDotxToPdf(
    file: File,
    variables: { [key: string]: string }
  ): Observable<Blob> {
    return new Observable((observer) => {
      const reader = new FileReader();

      reader.onload = async (event) => {
        try {
          const arrayBuffer = reader.result as ArrayBuffer;

          console.log('File loaded, processing...');
          console.log('Variables:', variables);

          const zip = new JSZip();
          const content = await zip.loadAsync(arrayBuffer);

          const documentXml = await content
            .file('word/document.xml')
            ?.async('string');

          if (!documentXml) {
            throw new Error('Erro ao ler o documento .dotx');
          }

          console.log('Original document XML:', documentXml);

          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
          const instrTexts = xmlDoc.getElementsByTagName('w:instrText');

          // Iterar de trás para frente
          for (let i = instrTexts.length - 1; i >= 0; i--) {
            const instrText = instrTexts[i];
            const textContent = instrText.textContent || '';

            if (textContent.includes('DOCVARIABLE')) {
              for (const key in variables) {
                if (textContent.includes(key)) {
                  const parent = instrText.parentNode;
                  if (parent) {
                    const newTextElement = xmlDoc.createElement('w:t');
                    newTextElement.textContent = variables[key];
                    parent.replaceChild(newTextElement, instrText);
                    console.log(`Replaced ${key} with ${variables[key]}`);
                  }
                }
              }
            }
          }

          const serializer = new XMLSerializer();
          const updatedXml = serializer.serializeToString(xmlDoc);

          console.log('Updated document XML:', updatedXml);

          const paragraphs = xmlDoc.getElementsByTagName('w:p');

          let extractedText = '';
          for (let i = 0; i < paragraphs.length; i++) {
            const texts = paragraphs[i].getElementsByTagName('w:t');
            for (let j = 0; j < texts.length; j++) {
              extractedText += texts[j].textContent + ' ';
            }
            extractedText += '\n';
          }

          console.log('Extracted text:', extractedText);

          const doc = new jsPDF();
          doc.text(extractedText, 10, 10);
          const pdfBlob = doc.output('blob');
          observer.next(pdfBlob);
          observer.complete();
        } catch (error) {
          console.error('Error processing file:', error);
          observer.error(error);
        }
      };

      reader.readAsArrayBuffer(file);
    });
  }

  downloadModel(): Observable<Blob> {
    // A URL deve corresponder à rota configurada no seu servidor Express
    const url = '/api/download-model';

    // A opção 'responseType: 'blob'' é necessária para tratar a resposta como um arquivo binário
    return this.http.get(url, { responseType: 'blob' });
  }
}
