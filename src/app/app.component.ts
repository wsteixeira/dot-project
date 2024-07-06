import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  formData = { name: '', age: 0 };

  constructor(private http: HttpClient) {}

  downloadTemplate() {
    this.http
      .get('/api/template', { responseType: 'blob' })
      .subscribe((blob) => {
        if (blob.size === 0) {
          console.error('Erro: Blob recebido está vazio.');
          return;
        }
        console.log('Blob recebido com sucesso.');

        const reader = new FileReader();
        reader.onload = (e: any) => {
          if (!e.target.result) {
            console.error('Erro: Falha ao ler o conteúdo do blob.');
            return;
          }
          console.log('Conteúdo do blob lido com sucesso.');
          this.convertDotxToDocx(e.target.result, 'documento.dotx');
        };
        reader.onerror = (error) => {
          console.error('Erro ao ler o blob:', error);
        };
        reader.readAsArrayBuffer(blob);
      });
  }

  convertDotxToDocx(content: string | ArrayBuffer, fileName: string) {
    try {
      console.log('Iniciando a conversão do documento...');
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Set data for template placeholders
      doc.setData(this.formData);
      doc.render(); // Prepara o documento para saída
      console.log('Documento convertido com sucesso.');

      // Ajuste na extensão do arquivo para .docx ao salvar
      const outFileName = fileName.replace('.dotx', '.docx');
      const out = doc.getZip().generate({
        type: 'blob',
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      saveAs(out, outFileName);
      console.log('Arquivo .docx salvo com sucesso.');
    } catch (error) {
      console.error('Erro durante a conversão do documento:', error);
    }
  }
}
