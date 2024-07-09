import { Component } from '@angular/core';
import { FileService } from './file.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  name: string = '';
  age: string = '';
  docNumber: string = '';
  selectedFile: File | null = null;

  constructor(private fileService: FileService) {}

  onFileSelected(event: any): void {
    this.selectedFile = event.target.files[0];
  }

  importConvertToPdf(): void {
    if (this.selectedFile) {
      const variables = {
        name: this.name,
        age: this.age,
        docNumber: this.docNumber,
      };

      this.fileService.convertDotxToPdf(this.selectedFile, variables).subscribe(
        (blob) => {
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'import_convert_document.pdf';
          document.body.appendChild(a);
          a.click();
          window.URL.revokeObjectURL(url);
        },
        (error) => {
          console.error('Error converting file:', error);
        }
      );
    } else {
      console.error('No file selected');
    }
  }

  downloadConvertToPdf(): void {
    this.fileService.downloadModel().subscribe(
      (dotxBlob: any) => {
        // Assumindo que o serviço convertDotxToPdf pode aceitar um Blob diretamente
        const variables = {
          name: this.name,
          age: this.age,
          docNumber: this.docNumber,
        };
        this.fileService.convertDotxToPdf(dotxBlob, variables).subscribe(
          (pdfBlob) => {
            const url = window.URL.createObjectURL(pdfBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'download_convert_document.pdf'; // Nome do arquivo PDF convertido
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
          },
          (error) => {
            console.error('Error converting .dotx to .pdf:', error);
          }
        );
      },
      (error) => {
        console.error('Error downloading model:', error);
      }
    );
  }
}
