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
  selectedFile: File | null = null;

  constructor(private fileService: FileService) {}

  onFileSelected(event: any): void {
    this.selectedFile = event.target.files[0];
  }

  convertToPdf(): void {
    if (this.selectedFile) {
      const variables = {
        name: this.name,
        age: this.age,
      };

      this.fileService.convertDotxToPdf(this.selectedFile, variables).subscribe(
        (blob) => {
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'document.pdf';
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
}
