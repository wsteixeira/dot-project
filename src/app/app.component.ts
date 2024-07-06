import { Component } from '@angular/core';
import { FileService } from './file.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  constructor(private fileService: FileService) {}

  onFileSelected(event: any): void {
    const file: File = event.target.files[0];

    if (file) {
      this.fileService.convertDotxToPdf(file).subscribe(
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
    }
  }
}
