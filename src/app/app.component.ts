import { Component } from '@angular/core';
import { AppService } from './app.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  constructor(private service: AppService) {}

  generateDoc() {
    const data = { name: 'John Doe', age: 30 };
    const templateUrl = 'assets/templates/model.dotx';
    this.service.generateDocx(data, templateUrl);
  }
}
