import { Component } from '@angular/core';
import { DataService } from './common/data.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'testapp';

  constructor(private service: DataService) { }

  callMyFunc() {
    this.service.generateXLSX();
  }
}
