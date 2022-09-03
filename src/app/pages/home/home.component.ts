import { NgModule, Component, enableProdMode } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { BcfReader, Topic } from '@parametricos/bcf-js';

if (!/localhost/.test(document.location.host)) {
  enableProdMode();
}

@Component({
  templateUrl: 'home.component.html',
  styleUrls: ['./home.component.scss'],
})
export class HomeComponent {
  topics: Topic[] = [];
  imageItemToDisplay: any = {};

  popupVisible = false;

  constructor() {
  }

  displayImagePopup(e: any) {
    this.imageItemToDisplay = e.file;
    this.popupVisible = true;
  }

  async onFileInput(event: Event) {
    const element = event.currentTarget as HTMLInputElement;
    let fileList: FileList | null = element.files;
    if (fileList) {
      const reader = new BcfReader();

      for (let index = 0; index < fileList.length; index++) {
        const file: File = fileList[index];
        await reader.read(await file.arrayBuffer());

        this.topics = reader.topics;
      }
    }
  }
}
