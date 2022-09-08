import { NgModule, Component, enableProdMode } from '@angular/core';
import { BrowserModule, SafeUrl } from '@angular/platform-browser';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { BcfReader, Topic } from '@parametricos/bcf-js';
import { DomSanitizer } from '@angular/platform-browser';

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

  constructor(private sanitizer: DomSanitizer) {
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
        
        this.topics.forEach(t=>{
          if (t.viewpoints.length > 0)
          {
            console.log(t.viewpoints[0].snapshot)
          }
        })
      }
    }
  }

  private sanitize(url: string) {
    return this.sanitizer.bypassSecurityTrustUrl(url);
  }

  private arrayBufferToBase64(buffer: ArrayBuffer | undefined) : string {
    if (!buffer) return ''
    var binary = '';
    var bytes = new Uint8Array(buffer);
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  topicToImage(topic: Topic) : SafeUrl | undefined{
    if ( topic.viewpoints.length > 0)
    {
      const arrayBuffer = topic.viewpoints[0].snapshot
      const b64 = this.arrayBufferToBase64(arrayBuffer)

      return this.sanitize('data:image/jpg;base64, ' + b64);
    }
    else
    {
      return;
    }
  }
}

export class DisplayedTopic {
  topic: Topic;
  image: Promise<SafeUrl | undefined>;

  constructor(topic : Topic, image : Promise<SafeUrl| undefined>) {
    this.topic = topic;
    this.image = image;
  }
}
