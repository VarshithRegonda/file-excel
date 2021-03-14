import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
// import { FileUploadComponent } from './file-upload/file-upload.component';
import { UploadComponent } from './upload/upload.component';
import { ReadFileComponent } from './read-file/read-file.component';

@NgModule({
  declarations: [
    AppComponent,
    UploadComponent,
    ReadFileComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
