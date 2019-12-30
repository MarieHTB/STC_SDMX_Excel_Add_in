import { Component } from "@angular/core";
const template = require("./app.component.html");
//const fileDialog = require("file-dialog");
//import FileReader from "filereader"

import FileLoader from "../file-loader/file-loader.component";
import SDMX from "../sdmx/sdmx.component";
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent {
  welcomeMessage = " Welcome to the Import SDMX data file";

  async selectFiles() {
    var fileLoader = new FileLoader();
    fileLoader.selectFiles("Process", "*.xml", files => this.processFiles(files));
  }

  processFiles(files: File[]) {
    const sdmx = new SDMX();
    sdmx.processFiles(files);
  }

  /*
  async selectFilesOld() {
    var fileLoader = new FileLoader();
    console.log("calling file-dialog");
    fileDialog({ multiple: true, accept: "*.xml" }).then(files => {
      console.log(files);
      var reader = new FileReader();
      reader.onload = (event: Event) => {
        console.log(event);
        var oParser = new DOMParser();
        var oDOM = oParser.parseFromString(reader.result as string, "text/xml");
        console.log(oDOM);
      };
      reader.readAsText(files[0]);
    });
  }
*/
  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
}
