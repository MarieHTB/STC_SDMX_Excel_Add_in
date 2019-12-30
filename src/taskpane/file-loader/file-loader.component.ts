/* global console, document, require */

const fileDialog = require("file-dialog");
import { orderBy } from "natural-orderby";

export default class FileLoader {
  files: File[];
  mainDiv: HTMLElement;

  constructor() {
    this.mainDiv = document.getElementById("file-loader");
  }

  selectFiles(label: string, filetype: string, callback: (files: File[]) => void) {
    this.files = [];
    fileDialog({ multiple: true, accept: filetype }).then(fileList => {
      console.log(fileList);
      for (let i = 0; i < fileList.length; i++) {
        this.files.push(fileList[i]);
      }
      // Sort files with natural sort
      this.files = orderBy(this.files, [v => v.name], ["asc"]);

      // Hide the file-loader div
      this.mainDiv.style.display = "none";

      // Add the listbox div
      var listbox = document.createElement("div");
      listbox.setAttribute("id", "file-loader-listbox");
      this.mainDiv.parentNode.insertBefore(listbox, this.mainDiv.nextSibling);

      listbox.style.width = "90%";
      var sel = document.createElement("select");
      listbox.appendChild(sel);
      sel.style.width = "inherit"; //document.documentElement["scrollWidth"] * 0.9;
      sel.size = 5;
      for (const file of this.files) {
        var opt = document.createElement("option");
        opt.text = file.name;
        sel.appendChild(opt);
      }

      // Add reset button
      listbox.appendChild(document.createElement("br"));
      var rstBtn = document.createElement("button");
      rstBtn.innerHTML = "Reset";
      rstBtn.addEventListener("click", ev => this.resetListbox(ev, "reset"));
      listbox.appendChild(rstBtn);

      // Add process button
      var prcBtn = document.createElement("button");
      prcBtn.innerHTML = "Process";
      prcBtn.addEventListener("click", ev => this.resetListbox(ev, callback));
      listbox.appendChild(prcBtn);
    });
  }
  resetListbox(ev, action) {
    // Remove listbox
    var listbox = document.getElementById("file-loader-listbox");
    listbox.parentNode.removeChild(listbox);

    // Show select button
    if (action == "reset") {
      // Hide the file-loader div
      this.mainDiv.style.display = "block";
    } // Call next function
    else if (typeof action == "function") {
      action(this.files);
    }
  }
}
