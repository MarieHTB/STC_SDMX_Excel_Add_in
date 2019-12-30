//import { c } from "@angular/core/src/render3";
//import { isFormattedError } from "@angular/compiler";
//import { exists } from "fs";

//import { e } from "@angular/core/src/render3";

//import { checkAndUpdateDirectiveDynamic } from "@angular/core/src/view/provider";
//const DataFrame = require("dataframe-js");
//import DataFrame from "dataframe-js";

/* global console, Excel, document DOMParser, Response, require, Office */
var et = require("elementtree");

const acceptedNS = ["message", "common", "TIME_PERIOD", "structure"];





/* ================================================================================== */
class Field {
  name: string;
  index: number;
  nbr: number;
  unique: string[];
  type: string;
  codelist: any;

  constructor(name: string, type: string, codelist: any) {
    this.name = name;
    this.type = type;
    this.codelist = codelist;
    this.unique = [];
    this.nbr = 0;
  }
  // Build the list of unique values
  addValue(value: string) {
    if (this.name != "OBS_VALUE" && !this.unique.includes(value)) {
      this.unique.push(value);
    }
    this.nbr++;
  }
}
/* ================================================================================== */
class SDMXstruc {
  baseURI = ".//mes:Structures/str:DataStructures/str:DataStructure/str:DataStructureComponents";
  dfURI = ".//mes:Structures/str:Dataflows/str:Dataflow";
  dimURI = this.baseURI + "/str:DimensionList/str:Dimension";
  attURI = this.baseURI + "/str:AttributeList/str:Attribute";
  meaURI = this.baseURI + "/str:MeasureList/str:PrimaryMeasure";
  timURI = this.baseURI + "/str:DimensionList/str:TimeDimension";
  refURI = "str:LocalRepresentation/str:Enumeration/Ref";
  codeURI = ".//mes:Structures/str:Codelists/str:Codelist[@id='{code}']/str:Code[@id='{value}']/com:Name";

  namespaces: { [id: string]: string } = {};

  eTree: any; //et.ElementTree;
  fields: { [id: string]: Field } = {};
  codesets: { [id: string]: any }[] = [];
  dfId: string = "NA";

  constructor(data: string) {
    // parse string to ElementTree
    this.eTree = et.parse(data);
  }
  registerNS() {
    // Get the namespaces
    for (let [key, obj] of Object.entries(this.eTree.getroot().attrib)) {
      var uri: string = String(obj);
      var [xml, prefix] = key.split(":");
      if (xml == "xmlns") {
        for (var name of acceptedNS) {
          if (uri.endsWith(name)) {
            this.namespaces[uri] = prefix;
            break;
          }
        }
      }
    }
  }
  // Read structure to find column names
  getFields() {
    this.registerNS();

    this.findIDs("D", this.dimURI);
    this.findIDs("A", this.attURI);
    this.findIDs("M", this.meaURI);
    this.findIDs("T", this.timURI);

    this.findDFID();

    return this.fields;
  }
  // Find DF ID
  findDFID() {
    var nodes = this.eTree.findall(this.dfURI, this.namespaces);
    for (var i = 0; i < nodes.length; i++) {
      var attrs = nodes[i].attrib;
      var name = attrs["id"];
      if (name) {
        this.dfId = name; //.replace("DF_", "");
        console.log(this.dfId);
        return;
      }
    }
  }
  //
  findIDs(code: string, uri: string) {
    var nodes = this.eTree.findall(uri, this.namespaces);
    for (var i = 0; i < nodes.length; i++) {
      var attrs = nodes[i].attrib;
      var name = attrs["id"];
      if (name) {
        this.fields[name] = new Field(name, code, this.getCodeList(nodes[i]));
      }
    }
  }
  getCodeList(pNode) {
    var codelist = pNode.findall(this.refURI);
    if (codelist == null || codelist.length == 0) {
      return "";
    }
    var value = codelist[0].attrib["id"];
    if (value == undefined) {
      return "";
    }
    return value;
  }
  getCodes(code: string, value: string) {
    var uri = this.codeURI.replace("{code}", code).replace("{value}", value);
    return this.eTree.findall(uri, this.namespaces);
  }
  getCodeSets() {
    for (var field of Object.values(this.fields)) {
      if (field.codelist != "" && field.unique.length > 0) {
        console.log("TABLE", field);
        var codeset: { [id: string]: any } = {};
        codeset["header"] = [field.name, "DescE", "DescF"];
        var records: string[][] = [];
        for (var code of field.unique) {
          var rec = [code, "", ""];
          var descriptions = this.getCodes(field.codelist, code);
          for (var k = 0; k < descriptions.length; k++) {
            var desc = descriptions[k];
            var language = desc.attrib["xml:lang"];
            if (language == "en") {
              rec[1] = desc.text;
            } else if (language == "fr") {
              rec[2] = desc.text;
            }
          }
          records.push(rec);
        }
        codeset["data"] = records;
        this.codesets.push(codeset);
      }
    }
  }
}
/* ================================================================================== */
class SDMXdoc {
  xmlDoc: Document;
  root: Element;
  docNS: string;
  orderedFields: string[] = [];

  constructor(data: string) {
    // parse string to Document
    var parser = new DOMParser();
    this.xmlDoc = parser.parseFromString(data, "text/xml");
    var tag = this.xmlDoc.documentElement.tagName;
    this.root = this.xmlDoc.getElementsByTagName(tag)[0];
    this.docNS = this.xmlDoc.documentElement.namespaceURI;
  }
  registerNS(ns: { [id: string]: string }) {
    // Get the namespaces
    var attrs = this.root.attributes;
    console.log("REG-NS", attrs);
    for (var i = 0; i < attrs.length; i++) {
      var attr = attrs[i];
      if (attr.prefix === "xmlns") {
        for (var name of acceptedNS) {
          if (attr.nodeValue.endsWith(name)) {
            ns[attr.nodeValue] = attr.localName;
            break;
          }
        }
      }
    }
  }
  getRoot() {
    return this.root;
  }
  // Add attributes to fields
  addFields(ele: any) {
    for (var attr of ele.attributes) {
      if (!this.orderedFields.includes(attr.name)) {
        this.orderedFields.push(attr.name);
      }
    }
  }

  // Add values to record and compute a key if has_obs
  addObs(sdmx: SDMX, rec: string[], ele: any) {
    for (var attr of ele.attributes) {
      if (attr.name in sdmx.fields) {
        var field = sdmx.fields[attr.name];
        var value = attr.value.trim();
        rec[field.index] = value;
        field.addValue(value);
      }
    }
    return rec;
  }
  // Make a key to retreive this record
  makeKey(sdmx: SDMX, rec: string[]) {
    var value = rec[sdmx.ObsIndex];
    if (value == "NaN") {
      console.log("Found NaN");
      return null; // null key when OBS_VALUE is NaN
    }

    // Make key by compacting all value in a string. Must remove OBS_VALUE
    rec[sdmx.ObsIndex] = "";
    var key = rec.join(" ");
    rec[sdmx.ObsIndex] = value;

    return key;
  }

  async findFields() {
    // Loop DataSet
    var nodes = this.root.childNodes;
    for (var i = 0; i < nodes.length; i++) {
      var node = nodes[i];
      if (node.namespaceURI == this.docNS && node.nodeName.includes("DataSet")) {
        this.addFields(<Element>node);
        for (var j = 0; j < node.childNodes.length; j++) {
          var child = node.childNodes[j];
          if (child.nodeName == "Series") {
            this.addFields(<Element>child);
            for (var k = 0; k < child.childNodes.length; k++) {
              var obs = child.childNodes[k];
              if (obs.nodeName == "Obs") {
                this.addFields(<Element>obs);
              }
            }
          }
        }
      }
    }
  }
  async findObs(sdmx: SDMX) {
    // Make empty record
    var rec_empty: string[] = [];
    const rec_size = sdmx.columns.length;
    for (var i = 0; i < rec_size; i++) {
      rec_empty.push("");
    }

    // Loop DataSet
    var nodes = this.root.childNodes;
    for (var dsId = 0; dsId < nodes.length; dsId++) {
      var node = nodes[dsId];
      if (node.namespaceURI == this.docNS && node.nodeName.includes("DataSet")) {
        var rec_ds = this.addObs(sdmx, Object.assign([], rec_empty), <Element>node);
        for (var serId = 0; serId < node.childNodes.length; serId++) {
          var child = node.childNodes[serId];
          if (child.nodeName == "Series") {
            var rec_ser = this.addObs(sdmx, Object.assign([], rec_ds), <Element>child);
            for (var obsId = 0; obsId < child.childNodes.length; obsId++) {
              var obs = child.childNodes[obsId];
              if (obs.nodeName == "Obs") {
                var record = this.addObs(sdmx, Object.assign([], rec_ser), <Element>obs);
                var key = this.makeKey(sdmx, record);
                if (key) {
                  sdmx.records[key] = record;
                  sdmx.newRecord();
                }
              }
            }
          }
        }
      }
    }
  }
}

export default class SDMX {
  namespaces: { [id: string]: string } = {};
  columns: string[] = [];
  fields: { [id: string]: Field } = {};
  records: { [id: string]: string[] } = {};
  ObsIndex: number;

  docFiles: File[] = [];
  sdmxDocs: SDMXdoc[] = [];
  strucDoc: SDMXstruc;
  mainDiv: HTMLElement;
  nbrObs: number = 0;

  constructor() {
    this.mainDiv = document.getElementById("sdmx");
  }
  // Read all files and load in Document
  async processFiles(files: File[]) {
    console.log("PROCESSING");
    for (const file of files) {
      if (file.name.includes("Structure")) {
        var data = await this.readFile(file);
        this.strucDoc = new SDMXstruc(data);
        this.strucDoc.getFields();
      } else {
        this.docFiles.push(file);
      }
    }
    await this.processDocs();
    this.strucDoc.getCodeSets();
    await this.updateDataSheet();
    await this.updateCodeSets();
    //await this.addPivotTable();
    console.log("END");
  }

  newRecord() {
    this.nbrObs++;
    if (this.nbrObs % 10000 == 0) {
      console.log("read", this.nbrObs, "records");
    }
  }

  async processDocs() {
    for (var docId = 0; docId < this.docFiles.length; docId++) {
      var sdmxdoc = new SDMXdoc(await this.readFile(this.docFiles[docId]));
      if (docId == 0) {
        sdmxdoc.findFields();
        this.cleanFields(sdmxdoc, this.strucDoc);
      }
      sdmxdoc.findObs(this);
      console.log("Nbr Obs ", Object.keys(this.records).length);
    }
  }
  async cleanFields(doc, struc) {
    var index = 0;
    console.log(doc.orderedFields);
    for (var key of doc.orderedFields) {
      if (key in struc.fields) {
        this.columns.push(key);
        var field = struc.fields[key];
        this.fields[key] = field;
        field.index = index;
        field.nbr = 0;
        index++;
      }
    }
  }

  // Read the SDMX file and load into SDMXdoc
  async readFile(file: File) {
    console.log("Reading " + file.name);
    var data = await new Response(file).text();
    console.log("End reading " + data.length);
    return data;
  }
  async updateDataSheet() {
    try {
      await Excel.run(async context => {
        context.workbook.worksheets.getItemOrNullObject("Data").delete();
        const dataSheet = context.workbook.worksheets.add("Data");
        const header = dataSheet.getRangeByIndexes(0, 0, 1, this.columns.length);
        var table = dataSheet.tables.add(header, true);
        table.name = this.strucDoc.dfId;
        table.getHeaderRowRange().values = [this.columns];
        table.rows.add(null, Object.values(this.records));

        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
          dataSheet.getUsedRange().format.autofitColumns();
          dataSheet.getUsedRange().format.autofitRows();
        }

        // Add pivot table
        var nbr_rows = Object.keys(this.records).length;
        var rangeToAnalyze = dataSheet.getRangeByIndexes(0, 0, nbr_rows, this.columns.length);
        var rangeToPlacePivot = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

        context.workbook.worksheets
          .getActiveWorksheet()
          .pivotTables.add("PT_" + this.strucDoc.dfId, rangeToAnalyze, rangeToPlacePivot);

        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
  async updateCodeSets() {
    try {
      await Excel.run(async context => {
        context.workbook.worksheets.getItemOrNullObject("CodeSets").delete();
        const dataSheet = context.workbook.worksheets.add("CodeSets");
        var index = 0;
        for (var codeset of this.strucDoc.codesets) {
          const header = dataSheet.getRangeByIndexes(index, 0, 1, 3);
          var table = dataSheet.tables.add(header, true);
          table.name = codeset["header"][0];
          table.getHeaderRowRange().values = [codeset["header"]];
          table.rows.add(null, codeset["data"]);
          index += codeset["data"].length + 2;
        }
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
          dataSheet.getUsedRange().format.autofitColumns();
          dataSheet.getUsedRange().format.autofitRows();
        }
        console.log(dataSheet.getUsedRange());
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }
}
