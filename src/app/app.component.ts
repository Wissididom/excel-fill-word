import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import { RouterOutlet } from '@angular/router';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  excelData: any[] = [];
  templateContent: any;

  onExcelUpload(event: any) {
  	const target: DataTransfer = <DataTransfer>(event.target);
  	const reader: FileReader = new FileReader();

  	reader.onload = (e: any) => {
  		const bstr: string = e.target.result;
  		const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
  		const wsname: string = wb.SheetNames[0];
  		const ws: XLSX.WorkSheet = wb.Sheets[wsname];
  		this.excelData = XLSX.utils.sheet_to_json(ws);
  	};
  	reader.readAsBinaryString(target.files[0]);
  }

  onDocxUpload(event: any) {
  	const file = event.target.files[0];
  	const reader = new FileReader();
  	reader.onload = (e: any) => {
  		this.templateContent = e.target.result;
  		if (this.excelData.length === 0) {
  			console.warn('No Excel data available.');
  			return;
  		}
  		if (this.excelData.length === 1) {
  			const docBlob = this.generateDocx(this.excelData[0]);
  			saveAs(docBlob, 'filled.docx');
  		} else {
  			const archive = new PizZip();
  			this.excelData.forEach((row, index) => {
  				const filename = `filled_${index + 1}.docx`;
  				const docBlob = this.generateDocx(row);
  				// Convert Blob to ArrayBuffer before adding to zip
  				// (PizZip only supports strings, binary strings, Uint8Array, etc.)
  				const reader = new FileReader();
  				reader.onload = (e: any) => {
  					const arrayBuffer = e.target.result;
  					archive.file(filename, new Uint8Array(arrayBuffer));
  					if (index === this.excelData.length - 1) {
  						// Wait until last file is added before saving
  						const zippedContent = archive.generate({ type: 'blob' });
  						saveAs(zippedContent, 'filled_docs.zip');
  					}
  				};
  				reader.readAsArrayBuffer(docBlob);
  			});
  		}
  	}
  	reader.readAsBinaryString(file);
  }
  generateDocx(data: any): Blob {
  	const zip = new PizZip(this.templateContent);
  	const doc = new Docxtemplater(zip, {
  		paragraphLoop: true,
  		linebreaks: true
  	});
  	doc.setData(data);
  	try {
  		doc.render();
  	} catch (error) {
  		console.error('Template render error:', error);
  		throw error;
  	}
  	return doc.getZip().generate({
  		type: 'blob',
  		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  	});
  }
}
