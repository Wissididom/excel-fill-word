import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { RouterOutlet } from '@angular/router';

@Component({
  selector: 'app-root',
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  excelData: any[] = [];
  templateContent: ArrayBuffer | null = null;

  onExcelUpload(event: any) {
	const target: DataTransfer = <DataTransfer>(event.target);
	const reader: FileReader = new FileReader();
	reader.onload = (e: any) => {
		const arrayBuffer: ArrayBuffer = e.target.result;
		const wb: XLSX.WorkBook = XLSX.read(arrayBuffer, { type: 'array' });
		const wsname: string = wb.SheetNames[0];
		const ws: XLSX.WorkSheet = wb.Sheets[wsname];
		this.excelData = XLSX.utils.sheet_to_json(ws);
	};
	reader.readAsArrayBuffer(target.files[0]);
  }

  onDocxUpload(event: any) {
	const file = event.target.files[0];
	const reader = new FileReader();
	reader.onload = async (e: any) => {
		this.templateContent = e.target.result;
		if (this.excelData.length === 0) {
			console.warn('No Excel data available');
			return;
		}
		if (this.excelData.length === 1) {
			const docBlob = this.generateDocx(this.excelData[0]);
			this.downloadFile(docBlob, 'filled.docx');
		} else {
			const archive = new PizZip();
			for (let i = 0; i < this.excelData.length; i++) {
				const row = this.excelData[i];
				const filename = `filled_${i + 1}.docx`;
				const docBlob = this.generateDocx(row);
				// Convert Blob to ArrayBuffer before adding to zip
				// (PizZip only supports strings, binary strings, Uint8Array, etc.)
				const arrayBuffer = await docBlob.arrayBuffer();
				archive.file(filename, arrayBuffer);
			}
			const zippedContent = archive.generate({ type: 'blob' });
			this.downloadFile(zippedContent, 'filled_docs.zip');
		}
	}
	reader.readAsArrayBuffer(file);
  }
  generateDocx(data: any): Blob {
	if (!this.templateContent) {
		throw new Error('Template content not loaded');
	}
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

  private downloadFile(blob: Blob, filename: string) {
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = filename;
	document.body.appendChild(a);
	a.click();
	document.body.removeChild(a);
	URL.revokeObjectURL(url);
  }
}
