import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import html2pdf from 'html2pdf.js';
import { jsPDF } from 'jspdf';

export interface EsercizioCircuito {
  nome: string;
  serie?: number;
  ripetizioni?: number;
  serieRipetizioni: string;
  recupero: string;
  hasImage: boolean;
  imagePreview?: string;
  note?: string;
}

export interface Circuito {
  nome: string;
  esercizi: EsercizioCircuito[];
}

export interface EsercizioAerobico {
  nome: string;
  durata: string;
  hasImage: boolean;
  imagePreview?: string;
  note: string;
}

export interface SchedaAllenamento {
  dataInizio: string;
  settimane: number | null;
  nomeCliente: string;
  cognomeCliente: string;
  circuiti: Circuito[];
  lavoroAerobico: EsercizioAerobico[];
}

@Component({
  selector: 'app-allenamento',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './allenamento.component.html',
  styleUrls: ['./allenamento.component.css']
})
export class AllenamentoComponent {
  schedaData: SchedaAllenamento = {
    dataInizio: '',
    settimane: null,
    nomeCliente: '',
    cognomeCliente: '',
    circuiti: [],
    lavoroAerobico: []
  };

  logoPreview: string | undefined;
  showPreviewModal: boolean = false;
  showAddAerobicModal: boolean = false;
  showAddCircuitExerciseModal: boolean = false;
  currentCircuitIndex: number = -1;
  
  // Variabili per gestire la modifica
  isEditingAerobico: boolean = false;
  isEditingCircuito: boolean = false;
  editingAerobicoIndex: number = -1;
  editingCircuitoIndex: number = -1;
  editingEsercizioIndex: number = -1;
  
  // Dati temporanei per le modali
  tempAerobicExercise = {
    nome: '',
    durata: '',
    hasImage: false,
    imagePreview: undefined as string | undefined,
    note: ''
  };
  
  tempCircuitExercise = {
    nome: '',
    serie: null as number | null,
    ripetizioni: null as number | null,
    serieRipetizioni: '',
    recupero: '',
    hasImage: false,
    imagePreview: undefined as string | undefined,
    note: '' // Aggiunta questa proprietà
  };

  // Proprietà per il controllo durata
  tempDurationHours: number = 0;
  tempDurationMinutes: number = 0;
  tempDurationSeconds: number = 0;

  // Proprietà per il controllo recupero
  tempRecuperoHours: number = 0;
  tempRecuperoMinutes: number = 0;
  tempRecuperoSeconds: number = 0;

  constructor() {
    // Costruttore vuoto: nessuna migrazione necessaria per EsercizioAerobico
  }

  // Funzione helper per organizzare esercizi in gruppi di 3
  getExerciseRows(esercizi: EsercizioCircuito[]): EsercizioCircuito[][] {
    const rows: EsercizioCircuito[][] = [];
    for (let i = 0; i < esercizi.length; i += 3) {
      rows.push(esercizi.slice(i, i + 3));
    }
    return rows;
  }

  // Metodo per aggiornare la stringa durata
  updateDurationString() {
    const hours = this.tempDurationHours || 0;
    const minutes = this.tempDurationMinutes || 0;
    const seconds = this.tempDurationSeconds || 0;
    
    if (hours === 0 && minutes === 0 && seconds === 0) {
      this.tempAerobicExercise.durata = '';
      return;
    }
    
    let durationStr = '';
    if (hours > 0) {
      durationStr += `${hours}h`;
    }
    if (minutes > 0) {
      if (durationStr) durationStr += ' ';
      durationStr += `${minutes}m`;
    }
    if (seconds > 0) {
      if (durationStr) durationStr += ' ';
      durationStr += `${seconds}s`;
    }
    
    this.tempAerobicExercise.durata = durationStr;
  }

  // Metodo per aggiornare la stringa recupero
  updateRecuperoString() {
    const hours = this.tempRecuperoHours || 0;
    const minutes = this.tempRecuperoMinutes || 0;
    const seconds = this.tempRecuperoSeconds || 0;
    
    if (hours === 0 && minutes === 0 && seconds === 0) {
      this.tempCircuitExercise.recupero = '';
      return;
    }
    
    let recuperoStr = '';
    if (hours > 0) {
      recuperoStr += `${hours}h`;
    }
    if (minutes > 0) {
      if (recuperoStr) recuperoStr += ' ';
      recuperoStr += `${minutes}m`;
    }
    if (seconds > 0) {
      if (recuperoStr) recuperoStr += ' ';
      recuperoStr += `${seconds}s`;
    }
    
    this.tempCircuitExercise.recupero = recuperoStr;
  }

  // Metodo per parsare la durata esistente
  parseDurationString(duration: string) {
    this.tempDurationHours = 0;
    this.tempDurationMinutes = 0;
    this.tempDurationSeconds = 0;
    
    if (!duration) return;
    
    // Regex per estrarre ore, minuti e secondi (formato breve: h, m, s)
    const hoursMatch = duration.match(/(\d+)h/);
    const minutesMatch = duration.match(/(\d+)m/);
    const secondsMatch = duration.match(/(\d+)s/);
    
    if (hoursMatch) {
      this.tempDurationHours = parseInt(hoursMatch[1]);
    }
    if (minutesMatch) {
      this.tempDurationMinutes = parseInt(minutesMatch[1]);
    }
    if (secondsMatch) {
      this.tempDurationSeconds = parseInt(secondsMatch[1]);
    }
  }

  // Metodo per parsare il recupero esistente
  parseRecuperoString(recupero: string) {
    this.tempRecuperoHours = 0;
    this.tempRecuperoMinutes = 0;
    this.tempRecuperoSeconds = 0;
    
    if (!recupero) return;
    
    // Regex per estrarre ore, minuti e secondi (formato breve: h, m, s)
    const hoursMatch = recupero.match(/(\d+)h/);
    const minutesMatch = recupero.match(/(\d+)m/);
    const secondsMatch = recupero.match(/(\d+)s/);
    
    if (hoursMatch) {
      this.tempRecuperoHours = parseInt(hoursMatch[1]);
    }
    if (minutesMatch) {
      this.tempRecuperoMinutes = parseInt(minutesMatch[1]);
    }
    if (secondsMatch) {
      this.tempRecuperoSeconds = parseInt(secondsMatch[1]);
    }
  }

  // Metodo per aggiornare la stringa serie x ripetizioni
  updateSerieRipetizioniString() {
    const serie = this.tempCircuitExercise.serie || 0;
    const ripetizioni = this.tempCircuitExercise.ripetizioni || 0;
    
    if (serie > 0 && ripetizioni > 0) {
      this.tempCircuitExercise.serieRipetizioni = `${serie} x ${ripetizioni}`;
    } else {
      this.tempCircuitExercise.serieRipetizioni = '';
    }
  }

  // Metodo per parsare le serie x ripetizioni esistenti (per editing)
  parseSerieRipetizioniString(serieRipetizioni: string) {
    if (serieRipetizioni && serieRipetizioni.includes(' x ')) {
      const parts = serieRipetizioni.split(' x ');
      if (parts.length === 2) {
        this.tempCircuitExercise.serie = parseInt(parts[0]) || null;
        this.tempCircuitExercise.ripetizioni = parseInt(parts[1]) || null;
      }
    } else {
      this.tempCircuitExercise.serie = null;
      this.tempCircuitExercise.ripetizioni = null;
    }
  }

  addCircuito() {
    this.schedaData.circuiti.push({
      nome: '',
      esercizi: []
    });
  }

  removeCircuito(index: number) {
    this.schedaData.circuiti.splice(index, 1);
  }

  removeEsercizioAerobico(index: number) {
    this.schedaData.lavoroAerobico.splice(index, 1);
  }

  onLogoChange(event: any) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        this.logoPreview = e.target?.result as string;
      };
      reader.readAsDataURL(file);
    }
  }

  removeEsercizioCircuito(circuitoIndex: number, esercizioIndex: number) {
    this.schedaData.circuiti[circuitoIndex].esercizi.splice(esercizioIndex, 1);
  }

  addEsercizioAerobico() {
    // Apri la modal per aggiungere esercizio aerobico
    this.isEditingAerobico = false;
    this.tempAerobicExercise = {
      nome: '',
      durata: '',
      hasImage: false,
      imagePreview: undefined,
      note: ''
    };
    this.tempDurationHours = 0;
    this.tempDurationMinutes = 0;
    this.tempDurationSeconds = 0;
    this.showAddAerobicModal = true;
  }

  editEsercizioAerobico(index: number) {
    // Apri la modal per modificare esercizio aerobico
    this.isEditingAerobico = true;
    this.editingAerobicoIndex = index;
    const esercizio = this.schedaData.lavoroAerobico[index];
    this.tempAerobicExercise = {
      nome: esercizio.nome,
      durata: esercizio.durata,
      hasImage: esercizio.hasImage,
      imagePreview: esercizio.imagePreview,
      note: esercizio.note
    };
    this.parseDurationString(esercizio.durata);
    this.showAddAerobicModal = true;
  }

  addEsercizioCircuito(circuitoIndex: number) {
    // Apri la modal per aggiungere esercizio del circuito
    this.isEditingCircuito = false;
    this.currentCircuitIndex = circuitoIndex;
    this.tempCircuitExercise = {
      nome: '',
      serie: null as number | null,
      ripetizioni: null as number | null,
      serieRipetizioni: '',
      recupero: '',
      hasImage: false,
      imagePreview: undefined as string | undefined,
      note: '' // Aggiunta questa proprietà
    };
    this.tempRecuperoHours = 0;
    this.tempRecuperoMinutes = 0;
    this.tempRecuperoSeconds = 0;
    this.showAddCircuitExerciseModal = true;
  }

  editEsercizioCircuito(circuitoIndex: number, esercizioIndex: number) {
    this.currentCircuitIndex = circuitoIndex;
    this.isEditingCircuito = true;
    this.editingCircuitoIndex = circuitoIndex;
    this.editingEsercizioIndex = esercizioIndex;
    
    const esercizio = this.schedaData.circuiti[circuitoIndex].esercizi[esercizioIndex];
    
    this.tempCircuitExercise = {
      nome: esercizio.nome,
      serie: null,
      ripetizioni: null,
      serieRipetizioni: esercizio.serieRipetizioni,
      recupero: esercizio.recupero,
      hasImage: esercizio.hasImage,
      imagePreview: esercizio.imagePreview,
      note: esercizio.note || ''
    };
    
    // Parsa le serie x ripetizioni esistenti
    this.parseSerieRipetizioniString(esercizio.serieRipetizioni);
    
    // Parsa il recupero esistente
    this.parseRecuperoString(esercizio.recupero);
    
    this.showAddCircuitExerciseModal = true;
  }

  // Conferma aggiunta/modifica esercizio aerobico
  confirmAddAerobicExercise() {
    if (this.tempAerobicExercise.nome && this.tempAerobicExercise.durata) {
      if (this.isEditingAerobico && this.editingAerobicoIndex >= 0) {
        // Modifica esercizio esistente
        this.schedaData.lavoroAerobico[this.editingAerobicoIndex] = {
          nome: this.tempAerobicExercise.nome,
          durata: this.tempAerobicExercise.durata,
          hasImage: this.tempAerobicExercise.hasImage,
          imagePreview: this.tempAerobicExercise.imagePreview,
          note: this.tempAerobicExercise.note
        };
      } else {
        // Aggiungi nuovo esercizio
        this.schedaData.lavoroAerobico.push({
          nome: this.tempAerobicExercise.nome,
          durata: this.tempAerobicExercise.durata,
          hasImage: this.tempAerobicExercise.hasImage,
          imagePreview: this.tempAerobicExercise.imagePreview,
          note: this.tempAerobicExercise.note
        });
      }
      this.closeAddAerobicModal();
    }
  }

  // Conferma aggiunta/modifica esercizio del circuito
  confirmAddCircuitExercise() {
    if (this.tempCircuitExercise.nome && this.tempCircuitExercise.serieRipetizioni) {
      if (this.isEditingCircuito && this.editingCircuitoIndex >= 0 && this.editingEsercizioIndex >= 0) {
        // Modifica esercizio esistente
        this.schedaData.circuiti[this.editingCircuitoIndex].esercizi[this.editingEsercizioIndex] = {
          nome: this.tempCircuitExercise.nome,
          serieRipetizioni: this.tempCircuitExercise.serieRipetizioni,
          recupero: this.tempCircuitExercise.recupero,
          hasImage: this.tempCircuitExercise.hasImage,
          imagePreview: this.tempCircuitExercise.imagePreview,
          note: this.tempCircuitExercise.note
        };
      } else if (this.currentCircuitIndex >= 0) {
        // Aggiungi nuovo esercizio
        this.schedaData.circuiti[this.currentCircuitIndex].esercizi.push({
          nome: this.tempCircuitExercise.nome,
          serieRipetizioni: this.tempCircuitExercise.serieRipetizioni,
          recupero: this.tempCircuitExercise.recupero,
          hasImage: this.tempCircuitExercise.hasImage,
          imagePreview: this.tempCircuitExercise.imagePreview,
          note: this.tempCircuitExercise.note
        });
      }
      this.closeAddCircuitExerciseModal();
    }
  }

  // Chiudi modali
  closeAddAerobicModal() {
    this.showAddAerobicModal = false;
    this.isEditingAerobico = false;
    this.editingAerobicoIndex = -1;
    this.tempAerobicExercise = {
      nome: '',
      durata: '',
      hasImage: false,
      imagePreview: undefined,
      note: ''
    };
  }

  closeAddCircuitExerciseModal() {
    this.showAddCircuitExerciseModal = false;
    this.isEditingCircuito = false;
    this.currentCircuitIndex = -1;
    this.editingCircuitoIndex = -1;
    this.editingEsercizioIndex = -1;
    this.tempCircuitExercise = {
      nome: '',
      serie: null as number | null,
      ripetizioni: null as number | null,
      serieRipetizioni: '',
      recupero: '',
      hasImage: false,
      imagePreview: undefined,
      note: '' // Aggiunta questa proprietà
    };
    this.tempRecuperoHours = 0;
    this.tempRecuperoMinutes = 0;
    this.tempRecuperoSeconds = 0;
  }

  // Gestione immagini nelle modali
  onTempImageChangeAerobico(event: any) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        this.tempAerobicExercise.imagePreview = e.target?.result as string;
        this.tempAerobicExercise.hasImage = true;
      };
      reader.readAsDataURL(file);
    }
  }

  onTempImageChangeCircuito(event: any) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        this.tempCircuitExercise.imagePreview = e.target?.result as string;
        this.tempCircuitExercise.hasImage = true;
      };
      reader.readAsDataURL(file);
    }
  }

  // Metodo helper per convertire base64 in ArrayBuffer
  base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binaryString = window.atob(base64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }

  // Metodo helper per determinare l'estensione dell'immagine
  getImageExtension(base64String: string): 'jpeg' | 'png' | 'gif' {
    if (base64String.includes('data:image/png')) return 'png';
    if (base64String.includes('data:image/gif')) return 'gif';
    return 'jpeg'; // default per JPEG/JPG
  }

  onImageChangeCircuito(event: any, circuitoIndex: number, esercizioIndex: number) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        this.schedaData.circuiti[circuitoIndex].esercizi[esercizioIndex].imagePreview = e.target?.result as string;
        this.schedaData.circuiti[circuitoIndex].esercizi[esercizioIndex].hasImage = true;
      };
      reader.readAsDataURL(file);
    }
  }

  onImageChangeAerobico(event: any, index: number) {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        this.schedaData.lavoroAerobico[index].imagePreview = e.target?.result as string;
        this.schedaData.lavoroAerobico[index].hasImage = true;
      };
      reader.readAsDataURL(file);
    }
  }

  async exportToExcel() {
    try {
      alert('Export in corso... Le immagini saranno incluse nel file Excel!');
      console.log('Inizio export Excel con ExcelJS...');
      
      const workbook = new ExcelJS.Workbook();
      const riepilogoSheet = workbook.addWorksheet('Riepilogo');
      
      // Impostazione delle dimensioni delle colonne (tutte 2,89 = 33 pixel)
      riepilogoSheet.columns = [
        { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, // A-D per logo
        { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, // E-J
        { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, // K-P
        { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, { width: 2.89 }, // Q-V
        { width: 2.89 }, { width: 2.89 } // W-X
      ];

      // Crea un foglio griglia pesi per ogni circuito
      this.schedaData.circuiti.forEach((circuito, circuitoIndex) => {
        if (circuito.esercizi.length > 0) {
          const nomeSheet = `Circuito ${circuitoIndex + 1} - Pesi`;
          const circuitoSheet = workbook.addWorksheet(nomeSheet);
          
          // Estendi la griglia per coprire tutto il foglio (fino alla colonna Z = 26 colonne)
          const maxColonneCircuito = 26; // A-Z
          const colonneEffettiveCircuito = Math.max(circuito.esercizi.length, maxColonneCircuito);
          
          // Imposta larghezza colonne per tutte le colonne (partendo dalla colonna A)
          for (let i = 1; i <= colonneEffettiveCircuito; i++) {
            circuitoSheet.getColumn(i).width = 15;
          }
          
          // Header con nomi esercizi del circuito (riga 1, partendo dalla colonna A)
          for (let index = 0; index < colonneEffettiveCircuito; index++) {
            const headerCell = circuitoSheet.getCell(1, index + 1); // Inizia dalla colonna A
            
            if (index < circuito.esercizi.length) {
              // Usa il nome dell'esercizio del circuito
              headerCell.value = circuito.esercizi[index].nome;
            } else {
              // Lascia vuoto per le colonne extra
              headerCell.value = '';
            }
            
            headerCell.font = { bold: true, size: 10 };
            headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            headerCell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
          }
          
          // Crea righe per le sessioni (senza etichette) - fino alla riga 39
          const numSessioniCircuito = 38; // 38 sessioni per arrivare alla riga 39
          for (let sessione = 1; sessione <= numSessioniCircuito; sessione++) {
            const rowIndex = sessione + 1; // Riga 2 in poi
            
            // Celle per tutte le colonne (fino alla colonna Z)
            for (let colIndex = 0; colIndex < colonneEffettiveCircuito; colIndex++) {
              const pesoCell = circuitoSheet.getCell(rowIndex, colIndex + 1);
              pesoCell.value = '';
              pesoCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
            }
          }
        }
      });
      
      // Impostazione dell'altezza delle righe (14,40 = 24 pixel circa in Excel)
      for (let i = 1; i <= 6; i++) {
        riepilogoSheet.getRow(i).height = 14.40;
      }
      
      // Aggiunta del titolo "SCHEDA D'ALLENAMENTO" nelle celle E1:X2
      // Unisci le celle E1:X2
      riepilogoSheet.mergeCells('E1:X2');
      
      // Imposta il testo e la formattazione
      const titleCell = riepilogoSheet.getCell('E1');
      titleCell.value = "SCHEDA D'ALLENAMENTO";
      titleCell.font = { 
        size: 16, 
        bold: true,
        name: 'Arial'
      };
      titleCell.alignment = { 
        horizontal: 'center', 
        vertical: 'middle' 
      };
      titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' } // Sfondo bianco
      };
      
      // Inserimento del logo se presente (celle A1:D6)
      if (this.logoPreview) {
        try {
          console.log('Aggiungendo logo all\'Excel...');
          
          // Converti base64 in ArrayBuffer
          const base64Data = this.logoPreview.split(',')[1];
          const imageBuffer = this.base64ToArrayBuffer(base64Data);
          
          // Determina l'estensione dell'immagine
          const imageExtension = this.getImageExtension(this.logoPreview);
          
          // Aggiungi l'immagine al workbook
          const imageId = workbook.addImage({
            buffer: imageBuffer,
            extension: imageExtension
          });
          
          // Posiziona l'immagine esattamente nelle celle A1:D6
          riepilogoSheet.addImage(imageId, 'A1:D6');
          
          console.log('Logo aggiunto con successo!');
        } catch (imageError) {
          console.error('Errore nell\'aggiunta del logo:', imageError);
        }
      }
      
      // Nome e Cognome: celle E3:X4 (righe 3-4, colonne E-X)
      riepilogoSheet.mergeCells('E3:X4');
      const clienteCell = riepilogoSheet.getCell('E3');
      clienteCell.value = this.schedaData.nomeCliente + ' ' + this.schedaData.cognomeCliente;
      clienteCell.alignment = { horizontal: 'center', vertical: 'middle' };
      clienteCell.font = { bold: true, size: 14 };
      
      // Data d'inizio: celle F5:M5 (riga 5, colonne F-M)
      riepilogoSheet.mergeCells('F5:M5');
      const dataInizioCell = riepilogoSheet.getCell('F5');
      dataInizioCell.value = {
        richText: [
          { text: 'Data d\'inizio: ', font: { size: 12, bold: false } },
          { text: this.schedaData.dataInizio, font: { size: 12, bold: true } }
        ]
      };
      dataInizioCell.alignment = { horizontal: 'center', vertical: 'middle' };
      dataInizioCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      
      // Settimane: celle O5:W5 (riga 5, colonne O-W)
      if (this.schedaData.settimane) {
        riepilogoSheet.mergeCells('O5:W5');
        const settimaneCell = riepilogoSheet.getCell('O5');
        settimaneCell.value = {
          richText: [
            { text: 'Settimane: ', font: { size: 12, bold: false } },
            { text: this.schedaData.settimane.toString(), font: { size: 12, bold: true } }
          ]
        };
        settimaneCell.alignment = { horizontal: 'center', vertical: 'middle' };
        settimaneCell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
      
      // LAVORO AEROBICO: riga 7, ampiezza A-X, sfondo giallo, carattere bold
      riepilogoSheet.mergeCells('A7:X7');
      const lavoroAerobicoTitleCell = riepilogoSheet.getCell('A7');
      lavoroAerobicoTitleCell.value = 'LAVORO AEROBICO';
      lavoroAerobicoTitleCell.font = { bold: true, size: 14 };
      lavoroAerobicoTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      lavoroAerobicoTitleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' } // Sfondo giallo
      };
      lavoroAerobicoTitleCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      
      // Gestione degli esercizi aerobici (massimo 3 per riga, 8 celle di ampiezza ciascuno)
      if (this.schedaData.lavoroAerobico.length > 0) {
        const colPositions = [
          { start: 0, end: 7 },   // A-H (colonne 0-7)
          { start: 8, end: 15 },  // I-P (colonne 8-15) 
          { start: 16, end: 23 }  // Q-X (colonne 16-23)
        ];
        
        this.schedaData.lavoroAerobico.forEach((esercizio, index) => {
          if (index < 3) { // Massimo 3 esercizi per riga
            const colPos = colPositions[index];
            const startCol = String.fromCharCode(65 + colPos.start); // A, I, Q
            const endCol = String.fromCharCode(65 + colPos.end);     // H, P, X
            
            // Immagine: righe 8-15, ampiezza 8 celle (sempre 8x8)
            if (esercizio.hasImage && esercizio.imagePreview) {
              try {
                const base64Data = esercizio.imagePreview.split(',')[1];
                const imageBuffer = this.base64ToArrayBuffer(base64Data);
                const imageExtension = this.getImageExtension(esercizio.imagePreview);
                
                const imageId = workbook.addImage({
                  buffer: imageBuffer,
                  extension: imageExtension
                });
                
                // Posiziona l'immagine nelle 8x8 celle
                riepilogoSheet.addImage(imageId, `${startCol}8:${endCol}15`);
              } catch (imageError) {
                console.error('Errore nell\'aggiunta dell\'immagine aerobica:', imageError);
              }
            }
            
            // Aggiungi bordi sottili all'area dell'immagine (8x8 celle)
            for (let row = 8; row <= 15; row++) {
              for (let col = colPos.start; col <= colPos.end; col++) {
                const cell = riepilogoSheet.getCell(row, col + 1); // +1 perché le colonne Excel iniziano da 1
                cell.border = {
                  top: { style: 'thin' },
                  bottom: { style: 'thin' },
                  left: { style: 'thin' },
                  right: { style: 'thin' }
                };
              }
            }
            
            // Descrizione: riga 16, sotto l'immagine, ampiezza 8 celle
            riepilogoSheet.mergeCells(`${startCol}16:${endCol}16`);
            const descrizioneCell = riepilogoSheet.getCell(`${startCol}16`);
            descrizioneCell.value = esercizio.nome;
            descrizioneCell.alignment = { horizontal: 'center', vertical: 'middle' };
            descrizioneCell.font = { size: 8, bold: true };
            descrizioneCell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            
            // Minuti: riga 17, sotto la descrizione, ampiezza 8 celle  
            riepilogoSheet.mergeCells(`${startCol}17:${endCol}17`);
            const minutiCell = riepilogoSheet.getCell(`${startCol}17`);
            minutiCell.value = esercizio.durata;
            minutiCell.alignment = { horizontal: 'center', vertical: 'middle' };
            minutiCell.font = { size: 10 };
            minutiCell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };

            // Note: riga 18, se presente, ampiezza 8 celle
            if (esercizio.note && esercizio.note.trim()) {
              riepilogoSheet.mergeCells(`${startCol}18:${endCol}18`);
              const noteCell = riepilogoSheet.getCell(`${startCol}18`);
              noteCell.value = esercizio.note;
              noteCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
              noteCell.font = { size: 8, italic: true, bold: true }; // Aggiunto bold per migliorare la leggibilità
              noteCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFFF' } // Sfondo bianco per le note aerobiche
              };
              noteCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
              
              // Adatta l'altezza della riga 18 in base al contenuto delle note aerobiche
              const noteLength = esercizio.note.length;
              const estimatedHeight = Math.max(20, Math.min(80, noteLength / 6 * 3)); // Stessa logica dei circuiti
              const currentRowHeight = riepilogoSheet.getRow(18).height || 0;
              // Usa l'altezza maggiore se ci sono più esercizi aerobici con note sulla stessa riga
              riepilogoSheet.getRow(18).height = Math.max(currentRowHeight, estimatedHeight);
            } else {
              // Cella vuota se non ci sono note
              riepilogoSheet.mergeCells(`${startCol}18:${endCol}18`);
              const emptyCell = riepilogoSheet.getCell(`${startCol}18`);
              emptyCell.value = '';
              emptyCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFFF' } // Sfondo bianco anche per celle vuote
              };
              emptyCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
            }
          }
        });
      }
      
      // Inizio contenuto circuiti dalla riga 19 (dopo il lavoro aerobico con note)
      let currentRow = 19;
      
      this.schedaData.circuiti.forEach((circuito, circuitoIndex) => {
        // Titolo del circuito: riga singola, ampiezza A-X, altezza 1 cella
        // Alterna colori: blu per circuiti dispari (1,3,5...), giallo per pari (2,4,6...)
        const isOdd = (circuitoIndex + 1) % 2 === 1;
        const backgroundColor = isOdd ? 'FF0080FF' : 'FFFFFF00'; // Blu o Giallo
        
        riepilogoSheet.mergeCells(`A${currentRow}:X${currentRow}`);
        const circuitoTitleCell = riepilogoSheet.getCell(`A${currentRow}`);
        circuitoTitleCell.value = `CIRCUITO ${circuitoIndex + 1}: ${circuito.nome.toUpperCase()}`;
        circuitoTitleCell.font = { bold: true, size: 14 };
        circuitoTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        circuitoTitleCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: backgroundColor }
        };
        circuitoTitleCell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        riepilogoSheet.getRow(currentRow).height = 20;
        currentRow++;
        
        // Esercizi del circuito (massimo 3 per riga, 8 celle di ampiezza ciascuno)
        if (circuito.esercizi.length > 0) {
          const colPositions = [
            { start: 0, end: 7 },   // A-H (colonne 0-7)
            { start: 8, end: 15 },  // I-P (colonne 8-15) 
            { start: 16, end: 23 }  // Q-X (colonne 16-23)
          ];
          
          for (let i = 0; i < circuito.esercizi.length; i += 3) {
            // Immagini degli esercizi (8 righe per le immagini)
            const eserciziFila = circuito.esercizi.slice(i, i + 3);
            
            eserciziFila.forEach((esercizio, index) => {
              const colPos = colPositions[index];
              const startCol = String.fromCharCode(65 + colPos.start); // A, I, Q
              const endCol = String.fromCharCode(65 + colPos.end);     // H, P, X
              
              // Immagine: 8 righe, ampiezza 8 celle (8x8)
              if (esercizio.hasImage && esercizio.imagePreview) {
                try {
                  const base64Data = esercizio.imagePreview.split(',')[1];
                  const imageBuffer = this.base64ToArrayBuffer(base64Data);
                  const imageExtension = this.getImageExtension(esercizio.imagePreview);
                  
                  const imageId = workbook.addImage({
                    buffer: imageBuffer,
                    extension: imageExtension
                  });
                  
                  // Posiziona l'immagine nelle 8x8 celle
                  riepilogoSheet.addImage(imageId, `${startCol}${currentRow}:${endCol}${currentRow + 7}`);
                } catch (imageError) {
                  console.error('Errore nell\'aggiunta dell\'immagine del circuito:', imageError);
                }
              }
              
              // Aggiungi bordi sottili all'area dell'immagine (8x8 celle)
              for (let row = currentRow; row < currentRow + 8; row++) {
                for (let col = colPos.start; col <= colPos.end; col++) {
                  const cell = riepilogoSheet.getCell(row, col + 1); // +1 perché le colonne Excel iniziano da 1
                  cell.border = {
                    top: { style: 'thin' },
                    bottom: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' }
                  };
                }
              }
            });
            
            currentRow += 8; // Salta le 8 righe delle immagini
            
            // Nome esercizio e serie/ripetizioni: prima riga sotto l'immagine
            eserciziFila.forEach((esercizio, index) => {
              const colPos = colPositions[index];
              const startCol = String.fromCharCode(65 + colPos.start); // A, I, Q
              const endCol = String.fromCharCode(65 + colPos.end);     // H, P, X
              
              // Nome esercizio e serie/ripetizioni: prima riga, ampiezza 8 celle
              riepilogoSheet.mergeCells(`${startCol}${currentRow}:${endCol}${currentRow}`);
              const nomeEsercizioCell = riepilogoSheet.getCell(`${startCol}${currentRow}`);
              nomeEsercizioCell.value = `${esercizio.nome} - ${esercizio.serieRipetizioni}`;
              nomeEsercizioCell.alignment = { horizontal: 'center', vertical: 'middle' };
              nomeEsercizioCell.font = { size: 8, bold: true };
              nomeEsercizioCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
            });
            
            currentRow++; // Passa alla riga successiva per il recupero
            
            // Recupero: seconda riga sotto il nome
            eserciziFila.forEach((esercizio, index) => {
              const colPos = colPositions[index];
              const startCol = String.fromCharCode(65 + colPos.start); // A, I, Q
              const endCol = String.fromCharCode(65 + colPos.end);     // H, P, X
              
              // Recupero: seconda riga, ampiezza 8 celle
              riepilogoSheet.mergeCells(`${startCol}${currentRow}:${endCol}${currentRow}`);
              const recuperoCell = riepilogoSheet.getCell(`${startCol}${currentRow}`);
              recuperoCell.value = esercizio.recupero ? `Recupero: ${esercizio.recupero}` : '';
              recuperoCell.alignment = { horizontal: 'center', vertical: 'middle' };
              recuperoCell.font = { size: 8 };
              recuperoCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
              };
            });
            
            currentRow++; // Passa alla riga successiva per le note
            
            // Note: terza riga sotto il recupero, se presente
            eserciziFila.forEach((esercizio, index) => {
              const colPos = colPositions[index];
              const startCol = String.fromCharCode(65 + colPos.start); // A, I, Q
              const endCol = String.fromCharCode(65 + colPos.end);     // H, P, X
              
              if (esercizio.note && esercizio.note.trim()) {
                // Note: terza riga, ampiezza 8 celle con adattamento altezza
                riepilogoSheet.mergeCells(`${startCol}${currentRow}:${endCol}${currentRow}`);
                const noteCell = riepilogoSheet.getCell(`${startCol}${currentRow}`);
                noteCell.value = esercizio.note;
                noteCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                noteCell.font = { size: 8, italic: true, bold: true }; // Aggiunto bold per migliorare la leggibilità
                noteCell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' } // Sfondo bianco per le note
                };
                noteCell.border = {
                  top: { style: 'thin' },
                  left: { style: 'thin' },
                  bottom: { style: 'thin' },
                  right: { style: 'thin' }
                };
                
                // Adatta l'altezza della riga in base al contenuto delle note
                const noteLength = esercizio.note.length;
                const estimatedHeight = Math.max(20, Math.min(80, noteLength / 6 * 3)); // Aumentata l'altezza minima
                riepilogoSheet.getRow(currentRow).height = estimatedHeight;
              } else {
                // Cella vuota se non ci sono note
                riepilogoSheet.mergeCells(`${startCol}${currentRow}:${endCol}${currentRow}`);
                const emptyCell = riepilogoSheet.getCell(`${startCol}${currentRow}`);
                emptyCell.value = '';
                emptyCell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' } // Sfondo bianco anche per celle vuote
                };
                emptyCell.border = {
                  top: { style: 'thin' },
                  left: { style: 'thin' },
                  bottom: { style: 'thin' },
                  right: { style: 'thin' }
                };
              }
            });
            
            currentRow++; // Passa alla riga successiva per la prossima serie di esercizi
          }
        }
      });

      // Aggiungi contorno nero sottile a tutta la scheda
      // Determina l'ultima riga utilizzata
      const lastRow = currentRow - 1;
      
      // Contorno superiore (riga 1 da A a X)
      for (let col = 1; col <= 24; col++) { // A=1, X=24
        const cell = riepilogoSheet.getCell(1, col);
        cell.border = {
          ...cell.border,
          top: { style: 'thin', color: { argb: 'FF000000' } }
        };
      }
      
      // Contorno inferiore (ultima riga da A a X)
      for (let col = 1; col <= 24; col++) { // A=1, X=24
        const cell = riepilogoSheet.getCell(lastRow, col);
        cell.border = {
          ...cell.border,
          bottom: { style: 'thin', color: { argb: 'FF000000' } }
        };
      }
      
      // Contorno sinistro (colonna A da riga 1 a ultima riga)
      for (let row = 1; row <= lastRow; row++) {
        const cell = riepilogoSheet.getCell(row, 1); // Colonna A
        cell.border = {
          ...cell.border,
          left: { style: 'thin', color: { argb: 'FF000000' } }
        };
      }
      
      // Contorno destro (colonna X da riga 1 a ultima riga)
      for (let row = 1; row <= lastRow; row++) {
        const cell = riepilogoSheet.getCell(row, 24); // Colonna X
        cell.border = {
          ...cell.border,
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
      }

      const nomeCompleto = (this.schedaData.nomeCliente + '_' + this.schedaData.cognomeCliente).replace(/[^a-zA-Z0-9]/g, '_');
      const nomeFile = nomeCompleto ? 'Scheda_' + nomeCompleto + '_' + new Date().toISOString().split('T')[0] + '.xlsx' : 'Scheda_Allenamento_' + new Date().toISOString().split('T')[0] + '.xlsx';
      
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = nomeFile;
      link.click();
      
      window.URL.revokeObjectURL(url);
      
      console.log('Export Excel completato!');
      
    } catch (error: any) {
      console.error('Errore durante export Excel:', error);
      alert('Errore durante la generazione del file Excel. Controlla la console.');
    }
  }

  // Aggiungi il metodo clearForm mancante
  clearForm() {
    // Conferma prima di cancellare
    if (confirm('Sei sicuro di voler cancellare tutti i dati della scheda?')) {
      this.schedaData = {
        dataInizio: '',
        settimane: null,
        nomeCliente: '',
        cognomeCliente: '',
        circuiti: [],
        lavoroAerobico: []
      };
      
      // Reset logo
      this.logoPreview = undefined;
      
      // Chiudi eventuali modal aperte
      this.showPreviewModal = false;
      this.showAddAerobicModal = false;
      this.showAddCircuitExerciseModal = false;
      
      console.log('Form cancellato con successo');
    }
  }

  showPreview() {
    // Verifica che ci siano dati da mostrare
    if (!this.schedaData.nomeCliente && !this.schedaData.cognomeCliente && 
        this.schedaData.circuiti.length === 0 && this.schedaData.lavoroAerobico.length === 0) {
      alert('Inserisci almeno alcuni dati prima di visualizzare l\'anteprima');
      return;
    }
    
    console.log('Apertura anteprima...');
    this.showPreviewModal = true;
  }
 /**
   * NUOVA FUNZIONE EXPORT TO PDF
   * Usa jsPDF per costruire il documento in modo programmatico, garantendo un layout preciso.
   */
  async exportToPDF() {
    try {
      // --- IMPOSTAZIONI GENERALI DEL DOCUMENTO ---
      const doc = new jsPDF('p', 'mm', 'a4'); // Orientamento portrait, unità in mm, formato A4
      const pageHeight = doc.internal.pageSize.getHeight();
      const pageWidth = doc.internal.pageSize.getWidth();
      const margin = 10;
      let currentY = margin; // Posizione verticale corrente, parte dal margine superiore

      // Funzione helper per aggiungere una nuova pagina e resettare la Y
      const addPageIfNeeded = (neededHeight: number) => {
        if (currentY + neededHeight > pageHeight - margin) {
          doc.addPage();
          currentY = margin;
        }
      };

    // --- 1. HEADER (LOGO, TITOLO, CLIENTE) ---
      const headerHeight = 30;
      addPageIfNeeded(headerHeight);

      const logoSize = 25;
      const infoX = margin + logoSize + 5; // Spazio dopo il logo
      const infoY = currentY + 6; // Piccolo offset per centrare verticalmente il testo rispetto al logo

      // Logo (se presente)
      if (this.logoPreview) {
        const imgData = this.logoPreview;
        const imgExtension = this.getImageExtension(imgData);
        doc.addImage(imgData, imgExtension, margin, currentY, logoSize, logoSize);
      }

      // Titolo principale
      doc.setFontSize(18);
      doc.setFont('helvetica', 'bold');
      doc.text("SCHEDA D'ALLENAMENTO", infoX, infoY, { align: 'left' });

      // Nome Cliente
      doc.setFontSize(14);
      doc.text(`${this.schedaData.nomeCliente} ${this.schedaData.cognomeCliente}`, infoX, infoY + 8, { align: 'left' });

      // Data Inizio e Settimane (vicini)
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.text('Data Inizio:', infoX, infoY + 16, { align: 'left' });
      doc.setFont('helvetica', 'normal');
      doc.text(` ${this.schedaData.dataInizio}`, infoX + 20, infoY + 16, { align: 'left' });

      if (this.schedaData.settimane) {
        doc.setFont('helvetica', 'bold');
        doc.text('Settimane:', infoX + 45, infoY + 16, { align: 'left' });
        doc.setFont('helvetica', 'normal');
        doc.text(` ${this.schedaData.settimane}`, infoX + 65, infoY + 16, { align: 'left' });
      }
      currentY += logoSize + 5; // Aggiorna la posizione verticale dopo il logo

      // --- 2. SEZIONE LAVORO AEROBICO ---
      if (this.schedaData.lavoroAerobico.length > 0) {
        const aerobicSectionHeight = 65; // Altezza stimata (titolo + 1 riga di esercizi)
        addPageIfNeeded(aerobicSectionHeight);

        // Titolo sezione
        doc.setFillColor('#FFC107'); // Giallo
        doc.rect(margin, currentY, pageWidth - (margin * 2), 10, 'F'); // x, y, larg, alt, stile
        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 0, 0);
        doc.text('LAVORO AEROBICO', pageWidth / 2, currentY + 7, { align: 'center' });
        currentY += 12;

        // Esercizi (layout a 3 colonne)
        const colWidth = (pageWidth - margin * 2) / 3;
        this.schedaData.lavoroAerobico.forEach((esercizio, index) => {
          const colIndex = index % 3;
          const xPos = margin + colIndex * colWidth;

          // Immagine
          if (esercizio.imagePreview) {
            doc.addImage(esercizio.imagePreview, this.getImageExtension(esercizio.imagePreview), xPos + 5, currentY, 50, 30);
          }
          doc.rect(xPos + 5, currentY, 50, 30); // Bordo intorno all'immagine

          // Testi
          doc.setFontSize(9);
          doc.setFont('helvetica', 'bold');
          doc.text(esercizio.nome, xPos + 30, currentY + 35, { align: 'center' });

          doc.setFontSize(8);
          doc.setFont('helvetica', 'normal');
          doc.text(esercizio.durata, xPos + 30, currentY + 40, { align: 'center' });
          
          if(esercizio.note) {
              doc.setFont('helvetica', 'italic');
              // Il metodo text gestisce il wrapping automatico se si fornisce una larghezza massima
              const noteLines = doc.splitTextToSize(esercizio.note, colWidth - 10);
              doc.text(noteLines, xPos + 30, currentY + 45, { align: 'center' });
          }
        });
        currentY += aerobicSectionHeight - 12; // Aggiorna Y dopo la sezione
      }
      /*
         // --- 2. SEZIONE LAVORO AEROBICO ---
      if ( this.schedaData.circuiti.length > 0) {
        const circuitoSectionHeight = 65; // Altezza stimata (titolo + 1 riga di esercizi)
        addPageIfNeeded(circuitoSectionHeight);

        // Titolo sezione
        doc.setFillColor('#FFC107'); // Giallo
        doc.rect(margin, currentY, pageWidth - (margin * 2), 10, 'F'); // x, y, larg, alt, stile
        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 0, 0);
        doc.text('LAVORO AEROBICO', pageWidth / 2, currentY + 7, { align: 'center' });
        currentY += 12;

        // Esercizi (layout a 3 colonne)
        const colWidth = (pageWidth - margin * 2) / 3;
        this.schedaData.lavoroAerobico.forEach((esercizio, index) => {
          const colIndex = index % 3;
          const xPos = margin + colIndex * colWidth;

          // Immagine
          if (esercizio.imagePreview) {
            doc.addImage(esercizio.imagePreview, this.getImageExtension(esercizio.imagePreview), xPos + 5, currentY, 50, 30);
          }
          doc.rect(xPos + 5, currentY, 50, 30); // Bordo intorno all'immagine

          // Testi
          doc.setFontSize(9);
          doc.setFont('helvetica', 'bold');
          doc.text(esercizio.nome, xPos + 30, currentY + 35, { align: 'center' });

          doc.setFontSize(8);
          doc.setFont('helvetica', 'normal');
          doc.text(esercizio.durata, xPos + 30, currentY + 40, { align: 'center' });
          
          if(esercizio.note) {
              doc.setFont('helvetica', 'italic');
              // Il metodo text gestisce il wrapping automatico se si fornisce una larghezza massima
              const noteLines = doc.splitTextToSize(esercizio.note, colWidth - 10);
              doc.text(noteLines, xPos + 30, currentY + 45, { align: 'center' });
          }
        });
        currentY += circuitoSectionHeight - 12; // Aggiorna Y dopo la sezione
      }
      */
      // --- 3. SEZIONE CIRCUITI (stile lavoro aerobico) ---
      if (this.schedaData.circuiti.length > 0) {
        this.schedaData.circuiti.forEach((circuito, circuitoIndex) => {
          // Titolo circuito
          addPageIfNeeded(15);
          doc.setFillColor(circuitoIndex % 2 === 0 ? '#2196F3' : '#FFC107'); // Blu alternato a giallo
          doc.rect(margin, currentY, pageWidth - (margin * 2), 10, 'F');
          doc.setFontSize(13);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(0, 0, 0);
          doc.text(
            `CIRCUITO ${circuitoIndex + 1}: ${circuito.nome.toUpperCase()}`,
            pageWidth / 2,
            currentY + 7,
            { align: 'center' }
          );
          currentY += 12;

          // Esercizi del circuito (max 3 per riga)
          const colWidth = (pageWidth - margin * 2) / 3;
          circuito.esercizi.forEach((esercizio, index) => {
            const colIndex = index % 3;
            const xPos = margin + colIndex * colWidth;

            // Immagine
            if (esercizio.imagePreview) {
              doc.addImage(
                esercizio.imagePreview,
                this.getImageExtension(esercizio.imagePreview),
                xPos + 5,
                currentY,
                50,
                30
              );
            }
            doc.rect(xPos + 5, currentY, 50, 30);

            // Nome + serie x ripetizioni
            doc.setFontSize(9);
            doc.setFont('helvetica', 'bold');
            doc.text(
              `${esercizio.nome} - ${esercizio.serieRipetizioni}`,
              xPos + 30,
              currentY + 35,
              { align: 'center' }
            );

            // Recupero
            doc.setFontSize(8);
            doc.setFont('helvetica', 'normal');
            if (esercizio.recupero) {
              doc.text(
                `Recupero: ${esercizio.recupero}`,
                xPos + 30,
                currentY + 40,
                { align: 'center' }
              );
            }

            // Note
            if (esercizio.note) {
              doc.setFont('helvetica', 'italic');
              const noteLines = doc.splitTextToSize(esercizio.note, colWidth - 10);
              doc.text(noteLines, xPos + 30, currentY + 45, { align: 'center' });
            }

            // Vai a nuova riga ogni 3 esercizi
            if (colIndex === 2 || index === circuito.esercizi.length - 1) {
              currentY += 55;
              addPageIfNeeded(55);
            }
          });
          currentY += 5; // Spazio dopo ogni circuito
        });
      }
      /*
      // --- 3. GRIGLIA PESI PER OGNI CIRCUITO ---
      this.schedaData.circuiti.forEach((circuito, circuitoIndex) => {
        // Titolo del foglio griglia pesi
        addPageIfNeeded(20);
        doc.setFontSize(13);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(0, 0, 128);
        doc.text(`Circuito ${circuitoIndex + 1} - Griglia Pesi`, pageWidth / 2, currentY + 7, { align: 'center' });
        currentY += 10;

        // Parametri tabella
        const maxColonne = Math.max(circuito.esercizi.length, 10); // almeno 10 colonne per leggibilità
        const colonneEffettive = Math.max(circuito.esercizi.length, 10);
        const colWidth = (pageWidth - margin * 2) / colonneEffettive;
        const rowHeight = 6;
        const numSessioni = 38;

        // Header: nomi esercizi (o celle vuote)
        addPageIfNeeded(rowHeight);
        for (let col = 0; col < colonneEffettive; col++) {
          const x = margin + col * colWidth;
          doc.setFontSize(8);
          doc.setFont('helvetica', 'bold');
          doc.setDrawColor(0);
          doc.setLineWidth(0.2);
          doc.rect(x, currentY, colWidth, rowHeight);
          doc.text(
            col < circuito.esercizi.length ? circuito.esercizi[col].nome : '',
            x + colWidth / 2,
            currentY + rowHeight - 2,
            { align: 'center' }
          );
        }
        currentY += rowHeight;

        // Righe vuote per le sessioni
        for (let r = 0; r < numSessioni; r++) {
          addPageIfNeeded(rowHeight);
          for (let col = 0; col < colonneEffettive; col++) {
            const x = margin + col * colWidth;
            doc.setDrawColor(0);
            doc.setLineWidth(0.2);
            doc.rect(x, currentY, colWidth, rowHeight);
            // Lascia vuoto per la scrittura manuale
          }
          currentY += rowHeight;
        }

        // Spazio dopo la tabella
        currentY += 5;
      });
      */
      // --- 4. SALVATAGGIO DEL FILE ---
      const nomeFile = `Scheda_${this.schedaData.nomeCliente || 'Allenamento'}.pdf`;
      doc.save(nomeFile);

      console.log('PDF generato con successo usando jsPDF!');

    } catch (err) {
      console.error('Errore durante la generazione del PDF con jsPDF:', err);
      alert(`Errore generazione PDF: ${err instanceof Error ? err.message : 'errore sconosciuto'}`);
    }
  }


  // Il resto del codice rimane invariato...
  closePreview() {
    this.showPreviewModal = false;
  }

  calculateNoteHeight(noteText: string): number {
    if (!noteText || !noteText.trim()) {
      return 20;
    }
    const charactersPerLine = 40;
    const lineHeight = 20;
    const minHeight = 20;
    const maxHeight = 120;
    const estimatedLines = Math.ceil(noteText.length / charactersPerLine);
    const calculatedHeight = estimatedLines * lineHeight + 16;
    return Math.max(minHeight, Math.min(maxHeight, calculatedHeight));
  }
}
