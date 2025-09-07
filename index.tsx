/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
import { GoogleGenAI, Type } from "@google/genai";

// Assume API_KEY is set in the environment
const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxI5cJ9okwcWPGk6fm95OSxmGUxHSI39nGTWti4M-kRwIxVfIee_aLCLueY_gMhh0kKZg/exec';


// --- Google Sheet Sync Helper ---
const syncWithGoogleSheet = async (action: string, payload: object, retries = 2, delay = 1000) => {
    try {
        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            mode: 'cors',
            cache: 'no-cache',
            headers: {
                'Content-Type': 'text/plain',
            },
            redirect: 'follow', // Important for Apps Script
            body: JSON.stringify({ action, payload }),
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const textData = await response.text();
        try {
            const result = JSON.parse(textData);
            if (result.success === false) { // Check for explicit false
                throw new Error(result.error || 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุกับการซิงค์ Google Sheets');
            }
            return result.data; // For 'read' action, this will contain the documents
        } catch (e) {
            console.error("Failed to parse response from Google Sheet as JSON:", textData);
            throw new Error("ได้รับข้อมูลตอบกลับที่ไม่ถูกต้องจากเซิร์ฟเวอร์");
        }

    } catch (error) {
        const isFetchError = error instanceof TypeError && error.message === 'Failed to fetch';

        if (isFetchError && retries > 0) {
            console.warn(`Sync failed, retrying in ${delay / 1000}s... (${retries} retries left)`);
            await new Promise(resolve => setTimeout(resolve, delay));
            // Recurse with decremented retries and increased delay (exponential backoff)
            return syncWithGoogleSheet(action, payload, retries - 1, delay * 2);
        }

        console.error('Error syncing with Google Sheet:', error);
        if (isFetchError) {
            const detailedMessage = 'การซิงค์ข้อมูลล้มเหลว (Failed to fetch) กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ต และตรวจสอบว่าสิทธิ์ของ Google Sheets Web App ถูกตั้งค่าให้ "ทุกคน" เข้าถึงได้';
            throw new Error(detailedMessage);
        }
        // Re-throw other errors as they are.
        throw error;
    }
};


interface Document {
  id: string;
  docNumber: string;
  source: string;
  subject: string;
  docDate: string; // Stored as DD/MM/YYYY (BE)
  notes: string;
  tags: string[];
  createdAt: string;
  fileName?: string;
  fileContent?: string; // base64
  fileType?: string;
}

const app = document.getElementById('app') as HTMLDivElement;
let documents: Document[] = JSON.parse(localStorage.getItem('documents') || '[]');

// State for table sorting and re-rendering
let currentSort = { key: 'createdAt', direction: 'desc' };
let lastRenderedDocs: Document[] = [];
let lastRenderedTitle = '';


const saveDocuments = () => {
  localStorage.setItem('documents', JSON.stringify(documents));
  // After any change, re-render the current view
  if (document.getElementById('search-view')?.style.display !== 'none') {
    const searchInput = document.getElementById('search-input') as HTMLInputElement;
    if (searchInput.value.trim()) {
        handleSearch();
    } else {
        renderRecentDocuments();
    }
  }
};

const beDateToTimestamp = (beDateString: string): number => {
    if (!beDateString) return 0;
    const parts = beDateString.split('/');
    if (parts.length !== 3) return 0;
    const day = parts[0];
    const month = parts[1];
    const beYear = parseInt(parts[2], 10);
    const gregorianYear = beYear - 543;
    // JS months are 0-indexed
    return new Date(gregorianYear, parseInt(month, 10) - 1, parseInt(day, 10)).getTime();
};

const formatDateInput = (e: Event) => {
    const input = e.target as HTMLInputElement;
    let value = input.value.replace(/\D/g, ''); // Remove all non-digit characters
    if (value.length > 2) {
        value = `${value.slice(0, 2)}/${value.slice(2)}`;
    }
    if (value.length > 5) {
        // Limit the year to 4 digits
        value = `${value.slice(0, 5)}/${value.slice(5, 9)}`;
    }
    input.value = value;
};


const showToast = (message: string, type: 'success' | 'error' = 'success') => {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const textNode = document.createTextNode(message);
    toast.appendChild(textNode);
    
    container.appendChild(toast);

    // Animate in
    setTimeout(() => {
        toast.classList.add('show');
    }, 10); // small delay to allow CSS transition

    // Animate out and remove after 4 seconds
    setTimeout(() => {
        toast.classList.remove('show');
        toast.addEventListener('transitionend', () => {
            if (toast.parentElement) {
                toast.remove();
            }
        }, { once: true });
    }, 4000);
};

const setupMaxLengthWarning = (inputId: string) => {
    const input = document.getElementById(inputId) as HTMLInputElement;
    if (!input || input.maxLength <= 0) return;

    // Create a warning message element that is visually distinct but doesn't block submission
    const warningElement = document.createElement('p');
    warningElement.className = 'form-hint'; // Re-use existing styling for hints
    warningElement.style.color = 'var(--error-color)';
    warningElement.style.display = 'none'; // Hidden by default
    warningElement.setAttribute('aria-live', 'polite'); // For screen readers
    warningElement.textContent = `สามารถกรอกได้สูงสุด ${input.maxLength} ตัวอักษร`;

    // Insert after the input, before the main error message container
    input.parentNode?.insertBefore(warningElement, input.nextSibling);

    input.addEventListener('input', () => {
        // Show warning when length reaches max
        if (input.value.length >= input.maxLength) {
            warningElement.style.display = 'block';
        } else {
            // Hide it otherwise
            warningElement.style.display = 'none';
        }
    });
};

const handleFileChange = (e: Event) => {
    const input = e.target as HTMLInputElement;
    const displayArea = input.closest('.file-input-wrapper')?.querySelector('.file-display-area span');
    const errorElement = input.closest('.form-group')?.querySelector('.error-message');
    if (!displayArea || !errorElement) return;

    errorElement.textContent = ''; // Clear previous errors
    input.classList.remove('invalid');

    if (input.files && input.files.length > 0) {
        const file = input.files[0];
        const maxSize = 50 * 1024 * 1024; // 50 MB
        const allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'application/pdf'];

        if (file.size > maxSize) {
            errorElement.textContent = 'ขนาดไฟล์เกิน 50MB';
            input.classList.add('invalid');
            input.value = ''; // Clear the invalid file
            displayArea.textContent = 'ยังไม่ได้เลือกไฟล์';
            return;
        }

        if (!allowedTypes.includes(file.type)) {
            errorElement.textContent = 'ชนิดไฟล์ไม่ถูกต้อง (ต้องเป็นรูปภาพหรือ PDF)';
            input.classList.add('invalid');
            input.value = ''; // Clear the invalid file
            displayArea.textContent = 'ยังไม่ได้เลือกไฟล์';
            return;
        }

        displayArea.textContent = file.name;
    } else {
        displayArea.textContent = 'ยังไม่ได้เลือกไฟล์';
    }
};

const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            // result is "data:mime/type;base64,..."
            // We only want the base64 part
            const base64String = (reader.result as string).split(',')[1];
            resolve(base64String);
        };
        reader.onerror = error => reject(error);
    });
};

const downloadFileFromBase64 = (base64String: string, fileName: string, fileType: string) => {
    const byteCharacters = atob(base64String);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: fileType });
    const blobUrl = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(blobUrl);
};


// --- MODAL AND ACTION HANDLERS ---

const showViewModal = (docId: string) => {
  const doc = documents.find(d => d.id === docId);
  if (!doc) return;

  const modal = document.getElementById('view-modal') as HTMLDivElement;
  (document.getElementById('view-docNumber') as HTMLElement).textContent = doc.docNumber;
  (document.getElementById('view-docDate') as HTMLElement).textContent = doc.docDate;
  (document.getElementById('view-source') as HTMLElement).textContent = doc.source;
  (document.getElementById('view-subject') as HTMLElement).textContent = doc.subject;
  (document.getElementById('view-notes') as HTMLElement).textContent = doc.notes || '-';
  (document.getElementById('view-tags') as HTMLElement).innerHTML = doc.tags.map(tag => `<span class="tag">${tag}</span>`).join(' ');
  
  const attachmentContainer = document.getElementById('view-attachment-container') as HTMLDivElement;
  const attachmentContent = document.getElementById('view-attachment') as HTMLDivElement;
  if (doc.fileName && doc.fileContent && doc.fileType) {
      attachmentContent.innerHTML = `
          <span>${doc.fileName}</span>
          <button class="download-btn" data-id="${doc.id}">ดาวน์โหลด</button>
      `;
      attachmentContainer.style.display = 'block';
  } else {
      attachmentContainer.style.display = 'none';
  }

  modal.style.display = 'flex';
};

const showEditModal = (docId: string) => {
  const doc = documents.find(d => d.id === docId);
  if (!doc) return;

  const modal = document.getElementById('edit-modal') as HTMLDivElement;
  (document.getElementById('edit-docId') as HTMLInputElement).value = doc.id;
  (document.getElementById('edit-docNumber') as HTMLInputElement).value = doc.docNumber;
  (document.getElementById('edit-docDate') as HTMLInputElement).value = doc.docDate;
  (document.getElementById('edit-source') as HTMLInputElement).value = doc.source;
  (document.getElementById('edit-subject') as HTMLTextAreaElement).value = doc.subject;
  (document.getElementById('edit-notes') as HTMLInputElement).value = doc.notes;

  const fileDisplay = document.getElementById('edit-file-display') as HTMLSpanElement;
  const fileInput = document.getElementById('edit-attachment') as HTMLInputElement;
  fileInput.value = ''; // Reset file input
  if(doc.fileName) {
      fileDisplay.textContent = doc.fileName;
  } else {
      fileDisplay.textContent = 'ยังไม่ได้เลือกไฟล์';
  }

  modal.style.display = 'flex';
};


const handleUpdateDocument = async (e: Event) => {
    e.preventDefault();
    const form = e.target as HTMLFormElement;

    // --- Inline Validation ---
    let isValid = true;
    const errorMessages = form.querySelectorAll('.error-message');
    const invalidInputs = form.querySelectorAll('.invalid');
    errorMessages.forEach(msg => msg.textContent = '');
    invalidInputs.forEach(input => input.classList.remove('invalid'));

    const docNumberInput = form.elements.namedItem('edit-docNumber') as HTMLInputElement;
    const docDateInput = form.elements.namedItem('edit-docDate') as HTMLInputElement;
    const sourceInput = form.elements.namedItem('edit-source') as HTMLInputElement;
    const subjectInput = form.elements.namedItem('edit-subject') as HTMLTextAreaElement;
    const fileInput = form.elements.namedItem('edit-attachment') as HTMLInputElement;
    const dateRegex = /^\d{2}\/\d{2}\/\d{4}$/;


    if (!docNumberInput.value.trim()) {
        docNumberInput.classList.add('invalid');
        docNumberInput.nextElementSibling!.textContent = 'กรุณากรอกเลขที่หนังสือ';
        isValid = false;
    }
     if (!docDateInput.value.trim() || !dateRegex.test(docDateInput.value.trim())) {
        docDateInput.classList.add('invalid');
        docDateInput.nextElementSibling!.textContent = 'กรุณากรอกวันที่ให้ถูกต้อง (วว/ดด/ปปปป)';
        isValid = false;
    }
    if (!sourceInput.value.trim()) {
        sourceInput.classList.add('invalid');
        sourceInput.nextElementSibling!.textContent = 'กรุณากรอกที่มาหนังสือ';
        isValid = false;
    }
    if (!subjectInput.value.trim()) {
        subjectInput.classList.add('invalid');
        subjectInput.nextElementSibling!.textContent = 'กรุณากรอกเรื่องของหนังสือ';
        isValid = false;
    }
   
    if (!isValid) return;

    const docId = (form.elements.namedItem('edit-docId') as HTMLInputElement).value;
    const docIndex = documents.findIndex(d => d.id === docId);
    if (docIndex === -1) return;

    const saveBtn = form.querySelector('button[type="submit"]') as HTMLButtonElement;
    saveBtn.disabled = true;
    saveBtn.textContent = 'กำลังบันทึก...';
    
    const formattedDate = docDateInput.value.trim();

    const updatedDocData: Document = {
        ...documents[docIndex],
        docNumber: docNumberInput.value,
        docDate: formattedDate,
        source: sourceInput.value,
        subject: subjectInput.value,
        notes: (form.elements.namedItem('edit-notes') as HTMLInputElement).value,
    };
    
    // Handle file update
    const file = fileInput.files?.[0];
    if (file) {
        updatedDocData.fileName = file.name;
        updatedDocData.fileType = file.type;
        updatedDocData.fileContent = await fileToBase64(file);
    } // If no new file, the old file info remains from the spread operator

    // Remove legacy second file properties if they exist
    delete (updatedDocData as any).fileName2;
    delete (updatedDocData as any).fileContent2;
    delete (updatedDocData as any).fileType2;

    try {
        await syncWithGoogleSheet('update', updatedDocData);

        documents[docIndex] = updatedDocData;
        saveDocuments();
        (document.getElementById('edit-modal') as HTMLDivElement).style.display = 'none';
        form.reset();
        showToast('แก้ไขเอกสารเรียบร้อยแล้ว', 'success');

    } catch (error) {
        console.error("Error updating document:", error);
        const errorMessage = error instanceof Error ? error.message : 'เกิดข้อผิดพลาดในการซิงค์ข้อมูล';
        showToast(errorMessage, 'error');
    } finally {
        saveBtn.disabled = false;
        saveBtn.textContent = 'บันทึกการเปลี่ยนแปลง';
    }
};

const handleDeleteDocument = async (docId: string) => {
  const doc = documents.find(d => d.id === docId);
  if (!doc) return;
  
  if (confirm(`คุณแน่ใจหรือไม่ว่าต้องการลบเอกสารเรื่องของหนังสือ "${doc.subject}"?`)) {
    try {
        await syncWithGoogleSheet('delete', { id: docId });
        documents = documents.filter(d => d.id !== docId);
        saveDocuments();
        showToast('ลบเอกสารเรียบร้อยแล้ว', 'success');
    } catch (error) {
        console.error("Failed to delete document from sheet:", error);
        const errorMessage = error instanceof Error ? error.message : 'เกิดข้อผิดพลาดในการลบเอกสาร';
        showToast(errorMessage, 'error');
    }
  }
};

const renderResultsTable = (docs: Document[], title: string) => {
  lastRenderedDocs = docs;
  lastRenderedTitle = title;
  const resultsContainer = document.getElementById('search-results') as HTMLDivElement;
  if (!resultsContainer) return;

  if (docs.length === 0) {
    resultsContainer.innerHTML = `
      <h3>${title}</h3>
      <p class="placeholder">ไม่พบเอกสาร</p>
    `;
    return;
  }

  const sortedDocs = [...docs].sort((a: Document, b: Document) => {
    const key = currentSort.key as keyof Document;

    if (key === 'docDate') {
        const dateA = beDateToTimestamp(a.docDate);
        const dateB = beDateToTimestamp(b.docDate);
        return currentSort.direction === 'asc' ? dateA - dateB : dateB - dateA;
    }
     if (key === 'createdAt') {
       const dateA = new Date(a.createdAt).getTime();
       const dateB = new Date(b.createdAt).getTime();
       return currentSort.direction === 'asc' ? dateA - dateB : dateB - dateA;
    }

    const valA = (a[key] as string) || '';
    const valB = (b[key] as string) || '';

    return currentSort.direction === 'asc'
        ? valA.localeCompare(valB, 'th')
        : valB.localeCompare(valA, 'th');
  });

  const tableRows = sortedDocs.map(doc => `
      <tr>
        <td data-label="เลขที่หนังสือ">${doc.docNumber}</td>
        <td data-label="วันที่ลงหนังสือ">${doc.docDate}</td>
        <td data-label="ที่มาหนังสือ">${doc.source}</td>
        <td data-label="เรื่องของหนังสือ">${doc.subject}</td>
        <td data-label="ไฟล์แนบ">
            ${doc.fileName ? `<button class="download-btn" data-id="${doc.id}">ดาวน์โหลด</button>` : '<i>-</i>'}
        </td>
        <td data-label="การจัดการ">
          <div class="action-buttons">
            <button class="action-btn view-btn" data-id="${doc.id}" aria-label="ดูเอกสาร ${doc.subject}">ดู</button>
            <button class="action-btn edit-btn" data-id="${doc.id}" aria-label="แก้ไขเอกสาร ${doc.subject}">แก้ไข</button>
            <button class="action-btn delete-btn" data-id="${doc.id}" aria-label="ลบเอกสาร ${doc.subject}">ลบ</button>
          </div>
        </td>
      </tr>
    `).join('');

  const getSortAttr = (key: string) => {
    return currentSort.key === key ? `data-sort-direction="${currentSort.direction}"` : '';
  }

  resultsContainer.innerHTML = `
    <h3>${title}</h3>
    <div class="table-wrapper">
      <table class="results-table">
        <thead>
          <tr>
            <th class="sortable-header" data-sort-key="docNumber" ${getSortAttr('docNumber')}>เลขที่หนังสือ</th>
            <th class="sortable-header" data-sort-key="docDate" ${getSortAttr('docDate')}>วันที่ลงหนังสือ</th>
            <th class="sortable-header" data-sort-key="source" ${getSortAttr('source')}>ที่มาหนังสือ</th>
            <th class="sortable-header" data-sort-key="subject" ${getSortAttr('subject')}>เรื่องของหนังสือ</th>
            <th>ไฟล์แนบ</th>
            <th>การจัดการ</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
    </div>
  `;
};

const renderRecentDocuments = () => {
  const recentDocs = [...documents].sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()).slice(0, 10);
  if (documents.length === 0) {
     const resultsContainer = document.getElementById('search-results') as HTMLDivElement;
     if(resultsContainer) {
        resultsContainer.innerHTML = `<p class="placeholder">ยังไม่มีเอกสารในระบบ ไปที่แท็บ "จัดเก็บเอกสาร" เพื่อเพิ่มเอกสารใหม่</p>`;
     }
  } else {
    // Reset sort to default when showing recent docs
    currentSort = { key: 'createdAt', direction: 'desc' };
    renderResultsTable(recentDocs, "เอกสารล่าสุด 10 รายการ");
  }
};

const handleSearch = () => {
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const query = searchInput.value.toLowerCase().trim();

  if (!query) {
    renderRecentDocuments();
    return;
  }

  const results = documents.filter(doc =>
    doc.subject.toLowerCase().includes(query) ||
    doc.docNumber.toLowerCase().includes(query) ||
    doc.source.toLowerCase().includes(query) ||
    doc.notes.toLowerCase().includes(query) ||
    (doc.fileName && doc.fileName.toLowerCase().includes(query)) ||
    doc.tags.some(tag => tag.toLowerCase().includes(query))
  );

  renderResultsTable(results, "ผลการค้นหา");
};

const handleSaveDocument = async (e: Event) => {
  e.preventDefault();
  const form = e.target as HTMLFormElement;
  
  // --- Inline Validation ---
  let isValid = true;
  const errorMessages = form.querySelectorAll('.error-message');
  const invalidInputs = form.querySelectorAll('.invalid');
  errorMessages.forEach(msg => msg.textContent = '');
  invalidInputs.forEach(input => input.classList.remove('invalid'));

  const docNumberInput = form.elements.namedItem('docNumber') as HTMLInputElement;
  const docDateInput = form.elements.namedItem('docDate') as HTMLInputElement;
  const sourceInput = form.elements.namedItem('source') as HTMLInputElement;
  const subjectInput = form.elements.namedItem('subject') as HTMLTextAreaElement;
  const notesInput = form.elements.namedItem('notes') as HTMLInputElement;
  const fileInput = form.elements.namedItem('attachment') as HTMLInputElement;
  const dateRegex = /^\d{2}\/\d{2}\/\d{4}$/;


  if (!docNumberInput.value.trim()) {
      docNumberInput.classList.add('invalid');
      docNumberInput.nextElementSibling!.textContent = 'กรุณากรอกเลขที่หนังสือ';
      isValid = false;
  }
  if (!docDateInput.value.trim() || !dateRegex.test(docDateInput.value.trim())) {
      docDateInput.classList.add('invalid');
      docDateInput.nextElementSibling!.textContent = 'กรุณากรอกวันที่ให้ถูกต้อง (วว/ดด/ปปปป)';
      isValid = false;
  }
  if (!sourceInput.value.trim()) {
      sourceInput.classList.add('invalid');
      sourceInput.nextElementSibling!.textContent = 'กรุณากรอกที่มาหนังสือ';
      isValid = false;
  }
  if (!subjectInput.value.trim()) {
      subjectInput.classList.add('invalid');
      subjectInput.nextElementSibling!.textContent = 'กรุณากรอกเรื่องของหนังสือ';
      isValid = false;
  }

  if (!isValid) return;

  const saveBtn = document.getElementById('save-btn') as HTMLButtonElement;
  const btnText = saveBtn.querySelector('.btn-text') as HTMLSpanElement;
  const spinner = saveBtn.querySelector('.spinner') as HTMLDivElement;

  const startLoading = () => {
    btnText.textContent = 'กำลังประมวลผล...';
    spinner.classList.remove('hidden');
    saveBtn.disabled = true;
  };

  const stopLoading = (message: string, isError: boolean) => {
    showToast(message, isError ? 'error' : 'success');
    btnText.textContent = 'บันทึกเอกสาร';
    spinner.classList.add('hidden');
    saveBtn.disabled = false;
    if (!isError) {
      form.reset();
      // Also reset custom file input display
      const fileDisplay = form.querySelector('.file-display-area span');
      if (fileDisplay) fileDisplay.textContent = 'ยังไม่ได้เลือกไฟล์';
    }
  };
  
  startLoading();

  const formattedDate = docDateInput.value.trim();

  try {
    const prompt = `จากเรื่องของหนังสือและหมายเหตุต่อไปนี้ ช่วยสกัดคำสำคัญ (keywords) ที่เกี่ยวข้องมา 5-7 คำในภาษาไทย:\n\nเรื่องของหนังสือ: ${subjectInput.value}\nหมายเหตุ: ${notesInput.value}\n---`;

    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            keywords: {
              type: Type.ARRAY,
              items: { type: Type.STRING }
            }
          }
        },
      },
    });

    const resultJson = JSON.parse(response.text);
    const tags = resultJson.keywords || [];

    const newDoc: Document = {
      id: `doc_${Date.now()}`,
      docNumber: docNumberInput.value.trim(),
      docDate: formattedDate,
      source: sourceInput.value.trim(),
      subject: subjectInput.value.trim(),
      notes: notesInput.value.trim(),
      tags: tags,
      createdAt: new Date().toISOString()
    };
    
    const file = fileInput.files?.[0];
    if (file) {
        newDoc.fileName = file.name;
        newDoc.fileType = file.type;
        newDoc.fileContent = await fileToBase64(file);
    }
    
    await syncWithGoogleSheet('create', newDoc);
    
    documents.unshift(newDoc);
    saveDocuments();

    stopLoading('บันทึกเอกสารเรียบร้อยแล้ว!', false);

  } catch (error) {
    console.error('Error generating tags or saving document:', error);
    const errorMessage = error instanceof Error ? error.message : 'เกิดข้อผิดพลาดในการบันทึกเอกสาร';
    stopLoading(errorMessage, true);
  }
};

const addAppEventListeners = () => {
    const storeTabBtn = document.getElementById('store-tab-btn') as HTMLButtonElement;
    const searchTabBtn = document.getElementById('search-tab-btn') as HTMLButtonElement;
    const storeView = document.getElementById('store-view') as HTMLDivElement;
    const searchView = document.getElementById('search-view') as HTMLDivElement;
    
    // Setup real-time validation warnings
    setupMaxLengthWarning('docNumber');
    setupMaxLengthWarning('edit-docNumber');

    // Setup auto-formatting for date inputs
    document.getElementById('docDate')?.addEventListener('input', formatDateInput);
    document.getElementById('edit-docDate')?.addEventListener('input', formatDateInput);
    
    // Setup file input change listeners
    document.getElementById('attachment')?.addEventListener('change', handleFileChange);
    document.getElementById('edit-attachment')?.addEventListener('change', handleFileChange);

    const switchTab = (activeBtn: HTMLButtonElement, activeView: HTMLDivElement, inactiveBtn: HTMLButtonElement, inactiveView: HTMLDivElement) => {
        activeBtn.classList.add('active');
        activeView.style.display = 'block';
        inactiveBtn.classList.remove('active');
        inactiveView.style.display = 'none';
    };

    storeTabBtn.addEventListener('click', () => switchTab(storeTabBtn, storeView, searchTabBtn, searchView));
    searchTabBtn.addEventListener('click', () => {
      switchTab(searchTabBtn, searchView, storeTabBtn, storeView);
      handleSearch(); // Render recent docs when switching to search tab
    });

    // Save form submission
    const saveForm = document.getElementById('save-form') as HTMLFormElement;
    saveForm.addEventListener('submit', handleSaveDocument);
    
    // Search form submission
    const searchForm = document.getElementById('search-form') as HTMLFormElement;
    searchForm.addEventListener('submit', (e) => {
        e.preventDefault();
        handleSearch();
    });

    // Edit form submission
    const editForm = document.getElementById('edit-form') as HTMLFormElement;
    editForm.addEventListener('submit', handleUpdateDocument);

    // Close modals
    app.addEventListener('click', (e) => {
      const target = e.target as HTMLElement;
      if (target.classList.contains('modal-overlay') || target.classList.contains('close-btn')) {
        const modal = target.closest('.modal-overlay') as HTMLDivElement;
        if(modal) {
          modal.style.display = 'none';
        }
      }
    });

    // Event delegation for action buttons
    const resultsContainer = document.getElementById('search-results') as HTMLDivElement;
    resultsContainer.addEventListener('click', (e) => {
      const target = e.target as HTMLButtonElement;
      const docId = target.dataset.id;
      if (docId) {
        if (target.classList.contains('view-btn')) {
          showViewModal(docId);
        } else if (target.classList.contains('edit-btn')) {
          showEditModal(docId);
        } else if (target.classList.contains('delete-btn')) {
          handleDeleteDocument(docId);
        } else if (target.classList.contains('download-btn')) {
            const doc = documents.find(d => d.id === docId);
            if (doc && doc.fileContent && doc.fileName && doc.fileType) {
                downloadFileFromBase64(doc.fileContent, doc.fileName, doc.fileType);
            }
        }
      }
    });
    
    // Event delegation for modal downloads
    const viewModal = document.getElementById('view-modal');
    viewModal?.addEventListener('click', (e) => {
        const target = e.target as HTMLButtonElement;
        if (target.classList.contains('download-btn')) {
            const docId = target.dataset.id;
            const doc = documents.find(d => d.id === docId);
            if (doc && doc.fileContent && doc.fileName && doc.fileType) {
                downloadFileFromBase64(doc.fileContent, doc.fileName, doc.fileType);
            }
        }
    });

    // Event delegation for table sorting
    resultsContainer.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (target.classList.contains('sortable-header')) {
            const sortKey = target.dataset.sortKey;
            if (!sortKey) return;

            const currentDirection = target.dataset.sortDirection;
            if (currentSort.key === sortKey && currentDirection === 'desc') {
                currentSort.direction = 'asc';
            } else {
                currentSort.direction = 'desc';
            }
            currentSort.key = sortKey;
            
            renderResultsTable(lastRenderedDocs, lastRenderedTitle);
        }
    });

    // Synchronize with Google Sheet on initial load
    (async () => {
        const resultsContainer = document.getElementById('search-results') as HTMLDivElement;
        if(resultsContainer) {
            resultsContainer.innerHTML = `<p class="placeholder"><div class="spinner"></div>กำลังซิงค์ข้อมูลจาก Google Sheet...</p>`;
        }
        try {
            const sheetData = await syncWithGoogleSheet('read', {});
            documents = sheetData;
            saveDocuments();
            showToast('ซิงค์ข้อมูลสำเร็จ', 'success');
        } catch (error) {
            console.error("Initial sync failed, using local data:", error);
            const errorMessage = error instanceof Error ? error.message : 'เกิดข้อผิดพลาดในการซิงค์ข้อมูล';
            showToast(`${errorMessage}, แอปจะใช้ข้อมูลที่บันทึกไว้ล่าสุด`, 'error');
        } finally {
            if(document.getElementById('search-view')?.style.display !== 'none') {
                renderRecentDocuments();
            }
        }
    })();
}


const renderApp = () => {
  app.innerHTML = `
    <header>
        <img src="https://odc.mcu.ac.th/wp-content/uploads/2025/09/มจร-scaled.jpg" alt="Logo" class="header-logo">
        <div>
            <h1>ระบบจัดเก็บและสืบค้นเอกสาร</h1>
            <h2>วิทยาลัยพระธรรมทูต</h2>
        </div>
    </header>
    <main>
        <div class="tabs">
            <button id="store-tab-btn" class="tab-btn active">จัดเก็บเอกสาร</button>
            <button id="search-tab-btn" class="tab-btn">สืบค้นเอกสาร</button>
        </div>
        <div id="store-view" class="tab-content">
            <form id="save-form" novalidate>
                 <div class="form-row">
                    <div class="form-group form-group-doc-number">
                        <label for="docNumber">เลขที่หนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="docNumber" name="docNumber" required maxlength="50">
                        <div class="error-message"></div>
                    </div>
                    <div class="form-group form-group-date">
                        <label for="docDate">วันที่ลงหนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="docDate" name="docDate" placeholder="วว/ดด/ปปปป" required>
                         <div class="error-message"></div>
                    </div>
                    <div class="form-group form-group-fill">
                        <label for="source">ที่มาหนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="source" name="source" required>
                        <div class="error-message"></div>
                    </div>
                </div>
                <div class="form-group">
                    <label for="subject">เรื่องของหนังสือ<span class="required-asterisk">*</span></label>
                    <textarea id="subject" name="subject" rows="3" required></textarea>
                    <div class="error-message"></div>
                </div>
                 <div class="form-row">
                    <div class="form-group form-group-fill">
                         <label for="attachment">ไฟล์แนบ<span class="attachment-hint">(ขนาดไฟล์ไม่เกิน 50MB, รูปภาพ/PDF)</span></label>
                        <div class="file-input-wrapper">
                            <label for="attachment" class="file-input-label">เลือกไฟล์</label>
                            <input type="file" id="attachment" name="attachment" accept="image/*,application/pdf">
                            <div class="file-display-area"><span>ยังไม่ได้เลือกไฟล์</span></div>
                        </div>
                        <div class="error-message"></div>
                    </div>
                </div>
                <div class="form-group">
                    <label for="notes">หมายเหตุ</label>
                    <input type="text" id="notes" name="notes">
                    <div class="error-message"></div>
                </div>
                <div class="form-actions">
                    <button type="submit" id="save-btn">
                      <span class="btn-text">บันทึกเอกสาร</span>
                      <div class="spinner hidden"></div>
                    </button>
                </div>
            </form>
        </div>
        <div id="search-view" class="tab-content" style="display: none;">
            <form id="search-form">
                <div class="search-bar">
                    <input type="search" id="search-input" placeholder="ค้นหาจากเรื่องของหนังสือ, เลขที่, ที่มา, หมายเหตุ...">
                    <button type="submit" id="search-btn">ค้นหา</button>
                </div>
            </form>
            <div id="search-results">
                
            </div>
        </div>
    </main>
    
    <!-- View Modal -->
    <div id="view-modal" class="modal-overlay" style="display: none;">
      <div class="modal-content">
        <div class="modal-header">
          <h2>รายละเอียดเอกสาร</h2>
          <button class="close-btn" aria-label="ปิด">&times;</button>
        </div>
        <div class="modal-body">
          <p><strong>เลขที่หนังสือ:</strong> <span id="view-docNumber"></span></p>
          <p><strong>วันที่ลงหนังสือ:</strong> <span id="view-docDate"></span></p>
          <p><strong>ที่มาหนังสือ:</strong> <span id="view-source"></span></p>
          <p><strong>เรื่องของหนังสือ:</strong> <span id="view-subject"></span></p>
           <div id="view-attachment-container" style="display: none;">
             <p><strong>ไฟล์แนบ:</strong></p>
             <div id="view-attachment" class="attachment-display"></div>
           </div>
          <p><strong>หมายเหตุ:</strong> <span id="view-notes"></span></p>
          <hr>
          <p><strong>คำสำคัญ (Tags):</strong></p>
          <div id="view-tags" class="tags-container"></div>
        </div>
      </div>
    </div>

    <!-- Edit Modal -->
     <div id="edit-modal" class="modal-overlay" style="display: none;">
      <div class="modal-content">
        <form id="edit-form" novalidate>
            <div class="modal-header">
              <h2>แก้ไขเอกสาร</h2>
              <button type="button" class="close-btn" aria-label="ปิด">&times;</button>
            </div>
            <div class="modal-body">
                <input type="hidden" id="edit-docId" name="edit-docId">
                <div class="form-row">
                    <div class="form-group form-group-doc-number">
                        <label for="edit-docNumber">เลขที่หนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="edit-docNumber" name="edit-docNumber" required maxlength="50">
                        <div class="error-message"></div>
                    </div>
                    <div class="form-group form-group-date">
                        <label for="edit-docDate">วันที่ลงหนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="edit-docDate" name="edit-docDate" placeholder="วว/ดด/ปปปป" required>
                        <div class="error-message"></div>
                    </div>
                    <div class="form-group form-group-fill">
                        <label for="edit-source">ที่มาหนังสือ<span class="required-asterisk">*</span></label>
                        <input type="text" id="edit-source" name="edit-source" required>
                        <div class="error-message"></div>
                    </div>
                </div>
                <div class="form-group">
                    <label for="edit-subject">เรื่องของหนังสือ<span class="required-asterisk">*</span></label>
                    <textarea id="edit-subject" name="edit-subject" rows="3" required></textarea>
                    <div class="error-message"></div>
                </div>
                 <div class="form-row">
                    <div class="form-group form-group-fill">
                         <label for="edit-attachment">ไฟล์แนบ<span class="attachment-hint">(ขนาดไฟล์ไม่เกิน 50MB, รูปภาพ/PDF)</span></label>
                        <div class="file-input-wrapper">
                            <label for="edit-attachment" class="file-input-label">เลือกไฟล์ใหม่</label>
                            <input type="file" id="edit-attachment" name="edit-attachment" accept="image/*,application/pdf">
                            <div class="file-display-area"><span id="edit-file-display">ยังไม่ได้เลือกไฟล์</span></div>
                        </div>
                        <div class="error-message"></div>
                    </div>
                </div>
                <div class="form-group">
                    <label for="edit-notes">หมายเหตุ</label>
                    <input type="text" id="edit-notes" name="edit-notes">
                    <div class="error-message"></div>
                </div>
            </div>
            <div class="modal-footer">
                <button type="submit">บันทึกการเปลี่ยนแปลง</button>
            </div>
        </form>
      </div>
    </div>
  `;
  addAppEventListeners();
};


// Initial Render
renderApp();