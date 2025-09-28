class FinancialAnalyzer {
    constructor() {
        this.data = [];
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('excelFile');
        const processBtn = document.getElementById('processBtn');
        
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        processBtn.addEventListener('click', () => this.processData());
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        const fileInfo = document.getElementById('fileInfo');
        const processBtn = document.getElementById('processBtn');
        
        if (file) {
            fileInfo.textContent = `تم اختيار: ${file.name}`;
            processBtn.disabled = false;
            this.selectedFile = file;
        } else {
            fileInfo.textContent = 'لم يتم اختيار ملف';
            processBtn.disabled = true;
        }
    }

    async processData() {
        if (!this.selectedFile) return;
        
        this.showLoading(true);
        
        try {
            const data = await this.readExcelFile(this.selectedFile);
            this.data = data;
            this.analyzeData();
            this.showResults();
        } catch (error) {
            alert('حدث خطأ في قراءة الملف: ' + error.message);
        }
        
        this.showLoading(false);
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    resolve(this.parseExcelData(jsonData));
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('فشل في قراءة الملف'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseExcelData(rawData) {
        if (rawData.length < 2) {
            throw new Error('الملف يجب أن يحتوي على بيانات');
        }
        
        const headers = rawData[0].map(h => String(h).trim());
        
        // البحث عن الأعمدة المطلوبة
        const debitIndex = this.findColumnIndex(headers, ['المدين', 'مدين', 'debit']);
        const creditIndex = this.findColumnIndex(headers, ['الدائن', 'دائن', 'credit']);
        const notesIndex = this.findColumnIndex(headers, ['الملاحظات', 'ملاحظات', 'notes', 'description']);
        
        if (debitIndex === -1 || creditIndex === -1 || notesIndex === -1) {
            throw new Error('لم يتم العثور على الأعمدة المطلوبة (المدين، الدائن، الملاحظات)');
        }
        
        const parsedData = [];
        
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            
            const debit = this.parseNumber(row[debitIndex]);
            const credit = this.parseNumber(row[creditIndex]);
            const notes = String(row[notesIndex] || '').trim();
            
            if (debit !== 0 || credit !== 0 || notes !== '') {
                parsedData.push({
                    debit,
                    credit,
                    notes
                });
            }
        }
        
        return parsedData;
    }

    findColumnIndex(headers, searchTerms) {
        for (let term of searchTerms) {
            const index = headers.findIndex(h => 
                h.toLowerCase().includes(term.toLowerCase())
            );
            if (index !== -1) return index;
        }
        return -1;
    }

    parseNumber(value) {
        if (value === null || value === undefined || value === '') return 0;
        
        // تحويل إلى نص أولاً
        let str = String(value).trim();
        
        // إزالة الفواصل والرموز غير الضرورية
        str = str.replace(/[^0-9.-]/g, '');
        
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    analyzeData() {
        const results = {
            deficit: { debit: 0, credit: 0, net: 0 },
            service: { debit: 0, credit: 0, net: 0 },
            advances: { debit: 0, credit: 0, net: 0 }
        };
        
        this.data.forEach(row => {
            const notes = row.notes.toLowerCase();
            
            // تحليل العجز
            if (notes.includes('عجز')) {
                results.deficit.debit += row.debit;
                results.deficit.credit += row.credit;
            }
            
            // تحليل الخدمات
            if (notes.includes('خدمة') || notes.includes('خدمات')) {
                results.service.debit += row.debit;
                results.service.credit += row.credit;
            }
            
            // السلف والمشتريات (كل ما لا يحتوي على عجز أو خدمة)
            if (!notes.includes('عجز') && !notes.includes('خدمة') && !notes.includes('خدمات')) {
                results.advances.debit += row.debit;
                results.advances.credit += row.credit;
            }
        });
        
        // حساب الصافي
        results.deficit.net = results.deficit.debit - results.deficit.credit;
        results.service.net = results.service.debit - results.service.credit;
        results.advances.net = results.advances.debit - results.advances.credit;
        
        this.results = results;
    }

    showResults() {
        // عرض نتائج العجز
        document.getElementById('deficitDebit').textContent = this.formatNumber(this.results.deficit.debit);
        document.getElementById('deficitCredit').textContent = this.formatNumber(this.results.deficit.credit);
        document.getElementById('deficitNet').textContent = this.formatNumber(this.results.deficit.net);
        
        // عرض نتائج الخدمات
        document.getElementById('serviceDebit').textContent = this.formatNumber(this.results.service.debit);
        document.getElementById('serviceCredit').textContent = this.formatNumber(this.results.service.credit);
        document.getElementById('serviceNet').textContent = this.formatNumber(this.results.service.net);
        
        // عرض نتائج السلف والمشتريات
        document.getElementById('advancesDebit').textContent = this.formatNumber(this.results.advances.debit);
        document.getElementById('advancesCredit').textContent = this.formatNumber(this.results.advances.credit);
        document.getElementById('advancesNet').textContent = this.formatNumber(this.results.advances.net);
        
        // عرض الملخص النهائي
        document.getElementById('totalAdvances').textContent = this.formatNumber(this.results.advances.net);
        document.getElementById('totalServices').textContent = this.formatNumber(this.results.service.net);
        document.getElementById('totalDeficit').textContent = this.formatNumber(this.results.deficit.net);
        
        const grandTotal = this.results.advances.net + this.results.service.net + this.results.deficit.net;
        document.getElementById('grandTotal').textContent = this.formatNumber(grandTotal);
        
        // إظهار قسم النتائج
        document.getElementById('resultsSection').style.display = 'block';
        
        // التمرير إلى النتائج
        document.getElementById('resultsSection').scrollIntoView({ 
            behavior: 'smooth' 
        });
    }

    formatNumber(num) {
        return new Intl.NumberFormat('ar-EG', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    showLoading(show) {
        const loading = document.getElementById('loading');
        const resultsSection = document.getElementById('resultsSection');
        
        if (show) {
            loading.style.display = 'block';
            resultsSection.style.display = 'none';
        } else {
            loading.style.display = 'none';
        }
    }
}

// تشغيل التطبيق عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', () => {
    new FinancialAnalyzer();
});