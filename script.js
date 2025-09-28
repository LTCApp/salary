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
        
        // تحليل الأسماء
        const namesAnalysis = this.analyzeNames();
        
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
        this.namesResults = namesAnalysis;
    }

    analyzeNames() {
        const nameFrequency = {};
        const nameData = {};
        
        // تجميع الأسماء وحساب تكرارها
        this.data.forEach(row => {
            // استخراج الاسم من الملاحظات
            const name = this.extractNameFromNotes(row.notes);
            if (name) {
                // حساب التكرار
                nameFrequency[name] = (nameFrequency[name] || 0) + 1;
                
                // تجميع البيانات المالية
                if (!nameData[name]) {
                    nameData[name] = { debit: 0, credit: 0, net: 0, count: 0 };
                }
                
                nameData[name].debit += row.debit;
                nameData[name].credit += row.credit;
                nameData[name].count += 1;
            }
        });
        
        // حساب الصافي لكل اسم
        Object.keys(nameData).forEach(name => {
            nameData[name].net = nameData[name].debit - nameData[name].credit;
        });
        
        // تحديد الاسم الأكثر تكراراً (الأصلي)
        let originalName = '';
        let maxFrequency = 0;
        
        Object.entries(nameFrequency).forEach(([name, freq]) => {
            if (freq > maxFrequency) {
                maxFrequency = freq;
                originalName = name;
            }
        });
        
        // تحديد الأسماء المختلفة
        const differentNames = Object.keys(nameData).filter(name => name !== originalName);
        
        return {
            originalName,
            differentNames,
            nameData,
            nameFrequency
        };
    }
    
    extractNameFromNotes(notes) {
        if (!notes || notes.trim() === '') return null;
        
        // إزالة الكلمات الزائدة واستخراج الاسم
        let cleanNotes = notes.trim();
        
        // إزالة الكلمات الشائعة مثل "عجز" و "خدمة"
        const wordsToRemove = ['عجز', 'خدمة', 'خدمات', 'مدين', 'دائن', 'سلف', 'مشتريات'];
        
        wordsToRemove.forEach(word => {
            cleanNotes = cleanNotes.replace(new RegExp(word, 'gi'), '');
        });
        
        // إزالة الفواصل والرقام والرموز غير الضرورية
        cleanNotes = cleanNotes.replace(/[0-9\.,\-\/\\:;(){}\[\]]/g, ' ');
        
        // إزالة الفراغات الزائدة
        cleanNotes = cleanNotes.replace(/\s+/g, ' ').trim();
        
        // التحقق من وجود اسم (على الأقل كلمتين)
        const words = cleanNotes.split(' ').filter(word => word.length > 1);
        
        if (words.length >= 2) {
            // أخذ أول كلمتين كاسم
            return words.slice(0, 2).join(' ');
        } else if (words.length === 1 && words[0].length >= 3) {
            // إذا كانت كلمة واحدة وطويلة
            return words[0];
        }
        
        return null;
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
        
        // عرض نتائج الأسماء
        this.showNamesResults();
        
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
    
    showNamesResults() {
        if (!this.namesResults || !this.namesResults.nameData) return;
        
        const { originalName, differentNames, nameData } = this.namesResults;
        const namesContainer = document.getElementById('namesContainer');
        const namesSection = document.getElementById('namesSection');
        
        // مسح المحتوى السابق
        namesContainer.innerHTML = '';
        
        // إذا لم يتم العثور على أسماء، إخفاء هذا القسم
        if (!originalName && differentNames.length === 0) {
            namesSection.style.display = 'none';
            return;
        }
        
        // عرض قسم الأسماء
        namesSection.style.display = 'block';
        
        // عرض الاسم الأصلي
        if (originalName && nameData[originalName]) {
            const nameBox = this.createNameBox(originalName, nameData[originalName], true);
            namesContainer.appendChild(nameBox);
        }
        
        // عرض الأسماء المختلفة
        differentNames.forEach(name => {
            if (nameData[name]) {
                const nameBox = this.createNameBox(name, nameData[name], false);
                namesContainer.appendChild(nameBox);
            }
        });
    }
    
    createNameBox(name, data, isOriginal) {
        const nameBox = document.createElement('div');
        nameBox.className = `name-box ${isOriginal ? 'original-name' : 'different-name'}`;
        
        nameBox.innerHTML = `
            <h4>${name} ${isOriginal ? '(الاسم الأصلي)' : '(اسم مختلف)'}</h4>
            <div class="calculation-grid">
                <div class="calc-item">
                    <span class="label">مدين:</span>
                    <span class="value">${this.formatNumber(data.debit)}</span>
                </div>
                <div class="calc-item">
                    <span class="label">دائن:</span>
                    <span class="value">${this.formatNumber(data.credit)}</span>
                </div>
                <div class="calc-item total">
                    <span class="label">الصافي:</span>
                    <span class="value">${this.formatNumber(data.net)}</span>
                </div>
                <div class="calc-item">
                    <span class="label">عدد العمليات:</span>
                    <span class="value">${data.count}</span>
                </div>
            </div>
        `;
        
        return nameBox;
    }
}

// تشغيل التطبيق عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', () => {
    new FinancialAnalyzer();
});