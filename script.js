const STORAGE_KEY = 'patientManagerData';
    const themeKey = 'patientManagerTheme';
    const FIREBASE_COLLECTION = 'patientManager';
    const FIREBASE_DOC_ID = 'appData';
    let patients = [];
    let currentSort = { field: 'patientId', direction: 'asc' };
    let currentPatientId = null;
    let currentResultsPatientId = null;
    let reminderTableFilter = null;
    let patientTableFilter = null;
    let reminderSearchFilter = '';
    let reminderStatusFilter = 'all';
    let selectedPatientIds = new Set();
    let pendingReceptionPatientId = null;

    const receptionTestDefinitions = {
      hematology: ['wbc','neu','lym','mono','eos','baso','neuPercent','lymPercent','monoPercent','eosPercent','basoPercent','rbc','hgb','hct','mcv','mch','mchc','rdw','rdwSd','plt','mpv','pdw','pct','plcr','plcc'],
      biochemistry: ['glucose','hba1c','ure','creatinine','ast','alt','ggt','bilirubin','cholesterol','triglyceride','hdl','ldl','uricAcid','totalProtein','albumin','globulin','agRatio','ldh','calcium','phosphate','magnesium'],
      immunology: ['hbsag','antiHbs','hcv','hiv','afp','cea','psa','ca125','ca199','ca153','crp','rf','ana','igg','igm','iga','tsh','t3','t4','ft3','ft4','cortisol'],
      coagulation: ['pt_s','pt_percent','inr','aptt','fibrinogen'],
      urine: ['leu','nit','ubg','bil','pro','ph','bld','ket','glu','sg']
    };

    async function saveDataToFirebase() {
      try {
        if (!window.firebaseDb || !window.firebaseFirestoreImports || !window.firebaseAuth || !window.firebaseAuth.currentUser) return;
        const { doc, setDoc } = window.firebaseFirestoreImports;
        const db = window.firebaseDb;
        const docRef = doc(db, FIREBASE_COLLECTION, FIREBASE_DOC_ID);
        await setDoc(docRef, {
          patients,
          updatedAt: new Date().toISOString()
        });
      } catch (error) {
        console.warn('Firebase sync failed:', error);
      }
    }

    async function loadDataFromFirebase() {
      try {
        if (!window.firebaseDb || !window.firebaseFirestoreImports || !window.firebaseAuth || !window.firebaseAuth.currentUser) return null;
        const { doc, getDoc } = window.firebaseFirestoreImports;
        const db = window.firebaseDb;
        const docRef = doc(db, FIREBASE_COLLECTION, FIREBASE_DOC_ID);
        const snap = await getDoc(docRef);
        if (!snap.exists()) return null;
        const data = snap.data();
        return Array.isArray(data?.patients) ? data.patients : null;
      } catch (error) {
        console.warn('Firebase load failed:', error);
        return null;
      }
    }

    async function saveData() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(patients));
      await saveDataToFirebase();
    }

    const samplePatients = [
      {
        patientId: 'BN001', fullName: 'Nguyễn Văn A', birthDate: '1986-04-21', gender: 'Nam', healthInsurance: 'BHYT12345', cccd: '012345678901', phone: '0912345678', ethnicity: 'Kinh', job: 'Nhân viên', address: 'Hà Nội', note: 'Tiền sử dị ứng penicillin.', lastExamDate: '2024-05-22', nextAppointment: '2024-06-20', clinical: { height: '172', weight: '70', bmi: '23.7', bp: '120/80', pulse: '72', temp: '36.5', spo2: '98' }, hematology: { wbc: '8.96', neu: '7.04', lym: '1.43', mono: '0.42', eos: '0.05', baso: '0.02', neuPercent: '78.6', lymPercent: '16.0', monoPercent: '4.7', eosPercent: '0.5', basoPercent: '0.2', rbc: '5.23', hgb: '100', hct: '31.1', mcv: '59.5', mch: '19.1', mchc: '320', rdw: '14.0', rdwSd: '29.4', plt: '286', mpv: '11.0', pdw: '20.4', pct: '0.314', plcr: '27.3', plcc: '78' }, biochemistry: { glucose: '5.2', hba1c: '5.4', ure: '5.6', creatinine: '0.9', ast: '22', alt: '20', ggt: '28', bilirubin: '12', cholesterol: '4.8', triglyceride: '1.4', hdl: '1.1', ldl: '2.7', uricAcid: '6.1' }, immunology: { hbsag: 'Âm tính', antiHbs: 'Dương tính', hcv: 'Âm tính', hiv: 'Âm tính', afp: '4.5', cea: '1.8', psa: '0.9', ca125: '12', ca199: '15', ca153: '14' }, conclusion: { diagnosis: 'Tăng huyết áp độ 1', treatment: 'Theo dõi huyết áp, thay đổi lối sống', note: 'Khuyến nghị tập thể dục 30 phút mỗi ngày.' }, history: [ { date: '2024-05-22', diagnosis: 'Khám định kỳ', result: 'Ổn định', note: 'Huyết áp 130/85.' } ], reminders: [ { examDate: '2024-05-22', revisitDate: '2024-06-20', content: 'Tái khám huyết áp', note: 'Nhắc uống thuốc đều.' } ]
      }
    ];

    function normalizeExcelDate(value) {
      if (typeof value === 'number') {
        const date = XLSX.SSF.parse_date_code(value);
        if (!date) return '';
        return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
      }
      return value ? value.toString().trim() : '';
    }

    const testFieldDefinitionsKey = 'testFieldDefinitions';
    let testFieldDefinitions = loadTestFieldDefinitions();
    let testDefinitionEditCategory = null;
    let testDefinitionEditIndex = -1;

    function getExcelField(row, keys) {
      for (const key of keys) {
        if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
          return row[key];
        }
      }
      return '';
    }

    function loadTestFieldDefinitions() {
      const stored = localStorage.getItem(testFieldDefinitionsKey);
      if (stored) {
        try {
          return JSON.parse(stored);
        } catch (error) {
          console.warn('Không thể đọc cấu hình xét nghiệm từ localStorage', error);
        }
      }
      return {
        hematology: [],
        biochemistry: [],
        immunology: []
      };
    }

    function saveTestFieldDefinitions() {
      localStorage.setItem(testFieldDefinitionsKey, JSON.stringify(testFieldDefinitions));
    }

    function sanitizeFieldKey(value) {
      return String(value || '').trim().replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '');
    }

    function renderTestDefinitionsTable() {
      const category = document.getElementById('testDefinitionCategorySelect').value;
      const body = document.getElementById('testDefinitionsTableBody');
      const items = testFieldDefinitions[category] || [];
      body.innerHTML = items.length ? items.map((item, index) => `
        <tr>
          <td>${item.key}</td>
          <td>${item.title}</td>
          <td>${item.unit || ''}</td>
          <td>
            <button type="button" class="btn btn-sm btn-outline-primary me-1" onclick="editTestDefinition('${category}', ${index})">Sửa</button>
            <button type="button" class="btn btn-sm btn-outline-danger" onclick="deleteTestDefinition('${category}', ${index})">Xóa</button>
          </td>
        </tr>
      `).join('') : '<tr><td colspan="4" class="text-muted">Chưa có xét nghiệm tuỳ chỉnh</td></tr>';
    }

    function clearTestDefinitionForm() {
      testDefinitionEditCategory = null;
      testDefinitionEditIndex = -1;
      document.getElementById('testDefinitionKey').value = '';
      document.getElementById('testDefinitionTitle').value = '';
      document.getElementById('testDefinitionUnit').value = '';
      document.getElementById('cancelTestDefinitionEditButton').style.display = 'none';
      document.getElementById('saveTestDefinitionButton').textContent = 'Lưu xét nghiệm';
    }

    function editTestDefinition(category, index) {
      const item = (testFieldDefinitions[category] || [])[index];
      if (!item) return;
      testDefinitionEditCategory = category;
      testDefinitionEditIndex = index;
      document.getElementById('testDefinitionCategorySelect').value = category;
      document.getElementById('testDefinitionKey').value = item.key;
      document.getElementById('testDefinitionTitle').value = item.title;
      document.getElementById('testDefinitionUnit').value = item.unit || '';
      document.getElementById('cancelTestDefinitionEditButton').style.display = 'inline-block';
      document.getElementById('saveTestDefinitionButton').textContent = 'Cập nhật xét nghiệm';
      renderTestDefinitionsTable();
    }

    function deleteTestDefinition(category, index) {
      if (!confirm('Xoá xét nghiệm này?')) return;
      testFieldDefinitions[category].splice(index, 1);
      saveTestFieldDefinitions();
      if (testDefinitionEditCategory === category && testDefinitionEditIndex === index) {
        clearTestDefinitionForm();
      }
      renderTestDefinitionsTable();
    }

    function saveTestDefinition() {
      const category = document.getElementById('testDefinitionCategorySelect').value;
      const rawKey = document.getElementById('testDefinitionKey').value;
      const key = sanitizeFieldKey(rawKey);
      const title = document.getElementById('testDefinitionTitle').value.trim();
      const unit = document.getElementById('testDefinitionUnit').value.trim();
      if (!key || !title) {
        alert('Vui lòng nhập cả mã trường và tên hiển thị.');
        return;
      }
      const items = testFieldDefinitions[category] || [];
      const duplicateIndex = items.findIndex((item, idx) => item.key === key && idx !== testDefinitionEditIndex);
      if (duplicateIndex !== -1) {
        alert('Mã trường đã tồn tại trong nhóm này.');
        return;
      }
      const definition = { key, title, unit };
      if (testDefinitionEditIndex >= 0 && testDefinitionEditCategory === category) {
        items[testDefinitionEditIndex] = definition;
      } else {
        items.push(definition);
      }
      testFieldDefinitions[category] = items;
      saveTestFieldDefinitions();
      clearTestDefinitionForm();
      renderTestDefinitionsTable();
    }

    function renderCustomFieldsForCategory(category, containerId, data = {}) {
      const container = document.getElementById(containerId);
      if (!container) return;
      const defs = testFieldDefinitions[category] || [];
      if (!defs.length) {
        container.innerHTML = '';
        return;
      }
      container.innerHTML = defs.map(def => {
        const value = (data && data[def.key]) || '';
        return `
          <div class="col-md-3">
            <label class="form-label">${def.title}</label>
            <input id="${category}Custom_${def.key}" class="form-control" value="${String(value).replace(/"/g, '&quot;')}" />
          </div>
          <div class="col-md-3">
            <label class="form-label">Đơn vị</label>
            <input class="form-control" disabled value="${def.unit || ''}" />
          </div>
        `;
      }).join('');
    }

    const normalValueRanges = {
      hematologyWbc: '4.0-10.0',
      hematologyNeu: '2.0-7.5',
      hematologyLym: '1.0-4.0',
      hematologyMono: '0.2-0.8',
      hematologyEos: '0.0-0.5',
      hematologyBaso: '0.0-0.2',
      hematologyNeuPercent: '40-75',
      hematologyLymPercent: '20-45',
      hematologyMonoPercent: '2-8',
      hematologyEosPercent: '0.0-5',
      hematologyBasoPercent: '0.0-2',
      hematologyRbc: '4.0-5.5',
      hematologyHgb: '130-170',
      hematologyHct: '40-50',
      hematologyMcv: '80-100',
      hematologyMch: '27-33',
      hematologyMchc: '320-360',
      hematologyRdw: '11.5-14.5',
      hematologyRdwSd: '35-56',
      hematologyPlt: '150-450',
      hematologyMpv: '7.5-11.5',
      hematologyPdw: '9-17',
      hematologyPct: '0.19-0.39',
      hematologyPlcr: '15-35',
      hematologyPlcc: '15-35',
      biochemistryGlucose: '70-99',
      biochemistryHba1c: '<5.7',
      biochemistryUre: '7-20',
      biochemistryCreatinine: '0.7-1.2',
      biochemistryAst: '10-40',
      biochemistryAlt: '7-56',
      biochemistryGgt: '9-48',
      biochemistryBilirubin: '0.1-1.2',
      biochemistryCholesterol: '<200',
      biochemistryTriglyceride: '<150',
      biochemistryHdl: '>40',
      biochemistryLdl: '<100',
      biochemistryUricAcid: '3.5-7.2',
      biochemistryTotalProtein: '6.0-8.3',
      biochemistryAlbumin: '3.5-5.0',
      biochemistryGlobulin: '2.0-3.5',
      biochemistryAgRatio: '1.0-2.2',
      biochemistryLdh: '140-280',
      biochemistryCalcium: '8.4-10.2',
      biochemistryPhosphate: '2.5-4.5',
      biochemistryMagnesium: '1.7-2.2',
      immunologyAfp: '<10',
      immunologyCea: '<5',
      immunologyPsa: '<4',
      immunologyCa125: '<35',
      immunologyCa199: '<37',
      immunologyCa153: '<30',
      immunologyCrp: '<5',
      immunologyRf: '<14',
      immunologyIgg: '700-1600',
      immunologyIgm: '40-230',
      immunologyIga: '70-400',
      immunologyTsh: '0.4-4.0',
      immunologyT3: '0.8-2.0',
      immunologyT4: '5-12',
      immunologyFt3: '2.3-4.2',
      immunologyFt4: '0.8-1.8',
      immunologyCortisol: '5-25'
    };

    function parseRange(range) {
      if (!range) return {};
      const source = String(range).trim();
      const between = source.match(/^([\d.,]+)\s*[-–]\s*([\d.,]+)$/);
      if (between) {
        return { min: parseFloat(between[1].replace(',', '.')), max: parseFloat(between[2].replace(',', '.')) };
      }
      const lessEqual = source.match(/^(?:<=|≤)\s*([\d.,]+)$/);
      if (lessEqual) {
        return { max: parseFloat(lessEqual[1].replace(',', '.')) };
      }
      const greaterEqual = source.match(/^(?:>=|≥)\s*([\d.,]+)$/);
      if (greaterEqual) {
        return { min: parseFloat(greaterEqual[1].replace(',', '.')) };
      }
      const less = source.match(/^<\s*([\d.,]+)$/);
      if (less) {
        return { max: parseFloat(less[1].replace(',', '.')) };
      }
      const greater = source.match(/^>\s*([\d.,]+)$/);
      if (greater) {
        return { min: parseFloat(greater[1].replace(',', '.')) };
      }
      return {};
    }

    function getResultColor(value, range) {
      if (!value || !range) return '';
      const numeric = parseFloat(String(value).replace(',', '.'));
      if (Number.isNaN(numeric)) return '';
      const parsed = parseRange(range);
      if (parsed.min != null && parsed.max != null) {
        if (numeric < parsed.min) return 'blue';
        if (numeric > parsed.max) return 'red';
      } else if (parsed.max != null) {
        if (numeric > parsed.max) return 'red';
      } else if (parsed.min != null) {
        if (numeric < parsed.min) return 'blue';
      }
      return '';
    }

    function setResultInputColor(input) {
      if (!input) return;
      const range = normalValueRanges[input.id] || '';
      const color = getResultColor(input.value.trim(), range);
      input.style.color = color || '';
    }

    function importPatientsFromExcel(file) {
      const reader = new FileReader();
      reader.onload = async event => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames.length) {
          alert('Không tìm thấy sheet trong file Excel.');
          return;
        }
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (!rows.length) {
          alert('Không tìm thấy dữ liệu trong file Excel.');
          return;
        }

        let updatedCount = 0;
        let newCount = 0;
        rows.forEach(raw => {
          const row = Object.fromEntries(Object.entries(raw).map(([key, value]) => [key.toString().trim(), value]));
          const patientId = getExcelField(row, ['Mã BN', 'patientId', 'Patient ID', 'ID']).toString().trim();
          if (!patientId) return;
          const existing = patients.find(p => p.patientId === patientId);

          const revisitDateVal = normalizeExcelDate(getExcelField(row, ['Ngày tái khám', 'revisitDate', 'Revisit Date', 'RevisitDate', 'revisit_date', 'examDate', 'Ngày hẹn']));
          const revisitContent = getExcelField(row, ['Nội dung', 'content', 'reminderContent', 'Reminder Content', 'Content']);
          const revisitNote = getExcelField(row, ['Ghi chú nhắc', 'reminderNote', 'reminder_note', 'note', 'Note']);

          const normalized = {
            patientId,
            fullName: getExcelField(row, ['Họ tên', 'fullName', 'Full Name']),
            birthDate: normalizeExcelDate(getExcelField(row, ['Ngày sinh', 'birthDate', 'Birth Date'])),
            gender: getExcelField(row, ['Giới tính', 'gender']),
            healthInsurance: getExcelField(row, ['Mã BHYT', 'healthInsurance', 'Insurance']),
            cccd: getExcelField(row, ['CCCD', 'cccd', 'Citizen ID']),
            phone: getExcelField(row, ['SĐT', 'phone', 'Phone']),
            ethnicity: getExcelField(row, ['Dân tộc', 'ethnicity']),
            job: getExcelField(row, ['Nghề nghiệp', 'job', 'Job']),
            address: getExcelField(row, ['Địa chỉ', 'address', 'Address']),
            note: getExcelField(row, ['Ghi chú', 'note', 'Note']),
            lastExamDate: normalizeExcelDate(getExcelField(row, ['Ngày khám gần nhất', 'lastExamDate', 'Last Exam Date'])),
            nextAppointment: normalizeExcelDate(getExcelField(row, ['Lịch hẹn tiếp theo', 'nextAppointment', 'Next Appointment'])),
            hematology: {
              wbc: getExcelField(row, ['WBC', 'wbc']),
              rbc: getExcelField(row, ['RBC', 'rbc']),
              hgb: getExcelField(row, ['HGB', 'hgb']),
              hct: getExcelField(row, ['HCT', 'hct']),
              mcv: getExcelField(row, ['MCV', 'mcv']),
              mch: getExcelField(row, ['MCH', 'mch']),
              mchc: getExcelField(row, ['MCHC', 'mchc']),
              rdw: getExcelField(row, ['RDW', 'rdw']),
              plt: getExcelField(row, ['PLT', 'plt']),
              mpv: getExcelField(row, ['MPV', 'mpv']),
              neu: getExcelField(row, ['NEU', 'neu']),
              lym: getExcelField(row, ['LYM', 'lym']),
              mono: getExcelField(row, ['MONO', 'mono']),
              eos: getExcelField(row, ['EOS', 'eos']),
              baso: getExcelField(row, ['BASO', 'baso'])
            },
            biochemistry: {
              glucose: getExcelField(row, ['Glucose', 'glucose']),
              hba1c: getExcelField(row, ['HbA1c', 'hba1c']),
              ure: getExcelField(row, ['Ure', 'ure']),
              creatinine: getExcelField(row, ['Creatinine', 'creatinine']),
              ast: getExcelField(row, ['AST', 'ast']),
              alt: getExcelField(row, ['ALT', 'alt']),
              ggt: getExcelField(row, ['GGT', 'ggt']),
              bilirubin: getExcelField(row, ['Bilirubin', 'bilirubin']),
              cholesterol: getExcelField(row, ['Cholesterol', 'cholesterol']),
              triglyceride: getExcelField(row, ['Triglyceride', 'triglyceride']),
              hdl: getExcelField(row, ['HDL', 'hdl']),
              ldl: getExcelField(row, ['LDL', 'ldl']),
              uricAcid: getExcelField(row, ['Acid Uric', 'uricAcid', 'Uric Acid']),
              totalProtein: getExcelField(row, ['Total Protein', 'TotalProtein', 'totalProtein']),
              albumin: getExcelField(row, ['Albumin', 'albumin']),
              globulin: getExcelField(row, ['Globulin', 'globulin']),
              agRatio: getExcelField(row, ['A/G Ratio', 'AgRatio', 'agRatio']),
              ldh: getExcelField(row, ['LDH', 'ldh']),
              calcium: getExcelField(row, ['Calcium', 'calcium']),
              phosphate: getExcelField(row, ['Phosphate', 'phosphate']),
              magnesium: getExcelField(row, ['Magnesium', 'magnesium'])
            },
            immunology: {
              hbsag: getExcelField(row, ['HBsAg', 'hbsag']),
              antiHbs: getExcelField(row, ['Anti-HBs', 'antiHbs']),
              hcv: getExcelField(row, ['HCV', 'hcv']),
              hiv: getExcelField(row, ['HIV', 'hiv']),
              afp: getExcelField(row, ['AFP', 'afp']),
              cea: getExcelField(row, ['CEA', 'cea']),
              psa: getExcelField(row, ['PSA', 'psa']),
              ca125: getExcelField(row, ['CA125', 'ca125']),
              ca199: getExcelField(row, ['CA19-9', 'ca199']),
              ca153: getExcelField(row, ['CA15-3', 'ca153']),
              crp: getExcelField(row, ['CRP', 'crp']),
              rf: getExcelField(row, ['RF', 'rf']),
              ana: getExcelField(row, ['ANA', 'ana']),
              igg: getExcelField(row, ['IgG', 'igg']),
              igm: getExcelField(row, ['IgM', 'igm']),
              iga: getExcelField(row, ['IgA', 'iga']),
              tsh: getExcelField(row, ['TSH', 'tsh']),
              t3: getExcelField(row, ['T3', 't3']),
              t4: getExcelField(row, ['T4', 't4']),
              ft3: getExcelField(row, ['FT3', 'ft3']),
              ft4: getExcelField(row, ['FT4', 'ft4']),
              cortisol: getExcelField(row, ['Cortisol', 'cortisol'])
            }
          };

          // Build reminders from row if present
          normalized.reminders = [];
          if (revisitDateVal || revisitContent) {
            normalized.reminders.push({
              revisitDate: revisitDateVal || '',
              content: revisitContent || '',
              note: revisitNote || '',
              contacted: false,
              done: false
            });
          }

          if (existing) {
            // Merge basic fields
            Object.assign(existing, normalized);
            // Merge reminders: append new reminders if not duplicate (by revisitDate+content)
            existing.reminders = existing.reminders || [];
            (normalized.reminders || []).forEach(r => {
              const duplicate = existing.reminders.find(er => (er.revisitDate || '') === (r.revisitDate || '') && (er.content || '') === (r.content || ''));
              if (!duplicate) existing.reminders.push(r);
            });
            // Update nextAppointment to nearest upcoming revisitDate if available
            const upcomingDates = (existing.reminders || []).map(x => x.revisitDate).filter(Boolean).sort();
            existing.nextAppointment = upcomingDates.length ? upcomingDates[0] : existing.nextAppointment;
            updatedCount++;
          } else {
            normalized.conclusion = {};
            normalized.history = [];
            normalized.reminders = normalized.reminders || [];
            if (normalized.reminders.length) {
              const upcomingDates = normalized.reminders.map(x => x.revisitDate).filter(Boolean).sort();
              normalized.nextAppointment = upcomingDates.length ? upcomingDates[0] : (normalized.nextAppointment || '');
            }
            patients.push(normalized);
            newCount++;
          }
        });

        await saveData();
        renderDashboard();
        renderPatientTable();
        renderReminders();
        if (currentPatientId) {
          const p = patients.find(x => x.patientId === currentPatientId);
          if (p) renderPatientReminders(p);
        }
        alert(`Đã đồng bộ xong. Tạo mới: ${newCount}, Cập nhật: ${updatedCount}`);
      };
      reader.readAsArrayBuffer(file);
    }

    function exportPatientsToExcel() {
      const includeInfo = document.getElementById('exportInfoCheckbox').checked;
      const includeHematology = document.getElementById('exportHematologyCheckbox').checked;
      const includeBiochemistry = document.getElementById('exportBiochemistryCheckbox').checked;
      const includeImmunology = document.getElementById('exportImmunologyCheckbox').checked;
      if (!includeInfo && !includeHematology && !includeBiochemistry && !includeImmunology) {
        alert('Vui lòng chọn ít nhất một loại dữ liệu để xuất.');
        return;
      }

      const rows = patients.map(patient => {
        const row = {};
        if (includeInfo) {
          row['Mã BN'] = patient.patientId;
          row['Họ tên'] = patient.fullName;
          row['Ngày sinh'] = patient.birthDate || '';
          row['Giới tính'] = patient.gender || '';
          row['Mã BHYT'] = patient.healthInsurance || '';
          row['CCCD'] = patient.cccd || '';
          row['SĐT'] = patient.phone || '';
          row['Dân tộc'] = patient.ethnicity || '';
          row['Nghề nghiệp'] = patient.job || '';
          row['Địa chỉ'] = patient.address || '';
          row['Ghi chú'] = patient.note || '';
          row['Ngày khám gần nhất'] = patient.lastExamDate || '';
          row['Lịch hẹn tiếp theo'] = patient.nextAppointment || '';
        }
        if (includeHematology) {
          const h = patient.hematology || {};
          row['WBC'] = h.wbc || '';
          row['RBC'] = h.rbc || '';
          row['HGB'] = h.hgb || '';
          row['HCT'] = h.hct || '';
          row['MCV'] = h.mcv || '';
          row['MCH'] = h.mch || '';
          row['MCHC'] = h.mchc || '';
          row['RDW'] = h.rdw || '';
          row['PLT'] = h.plt || '';
          row['MPV'] = h.mpv || '';
          row['NEU'] = h.neu || '';
          row['LYM'] = h.lym || '';
          row['MONO'] = h.mono || '';
          row['EOS'] = h.eos || '';
          row['BASO'] = h.baso || '';
        }
        if (includeBiochemistry) {
          const b = patient.biochemistry || {};
          row['Glucose'] = b.glucose || '';
          row['HbA1c'] = b.hba1c || '';
          row['Ure'] = b.ure || '';
          row['Creatinine'] = b.creatinine || '';
          row['AST'] = b.ast || '';
          row['ALT'] = b.alt || '';
          row['GGT'] = b.ggt || '';
          row['Bilirubin'] = b.bilirubin || '';
          row['Cholesterol'] = b.cholesterol || '';
          row['Triglyceride'] = b.triglyceride || '';
          row['HDL'] = b.hdl || '';
          row['LDL'] = b.ldl || '';
          row['Acid Uric'] = b.uricAcid || '';
          row['Total Protein'] = b.totalProtein || '';
          row['Albumin'] = b.albumin || '';
          row['Globulin'] = b.globulin || '';
          row['A/G Ratio'] = b.agRatio || '';
          row['LDH'] = b.ldh || '';
          row['Calcium'] = b.calcium || '';
          row['Phosphate'] = b.phosphate || '';
          row['Magnesium'] = b.magnesium || '';
        }
        if (includeImmunology) {
          const i = patient.immunology || {};
          row['HBsAg'] = i.hbsag || '';
          row['Anti-HBs'] = i.antiHbs || '';
          row['HCV'] = i.hcv || '';
          row['HIV'] = i.hiv || '';
          row['AFP'] = i.afp || '';
          row['CEA'] = i.cea || '';
          row['PSA'] = i.psa || '';
          row['CA125'] = i.ca125 || '';
          row['CA19-9'] = i.ca199 || '';
          row['CA15-3'] = i.ca153 || '';
          row['CRP'] = i.crp || '';
          row['RF'] = i.rf || '';
          row['ANA'] = i.ana || '';
          row['IgG'] = i.igg || '';
          row['IgM'] = i.igm || '';
          row['IgA'] = i.iga || '';
          row['TSH'] = i.tsh || '';
          row['T3'] = i.t3 || '';
          row['T4'] = i.t4 || '';
          row['FT3'] = i.ft3 || '';
          row['FT4'] = i.ft4 || '';
          row['Cortisol'] = i.cortisol || '';
        }
        return row;
      });

      const worksheet = XLSX.utils.json_to_sheet(rows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Bệnh nhân');
      const fileName = `patients_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(workbook, fileName);
    }

    async function logout() {
      if (window.firebaseAuth && window.firebaseAuthImports) {
        const { signOut } = window.firebaseAuthImports;
        try {
          await signOut(window.firebaseAuth);
        } catch (error) {
          console.warn('Firebase sign out failed:', error);
        }
      }
      currentPatientId = null;
      showLoginScreen();
      alert('Đã đăng xuất.');
    }

    function showLoginScreen() {
      const loginScreen = document.getElementById('loginScreen');
      const appWrapper = document.getElementById('appWrapper');
      if (loginScreen && appWrapper) {
        loginScreen.style.display = 'flex';
        appWrapper.style.display = 'none';
      }
    }

    function showApp() {
      const loginScreen = document.getElementById('loginScreen');
      const appWrapper = document.getElementById('appWrapper');
      if (!window.firebaseAuth || !window.firebaseAuth.currentUser) {
        showLoginScreen();
        return;
      }
      if (loginScreen && appWrapper) {
        loginScreen.style.display = 'none';
        appWrapper.style.display = 'block';
      }
    }

    async function handleLogin() {
      const email = document.getElementById('loginUser').value.trim();
      const password = document.getElementById('loginPassword').value.trim();
      if (!email || !password) {
        alert('Vui lòng nhập email và mật khẩu.');
        return false;
      }
      if (!window.firebaseAuth || !window.firebaseAuthImports) {
        alert('Firebase chưa khởi tạo. Vui lòng thử lại sau.');
        return false;
      }
      try {
        const { signInWithEmailAndPassword } = window.firebaseAuthImports;
        await signInWithEmailAndPassword(window.firebaseAuth, email, password);
        await loadData();
        renderDashboard();
        renderPatientTable();
        renderReminders();
        showApp();
        return true;
      } catch (error) {
        alert('Đăng nhập thất bại: ' + (error.message || 'Kiểm tra lại email và mật khẩu.'));
        showLoginScreen();
        return false;
      }
    }

    async function loadData() {
      let firebaseData = null;
      if (window.firebaseAuth && window.firebaseAuth.currentUser) {
        firebaseData = await loadDataFromFirebase();
      }
      const stored = localStorage.getItem(STORAGE_KEY);
      if (firebaseData && firebaseData.length) {
        patients = firebaseData;
        localStorage.setItem(STORAGE_KEY, JSON.stringify(patients));
      } else if (stored) {
        patients = JSON.parse(stored);
        if (window.firebaseAuth && window.firebaseAuth.currentUser) {
          await saveDataToFirebase();
        }
      } else {
        patients = samplePatients;
        await saveData();
      }
      patients.forEach(patient => {
        ensureReminderForNextAppointment(patient);
        syncNextAppointmentWithReminders(patient);
      });
    }

    function formatDate(value) {
      if (!value) return '-';
      return new Date(value).toLocaleDateString('vi-VN');
    }

    function calculateAge(birthDate) {
      if (!birthDate) return '';
      const birth = new Date(birthDate);
      const today = new Date();
      let age = today.getFullYear() - birth.getFullYear();
      const m = today.getMonth() - birth.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) age--;
      return age;
    }

    function compareValues(a, b) {
      if (a === null || a === undefined) return 1;
      if (b === null || b === undefined) return -1;
      if (!isNaN(Date.parse(a)) && !isNaN(Date.parse(b))) return new Date(a) - new Date(b);
      if (!isNaN(parseFloat(a)) && !isNaN(parseFloat(b))) return parseFloat(a) - parseFloat(b);
      return a.toString().localeCompare(b.toString(), 'vi');
    }

    function formatAppointmentLabel(diff) {
      if (diff < 0) return 'Quá hạn';
      if (diff === 0) return 'Hôm nay';
      if (diff === 1) return '1 ngày';
      if (diff <= 7) return `${diff} ngày`;
      if (diff <= 30) {
        const weeks = Math.ceil(diff / 7);
        return weeks === 1 ? '1 tuần' : `${weeks} tuần`;
      }
      const months = Math.ceil(diff / 30);
      return months === 1 ? '1 tháng' : `${months} tháng`;
    }

    function getStatus(patient) {
      if (!patient.nextAppointment) return { label: 'Không có gì', type: 'gray' };
      const today = new Date();
      const next = new Date(patient.nextAppointment);
      const diff = Math.ceil((next - today) / (1000 * 60 * 60 * 24));
      if (diff < 0) return { label: 'Quá hạn', type: 'red' };
      if (diff === 0) return { label: 'Hôm nay', type: 'orange' };
      const label = formatAppointmentLabel(diff);
      const type = diff <= 7 ? 'yellow' : 'green';
      return { label, type };
    }

    function getPatientDisplayStatus(patient) {
      if (!patient) return { label: 'Không có gì', type: 'gray' };
      if (patient.examinationState === 'in-progress') {
        return { label: 'Đang khám - chờ kết quả', type: 'orange' };
      }
      if (patient.nextAppointment) {
        const status = getStatus(patient);
        let label;
        if (status.label === 'Hôm nay') {
          label = 'Tái khám hôm nay';
        } else if (status.label === 'Quá hạn') {
          label = 'Tái khám quá hạn';
        } else {
          label = `Tái khám sau ${status.label}`;
        }
        return { label, type: status.type };
      }
      if (patient.examinationState === 'examined' || patient.resultsReturned) {
        return { label: 'Đã trả kết quả, chờ tái khám', type: 'green' };
      }
      return { label: 'Không có gì', type: 'gray' };
    }

    function renderDashboard() {
      const total = patients.length;
      const upcoming = patients.reduce((acc, patient) => {
        if (!patient.nextAppointment) return acc;
        const diff = Math.ceil((new Date(patient.nextAppointment) - new Date()) / (1000 * 60 * 60 * 24));
        return acc + (diff >= 0 && diff <= 7 ? 1 : 0);
      }, 0);
      const overdue = patients.reduce((acc, patient) => {
        if (!patient.nextAppointment) return acc;
        const diff = Math.ceil((new Date(patient.nextAppointment) - new Date()) / (1000 * 60 * 60 * 24));
        return acc + (diff < 0 ? 1 : 0);
      }, 0);
      const currentMonthReminders = getCurrentMonthReminders();
      // compute today / week / month / contacted / done
      const today = new Date();
      const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const endOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23,59,59);
      const msDay = 1000*60*60*24;
      const todayCount = patients.reduce((acc, p) => p.nextAppointment && new Date(p.nextAppointment) >= startOfToday && new Date(p.nextAppointment) <= endOfToday ? acc+1 : acc, 0);
      const weekCount = patients.reduce((acc, p) => {
        if (!p.nextAppointment) return acc;
        const diff = Math.ceil((new Date(p.nextAppointment) - new Date()) / msDay);
        return acc + (diff >= 0 && diff <= 7 ? 1 : 0);
      }, 0);
      const overdueCount = patients.reduce((acc, p) => {
        if (!p.nextAppointment) return acc;
        const diff = Math.ceil((new Date(p.nextAppointment) - new Date()) / msDay);
        return acc + (diff < 0 ? 1 : 0);
      }, 0);
      const monthCount = currentMonthReminders.length;
      const contactedCount = patients.reduce((acc, p) => {
        return acc + (p.reminders ? p.reminders.filter(r => r.contacted && !r.done).length : 0);
      }, 0);
      const doneCount = patients.reduce((acc, p) => {
        return acc + (p.reminders ? p.reminders.filter(r => r.done).length : 0);
      }, 0);

      document.getElementById('todayAppointments') && (document.getElementById('todayAppointments').textContent = todayCount);
      document.getElementById('weekAppointments') && (document.getElementById('weekAppointments').textContent = weekCount);
      document.getElementById('overdueAppointments') && (document.getElementById('overdueAppointments').textContent = overdueCount);
      document.getElementById('monthAppointments') && (document.getElementById('monthAppointments').textContent = monthCount);
      document.getElementById('contactedCount') && (document.getElementById('contactedCount').textContent = contactedCount);
      document.getElementById('doneCount') && (document.getElementById('doneCount').textContent = doneCount);
      document.getElementById('todayCountBadge') && (document.getElementById('todayCountBadge').textContent = todayCount);

      // day progress
      const dayTotal = todayCount;
      const dayDone = patients.reduce((acc, p) => {
        if (!p.nextAppointment) return acc;
        const na = new Date(p.nextAppointment);
        if (na >= startOfToday && na <= endOfToday) {
          return acc + (p.reminders ? p.reminders.filter(r => r.done).length : 0);
        }
        return acc;
      }, 0);
      const percent = dayTotal ? Math.round((dayDone/dayTotal)*100) : 0;
      const dayTextEl = document.getElementById('dayProgressText');
      const dayBarEl = document.getElementById('dayProgressBar');
      if (dayTextEl) dayTextEl.textContent = `${dayDone} / ${dayTotal}`;
      if (dayBarEl) { dayBarEl.style.width = percent + '%'; dayBarEl.setAttribute('aria-valuenow', percent); }

      // update older elements for compatibility
      if (document.getElementById('totalPatients')) document.getElementById('totalPatients').textContent = total;
      if (document.getElementById('totalPatientsBadge')) document.getElementById('totalPatientsBadge').textContent = total;
      if (document.getElementById('upcomingAppointments')) document.getElementById('upcomingAppointments').textContent = upcoming;
      if (document.getElementById('overdueAppointments')) document.getElementById('overdueAppointments').textContent = overdue;
      if (document.getElementById('currentMonthRevisitCount')) document.getElementById('currentMonthRevisitCount').textContent = monthCount;

      renderCurrentMonthAppointments();
      drawAppointmentChart();
      renderTodayAppointments();
      updateReceptionStats();
    }

    function updateReceptionStats() {
      const total = patients.length;
      const scheduled = patients.filter(p => p.nextAppointment).length;
      const noSchedule = patients.filter(p => !p.nextAppointment).length;
      const totalEl = document.getElementById('receptionTotalBadge');
      const scheduledEl = document.getElementById('receptionScheduledBadge');
      const noScheduleEl = document.getElementById('receptionNoScheduleBadge');
      if (totalEl) totalEl.textContent = total;
      if (scheduledEl) scheduledEl.textContent = scheduled;
      if (noScheduleEl) noScheduleEl.textContent = noSchedule;
    }

    function renderTodayAppointments() {
      const tbody = document.getElementById('todayAppointmentsBody');
      if (!tbody) return;
      const today = new Date();
      const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const endOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate(),23,59,59);
      const rows = patients.filter(p => p.nextAppointment && new Date(p.nextAppointment) >= startOfToday && new Date(p.nextAppointment) <= endOfToday)
        .map(p => {
          const st = getPatientDisplayStatus(p);
          const badgeClass = st.type === 'green' ? 'bg-success' : st.type === 'red' ? 'bg-danger' : st.type === 'gray' ? 'bg-secondary' : 'bg-warning';
          return `<tr>
            <td>${p.patientId}</td>
            <td>${p.fullName}</td>
            <td>${p.phone || '-'}</td>
            <td>${formatDate(p.nextAppointment)}</td>
            <td><span class="badge ${badgeClass} text-white">${st.label}</span></td>
          </tr>`;
        }).join('');
      tbody.innerHTML = rows || '<tr><td colspan="5" class="text-muted">Không có lịch khám hôm nay</td></tr>';
    }

    function openStatsModal(category) {
      const modalEl = document.getElementById('statsModal');
      const title = document.getElementById('statsModalTitle');
      const body = document.getElementById('statsModalBody');
      if (!modalEl || !body || !title) return;
      let list = [];
      const now = new Date();
      const msDay = 1000*60*60*24;
      if (category === 'today') {
        title.textContent = 'Danh sách tái khám hôm nay';
        const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(),23,59,59);
        list = patients.filter(p => p.nextAppointment && new Date(p.nextAppointment) >= start && new Date(p.nextAppointment) <= end);
      } else if (category === 'week') {
        title.textContent = 'Danh sách tái khám tuần này';
        list = patients.filter(p => p.nextAppointment && Math.ceil((new Date(p.nextAppointment)-now)/msDay) >=0 && Math.ceil((new Date(p.nextAppointment)-now)/msDay) <=7);
      } else if (category === 'month') {
        title.textContent = 'Danh sách tái khám tháng này';
        const m = now.getMonth(), y = now.getFullYear();
        list = patients.filter(p => p.nextAppointment && new Date(p.nextAppointment).getMonth()===m && new Date(p.nextAppointment).getFullYear()===y);
      } else if (category === 'overdue') {
        title.textContent = 'Danh sách tái khám quá hạn';
        list = patients.filter(p => p.nextAppointment && Math.ceil((new Date(p.nextAppointment)-now)/msDay) < 0);
      } else if (category === 'contacted') {
        title.textContent = 'Danh sách đã liên hệ';
        list = patients.filter(p => p.reminders && p.reminders.some(r => r.contacted && !r.done));
      } else if (category === 'done') {
        title.textContent = 'Danh sách đã khám';
        list = patients.filter(p => p.reminders && p.reminders.some(r => r.done));
      }
      body.innerHTML = list.map(p => {
        const st = getPatientDisplayStatus(p);
        const badgeClass = st.type === 'green' ? 'bg-success' : st.type === 'red' ? 'bg-danger' : st.type === 'gray' ? 'bg-secondary' : 'bg-warning';
        return `<tr><td>${p.patientId}</td><td>${p.fullName}</td><td>${p.phone||'-'}</td><td>${formatDate(p.nextAppointment)}</td><td><span class="badge ${badgeClass} text-white">${st.label}</span></td></tr>`;
      }).join('') || '<tr><td colspan="5" class="text-muted">Không có bệnh nhân phù hợp</td></tr>';
      const bs = new bootstrap.Modal(modalEl);
      bs.show();
    }

    // Helpers for Excel import mapping
    function normalizeHeader(h) {
      if (!h) return '';
      return h.toString().trim().toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-z0-9]/g, '');
    }

    function mapRowToPatient(row) {
      const obj = {};
      const map = {};
      Object.keys(row).forEach(k => map[normalizeHeader(k)] = row[k]);
      const pick = (names) => {
        for (const n of names) if (map[n] !== undefined) return map[n];
        return '';
      };
      obj.patientId = (pick(['mabn','mabenhnhan','mabenh','patientid','id','code']) || '').toString().trim();
      obj.fullName = (pick(['hoten','hovaten','name','fullname']) || '').toString().trim();
      obj.birthDate = normalizeExcelDate(pick(['ngaysinh','birthdate','dob']));
      obj.gender = (pick(['gioitinh','gender']) || '').toString().trim();
      obj.healthInsurance = (pick(['bhyt','mabhyt','insurance']) || '').toString().trim();
      obj.cccd = (pick(['cccd','cmnd','idcard']) || '').toString().trim();
      obj.phone = (pick(['sdt','dienthoai','phone']) || '').toString().trim();
      obj.address = (pick(['diachi','address']) || '').toString().trim();
      obj.note = (pick(['ghichu','note','remark']) || '').toString().trim();
      obj.nextAppointment = normalizeExcelDate(pick(['lichhen','lich hen','lichhenngay','nextappointment','revisit','appointment']));
      obj.lastExamDate = normalizeExcelDate(pick(['ngaytiepnhan','ngaynhan','lastexamdate','examdate']));
      obj.clinical = {};
      obj.hematology = {};
      obj.biochemistry = {};
      obj.immunology = {};
      obj.conclusion = {};
      obj.history = [];
      obj.reminders = [];
      return obj;
    }

    async function handleReceptionImportPatients(file) {
      if (!file) return;
      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        let added = 0, updated = 0;
        rows.forEach(r => {
          const p = mapRowToPatient(r);
          if (!p.patientId) return;
          const idx = patients.findIndex(x => x.patientId === p.patientId);
          if (idx !== -1) {
            patients[idx] = { ...patients[idx], ...p };
            updated++;
          } else {
            patients.push(p);
            added++;
          }
        });
        await saveData();
        renderDashboard();
        renderPatientTable();
        alert(`Import xong. Thêm ${added}, Cập nhật ${updated}.`);
      } catch (err) {
        console.error(err);
        alert('Lỗi import file bệnh nhân.');
      }
    }

    async function handleReceptionImportResults(file) {
      if (!file) return;
      try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        let matched = 0, unmatched = 0;
        rows.forEach(r => {
          const map = {};
          Object.keys(r).forEach(k => map[normalizeHeader(k)] = r[k]);
          const id = (map['mabn'] || map['mabenhnhan'] || map['id'] || map['patientid'] || '').toString().trim();
          if (!id) { unmatched++; return; }
          const idx = patients.findIndex(p => p.patientId === id);
          if (idx === -1) { unmatched++; return; }
          patients[idx].tests = patients[idx].tests || [];
          patients[idx].tests.push(r);
          matched++;
        });
        await saveData();
        renderDashboard();
        renderPatientTable();
        alert(`Import kết quả xong. Ghép được ${matched}. Không tìm thấy ${unmatched}.`);
      } catch (err) {
        console.error(err);
        alert('Lỗi import file kết quả.');
      }
    }

    function getCurrentMonthReminders() {
      const now = new Date();
      const year = now.getFullYear();
      const month = now.getMonth();
      return patients.flatMap(patient => {
        const reminders = (patient.reminders || []).map(reminder => ({
          ...reminder,
          patientId: patient.patientId,
          fullName: patient.fullName
        }));
        if (patient.nextAppointment) {
          const hasNextAppointmentReminder = reminders.some(r => r.revisitDate === patient.nextAppointment);
          if (!hasNextAppointmentReminder) {
            reminders.push({
              revisitDate: patient.nextAppointment,
              content: 'Lịch hẹn tiếp theo',
              note: '',
              patientId: patient.patientId,
              fullName: patient.fullName
            });
          }
        }
        return reminders;
      }).filter(reminder => {
        if (!reminder.revisitDate) return false;
        const date = new Date(reminder.revisitDate);
        return date.getFullYear() === year && date.getMonth() === month;
      }).sort((a, b) => new Date(a.revisitDate) - new Date(b.revisitDate));
    }

    function getPatientReminderList(patient) {
      if (!patient) return [];
      const reminders = (patient.reminders || []).slice();
      if (patient.nextAppointment) {
        const hasNextAppointmentReminder = reminders.some(r => r.revisitDate === patient.nextAppointment);
        if (!hasNextAppointmentReminder) {
          reminders.push({
            examDate: patient.lastExamDate || new Date().toISOString().slice(0, 10),
            revisitDate: patient.nextAppointment,
            content: 'Lịch hẹn tiếp theo',
            note: '',
            contacted: false,
            done: false,
            isSynthetic: true
          });
        }
      }
      return reminders.sort((a, b) => new Date(a.revisitDate) - new Date(b.revisitDate));
    }

    function getActiveReminderDates(patient) {
      if (!patient || !Array.isArray(patient.reminders)) return [];
      return patient.reminders
        .filter(r => r.revisitDate && !r.done && !r.contacted)
        .map(r => r.revisitDate)
        .sort();
    }

    function syncNextAppointmentWithReminders(patient) {
      if (!patient) return;
      patient.reminders = patient.reminders || [];
      const activeDates = getActiveReminderDates(patient);
      patient.nextAppointment = activeDates.length ? activeDates[0] : '';
    }

    function archiveNonMatchingActiveReminders(patient) {
      if (!patient || !Array.isArray(patient.reminders) || !patient.nextAppointment) return;
      const kept = [];
      const archived = [];
      patient.reminders.forEach(reminder => {
        if (!reminder.revisitDate || reminder.done || reminder.contacted || reminder.revisitDate === patient.nextAppointment) {
          kept.push(reminder);
        } else {
          archived.push(reminder);
        }
      });
      if (!archived.length) {
        patient.reminders = kept;
        return;
      }
      patient.history = patient.history || [];
      archived.forEach(reminder => {
        patient.history.push({
          date: reminder.revisitDate || reminder.examDate,
          diagnosis: reminder.content || 'Tái khám',
          result: reminder.contacted ? 'Đã liên hệ' : 'Đã khám',
          note: reminder.note || ''
        });
      });
      patient.reminders = kept;
    }

    function ensureReminderForNextAppointment(patient) {
      if (!patient) return;
      patient.reminders = patient.reminders || [];
      if (!patient.nextAppointment) return;
      archiveNonMatchingActiveReminders(patient);
      const hasActiveReminderForDate = patient.reminders.some(r => r.revisitDate === patient.nextAppointment && !r.done && !r.contacted);
      if (hasActiveReminderForDate) return;
      patient.reminders.push({
        examDate: patient.lastExamDate || new Date().toISOString().slice(0, 10),
        revisitDate: patient.nextAppointment,
        content: 'Tái khám theo lịch hẹn',
        note: '',
        contacted: false,
        done: false
      });
    }

    function renderCurrentMonthAppointments() {
      const reminders = getCurrentMonthReminders();
      const tbody = document.getElementById('currentMonthAppointmentsBody');
      document.getElementById('monthCount').textContent = reminders.length;
      if (!reminders.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-muted">Không có lịch tái khám trong tháng này</td></tr>';
        return;
      }
      tbody.innerHTML = reminders.map(reminder => {
        const status = getReminderStatus(reminder);
        return `
          <tr>
            <td>${formatDate(reminder.revisitDate)}</td>
            <td>${reminder.patientId}</td>
            <td>${reminder.fullName}</td>
            <td><span class="status-chip ${status.type}">${status.label}</span></td>
          </tr>
        `;
      }).join('');
    }

    function getWeeklyAppointmentCounts(reminders) {
      const daysInMonth = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0).getDate();
      const ranges = [
        { label: 'Tuần 1', start: 1, end: 7 },
        { label: 'Tuần 2', start: 8, end: 14 },
        { label: 'Tuần 3', start: 15, end: 21 },
        { label: 'Tuần 4', start: 22, end: daysInMonth }
      ];
      return ranges.map(range => ({
        label: range.label,
        count: reminders.filter(reminder => {
          const day = new Date(reminder.revisitDate).getDate();
          return day >= range.start && day <= range.end;
        }).length
      }));
    }

    let appointmentChartInstance = null;

    function drawAppointmentChart() {
      const canvas = document.getElementById('appointmentChart');
      if (!canvas) return;
      
      const reminders = getCurrentMonthReminders();
      const data = getWeeklyAppointmentCounts(reminders);
      const values = data.map(item => item.count);
      const weekLabels = ['Tuần 1', 'Tuần 2', 'Tuần 3', 'Tuần 4', 'Tuần 5'];
      
      const ctx = canvas.getContext('2d');
      
      // Destroy old chart instance if exists
      if (appointmentChartInstance) {
        appointmentChartInstance.destroy();
      }
      
      appointmentChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: weekLabels.slice(0, values.length),
          datasets: [{
            label: 'Số bệnh nhân tái khám',
            data: values,
            backgroundColor: [
              '#4F46E5',
              '#4F46E5',
              '#4F46E5',
              '#4F46E5',
              '#4F46E5'
            ],
            borderColor: ['#4338CA', '#4338CA', '#4338CA', '#4338CA', '#4338CA'],
            borderWidth: 0,
            borderRadius: [8, 8, 0, 0],
            borderSkipped: false,
            barThickness: 50,
            categoryPercentage: 0.75,
            barPercentage: 0.85
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: 'x',
          plugins: {
            legend: {
              display: false
            },
            tooltip: {
              enabled: true,
              backgroundColor: 'rgba(0, 0, 0, 0.8)',
              titleFont: { size: 13, weight: '600' },
              bodyFont: { size: 12 },
              padding: 12,
              cornerRadius: 8,
              displayColors: false,
              callbacks: {
                title: function(context) {
                  return context[0].label;
                },
                label: function(context) {
                  return 'Số bệnh nhân: ' + context.parsed.y;
                }
              }
            }
          },
          scales: {
            y: {
              beginAtZero: true,
              max: Math.max(...values, 5),
              ticks: {
                font: { size: 12, family: "'Inter', system-ui" },
                color: '#6b7280',
                stepSize: Math.max(1, Math.ceil(Math.max(...values, 5) / 5))
              },
              grid: {
                color: '#E5E7EB',
                lineWidth: 0.75,
                drawBorder: false
              }
            },
            x: {
              ticks: {
                font: { size: 13, weight: '600', family: "'Inter', system-ui" },
                color: '#1f2937'
              },
              grid: {
                display: false,
                drawBorder: false
              }
            }
          }
        }
      });
    }

    function renderCurrentTime() {
      const now = new Date();
      const dateEl = document.getElementById('currentDate');
      const timeEl = document.getElementById('currentClock');
      if (dateEl) dateEl.textContent = now.toLocaleDateString('vi-VN', { weekday: 'long', year: 'numeric', month: 'short', day: 'numeric' });
      if (timeEl) timeEl.textContent = now.toLocaleTimeString('vi-VN');
    }

    function renderPatientTable() {
      const tbody = document.getElementById('patientTableBody');
      tbody.innerHTML = '';
      const filtered = getFilteredPatients();
      filtered.sort((a, b) => {
        const cmp = compareValues(a[currentSort.field], b[currentSort.field]);
        return currentSort.direction === 'asc' ? cmp : -cmp;
      });
      filtered.forEach(patient => {
        const status = getPatientDisplayStatus(patient);
        const row = document.createElement('tr');
        const checked = selectedPatientIds.has(patient.patientId) ? 'checked' : '';
        row.innerHTML = `
          <td><input type="checkbox" class="form-check-input" onchange="toggleSelectPatient('${patient.patientId}')" ${checked} /></td>
          <td>${patient.patientId}</td>
          <td>${patient.fullName}</td>
          <td>${formatDate(patient.birthDate)}</td>
          <td>${patient.gender}</td>
          <td>${patient.phone}</td>
          <td>${patient.cccd}</td>
          <td>${patient.lastExamDate ? formatDate(patient.lastExamDate) : '-'}</td>
          <td>${patient.nextAppointment ? formatDate(patient.nextAppointment) : '-'}</td>
          <td><span class="status-chip ${status.type}">${status.label}</span></td>
          <td>
            <button class="btn btn-sm btn-info me-1" onclick="viewPatient('${patient.patientId}')"><i class="fa-solid fa-eye"></i></button>
            <button class="btn btn-sm btn-warning me-1" onclick="editPatient('${patient.patientId}')"><i class="fa-solid fa-pen-to-square"></i></button>
            <button class="btn btn-sm btn-danger" onclick="deletePatient('${patient.patientId}')"><i class="fa-solid fa-trash"></i></button>
          </td>
        `;
        tbody.appendChild(row);
      });
      updateSelectedCount();
    }

    const resultGroups = {
      hematology: {
        title: 'Công thức máu',
        fixedFields: [
          { key: 'wbc', title: 'WBC', normal: '4.0-10.0', unit: '×10^9/L' },
          { key: 'neu', title: 'Neu#', normal: '2.0-7.5', unit: '×10^9/L' },
          { key: 'lym', title: 'Lym#', normal: '1.0-4.0', unit: '×10^9/L' },
          { key: 'mono', title: 'Mon#', normal: '0.2-0.8', unit: '×10^9/L' },
          { key: 'eos', title: 'Eos#', normal: '0.0-0.5', unit: '×10^9/L' },
          { key: 'baso', title: 'Baso#', normal: '0.0-0.2', unit: '×10^9/L' },
          { key: 'neuPercent', title: 'Neu%', normal: '40-75', unit: '%' },
          { key: 'lymPercent', title: 'Lym%', normal: '20-45', unit: '%' },
          { key: 'monoPercent', title: 'Mon%', normal: '2-8', unit: '%' },
          { key: 'eosPercent', title: 'Eos%', normal: '0.0-5', unit: '%' },
          { key: 'basoPercent', title: 'Baso%', normal: '0.0-2', unit: '%' },
          { key: 'rbc', title: 'RBC', normal: '4.0-5.5', unit: '×10^12/L' },
          { key: 'hgb', title: 'HGB', normal: '130-170 Nam / 120-160 Nữ', unit: 'g/L' },
          { key: 'hct', title: 'HCT', normal: '40-50', unit: '%' },
          { key: 'mcv', title: 'MCV', normal: '80-100', unit: 'fL' },
          { key: 'mch', title: 'MCH', normal: '27-33', unit: 'pg' },
          { key: 'mchc', title: 'MCHC', normal: '320-360', unit: 'g/L' },
          { key: 'rdw', title: 'RDW-CV', normal: '11.5-14.5', unit: '%' },
          { key: 'rdwSd', title: 'RDW-SD', normal: '35-56', unit: 'fL' },
          { key: 'plt', title: 'PLT', normal: '150-450', unit: '×10^9/L' },
          { key: 'mpv', title: 'MPV', normal: '7.5-11.5', unit: 'fL' },
          { key: 'pdw', title: 'PDW', normal: '9-17', unit: 'fL' },
          { key: 'pct', title: 'PCT', normal: '0.19-0.39', unit: '%' },
          { key: 'plcr', title: 'PLCR', normal: '15-35', unit: '%' },
          { key: 'plcc', title: 'PLCC', normal: '15-35', unit: '×10^9/L' }
        ]
      },
      biochemistry: {
        title: 'Sinh hóa',
        fixedFields: [
          { key: 'glucose', title: 'Glucose', normal: '70-99', unit: 'mg/dL' },
          { key: 'hba1c', title: 'HbA1c', normal: '<5.7', unit: '%' },
          { key: 'ure', title: 'Ure', normal: '7-20', unit: 'mg/dL' },
          { key: 'creatinine', title: 'Creatinine', normal: '0.7-1.2', unit: 'mg/dL' },
          { key: 'ast', title: 'AST', normal: '10-40', unit: 'U/L' },
          { key: 'alt', title: 'ALT', normal: '7-56', unit: 'U/L' },
          { key: 'ggt', title: 'GGT', normal: '9-48', unit: 'U/L' },
          { key: 'bilirubin', title: 'Bilirubin', normal: '0.1-1.2', unit: 'mg/dL' },
          { key: 'cholesterol', title: 'Cholesterol', normal: '<200', unit: 'mg/dL' },
          { key: 'triglyceride', title: 'Triglyceride', normal: '<150', unit: 'mg/dL' },
          { key: 'hdl', title: 'HDL', normal: '>40', unit: 'mg/dL' },
          { key: 'ldl', title: 'LDL', normal: '<100', unit: 'mg/dL' },
          { key: 'uricAcid', title: 'Uric Acid', normal: '3.5-7.2', unit: 'mg/dL' },
          { key: 'totalProtein', title: 'Total Protein', normal: '6.0-8.3', unit: 'g/dL' },
          { key: 'albumin', title: 'Albumin', normal: '3.5-5.0', unit: 'g/dL' },
          { key: 'globulin', title: 'Globulin', normal: '2.0-3.5', unit: 'g/dL' },
          { key: 'agRatio', title: 'A/G Ratio', normal: '1.0-2.2', unit: '' },
          { key: 'ldh', title: 'LDH', normal: '140-280', unit: 'U/L' },
          { key: 'calcium', title: 'Calcium', normal: '8.4-10.2', unit: 'mg/dL' },
          { key: 'phosphate', title: 'Phosphate', normal: '2.5-4.5', unit: 'mg/dL' },
          { key: 'magnesium', title: 'Magnesium', normal: '1.7-2.2', unit: 'mg/dL' }
        ]
      },
      immunology: {
        title: 'Miễn dịch',
        fixedFields: [
          { key: 'hbsag', title: 'HBsAg', normal: 'Negative', unit: '' },
          { key: 'antiHbs', title: 'Anti-HBs', normal: '≥10', unit: 'mIU/mL' },
          { key: 'hcv', title: 'HCV', normal: 'Negative', unit: '' },
          { key: 'hiv', title: 'HIV', normal: 'Negative', unit: '' },
          { key: 'afp', title: 'AFP', normal: '<10', unit: 'ng/mL' },
          { key: 'cea', title: 'CEA', normal: '<5', unit: 'ng/mL' },
          { key: 'psa', title: 'PSA', normal: '<4', unit: 'ng/mL' },
          { key: 'ca125', title: 'CA125', normal: '<35', unit: 'U/mL' },
          { key: 'ca199', title: 'CA19-9', normal: '<37', unit: 'U/mL' },
          { key: 'ca153', title: 'CA15-3', normal: '<30', unit: 'U/mL' },
          { key: 'crp', title: 'CRP', normal: '<5', unit: 'mg/L' },
          { key: 'rf', title: 'RF', normal: '<14', unit: 'IU/mL' },
          { key: 'ana', title: 'ANA', normal: 'Negative', unit: '' }
          ,{ key: 'igg', title: 'IgG', normal: '700-1600', unit: 'mg/dL' },
          { key: 'igm', title: 'IgM', normal: '40-230', unit: 'mg/dL' },
          { key: 'iga', title: 'IgA', normal: '70-400', unit: 'mg/dL' },
          { key: 'tsh', title: 'TSH', normal: '0.4-4.0', unit: 'mIU/L' },
          { key: 't3', title: 'T3', normal: '0.8-2.0', unit: 'ng/mL' },
          { key: 't4', title: 'T4', normal: '5-12', unit: 'µg/dL' },
          { key: 'ft3', title: 'FT3', normal: '2.3-4.2', unit: 'pg/mL' },
          { key: 'ft4', title: 'FT4', normal: '0.8-1.8', unit: 'ng/dL' },
          { key: 'cortisol', title: 'Cortisol', normal: '5-25', unit: 'µg/dL' }
        ]
      },
      coagulation: {
        title: 'Đông máu',
        fixedFields: [
          { key: 'pt_s', title: 'PT (s)', normal: '11-15', unit: 's' },
          { key: 'pt_percent', title: 'PT (%)', normal: '70-120', unit: '%' },
          { key: 'inr', title: 'INR', normal: '0.8-1.2', unit: '' },
          { key: 'aptt', title: 'APTT', normal: '25-40', unit: 's' },
          { key: 'fibrinogen', title: 'Fibrinogen', normal: '200-400', unit: 'mg/dL' }
        ]
      },
      urine: {
        title: 'Nước tiểu',
        fixedFields: [
          { key: 'leu', title: 'LEU', normal: 'Negative', unit: '' },
          { key: 'nit', title: 'NIT', normal: 'Negative', unit: '' },
          { key: 'ubg', title: 'UBG', normal: 'Negative', unit: '' },
          { key: 'bil', title: 'BIL', normal: 'Negative', unit: '' },
          { key: 'pro', title: 'PRO', normal: 'Negative', unit: '' },
          { key: 'ph', title: 'pH', normal: '4.5-8.0', unit: '' },
          { key: 'bld', title: 'BLD', normal: 'Negative', unit: '' },
          { key: 'ket', title: 'KET', normal: 'Negative', unit: '' },
          { key: 'glu', title: 'GLU', normal: 'Negative', unit: '' },
          { key: 'sg', title: 'SG', normal: '1.005-1.030', unit: '' }
        ]
      }
    };

    function getPatientLabValue(patient, category, key) {
      if (!patient || !category || !key) return '';
      const hasSelection = patient.testsSelection && Array.isArray(patient.testsSelection[category]) && patient.testsSelection[category].includes(key);
      if (hasSelection) {
        if (patient.pendingResults && patient.pendingResults[category] && patient.pendingResults[category][key] !== undefined) {
          return patient.pendingResults[category][key];
        }
        return '';
      }
      if (patient.pendingResults && patient.pendingResults[category] && patient.pendingResults[category][key] !== undefined) {
        return patient.pendingResults[category][key];
      }
      if (patient[category] && patient[category][key] !== undefined) {
        return patient[category][key];
      }
      return '';
    }

    function getPatientEffectiveLabData(patient, category) {
      if (!patient || !category) return {};
      // When entering results (has testsSelection), show only pendingResults
      // This prevents showing old saved results while entering new ones
      if (patient.testsSelection && Object.keys(patient.testsSelection).some(cat => Array.isArray(patient.testsSelection[cat]) && patient.testsSelection[cat].length > 0)) {
        return (patient.pendingResults && patient.pendingResults[category]) || {};
      }
      // Otherwise show saved results from patient[category]
      return patient[category] || {};
    }

    function setPatientPendingLabValue(patient, category, key, value) {
      if (!patient || !category || !key) return;
      patient.pendingResults = patient.pendingResults || {};
      patient.pendingResults[category] = patient.pendingResults[category] || {};
      patient.pendingResults[category][key] = value;
    }

    function patientHasResults(patient) {
      if (!patient) return false;
      if (patient.results) {
        for (const k in patient.results) {
          if (Array.isArray(patient.results[k]) && patient.results[k].length > 0) return true;
        }
      }
      // also check fixed fields presence
      for (const cat in resultGroups) {
        const fixed = (resultGroups[cat].fixedFields || []).map(f => f.key);
        for (const key of fixed) {
          const value = getPatientLabValue(patient, cat, key);
          if (value !== undefined && value !== '') return true;
        }
      }
      return false;
    }

    function isResultsComplete(patient) {
      if (!patient) return false;
      const sel = patient.testsSelection || {};
      let anySelected = false;
      for (const cat of ['hematology','biochemistry','immunology','coagulation','urine']) {
        const arr = sel[cat];
        if (Array.isArray(arr) && arr.length > 0) {
          anySelected = true;
          for (const key of arr) {
            const value = getPatientLabValue(patient, cat, key);
            if (value === undefined || value === '') return false;
          }
        }
      }
      if (anySelected) return true;
      // If no active selection, consider returned results as complete.
      return !!patient.resultsReturned;
    }

    function updateResultsStatusLabel(patient) {
      const statusLabels = [
        document.getElementById('resultsModalStatusLabel'),
        document.getElementById('resultsStatusLabel'),
        document.getElementById('resultsDetailStatusLabel')
      ].filter(Boolean);
      if (!statusLabels.length) return;
      if (!patient) {
        statusLabels.forEach(label => {
          label.textContent = 'Trạng thái: Chưa chọn';
          label.className = 'badge bg-secondary me-auto';
        });
        return;
      }
      if (patient.resultsReturned) {
        statusLabels.forEach(label => {
          label.textContent = 'Trạng thái: Đã trả';
          label.className = 'badge bg-success me-auto';
        });
        return;
      }
      const hasAny = patientHasResults(patient);
      const complete = isResultsComplete(patient);
      statusLabels.forEach(label => {
        if (!hasAny) {
          label.textContent = 'Trạng thái: Chưa nhập';
          label.className = 'badge bg-primary me-auto';
        } else if (complete) {
          label.textContent = 'Trạng thái: Nhập đủ, chưa duyệt';
          label.className = 'badge bg-warning text-dark me-auto';
        } else {
          label.textContent = 'Trạng thái: Nhập chưa đủ';
          label.className = 'badge bg-danger me-auto';
        }
      });
    }

    function refreshResultsPatientInfo(patient) {
      const inlineInfo = document.getElementById('resultsPatientInfo');
      const modalInfo = document.getElementById('resultsModalPatientInfo');
      const detailsText = `${patient.phone || '-'} • ${patient.cccd || '-'} • ${patient.nextAppointment ? formatDate(patient.nextAppointment) : 'Không có lịch hẹn'} • ${patient.resultsReturned ? 'Đã trả' : 'Chưa trả'}`;
      if (inlineInfo) inlineInfo.textContent = detailsText;
      if (modalInfo) modalInfo.textContent = detailsText;
    }

    function approvePatientResults(patient) {
      if (!patient) return false;
      const complete = isResultsComplete(patient);
      if (!complete) {
        if (!confirm('Kết quả chưa nhập đủ. Bạn vẫn muốn duyệt và đánh dấu là đã trả kết quả?')) return false;
      }
      if (patient.pendingResults) {
        Object.keys(patient.pendingResults).forEach(category => {
          const pendingCategory = patient.pendingResults[category];
          if (!pendingCategory || !Object.keys(pendingCategory).length) return;
          const current = patient[category] ? { ...patient[category] } : {};
          const merged = { ...current, ...pendingCategory };
          if (JSON.stringify(current) !== JSON.stringify(merged)) {
            patient[`previous${category.charAt(0).toUpperCase() + category.slice(1)}`] = current;
          }
          patient[category] = merged;
        });
      }
      patient.pendingResults = {};
      patient.testsSelection = {};
      patient.examinationState = 'examined';
      patient.resultsReturned = true;
      patient.resultsApproved = true;
      patient.resultsReturnedDate = new Date().toISOString();
      patient.resultsApprovedDate = new Date().toISOString();
      patient.labUpdatedDate = new Date().toISOString();
      saveData();
      renderDashboard();
      renderPatientTable();
      renderResultsPatientTable();
      renderResultsPatientList();
      renderReminders();
      return true;
    }

    function getResultsFilteredPatients() {
      const filters = {
        name: (document.getElementById('resultsSearchName') || { value: '' }).value.trim().toLowerCase(),
        code: (document.getElementById('resultsSearchCode') || { value: '' }).value.trim().toLowerCase(),
        date: (document.getElementById('resultsSearchDate') || { value: '' }).value,
        status: (document.getElementById('resultsSearchStatus') || { value: 'all' }).value
      };
      return patients.filter(patient => {
        if (filters.name && !(patient.fullName || '').toLowerCase().includes(filters.name)) return false;
        if (filters.code && !(patient.patientId || '').toLowerCase().includes(filters.code)) return false;
        if (filters.date) {
          if (!patient.lastExamDate) return false;
          const pDate = new Date(patient.lastExamDate);
          if (isNaN(pDate)) return false;
          if (pDate.toISOString().slice(0,10) !== filters.date) return false;
        }
        if (filters.status === 'returned' && !patient.resultsReturned) return false;
        if (filters.status === 'notReturned' && patient.resultsReturned) return false;
        return true;
      });
    }

    function renderResultsPatientTable() {
      const totalBadge = document.getElementById('resultsTotalBadge');
      const patientsList = getResultsFilteredPatients();
      const returnedBadge = document.getElementById('resultsReturnedBadge');
      const notReturnedBadge = document.getElementById('resultsNotReturnedBadge');
      if (totalBadge) totalBadge.textContent = patientsList.length;
      const returnedCount = patientsList.filter(p => p.resultsReturned).length;
      if (returnedBadge) returnedBadge.textContent = returnedCount;
      if (notReturnedBadge) notReturnedBadge.textContent = (patientsList.length - returnedCount);
      renderResultsPatientList();
    }

    function getResultsRows(patient, category) {
      const group = resultGroups[category];
      const rows = [];
      const fixed = group.fixedFields || [];
      const fixedData = getPatientEffectiveLabData(patient, category);
      fixed.forEach(field => {
        rows.push({
          name: field.title,
          result: fixedData[field.key] || '',
          normalRange: field.normal,
          unit: field.unit,
          fixedKey: field.key,
          fixed: true
        });
      });
      const customRows = ((patient.results || {})[category] || []).map(row => ({
        name: row.name || '',
        result: row.result || '',
        normalRange: row.normalRange || '',
        unit: row.unit || '',
        fixed: false
      }));
      return rows.concat(customRows);
    }

    function renderResultsCategoryTable(patient, category) {
      const body = document.getElementById(`results${category.charAt(0).toUpperCase() + category.slice(1)}Body`);
      if (!body) return;
      const rows = getResultsRows(patient, category);
      body.innerHTML = rows.map((row, index) => {
        const deleteButton = row.fixed ? '' : `<button class="btn btn-sm btn-danger" type="button" onclick="deleteResultsRow('${category}', ${index})"><i class="fa-solid fa-trash"></i></button>`;
        return `
          <tr data-fixed-key="${row.fixedKey || ''}" data-fixed="${row.fixed}">
            <td><input class="form-control form-control-sm" value="${row.name.replace(/"/g, '&quot;')}" /></td>
            <td><input class="form-control form-control-sm" value="${row.result.replace(/"/g, '&quot;')}" /></td>
            <td><input class="form-control form-control-sm" value="${row.normalRange.replace(/"/g, '&quot;')}" /></td>
            <td><input class="form-control form-control-sm" value="${row.unit.replace(/"/g, '&quot;')}" /></td>
            <td>${deleteButton}</td>
          </tr>
        `;
      }).join('');
      if (!rows.length) {
        body.innerHTML = `<tr><td colspan="5" class="text-muted">Chưa có dữ liệu ${resultGroups[category].title.toLowerCase()}.</td></tr>`;
      }
    }

    function openResultsDetail(patientId) {
      const patient = patients.find(item => item.patientId === patientId);
      if (!patient) return;
      currentResultsPatientId = patientId;
      document.getElementById('resultsModalPatientLabel').textContent = `${patient.patientId} - ${patient.fullName}`;
      document.getElementById('resultsModalPatientInfo').textContent = `${patient.phone || '-'} • ${patient.cccd || '-'} • ${patient.nextAppointment ? formatDate(patient.nextAppointment) : 'Không có lịch hẹn'} • ${patient.resultsReturned ? 'Đã trả' : 'Chưa trả'}`;
      document.getElementById('resultsConclusionDiagnosis').value = (patient.conclusion && patient.conclusion.diagnosis) || '';
      document.getElementById('resultsConclusionTreatment').value = (patient.conclusion && patient.conclusion.treatment) || '';
      document.getElementById('resultsConclusionNote').value = (patient.conclusion && patient.conclusion.note) || '';
      Object.keys(resultGroups).forEach(category => renderResultsCategoryTable(patient, category));
      const approveResultsButton = document.getElementById('approveResultsButton');
      const saveResultsModalButton = document.getElementById('saveResultsModalButton');
      if (approveResultsButton) approveResultsButton.disabled = false;
      if (saveResultsModalButton) saveResultsModalButton.disabled = false;
      updateResultsStatusLabel(patient);
      const modal = new bootstrap.Modal(document.getElementById('resultsDetailModal'));
      modal.show();
    }

    function getActiveResultsCategory() {
      const activeTab = document.querySelector('#resultsDetailModal .nav-link.active');
      if (!activeTab) return 'hematology';
      const target = activeTab.dataset.bsTarget || '';
      return target.replace('#resultsTab', '').toLowerCase() || 'hematology';
    }

    function addResultsRow(category) {
      if (!currentResultsPatientId) return;
      const patient = patients.find(item => item.patientId === currentResultsPatientId);
      if (!patient) return;
      patient.results = patient.results || {};
      patient.results[category] = patient.results[category] || [];
      patient.results[category].push({ name: '', result: '', normalRange: '', unit: '' });
      renderResultsCategoryTable(patient, category);
    }

    function deleteResultsRow(category, index) {
      if (!currentResultsPatientId) return;
      const patient = patients.find(item => item.patientId === currentResultsPatientId);
      if (!patient || !patient.results || !patient.results[category]) return;
      patient.results[category].splice(index - (resultGroups[category].fixedFields || []).length, 1);
      renderResultsCategoryTable(patient, category);
    }

    function saveResultsModal() {
      if (!currentResultsPatientId) return;
      const patient = patients.find(item => item.patientId === currentResultsPatientId);
      if (!patient) return;
      patient.results = patient.results || {};
      Object.keys(resultGroups).forEach(category => {
        const body = document.getElementById(`results${category.charAt(0).toUpperCase() + category.slice(1)}Body`);
        if (!body) return;
        const customRows = [];
        Array.from(body.querySelectorAll('tr')).forEach(tr => {
          const inputs = tr.querySelectorAll('input');
          if (!inputs.length) return;
          const name = inputs[0].value.trim();
          const result = inputs[1].value.trim();
          const normalRange = inputs[2].value.trim();
          const unit = inputs[3].value.trim();
          const fixedKey = tr.dataset.fixedKey;
          const isFixed = tr.dataset.fixed === 'true';
          if (isFixed && fixedKey) {
            setPatientPendingLabValue(patient, category, fixedKey, result);
          } else if (name) {
            customRows.push({ name, result, normalRange, unit });
          }
        });
        patient.results[category] = customRows;
      });
      patient.conclusion = {
        diagnosis: document.getElementById('resultsConclusionDiagnosis').value.trim(),
        treatment: document.getElementById('resultsConclusionTreatment').value.trim(),
        note: document.getElementById('resultsConclusionNote').value.trim()
      };
      patient.examinationState = 'examined';
      patient.lastExamDate = new Date().toISOString().slice(0,10);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      updateResultsStatusLabel(patient);
      alert('Lưu kết quả xét nghiệm thành công.');
    }

    function isCurrentMonthDate(dateString) {
      if (!dateString) return false;
      const date = new Date(dateString);
      const now = new Date();
      return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
    }

    function getFilteredPatients() {
      const filters = {
        name: document.getElementById('searchName').value.trim().toLowerCase(),
        phone: document.getElementById('searchPhone').value.trim(),
        cccd: document.getElementById('searchCccd').value.trim(),
        code: document.getElementById('searchCode').value.trim().toLowerCase()
      };
      return patients.filter(patient => {
        const matchesSearch = (
          (!filters.name || patient.fullName.toLowerCase().includes(filters.name)) &&
          (!filters.phone || patient.phone.includes(filters.phone)) &&
          (!filters.cccd || patient.cccd.includes(filters.cccd)) &&
          (!filters.code || patient.patientId.toLowerCase().includes(filters.code))
        );
        if (!matchesSearch) return false;

        const currentReminder = (patient.reminders || [])
          .filter(r => r.revisitDate)
          .sort((a, b) => new Date(a.revisitDate) - new Date(b.revisitDate))[0];

        const isDoneOrContacted = currentReminder && (currentReminder.done || currentReminder.contacted);

        if (patientTableFilter === 'upcoming') {
          if (!patient.nextAppointment || isDoneOrContacted) return false;
          const diff = Math.ceil((new Date(patient.nextAppointment) - new Date()) / (1000 * 60 * 60 * 24));
          return diff >= 0 && diff <= 7;
        }
        if (patientTableFilter === 'overdue') {
          if (!patient.nextAppointment || isDoneOrContacted) return false;
          const diff = Math.ceil((new Date(patient.nextAppointment) - new Date()) / (1000 * 60 * 60 * 24));
          return diff < 0;
        }
        return true;
      });
    }

    function addMonths(date, months) {
      const d = new Date(date);
      const day = d.getDate();
      d.setMonth(d.getMonth() + months);
      // Handle month overflow
      if (d.getDate() !== day) {
        d.setDate(0);
      }
      return d;
    }

    function toggleSelectPatient(id) {
      if (selectedPatientIds.has(id)) selectedPatientIds.delete(id);
      else selectedPatientIds.add(id);
      updateSelectedCount();
      renderPatientTable();
    }

    function handleTopDelete() {
      if (!selectedPatientIds.size) return alert('Vui lòng chọn ít nhất một bệnh nhân để xóa.');
      if (!confirm(`Xác nhận xóa ${selectedPatientIds.size} bệnh nhân đã chọn?`)) return;
      deleteSelectedPatients();
    }

    function toggleSelectAll(checked) {
      const displayed = getFilteredPatients().map(p => p.patientId);
      if (checked) {
        displayed.forEach(id => selectedPatientIds.add(id));
      } else {
        displayed.forEach(id => selectedPatientIds.delete(id));
      }
      updateSelectedCount();
      renderPatientTable();
    }

    function updateSelectedCount() {
      const badge = document.getElementById('selectedCountBadge');
      const displayed = getFilteredPatients().map(p => p.patientId);
      const selectedCount = Array.from(selectedPatientIds).filter(id => displayed.includes(id)).length;
      if (badge) badge.textContent = `${selectedCount} đã chọn`;
      const selectAll = document.getElementById('selectAllCheckbox');
      if (selectAll) selectAll.checked = displayed.length > 0 && displayed.every(id => selectedPatientIds.has(id));
    }

    function deleteSelectedPatients() {
      const ids = Array.from(selectedPatientIds);
      if (!ids.length) return;
      // remove patients
      patients = patients.filter(p => !selectedPatientIds.has(p.patientId));
      // clear selection
      selectedPatientIds.clear();
      // persist and refresh UI
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      renderResultsPatientList();
      if (ids.includes(currentPatientId)) {
        currentPatientId = null;
      }
      updateSelectedCount();
      alert('Đã xóa bệnh nhân đã chọn.');
    }

    function toggleQuickRevisitInputs() {
      const useMonths = document.getElementById('quickRevisitModeMonths').checked;
      document.getElementById('quickRevisitMonthsGroup').classList.toggle('d-none', !useMonths);
      document.getElementById('quickRevisitDateGroup').classList.toggle('d-none', useMonths);
    }

    function openQuickRevisitModal() {
      const modalEl = document.getElementById('quickRevisitModal');
      if (!modalEl) return;
      document.getElementById('quickRevisitModeMonths').checked = true;
      document.getElementById('quickRevisitModeDate').checked = false;
      document.getElementById('quickRevisitDate').value = new Date().toISOString().slice(0, 10);
      toggleQuickRevisitInputs();
      const modal = new bootstrap.Modal(modalEl);
      modal.show();
    }

    async function applyQuickRevisit() {
      const mode = document.querySelector('input[name="quickRevisitMode"]:checked')?.value || 'months';
      const today = new Date();
      let revisitDate = '';
      if (mode === 'date') {
        revisitDate = document.getElementById('quickRevisitDate').value;
        if (!revisitDate) {
          alert('Vui lòng chọn ngày tái khám chính xác.');
          return;
        }
        const dateObj = new Date(revisitDate);
        if (isNaN(dateObj.getTime())) {
          alert('Ngày tái khám không hợp lệ.');
          return;
        }
      } else {
        const months = parseInt(document.getElementById('quickRevisitMonths').value, 10);
        if (!months || months <= 0) {
          alert('Vui lòng nhập số tháng hợp lệ.');
          return;
        }
        revisitDate = addMonths(today, months).toISOString().slice(0,10);
      }
      if (!selectedPatientIds.size) {
        alert('Vui lòng chọn ít nhất một bệnh nhân.');
        return;
      }
      const selectedIds = Array.from(selectedPatientIds);
      selectedIds.forEach(id => {
        const patient = patients.find(p => p.patientId === id);
        if (!patient) return;
        patient.reminders = patient.reminders || [];
        const newReminder = { revisitDate, content: 'Tái khám nhanh', note: '', contacted: false, done: false };
        if (patient.reminders.length) {
          archivePatientReminders(patient);
        }
        patient.reminders = [newReminder];
        syncNextAppointmentWithReminders(patient);
      });
      await saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId && selectedPatientIds.has(currentPatientId)) {
        const p = patients.find(x => x.patientId === currentPatientId);
        if (p) renderPatientReminders(p);
      }
      // clear selection
      selectedPatientIds.clear();
      updateSelectedCount();
      // hide modal
      const modalEl = document.getElementById('quickRevisitModal');
      const bsModal = bootstrap.Modal.getInstance(modalEl);
      if (bsModal) bsModal.hide();
      alert('Đã cập nhật tái khám cho các bệnh nhân đã chọn.');
    }

    function renderReminders() {
      const tbody = document.getElementById('reminderTableBody');
      tbody.innerHTML = '';
      const rows = [];
      const search = reminderSearchFilter.trim().toLowerCase();
      const selectedStatus = reminderStatusFilter;
      const reminderItems = [];

      patients.forEach((patient, patientIndex) => {
        getPatientReminderList(patient).forEach((reminder) => {
          if (reminderTableFilter === 'currentMonth' && !isCurrentMonthDate(reminder.revisitDate)) {
            return;
          }
          const status = getReminderStatus(reminder);
          if (selectedStatus === 'today' && status.label !== 'Tái khám hôm nay') return;
          if (selectedStatus === 'upcoming' && status.label !== 'Tái khám sắp tới') return;
          if (selectedStatus === 'overdue' && status.label !== 'Quá hạn') return;
          if (selectedStatus === 'no-schedule' && status.label !== 'Không có lịch') return;
          if (selectedStatus === 'contacted' && !reminder.contacted) return;
          if (selectedStatus === 'done' && !reminder.done) return;
          if (search) {
            const matchPatient = patient.patientId.toLowerCase().includes(search) || patient.fullName.toLowerCase().includes(search);
            if (!matchPatient) return;
          }
          reminderItems.push({ patient, patientIndex, reminder, status });
        });
      });

      reminderItems.sort((a, b) => {
        if (a.status.order !== b.status.order) return a.status.order - b.status.order;
        const dateA = a.reminder.revisitDate ? new Date(a.reminder.revisitDate).getTime() : Infinity;
        const dateB = b.reminder.revisitDate ? new Date(b.reminder.revisitDate).getTime() : Infinity;
        if (dateA !== dateB) return dateA - dateB;
        return a.patient.patientId.localeCompare(b.patient.patientId, 'vi');
      });

      reminderItems.forEach(({ patient, patientIndex, reminder, status }) => {
        const isSynthetic = reminder.isSynthetic;
        const patientReminderIndex = patient.reminders ? patient.reminders.findIndex(r => r === reminder) : -1;
        rows.push(`
          <tr>
            <td>${patient.patientId}</td>
            <td>${patient.fullName}</td>
            <td>${formatDate(reminder.revisitDate)}</td>
            <td>${reminder.content || '-'}</td>
            <td><span class="status-chip ${status.type}">${status.label}</span></td>
            <td>
              ${isSynthetic ? '<span class="text-muted">Tự động</span>' : `
                ${!reminder.revisitDate ? `<button class="btn btn-sm btn-warning me-1" onclick="promptSetRevisitDate(${patientIndex}, ${patientReminderIndex})">Đặt lịch</button>` : ''}
                <button class="btn btn-sm btn-success me-1" onclick="markReminderDone(${patientIndex}, ${patientReminderIndex})">Đã khám</button>
                <button class="btn btn-sm btn-outline-primary me-1" onclick="markReminderContact(${patientIndex}, ${patientReminderIndex})">Đã liên hệ</button>
                <button class="btn btn-sm btn-outline-danger" onclick="deleteReminder(${patientIndex}, ${patientReminderIndex})">Xóa lịch</button>
              `}
            </td>
          </tr>
        `);
      });

      tbody.innerHTML = rows.join('') || '<tr><td class="text-muted" colspan="6">Không có lịch tái khám</td></tr>';
    }

    function showCurrentMonthReminders() {
      reminderTableFilter = 'currentMonth';
      patientTableFilter = null;
      showSection('reminderSection');
      renderReminders();
    }

    function showPatientSectionWithFilter(filter) {
      patientTableFilter = filter || null;
      reminderTableFilter = null;
      showSection('patientSection');
      renderPatientTable();
    }

    function syncPatientReminderStatus(patient, contacted) {
      if (!patient || !patient.reminders) return;
      patient.reminders.forEach(reminder => {
        reminder.contacted = contacted;
      });
    }

    function archivePatientReminders(patient) {
      if (!patient || !patient.reminders || !patient.reminders.length) return;
      patient.history = patient.history || [];
      patient.reminders.forEach(reminder => {
        patient.history.push({
          date: reminder.revisitDate || reminder.examDate,
          diagnosis: reminder.content || 'Tái khám',
          result: reminder.contacted ? 'Đã liên hệ' : 'Đã khám',
          note: reminder.note || ''
        });
      });
      patient.reminders = [];
    }

    function toggleReminderContact(patientIndex, reminderIndex) {
      const patient = patients[patientIndex];
      if (!patient || !patient.reminders || !patient.reminders[reminderIndex]) return;
      const reminder = patient.reminders[reminderIndex];
      const newContactedState = !reminder.contacted;
      syncPatientReminderStatus(patient, newContactedState);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId === patient.patientId) {
        renderPatientReminders(patient);
      }
    }

    function resetPatientForm() {
      document.getElementById('patientModalTitle').textContent = 'Thêm bệnh nhân';
      document.getElementById('patientId').readOnly = false;
      ['patientId','fullName','birthDate','gender','healthInsurance','cccd','phone','ethnicity','job','address','note','lastExamDate','nextAppointment'].forEach(id => {
        document.getElementById(id).value = '';
      });
      currentPatientId = null;
    }

    function editPatient(id) {
      const patient = patients.find(item => item.patientId === id);
      if (!patient) return;
      currentPatientId = id;
      document.getElementById('patientModalTitle').textContent = 'Sửa bệnh nhân';
      document.getElementById('patientId').value = patient.patientId;
      document.getElementById('patientId').readOnly = true;
      document.getElementById('fullName').value = patient.fullName;
      document.getElementById('birthDate').value = patient.birthDate;
      document.getElementById('gender').value = patient.gender;
      document.getElementById('healthInsurance').value = patient.healthInsurance || '';
      document.getElementById('cccd').value = patient.cccd;
      document.getElementById('phone').value = patient.phone;
      document.getElementById('ethnicity').value = patient.ethnicity || '';
      document.getElementById('job').value = patient.job || '';
      document.getElementById('address').value = patient.address || '';
      document.getElementById('note').value = patient.note || '';
      document.getElementById('lastExamDate').value = patient.lastExamDate || '';
      document.getElementById('nextAppointment').value = patient.nextAppointment || '';
      const modal = new bootstrap.Modal(document.getElementById('patientModal'));
      modal.show();
    }

    function deletePatient(id) {
      if (!confirm('Bạn có chắc muốn xóa bệnh nhân này?')) return;
      patients = patients.filter(patient => patient.patientId !== id);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
    }

    function viewPatient(id) {
      const patient = patients.find(item => item.patientId === id);
      if (!patient) return;
      currentPatientId = id;
      document.getElementById('detailPatientId').value = patient.patientId;
      document.getElementById('detailFullName').value = patient.fullName;
      document.getElementById('detailBirthDate').value = patient.birthDate || '';
      document.getElementById('detailAge').value = calculateAge(patient.birthDate);
      document.getElementById('detailGender').value = patient.gender;
      document.getElementById('detailHealthInsurance').value = patient.healthInsurance || '';
      document.getElementById('detailCccd').value = patient.cccd;
      document.getElementById('detailPhone').value = patient.phone;
      document.getElementById('detailEthnicity').value = patient.ethnicity || '';
      document.getElementById('detailJob').value = patient.job || '';
      document.getElementById('detailAddress').value = patient.address || '';
      document.getElementById('detailNote').value = patient.note || '';
      document.getElementById('detailLastExamDate').value = patient.lastExamDate || '';
      document.getElementById('detailNextAppointment').value = patient.nextAppointment || '';
      document.getElementById('detailPatientLabel').textContent = `${patient.patientId} - ${patient.fullName}`;
      document.querySelectorAll('#detailSection input, #detailSection textarea').forEach(el => {
        if (el.id === 'detailPatientId' || el.id.endsWith('Prev')) return;
        el.disabled = false;
      });
      ['detailGender', 'detailHealthInsurance', 'detailCccd', 'detailPhone', 'detailEthnicity', 'detailJob', 'detailAddress', 'detailNote', 'detailBirthDate', 'detailLastExamDate', 'detailNextAppointment'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.disabled = false;
      });
      const sections = [
        { key: 'clinical', ids: ['Height','Weight','Bmi','Bp','Pulse','Temp','Spo2'] },
        { key: 'hematology', ids: ['Wbc','Neu','Lym','Mono','Eos','Baso','NeuPercent','LymPercent','MonoPercent','EosPercent','BasoPercent','Rbc','Hgb','Hct','Mcv','Mch','Mchc','Rdw','RdwSd','Plt','Mpv','Pdw','Pct','Plcr','Plcc'] },
        { key: 'biochemistry', ids: ['Glucose','Hba1c','Ure','Creatinine','Ast','Alt','Ggt','Bilirubin','Cholesterol','Triglyceride','Hdl','Ldl','UricAcid','TotalProtein','Albumin','Globulin','AgRatio','Ldh','Calcium','Phosphate','Magnesium'] },
        { key: 'immunology', ids: ['Hbsag','AntiHbs','Hcv','Hiv','Afp','Cea','Psa','Ca125','Ca199','Ca153','Crp','Rf','Ana','Igg','Igm','Iga','Tsh','T3','T4','Ft3','Ft4','Cortisol'] },
        { key: 'coagulation', ids: ['PtS','PtPercent','Inr','Aptt','Fibrinogen'] },
        { key: 'urine', ids: ['Leu','Nit','Ubg','Bil','Pro','Ph','Bld','Ket','Glu','Sg'] }
      ];
      const getDataKey = (sectionKey, id) => {
        const normalized = id.charAt(0).toLowerCase() + id.slice(1);
        if (sectionKey === 'coagulation') {
          if (normalized === 'ptS') return 'pt_s';
          if (normalized === 'ptPercent') return 'pt_percent';
        }
        return normalized;
      };
      sections.forEach(section => {
        const data = patient[section.key] || {};
        section.ids.forEach(id => {
          const element = document.getElementById(`${section.key}${id}`);
          if (element) {
            const dataKey = getDataKey(section.key, id);
            element.value = data[dataKey] || data[id.toLowerCase()] || '';
            setResultInputColor(element);
          }
        });
      });
      const prevSections = [
        { key: 'hematology', ids: ['Wbc','Neu','Lym','Mono','Eos','Baso','NeuPercent','LymPercent','MonoPercent','EosPercent','BasoPercent','Rbc','Hgb','Hct','Mcv','Mch','Mchc','Rdw','RdwSd','Plt','Mpv','Pdw','Pct','Plcr','Plcc'] },
        { key: 'biochemistry', ids: ['Glucose','Hba1c','Ure','Creatinine','Ast','Alt','Ggt','Bilirubin','Cholesterol','Triglyceride','Hdl','Ldl','UricAcid','TotalProtein','Albumin','Globulin','AgRatio','Ldh','Calcium','Phosphate','Magnesium'] },
        { key: 'immunology', ids: ['Hbsag','AntiHbs','Hcv','Hiv','Afp','Cea','Psa','Ca125','Ca199','Ca153','Crp','Rf','Ana','Igg','Igm','Iga','Tsh','T3','T4','Ft3','Ft4','Cortisol'] },
        { key: 'coagulation', ids: ['PtS','PtPercent','Inr','Aptt','Fibrinogen'] },
        { key: 'urine', ids: ['Leu','Nit','Ubg','Bil','Pro','Ph','Bld','Ket','Glu','Sg'] }
      ];
      const getPrevDataKey = (sectionKey, id) => {
        const normalized = id.charAt(0).toLowerCase() + id.slice(1);
        if (sectionKey === 'coagulation') {
          if (normalized === 'ptS') return 'pt_s';
          if (normalized === 'ptPercent') return 'pt_percent';
        }
        return normalized;
      };
      prevSections.forEach(section => {
        const prevData = patient[`previous${section.key.charAt(0).toUpperCase() + section.key.slice(1)}`] || {};
        section.ids.forEach(id => {
          const prevElement = document.getElementById(`${section.key}${id}Prev`);
          if (prevElement) {
            const prevKey = getPrevDataKey(section.key, id);
            prevElement.value = prevData[prevKey] || prevData[id.toLowerCase()] || '';
          }
        });
      });
      ['detailGender', 'detailHealthInsurance', 'detailCccd', 'detailPhone', 'detailEthnicity', 'detailJob', 'detailAddress', 'detailNote', 'detailBirthDate', 'detailLastExamDate', 'detailNextAppointment'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.disabled = false;
      });
      document.querySelectorAll('#tabHematology input:not([id$="Prev"]), #tabBiochemistry input:not([id$="Prev"]), #tabImmunology input:not([id$="Prev"]), #tabCoagulation input:not([id$="Prev"]), #tabUrine input:not([id$="Prev"])').forEach(element => {
        element.disabled = true;
      });
      document.getElementById('conclusionDiagnosis').value = (patient.conclusion && patient.conclusion.diagnosis) || '';
      document.getElementById('conclusionTreatment').value = (patient.conclusion && patient.conclusion.treatment) || '';
      document.getElementById('conclusionNote').value = (patient.conclusion && patient.conclusion.note) || '';
      document.getElementById('conclusionTime').value = (patient.conclusion && patient.conclusion.time) || '';
      renderHistory(patient.history || []);
      renderPatientReminders(patient);
      resetReminderForm();
      // Show detail section
      showSection('detailSection');
    }

    function savePatientDetails() {
      if (!currentPatientId) return;
      const patient = patients.find(item => item.patientId === currentPatientId);
      if (!patient) return;
      const getDataKey = (sectionKey, id) => {
        const normalized = id.charAt(0).toLowerCase() + id.slice(1);
        if (sectionKey === 'coagulation') {
          if (normalized === 'ptS') return 'pt_s';
          if (normalized === 'ptPercent') return 'pt_percent';
        }
        return normalized;
      };
      const readSectionValues = (sectionKey, ids) => {
        const values = {};
        ids.forEach(id => {
          const element = document.getElementById(`${sectionKey}${id}`);
          if (!element) return;
          values[getDataKey(sectionKey, id)] = element.value.trim();
        });
        return values;
      };
      patient.fullName = document.getElementById('detailFullName').value.trim();
      patient.birthDate = document.getElementById('detailBirthDate').value || patient.birthDate;
      patient.gender = document.getElementById('detailGender').value.trim();
      patient.healthInsurance = document.getElementById('detailHealthInsurance').value.trim();
      patient.cccd = document.getElementById('detailCccd').value.trim();
      patient.phone = document.getElementById('detailPhone').value.trim();
      patient.ethnicity = document.getElementById('detailEthnicity').value.trim();
      patient.job = document.getElementById('detailJob').value.trim();
      patient.address = document.getElementById('detailAddress').value.trim();
      patient.note = document.getElementById('detailNote').value.trim();
      patient.lastExamDate = document.getElementById('detailLastExamDate').value || patient.lastExamDate;
      patient.nextAppointment = document.getElementById('detailNextAppointment').value || patient.nextAppointment;
      ensureReminderForNextAppointment(patient);
      syncNextAppointmentWithReminders(patient);
      const getOptionalValueExisting = (id, existingValue = '') => {
        const el = document.getElementById(id);
        return el ? el.value.trim() : existingValue;
      };
      patient.clinical = {
        height: getOptionalValueExisting('clinicalHeight', patient.clinical?.height || ''),
        weight: getOptionalValueExisting('clinicalWeight', patient.clinical?.weight || ''),
        bmi: getOptionalValueExisting('clinicalBmi', patient.clinical?.bmi || ''),
        bp: getOptionalValueExisting('clinicalBp', patient.clinical?.bp || ''),
        pulse: getOptionalValueExisting('clinicalPulse', patient.clinical?.pulse || ''),
        temp: getOptionalValueExisting('clinicalTemp', patient.clinical?.temp || ''),
        spo2: getOptionalValueExisting('clinicalSpo2', patient.clinical?.spo2 || '')
      };
      patient.hematology = readSectionValues('hematology', ['Wbc','Neu','Lym','Mono','Eos','Baso','NeuPercent','LymPercent','MonoPercent','EosPercent','BasoPercent','Rbc','Hgb','Hct','Mcv','Mch','Mchc','Rdw','RdwSd','Plt','Mpv','Pdw','Pct','Plcr','Plcc']);
      patient.biochemistry = readSectionValues('biochemistry', ['Glucose','Hba1c','Ure','Creatinine','Ast','Alt','Ggt','Bilirubin','Cholesterol','Triglyceride','Hdl','Ldl','UricAcid','TotalProtein','Albumin','Globulin','AgRatio','Ldh','Calcium','Phosphate','Magnesium']);
      patient.immunology = readSectionValues('immunology', ['Hbsag','AntiHbs','Hcv','Hiv','Afp','Cea','Psa','Ca125','Ca199','Ca153','Crp','Rf','Ana','Igg','Igm','Iga','Tsh','T3','T4','Ft3','Ft4','Cortisol']);
      patient.coagulation = readSectionValues('coagulation', ['PtS','PtPercent','Inr','Aptt','Fibrinogen']);
      patient.urine = readSectionValues('urine', ['Leu','Nit','Ubg','Bil','Pro','Ph','Bld','Ket','Glu','Sg']);
      patient.conclusion = {
        diagnosis: document.getElementById('conclusionDiagnosis').value.trim(),
        treatment: document.getElementById('conclusionTreatment').value.trim(),
        note: document.getElementById('conclusionNote').value.trim(),
        time: document.getElementById('conclusionTime').value || (patient.conclusion && patient.conclusion.time) || ''
      };
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      renderPatientReminders(patient);
      document.getElementById('detailPatientLabel').textContent = `${patient.patientId} - ${patient.fullName}`;
      alert('Đã lưu cập nhật hồ sơ bệnh nhân.');
    }

    function backToPatient() {
      currentPatientId = null;
      showSection('patientSection');
    }

    function renderHistory(history = []) {
      const tbody = document.getElementById('historyTableBody');
      if (!history.length) {
        tbody.innerHTML = '<tr><td colspan="5" class="text-muted">Chưa có lịch sử khám</td></tr>';
        return;
      }
      tbody.innerHTML = history.sort((a, b) => new Date(b.date) - new Date(a.date)).map((item, index) => `
        <tr>
          <td>${formatDate(item.date)}</td>
          <td>${item.diagnosis || '-'}</td>
          <td>${item.result || '-'}</td>
          <td>${item.note || '-'}</td>
          <td><button class="btn btn-sm btn-danger" onclick="deleteHistory('${formatDate(item.date)}', ${index})"><i class="fa-solid fa-trash"></i></button></td>
        </tr>
      `).join('');
    }

    function deleteHistory(dateStr, index) {
      if (!currentPatientId) return;
      if (!confirm('Bạn có chắc muốn xóa lịch sử khám này?')) return;
      const patient = patients.find(item => item.patientId === currentPatientId);
      if (!patient || !patient.history) return;
      patient.history = patient.history.filter(h => formatDate(h.date) !== dateStr);
      saveData();
      renderHistory(patient.history);
    }

    function getReminderStatus(reminder) {
      if (reminder.done) return { label: 'Đã khám', type: 'green', days: '-', order: 5 };
      if (reminder.contacted) return { label: 'Đã liên hệ', type: 'green', days: '-', order: 4 };
      if (!reminder.revisitDate) return { label: 'Không có lịch', type: 'gray', days: '-', order: 3 };
      const diff = Math.ceil((new Date(reminder.revisitDate) - new Date()) / (1000 * 60 * 60 * 24));
      if (diff < 0) return { label: 'Quá hạn', type: 'red', days: `${Math.abs(diff)} ngày`, order: 2 };
      if (diff === 0) return { label: 'Tái khám hôm nay', type: 'orange', days: '0 ngày', order: 0 };
      if (diff <= 7) return { label: 'Tái khám sắp tới', type: 'yellow', days: `${diff} ngày`, order: 1 };
      const label = formatAppointmentLabel(diff);
      return { label: `Tái khám sau ${label}`, type: 'green', days: `${diff} ngày`, order: 3 };
    }

    function sortRemindersByStatusAndDate(reminders = []) {
      return reminders.slice().sort((a, b) => {
        const statusA = getReminderStatus(a);
        const statusB = getReminderStatus(b);
        if (statusA.order !== statusB.order) return statusA.order - statusB.order;
        const dateA = a.revisitDate ? new Date(a.revisitDate).getTime() : Infinity;
        const dateB = b.revisitDate ? new Date(b.revisitDate).getTime() : Infinity;
        return dateA - dateB;
      });
    }

    function markReminderDone(patientIndex, reminderIndex) {
      const patient = patients[patientIndex];
      if (!patient || !patient.reminders || !patient.reminders[reminderIndex]) return;
      const reminder = patient.reminders[reminderIndex];
      if (reminder.done) {
        reminder.done = false;
        reminder.contacted = false;
      } else {
        reminder.done = true;
        reminder.contacted = false;
      }
      syncNextAppointmentWithReminders(patient);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId === patient.patientId) {
        renderPatientReminders(patient);
      }
    }

    function markReminderContact(patientIndex, reminderIndex) {
      const patient = patients[patientIndex];
      if (!patient || !patient.reminders || !patient.reminders[reminderIndex]) return;
      const reminder = patient.reminders[reminderIndex];
      if (reminder.contacted) {
        reminder.contacted = false;
        reminder.done = false;
      } else {
        reminder.contacted = true;
        reminder.done = false;
      }
      syncNextAppointmentWithReminders(patient);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId === patient.patientId) {
        renderPatientReminders(patient);
      }
    }

    function deleteReminder(patientIndex, reminderIndex) {
      if (!confirm('Bạn có chắc muốn xóa lịch tái khám này?')) return;
      const patient = patients[patientIndex];
      if (!patient || !patient.reminders || !patient.reminders[reminderIndex]) return;
      patient.reminders.splice(reminderIndex, 1);
      if (!patient.reminders.length) {
        patient.reminders = [];
      }
      syncNextAppointmentWithReminders(patient);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId === patient.patientId) {
        renderPatientReminders(patient);
      }
    }

    function promptSetRevisitDate(patientIndex, reminderIndex) {
      const patient = patients[patientIndex];
      if (!patient || !patient.reminders || !patient.reminders[reminderIndex]) return;
      const reminder = patient.reminders[reminderIndex];
      const newDate = prompt('Nhập ngày tái khám (YYYY-MM-DD):', '');
      if (!newDate || !newDate.trim()) return;
      const dateObj = new Date(newDate);
      if (isNaN(dateObj.getTime())) {
        alert('Định dạng ngày không hợp lệ. Vui lòng nhập YYYY-MM-DD');
        return;
      }
      reminder.revisitDate = newDate;
      syncNextAppointmentWithReminders(patient);
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      if (currentPatientId === patient.patientId) {
        renderPatientReminders(patient);
      }
    }

    function renderPatientReminders(patientOrReminders = []) {
      const tbody = document.getElementById('reminderDetailTableBody');
      const nextLabelElement = document.getElementById('reminderNextLabel');
      if (!tbody || !nextLabelElement) return;
      let reminders = [];
      if (Array.isArray(patientOrReminders)) {
        reminders = patientOrReminders.slice();
      } else if (patientOrReminders && typeof patientOrReminders === 'object') {
        reminders = getPatientReminderList(patientOrReminders);
      }
      if (!reminders.length) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-muted">Chưa có nhắc tái khám</td></tr>';
        nextLabelElement.textContent = 'Không có lịch tái khám';
        return;
      }
      const sorted = sortRemindersByStatusAndDate(reminders);
      const pendingReminder = sorted.find(reminder => !reminder.done && !reminder.contacted) || sorted[0];
      const nextLabel = pendingReminder.done ? 'Đã khám' : (pendingReminder && pendingReminder.revisitDate ? `Tái khám: ${formatDate(pendingReminder.revisitDate)}` : 'Không có lịch tái khám');
      nextLabelElement.textContent = nextLabel;
      tbody.innerHTML = sorted.map(reminder => {
        const statusItem = getReminderStatus(reminder);
        return `
          <tr>
            <td>${formatDate(reminder.revisitDate)}</td>
            <td>${reminder.content || '-'}</td>
            <td><span class="status-chip ${statusItem.type}">${statusItem.label}</span></td>
            <td>${statusItem.days}</td>
          </tr>
        `;
      }).join('');
    }

    function resetReminderForm() {
      ['reminderRevisitDate','reminderContent','reminderNote'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.value = '';
      });
    }

    function getReceptionTestCheckbox(category, key) {
      const prefix = category.charAt(0).toUpperCase() + category.slice(1);
      const suffix = key ? `_${key}` : 'All';
      return document.getElementById(`receptionTest${prefix}${suffix}`);
    }

    function updateReceptionGroupCheckbox(category) {
      const definition = receptionTestDefinitions[category] || [];
      if (!definition.length) return;
      const allCheckbox = getReceptionTestCheckbox(category);
      const individualCheckboxes = definition.map(key => getReceptionTestCheckbox(category, key));
      const checkedCount = individualCheckboxes.filter(cb => cb && cb.checked).length;
      if (allCheckbox) {
        allCheckbox.checked = checkedCount === definition.length;
      }
    }

    function updateReceptionIndividualCheckboxes(category, checked) {
      const definition = receptionTestDefinitions[category] || [];
      definition.forEach(key => {
        const cb = getReceptionTestCheckbox(category, key);
        if (cb) cb.checked = checked;
      });
    }

    function showReceptionTestSelection(patientId) {
      const patient = patients.find(item => item.patientId === patientId);
      if (!patient) return;
      const panel = document.getElementById('receptionTestSelection');
      const label = document.getElementById('receptionSelectedPatientLabel');
      if (label) label.textContent = `${patient.patientId} - ${patient.fullName}`;
      if (panel) panel.style.display = 'block';
      const existing = patient.testsSelection || {};
      const biochemAll = document.getElementById('receptionTestBiochemistryAll');
      const immunoAll = document.getElementById('receptionTestImmunologyAll');
      const hematologyAll = document.getElementById('receptionTestHematologyAll');
      const coagulationAll = document.getElementById('receptionTestCoagulationAll');
      const urineAll = document.getElementById('receptionTestUrineAll');
      if (biochemAll) biochemAll.checked = Array.isArray(existing.biochemistry) && existing.biochemistry.length === receptionTestDefinitions.biochemistry.length;
      if (immunoAll) immunoAll.checked = Array.isArray(existing.immunology) && existing.immunology.length === receptionTestDefinitions.immunology.length;
      if (hematologyAll) hematologyAll.checked = Array.isArray(existing.hematology) && existing.hematology.length === receptionTestDefinitions.hematology.length;
      if (coagulationAll) coagulationAll.checked = Array.isArray(existing.coagulation) && existing.coagulation.length === receptionTestDefinitions.coagulation.length;
      if (urineAll) urineAll.checked = Array.isArray(existing.urine) && existing.urine.length === receptionTestDefinitions.urine.length;
      receptionTestDefinitions.biochemistry.forEach(key => {
        const cb = getReceptionTestCheckbox('biochemistry', key);
        if (cb) cb.checked = Array.isArray(existing.biochemistry) ? existing.biochemistry.includes(key) : false;
      });
      receptionTestDefinitions.immunology.forEach(key => {
        const cb = getReceptionTestCheckbox('immunology', key);
        if (cb) cb.checked = Array.isArray(existing.immunology) ? existing.immunology.includes(key) : false;
      });
      receptionTestDefinitions.hematology.forEach(key => {
        const cb = getReceptionTestCheckbox('hematology', key);
        if (cb) cb.checked = Array.isArray(existing.hematology) ? existing.hematology.includes(key) : false;
      });
      receptionTestDefinitions.coagulation.forEach(key => {
        const cb = getReceptionTestCheckbox('coagulation', key);
        if (cb) cb.checked = Array.isArray(existing.coagulation) ? existing.coagulation.includes(key) : false;
      });
      receptionTestDefinitions.urine.forEach(key => {
        const cb = getReceptionTestCheckbox('urine', key);
        if (cb) cb.checked = Array.isArray(existing.urine) ? existing.urine.includes(key) : false;
      });
    }

    function hideReceptionTestSelection() {
      const panel = document.getElementById('receptionTestSelection');
      if (panel) panel.style.display = 'none';
      ['receptionTestBiochemistryAll','receptionTestImmunologyAll','receptionTestHematologyAll','receptionTestCoagulationAll','receptionTestUrineAll'].forEach(id => {
        const cb = document.getElementById(id);
        if (cb) cb.checked = false;
      });
      receptionTestDefinitions.biochemistry.forEach(key => {
        const cb = getReceptionTestCheckbox('biochemistry', key);
        if (cb) cb.checked = false;
      });
      receptionTestDefinitions.immunology.forEach(key => {
        const cb = getReceptionTestCheckbox('immunology', key);
        if (cb) cb.checked = false;
      });
      receptionTestDefinitions.hematology.forEach(key => {
        const cb = getReceptionTestCheckbox('hematology', key);
        if (cb) cb.checked = false;
      });
      receptionTestDefinitions.coagulation.forEach(key => {
        const cb = getReceptionTestCheckbox('coagulation', key);
        if (cb) cb.checked = false;
      });
      receptionTestDefinitions.urine.forEach(key => {
        const cb = getReceptionTestCheckbox('urine', key);
        if (cb) cb.checked = false;
      });
      pendingReceptionPatientId = null;
    }

    function toggleReceptionMode(mode) {
      const newPanel = document.getElementById('receptionModeNew');
      const revisitPanel = document.getElementById('receptionModeRevisit');
      const newBtn = document.getElementById('receptionModeNewBtn');
      const revisitBtn = document.getElementById('receptionModeRevisitBtn');
      if (mode === 'revisit') {
        if (newPanel) newPanel.style.display = 'none';
        if (revisitPanel) revisitPanel.style.display = 'block';
        if (newBtn) newBtn.classList.remove('active');
        if (revisitBtn) revisitBtn.classList.add('active');
      } else {
        if (newPanel) newPanel.style.display = 'block';
        if (revisitPanel) revisitPanel.style.display = 'none';
        if (newBtn) newBtn.classList.add('active');
        if (revisitBtn) revisitBtn.classList.remove('active');
      }
    }

    function confirmRevisit(patientIdInput) {
      const pid = (patientIdInput || document.getElementById('r_revisitPatientId')).value?.trim();
      if (!pid) return alert('Vui lòng nhập Mã BN.');
      const patient = patients.find(p => p.patientId === pid);
      if (!patient) return alert('Không tìm thấy bệnh nhân với mã này.');
      // mark upcoming reminders (today and next 7 days) as done
      if (patient.reminders && Array.isArray(patient.reminders)) {
        const now = new Date();
        patient.reminders.forEach(r => {
          if (!r.revisitDate) return;
          const diff = Math.ceil((new Date(r.revisitDate) - new Date()) / (1000*60*60*24));
          if (diff >= 0 && diff <= 7) {
            r.done = true;
            r.contacted = false;
          }
        });
      }
      // update lastExamDate to today
      patient.lastExamDate = new Date().toISOString().slice(0,10);
      const status = getStatus(patient);
      if (status.label === 'Hôm nay' || status.label === 'Sắp tới hạn') {
        patient.examinationState = 'in-progress';
      }
      syncNextAppointmentWithReminders(patient);
      pendingReceptionPatientId = patient.patientId;
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      showReceptionTestSelection(patient.patientId);
      alert('Đã xác nhận tái khám và cập nhật trạng thái lịch hẹn. Bạn có thể chọn xét nghiệm.');
    }

    function saveReceptionTestSelection() {
      if (!pendingReceptionPatientId) return;
      const patient = patients.find(item => item.patientId === pendingReceptionPatientId);
      if (!patient) return;
      const selection = {
        biochemistry: receptionTestDefinitions.biochemistry.filter(key => getReceptionTestCheckbox('biochemistry', key)?.checked),
        immunology: receptionTestDefinitions.immunology.filter(key => getReceptionTestCheckbox('immunology', key)?.checked),
        hematology: receptionTestDefinitions.hematology.filter(key => getReceptionTestCheckbox('hematology', key)?.checked),
        coagulation: receptionTestDefinitions.coagulation.filter(key => getReceptionTestCheckbox('coagulation', key)?.checked),
        urine: receptionTestDefinitions.urine.filter(key => getReceptionTestCheckbox('urine', key)?.checked)
      };
      if (!selection.hematology.length && !selection.coagulation.length && !selection.urine.length && !selection.biochemistry.length && !selection.immunology.length) {
        alert('Vui lòng chọn ít nhất một xét nghiệm.');
        return;
      }
      patient.testsSelection = selection;
      patient.pendingResults = {};
      patient.examinationState = 'in-progress';
      patient.resultsReturned = false;
      patient.resultsApproved = false;
      saveData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
      // Refresh results list and detail view if needed so UI stays in sync
      renderResultsPatientList();
      if (currentResultsPatientId === patient.patientId) showResultsDetail(currentResultsPatientId);
      hideReceptionTestSelection();
      alert('Đã lưu lựa chọn xét nghiệm.');
    }

    // Results section - New layout functions
    function getTestInfoFromSelection(patient) {
      if (!patient) return [];
      const tests = [];
      const categories = ['hematology','biochemistry','immunology','coagulation','urine'];
      const selection = patient.testsSelection || {};
      const effectiveData = {
        hematology: getPatientEffectiveLabData(patient, 'hematology'),
        biochemistry: getPatientEffectiveLabData(patient, 'biochemistry'),
        immunology: getPatientEffectiveLabData(patient, 'immunology'),
        coagulation: getPatientEffectiveLabData(patient, 'coagulation'),
        urine: getPatientEffectiveLabData(patient, 'urine')
      };

      categories.forEach(category => {
        const selectedKeys = Array.isArray(selection[category]) ? selection[category].slice() : [];
        const effectiveKeys = Object.keys(effectiveData[category] || {}).filter(key => effectiveData[category][key] !== undefined && effectiveData[category][key] !== '');
        const allKeys = [];

        (resultGroups[category].fixedFields || []).forEach(field => {
          if (selectedKeys.includes(field.key) || effectiveKeys.includes(field.key)) {
            allKeys.push(field.key);
          }
        });

        allKeys.forEach(key => {
          const field = resultGroups[category].fixedFields.find(f => f.key === key);
          if (field) {
            tests.push({
              category,
              key: field.key,
              title: field.title,
              normal: field.normal,
              unit: field.unit
            });
          }
        });
      });
      return tests;
    }

    function getActiveResultsCategories(patient) {
      if (!patient) return [];
      const categories = [];
      const selection = patient.testsSelection || {};
      // Only show input fields for categories that have testsSelection (actively entering results)
      ['hematology','biochemistry','immunology','coagulation','urine'].forEach(category => {
        if (Array.isArray(selection[category]) && selection[category].length > 0) {
          categories.push(category);
        }
      });
      return categories;
    }

    function renderResultsPatientList() {
      const container = document.getElementById('resultsPatientListContainer');
      if (!container) return;
      const filteredPatients = getResultsFilteredPatients();
      
      container.innerHTML = filteredPatients.map(patient => {
        const isActive = patient.patientId === currentResultsPatientId ? 'active' : '';
        const hasAny = patientHasResults(patient);
        const complete = isResultsComplete(patient);
        let color = '#0d6efd';
        let title = 'Chưa nhập kết quả';
        if (!hasAny) { color = '#0d6efd'; title = 'Chưa nhập gì'; }
        else if (complete) {
          if (patient.resultsReturned) { color = '#198754'; title = 'Đã nhập đầy đủ và đã duyệt'; }
          else { color = '#ffc107'; title = 'Đã nhập đầy đủ, chưa duyệt'; }
        } else { color = '#dc3545'; title = 'Nhập chưa đủ kết quả'; }
        const statusDot = `<span class="me-2 d-inline-block rounded-circle" style="width:12px;height:12px;background:${color};" title="${title}"></span>`;
        return `
          <a href="#" class="list-group-item list-group-item-action ${isActive}" onclick="event.preventDefault(); showResultsDetail('${patient.patientId}')">
            <div class="d-flex w-100 justify-content-between align-items-start">
              <div class="flex-grow-1 d-flex align-items-start">
                ${statusDot}
                <div class="flex-grow-1">
                  <h6 class="mb-1 text-truncate">${patient.fullName}</h6>
                  <small class="text-muted d-block text-truncate">${patient.patientId}</small>
                </div>
              </div>
              ${patient.nextAppointment ? `<span class="badge bg-success-subtle text-success text-nowrap ms-2">${formatDate(patient.nextAppointment)}</span>` : ''}
            </div>
          </a>
        `;
      }).join('');
      
      if (!filteredPatients.length) {
        container.innerHTML = '<div class="text-muted text-center py-3"><small>Không tìm thấy bệnh nhân</small></div>';
      }
    }

    function showResultsDetail(patientId) {
      const patient = patients.find(item => item.patientId === patientId);
      if (!patient) return;
      
      currentResultsPatientId = patientId;
      renderResultsPatientList();
      
      const placeholder = document.getElementById('resultsDetailPlaceholder');
      const detailContainer = document.getElementById('resultsDetailContainer');
      const tabsContainer = document.getElementById('resultsDetailTabs');
      const tabContent = document.getElementById('resultsDetailTabContent');
      
      placeholder.style.display = 'none';
      detailContainer.style.display = 'block';
      
      document.getElementById('resultsPatientLabel').textContent = `${patient.patientId} - ${patient.fullName}`;
      document.getElementById('resultsPatientInfo').textContent = 
        `${patient.phone || '-'} • ${patient.cccd || '-'} • Ngày khám: ${patient.lastExamDate ? formatDate(patient.lastExamDate) : '-'} • ${patient.resultsReturned ? 'Đã trả' : 'Chưa trả'}`;
      const approveResultsButton = document.getElementById('approveResultsButton');
      const approveResultsInlineButton = document.getElementById('approveResultsInlineButton');
      const saveResultsModalButton = document.getElementById('saveResultsModalButton');
      if (approveResultsButton) approveResultsButton.disabled = false;
      if (approveResultsInlineButton) approveResultsInlineButton.disabled = false;
      if (saveResultsModalButton) saveResultsModalButton.disabled = false;
      updateResultsStatusLabel(patient);
      
      // Get active categories based on testsSelection
      const categories = getActiveResultsCategories(patient);
      const tests = getTestInfoFromSelection(patient);
      
      // If no active input categories, clear detail view
      if (categories.length === 0) {
        tabsContainer.innerHTML = '';
        tabContent.innerHTML = `<div class="text-center text-muted py-5"><p>Chọn xét nghiệm để nhập kết quả</p></div>`;
        return;
      }
      
      // Create tabs
      tabsContainer.innerHTML = categories.map((cat, idx) => {
        const isActive = idx === 0 ? 'active' : '';
        const categoryName = resultGroups[cat]?.title || cat;
        return `
          <li class="nav-item" role="presentation">
            <button class="nav-link ${isActive}" id="resultsTab${cat}" data-bs-toggle="tab" 
                    data-bs-target="#resultsTabContent${cat}" type="button" role="tab">
              ${categoryName}
            </button>
          </li>
        `;
      }).join('');
      
      // Create tab content
      tabContent.innerHTML = categories.map((cat, idx) => {
        const isActive = idx === 0 ? 'show active' : '';
        const categoryTests = tests.filter(t => t.category === cat);
        const tableRows = categoryTests.map(test => {
          const result = getPatientLabValue(patient, cat, test.key) || '';
          return `
            <tr>
              <td class="fw-500">${test.title}</td>
              <td><input class="form-control form-control-sm" value="${result}" data-category="${cat}" data-key="${test.key}" onchange="saveTestResult(this)" /></td>
              <td class="text-nowrap"><small>${test.normal}</small></td>
              <td class="text-nowrap"><small>${test.unit}</small></td>
            </tr>
          `;
        }).join('');
        
        return `
          <div class="tab-pane fade ${isActive}" id="resultsTabContent${cat}" role="tabpanel">
            <div class="table-responsive">
              <table class="table table-bordered align-middle mb-0 table-sm">
                <thead class="table-light">
                  <tr>
                    <th style="width: 25%;">Thông số</th>
                    <th style="width: 25%;">Kết quả</th>
                    <th style="width: 25%;">Giá trị bình thường</th>
                    <th style="width: 25%;">Đơn vị</th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRows || '<tr><td colspan="4" class="text-muted">Không có dữ liệu</td></tr>'}
                </tbody>
              </table>
            </div>
          </div>
        `;
      }).join('');
    }

    function saveTestResult(input) {
      if (!currentResultsPatientId) return;
      const patient = patients.find(item => item.patientId === currentResultsPatientId);
      if (!patient) return;
      
      const category = input.dataset.category;
      const key = input.dataset.key;
      const value = input.value.trim();
      
      setPatientPendingLabValue(patient, category, key, value);
      saveData();
      // Update status label in real-time
      updateResultsStatusLabel(patient);
    }

    function attachEvents() {
      document.getElementById('patientForm').addEventListener('submit', event => {
        event.preventDefault();
        const patient = {
          patientId: document.getElementById('patientId').value.trim(),
          fullName: document.getElementById('fullName').value.trim(),
          birthDate: document.getElementById('birthDate').value,
          gender: document.getElementById('gender').value,
          healthInsurance: document.getElementById('healthInsurance').value.trim(),
          cccd: document.getElementById('cccd').value.trim(),
          phone: document.getElementById('phone').value.trim(),
          ethnicity: document.getElementById('ethnicity').value.trim(),
          job: document.getElementById('job').value.trim(),
          address: document.getElementById('address').value.trim(),
          note: document.getElementById('note').value.trim(),
          lastExamDate: document.getElementById('lastExamDate').value,
          nextAppointment: document.getElementById('nextAppointment').value,
          clinical: {}, hematology: {}, biochemistry: {}, immunology: {}, coagulation: {}, urine: {}, conclusion: {}, history: [], reminders: []
        };
        if (currentPatientId) {
          const index = patients.findIndex(item => item.patientId === currentPatientId);
          if (index !== -1) {
            const existing = patients[index];
            patient.clinical = existing.clinical || {};
            patient.hematology = existing.hematology || {};
            patient.biochemistry = existing.biochemistry || {};
            patient.immunology = existing.immunology || {};
            patient.conclusion = existing.conclusion || {};
            patient.history = existing.history || [];
            patient.reminders = existing.reminders || [];
            patients[index] = { ...existing, ...patient };
          }
        } else {
          if (patients.some(item => item.patientId === patient.patientId)) {
            alert('Mã bệnh nhân đã tồn tại.');
            return;
          }
          patients.push(patient);
        }
        
        // Sync appointment to reminders
        ensureReminderForNextAppointment(patient);
        syncNextAppointmentWithReminders(patient);
        
        saveData();
        renderDashboard();
        renderPatientTable();
        renderReminders();
        bootstrap.Modal.getInstance(document.getElementById('patientModal')).hide();
      });

      // Reception form submit
      const receptionForm = document.getElementById('receptionForm');
      if (receptionForm) {
        receptionForm.addEventListener('submit', async event => {
          event.preventDefault();
          const p = {
            patientId: document.getElementById('r_patientId').value.trim(),
            fullName: document.getElementById('r_fullName').value.trim(),
            birthDate: document.getElementById('r_birthDate').value,
            gender: document.getElementById('r_gender').value,
            healthInsurance: document.getElementById('r_healthInsurance').value.trim(),
            cccd: document.getElementById('r_cccd').value.trim(),
            phone: document.getElementById('r_phone').value.trim(),
            address: document.getElementById('r_address').value.trim(),
            note: document.getElementById('r_note').value.trim(),
            lastExamDate: document.getElementById('r_lastExamDate').value,
            nextAppointment: document.getElementById('r_nextAppointment').value,
            clinical: {}, hematology: {}, biochemistry: {}, immunology: {}, coagulation: {}, urine: {}, conclusion: {}, history: [], reminders: []
          };
          if (!p.patientId) { alert('Vui lòng nhập Mã bệnh nhân'); return; }
          if (p.nextAppointment) {
            const status = getStatus(p);
            if (status.label === 'Hôm nay' || status.label === 'Sắp tới hạn') {
              p.examinationState = 'in-progress';
            }
          }
          const idx = patients.findIndex(x => x.patientId === p.patientId);
          if (idx !== -1) {
            const existing = patients[idx];
            p.clinical = existing.clinical || {};
            p.hematology = existing.hematology || {};
            p.biochemistry = existing.biochemistry || {};
            p.immunology = existing.immunology || {};
            p.conclusion = existing.conclusion || {};
            p.history = existing.history || [];
            p.reminders = existing.reminders || [];
            patients[idx] = { ...existing, ...p };
          } else {
            patients.push(p);
          }
          if (p.nextAppointment) {
            const has = p.reminders.some(r => r.revisitDate === p.nextAppointment);
            if (!has) p.reminders.push({ examDate: p.lastExamDate || new Date().toISOString().split('T')[0], revisitDate: p.nextAppointment, content: 'Tái khám theo lịch hẹn', note: '' });
          }
          ensureReminderForNextAppointment(p);
          syncNextAppointmentWithReminders(p);
          await saveData();
          renderDashboard();
          renderPatientTable();
          renderReminders();
          receptionForm.reset();
          pendingReceptionPatientId = p.patientId;
          showReceptionTestSelection(p.patientId);
        });
        document.getElementById('receptionResetBtn')?.addEventListener('click', () => receptionForm.reset());
      }

      const saveReceptionTestSelectionButton = document.getElementById('saveReceptionTestSelection');
      if (saveReceptionTestSelectionButton) {
        saveReceptionTestSelectionButton.addEventListener('click', saveReceptionTestSelection);
      }
      // Reception mode tab buttons
      const receptionModeNewBtn = document.getElementById('receptionModeNewBtn');
      const receptionModeRevisitBtn = document.getElementById('receptionModeRevisitBtn');
      if (receptionModeNewBtn) receptionModeNewBtn.addEventListener('click', () => toggleReceptionMode('new'));
      if (receptionModeRevisitBtn) receptionModeRevisitBtn.addEventListener('click', () => toggleReceptionMode('revisit'));

      const rLoadRevisitBtn = document.getElementById('r_loadRevisitBtn');
      const rConfirmRevisitBtn = document.getElementById('r_confirmRevisitBtn');
      if (rLoadRevisitBtn) rLoadRevisitBtn.addEventListener('click', () => {
        const pid = (document.getElementById('r_revisitPatientId') || { value: '' }).value.trim();
        if (!pid) return alert('Nhập mã BN để tìm.');
        const p = patients.find(x => x.patientId === pid);
        if (!p) return alert('Không tìm thấy bệnh nhân.');
        // show brief info
        alert(`Tìm thấy: ${p.patientId} - ${p.fullName}`);
      });
      if (rConfirmRevisitBtn) rConfirmRevisitBtn.addEventListener('click', () => confirmRevisit());
      const cancelReceptionTestSelectionButton = document.getElementById('cancelReceptionTestSelection');
      if (cancelReceptionTestSelectionButton) {
        cancelReceptionTestSelectionButton.addEventListener('click', hideReceptionTestSelection);
      }
      const editResultsTestsBtn = document.getElementById('editResultsTestsBtn');
      if (editResultsTestsBtn) {
        editResultsTestsBtn.addEventListener('click', () => {
          if (!currentResultsPatientId) return alert('Chọn một bệnh nhân trước.');
          pendingReceptionPatientId = currentResultsPatientId;
          showReceptionTestSelection(currentResultsPatientId);
        });
      }
      ['receptionTestBiochemistryAll','receptionTestImmunologyAll','receptionTestHematologyAll','receptionTestCoagulationAll','receptionTestUrineAll'].forEach(id => {
        const cb = document.getElementById(id);
        if (cb) {
          cb.addEventListener('change', () => {
            const category = id.replace('receptionTest','').replace('All','').toLowerCase();
            updateReceptionIndividualCheckboxes(category, cb.checked);
          });
        }
      });
      receptionTestDefinitions.biochemistry.forEach(key => {
        const cb = getReceptionTestCheckbox('biochemistry', key);
        if (cb) cb.addEventListener('change', () => updateReceptionGroupCheckbox('biochemistry'));
      });
      receptionTestDefinitions.immunology.forEach(key => {
        const cb = getReceptionTestCheckbox('immunology', key);
        if (cb) cb.addEventListener('change', () => updateReceptionGroupCheckbox('immunology'));
      });
      receptionTestDefinitions.hematology.forEach(key => {
        const cb = getReceptionTestCheckbox('hematology', key);
        if (cb) cb.addEventListener('change', () => updateReceptionGroupCheckbox('hematology'));
      });
      receptionTestDefinitions.coagulation.forEach(key => {
        const cb = getReceptionTestCheckbox('coagulation', key);
        if (cb) cb.addEventListener('change', () => updateReceptionGroupCheckbox('coagulation'));
      });
      receptionTestDefinitions.urine.forEach(key => {
        const cb = getReceptionTestCheckbox('urine', key);
        if (cb) cb.addEventListener('change', () => updateReceptionGroupCheckbox('urine'));
      });

      // Import inputs
      const importPatientsInput = document.getElementById('importPatientsInput');
      if (importPatientsInput) importPatientsInput.addEventListener('change', e => handleReceptionImportPatients(e.target.files[0]));
      const importResultsInput = document.getElementById('importResultsInput');
      if (importResultsInput) importResultsInput.addEventListener('change', e => handleReceptionImportResults(e.target.files[0]));

      const importPatientsButton = document.getElementById('importPatientsButton');
      if (importPatientsButton && importPatientsInput) {
        importPatientsButton.addEventListener('click', () => importPatientsInput.click());
      }
      const importResultsButton = document.getElementById('importResultsButton');
      if (importResultsButton && importResultsInput) {
        importResultsButton.addEventListener('click', () => importResultsInput.click());
      }

      ['searchName', 'searchPhone', 'searchCccd', 'searchCode'].forEach(id => {
        document.getElementById(id).addEventListener('input', renderPatientTable);
      });
      ['resultsSearchName', 'resultsSearchCode', 'resultsSearchDate', 'resultsSearchStatus'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', renderResultsPatientTable);
      });
      // default date to today
      const resultsDateEl = document.getElementById('resultsSearchDate');
      if (resultsDateEl && !resultsDateEl.value) {
        resultsDateEl.value = new Date().toISOString().slice(0,10);
      }
      const addResultsRowButton = document.getElementById('addResultRowButton');
      if (addResultsRowButton) addResultsRowButton.addEventListener('click', () => addResultsRow(getActiveResultsCategory()));
      const saveResultsModalButton = document.getElementById('saveResultsModalButton');
      if (saveResultsModalButton) {
        saveResultsModalButton.disabled = true;
        saveResultsModalButton.addEventListener('click', saveResultsModal);
      }
      const approveResultsButton = document.getElementById('approveResultsButton');
      if (approveResultsButton) {
        approveResultsButton.disabled = true;
        approveResultsButton.addEventListener('click', () => {
          if (!currentResultsPatientId) return alert('Chọn một bệnh nhân trước.');
          const patient = patients.find(p => p.patientId === currentResultsPatientId);
          if (!patient) return;
          if (!approvePatientResults(patient)) return;
          refreshResultsPatientInfo(patient);
          updateResultsStatusLabel(patient);
          alert('Đã duyệt kết quả cho bệnh nhân.');
        });
      }
      const approveResultsInlineButton = document.getElementById('approveResultsInlineButton');
      if (approveResultsInlineButton) {
        approveResultsInlineButton.disabled = true;
        approveResultsInlineButton.addEventListener('click', () => {
          if (!currentResultsPatientId) return alert('Chọn một bệnh nhân trước.');
          const patient = patients.find(p => p.patientId === currentResultsPatientId);
          if (!patient) return;
          if (!approvePatientResults(patient)) return;
          refreshResultsPatientInfo(patient);
          updateResultsStatusLabel(patient);
          renderResultsPatientList();
          alert('Đã duyệt kết quả cho bệnh nhân.');
        });
      }

      document.querySelectorAll('[data-sort]').forEach(header => {
        header.addEventListener('click', () => {
          const field = header.dataset.sort;
          if (currentSort.field === field) {
            currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
          } else {
            currentSort.field = field;
            currentSort.direction = 'asc';
          }
          renderPatientTable();
        });
      });

      const selectAllCheckbox = document.getElementById('selectAllCheckbox');
      if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', () => toggleSelectAll(selectAllCheckbox.checked));
      }

      const quickRevisitButton = document.getElementById('quickRevisitButton');
      if (quickRevisitButton) quickRevisitButton.addEventListener('click', openQuickRevisitModal);

      const quickRevisitConfirm = document.getElementById('quickRevisitConfirm');
      if (quickRevisitConfirm) quickRevisitConfirm.addEventListener('click', applyQuickRevisit);

      ['quickRevisitModeMonths', 'quickRevisitModeDate'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.addEventListener('change', toggleQuickRevisitInputs);
      });



      document.getElementById('saveConclusionButton').addEventListener('click', () => {
        if (!currentPatientId) return;
        const patient = patients.find(item => item.patientId === currentPatientId);
        if (!patient) return;
        const conclusionTime = document.getElementById('conclusionTime').value;
        const diagnosis = document.getElementById('conclusionDiagnosis').value.trim();
        const treatment = document.getElementById('conclusionTreatment').value.trim();
        const note = document.getElementById('conclusionNote').value.trim();
        const conclusionDate = patient.lastExamDate || new Date().toISOString().split('T')[0];
        
        // Save to patient's conclusion
        patient.conclusion = { diagnosis, treatment, note, time: conclusionTime };
        if (!patient.lastExamDate) {
          patient.lastExamDate = conclusionDate;
        }
        
        // Save to history
        patient.history = patient.history || [];
        const historyEntry = {
          date: conclusionDate,
          diagnosis: diagnosis || '-',
          result: treatment || '-',
          note: note || '-'
        };
        patient.history.push(historyEntry);
        
        saveData();
        renderHistory(patient.history);
        alert('Lưu kết luận khám thành công.');
      });

      const saveReminderButton = document.getElementById('saveReminderButton');
      if (saveReminderButton) {
        saveReminderButton.addEventListener('click', () => {
          if (!currentPatientId) return;
          const patient = patients.find(item => item.patientId === currentPatientId);
          if (!patient) return;
          const reminder = {
            examDate: patient.lastExamDate || new Date().toISOString().split('T')[0],
            revisitDate: document.getElementById('reminderRevisitDate').value,
            content: document.getElementById('reminderContent').value.trim(),
            note: document.getElementById('reminderNote').value.trim()
          };
          if (!reminder.revisitDate) {
            alert('Vui lòng nhập ngày tái khám.');
            return;
          }
          // Archive previous reminders into history when adding a new appointment
          if (patient.reminders && patient.reminders.length) {
            archivePatientReminders(patient);
          }
          // Save current exam and conclusion to history
          patient.history = patient.history || [];
          const historyEntry = {
            date: reminder.examDate,
            diagnosis: document.getElementById('conclusionDiagnosis').value.trim() || reminder.content,
            result: document.getElementById('conclusionTreatment').value.trim() || 'Theo dõi',
            note: document.getElementById('conclusionNote').value.trim() || reminder.note || ''
          };
          patient.history.push(historyEntry);
          patient.reminders = [reminder];
          patient.nextAppointment = reminder.revisitDate;
          saveData();
          renderDashboard();
          renderPatientTable();
          renderReminders();
          renderPatientReminders(patient);
          renderHistory(patient.history);
          resetReminderForm();
        });
      }

      document.getElementById('themeToggle').addEventListener('click', toggleTheme);
      const addPatientButton = document.getElementById('addPatientBtn') || document.getElementById('topAddBtn');
      if (addPatientButton) {
        addPatientButton.addEventListener('click', () => {
          currentPatientId = null;
          resetPatientForm();
        });
      }
      document.getElementById('savePatientDetailButton').addEventListener('click', savePatientDetails);
      document.getElementById('backToPatientBtn').addEventListener('click', backToPatient);
      const markDoneBtn = document.getElementById('markReminderDoneButton');
      if (markDoneBtn) {
        markDoneBtn.addEventListener('click', () => {
          if (!currentPatientId) return;
          const patientIndex = patients.findIndex(item => item.patientId === currentPatientId);
          if (patientIndex === -1) return;
          markReminderDone(patientIndex, 0);
        });
      }
      // dashboard card clicks -> open modal with list
      ['todayAppointmentsCard','weekAppointmentsCard','monthAppointmentsCard','contactedCard','doneCard'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => openStatsModal(el.getAttribute('data-cat')));
      });
      const markContactBtn = document.getElementById('markReminderContactButton');
      if (markContactBtn) {
        markContactBtn.addEventListener('click', () => {
          if (!currentPatientId) return;
          const patientIndex = patients.findIndex(item => item.patientId === currentPatientId);
          if (patientIndex === -1) return;
          markReminderContact(patientIndex, 0);
        });
      }
      const deleteScheduleBtn = document.getElementById('deleteReminderScheduleButton');
      if (deleteScheduleBtn) {
        deleteScheduleBtn.addEventListener('click', () => {
          if (!currentPatientId) return;
          const patientIndex = patients.findIndex(item => item.patientId === currentPatientId);
          if (patientIndex === -1) return;
          deleteReminder(patientIndex, 0);
        });
      }

      const importExcelButton = document.getElementById('importExcelButton');
      if (importExcelButton) {
        importExcelButton.addEventListener('click', () => {
          const fileInput = document.getElementById('importExcelInput');
          if (!fileInput || !fileInput.files || !fileInput.files[0]) {
            alert('Vui lòng chọn file Excel để nhập.');
            return;
          }
          importPatientsFromExcel(fileInput.files[0]);
        });
      }

      const exportExcelButton = document.getElementById('exportExcelButton');
      if (exportExcelButton) {
        exportExcelButton.addEventListener('click', exportPatientsToExcel);
      }

      const logoutButton = document.getElementById('logoutButton');
      if (logoutButton) {
        logoutButton.addEventListener('click', logout);
      }

      const loginButton = document.getElementById('loginButton');
      if (loginButton) {
        loginButton.addEventListener('click', async () => {
          await handleLogin();
        });
      }
      const loginPassword = document.getElementById('loginPassword');
      if (loginPassword) {
        loginPassword.addEventListener('keypress', async event => {
          if (event.key === 'Enter') {
            event.preventDefault();
            await handleLogin();
          }
        });
      }

      const totalPatientsCard = document.getElementById('totalPatientsCard');
      if (totalPatientsCard) {
        totalPatientsCard.addEventListener('click', () => showPatientSectionWithFilter('all'));
      }
      const upcomingAppointmentsCard = document.getElementById('upcomingAppointmentsCard');
      if (upcomingAppointmentsCard) {
        upcomingAppointmentsCard.addEventListener('click', () => showPatientSectionWithFilter('upcoming'));
      }
      const currentMonthRevisitCard = document.getElementById('currentMonthRevisitCard');
      if (currentMonthRevisitCard) {
        currentMonthRevisitCard.addEventListener('click', showCurrentMonthReminders);
      }
      const contactedCard = document.getElementById('contactedCard');
      if (contactedCard) {
        contactedCard.addEventListener('click', () => {
          showSection('reminderSection');
          reminderStatusFilter = 'contacted';
          renderReminders();
        });
      }
      const overdueAppointmentsCard = document.getElementById('overdueAppointmentsCard');
      if (overdueAppointmentsCard) {
        overdueAppointmentsCard.addEventListener('click', () => openStatsModal('overdue'));
      }
      const doneCard = document.getElementById('doneCard');
      if (doneCard) {
        doneCard.addEventListener('click', () => {
          showSection('reminderSection');
          reminderStatusFilter = 'done';
          renderReminders();
        });
      }

      const reminderSearchInput = document.getElementById('reminderSearchInput');
      if (reminderSearchInput) {
        reminderSearchInput.addEventListener('input', event => {
          reminderSearchFilter = event.target.value;
          renderReminders();
        });
      }

      const reminderStatusSelect = document.getElementById('reminderStatusSelect');
      if (reminderStatusSelect) {
        reminderStatusSelect.addEventListener('change', event => {
          reminderStatusFilter = event.target.value;
          renderReminders();
        });
      }

      const clearReminderFilters = document.getElementById('clearReminderFilters');
      if (clearReminderFilters) {
        clearReminderFilters.addEventListener('click', () => {
          reminderSearchFilter = '';
          reminderStatusFilter = 'all';
          document.getElementById('reminderSearchInput').value = '';
          document.getElementById('reminderStatusSelect').value = 'all';
          renderReminders();
        });
      }

      // Add navigation listeners
      document.querySelectorAll('.sidebar-nav .nav-link').forEach(link => {
        link.addEventListener('click', (e) => {
          e.preventDefault();
          const sectionId = link.getAttribute('href').substring(1);
          reminderTableFilter = null;
          patientTableFilter = null;
          showSection(sectionId);
          if (sectionId === 'reminderSection') {
            renderReminders();
          }
          if (sectionId === 'patientSection') {
            renderPatientTable();
          }
        });
      });
    }

    function updateActiveNav() {
      const links = Array.from(document.querySelectorAll('.sidebar-nav .nav-link'));
      const sections = links.map(link => document.querySelector(link.getAttribute('href'))).filter(Boolean);
      const scrollPos = window.scrollY + 120;
      let activeIndex = 0;
      sections.forEach((section, index) => {
        if (section.offsetTop <= scrollPos) activeIndex = index;
      });
      links.forEach((link, index) => link.classList.toggle('active', index === activeIndex));
    }

    function showSection(sectionId) {
      // Hide all sections
      document.getElementById('dashboardSection').style.display = 'none';
      document.getElementById('receptionSection').style.display = 'none';
      document.getElementById('patientSection').style.display = 'none';
      document.getElementById('reminderSection').style.display = 'none';
      document.getElementById('resultsSection').style.display = 'none';
      document.getElementById('detailSection').style.display = 'none';
      document.getElementById('settingsSection').style.display = 'none';
      
      // Show the requested section
      const targetSection = document.getElementById(sectionId);
      if (targetSection) {
        targetSection.style.display = 'block';
      }
      
      if (sectionId === 'resultsSection') {
        renderResultsPatientTable();
      }
      
      // Update active nav link
      document.querySelectorAll('.sidebar-nav .nav-link').forEach(link => {
        const href = link.getAttribute('href');
        link.classList.toggle('active', href === `#${sectionId}`);
      });
      
      // Scroll to top
      window.scrollTo(0, 0);
    }

    function updateThemeButton(isDark) {
      const themeToggle = document.getElementById('themeToggle');
      if (!themeToggle) return;
      themeToggle.innerHTML = isDark ? '<i class="fa-solid fa-sun"></i> Light Mode' : '<i class="fa-solid fa-moon"></i> Dark Mode';
      themeToggle.classList.toggle('btn-outline-secondary', !isDark);
      themeToggle.classList.toggle('btn-outline-light', isDark);
    }

    function openCalendarSection() {
      reminderStatusFilter = 'all';
      reminderSearchFilter = '';
      const searchInput = document.getElementById('reminderSearchInput');
      const statusSelect = document.getElementById('reminderStatusSelect');
      if (searchInput) searchInput.value = '';
      if (statusSelect) statusSelect.value = 'all';
      showSection('reminderSection');
      renderReminders();
    }

    function setTheme(theme) {
      const isDark = theme === 'dark';
      document.body.classList.toggle('dark-mode', isDark);
      localStorage.setItem(themeKey, theme);
      updateThemeButton(isDark);
    }

    function toggleTheme() {
      const currentTheme = document.body.classList.contains('dark-mode') ? 'dark' : 'light';
      setTheme(currentTheme === 'dark' ? 'light' : 'dark');
    }

    function loadTheme() {
      const saved = localStorage.getItem(themeKey) || 'light';
      setTheme(saved);
    }

    async function clearStorage() {
      if (!confirm('Xóa toàn bộ dữ liệu bệnh nhân và đặt lại dữ liệu mẫu?')) return;
      localStorage.removeItem(STORAGE_KEY);
      await loadData();
      renderDashboard();
      renderPatientTable();
      renderReminders();
    }

    document.addEventListener('DOMContentLoaded', async () => {
      loadTheme();
      attachEvents();
      renderCurrentTime();
      setInterval(renderCurrentTime, 1000);
      
      // Redraw chart on resize
      let resizeTimeout;
      window.addEventListener('resize', () => {
        clearTimeout(resizeTimeout);
        resizeTimeout = setTimeout(() => {
          if (appointmentChartInstance) {
            appointmentChartInstance.resize();
          }
        }, 250);
      });
      if (window.firebaseAuth && window.firebaseAuthImports) {
        const { onAuthStateChanged } = window.firebaseAuthImports;
        onAuthStateChanged(window.firebaseAuth, async user => {
          if (user) {
            await loadData();
            renderDashboard();
            renderPatientTable();
            renderReminders();
            showApp();
          } else {
            showLoginScreen();
          }
        });
      } else {
        showLoginScreen();
      }
      renderDashboard();
      renderPatientTable();
      renderReminders();
      // Always show dashboard after reload, do not restore previous section automatically
      window.location.hash = '';
      showSection('dashboardSection');
      const patientModal = document.getElementById('patientModal');
      patientModal.addEventListener('show.bs.modal', event => {
        if (!currentPatientId) resetPatientForm();
      });
      patientModal.addEventListener('hidden.bs.modal', () => {
        currentPatientId = null;
        resetPatientForm();
      });
      const resultsModal = document.getElementById('resultsDetailModal');
      if (resultsModal) {
        resultsModal.addEventListener('hidden.bs.modal', () => {
          currentResultsPatientId = null;
          const approveResultsButton = document.getElementById('approveResultsButton');
          const approveResultsDetailButton = document.getElementById('approveResultsDetailButton');
          if (approveResultsButton) approveResultsButton.disabled = true;
          if (approveResultsDetailButton) approveResultsDetailButton.disabled = true;
        });
      }
    });

    
