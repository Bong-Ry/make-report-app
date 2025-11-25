// public/script.js (修正版・全文)

// --- グローバル変数 ---
let selectedPeriod = {}; 
let currentClinicForModal = '';
let currentAnalysisTarget = 'L'; 
let currentDetailedAnalysisType = 'L'; 
let isEditingDetailedAnalysis = false; 
let currentCentralSheetId = null; 
let currentPeriodText = ""; 
let currentAiCompletionStatus = {};
let overallDataCache = null;
let clinicDataCache = null;
// WCグラフとAI分析のキャッシュ（PDF出力用）
let wcAnalysisCache = {}; // { 'L': { html, analysisResults }, 'I': {...}, ... }
let aiAnalysisCache = {}; // { 'L': { analysis: '...', suggestions: '...', overall: '...' }, ... } 

/** @type {string | null} 'nps' | 'feedback_i' | 'feedback_j' | 'feedback_m' */
let currentCommentType = null;
/** @type {string | null} 例: "クリニックA_NPS10" */
let currentCommentSheetName = null;
/** @type {string[][] | null} 例: [ ['A列c1', 'A列c2'], ['B列c1'] ] */
let currentCommentData = null; 
/** @type {number} 0=A列, 1=B列 */
let currentCommentPageIndex = 0; 
// --- ▲▲▲ ---

// --- Google Charts ロード (変更なし) ---
const googleChartsLoaded = new Promise(resolve => {
  google.charts.load('current', { packages: ['corechart', 'bar'] });
  google.charts.setOnLoadCallback(resolve);
});

// --- ▼▼▼ [修正] イベントリスナー設定 (新UI対応) ▼▼▼ ---
function setupEventListeners() {
  // Screen 1/2
  document.getElementById('next-to-clinics').addEventListener('click', handleNextToClinics); 
  document.getElementById('issue-btn').addEventListener('click', handleIssueReport); 
  document.getElementById('issued-list-container').addEventListener('click', handleIssuedListClickDelegator);
  document.getElementById('back-to-period').addEventListener('click', () => {
    currentCentralSheetId = null; 
    currentPeriodText = "";
    currentAiCompletionStatus = {}; 
    showScreen('screen1');
  });
  
  // Screen 3 (Nav)
  document.getElementById('report-nav').addEventListener('click', handleReportNavClick);
  document.getElementById('back-to-clinics').addEventListener('click', () => showScreen('screen2'));
  
  // Screen 3 (Header - コメントUI)
  document.getElementById('slide-header').addEventListener('click', (e) => {
    // (サブナビ NPS 10, 9...)
    const subNavBtn = e.target.closest('.comment-sub-nav-btn');
    if (subNavBtn) {
      e.preventDefault();
      const key = subNavBtn.dataset.key;
      document.querySelectorAll('.comment-sub-nav-btn').forEach(b => b.classList.remove('btn-active'));
      subNavBtn.classList.add('btn-active');
      fetchAndRenderCommentPage(key);
      return;
    }
      
    // (ヘッダー内 ページネーション)
    const prevBtn = e.target.closest('#comment-prev');
    if (prevBtn) {
      e.preventDefault();
      if (currentCommentPageIndex > 0) renderCommentPage(currentCommentPageIndex - 1);
      return;
    }
      
    const nextBtn = e.target.closest('#comment-next');
    if (nextBtn) {
      e.preventDefault();
      if (currentCommentData && currentCommentPageIndex < currentCommentData.length - 1) {
        renderCommentPage(currentCommentPageIndex + 1);
      }
      return;
    }

    // (編集ボタン)
    const editBtn = e.target.closest('#comment-edit-btn');
    if (editBtn) {
      e.preventDefault();
      enterCommentEditMode();
      return;
    }

    // (保存ボタン)
    const saveBtn = e.target.closest('#comment-save-btn');
    if (saveBtn) {
      e.preventDefault();
      saveCommentEdit();
      return;
    }
  });

  // Screen 5 (Nav)
  document.getElementById('report-nav-screen5').addEventListener('click', handleReportNavClick);
  document.getElementById('back-to-clinics-from-detailed-analysis').addEventListener('click', () => {
    toggleEditDetailedAnalysis(false); 
    showScreen('screen2');
  });

  // Screen 5 (AI Tabs)
  document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => {
    button.addEventListener('click', handleTabClick);
  });
  
  // Screen 5 (AI Controls)
  document.getElementById('regenerate-detailed-analysis-btn').addEventListener('click', handleRegenerateDetailedAnalysis);
  document.getElementById('edit-detailed-analysis-btn').addEventListener('click', () => toggleEditDetailedAnalysis(true));
  document.getElementById('save-detailed-analysis-btn').addEventListener('click', saveDetailedAnalysisEdits);
  document.getElementById('cancel-edit-detailed-analysis-btn').addEventListener('click', () => {
    if (confirm('編集内容を破棄しますか？')) {
      toggleEditDetailedAnalysis(false);
    }
  });

  // PDF出力ボタン
  document.getElementById('pdf-export-btn').addEventListener('click', handlePdfExport);
  document.getElementById('close-print-popup').addEventListener('click', () => {
    document.getElementById('print-ready-popup').classList.add('hidden');
  });
}

// --- 画面1/2 処理 (変更なし) ---

function populateDateSelectors() {
  const now = new Date();
  const cy = now.getFullYear();
  const sy = document.getElementById('start-year');
  const ey = document.getElementById('end-year');
  
  if (!sy || !ey) {
    console.error("日付選択プルダウン（#start-year または #end-year）が見つかりません。");
    return;
  }
  
  for (let i = 0; i < 5; i++) {
    const y = cy - i;
    sy.add(new Option(`${y}年`, y));
    ey.add(new Option(`${y}年`, y));
  }

  const sm = document.getElementById('start-month');
  const em = document.getElementById('end-month');

  if (!sm || !em) {
    console.error("日付選択プルダウン（#start-month または #end-month）が見つかりません。");
    return;
  }
  
  for (let i = 1; i <= 12; i++) {
    const m = String(i).padStart(2, '0');
    sm.add(new Option(`${i}月`, m));
    em.add(new Option(`${i}月`, m));
  }
  
  // デフォルト値を設定
  const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
  
  sy.value = String(cy); // 開始年 = 現在の年
  sm.value = currentMonth; // 開始月 = 現在の月
  ey.value = String(cy); // 終了年 = 現在の年
  em.value = currentMonth; // 終了月 = 現在の月
}

async function handleNextToClinics() {
  const sy = document.getElementById('start-year').value;
  const sm = document.getElementById('start-month').value;
  const ey = document.getElementById('end-year').value;
  const em = document.getElementById('end-month').value;
  const sd = new Date(`${sy}-${sm}-01`);
  const ed = new Date(`${ey}-${em}-01`);

  if (sd > ed) {
    alert('開始年月<=終了年月で設定');
    return;
  }

  selectedPeriod = { start: `${sy}-${sm}`, end: `${ey}-${em}` };
  currentPeriodText = `${sy}-${sm}～${ey}-${em}`; 
  const displayPeriod = `${sy}年${sm}月～${ey}年${em}月`;

  showLoading(true, '集計シートを検索・準備中...');

  try {
    const response = await fetch('/api/findOrCreateSheet', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ periodText: currentPeriodText })
    });

    if (!response.ok) {
      const et = await response.text();
      throw new Error(`サーバーエラー(${response.status}): ${et}`);
    }
      
    const data = await response.json();
    currentCentralSheetId = data.centralSheetId; 
      
    document.getElementById('period-display').textContent = `集計期間：${displayPeriod}`;
    showScreen('screen2');
    loadClinics();
  } catch (err) {
    console.error('!!! Find/Create Sheet failed:', err);
    alert(`集計シートの準備に失敗\n${err.message}`);
  } finally {
    showLoading(false);
  }
}

async function loadClinics() {
  showLoading(true, 'クリニック一覧と転記状況を読込中...');
  
  const clinicListContainer = document.getElementById('clinic-list-container');
  const issuedListContainer = document.getElementById('issued-list-container');
  clinicListContainer.innerHTML = '<p class="text-sm text-gray-500 text-center py-4">読込中...</p>';
  issuedListContainer.innerHTML = '<p class="text-sm text-gray-500 text-center py-4">読込中...</p>';
  document.getElementById('issue-btn').disabled = true; 

  try {
    const [clinicListRes, transferredListRes] = await Promise.all([
      fetch('/api/getClinicList'),
      fetch('/api/getTransferredList', { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ centralSheetId: currentCentralSheetId })
      })
    ]);

    if (!clinicListRes.ok) throw new Error(`クリニック一覧の取得失敗: ${await clinicListRes.text()}`);
    if (!transferredListRes.ok) throw new Error(`転記済み一覧の取得失敗: ${await transferredListRes.text()}`);
      
    const masterClinics = await clinicListRes.json(); 
    const { sheetTitles, aiCompletionStatus } = await transferredListRes.json(); 
      
    currentAiCompletionStatus = aiCompletionStatus;
    const transferredTitlesSet = new Set(sheetTitles); 
      
    clinicListContainer.innerHTML = '';
    issuedListContainer.innerHTML = '';
      
    let untransferredCount = 0;
    let transferredCount = 0;

    if (!Array.isArray(masterClinics) || masterClinics.length === 0) {
      clinicListContainer.innerHTML = '<p class="text-sm text-gray-500 text-center py-4">対象なし</p>';
    } else {
      masterClinics.forEach(clinicName => {
        if (transferredTitlesSet.has(clinicName)) {
          const d = document.createElement('div');
          d.className = 'flex items-center justify-between p-4 border rounded-lg hover:bg-gray-50 hover:shadow-md transition-all duration-200';
          d.innerHTML = `<p class="font-bold text-base view-report-btn" data-clinic-name="${clinicName}">${clinicName}</p>`;
          issuedListContainer.appendChild(d);
          transferredCount++;
        } else {
          const d = document.createElement('div');
          d.className = 'flex items-center p-2 rounded-md hover:bg-gray-100';
          d.innerHTML = `
            <input type="checkbox" id="clinic-${clinicName}" name="clinic" value="${clinicName}" class="mr-3 h-4 w-4 rounded border-gray-300 text-gray-800 focus:ring-gray-800 cursor-pointer">
            <label for="clinic-${clinicName}" class="text-sm select-none cursor-pointer text-gray-900">${clinicName}</label>
          `;
          clinicListContainer.appendChild(d);
          untransferredCount++;
        }
      });
    }
      
    if (untransferredCount === 0) {
      clinicListContainer.innerHTML = '<p class="text-sm text-gray-500 text-center py-4">未転記のクリニックはありません</p>';
    }
    if (transferredCount === 0) {
      issuedListContainer.innerHTML = '<p class="text-sm text-gray-500 text-center py-8">転記済みレポートなし</p>';
    }

    clinicListContainer.removeEventListener('change', handleClinicCheckboxChange);
    clinicListContainer.addEventListener('change', handleClinicCheckboxChange);
    handleClinicCheckboxChange();

  } catch (err) {
    console.error('!!! Load failed:', err);
    clinicListContainer.innerHTML = `<p class="text-sm text-red-500 text-center py-4">読込失敗<br>${err.message}</p>`;
    issuedListContainer.innerHTML = `<p class="text-sm text-red-500 text-center py-4">読込失敗<br>${err.message}</p>`;
  } finally {
    showLoading(false);
  }
}

function handleClinicCheckboxChange(event) {
  const checkedCheckboxes = document.querySelectorAll('#clinic-list-container input:checked');
  const checkedCount = checkedCheckboxes.length;
  const issueBtn = document.getElementById('issue-btn');
  
  if (checkedCount > 10) {
    alert('一度に選択できる件数は10件までです。');
    if (event && event.target) {
      event.target.checked = false;
    }
  }
  
  const finalCheckedCount = document.querySelectorAll('#clinic-list-container input:checked').length;
  
  document.querySelectorAll('#clinic-list-container input[type="checkbox"]').forEach(cb => {
    if (!cb.checked) {
      cb.disabled = (finalCheckedCount >= 10);
    } else {
      cb.disabled = false; 
    }
  });
  
  issueBtn.disabled = (finalCheckedCount === 0);
}

async function handleIssueReport() {
  const sc = Array.from(document.querySelectorAll('#clinic-list-container input:checked')).map(cb => cb.value);
  if (sc.length === 0) { alert('転記対象のクリニックが選択されていません。'); return; }
  if (sc.length > 10) { alert('一度に選択できる件数は10件までです。チェックを減らしてください。'); return; }

  showLoading(true, '集計スプレッドシートへデータ転記中...\n完了後、バックグラウンドで分析タブが自動生成されます。');

  try {
    const r = await fetch('/api/getReportData', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        period: selectedPeriod,
        selectedClinics: sc,
        centralSheetId: currentCentralSheetId 
      })
    });
    if (!r.ok) { const et = await r.text(); throw new Error(`サーバーエラー(${r.status}): ${et}`); }
    await r.json(); 
    loadClinics(); 
  } catch (err) {
    console.error('!!! ETL Issue failed:', err);
    alert(`データ転記失敗\n${err.message}`);
  } finally {
    showLoading(false);
  }
}

// --- 画面3/5 (レポート表示/AI分析) 処理 ---

function handleIssuedListClickDelegator(e) {
  const viewBtn = e.target.closest('.view-report-btn');
  if (viewBtn) {
    handleIssuedListClick(viewBtn.dataset.clinicName);
  }
}

async function handleIssuedListClick(clinicName) {
  currentClinicForModal = clinicName;
  console.log(`Issued clinic clicked: ${currentClinicForModal}`);
  
  clinicDataCache = null;
  overallDataCache = null;
  
  prepareAndShowReport('cover');
}

function handleReportNavClick(e) {
  const targetButton = e.target.closest('.btn');
  if (!targetButton) return;
  const reportType = targetButton.dataset.reportType;
  const analysisType = targetButton.dataset.analysisType;
  const detailedAnalysisType = targetButton.dataset.detailedAnalysisType;
  
  // コメント系の種別保持
  const commentTypeMap = { 'nps': 'nps', 'feedback_i': 'feedback_i', 'feedback_j': 'feedback_j', 'feedback_m': 'feedback_m' };
  currentCommentType = commentTypeMap[reportType] || null;

  if (reportType) {
    showScreen('screen3');
    prepareAndShowReport(reportType);
  } else if (analysisType) {
    currentCommentType = null; // WCはコメント系ではない
    currentAnalysisTarget = analysisType;
    showScreen('screen3');
    prepareAndShowAnalysis(analysisType); 
  } else if (detailedAnalysisType) {
    currentCommentType = null; // AI分析もコメント系ではない
    currentDetailedAnalysisType = detailedAnalysisType;
    prepareAndShowDetailedAnalysis(detailedAnalysisType); 
  }
}

// 共通データ取得関数 (キャッシュ対応)
async function getReportDataForCurrentClinic(sheetName) {
  const isOverall = sheetName === "全体";
  const cache = isOverall ? overallDataCache : clinicDataCache;
  if (cache) {
    console.log(`[Cache] Using cached data for ${sheetName}`);
    return cache;
  }
  
  console.log(`[Cache] Fetching fresh data for ${sheetName}`);
  
  const response = await fetch('/api/getChartData', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ centralSheetId: currentCentralSheetId, sheetName: sheetName })
  });
  if (!response.ok) throw new Error(`データ取得失敗(${sheetName}): ${await response.text()}`);
  const data = await response.json();
  
  if (isOverall) overallDataCache = data;
  else clinicDataCache = data;
  
  return data;
}

// --- ▼▼▼ レポート表示メイン (Screen 3) ▼▼▼ ---
async function prepareAndShowReport(reportType) {
  console.log(`Prepare report: ${reportType}`);
  showLoading(true,'レポートデータ集計中...');

  showScreen('screen3');
  updateNavActiveState(reportType, null, null);

  // グラフ位置固定クラスの制御
  const reportBody = document.querySelector('.report-body');
  const fixedLayoutTypes = [
    'age', 'children', 'income', 'recommendation',
    'satisfaction_b', 'satisfaction_c', 'satisfaction_d',
    'satisfaction_e', 'satisfaction_f', 'satisfaction_g', 'satisfaction_h'
  ];

  if (fixedLayoutTypes.includes(reportType)) {
    reportBody.classList.add('fixed-chart-layout');
  } else {
    reportBody.classList.remove('fixed-chart-layout');
  }

  // UI初期化
  document.getElementById('report-title').textContent = '';
  document.getElementById('report-subtitle').textContent = '';
  document.getElementById('report-title').style.textAlign = 'left';
  document.getElementById('report-separator').style.display = 'block'; 
  
  const subNav = document.getElementById('comment-sub-nav');
  const controls = document.getElementById('comment-controls');
  if (subNav) subNav.innerHTML = '';
  if (controls) controls.innerHTML = '';
  
  const slideBody = document.getElementById('slide-body');
  slideBody.innerHTML = '';
  slideBody.style.overflowY = 'hidden'; // ②スクロール禁止
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  showCopyrightFooter(reportType !== 'cover');

  // --- 1. コメント系レポート ---
  if (currentCommentType) {
    try {
      const initialKey = currentCommentType === 'nps'
        ? 'L_10'
        : (currentCommentType === 'feedback_i' ? 'I' : (currentCommentType === 'feedback_j' ? 'J' : 'M'));
      showCommentSubNav(currentCommentType);
      await fetchAndRenderCommentPage(initialKey);
    } catch (e) {
      console.error('Comment data fetch error:', e);
      document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">コメントデータ取得失敗<br>(${e.message})</p>`;
    } finally {
      showLoading(false);
    }
    return;
  }

  // --- 2. 例外構成 ---
  if (reportType === 'cover' || reportType === 'toc' || reportType === 'summary') {
    await prepareAndShowIntroPages(reportType);
    showLoading(false);
    return;
  }
  
  if (reportType === 'municipality') {
    await prepareAndShowMunicipalityReport(); 
    return;
  }
  
  if (reportType === 'recommendation') {
    await prepareAndShowRecommendationReport(); 
    return;
  }

  // --- 3. グラフ ---
  let isChart = false;
  let clinicData, overallData;

  try {
    clinicData = await getReportDataForCurrentClinic(currentClinicForModal);
    overallData = await getReportDataForCurrentClinic("全体");
  } catch (e) {
    console.error('Chart data fetch error:', e);
    document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">グラフデータ取得失敗<br>(${e.message})</p>`;
    showLoading(false);
    return;
  }
  
  if (reportType === 'nps_score') {
    prepareChartPage(
      'アンケート結果　ーNPS(ネットプロモータースコア)＝推奨度ー',
      'これから初めてお産を迎える友人知人がいた場合、\nご出産された産婦人科医院をどのくらいお勧めしたいですか。\n友人知人への推奨度を教えてください。＜推奨度＞ 10:強くお勧めする〜 0:全くお勧めしない',
      'nps_score'
    );
    isChart = true;
  } else if (reportType === 'satisfaction_b') {
    prepareChartPage('アンケート結果　ー満足度ー','ご出産された産婦人科医院への満足度について、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_b');
    isChart = true;
  } else if (reportType === 'satisfaction_c') {
    prepareChartPage('アンケート結果　ー施設の充実度・快適さー','ご出産された産婦人科医院への施設の充実度・快適さについて、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_c');
    isChart = true;
  } else if (reportType === 'satisfaction_d') {
    prepareChartPage('アンケート結果　ーアクセスの良さー','ご出産された産婦人科医院へのアクセスの良さについて、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_d');
    isChart = true;
  } else if (reportType === 'satisfaction_e') {
    prepareChartPage('アンケート結果　ー費用ー','ご出産された産婦人科医院への費用について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_e');
    isChart = true;
  } else if (reportType === 'satisfaction_f') {
    prepareChartPage('アンケート結果　ー病院の雰囲気ー','ご出産された産婦人科医院への病院の雰囲気について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_f');
    isChart = true;
  } else if (reportType === 'satisfaction_g') {
    prepareChartPage('アンケート結果　ースタッフの対応ー','ご出産された産婦人科医院へのスタッフの対応について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_g');
    isChart = true;
  } else if (reportType === 'satisfaction_h') {
    prepareChartPage('アンケート結果　ー先生の診断・説明ー','ご出産された産婦人科医院への先生の診断・説明について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満','satisfaction_h');
    isChart = true;
  } else if (reportType === 'age') {
    prepareChartPage('アンケート結果　ーご回答者さまの年代ー','ご出産された方の年代について教えてください。','age');
    isChart = true;
  } else if (reportType === 'children') {
    prepareChartPage('アンケート結果　ーご回答者さまのお子様の人数ー','ご出産された方のお子様の人数について教えてください。','children');
    isChart = true;
  } else if (reportType === 'income') {
    prepareChartPage('アンケート結果　ーご回答者さまの世帯年収ー','ご出産された方の世帯年収について教えてください。','income', true);
    isChart = true;
  }

  if (isChart) {
    // Google Chartsがロードされるまで待つ
    const waitForGoogleCharts = () => {
      return new Promise((resolve) => {
        if (typeof google !== 'undefined' && google.visualization) {
          resolve();
        } else {
          const checkInterval = setInterval(() => {
            if (typeof google !== 'undefined' && google.visualization) {
              clearInterval(checkInterval);
              resolve();
            }
          }, 100);
          // タイムアウト
          setTimeout(() => {
            clearInterval(checkInterval);
            resolve();
          }, 5000);
        }
      });
    };

    waitForGoogleCharts().then(() => {
      setTimeout(() => {
        try {
          if (reportType === 'nps_score') {
            drawNpsScoreCharts(clinicData.npsScoreData, overallData.npsScoreData);
        } else if (reportType.startsWith('satisfaction')) {
          const type = reportType.split('_')[1] + '_column';
          drawSatisfactionCharts(clinicData.satisfactionData[type].results, overallData.satisfactionData[type].results);
        } else if (reportType === 'age') {
          drawSatisfactionCharts(clinicData.ageData.results, overallData.ageData.results);
        } else if (reportType === 'children') {
          drawSatisfactionCharts(clinicData.childrenCountData.results, overallData.childrenCountData.results);
        } else if (reportType === 'income') {
          drawIncomeCharts(clinicData.incomeData, overallData.incomeData);
        }
      } catch (e) {
          console.error('Chart draw error:', e);
          document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">グラフ描画失敗<br>(${e.message})</p>`;
        } finally {
          showLoading(false);
        }
      }, 100);
    });
  } else {
    showLoading(false);
  } 
}

// (変更なし) 例外構成（表紙・目次・概要）の表示
async function prepareAndShowIntroPages(reportType) {
  document.getElementById('report-separator').style.display = 'none'; 
  document.getElementById('report-subtitle').style.textAlign = 'center'; 
  document.getElementById('slide-body').style.whiteSpace = 'pre-wrap';
  document.getElementById('slide-body').classList.remove('flex', 'items-center', 'justify-center', 'text-center');

  if (reportType === 'cover') {
    // 表紙：タイトル・サブタイトル・点線を非表示
    document.getElementById('report-title').textContent = '';
    document.getElementById('report-subtitle').textContent = '';
    document.getElementById('report-title').style.textAlign = 'center';

    // 背景画像URL
    const bgImageUrl = convertGoogleDriveUrl('https://drive.google.com/file/d/1-a8Hw5h15t6wvAafU2zoxg5uLCOmjIt-/view?usp=drive_link');

    // 表紙の内容
    const [sy, sm] = selectedPeriod.start.split('-').map(Number);
    const [ey, em] = selectedPeriod.end.split('-').map(Number);
    const startYearMonth = `${sy}年${sm}月`;
    const endYearMonth = `${ey}年${em}月`;

    document.getElementById('slide-body').innerHTML = `
      <div class="w-full h-full cover-background" style="position: relative; display: flex; align-items: center; justify-content: center;">
        <img src="${bgImageUrl}" alt="表紙背景" style="width: 95%; height: 95%; object-fit: cover; z-index: 0;">
        <div class="flex items-start justify-start h-full p-12" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: 1; padding-top: 120px;">
          <div class="text-left">
            <h2 class="text-[31px] font-bold mb-6">${currentClinicForModal}様<br>満足度調査結果報告書</h2>
            <p class="text-base mt-4">調査期間：${startYearMonth}〜${endYearMonth}</p>
          </div>
        </div>
      </div>
    `;
  } else if (reportType === 'toc') {
    document.getElementById('report-title').textContent = '目次';
    document.getElementById('report-title').style.textAlign = 'center';
    document.getElementById('report-title').style.marginBottom = '4px';
    document.getElementById('report-subtitle').textContent = '';

    document.getElementById('slide-body').innerHTML = `
      <div class="flex justify-center items-start h-full pt-8">

        <ul class="text-lg font-normal space-y-4 text-left">
          <li>１．アンケート概要</li>
          <li>２．アンケート結果</li>

          <ul class="pl-8 space-y-2 font-normal">
            <li>―１　顧客属性</li>
            <li>―２　病院への満足度（施設・ハード面）</li>
            <li>―３　病院への満足度（質・スタッフ面）</li>
            <li>―４　NPS推奨度・理由</li>
          </ul>

          <li>３．アンケート結果からの考察</li>
        </ul>
      </div>
    `;
  } else if (reportType === 'summary') {
    let overallCount = 0, clinicCount = 0, clinicListCount = 0;
    try {
      const rowCounts = await fetch('/api/getSheetRowCounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          centralSheetId: currentCentralSheetId,
          clinicName: currentClinicForModal
        })
      }).then(r => r.json());

      overallCount = rowCounts.overallCount || 0;
      clinicListCount = rowCounts.managementCount || 0;
      clinicCount = rowCounts.clinicCount || 0;
    } catch (e) { console.warn("Error fetching data for summary:", e); }

    const [sy, sm] = selectedPeriod.start.split('-').map(Number);
    const [ey, em] = selectedPeriod.end.split('-').map(Number);
    const startDay = new Date(sy, sm - 1, 1).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
    const endDay = new Date(ey, em, 0).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
      
    document.getElementById('report-title').textContent = 'アンケート概要';
    document.getElementById('report-title').style.textAlign = 'left';
    document.getElementById('report-subtitle').textContent = '';
    document.getElementById('slide-body').innerHTML = `
      <div class="flex justify-center h-full items-start" style="padding-top: 0;">
        <ul class="text-lg font-normal space-y-4 text-left">
          <li><span class="font-bold text-gray-800 w-32 inline-block">調査目的</span>｜貴院に対する満足度調査</li>
          <li><span class="font-bold text-gray-800 w-32 inline-block">調査方法</span>｜スマホ利用してのアンケートフォームによるインターネット調査</li>
          <li><span class="font-bold text-gray-800 w-32 inline-block">調査対象</span>｜貴院で出産された方（退院後～１か月健診までの期間）</li>
          <li><span class="font-bold text-gray-800 w-32 inline-block">調査期間</span>｜${startDay}〜${endDay}</li>
          <li><span class="font-bold text-gray-800 w-32 inline-block">回答件数</span>｜全体：${overallCount}件（${clinicListCount}病院）　貴院：${clinicCount}件</li>
        </ul>
      </div>
    `;
  }
}

// ▼▼▼ グラフ描画用シェル設定 ▼▼▼
function prepareChartPage(title, subtitle, type, isBar = false) { 
  document.getElementById('report-title').textContent = title;
  document.getElementById('report-subtitle').textContent = subtitle;
  document.getElementById('report-subtitle').style.textAlign = 'center';
  document.getElementById('report-separator').style.display = 'block';

  let htmlContent = '';
  const cid = isBar ? 'bar-chart' : 'pie-chart';
  const chartHeightClass = 'h-[320px]';

  if (type === 'nps_score') {
    const npsGraphImageUrl = convertGoogleDriveUrl('https://drive.google.com/file/d/1jAnKR5iG4BY2xTfcqnedvfiPpm-j-ZbX/view?usp=drive_link');
    const npsBoxImageUrl = convertGoogleDriveUrl('https://drive.google.com/file/d/1QKO6nlee3DQQmoYUEzKyBZcqTmke_HVe/view?usp=drive_link');

    htmlContent = `
      <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-full">
        <div class="flex flex-col h-full">
          <h3 id="clinic-chart-header" class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
          <div id="clinic-bar-chart" class="w-full ${chartHeightClass} border border-gray-200 flex items-center justify-center"></div>
          <div class="w-full flex-1 flex flex-col justify-start items-center mt-4 overflow-hidden">
            <img src="${npsGraphImageUrl}" alt="NPS説明画像" class="w-full max-h-full object-contain" />
          </div>
        </div>
        <div class="flex flex-col h-full">
          <div class="w-full flex justify-center mb-10 mt-30">
            <img src="${npsBoxImageUrl}" alt="NPSアイコン" class="h-32 object-contain" />
          </div>
          <div id="nps-summary-area" class="flex flex-col justify-start items-center space-y-6 h-full pt-16">
            <p class="text-gray-500">NPSスコア計算中...</p>
          </div>
        </div>
      </div>
    `;
  } else {
    // [修正] グラフの位置を統一（上詰め＋固定パディング）
    htmlContent = `
      <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-full pt-12">
        <div class="flex flex-col items-center">
          <h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
          <div id="clinic-${cid}" class="w-full ${chartHeightClass} clinic-graph-bg-yellow"></div>
        </div>
        <div class="flex flex-col items-center">
          <h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3>
          <div id="average-${cid}" class="w-full ${chartHeightClass}"></div>
        </div>
      </div>
    `;
  }
  
  document.getElementById('slide-body').innerHTML = htmlContent;
}

// --- グラフ描画関数 ---
function drawSatisfactionCharts(clinicChartData, overallChartData) {
  // 左側（貴院）のグラフのみ背景色を設定
  const clinicOpt = {
    is3D: true,
    chartArea: { left: '5%', top: '5%', width: '90%', height: '90%', backgroundColor: '#ffff95' },
    pieSliceText: 'percentage',
    pieSliceTextStyle: { color: 'black', fontSize: 12, bold: true },
    legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 12 } },
    tooltip: { showColorCode: true, textStyle: { fontSize: 12 }, trigger: 'focus' },
    backgroundColor: '#ffff95',
    sliceVisibilityThreshold: 0 // [修正] 自動で「その他」にまとめない
  };
  // 右側（全体平均）のグラフは背景色なし
  const overallOpt = {
    is3D: true,
    chartArea: { left: '5%', top: '5%', width: '90%', height: '90%' },
    pieSliceText: 'percentage',
    pieSliceTextStyle: { color: 'black', fontSize: 12, bold: true },
    legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 12 } },
    tooltip: { showColorCode: true, textStyle: { fontSize: 12 }, trigger: 'focus' },
    sliceVisibilityThreshold: 0 // [修正] 自動で「その他」にまとめない
  };
  const cdEl = document.getElementById('clinic-pie-chart');
  if (!cdEl) throw new Error('グラフ描画エリア(clinic-pie-chart)が見つかりません。');
  if (clinicChartData && clinicChartData.length > 1 && clinicChartData.slice(1).some(row => row[1] > 0)) {
    const d = google.visualization.arrayToDataTable(clinicChartData);
    const c = new google.visualization.PieChart(cdEl);
    c.draw(d, clinicOpt);
  } else {
    cdEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
  }
  const avgEl = document.getElementById('average-pie-chart');
  if (!avgEl) throw new Error('グラフ描画エリア(average-pie-chart)が見つかりません。');
  if (overallChartData && overallChartData.length > 1 && overallChartData.slice(1).some(row => row[1] > 0)) {
    const avgD = google.visualization.arrayToDataTable(overallChartData);
    const avgC = new google.visualization.PieChart(avgEl);
    avgC.draw(avgD, overallOpt);
  } else {
    avgEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
  }
}

function drawIncomeCharts(clinicData, overallData) {
  // 左側（貴院）のグラフのみ背景色を設定
  const clinicOpt = {
    legend: { position: 'none', textStyle: { fontSize: 14 } },
    annotations: { textStyle: { fontSize: 14, color: 'black', auraColor: 'none' }, alwaysOutside: false, stem: { color: 'transparent' } },
    vAxis: { format: "#.##'%'", viewWindow: { min: 0 }, textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    hAxis: { textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    chartArea: { height: '75%', width: '90%', left: '5%', top: '5%', backgroundColor: '#ffff95' },
    backgroundColor: '#ffff95',
    colors: ['#DE5D83']
  };
  // 右側（全体平均）のグラフは背景色なし
  const overallOpt = {
    legend: { position: 'none', textStyle: { fontSize: 14 } },
    annotations: { textStyle: { fontSize: 14, color: 'black', auraColor: 'none' }, alwaysOutside: false, stem: { color: 'transparent' } },
    vAxis: { format: "#.##'%'", viewWindow: { min: 0 }, textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    hAxis: { textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    chartArea: { height: '75%', width: '90%', left: '5%', top: '5%' },
    colors: ['#DE5D83']
  };
  const ccdEl = document.getElementById('clinic-bar-chart');
  if (!ccdEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');
  if (clinicData.totalCount > 0 && clinicData.results && clinicData.results.length > 1) {
    const cd = google.visualization.arrayToDataTable(clinicData.results);
    const cc = new google.visualization.ColumnChart(ccdEl);
    cc.draw(cd, clinicOpt);
  } else {
    ccdEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
  }
  const avgEl = document.getElementById('average-bar-chart');
  if (!avgEl) throw new Error('グラフ描画エリア(average-bar-chart)が見つかりません。');
  if (overallData.totalCount > 0 && overallData.results && overallData.results.length > 1) {
    const avgD = google.visualization.arrayToDataTable(overallData.results);
    const avgC = new google.visualization.ColumnChart(avgEl);
    avgC.draw(avgD, overallOpt);
  } else {
    avgEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
  }
}

function drawNpsScoreCharts(clinicData, overallData) {
  const clinicChartEl = document.getElementById('clinic-bar-chart');
  if (!clinicChartEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');
  const clinicNpsScore = calculateNps(clinicData.counts, clinicData.totalCount);
  const overallNpsScore = calculateNps(overallData.counts, overallData.totalCount);
  const clinicChartData = [['スコア', '割合', { role: 'annotation' }]];
  if (clinicData.totalCount > 0) {
    for (let i = 0; i <= 10; i++) {
      const count = clinicData.counts[i] || 0;
      const percentage = (count / clinicData.totalCount) * 100;
      clinicChartData.push([String(i), percentage, `${Math.round(percentage)}%`]);
    }
  }
  const opt = {
    legend: { position: 'none', textStyle: { fontSize: 14 } },
    annotations: { textStyle: { fontSize: 14, color: 'black', auraColor: 'none' }, alwaysOutside: false, stem: { color: 'transparent' } },
    vAxis: { format: "#.##'%'", title: '割合(%)', viewWindow: { min: 0 }, textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    hAxis: { title: '推奨度スコア (0〜10)', textStyle: { fontSize: 14 }, titleTextStyle: { fontSize: 14 } },
    bar: { groupWidth: '80%' },
    isStacked: false,
    chartArea: { height: '75%', width: '90%', left: '5%', top: '5%' },
    backgroundColor: 'transparent',
    colors: ['#DE5D83']
  };
  if (clinicData.totalCount > 0 && clinicChartData.length > 1) {
    const clinicDataVis = google.visualization.arrayToDataTable(clinicChartData);
    const clinicChart = new google.visualization.ColumnChart(clinicChartEl);
    clinicChart.draw(clinicDataVis, opt);
  } else {
    clinicChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
  }
  const summaryArea = document.getElementById('nps-summary-area');
  if (summaryArea) {
    summaryArea.innerHTML = `
      <div class="text-left text-3xl space-y-5 p-4 w-48">
        <p>全体：<span class="font-bold text-gray-800">${overallNpsScore.toFixed(1)}</span></p>
        <p>貴院：<span class="font-bold text-red-600">${clinicNpsScore.toFixed(1)}</span></p>
      </div>
    `;
  }
  const clinicHeaderEl = document.getElementById('clinic-chart-header');
  if (clinicHeaderEl) {
    clinicHeaderEl.textContent = `貴院の結果 (全 ${clinicData.totalCount} 件)`;
  }
}

function calculateNps(counts, totalCount) {
  if (totalCount === 0) return 0;
  let promoters = 0, passives = 0, detractors = 0;
  for (let i = 0; i <= 10; i++) {
    const count = counts[i] || 0;
    if (i >= 9) promoters += count;
    else if (i >= 7) passives += count;
    else detractors += count;
  }
  return ((promoters / totalCount) - (detractors / totalCount)) * 100;
}

// --- ▼▼▼ コメントスライド構築関数群 ▼▼▼ ---

function getCommentSheetName(clinicName, type) {
  switch (type) {
    case 'L_10': return `${clinicName}_NPS10`;
    case 'L_9': return `${clinicName}_NPS9`;
    case 'L_8': return `${clinicName}_NPS8`;
    case 'L_7': return `${clinicName}_NPS7`;
    case 'L_6_under': return `${clinicName}_NPS6以下`;
    case 'I': return `${clinicName}_よかった点悪かった点`;
    case 'J': return `${clinicName}_印象スタッフ`;
    case 'M': return `${clinicName}_お産意見`;
    default: return null;
  }
}

function getExcelColumnName(colIndex) {
  let colName = '';
  let dividend = colIndex + 1;
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    colName = String.fromCharCode(65 + modulo) + colName;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return colName;
}

function showCommentSubNav(reportType) {
  const navContainer = document.getElementById('comment-sub-nav');
  if (!navContainer) {
    console.error('comment-sub-nav element not found');
    return;
  }
  navContainer.innerHTML = '';

  let title = '';
  let subTitle = '';

  if (reportType === 'nps') {
    title = 'アンケート結果　ーNPS推奨度 理由ー';
    subTitle = 'データ一覧（1列12データずつ）';
    const groups = [
      { key: 'L_10', label: '10点' },
      { key: 'L_9', label: '9点' },
      { key: 'L_8', label: '8点' },
      { key: 'L_7', label: '7点' },
      { key: 'L_6_under', label: '6点以下' },
    ];
    groups.forEach((group, index) => {
      const btn = document.createElement('button');
      btn.className = 'btn comment-sub-nav-btn';
      if (index === 0) btn.classList.add('btn-active');
      btn.dataset.key = group.key;
      btn.textContent = group.label;
      navContainer.appendChild(btn);
    });
  } else {
    const titleMap = { 'feedback_i': '良かった点や悪かった点など', 'feedback_j': '印象に残ったスタッフへのコメント', 'feedback_m': 'お産にかかわるご意見・ご感想' };
    title = `アンケート結果　ー${titleMap[reportType]}ー`;
    subTitle = '';
  }

  document.getElementById('report-title').textContent = title;
  document.getElementById('report-subtitle').textContent = subTitle;
  document.getElementById('report-subtitle').style.textAlign = 'left';
}

// NPSコメントのサブタイトルを更新する関数
function updateCommentSubtitle(commentKey, totalCount) {
  if (!currentCommentType || currentCommentType !== 'nps') return;

  const labelMap = {
    'L_10': 'NPS10',
    'L_9': 'NPS9',
    'L_8': 'NPS8',
    'L_7': 'NPS7',
    'L_6_under': 'NPS6以下'
  };

  // 象のアイコンURL設定
  const elephantIconMap = {
    'L_10': convertGoogleDriveUrl('https://drive.google.com/file/d/1W2SGYDfVR0_0NgVeibKP4wXctHyDMNvy/view?usp=drive_link'),
    'L_9': convertGoogleDriveUrl('https://drive.google.com/file/d/1W2SGYDfVR0_0NgVeibKP4wXctHyDMNvy/view?usp=drive_link'),
    'L_8': convertGoogleDriveUrl('https://drive.google.com/file/d/1S-vNIRLbS2UAcZmkGE2cb_upnyaOAZfG/view?usp=drive_link'),
    'L_7': convertGoogleDriveUrl('https://drive.google.com/file/d/1S-vNIRLbS2UAcZmkGE2cb_upnyaOAZfG/view?usp=drive_link'),
    'L_6_under': convertGoogleDriveUrl('https://drive.google.com/file/d/1YnKxwWx6tEpssmSavl6-cww-YmnqlQDF/view?usp=drive_link')
  };

  const label = labelMap[commentKey] || commentKey;
  const elephantIcon = elephantIconMap[commentKey];

  const subtitleEl = document.getElementById('report-subtitle');
  subtitleEl.style.textAlign = 'left';

  // 象アイコンとテキストを組み合わせ
  if (elephantIcon) {
    subtitleEl.innerHTML = `<img src="${elephantIcon}" alt="象アイコン" style="display: inline-block; height: 1.5em; vertical-align: middle; margin-right: 0.5em;" /><span style="vertical-align: middle;">${label}　${totalCount}人</span>`;
  } else {
    subtitleEl.textContent = `${label}　${totalCount}人`;
  }
}

async function fetchAndRenderCommentPage(commentKey) {
  currentCommentData = null;
  currentCommentPageIndex = 0;
  currentCommentSheetName = getCommentSheetName(currentClinicForModal, commentKey);

  if (!currentCommentSheetName) {
    console.error("無効なコメントキー:", commentKey);
    return;
  }

  showLoading(true, `コメントシート (${currentCommentSheetName}) を読み込み中...`);
  document.getElementById('slide-body').innerHTML = '';

  // 編集モードをリセット
  window.commentEditMode = false;

  const controlsContainer = document.getElementById('comment-controls');
  if (controlsContainer) {
    controlsContainer.innerHTML = '';
  } else {
    console.error('fetchAndRenderCommentPage: comment-controls element not found!');
  }

  try {
    const response = await fetch('/api/getCommentData', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        centralSheetId: currentCentralSheetId,
        sheetName: currentCommentSheetName
      })
    });

    if (!response.ok) {
      throw new Error(`コメント取得APIエラー (${response.status}): ${await response.text()}`);
    }

    // dataは列ごとの配列 [ [A列データ...], [B列データ...], ... ]
    const data = await response.json(); 

    if (!data || data.length === 0 || (data.length > 0 && data[0].length === 0)) {
      currentCommentData = [];
      document.getElementById('slide-body').innerHTML = '<p class="text-center text-gray-500 py-16">コメントデータがありません</p>';
      updateCommentSubtitle(commentKey, 0);
      renderCommentControls();
    } else {
      // ▼▼▼ 修正：フラット化せずに、列＝ページとしてそのまま使用する ▼▼▼
      currentCommentData = data; 
      
      // 全コメント数をカウント（サブタイトル用）
      const totalCount = data.reduce((sum, col) => sum + (col ? col.length : 0), 0);
      
      updateCommentSubtitle(commentKey, totalCount);
      renderCommentPage(0); // 最初のページ（A列）を描画
    }

  } catch (e) {
    console.error('Comment data fetch error:', e);
    document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">コメントデータ取得失敗<br>(${e.message})</p>`;
  } finally {
    showLoading(false);
  }
}

function renderCommentPage(pageIndex) {
  if (!currentCommentData) return;

  currentCommentPageIndex = pageIndex;
  const columnData = currentCommentData[pageIndex] || [];

  const bodyEl = document.getElementById('slide-body');
  bodyEl.innerHTML = '';
  bodyEl.style.overflowY = window.commentEditMode ? 'auto' : 'hidden'; // 表示モードはスクロール禁止

  if (columnData.length === 0 && currentCommentData.length > 0 && !window.commentEditMode) {
    bodyEl.innerHTML = '<p class="text-center text-gray-500 py-16">(このページは空です)</p>';
  } else {
    if (!window.commentEditMode) {
      // 表示モード: フォントサイズ自動調整で1ページに収める
      const fragment = document.createDocumentFragment();
      columnData.forEach(comment => {
        if (comment) { // 空文字は表示しない
            const p = document.createElement('p');
            p.className = 'comment-display-item';
            p.textContent = comment;
            fragment.appendChild(p);
        }
      });
      bodyEl.appendChild(fragment);

      // コメント表示後にフォントサイズを調整
      adjustCommentFontSizes(bodyEl);
    } else {
      // 編集モード: 12個のテキストエリアを生成して、その列のデータを編集可能にする
      const editContainer = document.createElement('div');
      editContainer.className = 'comment-edit-container';
      editContainer.style.display = 'flex';
      editContainer.style.flexDirection = 'column';
      editContainer.style.gap = '8px';
      editContainer.style.padding = '8px';

      // ▼▼▼ 修正：必ず12個の入力欄を作る（空欄も編集可能にするため） ▼▼▼
      for (let i = 0; i < 12; i++) {
        const comment = columnData[i] || ''; // データがなければ空文字
        
        const itemWrapper = document.createElement('div');
        itemWrapper.className = 'comment-edit-item';
        itemWrapper.style.display = 'flex';
        itemWrapper.style.flexDirection = 'column';
        itemWrapper.style.gap = '4px';

        const label = document.createElement('label');
        label.textContent = `コメント ${i + 1}`;
        label.style.fontSize = '10pt';
        label.style.fontWeight = 'bold';
        label.style.color = '#4b5563';

        const textarea = document.createElement('textarea');
        textarea.className = 'comment-item-textarea';
        textarea.dataset.index = i;
        textarea.value = comment;
        textarea.style.width = '100%';
        textarea.style.minHeight = '60px';
        textarea.style.padding = '8px';
        textarea.style.border = '1px solid #d1d5db';
        textarea.style.borderRadius = '4px';
        textarea.style.fontSize = '11pt';
        textarea.style.fontFamily = 'inherit';
        textarea.style.resize = 'vertical';

        itemWrapper.appendChild(label);
        itemWrapper.appendChild(textarea);
        editContainer.appendChild(itemWrapper);
      }

      bodyEl.appendChild(editContainer);
    }
  }

  renderCommentControls();
}

// コメント一覧のフォントサイズを自動調整する関数
function adjustCommentFontSizes(containerEl) {
  if (!containerEl) return;

  const initialFontSizePt = 12;
  const minFontSizePt = 7;
  const step = 0.5;
  let currentSize = initialFontSizePt;

  // 全てのコメントアイテムに同じフォントサイズを適用
  const commentItems = containerEl.querySelectorAll('.comment-display-item');
  commentItems.forEach(item => {
    item.style.fontSize = currentSize + 'pt';
  });

  // コンテナがあふれている場合、フォントサイズを縮小
  for (let i = 0; i < 100; i++) {
    if (containerEl.scrollHeight <= containerEl.clientHeight) {
      break;
    }
    currentSize -= step;
    if (currentSize < minFontSizePt) {
      currentSize = minFontSizePt;
      break;
    }
    commentItems.forEach(item => {
      item.style.fontSize = currentSize + 'pt';
    });
  }
}

function renderCommentControls() {
  const controlsContainer = document.getElementById('comment-controls');
  if (!controlsContainer) {
    console.error('renderCommentControls: comment-controls element not found!');
    return;
  }
  controlsContainer.innerHTML = '';

  if (!currentCommentData) return;

  const totalPages = Math.max(1, currentCommentData.length);
  const currentPage = currentCommentPageIndex + 1;
  const currentCol = getExcelColumnName(currentCommentPageIndex);

  const prevDisabled = currentCommentPageIndex === 0;
  const nextDisabled = currentCommentPageIndex >= totalPages - 1;
  const prevCol = prevDisabled ? '' : getExcelColumnName(currentCommentPageIndex - 1);
  const nextCol = nextDisabled ? '' : getExcelColumnName(currentCommentPageIndex + 1);

  const isEditMode = window.commentEditMode;

  controlsContainer.innerHTML = `
    <button id="comment-prev" class="btn" ${prevDisabled || isEditMode ? 'disabled' : ''}>&lt; ${prevCol}</button>
    <span>${currentCol}列 (${currentPage} / ${totalPages})</span>
    <button id="comment-next" class="btn" ${nextDisabled || isEditMode ? 'disabled' : ''}>${nextCol} &gt;</button>
    ${isEditMode ?
      '<button id="comment-save-btn" class="btn comment-save-btn">保存</button>' :
      '<button id="comment-edit-btn" class="btn">編集</button>'
    }
  `;

  // イベントリスナーはsetupEventListeners内の委譲ハンドラーで処理されるため、ここでは追加しない
}

function enterCommentEditMode() {
  window.commentEditMode = true;
  renderCommentPage(currentCommentPageIndex);
}

async function saveCommentEdit() {
  const textareas = document.querySelectorAll('.comment-item-textarea');
  if (!textareas || textareas.length === 0) return;

  // 各テキストエリアから編集内容を取得
  const editedComments = [];
  textareas.forEach(textarea => {
    editedComments.push(textarea.value.trim());
  });

  // ローカルデータを更新 (現在のページのみ差し替え)
  // 元のデータより少ない場合も、12個固定で扱うので配列ごと置き換えてOK
  const oldData = currentCommentData[currentCommentPageIndex] || [];
  currentCommentData[currentCommentPageIndex] = editedComments;

  // スプレッドシートに保存
  showLoading(true, 'コメントを保存中...');

  try {
    // ▼▼▼ 修正：列名を動的に取得（A, B, C...） ▼▼▼
    const colName = getExcelColumnName(currentCommentPageIndex);

    // 各行をスプレッドシートに保存 (1〜12行目固定)
    const updatePromises = [];
    for (let i = 0; i < editedComments.length; i++) {
      // ▼▼▼ 修正：ヘッダーなしなので1行目から開始 (i + 1) ▼▼▼
      const rowNumber = i + 1; 
      const cellAddress = `${colName}${rowNumber}`;
      const value = editedComments[i] || ''; // 空の場合は空文字列

      updatePromises.push(
        fetch('/api/updateCommentData', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            centralSheetId: currentCentralSheetId,
            sheetName: currentCommentSheetName,
            cell: cellAddress,
            value: value
          })
        })
      );
    }

    // 12行目までループしているので、「残りのセルをクリア」する追加処理は不要
    // (editedCommentsは常に12要素あるため)

    await Promise.all(updatePromises);

    console.log('コメント保存完了');
    showLoading(false);

    // 編集モードを終了
    window.commentEditMode = false;
    renderCommentPage(currentCommentPageIndex);
  } catch (error) {
    console.error('コメント保存エラー:', error);
    alert('コメントの保存に失敗しました。もう一度お試しください。');
    // エラー時は元のデータに戻す
    currentCommentData[currentCommentPageIndex] = oldData;
    showLoading(false);
  }
}

// ▼▼▼ 市区町村 (テーブルはスクロール許可) ▼▼▼
async function prepareAndShowMunicipalityReport() {
  console.log('Prepare municipality report');
  updateNavActiveState('municipality', null, null);
  showScreen('screen3');
  document.getElementById('report-title').textContent = `アンケート結果　ーご回答者さまの市町村ー`;
  document.getElementById('report-subtitle').textContent = 'ご出産された方の住所（市町村）について教えてください。';
  document.getElementById('report-subtitle').style.textAlign = 'center'; 
  document.getElementById('report-separator').style.display = 'block';
  
  const slideBody = document.getElementById('slide-body');
  slideBody.style.whiteSpace = 'normal';
  slideBody.innerHTML = '<p class="text-center text-gray-500 py-16">市区町村データを読み込み中...</p>';
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.style.overflowY = 'auto'; // 市区町村表のみスクロール許可
  
  showLoading(true, '市区町村データを読み込み中...');
  
  try {
    const response = await fetch('/api/generateMunicipalityReport', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        centralSheetId: currentCentralSheetId,
        clinicName: currentClinicForModal
      })
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`集計APIエラー (${response.status}): ${errorText}`);
    }
    const tableData = await response.json();
    displayMunicipalityTable(tableData);
  } catch (err) {
    console.error('Failed to get municipality report:', err);
    slideBody.innerHTML = `<p class="text-center text-red-500 py-16">集計失敗: ${err.message}</p>`;
  } finally {
    showLoading(false);
  }
}

function displayMunicipalityTable(data) { 
  const slideBody = document.getElementById('slide-body'); 
  if (!data || data.length === 0) { 
    slideBody.innerHTML = '<p class="text-center text-gray-500 py-16">集計データがありません。</p>'; 
    return; 
  } 
  // ▼▼▼ [修正③] 「不明」を割合に関係なく最下段へ ▼▼▼
  const rows = Array.isArray(data) ? [...data] : [];
  const unknownRows = [];
  const normalRows = [];

  rows.forEach(row => {
    if (row.municipality === '不明') {
      unknownRows.push(row);
    } else {
      normalRows.push(row);
    }
  });

  const sortedRows = [...normalRows, ...unknownRows];

  let tableHtml = `
    <div class="municipality-table-container w-full h-full border border-gray-200 rounded-lg">
      <table class="w-full divide-y divide-gray-200">
        <thead class="bg-gray-50 sticky top-0 z-10">
          <tr>
            <th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">都道府県</th>
            <th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">市区町村</th>
            <th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">件数</th>
            <th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">割合</th>
          </tr>
        </thead>
        <tbody class="bg-white divide-y divide-gray-200">
  `;
  sortedRows.forEach(row => { 
    tableHtml += `
      <tr>
        <td class="py-2 font-medium text-gray-900">${row.prefecture}</td>
        <td class="py-2 text-gray-700">${row.municipality}</td>
        <td class="py-2 text-gray-700 text-right">${row.count}</td>
        <td class="py-2 text-gray-700 text-right">${row.percentage.toFixed(2)}%</td>
      </tr>
    `; 
  }); 
  tableHtml += '</tbody></table></div>'; 
  slideBody.innerHTML = tableHtml; 
}

// ▼▼▼ おすすめ理由 ▼▼▼
async function prepareAndShowRecommendationReport() {
  console.log('Prepare recommendation report');
  updateNavActiveState('recommendation', null, null);
  showScreen('screen3');
  document.getElementById('report-title').textContent = 'アンケート結果　ー本病院を選ぶ上で最も参考にしたものー';
  document.getElementById('report-subtitle').textContent = 'ご出産された産婦人科医院への本病院を選ぶ上で最も参考にしたものについて、教えてください。';
  document.getElementById('report-subtitle').style.textAlign = 'center';
  document.getElementById('report-separator').style.display = 'block';
    
  const slideBody = document.getElementById('slide-body');
  slideBody.style.whiteSpace = 'normal';
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.innerHTML = `
    <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-full pt-12">
      <div class="flex flex-col items-center">
        <h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
        <div id="clinic-pie-chart" class="w-full h-[320px] clinic-graph-bg-yellow"></div>
      </div>
      <div class="flex flex-col items-center">
        <h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3>
        <div id="average-pie-chart" class="w-full h-[320px]"></div>
      </div>
    </div>
  `;
    
  try {
    showLoading(true, '集計済みのおすすめ理由データを読込中...');
    const [clinicChartDataRes, overallChartDataRes] = await Promise.all([
      fetch('/api/getRecommendationReport', { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          centralSheetId: currentCentralSheetId, 
          clinicName: currentClinicForModal 
        })
      }),
      fetch('/api/getRecommendationReport', { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          centralSheetId: currentCentralSheetId, 
          clinicName: "全体" 
        })
      })
    ]);
        
    if (!clinicChartDataRes.ok) throw new Error(`貴院データ取得失敗: ${await clinicChartDataRes.text()}`);
    if (!overallChartDataRes.ok) throw new Error(`全体データ取得失敗: ${await overallChartDataRes.text()}`);

    let clinicChartData = await clinicChartDataRes.json();
    let overallChartData = await overallChartDataRes.json();

    // データ内の「インターネット（産院のホームページ）」を「公式サイト」に置換
    clinicChartData = clinicChartData.map(row => {
      if (Array.isArray(row) && row[0] === 'インターネット（産院のホームページ）') {
        return ['公式サイト', row[1]];
      }
      return row;
    });

    overallChartData = overallChartData.map(row => {
      if (Array.isArray(row) && row[0] === 'インターネット（産院のホームページ）') {
        return ['公式サイト', row[1]];
      }
      return row;
    });

    showLoading(false);

    // 左側（貴院）のグラフのみ背景色を設定
    const clinicOpt = {
      is3D: true,
      chartArea: { left: '5%', top: '5%', width: '90%', height: '90%', backgroundColor: '#ffff95' },
      pieSliceText: 'percentage',
      pieSliceTextStyle: { color: 'black', fontSize: 12, bold: true },
      legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 12 } },
      tooltip: { showColorCode: true, textStyle: { fontSize: 12 }, trigger: 'focus' },
      backgroundColor: '#ffff95',
      sliceVisibilityThreshold: 0 // [修正] 自動で「その他」にまとめない
    };
    // 右側（全体平均）のグラフは背景色なし
    const overallOpt = {
      is3D: true,
      chartArea: { left: '5%', top: '5%', width: '90%', height: '90%' },
      pieSliceText: 'percentage',
      pieSliceTextStyle: { color: 'black', fontSize: 12, bold: true },
      legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 12 } },
      tooltip: { showColorCode: true, textStyle: { fontSize: 12 }, trigger: 'focus' },
      sliceVisibilityThreshold: 0 // [修正] 自動で「その他」にまとめない
    };
    const clinicChartEl = document.getElementById('clinic-pie-chart');
    if (!clinicChartEl) throw new Error('グラフ描画エリア(clinic-pie-chart)が見つかりません。');
    const totalClinicCount = clinicChartData.slice(1).reduce((sum, row) => sum + row[1], 0);
    if (totalClinicCount > 0) {
      const d = google.visualization.arrayToDataTable(clinicChartData);
      new google.visualization.PieChart(clinicChartEl).draw(d, clinicOpt);
    } else {
      clinicChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
    }
    const averageChartEl = document.getElementById('average-pie-chart');
    if (!averageChartEl) throw new Error('グラフ描画エリア(average-pie-chart)が見つかりません。');
    const totalOverallCount = overallChartData.slice(1).reduce((sum, row) => sum + row[1], 0);
    if (totalOverallCount > 0) {
      const d = google.visualization.arrayToDataTable(overallChartData);
      new google.visualization.PieChart(averageChartEl).draw(d, overallOpt);
    } else {
      averageChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';
    }
  } catch (err) {
    console.error('Failed to get recommendation report:', err);
    slideBody.innerHTML = `<p class="text-center text-red-500">集計失敗: ${err.message}</p>`;
    showLoading(false);
  }
}

// ▼▼▼ Word Cloud表示 (Screen 3) ▼▼▼
async function prepareAndShowAnalysis(columnType) {
  showLoading(true, `テキスト分析中(${getColumnName(columnType)})...`);
  showScreen('screen3');
  clearAnalysisCharts();
  updateNavActiveState(null, columnType, null);
  showCopyrightFooter(true); 
  
  let tl = [];
  
  // ▼▼▼ [追加] 正しいデータ数を格納する変数 ▼▼▼
  let clinicTotalCount = 0;
  
  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, 0); 
  document.getElementById('report-subtitle').textContent =
    '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。';
  document.getElementById('report-subtitle').style.textAlign = 'left';
  document.getElementById('report-separator').style.display = 'block';
  
  const subNav = document.getElementById('comment-sub-nav');
  const controls = document.getElementById('comment-controls');
  if (subNav) subNav.innerHTML = '';
  if (controls) controls.innerHTML = '';
  
  const slideBody = document.getElementById('slide-body');
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.style.overflowY = 'hidden';
  
  try {
    // ▼▼▼ [修正] データ数(getSheetRowCounts)も並行して取得する ▼▼▼
    const [cd, rowCounts] = await Promise.all([
      getReportDataForCurrentClinic(currentClinicForModal),
      fetch('/api/getSheetRowCounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          centralSheetId: currentCentralSheetId,
          clinicName: currentClinicForModal
        })
      }).then(r => r.json())
    ]);

    // ▼▼▼ [追加] 取得した件数をセット ▼▼▼
    clinicTotalCount = rowCounts.clinicCount || 0;

    switch (columnType) {
      case 'L':
        tl = cd.npsData.rawText || [];
        break;
      case 'I':
        tl = cd.feedbackData.i_column.results || [];
        break;
      case 'J':
        tl = cd.feedbackData.j_column.results || [];
        break;
      case 'M':
        tl = cd.feedbackData.m_column.results || [];
        break;
      default:
        console.error("Invalid column:", columnType);
        showLoading(false);
        return;
    }
  } catch (e) {
    console.error("Error accessing text data:", e);
    slideBody.innerHTML = `<p class="text-center text-red-500 py-16">レポートデータアクセスエラー</p>`;
    showLoading(false);
    return;
  }
  
  // ▼▼▼ [修正] タイトルに clinicTotalCount を使用 ▼▼▼
  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, clinicTotalCount);
  
  if (tl.length === 0) {
    slideBody.innerHTML = `<p class="text-center text-red-500 py-16">分析対象テキストなし</p>`;
    showLoading(false);
    return;
  }
  
  slideBody.innerHTML = `
    <div class="grid grid-cols-2 gap-2 h-full">
      <div class="grid grid-cols-2 grid-rows-2 gap-1 chart-wc-left" style="height: 80%;">
        <div id="noun-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-blue-600 leading-none py-1" style="font-size: 12px;">名詞</h3>
          <div id="noun-chart" class="w-full flex-1"></div>
        </div>
        <div id="verb-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-red-600 leading-none py-1" style="font-size: 12px;">動詞</h3>
          <div id="verb-chart" class="w-full flex-1"></div>
        </div>
        <div id="adj-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-green-600 leading-none py-1" style="font-size: 12px;">形容詞</h3>
          <div id="adj-chart" class="w-full flex-1"></div>
        </div>
        <div id="int-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-gray-600 leading-none py-1" style="font-size: 12px;">感動詞</h3>
          <div id="int-chart" class="w-full flex-1"></div>
        </div>
      </div>
      <div class="flex flex-col justify-start" style="height: 80%;">
        <p class="text-gray-600 text-left leading-tight" style="font-size: 12px; margin-bottom: 8px;">スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。<br>単語の色は品詞の種類で異なります。<br><span class="text-blue-600 font-semibold">青色=名詞</span>、<span class="text-red-600 font-semibold">赤色=動詞</span>、<span class="text-green-600 font-semibold">緑色=形容詞</span>、<span class="text-gray-600 font-semibold">灰色=感動詞</span></p>
        <div id="word-cloud-container" class="border border-gray-200" style="flex: 1; min-height: 0; display: flex; align-items: center; justify-content: center; padding: 0; background: #ffffff;">
          <canvas id="word-cloud-canvas" style="width: 100%; height: 100%; display: block;"></canvas>
        </div>
        <div id="analysis-error" class="text-red-500 text-sm text-center hidden"></div>
      </div>
    </div>
  `;

  try {
    const r = await fetch('/api/analyzeText', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ textList: tl })
    });
    if (!r.ok) {
      const et = await r.text();
      throw new Error(`分析APIエラー(${r.status}): ${et}`);
    }
    const ad = await r.json();
    // ▼▼▼ [修正] キャッシュ保存時にも clinicTotalCount を使う ▼▼▼
    wcAnalysisCache[columnType] = {
      analysisResults: ad.results,
      title: getAnalysisTitle(columnType, clinicTotalCount),
      subtitle: '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。'
    };
    // レイアウト計算を確実に待つため、requestAnimationFrameとsetTimeoutを併用
    requestAnimationFrame(() => {
      setTimeout(() => drawAnalysisCharts(ad.results), 200);
    });
  } catch (error) {
    console.error('!!! Analyze fail:', error);
    document.getElementById('analysis-error').textContent = `分析失敗: ${error.message}`;
    document.getElementById('analysis-error').classList.remove('hidden');
  } finally {
    showLoading(false);
  }
}

// PDF出力専用：ローディングを表示しないバージョン
async function prepareAndShowAnalysisForPrint(columnType) {
  console.log(`[PDF WC] prepareAndShowAnalysisForPrint called for: ${columnType}`);
  showScreen('screen3');
  clearAnalysisCharts();
  updateNavActiveState(null, columnType, null);
  showCopyrightFooter(true);

  let tl = [];
  // ▼▼▼ [追加] 正しいデータ数を格納する変数 ▼▼▼
  let clinicTotalCount = 0;

  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, 0);
  document.getElementById('report-subtitle').textContent =
    '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。';
  document.getElementById('report-subtitle').style.textAlign = 'left';
  document.getElementById('report-separator').style.display = 'block';

  const subNav = document.getElementById('comment-sub-nav');
  const controls = document.getElementById('comment-controls');
  if (subNav) subNav.innerHTML = '';
  if (controls) controls.innerHTML = '';

  const slideBody = document.getElementById('slide-body');
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.style.overflowY = 'hidden';

  try {
    // ▼▼▼ [修正] データ数(getSheetRowCounts)も並行して取得する ▼▼▼
    const [cd, rowCounts] = await Promise.all([
      getReportDataForCurrentClinic(currentClinicForModal),
      fetch('/api/getSheetRowCounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          centralSheetId: currentCentralSheetId,
          clinicName: currentClinicForModal
        })
      }).then(r => r.json())
    ]);

    clinicTotalCount = rowCounts.clinicCount || 0;

    switch (columnType) {
      case 'L':
        tl = cd.npsData.rawText || [];
        break;
      case 'I':
        tl = cd.feedbackData.i_column.results || [];
        break;
      case 'J':
        tl = cd.feedbackData.j_column.results || [];
        break;
      case 'M':
        tl = cd.feedbackData.m_column.results || [];
        break;
      default:
        console.error("Invalid column:", columnType);
        return;
    }
  } catch (e) {
    console.error("Error accessing text data:", e);
    slideBody.innerHTML = `<p class="text-center text-red-500 py-16">レポートデータアクセスエラー</p>`;
    return;
  }

  console.log(`[PDF WC] Text count: ${tl.length}, Total: ${clinicTotalCount}`);
  // ▼▼▼ [修正] タイトルに clinicTotalCount を使用 ▼▼▼
  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, clinicTotalCount);

  if (tl.length === 0) {
    slideBody.innerHTML = `<p class="text-center text-red-500 py-16">分析対象テキストなし</p>`;
    return;
  }

  slideBody.innerHTML = `
    <div class="grid grid-cols-2 gap-2 h-full">
      <div class="grid grid-cols-2 grid-rows-2 gap-1 chart-wc-left" style="height: 80%;">
        <div id="noun-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-blue-600 leading-none py-1" style="font-size: 12px;">名詞</h3>
          <div id="noun-chart" class="w-full flex-1"></div>
        </div>
        <div id="verb-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-red-600 leading-none py-1" style="font-size: 12px;">動詞</h3>
          <div id="verb-chart" class="w-full flex-1"></div>
        </div>
        <div id="adj-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-green-600 leading-none py-1" style="font-size: 12px;">形容詞</h3>
          <div id="adj-chart" class="w-full flex-1"></div>
        </div>
        <div id="int-chart-container" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-gray-600 leading-none py-1" style="font-size: 12px;">感動詞</h3>
          <div id="int-chart" class="w-full flex-1"></div>
        </div>
      </div>
      <div class="flex flex-col justify-start" style="height: 80%;">
        <p class="text-gray-600 text-left leading-tight" style="font-size: 12px; margin-bottom: 8px;">スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。<br>単語の色は品詞の種類で異なります。<br><span class="text-blue-600 font-semibold">青色=名詞</span>、<span class="text-red-600 font-semibold">赤色=動詞</span>、<span class="text-green-600 font-semibold">緑色=形容詞</span>、<span class="text-gray-600 font-semibold">灰色=感動詞</span></p>
        <div id="word-cloud-container" class="border border-gray-200" style="flex: 1; min-height: 0; display: flex; align-items: center; justify-content: center; padding: 0; background: #ffffff;">
          <canvas id="word-cloud-canvas" style="width: 100%; height: 100%; display: block;"></canvas>
        </div>
        <div id="analysis-error" class="text-red-500 text-sm text-center hidden"></div>
      </div>
    </div>
  `;

  console.log(`[PDF WC] HTML layout created`);

  // Google Chartsがロードされるまで待つ
  const waitForGoogleCharts = () => {
    return new Promise((resolve) => {
      if (typeof google !== 'undefined' && google.visualization) {
        console.log(`[PDF WC] Google Charts already loaded`);
        resolve();
      } else {
        console.log(`[PDF WC] Waiting for Google Charts...`);
        const checkInterval = setInterval(() => {
          if (typeof google !== 'undefined' && google.visualization) {
            console.log(`[PDF WC] Google Charts loaded`);
            clearInterval(checkInterval);
            resolve();
          }
        }, 100);
        setTimeout(() => {
          console.log(`[PDF WC] Google Charts wait timeout`);
          clearInterval(checkInterval);
          resolve();
        }, 5000);
      }
    });
  };

  await waitForGoogleCharts();

  try {
    console.log(`[PDF WC] Checking cache for: ${columnType}`);
    // キャッシュがあればそれを使う
    if (wcAnalysisCache[columnType] && wcAnalysisCache[columnType].analysisResults) {
      console.log(`[PDF WC] Using cached analysis results`);
      const cachedResults = wcAnalysisCache[columnType].analysisResults;
      await new Promise(resolve => {
        requestAnimationFrame(() => {
          setTimeout(() => {
            console.log(`[PDF WC] Drawing cached charts`);
            drawAnalysisCharts(cachedResults);
            resolve();
          }, 200);
        });
      });
      return;
    }

    console.log(`[PDF WC] No cache, calling API...`);
    const r = await fetch('/api/analyzeText', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ textList: tl })
    });
    if (!r.ok) {
      throw new Error(`分析APIエラー(${r.status})`);
    }
    const ad = await r.json();
    console.log(`[PDF WC] API response received, results count: ${ad.results ? ad.results.length : 0}`);

    wcAnalysisCache[columnType] = {
      analysisResults: ad.results,
      title: getAnalysisTitle(columnType, clinicTotalCount), // [修正]
      subtitle: '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。'
    };

    await new Promise(resolve => {
      requestAnimationFrame(() => {
        setTimeout(() => {
          console.log(`[PDF WC] Drawing new charts`);
          drawAnalysisCharts(ad.results);
          resolve();
        }, 200);
      });
    });
  } catch (error) {
    console.error('[PDF WC] !!! Analyze fail:', error);
    document.getElementById('analysis-error').textContent = `分析失敗: ${error.message}`;
    document.getElementById('analysis-error').classList.remove('hidden');
  }
}

function drawAnalysisCharts(results) { 
  if (!results || results.length === 0) {
    console.log("No analysis results.");
    document.getElementById('analysis-error').textContent = '分析結果なし';
    document.getElementById('analysis-error').classList.remove('hidden');
    return;
  } 
  const nouns = results.filter(r => r.pos === '名詞');
  const verbs = results.filter(r => r.pos === '動詞');
  const adjs  = results.filter(r => r.pos === '形容詞');
  const ints  = results.filter(r => r.pos === '感動詞');
    
  const barOpt = {
    bars: 'horizontal',
    legend: { position: 'none' },
    hAxis: { title: 'スコア(出現頻度)', minValue: 0, textStyle: { fontSize: 12 }, titleTextStyle: { fontSize: 12 } },
    vAxis: { title: null, textStyle: { fontSize: 10 }, titleTextStyle: { fontSize: 12 } }, 
    chartArea: { height: '90%', width: '70%', left: '25%', top: '5%' }, 
    backgroundColor: 'transparent'
  };
    
  drawSingleBarChart(nouns.slice(0, 8), 'noun-chart', { ...barOpt, colors: ['#3b82f6'], width: '100%', height: '100%' });
  drawSingleBarChart(verbs.slice(0, 8), 'verb-chart', { ...barOpt, colors: ['#ef4444'], width: '100%', height: '100%' });
  drawSingleBarChart(adjs.slice(0, 8), 'adj-chart',  { ...barOpt, colors: ['#22c55e'], width: '100%', height: '100%' });
  drawSingleBarChart(ints.slice(0, 8), 'int-chart',  { ...barOpt, colors: ['#6b7280'], width: '100%', height: '100%' });
    
  const wl = results.map(r => [r.word, r.score]).slice(0, 100);
  const pm = results.reduce((map, item) => { map[item.word] = item.pos; return map; }, {});
  const cv = document.getElementById('word-cloud-canvas');

  if (WordCloud.isSupported && cv) {
    try {
      const dpr = window.devicePixelRatio || 1;
      const rect = cv.getBoundingClientRect();
      cv.width = rect.width * dpr;
      cv.height = rect.height * dpr;
      const ctx = cv.getContext('2d');
      ctx.scale(dpr, dpr);
      const logicalWidth = rect.width;
      const logicalHeight = rect.height;
      const minDimension = Math.min(logicalWidth, logicalHeight);

      // スコアの最大値と最小値を取得
      const scores = wl.map(item => item[1]);
      const maxScore = Math.max(...scores);
      const minScore = Math.min(...scores);
      const scoreRange = maxScore - minScore || 1;

      // データ数に応じて基本サイズを調整（より保守的なサイズ設定）
      const dataCount = wl.length;
      let baseMaxSize = 60;  // 最大スコア時のフォントサイズ
      let baseMinSize = 14;  // 最小スコア時のフォントサイズ

      if (dataCount <= 10) {
        baseMaxSize = 100;
        baseMinSize = 35;
      } else if (dataCount <= 20) {
        baseMaxSize = 80;
        baseMinSize = 28;
      } else if (dataCount <= 30) {
        baseMaxSize = 70;
        baseMinSize = 24;
      } else if (dataCount <= 40) {
        baseMaxSize = 65;
        baseMinSize = 20;
      } else if (dataCount <= 50) {
        baseMaxSize = 60;
        baseMinSize = 18;
      } else if (dataCount <= 70) {
        baseMaxSize = 55;
        baseMinSize = 16;
      } else if (dataCount <= 90) {
        baseMaxSize = 50;
        baseMinSize = 14;
      }

      // キャンバスサイズに基づいてスケーリング係数を計算
      const sizeScale = Math.min(logicalWidth, logicalHeight) / 500;
      const scaledMaxSize = baseMaxSize * sizeScale;
      const scaledMinSize = baseMinSize * sizeScale;

      // スコアの割合に基づいてサイズを計算
      const options = {
        list: wl,
        gridSize: Math.max(1, Math.round(minDimension / 200)),
        weightFactor: (score) => {
          // スコアを0-1の範囲に正規化
          const normalizedScore = (score - minScore) / scoreRange;
          // 最小サイズと最大サイズの間で線形補間
          return scaledMinSize + (scaledMaxSize - scaledMinSize) * normalizedScore;
        },
        fontFamily: 'Noto Sans JP,sans-serif',
        color: (w) => {
          const p = pm[w] || '不明';
          switch (p) {
            case '名詞': return '#3b82f6';
            case '動詞': return '#ef4444';
            case '形容詞': return '#22c55e';
            case '感動詞': return '#6b7280';
            default: return '#a8a29e';
          }
        },
        backgroundColor: 'transparent',
        clearCanvas: true,
        rotateRatio: 0,
        drawOutOfBound: false,
        shrinkToFit: true,
        minSize: scaledMinSize * 0.3,
        shuffle: true,
        wait: 5,
        abortThreshold: 2000,
        abort: () => false,
        origin: [logicalWidth / 2, logicalHeight / 2]
      };
      WordCloud(cv, options);
    } catch (wcError) {
      console.error("Error drawing WordCloud:", wcError);
      document.getElementById('word-cloud-container').innerHTML = `<p class="text-center text-red-500">ワードクラウド描画エラー:${wcError.message}</p>`;
    }
  } else {
    console.warn("WordCloud unsupported/canvas missing.");
    document.getElementById('word-cloud-container').innerHTML = '<p class="text-center text-gray-500">ワードクラウド非対応</p>';
  } 
}

function drawSingleBarChart(data, elementId, options) {
  const c = document.getElementById(elementId);
  if (!c) { console.error(`Element not found: ${elementId}`); return; }
  if (!data || data.length === 0) {
    c.innerHTML = '<p class="text-center text-gray-500 text-sm py-4">データなし</p>';
    return;
  }

  // ▼▼▼ [追加] 太さの計算ロジック ▼▼▼
  // 「最大8個」のときの太さを基準(例えば60%)とし、データ数に応じて比率を下げる
  // 例: 8個なら60%、1個なら (1/8)*60 = 7.5% の太さになる
  const maxItems = 8;
  const baseWidthPercent = 60; // 8個揃ったときの棒の太さ(%)
  const currentCount = data.length;
  const dynamicGroupWidth = (currentCount / maxItems) * baseWidthPercent + '%';

  // 計算した太さをオプションに適用
  const finalOptions = {
    ...options,
    bar: { groupWidth: dynamicGroupWidth }
  };
  // ▲▲▲ 追加ここまで ▲▲▲

  const cd = [['単語','スコア',{ role:'style' }]];
  const color = options.colors && options.colors.length > 0 ? options.colors[0] : '#a8a29e';
  data.slice().forEach(item => { cd.push([item.word, item.score, color]); });
  try {
    const dt = google.visualization.arrayToDataTable(cd);
    const chart = new google.visualization.BarChart(c);
    chart.draw(dt, finalOptions);
  } catch (chartError) {
    console.error(`Error drawing bar chart for ${elementId}:`, chartError);
    c.innerHTML = `<p class="text-center text-red-500 text-sm py-4">グラフ描画エラー<br>${chartError.message}</p>`;
  }
}

function clearAnalysisCharts() {
  const nounChart = document.getElementById('noun-chart'); if (nounChart) nounChart.innerHTML = '';
  const verbChart = document.getElementById('verb-chart'); if (verbChart) verbChart.innerHTML = '';
  const adjChart  = document.getElementById('adj-chart');  if (adjChart)  adjChart.innerHTML  = '';
  const intChart  = document.getElementById('int-chart');  if (intChart)  intChart.innerHTML  = '';
  const c = document.getElementById('word-cloud-canvas');
  if (c) { const x = c.getContext('2d'); x.clearRect(0, 0, c.width, c.height); }
  const analysisError = document.getElementById('analysis-error');
  if (analysisError) { analysisError.classList.add('hidden'); analysisError.textContent = ''; }
}

function getAnalysisTitle(columnType, count) {
  const bt = `アンケート結果　ー${getColumnName(columnType)}ー`;
  return `${bt}　※全回答数${count}件ー`;
}
function getColumnName(columnType) { 
  switch (columnType) {
    case 'L': return 'NPS推奨度 理由';
    case 'I':
    case 'I_good':
    case 'I_bad': return '良かった点や悪かった点など';
    case 'J': return '印象に残ったスタッフへのコメント';
    case 'M': return 'お産にかかわるご意見・ご感想';
    default: return '不明';
  } 
}

// --- ▼▼▼ AI詳細分析 (Screen 5) ▼▼▼

// PDF出力専用：ローディングを表示しないバージョン
async function prepareAndShowAIAnalysisForPrint(analysisType, tabId) {
  console.log(`[PDF] Prepare AI analysis: ${analysisType} - ${tabId}`);
  const clinicName = currentClinicForModal;

  showScreen('screen5');
  updateNavActiveState(null, null, analysisType);
  toggleEditDetailedAnalysis(false);
  showCopyrightFooter(true, 'screen5');

  const errorDiv = document.getElementById('detailed-analysis-error');
  errorDiv.classList.add('hidden');
  errorDiv.textContent = '';

  document.getElementById('detailed-analysis-title').textContent = getDetailedAnalysisTitleFull(analysisType);
  document.getElementById('detailed-analysis-subtitle').textContent = getSubtitleForItem(analysisType);

  clearDetailedAnalysisDisplay();

  // 指定されたタブをアクティブに
  document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => {
    const isActive = button.dataset.tabId === tabId;
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  // タブコンテンツエリアの表示切り替え
  document.querySelectorAll('.ai-panel').forEach(panel => {
    const panelId = panel.id;
    const isActive = panelId === `content-${tabId}`;
    panel.classList.toggle('is-active', isActive);
    panel.setAttribute('aria-hidden', isActive ? 'false' : 'true');
  });

  // タブをロード（ローディング表示なし）
  try {
    const response = await fetch('/api/getSingleAnalysisCell', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        centralSheetId: currentCentralSheetId,
        clinicName: clinicName,
        columnType: analysisType,
        tabId: tabId
      })
    });
    if (!response.ok) {
      throw new Error(`API error: ${response.status}`);
    }
    const data = await response.json();

    const content = data.content || '（データがありません）';
    const pentagonText = data.pentagonText || getTabLabel(tabId);

    console.log(`[PDF] Content loaded for ${analysisType}-${tabId}:`, content ? `${content.substring(0, 50)}...` : 'empty');

    // 表示を更新
    const displayElement = document.getElementById(`display-${tabId}`);
    const textareaElement = document.getElementById(`textarea-${tabId}`);

    console.log(`[PDF] Display element:`, displayElement ? 'found' : 'NOT FOUND');

    if (displayElement) {
      displayElement.textContent = content;
      adjustFontSize(displayElement);
    }
    if (textareaElement) {
      textareaElement.value = content;
    }

    // 五角形を更新（アクティブなタブ内のみ）
    const activePanel = document.querySelector(`#content-${tabId}`);
    const shapeElement = activePanel ? activePanel.querySelector('.ai-analysis-shape') : null;
    if (shapeElement) {
      shapeElement.textContent = pentagonText;
    }

    // サブタイトルを更新
    const subtitleElement = document.getElementById('detailed-analysis-subtitle');
    if (subtitleElement && data.subtitle) {
      subtitleElement.textContent = data.subtitle;
    }
  } catch (error) {
    console.error('[PDF] Error loading AI analysis:', error);
  }
}

// =================================================================
// === ▼▼▼ [置き換え 1/2] prepareAndShowDetailedAnalysis ▼▼▼ ===
async function prepareAndShowDetailedAnalysis(analysisType) {
  console.log(`Prepare detailed analysis: ${analysisType}`);
  const clinicName = currentClinicForModal;
  
  // 1. 画面表示 (ローディングは switchTab が行う)
  showScreen('screen5');
  updateNavActiveState(null, null, analysisType);
  toggleEditDetailedAnalysis(false); 
  showCopyrightFooter(true, 'screen5'); 
  
  // 2. エラー表示をリセット
  const errorDiv = document.getElementById('detailed-analysis-error');
  errorDiv.classList.add('hidden');
  errorDiv.textContent = '';
  
  // 3. メインタイトル・サブタイトルを設定
  document.getElementById('detailed-analysis-title').textContent = getDetailedAnalysisTitleFull(analysisType);
  document.getElementById('detailed-analysis-subtitle').textContent = getSubtitleForItem(analysisType);
  
  // 4. 表示エリアをクリア
  clearDetailedAnalysisDisplay();
  
  // 5. デフォルトタブ 'analysis' をアクティブに
  document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => {
    const isActive = button.dataset.tabId === 'analysis';
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });
  
  // 6. デフォルトタブをロード
  await switchTab('analysis'); 
}
// =================================================================

// 再実行
async function handleRegenerateDetailedAnalysis() {
  const typeName = getDetailedAnalysisTitleBase(currentDetailedAnalysisType);
  if (!confirm(`「${typeName}」のAI分析を再実行しますか？\n\n・現在の分析内容は破棄されます。\n・編集中の内容は保存されません。`)) {
    return;
  }
  showLoading(true, `AI分析を再実行中...\n(${typeName})`);
  toggleEditDetailedAnalysis(false); 
  await runDetailedAnalysisGeneration(currentDetailedAnalysisType);
}

// AI実行
async function runDetailedAnalysisGeneration(analysisType) {
  const errorDiv = document.getElementById('detailed-analysis-error');
  errorDiv.classList.add('hidden');
  try {
    const response = await fetch('/api/generateDetailedAnalysis', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        centralSheetId: currentCentralSheetId,
        clinicName: currentClinicForModal,
        columnType: analysisType
      })
    });
    if (!response.ok) {
      const errorText = await response.text();
      if (response.status === 404) {
        throw new Error('分析対象のテキストが0件のため、AI分析を実行できませんでした。');
      }
      throw new Error(`AI分析APIエラー (${response.status}): ${errorText}`);
    }
    // analysisJson = { analysis: { themes: [...] }, suggestions: { ... }, ... }
    const analysisJson = await response.json(); 
    
    // JSON → 表示用テキストへマッピング
    const mappedData = formatAiJsonToMappedObject(analysisJson, analysisType);
    
    // 表示エリア/Textareaにセット
    document.getElementById('display-analysis').textContent = mappedData.analysis;
    document.getElementById('textarea-analysis').value = mappedData.analysis;
    document.getElementById('display-suggestions').textContent = mappedData.suggestions;
    document.getElementById('textarea-suggestions').value = mappedData.suggestions;
    document.getElementById('display-overall').textContent = mappedData.overall;
    document.getElementById('textarea-overall').value = mappedData.overall;

    // 再実行後、'analysis' タブを更新
    await switchTab('analysis');
    
  } catch (err) {
    console.error('!!! AI Generation failed:', err);
    errorDiv.textContent = `AI分析実行エラー: ${err.message}`;
    errorDiv.classList.remove('hidden');
    clearDetailedAnalysisDisplay();
  } finally {
    showLoading(false);
  }
}

// [新規] JSON→Object マッピング
function formatAiJsonToMappedObject(analysisJson, columnType) {
  const analysisText = (analysisJson.analysis && analysisJson.analysis.themes)
    ? analysisJson.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n')
    : '（分析データがありません）';
  const suggestionsText = (analysisJson.suggestions && analysisJson.suggestions.items)
    ? analysisJson.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n')
    : '（改善提案データがありません）';
  const overallText = (analysisJson.overall && analysisJson.overall.summary)
    ? analysisJson.overall.summary
    : '（総評データがありません）';
  return { analysis: analysisText, suggestions: suggestionsText, overall: overallText };
}

/**
 * [修正] 表示＋フォントサイズ調整
 */
function displayDetailedAnalysis(data) {
  const displayAnalysis = document.getElementById('display-analysis');
  const displaySuggestions = document.getElementById('display-suggestions');
  const displayOverall = document.getElementById('display-overall');

  const analysisText = data.analysis || '（データがありません）';
  const suggestionsText = data.suggestions || '（データがありません）';
  const overallText = data.overall || '（データがありません）';

  displayAnalysis.textContent = analysisText;
  document.getElementById('textarea-analysis').value = analysisText;
  
  displaySuggestions.textContent = suggestionsText;
  document.getElementById('textarea-suggestions').value = suggestionsText;
  
  displayOverall.textContent = overallText;
  document.getElementById('textarea-overall').value = overallText;
  
  // 全タブのフォント調整
  adjustFontSize(displayAnalysis);
  adjustFontSize(displaySuggestions);
  adjustFontSize(displayOverall);

  // 編集用 Textarea の高さ調整
  ['analysis', 'suggestions', 'overall'].forEach(tabId => {
    const textarea = document.getElementById(`textarea-${tabId}`);
    if (textarea) {
      textarea.style.height = 'auto';
      textarea.style.height = (textarea.scrollHeight + 5) + 'px';
    }
  });
}

function clearDetailedAnalysisDisplay() {
  document.getElementById('display-analysis').textContent = '';
  document.getElementById('textarea-analysis').value = '';
  document.getElementById('display-suggestions').textContent = '';
  document.getElementById('textarea-suggestions').value = '';
  document.getElementById('display-overall').textContent = '';
  document.getElementById('textarea-overall').value = '';
}

// =================================================================
// === ▼▼▼ フォントサイズ調整（スクロール禁止対応） ▼▼▼ ===
function isOverflown(element) {
  if (!element) return false;
  return element.scrollHeight > element.clientHeight;
}
function adjustFontSize(element) {
  if (!element) return;
  const initialFontSizePt = 12;
  const minFontSizePt = 5;  // 7pt → 5pt に変更
  const step = 0.5;
  let currentSize = initialFontSizePt;

  // 初期状態に戻す
  element.style.fontSize = currentSize + 'pt';
  element.style.lineHeight = '1.5';

  // フォントサイズを縮小
  for (let i = 0; i < 100; i++) {
    if (!isOverflown(element)) return;
    currentSize -= step;
    element.style.fontSize = currentSize + 'pt';
    if (currentSize <= minFontSizePt) break;
  }

  // フォントサイズを最小にしても収まらない場合は行間を調整
  if (isOverflown(element)) {
    element.style.lineHeight = '1.3';
    if (isOverflown(element)) {
      element.style.lineHeight = '1.2';
      if (isOverflown(element)) {
        element.style.lineHeight = '1.1';
      }
    }
  }
}
// =================================================================

// =================================================================
// === ▼▼▼ タブ切り替え（サブタイトル/五角形/本文を正しく反映） ▼▼▼ ===
function handleTabClick(event) { 
  const btn = event.target.closest('.tab-button');
  if (!btn) return;
  const tabId = btn.dataset.tabId; // 'analysis', 'suggestions', 'overall'
  if (tabId) switchTab(tabId);
}

// [置き換え] タブ切り替え本体（①ページ切り替えの時と同じように差し替え描画）
async function switchTab(tabId) {
  console.log(`Switching tab to: ${tabId}`);

  // 1. APIから該当セルを取得
  showLoading(true, '分析データを読み込み中...');
  const errorDiv = document.getElementById('detailed-analysis-error');
  errorDiv.classList.add('hidden');

  let content = '（データがありません）';
  let pentagonText = '...';

  try {
    const response = await fetch('/api/getSingleAnalysisCell', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        centralSheetId: currentCentralSheetId,
        clinicName: currentClinicForModal,
        columnType: currentDetailedAnalysisType,
        tabId: tabId
      })
    });

    if (!response.ok) {
      throw new Error(`APIエラー (${response.status}): ${await response.text()}`);
    }

    const data = await response.json();
    content = data.content;
    pentagonText = data.pentagonText;

    // キャッシュに保存（PDF出力用）
    if (!aiAnalysisCache[currentDetailedAnalysisType]) {
      aiAnalysisCache[currentDetailedAnalysisType] = {};
    }
    aiAnalysisCache[currentDetailedAnalysisType][tabId] = {
      content: content,
      pentagonText: pentagonText
    };

  } catch (err) {
    console.error('Failed to fetch single analysis cell:', err);
    errorDiv.textContent = `データ読み込み失敗: ${err.message}`;
    errorDiv.classList.remove('hidden');

  } finally {
    showLoading(false);
  }

  // 2. タブの active 表示更新
  document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => {
    const isActive = button.dataset.tabId === tabId;
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-selected', isActive ? 'true' : 'false');
  });

  // 3. ▼▼▼ [修正①] CSS差し替え（.ai-panel の .is-active 切替）でスムーズに表示変更 ▼▼▼
  const allPanels = document.querySelectorAll('#detailed-analysis-content-area .ai-panel');
  const activePanel = document.getElementById(`content-${tabId}`);

  allPanels.forEach(panel => {
    panel.classList.remove('is-active');
    panel.setAttribute('aria-hidden', 'true');
  });

  if (activePanel) {
    activePanel.classList.add('is-active');
    activePanel.setAttribute('aria-hidden', 'false');

    // 五角形テキスト
    const shape = activePanel.querySelector('.ai-analysis-shape');
    if (shape) {
      console.log(`[switchTab] Updating pentagon text to: ${pentagonText}`);
      shape.textContent = pentagonText;
    } else {
      console.error(`[switchTab] Pentagon shape element not found in panel: content-${tabId}`);
    }

    // 本文
    const displayContent = activePanel.querySelector('.ai-analysis-content');
    if (displayContent) {
      console.log(`[switchTab] Updating content (length: ${content.length})`);
      displayContent.textContent = content;
      if (!isEditingDetailedAnalysis) {
        adjustFontSize(displayContent);
      }
    } else {
      console.error(`[switchTab] Display content element not found in panel: content-${tabId}`);
    }

    // 編集用Textarea
    const textarea = activePanel.querySelector('.edit-textarea');
    if (textarea) {
      textarea.value = content;
    }
  } else {
    console.error(`[switchTab] Active panel not found: content-${tabId}`);
  }
}
// =================================================================

// [変更なし] メインタイトル (H1)
function getDetailedAnalysisTitleFull(analysisType) {
  switch (analysisType) {
    case 'L': return '知人に病院を紹介したいと思う理由の分析';
    case 'I_bad': return '「悪かった点」の分析と改善策';
    case 'I_good': return '「良かった点」の分析と改善策';
    case 'J': return '印象に残ったスタッフへのコメント分析とスタッフ評価';
    case 'M': return 'お産に関わるご意見の分析と改善策の提案';
    default: return 'AI詳細分析レポート';
  }
}

// ▼▼▼ サブタイトル（項目で固定） ▼▼▼
function getSubtitleForItem(analysisType) {
  const base = '※コメントでいただいたフィードバックを元に分析しています\n';
  const subtitleMap = {
    'L': '知人への推奨理由を分析すると、以下の主要なテーマが浮かび上がります。',
    'I_bad': 'フィードバックの中で挙げられた「悪かった点」を分析すると、\n患者にとって以下の要素が特に課題として感じられていることが分かります。',
    'I_good': 'フィードバックの中で挙げられた「良かった点」を分析すると、\n以下の要素が患者にとって特に高く評価されていることが分かります。',
    'J': '印象に残ったスタッフに対するコメントから、いくつかの重要なテーマが浮かび上がります。\nこれらのテーマは、スタッフの評価においても重要なポイントとなります。',
    'M': '患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。'
  };
  return base + (subtitleMap[analysisType] || subtitleMap['L']);
}

// [変更なし] 
function getDetailedAnalysisTitleBase(analysisType) {
  switch (analysisType) {
    case 'L': return 'NPS推奨度 理由';
    case 'I_good': return '良かった点';
    case 'I_bad': return '悪かった点';
    case 'J': return '印象に残ったスタッフへのコメント';
    case 'M': return 'お産にかかわるご意見・ご感想';
    default: return '不明';
  }
}

// =================================================================
// === ▼▼▼ 編集モード切替（バグ修正） ▼▼▼ ===
// =================================================================
function toggleEditDetailedAnalysis(isEdit) {
  isEditingDetailedAnalysis = isEdit;
  const editBtn = document.getElementById('edit-detailed-analysis-btn');
  const regenBtn = document.getElementById('regenerate-detailed-analysis-btn');
  const saveBtn = document.getElementById('save-detailed-analysis-btn');
  const cancelBtn = document.getElementById('cancel-edit-detailed-analysis-btn');

  const displayAreas = [
    document.getElementById('display-analysis'),
    document.getElementById('display-suggestions'),
    document.getElementById('display-overall')
  ];
  const editAreas = [
    document.getElementById('edit-analysis'),
    document.getElementById('edit-suggestions'),
    document.getElementById('edit-overall')
  ];

  if (isEditingDetailedAnalysis) {
    // 編集開始
    editBtn.classList.add('hidden');
    regenBtn.classList.add('hidden');
    saveBtn.classList.remove('hidden');
    cancelBtn.classList.remove('hidden');
    displayAreas.forEach(el => el.classList.add('hidden'));
    editAreas.forEach(el => el.classList.remove('hidden'));
        
    const activeTab = document.querySelector('#ai-tab-nav .tab-button.active')?.dataset.tabId || 'analysis';
    const activeTextarea = document.getElementById(`textarea-${activeTab}`);
    if (activeTextarea) {
      activeTextarea.style.height = 'auto';
      activeTextarea.style.height = (activeTextarea.scrollHeight + 5) + 'px';
    }
    
  } else {
    // 編集終了
    editBtn.classList.remove('hidden');
    regenBtn.classList.remove('hidden');
    saveBtn.classList.add('hidden');
    cancelBtn.classList.add('hidden');
    editAreas.forEach(el => el.classList.add('hidden'));
    displayAreas.forEach(el => el.classList.remove('hidden'));

    const activeTab = document.querySelector('#ai-tab-nav .tab-button.active')?.dataset.tabId || 'analysis';
    const activeDisplayArea = document.getElementById(`display-${activeTab}`);
    if (activeDisplayArea) {
      adjustFontSize(activeDisplayArea);
    }
  }
}
// =================================================================

// 保存
async function saveDetailedAnalysisEdits() {
  showLoading(true, '変更を保存中...');
  const analysisContent = document.getElementById('textarea-analysis').value;
  const suggestionsContent = document.getElementById('textarea-suggestions').value;
  const overallContent = document.getElementById('textarea-overall').value;

  try {
    const response = await fetch('/api/updateDetailedAnalysis', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        centralSheetId: currentCentralSheetId,
        clinicName: currentClinicForModal,
        columnType: currentDetailedAnalysisType, 
        analysis: analysisContent,
        suggestions: suggestionsContent,
        overall: overallContent
      })
    });
        
    if (!response.ok) {
      throw new Error(`保存失敗 (${response.status}): ${await response.text()}`);
    }
        
    document.getElementById('display-analysis').textContent = analysisContent;
    document.getElementById('display-suggestions').textContent = suggestionsContent;
    document.getElementById('display-overall').textContent = overallContent;
        
    toggleEditDetailedAnalysis(false);
    alert('保存しました。');
  } catch (e) {
    console.error("Failed to save edits:", e);
    alert(`保存中にエラーが発生しました。\n${e.message}`);
  } finally {
    showLoading(false);
  }
}

// --- 汎用関数 ---
function updateNavActiveState(activeReportType, activeAnalysisType, activeDetailedAnalysisType) {
  const navs = ['#report-nav', '#report-nav-screen5']; 
  navs.forEach(navSelector => {
    document.querySelectorAll(`${navSelector} .btn`).forEach(btn => {
      if (
        (activeReportType && btn.dataset.reportType === activeReportType) ||
        (activeAnalysisType && btn.dataset.analysisType === activeAnalysisType) ||
        (activeDetailedAnalysisType && btn.dataset.detailedAnalysisType === activeDetailedAnalysisType)
      ) {
        btn.classList.add('btn-active');
      } else {
        btn.classList.remove('btn-active');
      }
    });
  });
}

function showScreen(screenId) {
  document.querySelectorAll('.screen').forEach(el => el.classList.add('hidden'));
  document.getElementById(screenId).classList.remove('hidden');
}
let isPdfGenerating = false; // PDF生成中フラグ

function showLoading(isLoading, message = '') {
  const o = document.getElementById('loading-overlay');
  const m = document.getElementById('loading-message');
  if (isLoading) {
    m.textContent = message;
    o.classList.remove('hidden');
  } else {
    // PDF生成中は非表示にしない
    if (isPdfGenerating) {
      console.log('[showLoading] PDF生成中のため、ローディングを非表示にしません');
      return;
    }
    o.classList.add('hidden');
    m.textContent = '';
  }
}

// Google DriveのURLからIDを抽出して/image/${id}形式のURLに変換
function convertGoogleDriveUrl(driveUrl) {
  const match = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return `/image/${match[1]}`;
  }
  return driveUrl;
}

// ▼▼▼ フッター表示切替 ▼▼▼
function showCopyrightFooter(show, screenId = 'screen3') {
  let footer;
  if (screenId === 'screen3') {
    footer = document.getElementById('report-copyright');
  } else if (screenId === 'screen5') {
    footer = document.querySelector('#screen5 .report-copyright');
  }
  if (footer) footer.style.display = show ? 'block' : 'none';
}

// --- PDF出力機能 ---
async function handlePdfExport() {
  console.log('[PDF Export] 開始');
  isPdfGenerating = true; // PDF生成開始
  showLoading(true, '印刷プレビューを準備中...');

  try {
    console.log('[PDF Export] クリニック名:', currentClinicForModal);
    console.log('[PDF Export] シートID:', currentCentralSheetId);

    const printContainer = document.getElementById('print-container');
    console.log('[PDF Export] printContainer:', printContainer);
    printContainer.innerHTML = '';

    // 全ページタイプを定義（コメント一覧とWCを交互に配置）
    const allPages = [
      'cover',
      'toc',
      'summary',
      'age',
      'children',
      'income',
      'municipality',
      'satisfaction_b',
      'satisfaction_c',
      'satisfaction_d',
      'satisfaction_e',
      'satisfaction_f',
      'satisfaction_g',
      'satisfaction_h',
      'recommendation',
      'nps_score',
      // NPSコメント一覧（5ページ）→ WC-NPS
      { type: 'nps_comments', key: 'L_10' },
      { type: 'nps_comments', key: 'L_9' },
      { type: 'nps_comments', key: 'L_8' },
      { type: 'nps_comments', key: 'L_7' },
      { type: 'nps_comments', key: 'L_6_under' },
      { type: 'word_cloud', analysisKey: 'L' },  // WC-NPS
      // 良かった点コメント → WC-良い点
      { type: 'nps_comments', key: 'I' },
      { type: 'word_cloud', analysisKey: 'I' },  // WC-良い点
      // 印象スタッフコメント → WC-スタッフ
      { type: 'nps_comments', key: 'J' },
      { type: 'word_cloud', analysisKey: 'J' },  // WC-スタッフ
      // お産意見コメント → WC-お産
      { type: 'nps_comments', key: 'M' },
      { type: 'word_cloud', analysisKey: 'M' },  // WC-お産
      // AI分析（5種類 × 3タブ = 15ページ）
      // AI-NPS
      { type: 'ai_analysis', analysisType: 'L', subtype: 'analysis' },
      { type: 'ai_analysis', analysisType: 'L', subtype: 'suggestions' },
      { type: 'ai_analysis', analysisType: 'L', subtype: 'overall' },
      // AI-悪い点
      { type: 'ai_analysis', analysisType: 'I_bad', subtype: 'analysis' },
      { type: 'ai_analysis', analysisType: 'I_bad', subtype: 'suggestions' },
      { type: 'ai_analysis', analysisType: 'I_bad', subtype: 'overall' },
      // AI-良い点
      { type: 'ai_analysis', analysisType: 'I_good', subtype: 'analysis' },
      { type: 'ai_analysis', analysisType: 'I_good', subtype: 'suggestions' },
      { type: 'ai_analysis', analysisType: 'I_good', subtype: 'overall' },
      // AI-スタッフ
      { type: 'ai_analysis', analysisType: 'J', subtype: 'analysis' },
      { type: 'ai_analysis', analysisType: 'J', subtype: 'suggestions' },
      { type: 'ai_analysis', analysisType: 'J', subtype: 'overall' },
      // AI-お産
      { type: 'ai_analysis', analysisType: 'M', subtype: 'analysis' },
      { type: 'ai_analysis', analysisType: 'M', subtype: 'suggestions' },
      { type: 'ai_analysis', analysisType: 'M', subtype: 'overall' }
    ];

    console.log('[PDF Export] 全', allPages.length, 'ページを生成します');

    // 元の画面を非表示にしてバックグラウンドで処理
    const appFrame = document.getElementById('app-frame');
    appFrame.style.opacity = '0';
    appFrame.style.pointerEvents = 'none';

    // 各ページを順番に表示してクローン
    for (let i = 0; i < allPages.length; i++) {
      const pageInfo = allPages[i];
      const pageType = typeof pageInfo === 'string' ? pageInfo : pageInfo.type;

      console.log(`[PDF Export] ページ ${i + 1}/${allPages.length}: ${pageType}`);
      const progress = Math.round(((i + 1) / allPages.length) * 100);
      showLoading(true, `PDF作成中... ${i + 1}/${allPages.length} (${progress}%)`);

      // ページを表示
      if (typeof pageInfo === 'string') {
        // 通常のレポートページ
        await prepareAndShowReport(pageInfo);

        // 表紙の場合は画像読み込みを待つ
        if (pageInfo === 'cover') {
          const coverImg = document.querySelector('.cover-background img');
          if (coverImg) {
            if (!coverImg.complete) {
              await new Promise(resolve => {
                coverImg.onload = resolve;
                coverImg.onerror = resolve;
                setTimeout(resolve, 3000); // タイムアウト
              });
            }
          }
          await new Promise(resolve => setTimeout(resolve, 2000));
        } else {
          // グラフの描画を待つ
          await new Promise(resolve => setTimeout(resolve, 1500));
        }
      } else if (pageInfo.type === 'nps_comments') {
        // コメントページ - 各カラムを取得して各ページを生成
        // keyに応じてcommentTypeを設定
        if (pageInfo.key === 'I') {
          currentCommentType = 'feedback_i';
          showCommentSubNav('feedback_i');
        } else if (pageInfo.key === 'J') {
          currentCommentType = 'feedback_j';
          showCommentSubNav('feedback_j');
        } else if (pageInfo.key === 'M') {
          currentCommentType = 'feedback_m';
          showCommentSubNav('feedback_m');
        } else {
          currentCommentType = 'nps';
          showCommentSubNav('nps');
        }
        await fetchAndRenderCommentPage(pageInfo.key);
        // データ取得とレンダリングを待つ
        await new Promise(resolve => setTimeout(resolve, 1000));

        // 各カラム（ページ）を別々に出力
        if (currentCommentData && currentCommentData.length > 0) {
          // 最初のカラムは既に表示されているのでクローン
          const firstPage = await cloneCurrentPageForPrint();
          if (firstPage) {
            printContainer.appendChild(firstPage);
            console.log(`[PDF Export] ページ ${i + 1}-1 クローン完了`);
          }

          // 残りのカラムをループ
          for (let colIndex = 1; colIndex < currentCommentData.length; colIndex++) {
            renderCommentPage(colIndex);
            await new Promise(resolve => setTimeout(resolve, 1000));
            const colPage = await cloneCurrentPageForPrint();
            if (colPage) {
              printContainer.appendChild(colPage);
              console.log(`[PDF Export] ページ ${i + 1}-${colIndex + 1} クローン完了`);
            }
          }
          continue; // forループの次の反復へ（下のクローン処理をスキップ）
        }
      } else if (pageInfo.type === 'word_cloud') {
        // ワードクラウドページ - generateWordCloudPageForPrint を使用
        console.log(`[PDF Export] ワードクラウド: ${pageInfo.analysisKey}`);
        const wcPage = await generateWordCloudPageForPrint(pageInfo.analysisKey);
        if (wcPage) {
          printContainer.appendChild(wcPage);
          console.log(`[PDF Export] ページ ${i + 1} クローン完了`);
        }
        continue; // forループの次の反復へ
      } else if (pageInfo.type === 'ai_analysis') {
        // AI分析ページ - 実際の画面を表示してからクローン
        console.log(`[PDF Export] AI分析: ${pageInfo.analysisType} - ${pageInfo.subtype}`);
        await prepareAndShowAIAnalysisForPrint(pageInfo.analysisType, pageInfo.subtype);
        // AI分析データの取得とレンダリングを十分に待つ（レイアウトとフォント調整）
        await new Promise(resolve => setTimeout(resolve, 2000));
        // フォントサイズ調整のための追加待機
        await new Promise(resolve => requestAnimationFrame(() => {
          setTimeout(resolve, 500);
        }));
        const aiPage = await cloneAIAnalysisPageForPrint();
        if (aiPage) {
          printContainer.appendChild(aiPage);
          console.log(`[PDF Export] ページ ${i + 1} クローン完了`);
        }
        continue; // forループの次の反復へ
      }

      // レンダリング完了を待つ
      await new Promise(resolve => setTimeout(resolve, 500));

      // 現在表示されているコンテンツをクローン
      let printPage;
      if (pageInfo.type === 'ai_analysis') {
        // AI分析ページは screen5 からクローン
        printPage = await cloneAIAnalysisPageForPrint();
      } else {
        // 通常ページは screen3 からクローン
        printPage = await cloneCurrentPageForPrint();
      }

      if (printPage) {
        // 表紙(cover)の場合のみ、専用クラスを追加して識別できるようにする
        if (pageType === 'cover') {
          printPage.classList.add('cover-page');
          console.log(`[PDF Export] 表紙ページにcover-pageクラスを追加`);
        }

        printContainer.appendChild(printPage);
        console.log(`[PDF Export] ページ ${i + 1} (${pageType}) クローン完了`);
      } else {
        console.warn(`[PDF Export] ページ ${i + 1} のクローンに失敗`);
      }
    }

    // 元の画面を復元
    appFrame.style.opacity = '1';
    appFrame.style.pointerEvents = 'auto';

    console.log('[PDF Export] 全ページ生成完了');

    isPdfGenerating = false; // PDF生成完了
    showLoading(false);
    console.log('[PDF Export] ローディング非表示完了');

    // 印刷モードに切り替え
    document.body.classList.add('print-mode-active');
    console.log('[PDF Export] 印刷モード切り替え完了');

    // ポップアップ表示
    setTimeout(() => {
      document.getElementById('print-ready-popup').classList.remove('hidden');
      console.log('[PDF Export] ポップアップ表示完了');
    }, 200);

  } catch (error) {
    isPdfGenerating = false; // PDF生成エラー終了
    showLoading(false);
    console.error('[PDF Export] エラー発生:', error);
    console.error('[PDF Export] エラースタック:', error.stack);
    alert('PDF出力の準備中にエラーが発生しました: ' + error.message);
  }
}

// 現在表示中のページをクローン（screen3用）
async function cloneCurrentPageForPrint() {
  const screen3 = document.getElementById('screen3');
  if (!screen3) {
    console.error('[cloneCurrentPageForPrint] screen3が見つかりません');
    return null;
  }

  const reportBody = screen3.querySelector('.report-body');
  if (!reportBody) {
    console.error('[cloneCurrentPageForPrint] report-bodyが見つかりません');
    return null;
  }

  // 画像の読み込みを待つ
  const images = reportBody.querySelectorAll('img');
  if (images.length > 0) {
    console.log(`[cloneCurrentPageForPrint] ${images.length}個の画像の読み込みを待機中...`);
    await Promise.all(
      Array.from(images).map(img => {
        if (img.complete) {
          return Promise.resolve();
        }
        return new Promise((resolve) => {
          img.addEventListener('load', resolve);
          img.addEventListener('error', resolve);
          setTimeout(resolve, 5000);
        });
      })
    );
    console.log('[cloneCurrentPageForPrint] すべての画像の読み込み完了');
  }

  // 印刷用ページを作成
  const printPage = document.createElement('div');
  printPage.className = 'print-page';

  // report-body全体をクローン（タイトル、サブタイトル、セパレータ、ボディすべて含む）
  const bodyClone = reportBody.cloneNode(true);

  // Canvas要素の描画内容をコピー
  const originalCanvases = reportBody.querySelectorAll('canvas');
  const clonedCanvases = bodyClone.querySelectorAll('canvas');
  originalCanvases.forEach((originalCanvas, index) => {
    if (clonedCanvases[index]) {
      const clonedCanvas = clonedCanvases[index];
      const context = clonedCanvas.getContext('2d');
      // 元のcanvasのサイズを設定
      clonedCanvas.width = originalCanvas.width;
      clonedCanvas.height = originalCanvas.height;
      // 元のcanvasの描画内容をコピー
      context.drawImage(originalCanvas, 0, 0);
    }
  });

  // クローンしたボディのスタイルは元のまま維持（CSSのスタイルを使用）
  // padding: 40px 40px 20px 40px などは保持

  printPage.appendChild(bodyClone);

  return printPage;
}

// AI分析ページをクローン（screen5用）
async function cloneAIAnalysisPageForPrint() {
  const screen5 = document.getElementById('screen5');
  if (!screen5) {
    console.error('[cloneAIAnalysisPageForPrint] screen5が見つかりません');
    return null;
  }

  const reportBody = screen5.querySelector('.report-body');
  if (!reportBody) {
    console.error('[cloneAIAnalysisPageForPrint] report-bodyが見つかりません');
    return null;
  }

  // 印刷用ページを作成
  const printPage = document.createElement('div');
  printPage.className = 'print-page';

  // report-bodyをクローン
  const bodyClone = reportBody.cloneNode(true);

  // アクティブなパネルのみを表示するように調整
  const allPanels = bodyClone.querySelectorAll('.ai-panel');
  allPanels.forEach(panel => {
    if (!panel.classList.contains('is-active')) {
      // 非アクティブなパネルは削除
      panel.remove();
    } else {
      // アクティブなパネルは常に表示（印刷用にスタイルを強制）
      panel.style.position = 'relative';
      panel.style.opacity = '1';
      panel.style.visibility = 'visible';
      panel.style.transform = 'none';
      panel.style.display = 'block';
    }
  });

  // Canvas要素の描画内容をコピー
  const originalCanvases = reportBody.querySelectorAll('canvas');
  const clonedCanvases = bodyClone.querySelectorAll('canvas');
  originalCanvases.forEach((originalCanvas, index) => {
    if (clonedCanvases[index]) {
      const clonedCanvas = clonedCanvases[index];
      const context = clonedCanvas.getContext('2d');
      // 元のcanvasのサイズを設定
      clonedCanvas.width = originalCanvas.width;
      clonedCanvas.height = originalCanvas.height;
      // 元のcanvasの描画内容をコピー
      context.drawImage(originalCanvas, 0, 0);
    }
  });

  printPage.appendChild(bodyClone);

  return printPage;
}

// ワードクラウドページを直接生成（PDF出力専用）
async function generateWordCloudPageForPrint(columnType) {
  console.log(`[generateWordCloudPageForPrint] Generating: ${columnType}`);

  let analysisResults, title, subtitle;

  // キャッシュから取得を試みる
  const cached = wcAnalysisCache[columnType];
  if (cached && cached.analysisResults) {
    console.log(`[generateWordCloudPageForPrint] Using cache for: ${columnType}`);
    analysisResults = cached.analysisResults;
    title = cached.title;
    subtitle = cached.subtitle;
  } else {
    // キャッシュがない場合はAPIから取得
    console.log(`[generateWordCloudPageForPrint] No cache, fetching from API: ${columnType}`);

    // データ取得
    let tl = [], td = 0;
    try {
      const cd = await getReportDataForCurrentClinic(currentClinicForModal);
      switch (columnType) {
        case 'L':
          tl = cd.npsData.rawText || [];
          td = cd.npsData.totalCount || 0;
          break;
        case 'I':
          tl = cd.feedbackData.i_column.results || [];
          td = cd.feedbackData.i_column.totalCount || 0;
          break;
        case 'J':
          tl = cd.feedbackData.j_column.results || [];
          td = cd.feedbackData.j_column.totalCount || 0;
          break;
        case 'M':
          tl = cd.feedbackData.m_column.results || [];
          td = cd.feedbackData.m_column.totalCount || 0;
          break;
        default:
          console.error("Invalid column:", columnType);
          return null;
      }
    } catch (e) {
      console.error("Error accessing text data:", e);
      return null;
    }

    if (tl.length === 0) {
      console.warn("No text data for:", columnType);
      return null;
    }

    // 分析実行
    try {
      const r = await fetch('/api/analyzeText', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ textList: tl })
      });
      if (!r.ok) {
        throw new Error(`分析APIエラー(${r.status})`);
      }
      const ad = await r.json();
      analysisResults = ad.results;
      title = getAnalysisTitle(columnType, td);
      subtitle = '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。';
    } catch (error) {
      console.error('Analyze fail:', error);
      return null;
    }
  }

  // ユニークなIDを生成
  const uniqueId = `wc-${columnType}-${Date.now()}`;

  // ページ作成
  const printPage = document.createElement('div');
  printPage.className = 'print-page';

  const reportBody = document.createElement('div');
  reportBody.className = 'report-body';

  // タイトル
  const titleElem = document.createElement('h1');
  titleElem.className = 'report-title';
  titleElem.textContent = title;
  reportBody.appendChild(titleElem);

  // サブタイトル
  const subtitleElem = document.createElement('p');
  subtitleElem.className = 'report-subtitle';
  subtitleElem.textContent = subtitle;
  subtitleElem.style.textAlign = 'left';
  reportBody.appendChild(subtitleElem);

  // セパレータ
  const separator = document.createElement('hr');
  separator.className = 'report-separator';
  reportBody.appendChild(separator);

  // コンテンツエリア
  const reportContent = document.createElement('div');
  reportContent.className = 'report-content';

  const slideBody = document.createElement('div');
  slideBody.id = `slide-body-${uniqueId}`;
  slideBody.style.height = '100%';
  slideBody.style.overflowY = 'hidden';
  slideBody.innerHTML = `
    <div class="grid grid-cols-2 gap-2 h-full">
      <div class="grid grid-cols-2 grid-rows-2 gap-1 chart-wc-left" style="height: 80%;">
        <div id="noun-chart-container-${uniqueId}" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-blue-600 leading-none py-1" style="font-size: 12px;">名詞</h3>
          <div id="noun-chart-${uniqueId}" class="w-full flex-1"></div>
        </div>
        <div id="verb-chart-container-${uniqueId}" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-red-600 leading-none py-1" style="font-size: 12px;">動詞</h3>
          <div id="verb-chart-${uniqueId}" class="w-full flex-1"></div>
        </div>
        <div id="adj-chart-container-${uniqueId}" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-green-600 leading-none py-1" style="font-size: 12px;">形容詞</h3>
          <div id="adj-chart-${uniqueId}" class="w-full flex-1"></div>
        </div>
        <div id="int-chart-container-${uniqueId}" class="chart-container h-full">
          <h3 class="font-bold text-center mb-0 text-gray-600 leading-none py-1" style="font-size: 12px;">感動詞</h3>
          <div id="int-chart-${uniqueId}" class="w-full flex-1"></div>
        </div>
      </div>
      <div class="flex flex-col justify-start" style="height: 80%;">
        <p class="text-gray-600 text-left leading-tight" style="font-size: 12px; margin-bottom: 8px;">スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。<br>単語の色は品詞の種類で異なります。<br><span class="text-blue-600 font-semibold">青色=名詞</span>、<span class="text-red-600 font-semibold">赤色=動詞</span>、<span class="text-green-600 font-semibold">緑色=形容詞</span>、<span class="text-gray-600 font-semibold">灰色=感動詞</span></p>
        <div id="word-cloud-container-${uniqueId}" class="border border-gray-200" style="flex: 1; min-height: 0; display: flex; align-items: center; justify-content: center; padding: 0; background: #ffffff;">
          <canvas id="word-cloud-canvas-${uniqueId}" style="width: 100%; height: 100%; display: block;"></canvas>
        </div>
      </div>
    </div>
  `;

  reportContent.appendChild(slideBody);
  reportBody.appendChild(reportContent);
  printPage.appendChild(reportBody);

  // DOMに追加（グラフ描画のため）
  const tempContainer = document.createElement('div');
  tempContainer.style.position = 'fixed';
  tempContainer.style.left = '-9999px';
  tempContainer.style.width = '1122px';
  tempContainer.style.height = '793px';
  tempContainer.appendChild(printPage);
  document.body.appendChild(tempContainer);

  // グラフ描画を待つ
  await new Promise(resolve => {
    requestAnimationFrame(() => {
      drawAnalysisChartsTemp(analysisResults, uniqueId, resolve);
    });
  });

  // DOMから削除
  document.body.removeChild(tempContainer);

  return printPage;
}

// 一時的なグラフ描画関数
function drawAnalysisChartsTemp(results, uniqueId, onComplete) {
  if (!results || results.length === 0) {
    if (onComplete) onComplete();
    return;
  }

  const noun = [], verb = [], adj = [], intj = [];
  results.forEach(w => {
    if (w.pos === '名詞') noun.push([w.word, w.score]);
    else if (w.pos === '動詞') verb.push([w.word, w.score]);
    else if (w.pos === '形容詞') adj.push([w.word, w.score]);
    else if (w.pos === '感動詞') intj.push([w.word, w.score]);
  });

  // ▼▼▼ [追加] 太さを計算してオプションを返す関数 ▼▼▼
  const getBarOption = (count) => {
    const maxItems = 8;
    const baseWidthPercent = 60;
    // データ数に応じた％を計算
    return { groupWidth: (count / maxItems) * baseWidthPercent + '%' };
  };
  // ▲▲▲ 追加ここまで ▲▲▲

  const opt = {
    legend: 'none', backgroundColor: 'transparent', chartArea: { width: '95%', height: '90%' },
    hAxis: { textStyle: { fontSize: 9 } }, vAxis: { textStyle: { fontSize: 9 } }
  };

  let chartsDrawn = 0;
  const totalCharts = [noun, verb, adj, intj].filter(arr => arr.length > 0).length;

  const checkComplete = () => {
    chartsDrawn++;
    if (chartsDrawn >= totalCharts) {
      drawWordCloudOnCanvas(results, uniqueId, onComplete);
    }
  };

  // 棒グラフ描画
  if (noun.length > 0) {
    const d = google.visualization.arrayToDataTable([['単語', 'スコア'], ...noun.slice(0, 10)]);
    const chart = new google.visualization.BarChart(document.getElementById(`noun-chart-${uniqueId}`));
    google.visualization.events.addListener(chart, 'ready', checkComplete);
    chart.draw(d, { ...opt, colors: ['#2563eb'], bar: getBarOption(noun.slice(0, 10).length) });
  }
  if (verb.length > 0) {
    const d = google.visualization.arrayToDataTable([['単語', 'スコア'], ...verb.slice(0, 10)]);
    const chart = new google.visualization.BarChart(document.getElementById(`verb-chart-${uniqueId}`));
    google.visualization.events.addListener(chart, 'ready', checkComplete);
    chart.draw(d, { ...opt, colors: ['#dc2626'], bar: getBarOption(verb.slice(0, 10).length) });
  }
  if (adj.length > 0) {
    const d = google.visualization.arrayToDataTable([['単語', 'スコア'], ...adj.slice(0, 10)]);
    const chart = new google.visualization.BarChart(document.getElementById(`adj-chart-${uniqueId}`));
    google.visualization.events.addListener(chart, 'ready', checkComplete);
    chart.draw(d, { ...opt, colors: ['#16a34a'], bar: getBarOption(adj.slice(0, 10).length) });
  }
  if (intj.length > 0) {
    const d = google.visualization.arrayToDataTable([['単語', 'スコア'], ...intj.slice(0, 10)]);
    const chart = new google.visualization.BarChart(document.getElementById(`int-chart-${uniqueId}`));
    google.visualization.events.addListener(chart, 'ready', checkComplete);
    chart.draw(d, { ...opt, colors: ['#6b7280'], bar: getBarOption(intj.slice(0, 10).length) });
  }

  if (totalCharts === 0) {
    drawWordCloudOnCanvas(results, uniqueId, onComplete);
  }
}

function drawWordCloudOnCanvas(results, uniqueId, onComplete) {
  const canvas = document.getElementById(`word-cloud-canvas-${uniqueId}`);
  if (!canvas) {
    if (onComplete) onComplete();
    return;
  }

  const container = document.getElementById(`word-cloud-container-${uniqueId}`);
  const rect = container.getBoundingClientRect();

  const dpr = window.devicePixelRatio || 1;
  const logicalWidth = rect.width;
  const logicalHeight = rect.height;

  canvas.width = logicalWidth * dpr;
  canvas.height = logicalHeight * dpr;
  canvas.style.width = logicalWidth + 'px';
  canvas.style.height = logicalHeight + 'px';

  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  if (!results || results.length === 0) {
    if (onComplete) onComplete();
    return;
  }

  const posMap = {};
  results.forEach(w => { posMap[w.word] = w.pos; });

  const scores = results.map(r => r.score);
  const minScore = Math.min(...scores);
  const maxScore = Math.max(...scores);
  const scoreRange = maxScore - minScore || 1;

  const wordList = results.slice(0, 100).map(r => [r.word, r.score]);

  const dataCount = wordList.length;
  let baseMaxSize = 60;
  let baseMinSize = 14;

  if (dataCount <= 10) {
    baseMaxSize = 100;
    baseMinSize = 35;
  } else if (dataCount <= 20) {
    baseMaxSize = 80;
    baseMinSize = 28;
  } else if (dataCount <= 30) {
    baseMaxSize = 70;
    baseMinSize = 24;
  } else if (dataCount <= 40) {
    baseMaxSize = 65;
    baseMinSize = 20;
  } else if (dataCount <= 50) {
    baseMaxSize = 60;
    baseMinSize = 18;
  } else if (dataCount <= 70) {
    baseMaxSize = 55;
    baseMinSize = 16;
  } else if (dataCount <= 90) {
    baseMaxSize = 50;
    baseMinSize = 14;
  }

  const minDimension = Math.min(logicalWidth, logicalHeight);
  const sizeScale = Math.min(logicalWidth, logicalHeight) / 500;
  const scaledMaxSize = baseMaxSize * sizeScale;
  const scaledMinSize = baseMinSize * sizeScale;

  try {
    const options = {
      list: wordList,
      gridSize: Math.max(1, Math.round(minDimension / 200)),
      weightFactor: (score) => {
        const normalizedScore = (score - minScore) / scoreRange;
        return scaledMinSize + (scaledMaxSize - scaledMinSize) * normalizedScore;
      },
      fontFamily: 'Noto Sans JP,sans-serif',
      color: (word) => {
        const pos = posMap[word] || '不明';
        switch (pos) {
          case '名詞': return '#3b82f6';
          case '動詞': return '#ef4444';
          case '形容詞': return '#22c55e';
          case '感動詞': return '#6b7280';
          default: return '#a8a29e';
        }
      },
      backgroundColor: 'transparent',
      clearCanvas: true,
      rotateRatio: 0,
      drawOutOfBound: false,
      shrinkToFit: true,
      minSize: scaledMinSize * 0.3,
      shuffle: true,
      wait: 5,
      abortThreshold: 2000,
      abort: () => false,
      origin: [logicalWidth / 2, logicalHeight / 2]
    };

    if (typeof WordCloud !== 'undefined') {
      canvas.addEventListener('wordcloudstop', function handleStop() {
        canvas.removeEventListener('wordcloudstop', handleStop);
        setTimeout(() => {
          if (onComplete) onComplete();
        }, 100);
      });
      WordCloud(canvas, options);
    } else {
      console.error('WordCloud library not loaded');
      if (onComplete) onComplete();
    }
  } catch (error) {
    console.error('Error drawing wordcloud:', error);
    if (onComplete) onComplete();
  }
}

// AI分析ページを直接生成（PDF出力専用）
async function generateAIAnalysisPageForPrint(analysisType, tabId) {
  console.log(`[generateAIAnalysisPageForPrint] Generating: ${analysisType} - ${tabId}`);

  let content = '';
  let pentagonText = getTabLabel(tabId);

  // キャッシュから取得を試みる
  const cached = aiAnalysisCache[analysisType];
  if (cached && cached[tabId]) {
    console.log(`[generateAIAnalysisPageForPrint] Using cache for: ${analysisType} - ${tabId}`);
    content = cached[tabId].content || '';
    pentagonText = cached[tabId].pentagonText || getTabLabel(tabId);
  } else {
    // キャッシュがない場合はAPIから取得
    console.log(`[generateAIAnalysisPageForPrint] No cache, fetching from API: ${analysisType} - ${tabId}`);

    try {
      const response = await fetch('/api/getSingleAnalysisCell', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          centralSheetId: currentCentralSheetId,
          clinicName: currentClinicForModal,
          columnType: analysisType,
          tabId: tabId
        })
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
      }

      const data = await response.json();
      content = data.content || '';
      pentagonText = data.pentagonText || getTabLabel(tabId);
    } catch (error) {
      console.error('[generateAIAnalysisPageForPrint] Error:', error);
      content = 'データの取得に失敗しました';
    }
  }

  // ページ作成
  const printPage = document.createElement('div');
  printPage.className = 'print-page';

  const reportBody = document.createElement('div');
  reportBody.className = 'report-body';

  // タイトル
  const title = document.createElement('h1');
  title.className = 'report-title';
  title.textContent = getDetailedAnalysisTitleFull(analysisType);
  reportBody.appendChild(title);

  // サブタイトル
  const subtitle = document.createElement('p');
  subtitle.className = 'report-subtitle';
  subtitle.textContent = getSubtitleForItem(analysisType);
  reportBody.appendChild(subtitle);

  // セパレータ
  const separator = document.createElement('hr');
  separator.className = 'report-separator';
  reportBody.appendChild(separator);

  // コンテンツエリア
  const reportContent = document.createElement('div');
  reportContent.className = 'report-content';

  const aiContainer = document.createElement('section');
  aiContainer.className = 'ai-analysis-container';

  // サイドバー（五角形）
  const sidebar = document.createElement('div');
  sidebar.className = 'ai-analysis-sidebar';
  const shape = document.createElement('div');
  shape.className = 'ai-analysis-shape';
  shape.textContent = pentagonText;
  sidebar.appendChild(shape);
  aiContainer.appendChild(sidebar);

  // コンテンツ
  const contentDiv = document.createElement('div');
  contentDiv.className = 'ai-analysis-content';
  contentDiv.textContent = content;
  contentDiv.style.fontSize = '12pt';
  contentDiv.style.lineHeight = '1.5';
  aiContainer.appendChild(contentDiv);

  reportContent.appendChild(aiContainer);
  reportBody.appendChild(reportContent);
  printPage.appendChild(reportBody);

  // DOMに一時追加してフォントサイズ調整
  const tempContainer = document.createElement('div');
  tempContainer.style.position = 'fixed';
  tempContainer.style.left = '-9999px';
  tempContainer.style.width = '1122px';
  tempContainer.style.height = '793px';
  tempContainer.appendChild(printPage);
  document.body.appendChild(tempContainer);

  // フォントサイズ調整
  await new Promise(resolve => {
    requestAnimationFrame(() => {
      setTimeout(() => {
        adjustFontSize(contentDiv);
        resolve();
      }, 100);
    });
  });

  // DOMから削除
  document.body.removeChild(tempContainer);

  return printPage;
}

// タブラベル取得
function getTabLabel(tabId) {
  switch (tabId) {
    case 'analysis': return '分析と考察';
    case 'suggestions': return '改善案';
    case 'overall': return '総評';
    default: return '';
  }
}

// --- 初期化 ---
(async () => {
  console.log('DOM Loaded (assumed).');
  populateDateSelectors();
  try {
    await googleChartsLoaded;
    console.log('Charts loaded.');
  } catch (err) {
    console.error('Chart load fail:', err);
    alert('グラフライブラリ読込失敗');
  }
  setupEventListeners();
  console.log('Listeners setup.');
})();

