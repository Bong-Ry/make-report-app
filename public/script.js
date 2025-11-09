// public/script.js (ファイル6)

// --- グローバル変数 ---
let selectedPeriod = {}; 
let currentClinicForModal = '';
// let slidesData = []; (廃止)
// let currentPage = 1; (廃止)
let currentAnalysisTarget = 'L'; 
let currentDetailedAnalysisType = 'L'; 
let isEditingDetailedAnalysis = false; 
let currentCentralSheetId = null; 
let currentPeriodText = ""; 
let currentAiCompletionStatus = {}; 
let overallDataCache = null; 
let clinicDataCache = null; 

// ▼▼▼ [新規] コメントページ用グローバル変数 ▼▼▼
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
const googleChartsLoaded = new Promise(resolve => { google.charts.load('current', {'packages':['corechart', 'bar']}); google.charts.setOnLoadCallback(resolve); });


// --- ▼▼▼ [新規] フォントサイズ自動調整 (ルール ⑥, ④) ▼▼▼ ---
/**
 * [新規] コンテナの高さにテキストが収まるようフォントサイズを調整する (ルール ⑥, ④)
 * @param {HTMLElement} container - .ai-analysis-content または #slide-body (コメント時)
 * @param {number} initialSize - ルール ⑦ (基本 12pt, コメント 9pt)
 */
function adjustFontSize(container, initialSize = 12) {
    if (!container) return;
    
    // (ルール ⑦) 基本フォントサイズを設定
    container.style.fontSize = `${initialSize}pt`;
    container.style.lineHeight = initialSize <= 9 ? '1.4' : '1.5';
    
    // スクロールNG (ルール ④) のため、はみ出ていたらフォントを小さくする
    let currentSize = initialSize;
    
    // HACK: わずかな時間待機してレンダリング後の高さを取得
    setTimeout(() => {
         // (20回まで試行)
        for (let i = 0; i < 20; i++) {
            // (clientHeight よりも scrollHeight が大きい ＝ はみ出ている)
            const isOverflowing = container.scrollHeight > container.clientHeight;
            
            if (isOverflowing && currentSize > 6) { // 6pt未満にはしない
                currentSize -= 0.5; // 0.5pt ずつ小さくする
                container.style.fontSize = `${currentSize}pt`;
                container.style.lineHeight = '1.4'; // 少し詰める
            } else {
                break; // 収まったか、最小サイズになった
            }
        }
    }, 50); // 50ms待ってから高さをチェック
}
// --- ▲▲▲ ---


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
  document.getElementById('pdf-export-btn').addEventListener('click', generatePdf); 
  document.getElementById('back-to-clinics').addEventListener('click', () => showScreen('screen2'));
  
  // Screen 3 (Header - コメントUI)
  document.getElementById('slide-header').addEventListener('click', (e) => {
      // (サブナビ NPS 10, 9...)
      const subNavBtn = e.target.closest('.comment-sub-nav-btn');
      if (subNavBtn) {
          e.preventDefault();
          const key = subNavBtn.dataset.key;
          // 他のボタンのアクティブを解除
          document.querySelectorAll('.comment-sub-nav-btn').forEach(b => b.classList.remove('btn-active'));
          subNavBtn.classList.add('btn-active');
          // データ読み込み
          fetchAndRenderCommentPage(key);
          return;
      }
      
      // (ヘッダー内 ページネーション)
      const prevBtn = e.target.closest('#comment-prev');
      if (prevBtn) {
          e.preventDefault();
          if (currentCommentPageIndex > 0) {
              renderCommentPage(currentCommentPageIndex - 1);
          }
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

      // (ヘッダー内 保存ボタン)
      // ▼▼▼ [修正] 保存ボタンのリスナーを削除 ▼▼▼
      /*
      const saveBtn = e.target.closest('#comment-save-btn');
      if (saveBtn) {
          e.preventDefault();
          handleSaveComment();
          return;
      }
      */
  });

  // Screen 5 (Nav)
  document.getElementById('report-nav-screen5').addEventListener('click', handleReportNavClick);
  document.getElementById('pdf-export-btn-screen5').addEventListener('click', generatePdf);
  document.getElementById('back-to-clinics-from-detailed-analysis').addEventListener('click', () => {
      toggleEditDetailedAnalysis(false); 
      showScreen('screen2');
  });

  // Screen 5 (AI Tabs)
  document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => { button.addEventListener('click', handleTabClick); });
  
  // Screen 5 (AI Controls)
  document.getElementById('regenerate-detailed-analysis-btn').addEventListener('click', handleRegenerateDetailedAnalysis);
  document.getElementById('edit-detailed-analysis-btn').addEventListener('click', () => toggleEditDetailedAnalysis(true));
  document.getElementById('save-detailed-analysis-btn').addEventListener('click', saveDetailedAnalysisEdits);
  document.getElementById('cancel-edit-detailed-analysis-btn').addEventListener('click', () => {
      if (confirm('編集内容を破棄しますか？')) {
          toggleEditDetailedAnalysis(false); 
      }
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
  
  for(let i = 0; i < 5; i++) {
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
  
  for(let i = 1; i <= 12; i++) {
      const m = String(i).padStart(2, '0');
      sm.add(new Option(`${i}月`, m));
      em.add(new Option(`${i}月`, m));
  }
  
  em.value = String(now.getMonth() + 1).padStart(2, '0');
  sy.value = String(cy); 
  sm.value = String(now.getMonth() + 1).padStart(2, '0');
}

async function handleNextToClinics() {
  const sy=document.getElementById('start-year').value;
  const sm=document.getElementById('start-month').value;
  const ey=document.getElementById('end-year').value;
  const em=document.getElementById('end-month').value;
  const sd=new Date(`${sy}-${sm}-01`);
  const ed=new Date(`${ey}-${em}-01`);

  if(sd > ed){
      alert('開始年月<=終了年月で設定');
      return;
  }

  selectedPeriod = {start: `${sy}-${sm}`, end: `${ey}-${em}`};
  currentPeriodText = `${sy}-${sm}～${ey}-${em}`; 
  const displayPeriod = `${sy}年${sm}月～${ey}年${em}月`;

  showLoading(true, '集計シートを検索・準備中...');

  try {
      const response = await fetch('/api/findOrCreateSheet', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
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
              headers: {'Content-Type': 'application/json'},
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

  } catch(err) {
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
  const sc = Array.from(document.querySelectorAll('#clinic-list-container input:checked'))
                .map(cb => cb.value);
  if (sc.length === 0) { alert('転記対象のクリニックが選択されていません。'); return; }
  if (sc.length > 10) { alert('一度に選択できる件数は10件までです。チェックを減らしてください。'); return; }

  showLoading(true, '集計スプレッドシートへデータ転記中...\n完了後、バックグラウンドで分析タブが自動生成されます。');

  try {
      const r = await fetch('/api/getReportData', {
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({
              period: selectedPeriod,
              selectedClinics: sc,
              centralSheetId: currentCentralSheetId 
          })
      });
      if (!r.ok) { const et = await r.text(); throw new Error(`サーバーエラー(${r.status}): ${et}`); }
      const data = await r.json(); 
      loadClinics(); 
  } catch(err) {
      console.error('!!! ETL Issue failed:', err);
      alert(`データ転記失敗\n${err.message}`);
  } finally {
      showLoading(false);
  }
}
// --- 画面1/2 処理 終わり ---


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
  
  // [修正] コメント系が押されたら、グローバル変数をセット
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
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ centralSheetId: currentCentralSheetId, sheetName: sheetName })
  });
  if (!response.ok) throw new Error(`データ取得失敗(${sheetName}): ${await response.text()}`);
  const data = await response.json();
  
  if (isOverall) overallDataCache = data;
  else clinicDataCache = data;
  
  return data;
}


// --- ▼▼▼ [大幅修正] レポート表示メイン (Screen 3) (新ロジック) ▼▼▼ ---
async function prepareAndShowReport(reportType){
  console.log(`Prepare report: ${reportType}`); 
  showLoading(true,'レポートデータ集計中...');
  
  showScreen('screen3');
  updateNavActiveState(reportType, null, null);
  
  // UI初期化
  document.getElementById('report-title').textContent = '';
  document.getElementById('report-subtitle').textContent = '';
  document.getElementById('report-title').style.textAlign = 'left';
  // ▼▼▼ [修正] サブタイトルのtextAlign指定を削除 (CSS側で left !important に任せる) ▼▼▼
  // document.getElementById('report-subtitle').style.textAlign = 'left'; 
  
  // ▼▼▼ [修正] (Req ⑤) グラフ・AI以外のセパレーターマージンをリセット ▼▼▼
  document.getElementById('report-separator').style.display = 'block'; 
  document.getElementById('report-separator').style.marginBottom = '16px';
  
  // ▼▼▼ [修正] slide-header の中身(innerHTML)を直接クリアするのをやめ、子要素を個別にクリアする ▼▼▼
  const subNav = document.getElementById('comment-sub-nav');
  const controls = document.getElementById('comment-controls');
  if (subNav) subNav.innerHTML = ''; // コメントサブナビをクリア
  if (controls) controls.innerHTML = ''; // コメントコントロールをクリア
  // ▲▲▲
  
  const slideBody = document.getElementById('slide-body');
  slideBody.innerHTML='';
  slideBody.style.overflowY = 'hidden'; // デフォルトはスクロールなし
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  showCopyrightFooter(reportType !== 'cover'); // フッター表示

  // --- 1. コメント系レポートの場合 (新APIを呼ぶ) ---
  if (currentCommentType) {
      try {
          // NPSか、それ以外(I,J,M)かで初期表示キーを決定
          const initialKey = currentCommentType === 'nps' ? 'L_10' : (currentCommentType === 'feedback_i' ? 'I' : (currentCommentType === 'feedback_j' ? 'J' : 'M'));
          
          // サブナビゲーションを描画 (NPSの場合は 10点, 9点...)
          showCommentSubNav(currentCommentType);
          
          // データを取得して最初のページ（列）を描画
          await fetchAndRenderCommentPage(initialKey);

      } catch(e) {
          console.error('Comment data fetch error:', e);
          document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">コメントデータ取得失敗<br>(${e.message})</p>`;
      } finally {
          showLoading(false);
      }
      return; // コメント系はここで終了
  }

  // --- 2. グラフ・概要系レポート (既存のロジック) ---
  
  // 例外構成 (表紙, 目次, 概要)
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

  // 基本構成 (グラフ)
  let isChart=false;
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
  
  // ▼▼▼ [修正] ページタイトル設定とグラフ描画準備 (サブタイトルを空文字('')またはNPS専用サブタイトルに) ▼▼▼
  if (reportType === 'nps_score'){ prepareChartPage('アンケート結果S(ネットプロモータースコア)＝推奨度ー', 'これから初めてお産を迎える友人知人がいた場合、\nご出産された産婦人科医院をどのくらいお勧めしたいですか。\n友人知人への推奨度を教えてください。＜推奨度＞ 10:強くお勧めする〜 0:全くお勧めしない', 'nps_score'); isChart=true; }
  else if (reportType === 'satisfaction_b'){ prepareChartPage('アンケート結果　ー満足度ー','', 'satisfaction_b'); isChart=true; }
  else if (reportType === 'satisfaction_c'){ prepareChartPage('アンケート結果　ー施設の充実度・快適さー','', 'satisfaction_c'); isChart=true; }
  else if (reportType === 'satisfaction_d'){ prepareChartPage('アンケート結果　ーアクセスの良さー','', 'satisfaction_d'); isChart=true; }
  else if (reportType === 'satisfaction_e'){ prepareChartPage('アンケート結果　ー費用ー','', 'satisfaction_e'); isChart=true; }
  else if (reportType === 'satisfaction_f'){ prepareChartPage('アンケート結果　ー病院の雰囲気ー','', 'satisfaction_f'); isChart=true; }
  else if (reportType === 'satisfaction_g'){ prepareChartPage('アンケート結果　ースタッフの対応ー','', 'satisfaction_g'); isChart=true; }
  else if (reportType === 'satisfaction_h'){ prepareChartPage('アンケート結果　ー先生の診断・説明ー','', 'satisfaction_h'); isChart=true; }
  else if (reportType === 'age'){ prepareChartPage('アンケート結果　ーご回答者さまの年代ー','', 'age'); isChart=true; }
  else if (reportType === 'children'){ prepareChartPage('アンケート結果　ーご回答者さまのお子様の人数ー','', 'children'); isChart=true; }
  else if (reportType === 'income'){ prepareChartPage('アンケート結果　ーご回答者さまの世帯年収ー','', 'income', true); isChart=true; }

  
  if (isChart) { 
      setTimeout(()=>{
          try{
              if(reportType==='nps_score'){
                  drawNpsScoreCharts(clinicData.npsScoreData, overallData.npsScoreData);
              }
              else if(reportType.startsWith('satisfaction')){
                  const type = reportType.split('_')[1] + '_column';
                  drawSatisfactionCharts(clinicData.satisfactionData[type].results, overallData.satisfactionData[type].results);
              }
              else if(reportType==='age'){
                  drawSatisfactionCharts(clinicData.ageData.results, overallData.ageData.results);
              }
              else if(reportType==='children'){
                  drawSatisfactionCharts(clinicData.childrenCountData.results, overallData.childrenCountData.results);
              }
              else if(reportType==='income'){
                  drawIncomeCharts(clinicData.incomeData, overallData.incomeData);
              }
          }catch(e){
              console.error('Chart draw error:',e);
              document.getElementById('slide-body').innerHTML=`<p class="text-center text-red-500 py-16">グラフ描画失敗<br>(${e.message})</p>`;
          }finally{
              showLoading(false);
          }
      },100);
  } else {
      showLoading(false);
  } 
}

// (変更なし) 例外構成（表紙・目次・概要）の表示
async function prepareAndShowIntroPages(reportType) {
  document.getElementById('report-separator').style.display='none'; 
  // ▼▼▼ [修正] サブタイトルのtextAlign指定を削除 (CSS側で left !important に任せる) ▼▼▼
  // document.getElementById('report-subtitle').style.textAlign = 'center'; 
  document.getElementById('slide-body').style.whiteSpace = 'pre-wrap';
  document.getElementById('slide-body').classList.remove('flex', 'items-center', 'justify-center', 'text-center');

  if (reportType === 'cover') {
      document.getElementById('report-title').textContent = currentClinicForModal;
      document.getElementById('report-title').style.textAlign = 'center'; 
      document.getElementById('slide-body').innerHTML = '<div class="flex items-center justify-center h-full"><h2 class="text-4xl font-bold">アンケートレポート</h2></div>';
      document.getElementById('report-subtitle').textContent = '';
  } else if (reportType === 'toc') {
      document.getElementById('report-title').textContent = '目次';
      document.getElementById('report-title').style.textAlign = 'left';
      document.getElementById('report-subtitle').textContent = '';
      document.getElementById('slide-body').innerHTML = `
          <div class="flex justify-center h-full items-start pt-8">
              <ul class="text-2xl font-semibold space-y-4 text-left">
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
          const overallData = await getReportDataForCurrentClinic("全体");
          overallCount = overallData.npsScoreData.totalCount || 0;
          
          // ▼▼▼ [修正] 転記済みリストの件数取得ロジックを修正 (currentAiCompletionStatus を利用) ▼▼▼
          clinicListCount = Object.keys(currentAiCompletionStatus).length; 
          
          const clinicData = await getReportDataForCurrentClinic(currentClinicForModal);
          clinicCount = clinicData.npsScoreData.totalCount || 0;
      } catch (e) { console.warn("Error fetching data for summary:", e); }

      const [sy, sm] = selectedPeriod.start.split('-').map(Number);
      const [ey, em] = selectedPeriod.end.split('-').map(Number);
      const startDay = new Date(sy, sm - 1, 1).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
      const endDay = new Date(ey, em, 0).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
      
      document.getElementById('report-title').textContent = 'アンケート概要';
      document.getElementById('report-title').style.textAlign = 'left';
      document.getElementById('report-subtitle').textContent = '';
      document.getElementById('slide-body').innerHTML = `
          <div class="flex justify-center h-full items-start pt-8">
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

// ▼▼▼ [修正] グラフ描画用シェル設定 (サブタイトル制御 + 余白変更) ▼▼▼
function prepareChartPage(title, subtitle, type, isBar=false) { 
  document.getElementById('report-title').textContent = title;
  document.getElementById('report-subtitle').textContent = subtitle; // (NPS以外は空文字が渡される)
  // ▼▼▼ [修正] サブタイトルのtextAlign指定を削除 (CSS側で left !important に任せる) ▼▼▼
  // document.getElementById('report-subtitle').style.textAlign = 'center'; 
  
  // ▼▼▼ [変更] (Req ⑤) グラフページではセパレーターマージンを 24px に変更 ▼▼▼
  document.getElementById('report-separator').style.display='block';
  document.getElementById('report-separator').style.marginBottom = '24px'; // 16px -> 24px

  let htmlContent = '';
  const cid = isBar ? 'bar-chart' : 'pie-chart';
  const chartHeightClass = 'h-[320px]'; // A4枠内に収まるように高さを固定

  if (type === 'nps_score') {
      // ▼▼▼ [修正] (Req ⑤) NPSスコアのセパレーターマージンは 16px のままにする ▼▼▼
      document.getElementById('report-separator').style.marginBottom = '16px';
      htmlContent = `
          <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-full">
              <div class="flex flex-col h-full">
                  <h3 id="clinic-chart-header" class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
                  <div id="clinic-bar-chart" class="w-full ${chartHeightClass} border border-gray-200 flex items-center justify-center"></div>
                  <div class="w-full h-[150px] flex flex-col justify-center items-center mt-2"> <p class="text-sm text-gray-500 mb-2">【画像入力エリア】</p>
                      <div class="w-full h-full border border-dashed border-gray-300 flex items-center justify-center text-gray-400">
                          [画像を入力する]
                      </div>
                  </div>
              </div>
              <div id="nps-summary-area" class="flex flex-col justify-center items-center space-y-6 pt-12 h-full">
                   <p class="text-gray-500">NPSスコア計算中...</p>
              </div>
          </div>
      `;
  } else {
      htmlContent = `
          <div class="grid grid-cols-1 md:grid-cols-2 gap-8 h-full">
              <div class="flex flex-col items-center h-full">
                  <h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
                  <div id="clinic-${cid}" class="w-full ${chartHeightClass} clinic-graph-bg-yellow"></div>
              </div>
              <div class="flex flex-col items-center h-full">
                  <h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3>
                  <div id="average-${cid}" class="w-full ${chartHeightClass}"></div>
              </div>
          </div>
      `;
  }
  
  document.getElementById('slide-body').innerHTML = htmlContent;
}

// --- ▼▼▼ [修正] グラフ描画関数 (フォントサイズ 14pt, %を黒文字に) ▼▼▼
function drawSatisfactionCharts(clinicChartData, overallChartData){ 
    // ▼▼▼ [修正] fontSize: 14, pieSliceTextStyle.color: 'black' (ご要望) ▼▼▼
    const opt={is3D:true,chartArea:{left:'5%',top:'5%',width:'90%',height:'90%'},pieSliceText:'percentage',pieSliceTextStyle:{color:'black',fontSize:14,bold:true},legend:{position:'labeled',textStyle:{color:'black',fontSize:14}},tooltip:{showColorCode:true,textStyle:{fontSize:14},trigger:'focus'},colors:['#4285F4','#DB4437','#F4B400','#0F9D58','#990099'], backgroundColor: 'transparent'};
    const cdEl=document.getElementById('clinic-pie-chart');if (!cdEl) throw new Error('グラフ描画エリア(clinic-pie-chart)が見つかりません。');if(clinicChartData&&clinicChartData.length>1&&clinicChartData.slice(1).some(row=>row[1]>0)){const d=google.visualization.arrayToDataTable(clinicChartData);const c=new google.visualization.PieChart(cdEl);c.draw(d,opt);} else {cdEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const avgEl=document.getElementById('average-pie-chart');if (!avgEl) throw new Error('グラフ描画エリア(average-pie-chart)が見つかりません。');if(overallChartData&&overallChartData.length>1&&overallChartData.slice(1).some(row=>row[1]>0)){const avgD=google.visualization.arrayToDataTable(overallChartData);const avgC=new google.visualization.PieChart(avgEl);avgC.draw(avgD,opt);} else {avgEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} }
function drawIncomeCharts(clinicData, overallData){ 
    // ▼▼▼ [修正] fontSize: 14, annotations.textStyle.color: 'black' (ご要望) ▼▼▼
    const opt={legend:{position:'none'},colors:['#DE5D83'],annotations:{textStyle:{fontSize:14,color:'black',auraColor:'none'},alwaysOutside:false,stem:{color:'transparent'}},vAxis:{format:"#.##'%'",viewWindow:{min:0}, textStyle:{fontSize:14}, titleTextStyle:{fontSize:14}}, hAxis:{textStyle:{fontSize:14}, titleTextStyle:{fontSize:14}}, backgroundColor: 'transparent'};
    const ccdEl=document.getElementById('clinic-bar-chart');if (!ccdEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');if(clinicData.totalCount > 0 && clinicData.results && clinicData.results.length > 1){const cd=google.visualization.arrayToDataTable(clinicData.results);const cc=new google.visualization.ColumnChart(ccdEl);cc.draw(cd,opt);} else {ccdEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const avgEl=document.getElementById('average-bar-chart');if (!avgEl) throw new Error('グラフ描画エリア(average-bar-chart)が見つかりません。');if(overallData.totalCount > 0 && overallData.results && overallData.results.length > 1){const avgD=google.visualization.arrayToDataTable(overallData.results);const avgC=new google.visualization.ColumnChart(avgEl);avgC.draw(avgD,opt); } else {avgEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} }
function drawNpsScoreCharts(clinicData, overallData) { const clinicChartEl = document.getElementById('clinic-bar-chart');if (!clinicChartEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');const clinicNpsScore = calculateNps(clinicData.counts, clinicData.totalCount);const overallNpsScore = calculateNps(overallData.counts, overallData.totalCount);const clinicChartData = [['スコア', '割合', { role: 'annotation' }]];if (clinicData.totalCount > 0) {for (let i = 0; i <= 10; i++) {const count = clinicData.counts[i] || 0;const percentage = (count / clinicData.totalCount) * 100;clinicChartData.push([String(i), percentage, `${Math.round(percentage)}%`]);}} 
    // ▼▼▼ [修正] fontSize: 14, annotations.textStyle.color: 'black' (ご要望) ▼▼▼
    const opt = {legend: { position: 'none' },colors: ['#DE5D83'], annotations: {textStyle: { fontSize: 14, color: 'black', auraColor: 'none' },alwaysOutside: false,stem: { color: 'transparent' }},vAxis: { format: "#.##'%'", title: '割合(%)', viewWindow: { min: 0 }, textStyle:{fontSize:14}, titleTextStyle:{fontSize:14}},hAxis: { title: '推奨度スコア (0〜10)', textStyle:{fontSize:14}, titleTextStyle:{fontSize:14}},bar: { groupWidth: '80%' },isStacked: false, chartArea:{height:'75%', width:'90%', left:'5%', top:'5%'}, backgroundColor: 'transparent'};
    if (clinicData.totalCount > 0 && clinicChartData.length > 1) {const clinicDataVis = google.visualization.arrayToDataTable(clinicChartData);const clinicChart = new google.visualization.ColumnChart(clinicChartEl);clinicChart.draw(clinicDataVis, opt);} else {clinicChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const summaryArea = document.getElementById('nps-summary-area');if (summaryArea) {summaryArea.innerHTML = ` <div class="text-left text-3xl space-y-5 p-6 border rounded-lg bg-gray-50 shadow-inner w-full max-w-xs"> <p>全体：<span class="font-bold text-gray-800">${overallNpsScore.toFixed(1)}</span></p> <p>貴院：<span class="font-bold text-red-600">${clinicNpsScore.toFixed(1)}</span></p> </div> `;} const clinicHeaderEl = document.getElementById('clinic-chart-header');if (clinicHeaderEl) {clinicHeaderEl.textContent = `貴院の結果 (全 ${clinicData.totalCount} 件)`;} }
function calculateNps(counts, totalCount) { if (totalCount === 0) return 0;let promoters = 0, passives = 0, detractors = 0;for (let i = 0; i <= 10; i++) {const count = counts[i] || 0;if (i >= 9) promoters += count;else if (i >= 7) passives += count;else detractors += count;} return ((promoters / totalCount) - (detractors / totalCount)) * 100; }


// --- ▼▼▼ [修正] コメントスライド構築関数群 (非編集化) ▼▼▼ ---

/**
 * [新規] JS側でコメントシート名を取得するヘルパー
 */
function getCommentSheetName(clinicName, type) {
  switch(type) {
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

/**
 * [新規] 0 -> A, 1 -> B ... のExcel列名変換
 */
function getExcelColumnName(colIndex) {
    let colName = '';
    let dividend = colIndex + 1;
    while (dividend > 0) {
        let modulo = (dividend - 1) % 26;
        colName = String.fromCharCode(65 + modulo) + colName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return colName;
}

/**
 * [修正] コメントのサブナビゲーション (サブタイトル廃止)
 */
function showCommentSubNav(reportType) {
    const navContainer = document.getElementById('comment-sub-nav');
    if (!navContainer) {
        console.error('comment-sub-nav element not found');
        return;
    }
    navContainer.innerHTML = '';
    
    let title = '';
    // ▼▼▼ [修正] サブタイトルを廃止 (空に) ▼▼▼
    let subTitle = ''; 
    
    if (reportType === 'nps') {
        title = 'アンケート結果　ーNPS推奨度 理由ー';
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
            // 最初のボタン(10点)をアクティブにする
            if (index === 0) btn.classList.add('btn-active');
            btn.dataset.key = group.key;
            btn.textContent = group.label;
            navContainer.appendChild(btn);
        });
    } else {
        const titleMap = { 'feedback_i': '良かった点や悪かった点など', 'feedback_j': '印象に残ったスタッフへのコメント', 'feedback_m': 'お産にかかわるご意見・ご感想' };
        title = `アンケート結果S${titleMap[reportType]}ー`;
    }
    
    document.getElementById('report-title').textContent = title;
    document.getElementById('report-subtitle').textContent = subTitle;
}

/**
 * [修正] APIからコメントシートのデータを取得し、最初のページを描画 (NPSサブタイトル・件数修正)
 */
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
    
    const controlsContainer = document.getElementById('comment-controls');
    if (controlsContainer) {
        controlsContainer.innerHTML = ''; // コントロールもクリア
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
        
        const data = await response.json(); // 例: [ ['A列c1', 'A列c2'], ['B列c1'] ]
        
        // ▼▼▼ [修正] (Req 3) 1ページ目の件数ではなく、全ページの合計件数を計算 ▼▼▼
        let totalCommentCount = 0;
        if (data && data.length > 0) {
            totalCommentCount = data.reduce((acc, column) => acc + column.length, 0);
        }
        // ▲▲▲

        if (totalCommentCount === 0) {
            currentCommentData = [];
            document.getElementById('slide-body').innerHTML = '<p class="text-center text-gray-500 py-16">コメントデータがありません</p>';
            renderCommentControls(); 
            // ▼▼▼ [修正] データ0件でもNPSサブタイトルを設定 ▼▼▼
            setNpsSubtitle(commentKey, 0); 
        } else {
            currentCommentData = data;
            // ▼▼▼ [修正] 描画の前にNPSサブタイトルを設定 (合計件数を渡す) ▼▼▼
            setNpsSubtitle(commentKey, totalCommentCount); 
            
            renderCommentPage(0); // 最初のページ (A列) を描画
        }
        
    } catch (e) {
        console.error('Comment data fetch error:', e);
        document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">コメントデータ取得失敗<br>(${e.message})</p>`;
    } finally {
        showLoading(false);
    }
}

/**
 * [修正] NPSコメント用のサブタイトルを設定する (ご要望の形式)
 * @param {string} commentKey (例: "L_10", "I")
 * @param {number} count (★グループの合計件数)
 */
function setNpsSubtitle(commentKey, count) {
    const subTitleEl = document.getElementById('report-subtitle');
    if (!subTitleEl) return;
    
    let subTitle = '';
    // (アイコンURLは未指定のため、プレースホルダ)
    if (commentKey === 'L_10' || commentKey === 'L_9') {
        const score = commentKey.split('_')[1];
        // ▼▼▼ [修正] count は既に合計件数 (Req 3) ▼▼▼
        subTitle = `[アイコン①] NPS ${score} ${count}人`; 
    } else if (commentKey === 'L_8' || commentKey === 'L_7') {
        const score = commentKey.split('_')[1];
        subTitle = `[アイコン②] NPS ${score} ${count}人`;
    } else if (commentKey === 'L_6_under') {
        subTitle = `[アイコン③] NPS 6以下 ${count}人`;
    } else {
        // "I", "J", "M" の場合はサブタイトルなし
        subTitle = '';
    }
    
    subTitleEl.textContent = subTitle;
}


/**
 * [修正] 指定されたページ(列)のコメントを <p> (編集不可) として描画
 */
function renderCommentPage(pageIndex) {
    if (!currentCommentData) return;
    
    currentCommentPageIndex = pageIndex;
    const columnData = currentCommentData[pageIndex] || []; // (例: ['A列c1', 'A列c2'])
    
    const bodyEl = document.getElementById('slide-body');
    bodyEl.innerHTML = '';
    
    // ▼▼▼ [修正] スクロールを 'hidden' に変更 (ルール ④) ▼▼▼
    bodyEl.style.overflowY = 'hidden'; 
    
    if (columnData.length === 0 && currentCommentData.length > 0) {
        bodyEl.innerHTML = '<p class="text-center text-gray-500 py-16">(このページは空です)</p>';
    } else {
        const fragment = document.createDocumentFragment();
        columnData.forEach((comment, index) => {
            // ▼▼▼ [修正] <textarea> の代わりに <p> を使用 ▼▼▼
            const p = document.createElement('p');
            p.className = 'comment-display-item';
            p.textContent = comment; // .textContent を使って安全にテキストを挿入
            
            fragment.appendChild(p);
        });
        bodyEl.appendChild(fragment);
    }
    
    // ヘッダーのコントロール（ページャー、保存ボタン）を再描画
    renderCommentControls();
    
    // ▼▼▼ [新規] フォントサイズ自動調整を実行 (ルール ④, ⑥) ▼▼▼
    // (CSSで 9pt が指定されているため、9pt から開始)
    adjustFontSize(bodyEl, 9);
}

/**
 * [修正] ヘッダーのコメントコントロール（保存ボタン削除）
 */
function renderCommentControls() {
    const controlsContainer = document.getElementById('comment-controls');
    if (!controlsContainer) {
        console.error('renderCommentControls: comment-controls element not found!');
        return;
    }
    controlsContainer.innerHTML = '';
    
    if (!currentCommentData) {
        return;
    }

    const totalPages = Math.max(1, currentCommentData.length); 
    const currentPage = currentCommentPageIndex + 1;
    const currentCol = getExcelColumnName(currentCommentPageIndex);
    
    const prevDisabled = currentCommentPageIndex === 0;
    const nextDisabled = currentCommentPageIndex >= totalPages - 1;
    const prevCol = prevDisabled ? '' : getExcelColumnName(currentCommentPageIndex - 1);
    const nextCol = nextDisabled ? '' : getExcelColumnName(currentCommentPageIndex + 1);

    // ▼▼▼ [修正] 保存ボタンを削除 ▼▼▼
    controlsContainer.innerHTML = `
        <button id="comment-prev" class="btn" ${prevDisabled ? 'disabled' : ''}>&lt; ${prevCol}</button>
        <span>${currentCol}列 (${currentPage} / ${totalPages})</span>
        <button id="comment-next" class="btn" ${nextDisabled ? 'disabled' : ''}>${nextCol} &gt;</button>
    `;
}

// ▼▼▼ [修正] handleSaveComment 関数を削除 ▼▼▼
// async function handleSaveComment() { ... } (削除)
// --- ▲▲▲ コメントスライド構築関数群 終わり ▲▲▲ ---


// ▼▼▼ [修正] 市区町村 (Req ④: サブタイトル廃止, ソート) (Req ⑤: 余白変更) ▼▼▼
async function prepareAndShowMunicipalityReport() {
  console.log('Prepare municipality report');
  updateNavActiveState('municipality', null, null);
  showScreen('screen3');
  document.getElementById('report-title').textContent = `アンケート結果　ーご回答者さまの市町村ー`;
  // ▼▼▼ [修正] サブタイトル廃止 ▼▼▼
  document.getElementById('report-subtitle').textContent = ''; 
  document.getElementById('report-separator').style.display='block';
  // ▼▼▼ [変更] (Req ⑤) グラフページではセパレーターマージンを 24px に変更 ▼▼▼
  document.getElementById('report-separator').style.marginBottom = '24px'; // 16px -> 24px
  
  const slideBody = document.getElementById('slide-body');
  slideBody.style.whiteSpace = 'normal';
  slideBody.innerHTML = '<p class="text-center text-gray-500 py-16">市区町村データを読み込み中...</p>';
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.style.overflowY = 'auto'; // ★ 市区町村テーブルはスクロール許可 (ルール④の例外)
  
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
      
      // ▼▼▼ [修正] (Req ④) 「不明」を最下部にソート ▼▼▼
      tableData.sort((a, b) => {
          if (a.prefecture === '不明') return 1;
          if (b.prefecture === '不明') return -1;
          return b.count - a.count; // 通常は件数で降順
      });
      
      displayMunicipalityTable(tableData);
  } catch (err) {
      console.error('Failed to get municipality report:', err);
      slideBody.innerHTML = `<p class="text-center text-red-500 py-16">集計失敗: ${err.message}</p>`;
  } finally {
      showLoading(false);
  }
}

// ▼▼▼ [修正] (Req ④) 市区町村テーブル描画 (スタイル適用) ▼▼▼
function displayMunicipalityTable(data) { 
  const slideBody = document.getElementById('slide-body'); 
  if (!data || data.length === 0) { 
      slideBody.innerHTML = '<p class="text-center text-gray-500 py-16">集計データがありません。</p>'; 
      return; 
  } 
  // ▼▼▼ [修正] (Req ④) コンテナに余白(px-8)、テーブルに 'municipality-table' クラスを追加 ▼▼▼
  let tableHtml = `<div class="municipality-table-container w-full h-full px-8"><table class="w-full municipality-table"><thead class="bg-gray-50 sticky top-0 z-10"><tr><th>都道府県</th><th>市区町村</th><th>件数</th><th>割合</th></tr></thead><tbody class="bg-white">`; 
  data.forEach(row => { 
      tableHtml += `<tr><td>${row.prefecture}</td><td>${row.municipality}</td><td class="text-right">${row.count}</td><td class="text-right">${row.percentage.toFixed(2)}%</td></tr>`; 
  }); 
  tableHtml += '</tbody></table></div>'; 
  slideBody.innerHTML = tableHtml; 
}


// ▼▼▼ [修正] おすすめ理由 (サブタイトル廃止) (Req ⑤: 余白変更) ▼▼▼
async function prepareAndShowRecommendationReport() {
    console.log('Prepare recommendation report');
    updateNavActiveState('recommendation', null, null);
    showScreen('screen3');
    document.getElementById('report-title').textContent = 'アンケート結果　ー本病院を選ぶ上で最も参考にしたものー';
    // ▼▼▼ [修正] サブタイトル廃止 ▼▼▼
    document.getElementById('report-subtitle').textContent = '';
    // ▼▼▼ [変更] (Req ⑤) グラフページではセパレーターマージンを 24px に変更 ▼▼▼
    document.getElementById('report-separator').style.display='block';
    document.getElementById('report-separator').style.marginBottom = '24px'; // 16px -> 24px
    
    const slideBody = document.getElementById('slide-body');
    slideBody.style.whiteSpace = 'normal';
    slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
    
    // ▼▼▼ [修正] 高さを h-[320px] に、左側に背景色クラスを追加 ▼▼▼
    slideBody.innerHTML = `<div class="grid grid-cols-1 md:grid-cols-2 gap-8 h-full"><div class="flex flex-col items-center"><h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3><div id="clinic-pie-chart" class="w-full h-[320px] clinic-graph-bg-yellow"></div></div><div class="flex flex-col items-center"><h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3><div id="average-pie-chart" class="w-full h-[320px]"></div></div></div>`;
    
    try {
        showLoading(true, '集計済みのおすすめ理由データを読込中...');
        const [clinicChartDataRes, overallChartDataRes] = await Promise.all([
              fetch('/api/getRecommendationReport', { 
                  method: 'POST',
                  headers: {'Content-Type': 'application/json'},
                  body: JSON.stringify({ 
                      centralSheetId: currentCentralSheetId, 
                      clinicName: currentClinicForModal 
                  })
              }),
              fetch('/api/getRecommendationReport', { 
                  method: 'POST',
                  headers: {'Content-Type': 'application/json'},
                  body: JSON.stringify({ 
                      centralSheetId: currentCentralSheetId, 
                      clinicName: "全体" 
                  })
              })
        ]);
        
        if (!clinicChartDataRes.ok) throw new Error(`貴院データ取得失敗: ${await clinicChartDataRes.text()}`);
        if (!overallChartDataRes.ok) throw new Error(`全体データ取得失敗: ${await overallChartDataRes.text()}`);

        const clinicChartData = await clinicChartDataRes.json(); 
        const overallChartData = await overallChartDataRes.json(); 
        
        showLoading(false);
        
        // ▼▼▼ [修正] フォントサイズを 14pt に, %を黒文字に (ルール②) ▼▼▼
        const opt = {is3D: true,chartArea: { left: '5%', top: '5%', width: '90%', height: '90%' },pieSliceText: 'percentage',pieSliceTextStyle: { color: 'black', fontSize: 14, bold: true },legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 14 } },tooltip: { showColorCode: true, textStyle: { fontSize: 14 }, trigger: 'focus' }, backgroundColor: 'transparent'};
        const clinicChartEl = document.getElementById('clinic-pie-chart');
        if (!clinicChartEl) throw new Error('グラフ描画エリア(clinic-pie-chart)が見つかりません。');
        const totalClinicCount = clinicChartData.slice(1).reduce((sum, row) => sum + row[1], 0);
        if (totalClinicCount > 0) {const d = google.visualization.arrayToDataTable(clinicChartData);new google.visualization.PieChart(clinicChartEl).draw(d, opt);} else {clinicChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';}
        const averageChartEl = document.getElementById('average-pie-chart');
        if (!averageChartEl) throw new Error('グラフ描画エリア(average-pie-chart)が見つかりません。');
        const totalOverallCount = overallChartData.slice(1).reduce((sum, row) => sum + row[1], 0);
         if (totalOverallCount > 0) {const d = google.visualization.arrayToDataTable(overallChartData);new google.visualization.PieChart(averageChartEl).draw(d, opt);} else {averageChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';}
    } catch (err) {
        console.error('Failed to get recommendation report:', err);
        slideBody.innerHTML = `<p class="text-center text-red-500">集計失敗: ${err.message}</p>`;
        showLoading(false);
    }
}


// ▼▼▼ [修正] Word Cloud表示 (Screen 3) (Req ⑤, ⑦.1, ⑦.2, ⑦.3) ▼▼▼
async function prepareAndShowAnalysis(columnType) {
  showLoading(true, `テキスト分析中(${getColumnName(columnType)})...`);
  showScreen('screen3');
  clearAnalysisCharts();
  updateNavActiveState(null, columnType, null);
  showCopyrightFooter(true); // WCページにもフッター表示
  
  let tl = [], td = 0;
  
  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, 0); 
  // ▼▼▼ [修正] (Req ⑦.3) サブタイトル廃止 ▼▼▼
  document.getElementById('report-subtitle').textContent = ''; 
  document.getElementById('report-subtitle').style.textAlign = 'left'; 
  
  // ▼▼▼ [修正] (Req ⑦.3) WCページではセパレーターマージンを 0 に ▼▼▼
  document.getElementById('report-separator').style.display='block';
  document.getElementById('report-separator').style.marginBottom = '0';
  
  const subNav = document.getElementById('comment-sub-nav');
  const controls = document.getElementById('comment-controls');
  if (subNav) subNav.innerHTML = '';
  if (controls) controls.innerHTML = '';
  

  const slideBody = document.getElementById('slide-body');
  slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
  slideBody.style.overflowY = 'hidden'; // スクロール禁止
  
  try {
      const cd = await getReportDataForCurrentClinic(currentClinicForModal);
      switch(columnType){
          case'L':tl=cd.npsData.rawText||[];td=cd.npsData.totalCount||0;break;
          case'I':tl=cd.feedbackData.i_column.results||[];td=cd.feedbackData.i_column.totalCount||0;break;
          case'J':tl=cd.feedbackData.j_column.results||[];td=cd.feedbackData.j_column.totalCount||0;break;
          case'M':tl=cd.feedbackData.m_column.results||[];td=cd.feedbackData.m_column.totalCount||0;break;
          default:console.error("Invalid column:",columnType);showLoading(false);return;
      }
  } catch(e) {
      console.error("Error accessing text data:", e);
      slideBody.innerHTML = `<p class="text-center text-red-500 py-16">レポートデータアクセスエラー</p>`;
      showLoading(false);
      return;
  }
  
  document.getElementById('report-title').textContent = getAnalysisTitle(columnType, td);
  
  if(tl.length === 0){
      slideBody.innerHTML = `<p class="text-center text-red-500 py-16">分析対象テキストなし</p>`;
      showLoading(false);
      return;
  }
  
  // ▼▼▼ [変更] WCシェル (Req ⑦.1: gap-1に変更, Req ⑦.2: text-left に変更) ▼▼▼
  slideBody.innerHTML = `
      <div class="grid grid-cols-2 gap-4 h-full">
          <div class="grid grid-cols-2 grid-rows-2 gap-1 h-full pr-2">
              <div id="noun-chart-container" class="chart-container h-full">
                  <h3 class="font-bold text-center mb-0 text-blue-600 text-sm">名詞</h3>
                  <div id="noun-chart" class="w-full h-[calc(100%-20px)]"></div>
              </div>
              <div id="verb-chart-container" class="chart-container h-full">
                  <h3 class="font-bold text-center mb-0 text-red-600 text-sm">動詞</h3>
                  <div id="verb-chart" class="w-full h-[calc(100%-20px)]"></div>
              </div>
              <div id="adj-chart-container" class="chart-container h-full">
                  <h3 class="font-bold text-center mb-0 text-green-600 text-sm">形容詞</h3>
                  <div id="adj-chart" class="w-full h-[calc(100%-20px)]"></div>
              </div>
              <div id="int-chart-container" class="chart-container h-full">
                  <h3 class="font-bold text-center mb-0 text-gray-600 text-sm">感動詞</h3>
                  <div id="int-chart" class="w-full h-[calc(100%-20px)]"></div>
              </div>
          </div>
          <div class="space-y-2 flex flex-col h-full">
              <p class="text-sm text-gray-600 text-left px-2">
                  スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。単語の色は品詞の種類で異なります。<br>
                  <span class="text-blue-600 font-semibold">青色=名詞</span>、<span class="text-red-600 font-semibold">赤色=動詞</span>、<span class="text-green-600 font-semibold">緑色=形容詞</span>、<span class="text-gray-600 font-semibold">灰色=感動詞</span>
              </p>
              <div id="word-cloud-container" class="h-[calc(100%-60px)] border border-gray-200"> <canvas id="word-cloud-canvas" class="!h-full !w-full"></canvas>
              </div>
              <div id="analysis-error" class="text-red-500 text-sm text-center hidden"></div>
          </div>
      </div>
  `;

  try{
      const r=await fetch('/api/analyzeText',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({textList:tl})});
      if(!r.ok){const et=await r.text();throw new Error(`分析APIエラー(${r.status}): ${et}`);}
      const ad=await r.json();
      setTimeout(() => drawAnalysisCharts(ad.results), 50);
  } catch(error){
      console.error('!!! Analyze fail:',error);
      document.getElementById('analysis-error').textContent=`分析失敗: ${error.message}`;
      document.getElementById('analysis-error').classList.remove('hidden');
  } finally {
      showLoading(false);
  }
}

// ▼▼▼ [修正] Word Cloud描画 (Req ⑦: 余白削除, ぼやけ解消, 10-17pt) ▼▼▼
function drawAnalysisCharts(results) { 
    if(!results||results.length===0){
        console.log("No analysis results.");
        document.getElementById('analysis-error').textContent='分析結果なし';
        document.getElementById('analysis-error').classList.remove('hidden');
        return;
    } 
    const nouns=results.filter(r=>r.pos==='名詞'),verbs=results.filter(r=>r.pos==='動詞'),adjs=results.filter(r=>r.pos==='形容詞'),ints=results.filter(r=>r.pos==='感動詞');
    
    // ▼▼▼ [修正] barOpt (レイアウト調整) ▼▼▼
    const barOpt={bars:'horizontal',legend:{position:'none'},
    hAxis:{title:'スコア(出現頻度)',minValue:0, textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}},
    vAxis:{title: null, textStyle:{fontSize:10}, titleTextStyle:{fontSize:12}}, 
    // ▼▼▼ [修正] chartArea の left/width を調整 (ご要望) ▼▼▼
    chartArea:{height:'90%', width:'70%', left:'25%', top:'5%'}, 
    backgroundColor: 'transparent'};
    
    // ▼▼▼ [修正] (Req ⑦) グラフの height: '100%' を削除し、余白を詰める ▼▼▼
    drawSingleBarChart(nouns.slice(0,8),'noun-chart',{...barOpt,colors:['#3b82f6'], width: '100%'});
    drawSingleBarChart(verbs.slice(0,8),'verb-chart',{...barOpt,colors:['#ef4444'], width: '100%'});
    drawSingleBarChart(adjs.slice(0,8),'adj-chart',{...barOpt,colors:['#22c55e'], width: '100%'});
    drawSingleBarChart(ints.slice(0,8),'int-chart',{...barOpt,colors:['#6b7280'], width: '100%'});
    
    // ▼▼▼ [修正] (Req 6) WordCloud (ぼやけ解消, 10-17pt) ▼▼▼
    const wl=results.map(r=>[r.word,r.score]).slice(0,100);
    const pm=results.reduce((map,item)=>{map[item.word]=item.pos;return map;},{});
    const cv=document.getElementById('word-cloud-canvas');
    
    if(WordCloud.isSupported&&cv){
        try{
            // ▼▼▼ [修正] (Req 6) 高解像度DPR対応 (ぼやけ解消) ▼▼▼
            const dpr = window.devicePixelRatio || 1;
            const rect = cv.getBoundingClientRect();
            cv.width = rect.width * dpr;
            cv.height = rect.height * dpr;
            const ctx = cv.getContext('2d');
            ctx.scale(dpr, dpr); // Scale context for high-res

            // ▼▼▼ [修正] (Req 6) サイズ計算ロジック (ご要望: 10〜17pt) ▼▼▼
            const minSize = 10;
            const maxSize = 17;
            
            let maxScore = 0;
            if (wl.length > 0) {
                 maxScore = wl[0][1]; // (scoreでソート済み前提)
            }
            
            // スコアを 10pt-17pt の範囲にマッピングする
            const weightFactor = (size) => {
                if (maxScore === 0) return minSize;
                const score = size;
                // スコアを 0-1 の範囲に正規化
                const normalizedScore = Math.max(0, score / maxScore);
                // 10pt (min) から 17pt (max) の範囲にマッピング
                return minSize + (maxSize - minSize) * normalizedScore;
            };
            // ▲▲▲
            
            const options={
                list:wl,
                gridSize: 8, // (密度)
                weightFactor: weightFactor, // (Req 6)
                minSize: minSize, // (Req 6)
                fontFamily:'Noto Sans JP,sans-serif',
                color:function(w,wt,fs,d,t){const p=pm[w]||'不明';switch(p){case'名詞':return'#3b82f6';case'動詞':return'#ef4444';case'形容詞':return'#22c55e';case'感動詞':return'#6b7280';default:return'#a8a29e';}},
                backgroundColor:'transparent',
                clearCanvas:true,
                // ▼▼▼ [修正] (Req 6) 安定化のため回転を無効化 ▼▼▼
                rotateRatio: 0
            };
            WordCloud(cv,options);
        }catch(wcError){
            console.error("Error drawing WordCloud:",wcError);
            document.getElementById('word-cloud-container').innerHTML=`<p class="text-center text-red-500">ワードクラウド描画エラー:${wcError.message}</p>`;
        }
    }else{
        console.warn("WordCloud unsupported/canvas missing.");
        document.getElementById('word-cloud-container').innerHTML='<p class="text-center text-gray-500">ワードクラウド非対応</p>';
    } 
}
function drawSingleBarChart(data, elementId, options) { const c=document.getElementById(elementId);if(!c){console.error(`Element not found: ${elementId}`);return;} if(!data||data.length===0){c.innerHTML='<p class="text-center text-gray-500 text-sm py-4">データなし</p>';return;} const cd=[['単語','スコア',{role:'style'}]];const color=options.colors&&options.colors.length>0?options.colors[0]:'#a8a29e';/* ▼▼▼ [修正] .reverse() を削除 (頻度高い順) ▼▼▼ */ data.slice().forEach(item=>{cd.push([item.word,item.score,color]);});try{const dt=google.visualization.arrayToDataTable(cd);const chart=new google.visualization.BarChart(c);chart.draw(dt,options);}catch(chartError){console.error(`Error drawing bar chart for ${elementId}:`,chartError);c.innerHTML=`<p class="text-center text-red-500 text-sm py-4">グラフ描画エラー<br>${chartError.message}</p>`;} }
function clearAnalysisCharts() { const nounChart = document.getElementById('noun-chart'); if(nounChart) nounChart.innerHTML=''; const verbChart = document.getElementById('verb-chart'); if(verbChart) verbChart.innerHTML=''; const adjChart = document.getElementById('adj-chart'); if(adjChart) adjChart.innerHTML=''; const intChart = document.getElementById('int-chart'); if(intChart) intChart.innerHTML=''; const c=document.getElementById('word-cloud-canvas');if(c){const x=c.getContext('2d');x.clearRect(0,0,c.width,c.height);} const analysisError = document.getElementById('analysis-error'); if(analysisError){analysisError.classList.add('hidden');analysisError.textContent='';} }
function getAnalysisTitle(columnType, count) { const bt=`アンケート結果　ー${getColumnName(columnType)}ー`;return`${bt}　※全回答数${count}件ー`; }
function getColumnName(columnType) { 
  switch(columnType){
      case'L':return'NPS推奨度 理由';
      case'I':case'I_good':case'I_bad':return'良かった点や悪かった点など';
      case'J':return'印象に残ったスタッフへのコメント';
      case'M':return'お産にかかわるご意見・ご感想';
      default:return'不明';
  } 
}


// --- ▼▼▼ [修正] AI詳細分析 (Screen 5) 処理 (Req ⑤, ⑥) ▼▼▼
async function prepareAndShowDetailedAnalysis(analysisType) {
  console.log(`Prepare detailed analysis: ${analysisType}`);
  const clinicName = currentClinicForModal;
  
  showLoading(true, `AI分析結果を読み込み中...\n(${getDetailedAnalysisTitleBase(analysisType)})`);
  showScreen('screen5');
  updateNavActiveState(null, null, analysisType);
  toggleEditDetailedAnalysis(false); 
  
  // ▼▼▼ [修正] Screen5のフッターも表示/非表示制御 ▼▼▼
  showCopyrightFooter(true, 'screen5'); 
  
  const errorDiv = document.getElementById('detailed-analysis-error');
  errorDiv.classList.add('hidden');
  errorDiv.textContent = '';
  
  document.getElementById('detailed-analysis-title').textContent = getDetailedAnalysisTitleFull(analysisType);
  // ▼▼▼ [修正] サブタイトル廃止 ▼▼▼
  document.getElementById('detailed-analysis-subtitle').textContent = ''; // getDetailedAnalysisSubtitleForUI(analysisType, 'analysis'); 
  
  // ▼▼▼ [修正] (Req ⑤) グラフページではセパレーターマージンを 16px に戻す ▼▼▼
  document.getElementById('report-separator').style.marginBottom = '16px';
  
  switchTab('analysis'); 
  
  try {
      console.log(`[AI] Fetching saved analysis from Sheet...`);
      const response = await fetch('/api/getDetailedAnalysis', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
              centralSheetId: currentCentralSheetId,
              clinicName: clinicName,
              columnType: analysisType
          })
      });

      if (!response.ok) {
          throw new Error(`分析結果の読み込み失敗 (${response.status}): ${await response.text()}`);
      }
      
      const savedData = await response.json(); 
      
      if (savedData.analysis && !savedData.analysis.includes('（データがありません）')) {
          console.log(`[AI] Found saved data in Sheet.`);
          displayDetailedAnalysis(savedData, analysisType, false); 
          showLoading(false);
      } else {
          console.log(`[AI] No saved data found. Running initial analysis...`);
          showLoading(true, `初回AI分析を実行中...\n(${getDetailedAnalysisTitleBase(analysisType)})`);
          await runDetailedAnalysisGeneration(analysisType);
      }

  } catch (err) {
      console.error('!!! Detailed analysis failed:', err);
      errorDiv.textContent = `分析失敗: ${err.message}`;
      errorDiv.classList.remove('hidden');
      clearDetailedAnalysisDisplay();
      showLoading(false);
  }
}

// (AI再実行 - 変更なし)
async function handleRegenerateDetailedAnalysis() {
    const typeName = getDetailedAnalysisTitleBase(currentDetailedAnalysisType);
    if (!confirm(`「${typeName}」のAI分析を再実行しますか？\n\n・現在の分析内容は破棄されます。\n・編集中の内容は保存されません。`)) {
        return;
    }
    showLoading(true, `AI分析を再実行中...\n(${typeName})`);
    toggleEditDetailedAnalysis(false); 
    await runDetailedAnalysisGeneration(currentDetailedAnalysisType);
}

// (AI実行 (共通) - 変更なし)
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
                 throw new Error(`分析対象のテキストが0件のため、AI分析を実行できませんでした。`);
            }
            throw new Error(`AI分析APIエラー (${response.status}): ${errorText}`);
        }
        const analysisJson = await response.json(); 
        displayDetailedAnalysis(analysisJson, analysisType, true); 
        switchTab('analysis');
    } catch (err) {
        console.error('!!! AI Generation failed:', err);
        errorDiv.textContent = `AI分析実行エラー: ${err.message}`;
        errorDiv.classList.remove('hidden');
        clearDetailedAnalysisDisplay();
    } finally {
        showLoading(false);
    }
}

// (AI表示 (共通) - ▼▼▼ [修正] フォント調整呼び出し追加 (ルール ②, ④, ⑥) ▼▼▼)
function displayDetailedAnalysis(data, analysisType, isRawJson) {
    let analysisText, suggestionsText, overallText;
    if (isRawJson) {
        analysisText = (data.analysis && data.analysis.themes) ? data.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n') : '（分析データがありません）';
        suggestionsText = (data.suggestions && data.suggestions.items) ? data.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n') : '（改善提案データがありません）';
        overallText = (data.overall && data.overall.summary) ? data.overall.summary : '（総評データがありません）';
    } else {
        analysisText = data.analysis;
        suggestionsText = data.suggestions;
        overallText = data.overall;
    }
    
    // --- 表示エリア (display-*) ---
    const displayAnalysis = document.getElementById('display-analysis');
    const displaySuggestions = document.getElementById('display-suggestions');
    const displayOverall = document.getElementById('display-overall');
    
    displayAnalysis.textContent = analysisText;
    displaySuggestions.textContent = suggestionsText;
    displayOverall.textContent = overallText;

    // --- 編集エリア (textarea-*) ---
    document.getElementById('textarea-analysis').value = analysisText;
    document.getElementById('textarea-suggestions').value = suggestionsText;
    document.getElementById('textarea-overall').value = overallText;
    
    // ▼▼▼ [新規] フォントサイズ自動調整 (ルール ⑥, ④) ▼▼▼
    // (ルール ⑦: 基本 12pt)
    adjustFontSize(displayAnalysis, 12);
    adjustFontSize(displaySuggestions, 12);
    adjustFontSize(displayOverall, 12);
    
    // (編集エリアの高さ調整)
    ['analysis', 'suggestions', 'overall'].forEach(tabId => {
         const textarea = document.getElementById(`textarea-${tabId}`);
         if (textarea) {
             textarea.style.height = 'auto';
             textarea.style.height = (textarea.scrollHeight + 5) + 'px';
         }
    });
}

// (AI表示 (クリア) - 変更なし)
function clearDetailedAnalysisDisplay() {
    document.getElementById('display-analysis').textContent = '';
    document.getElementById('textarea-analysis').value = '';
    document.getElementById('display-suggestions').textContent = '';
    document.getElementById('textarea-suggestions').value = '';
    document.getElementById('display-overall').textContent = '';
    document.getElementById('textarea-overall').value = '';
}

// (タブ切り替え - ▼▼▼ [修正] (Req ⑥) タブ切り替えセレクタ修正 ▼▼▼)
function handleTabClick(event) { 
  const tabId = event.target.dataset.tabId; 
  if (tabId) { 
      switchTab(tabId); 
  } 
}
function switchTab(tabId) { 
    // ▼▼▼ [修正] サブタイトル廃止 ▼▼▼
    document.getElementById('detailed-analysis-subtitle').textContent = ''; // getDetailedAnalysisSubtitleForUI(currentDetailedAnalysisType, tabId);
    
    document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => { if (button.dataset.tabId === tabId) { button.classList.add('active'); } else { button.classList.remove('active'); } }); 
    
    // ▼▼▼ [修正] (Req ⑥) セレクタを .tab-content から .ai-analysis-container に変更 ▼▼▼
    document.querySelectorAll('#detailed-analysis-content-area > .ai-analysis-container').forEach(content => { 
        if (content.id === `content-${tabId}`) { 
            content.classList.remove('hidden'); 
        } else { 
            content.classList.add('hidden'); 
        } 
    }); 

    if (isEditingDetailedAnalysis) {
        // 編集中の場合：textareaの高さを調整
        const activeTextarea = document.getElementById(`textarea-${tabId}`);
        if (activeTextarea) {
            activeTextarea.style.height = 'auto';
            activeTextarea.style.height = (activeTextarea.scrollHeight + 5) + 'px';
        }
    } else {
        // ▼▼▼ [新規] 表示中の場合：フォントサイズを再調整 (ルール ⑥, ④) ▼▼▼
        const activeDisplay = document.getElementById(`display-${tabId}`);
        if (activeDisplay) {
            adjustFontSize(activeDisplay, 12);
        }
    }
}

// (編集モード切り替え - ▼▼▼ [修正] (Req ⑥) 保存機能は存在するため変更なし ▼▼▼)
function toggleEditDetailedAnalysis(isEdit) {
    isEditingDetailedAnalysis = isEdit;
    const editBtn = document.getElementById('edit-detailed-analysis-btn');
    const regenBtn = document.getElementById('regenerate-detailed-analysis-btn');
    const saveBtn = document.getElementById('save-detailed-analysis-btn');
    const cancelBtn = document.getElementById('cancel-edit-detailed-analysis-btn');
    const displayAreas = document.querySelectorAll('.ai-analysis-content');
    const editAreas = document.querySelectorAll('[id^="edit-"]');

    if (isEditingDetailedAnalysis) {
        editBtn.classList.add('hidden');
        regenBtn.classList.add('hidden');
        saveBtn.classList.remove('hidden');
        cancelBtn.classList.remove('hidden');
        displayAreas.forEach(el => el.classList.add('hidden'));
        editAreas.forEach(el => el.classList.remove('hidden'));
        
        const activeTab = document.querySelector('#ai-tab-nav .tab-button.active').dataset.tabId;
        const activeTextarea = document.getElementById(`textarea-${activeTab}`);
        
        if (activeTextarea) {
            activeTextarea.style.height = 'auto';
            activeTextarea.style.height = (activeTextarea.scrollHeight + 5) + 'px';
        }
    } else {
        editBtn.classList.remove('hidden');
        regenBtn.classList.remove('hidden');
        saveBtn.classList.add('hidden');
        cancelBtn.classList.add('hidden');
        displayAreas.forEach(el => el.classList.remove('hidden'));
        editAreas.forEach(el => el.classList.add('hidden'));
        
        // ▼▼▼ [新規] 表示モードに戻る際、フォントサイズを再調整 (ルール ⑥, ④) ▼▼▼
        const activeTab = document.querySelector('#ai-tab-nav .tab-button.active').dataset.tabId;
        const activeDisplay = document.getElementById(`display-${activeTab}`);
        if (activeDisplay) {
            adjustFontSize(activeDisplay, 12);
        }
    }
}

// (AI詳細分析 (保存) - (Req ⑥) 機能は存在するため変更なし)
async function saveDetailedAnalysisEdits() {
    showLoading(true, '変更を保存中...');
    const analysisContent = document.getElementById('textarea-analysis').value;
    const suggestionsContent = document.getElementById('textarea-suggestions').value;
    const overallContent = document.getElementById('textarea-overall').value;
    const activeTabId = document.querySelector('#ai-tab-nav .tab-button.active').dataset.tabId;
    
    let contentToSave = '';
    if (activeTabId === 'analysis') contentToSave = analysisContent;
    else if (activeTabId === 'suggestions') contentToSave = suggestionsContent;
    else if (activeTabId === 'overall') contentToSave = overallContent;
    else {
        alert('不明なタブが選択されています。保存を中止します。');
        showLoading(false);
        return;
    }

    try {
        const response = await fetch('/api/updateDetailedAnalysis', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                centralSheetId: currentCentralSheetId,
                clinicName: currentClinicForModal,
                columnType: currentDetailedAnalysisType, 
                tabId: activeTabId, 
                content: contentToSave 
            })
        });
        
        if (!response.ok) {
            throw new Error(`保存失敗 (${response.status}): ${await response.text()}`);
        }
        
        document.getElementById('display-analysis').textContent = analysisContent;
        document.getElementById('display-suggestions').textContent = suggestionsContent;
        document.getElementById('display-overall').textContent = overallContent;
        
        toggleEditDetailedAnalysis(false);
        alert('変更を保存しました。');
    } catch (e) {
        console.error("Failed to save edits:", e);
        alert(`保存中にエラーが発生しました。\n${e.message}`);
    } finally {
        showLoading(false);
    }
}

// (AIタイトル・サブタイトル取得関数 - ▼▼▼ [修正] サブタイトル廃止 ▼▼▼)
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

// ▼▼▼ [修正] サブタイトルを常に空文字('')に (ご要望) ▼▼▼
function getDetailedAnalysisSubtitleForUI(analysisType, tabId) {
     return '';
}

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
// --- ▲▲▲ AI分析 終わり ▲▲▲


// --- 汎用関数 (変更なし) ---
function updateNavActiveState(activeReportType, activeAnalysisType, activeDetailedAnalysisType) {
  const navs = ['#report-nav', '#report-nav-screen5']; 
  navs.forEach(navSelector => {
      document.querySelectorAll(`${navSelector} .btn`).forEach(btn => {
          if ((activeReportType && btn.dataset.reportType === activeReportType) ||
              (activeAnalysisType && btn.dataset.analysisType === activeAnalysisType) ||
              (activeDetailedAnalysisType && btn.dataset.detailedAnalysisType === activeDetailedAnalysisType)) {
              btn.classList.add('btn-active');
          } else {
              btn.classList.remove('btn-active');
          }
      });
  });
}

function showScreen(screenId){ document.querySelectorAll('.screen').forEach(el=>el.classList.add('hidden'));document.getElementById(screenId).classList.remove('hidden'); }
function showLoading(isLoading, message=''){ const o=document.getElementById('loading-overlay'),m=document.getElementById('loading-message');if(isLoading){m.textContent=message;o.classList.remove('hidden');}else{o.classList.add('hidden');m.textContent='';} }

// (PDF生成 - 変更なし)
async function generatePdf(){
  console.log('PDF export clicked.');
  const cn=currentClinicForModal;
  const pt = currentPeriodText; 
  if(!cn || !pt || !currentCentralSheetId){
      alert('PDF生成に必要なレポートデータ（クリニック名、期間、集計ID）がありません。');
      return;
  }
  showLoading(true,'PDFを生成中...');
  try{
      const r=await fetch('/generate-pdf',{
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({
              clinicName: cn,
              periodText: pt,
              centralSheetId: currentCentralSheetId 
          })
      });
      if(!r.ok){let em=`サーバーエラー: ${r.status} ${r.statusText}`;try{const ed=await r.text();em+=`\n${ed}`;}catch(e){} throw new Error(em);}
      const b=await r.blob();const u=window.URL.createObjectURL(b);const a=document.createElement('a');a.style.display='none';a.href=u;document.body.appendChild(a);a.click();window.URL.revokeObjectURL(u);a.remove();
  } catch(error){
      console.error('PDF gen error:',error);
      alert('PDF生成失敗\n'+error.message);
  } finally {
      showLoading(false);
  }
}

// ▼▼▼ [新規] フッター表示切替 ▼▼▼
function showCopyrightFooter(show, screenId = 'screen3') {
    let footer;
    if (screenId === 'screen3') {
        footer = document.getElementById('report-copyright');
    } else if (screenId === 'screen5') {
        // Screen 5 のフッターも制御 (Screen 5 の .report-body 内の .report-copyright を探す)
        footer = document.querySelector('#screen5 .report-copyright');
    }
    
    if (footer) {
        footer.style.display = show ? 'block' : 'none';
    }
}


// --- ▼▼▼ [修正] 初期化ブロック (ファイル末尾に移動) ▼▼▼ ---
(async () => {
  console.log('DOM Loaded (assumed).');
  
  // 1. プルダウンを生成
  populateDateSelectors();
  
  // 2. Google Charts のロードを待つ
  try {
      await googleChartsLoaded;
      console.log('Charts loaded.');
  } catch (err) {
      console.error('Chart load fail:', err);
      alert('グラフライブラリ読込失敗');
      // (エラーが発生してもリスナーは設定する)
  }
  
  // 3. イベントリスナーを設定
  setupEventListeners();
  console.log('Listeners setup.');
})(); // 即時実行関数で全体をラップ
