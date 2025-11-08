// --- グローバル変数 ---
  let selectedPeriod = {}; 
  let currentClinicForModal = '';
  let slidesData = []; 
  let currentPage = 1; 
  let currentAnalysisTarget = 'L'; 
  let currentDetailedAnalysisType = 'L'; 
  let isEditingDetailedAnalysis = false; 
  let currentCentralSheetId = null; 
  let currentPeriodText = ""; 
  let currentAiCompletionStatus = {}; 
  let overallDataCache = null; 
  let clinicDataCache = null; 
  
  // ▼▼▼ [新規] コメント保存ボタンのDOM
  let commentSaveBtn = null;
  // --- ▲▲▲ ---

  // --- Google Charts ロード (変更なし) ---
  const googleChartsLoaded = new Promise(resolve => { google.charts.load('current', {'packages':['corechart', 'bar']}); google.charts.setOnLoadCallback(resolve); });

  // --- 初期化 (変更なし) ---
  document.addEventListener('DOMContentLoaded', async () => {
    console.log('DOM Loaded.');
    populateDateSelectors();
    try {
        await googleChartsLoaded;
        console.log('Charts loaded.');
    } catch (err) {
        console.error('Chart load fail:', err);
        alert('グラフライブラリ読込失敗');
        return;
    }
    setupEventListeners();
    console.log('Listeners setup.');
  });

  // --- ▼▼▼ [修正] イベントリスナー設定 (WC-が開けない問題解消のため修正 + コメント保存) ---
  function setupEventListeners() {
    document.getElementById('next-to-clinics').addEventListener('click', handleNextToClinics); 
    document.getElementById('issue-btn').addEventListener('click', handleIssueReport); 
    document.getElementById('issued-list-container').addEventListener('click', handleIssuedListClickDelegator);
    
    document.getElementById('report-nav').addEventListener('click', handleReportNavClick);
    document.getElementById('report-nav-screen5').addEventListener('click', handleReportNavClick);
    
    document.getElementById('prev-slide').addEventListener('click', handlePrevSlide);
    document.getElementById('next-slide').addEventListener('click', handleNextSlide);
    document.getElementById('back-to-clinics').addEventListener('click', () => showScreen('screen2'));
    
    // [新規] コメント保存ボタンのリスナー
    commentSaveBtn = document.getElementById('comment-save-btn');
    commentSaveBtn.addEventListener('click', handleSaveComment);
    
    document.getElementById('back-to-period').addEventListener('click', () => {
        currentCentralSheetId = null; 
        currentPeriodText = "";
        currentAiCompletionStatus = {}; 
        showScreen('screen1');
    });
    
    document.getElementById('pdf-export-btn').addEventListener('click', generatePdf); 
    document.getElementById('pdf-export-btn-screen5').addEventListener('click', generatePdf);
    
    // AI分析タブのリスナーは #ai-tab-nav に移動
    document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => { button.addEventListener('click', handleTabClick); });
    
    document.getElementById('regenerate-detailed-analysis-btn').addEventListener('click', handleRegenerateDetailedAnalysis);
    document.getElementById('edit-detailed-analysis-btn').addEventListener('click', () => toggleEditDetailedAnalysis(true));
    document.getElementById('save-detailed-analysis-btn').addEventListener('click', saveDetailedAnalysisEdits);
    document.getElementById('cancel-edit-detailed-analysis-btn').addEventListener('click', () => {
        if (confirm('編集内容を破棄しますか？')) {
            toggleEditDetailedAnalysis(false); 
        }
    });
    document.getElementById('back-to-clinics-from-detailed-analysis').addEventListener('click', () => {
        toggleEditDetailedAnalysis(false); 
        showScreen('screen2');
    });
  }

  // --- 画面1/2 処理 ---
  function populateDateSelectors() {
    const now = new Date();
    const cy = now.getFullYear();
    const sy = document.getElementById('start-year');
    const ey = document.getElementById('end-year');
    
    for(let i = 0; i < 5; i++) {
        const y = cy - i;
        sy.add(new Option(`${y}年`, y));
        ey.add(new Option(`${y}年`, y));
    }

    const sm = document.getElementById('start-month');
    const em = document.getElementById('end-month');
    
    for(let i = 1; i <= 12; i++) {
        const m = String(i).padStart(2, '0');
        sm.add(new Option(`${i}月`, m));
        em.add(new Option(`${i}月`, m));
    }
    
    em.value = String(now.getMonth() + 1).padStart(2, '0');
    
    // ▼▼▼ [修正] 開始年の初期値を正しく設定する
    sy.value = String(cy); 
    // ▲▲▲
    
    // (開始月も現在の月に設定)
    sm.value = String(now.getMonth() + 1).padStart(2, '0');
  }

  async function handleNextToClinics() {
    console.log('Next clicked.');
    const sy=document.getElementById('start-year').value;
    const sm=document.getElementById('start-month').value;
    const ey=document.getElementById('end-year').value;
    const em=document.getElementById('end-month').value;
    const sd=new Date(`${sy}-${sm}-01`);
    const ed=new Date(`${ey}-${em}-01`);

    if(sd > ed){
        console.warn('Start>End.');
        alert('開始年月<=終了年月で設定');
        document.getElementById('start-year').classList.add('border-red-500');
        document.getElementById('start-month').classList.add('border-red-500');
        return;
    }
    document.getElementById('start-year').classList.remove('border-red-500');
    document.getElementById('start-month').classList.remove('border-red-500');

    selectedPeriod = {start: `${sy}-${sm}`, end: `${ey}-${em}`};
    currentPeriodText = `${sy}-${sm}～${ey}-${em}`; 
    const displayPeriod = `${sy}年${sm}月～${ey}年${em}月`;

    console.log('Period:', selectedPeriod, 'File Name:', currentPeriodText);
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
        console.log(`Got Central Sheet ID: ${currentCentralSheetId}`);
        
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
    console.log('loadClinics (new)...');
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
                    
                    d.innerHTML = `
                        <p class="font-bold text-base view-report-btn" data-clinic-name="${clinicName}">${clinicName}</p>
                    `;
                    
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
    console.log('Issue report (ETL).');
    const sc = Array.from(document.querySelectorAll('#clinic-list-container input:checked'))
                  .map(cb => cb.value);
    if (sc.length === 0) { alert('転記対象のクリニックが選択されていません。'); return; }
    if (sc.length > 10) { alert('一度に選択できる件数は10件までです。チェックを減らしてください。'); return; }

    showLoading(true, '集計スプレッドシートへデータ転記中...\n完了後、バックグラウンドで分析タブが自動生成されます。');
    console.log('Selected for ETL:', sc);

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
        console.log('ETL process finished:', data);
        loadClinics(); 
    } catch(err) {
        console.error('!!! ETL Issue failed:', err);
        alert(`データ転記失敗\n${err.message}`);
    } finally {
        showLoading(false);
    }
  }

  // --- 画面3/5 (レポート表示/AI分析) 処理 ---

  // イベント委譲 (変更なし)
  function handleIssuedListClickDelegator(e) {
    const viewBtn = e.target.closest('.view-report-btn');
    if (viewBtn) {
        handleIssuedListClick(viewBtn.dataset.clinicName);
    }
  }

  // 転記済みレポートクリック (変更なし)
  async function handleIssuedListClick(clinicName) {
    currentClinicForModal = clinicName;
    console.log(`Issued clinic clicked: ${currentClinicForModal}`);
    
    clinicDataCache = null;
    overallDataCache = null;
    
    prepareAndShowReport('cover');
  }
  
  // ▼▼▼ [修正] ナビゲーションクリック (WC-が開けない問題解消) ▼▼▼
  function handleReportNavClick(e) {
    const targetButton = e.target.closest('.btn');
    if (!targetButton) return;
    const reportType = targetButton.dataset.reportType;
    const analysisType = targetButton.dataset.analysisType;
    const detailedAnalysisType = targetButton.dataset.detailedAnalysisType;
    
    if (reportType) {
        showScreen('screen3'); // コメント/グラフ系のボタンでも即座に画面3へ遷移
        prepareAndShowReport(reportType);
    } else if (analysisType) {
        currentAnalysisTarget = analysisType;
        // WC-ボタンクリック時にも showScreen('screen3') を実行し、遷移を確実にする
        showScreen('screen3');
        prepareAndShowAnalysis(analysisType); 
    } else if (detailedAnalysisType) {
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


  // --- ▼▼▼ [大幅修正] レポート表示メイン (Screen 3) (コメント白紙対策) ▼▼▼ ---
  async function prepareAndShowReport(reportType){
    console.log(`Prepare report: ${reportType}`); 
    showLoading(true,'レポートデータ集計中...');
    
    showScreen('screen3');
    updateNavActiveState(reportType, null, null);
    
    slidesData=[];
    document.getElementById('pagination').style.display='none'; 
    document.getElementById('report-separator').style.display='block'; 
    document.getElementById('report-title').style.textAlign = 'left'; // デフォルト左揃え
    
    // 初期化
    document.getElementById('report-title').textContent = '';
    document.getElementById('report-subtitle').textContent = '';
    document.getElementById('slide-header').innerHTML = '';
    document.getElementById('slide-body').style.whiteSpace='pre-wrap';
    document.getElementById('slide-body').innerHTML='';
    document.getElementById('slide-body').classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');

    // --- [変更] コメント系レポートの判定 ---
    const commentTypeMap = {
        'nps': 'L',
        'feedback_i': 'I',
        'feedback_j': 'J',
        'feedback_m': 'M'
    };
    const commentType = commentTypeMap[reportType];

    // --- 1. コメント系レポートの場合 (新APIを呼ぶ) ---
    if (commentType) {
        try {
            console.log(`[Comments] Fetching new comment data for type: ${commentType}`);
            showLoading(true, 'コメントデータを読み込み中...');
            
            const response = await fetch('/api/getCommentData', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    centralSheetId: currentCentralSheetId, 
                    clinicName: currentClinicForModal, 
                    commentType: commentType 
                })
            });

            if (!response.ok) {
                throw new Error(`コメント取得APIエラー (${response.status}): ${await response.text()}`);
            }
            
            const commentData = await response.json();
            
            // スライドデータを構築
            buildCommentSlides(reportType, commentData);
            
            if (slidesData.length > 0) {
                currentPage = 1; 
                renderSlide(currentPage);
            } else {
                document.getElementById('slide-body').innerHTML = '<p class="text-center text-gray-500 py-16">コメントデータがありません</p>';
                document.getElementById('pagination').style.display = 'none';
            }
            
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
        // (コメント系は除外されたので、ここはグラフ系のみ)
        clinicData = await getReportDataForCurrentClinic(currentClinicForModal);
        overallData = await getReportDataForCurrentClinic("全体");
    } catch (e) {
        console.error('Chart data fetch error:', e);
        document.getElementById('slide-body').innerHTML = `<p class="text-center text-red-500 py-16">グラフデータ取得失敗<br>(${e.message})</p>`;
        showLoading(false);
        return;
    }
    
    // ページタイトル設定とグラフ描画準備
    if (reportType === 'nps_score'){ prepareChartPage('アンケート結果　ーNPS(ネットプロモータースコア)＝推奨度ー', 'これから初めてお産を迎える友人知人がいた場合、\nご出産された産婦人科医院をどのくらいお勧めしたいですか。\n友人知人への推奨度を教えてください。＜推奨度＞ 10:強くお勧めする〜 0:全くお勧めしない', 'nps_score'); isChart=true; }
    else if (reportType === 'satisfaction_b'){ prepareChartPage('アンケート結果　ー満足度ー','ご出産された産婦人科医院への満足度について、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_b'); isChart=true; }
    else if (reportType === 'satisfaction_c'){ prepareChartPage('アンケート結果　ー施設の充実度・快適さー','ご出産された産婦人科医院への施設の充実度・快適さについて、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_c'); isChart=true; }
    else if (reportType === 'satisfaction_d'){ prepareChartPage('アンケート結果　ーアクセスの良さー','ご出産された産婦人科医院へのアクセスの良さについて、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_d'); isChart=true; }
    else if (reportType === 'satisfaction_e'){ prepareChartPage('アンケート結果　ー費用ー','ご出産された産婦人科医院への費用について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_e'); isChart=true; }
    else if (reportType === 'satisfaction_f'){ prepareChartPage('アンケート結果_ー病院の雰囲気ー','ご出産された産婦人科医院への病院の雰囲気について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_f'); isChart=true; }
    else if (reportType === 'satisfaction_g'){ prepareChartPage('アンケート結果　ーースタッフの対応ー','ご出産された産婦人科医院へのスタッフの対応について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_g'); isChart=true; }
    else if (reportType === 'satisfaction_h'){ prepareChartPage('アンケート結果　ー先生の診断・説明ー','ご出産された産婦人科医院への先生の診断・説明について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満', 'satisfaction_h'); isChart=true; }
    else if (reportType === 'age'){ prepareChartPage('アンケート結果　ーご回答者さまの年代ー','ご出産された方の年代について教えてください。', 'age'); isChart=true; }
    else if (reportType === 'children'){ prepareChartPage('アンケート結果　ーご回答者さまのお子様の人数ー','ご出産された方のお子様の人数について教えてください。', 'children'); isChart=true; }
    else if (reportType === 'income'){ prepareChartPage('アンケート結果　ーご回答者さまの世帯年収ー','ご出産された方の世帯年収について教えてください。', 'income', true); isChart=true; }

    
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
        // (WC分析は別ルート / コメントは上部で処理済みの)
        showLoading(false);
    } 
  }

  // 例外構成（表紙・目次・概要）の表示 (変更なし)
  async function prepareAndShowIntroPages(reportType) {
    document.getElementById('report-separator').style.display='none'; 
    document.getElementById('report-subtitle').style.textAlign = 'center'; 
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
        let overallCount = 0;
        let clinicCount = 0;
        let clinicListCount = 0;
        
        try {
            const overallData = await getReportDataForCurrentClinic("全体");
            overallCount = overallData.npsScoreData.totalCount || 0;
            clinicListCount = (await fetch('/api/getTransferredList', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ centralSheetId: currentCentralSheetId })
            }).then(r => r.json())).sheetTitles.length - 2; 
            
            const clinicData = await getReportDataForCurrentClinic(currentClinicForModal);
            clinicCount = clinicData.npsScoreData.totalCount || 0;
        } catch (e) {
             console.warn("Error fetching data for summary:", e);
        }

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
  
  // グラフ描画用シェル設定 
  function prepareChartPage(title, subtitle, type, isBar=false) { 
    document.getElementById('report-title').textContent = title;
    document.getElementById('report-subtitle').textContent = subtitle;
    document.getElementById('report-subtitle').style.textAlign = 'center'; // グラフページは中央揃え
    document.getElementById('report-separator').style.display='block';

    let htmlContent = '';
    const cid = isBar ? 'bar-chart' : 'pie-chart';
    
    // 本体部分の高さをレポート全体の高さに合わせて調整
    const chartHeightClass = 'h-[350px]'; // グラフ描画エリアの高さ調整

    if (type === 'nps_score') {
        htmlContent = `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-full">
                <div class="flex flex-col h-full">
                    <h3 id="clinic-chart-header" class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
                    <div id="clinic-bar-chart" class="w-full ${chartHeightClass} bg-gray-50 border border-gray-200 flex items-center justify-center"></div>
                    <div class="w-full h-[150px] flex flex-col justify-center items-center mt-4">
                        <p class="text-sm text-gray-500 mb-2">【画像入力エリア】</p>
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
                    <div id="clinic-${cid}" class="w-full ${chartHeightClass}"></div>
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

  // --- ▼▼▼ [修正] グラフ描画関数 (フォントサイズを 12pt に修正) ▼▼▼
  function drawSatisfactionCharts(clinicChartData, overallChartData){ const opt={is3D:true,chartArea:{left:'5%',top:'5%',width:'90%',height:'90%'},pieSliceText:'percentage',pieSliceTextStyle:{color:'black',fontSize:12,bold:true},legend:{position:'labeled',textStyle:{color:'black',fontSize:12}},tooltip:{showColorCode:true,textStyle:{fontSize:12},trigger:'focus'},colors:['#4285F4','#DB4437','#F4B400','#0F9D58','#990099']};const cdEl=document.getElementById('clinic-pie-chart');if (!cdEl) throw new Error('グラフ描画エリア(clinic-pie-chart)が見つかりません。');if(clinicChartData&&clinicChartData.length>1&&clinicChartData.slice(1).some(row=>row[1]>0)){const d=google.visualization.arrayToDataTable(clinicChartData);const c=new google.visualization.PieChart(cdEl);c.draw(d,opt);} else {cdEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const avgEl=document.getElementById('average-pie-chart');if (!avgEl) throw new Error('グラフ描画エリア(average-pie-chart)が見つかりません。');if(overallChartData&&overallChartData.length>1&&overallChartData.slice(1).some(row=>row[1]>0)){const avgD=google.visualization.arrayToDataTable(overallChartData);const avgC=new google.visualization.PieChart(avgEl);avgC.draw(avgD,opt);} else {avgEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} }
  function drawIncomeCharts(clinicData, overallData){ const opt={legend:{position:'none'},colors:['#DE5D83'],annotations:{textStyle:{fontSize:12,color:'black',auraColor:'none'},alwaysOutside:false,stem:{color:'transparent'}},vAxis:{format:"#.##'%'",viewWindow:{min:0}, textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}}, hAxis:{textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}}};const ccdEl=document.getElementById('clinic-bar-chart');if (!ccdEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');if(clinicData.totalCount > 0 && clinicData.results && clinicData.results.length > 1){const cd=google.visualization.arrayToDataTable(clinicData.results);const cc=new google.visualization.ColumnChart(ccdEl);cc.draw(cd,opt);} else {ccdEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const avgEl=document.getElementById('average-bar-chart');if (!avgEl) throw new Error('グラフ描画エリア(average-bar-chart)が見つかりません。');if(overallData.totalCount > 0 && overallData.results && overallData.results.length > 1){const avgD=google.visualization.arrayToDataTable(overallData.results);const avgC=new google.visualization.ColumnChart(avgEl);avgC.draw(avgD,opt); } else {avgEl.innerHTML='<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} }
  function drawNpsScoreCharts(clinicData, overallData) { const clinicChartEl = document.getElementById('clinic-bar-chart');if (!clinicChartEl) throw new Error('グラフ描画エリア(clinic-bar-chart)が見つかりません。');const clinicNpsScore = calculateNps(clinicData.counts, clinicData.totalCount);const overallNpsScore = calculateNps(overallData.counts, overallData.totalCount);const clinicChartData = [['スコア', '割合', { role: 'annotation' }]];if (clinicData.totalCount > 0) {for (let i = 0; i <= 10; i++) {const count = clinicData.counts[i] || 0;const percentage = (count / clinicData.totalCount) * 100;clinicChartData.push([String(i), percentage, `${Math.round(percentage)}%`]);}} const opt = {legend: { position: 'none' },colors: ['#DE5D83'], annotations: {textStyle: { fontSize: 12, color: 'black', auraColor: 'none' },alwaysOutside: false,stem: { color: 'transparent' }},vAxis: { format: "#.##'%'", title: '割合(%)', viewWindow: { min: 0 }, textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}},hAxis: { title: '推奨度スコア (0〜10)', textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}},bar: { groupWidth: '80%' },isStacked: false, chartArea:{height:'75%', width:'90%', left:'5%', top:'5%'}};if (clinicData.totalCount > 0 && clinicChartData.length > 1) {const clinicDataVis = google.visualization.arrayToDataTable(clinicChartData);const clinicChart = new google.visualization.ColumnChart(clinicChartEl);clinicChart.draw(clinicDataVis, opt);} else {clinicChartEl.innerHTML = '<div class="flex items-center justify-center h-full"><p class="text-gray-500">データなし</p></div>';} const summaryArea = document.getElementById('nps-summary-area');if (summaryArea) {summaryArea.innerHTML = ` <div class="text-left text-3xl space-y-5 p-6 border rounded-lg bg-gray-50 shadow-inner w-full max-w-xs"> <p>全体：<span class="font-bold text-gray-800">${overallNpsScore.toFixed(1)}</span></p> <p>貴院：<span class="font-bold text-red-600">${clinicNpsScore.toFixed(1)}</span></p> </div> `;} const clinicHeaderEl = document.getElementById('clinic-chart-header');if (clinicHeaderEl) {clinicHeaderEl.textContent = `貴院の結果 (全 ${clinicData.totalCount} 件)`;} }
  function calculateNps(counts, totalCount) { if (totalCount === 0) return 0;let promoters = 0, passives = 0, detractors = 0;for (let i = 0; i <= 10; i++) {const count = counts[i] || 0;if (i >= 9) promoters += count;else if (i >= 7) passives += count;else detractors += count;} return ((promoters / totalCount) - (detractors / totalCount)) * 100; }

  // --- ▼▼▼ [削除] 古いコメントページ作成関数 ▼▼▼ ---
  // function prepareNpsCommentPages(results, baseTitle, totalCount) { /* ... 削除 ... */ }
  // function prepareCommentPages(data, title){ /* ... 削除 ... */ }
  
  // --- ▼▼▼ [新規] コメントスライド構築関数 ▼▼▼ ---
  /**
   * [新規] APIから取得したコメントデータ(新形式)を
   * 編集可能な<textarea>を含むスライドとして `slidesData` に構築する
   */
  function buildCommentSlides(reportType, commentData) {
      slidesData = [];
      const MAX_PER_PAGE = 20; // 1ページあたり20件
      
      let baseTitle = '';
      let isNPS = (reportType === 'nps');
      
      // 1. NPSの場合 (commentData は { "10": [...], "9": [...] } オブジェクト)
      if (isNPS) {
          baseTitle = 'アンケート結果　ーNPS推奨度 理由ー';
          document.getElementById('report-title').textContent = baseTitle;
          document.getElementById('report-subtitle').textContent = 'データ一覧（20データずつ）';

          const scoreGroups = [
              { key: '10', label: '推奨度10点', col: 'A', data: commentData['10'] || [] },
              { key: '9', label: '推奨度9点', col: 'B', data: commentData['9'] || [] },
              { key: '8', label: '推奨度8点', col: 'C', data: commentData['8'] || [] },
              { key: '7', label: '推奨度7点', col: 'D', data: commentData['7'] || [] },
              { key: '6_under', label: '推奨度6点以下', col: 'E', data: commentData['6_under'] || [] }
          ];

          for (const group of scoreGroups) {
              const comments = group.data;
              if (comments.length === 0) continue;

              const totalPages = Math.ceil(comments.length / MAX_PER_PAGE);
              
              for (let i = 0; i < comments.length; i += MAX_PER_PAGE) {
                  const chunk = comments.slice(i, i + MAX_PER_PAGE);
                  const pageNum = (i / MAX_PER_PAGE) + 1;
                  const slideHeader = `${group.label} (${comments.length}件) - ( ${pageNum} / ${totalPages} )`;
                  
                  let bodyHtml = '<div>';
                  chunk.forEach((comment, index) => {
                      const sheetRowIndex = i + index + 2; // (1-based index + ヘッダー行1)
                      // (XSS対策のため、< > & " ' をエスケープ)
                      const escapedComment = String(comment).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[m]);
                      
                      bodyHtml += `<textarea 
                          class="comment-edit-textarea" 
                          data-type="L" 
                          data-col="${group.col}" 
                          data-row="${sheetRowIndex}"
                          data-original-value="${escapedComment}"
                      >${escapedComment}</textarea>`;
                  });
                  bodyHtml += '</div>';
                  
                  slidesData.push({ 
                      header: slideHeader, 
                      body: bodyHtml, 
                      isEditable: true 
                  });
              }
          }
      }
      // 2. NPS以外の場合 (commentData は [ "c1", "c2" ] 配列)
      else {
          const typeMap = { 'feedback_i': 'I', 'feedback_j': 'J', 'feedback_m': 'M' };
          const titleMap = { 'feedback_i': '良かった点や悪かった点など', 'feedback_j': '印象に残ったスタッフへのコメント', 'feedback_m': 'お産にかかわるご意見・ご感想' };
          
          baseTitle = `アンケート結果　ー${titleMap[reportType]}ー`;
          document.getElementById('report-title').textContent = baseTitle;
          document.getElementById('report-subtitle').textContent = 'データ一覧（20データずつ）';

          const comments = commentData; // (commentData が配列)
          const commentType = typeMap[reportType];
          
          if (comments.length === 0) return;

          const totalPages = Math.ceil(comments.length / MAX_PER_PAGE);
          
          for (let i = 0; i < comments.length; i += MAX_PER_PAGE) {
              const chunk = comments.slice(i, i + MAX_PER_PAGE);
              const pageNum = (i / MAX_PER_PAGE) + 1;
              const slideHeader = `( ${pageNum} / ${totalPages} )`;
              
              let bodyHtml = '<div>';
              chunk.forEach((comment, index) => {
                  const sheetRowIndex = i + index + 2; // (1-based index + ヘッダー行1)
                  const escapedComment = String(comment).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[m]);

                  bodyHtml += `<textarea 
                      class="comment-edit-textarea" 
                      data-type="${commentType}" 
                      data-col="A" 
                      data-row="${sheetRowIndex}"
                      data-original-value="${escapedComment}"
                  >${escapedComment}</textarea>`;
              });
              bodyHtml += '</div>';
              
              slidesData.push({ 
                  header: slideHeader, 
                  body: bodyHtml, 
                  isEditable: true 
              });
          }
      }
      
      // ページネーションを表示
      if (slidesData.length > 0) {
          document.getElementById('pagination').style.display = 'flex';
      }
  }

  // --- ▼▼▼ [新規] コメント保存処理 ▼▼▼ ---
  async function handleSaveComment() {
      const textareas = document.querySelectorAll('#slide-body .comment-edit-textarea');
      if (textareas.length === 0) {
          alert('保存対象のコメントがありません。');
          return;
      }
      
      showLoading(true, '変更を保存中...');
      
      const promises = [];
      const changedTextareas = [];

      textareas.forEach(ta => {
          const newValue = ta.value;
          const originalValue = ta.dataset.originalValue;

          if (newValue !== originalValue) {
              const payload = {
                  centralSheetId: currentCentralSheetId,
                  clinicName: currentClinicForModal,
                  commentType: ta.dataset.type,
                  col: ta.dataset.col,
                  row: parseInt(ta.dataset.row),
                  value: newValue
              };
              
              console.log('Saving change:', payload);
              promises.push(fetch('/api/updateCommentData', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(payload)
              }));
              changedTextareas.push(ta); // 成功後に data-original-value を更新するため
          }
      });

      if (promises.length === 0) {
          showLoading(false);
          alert('変更されたコメントはありません。');
          return;
      }

      try {
          const results = await Promise.all(promises);
          
          let failedCount = 0;
          results.forEach(async (res, index) => {
              if (res.ok) {
                  // 成功した場合、DOMの-original-valueを更新
                  changedTextareas[index].dataset.originalValue = changedTextareas[index].value;
              } else {
                  failedCount++;
                  console.error('Save failed:', await res.text());
              }
          });

          if (failedCount > 0) {
              alert(`${failedCount}件の保存に失敗しました。`);
          } else {
              alert(`${promises.length}件の変更を保存しました。`);
          }

      } catch (err) {
          console.error('Error saving comments:', err);
          alert(`保存中にエラーが発生しました: ${err.message}`);
      } finally {
          showLoading(false);
      }
  }


  // --- ▼▼▼ [修正] スライド描画ヘルパー (保存ボタンの表示/非表示) ▼▼▼ ---
  function renderSlide(page){ 
    if(page<1||page>slidesData.length)return;
    const s=slidesData[page-1];
    
    // (コメントスライドのタイトル/サブタイトルは buildCommentSlides で設定済み)
    if (!s.isEditable) {
        document.getElementById('report-title').textContent = s.title || '';
        document.getElementById('report-subtitle').textContent = s.subtitle || '';
    }
    
    document.getElementById('slide-header').innerHTML=s.header;
    document.getElementById('slide-header').classList.toggle('hidden', !s.header);
    
    const bodyEl = document.getElementById('slide-body');
    bodyEl.innerHTML = s.body; 

    document.getElementById('page-indicator').textContent=`${page} / ${slidesData.length}`;
    document.getElementById('prev-slide').disabled=(page===1);
    document.getElementById('next-slide').disabled=(page===slidesData.length); 
    
    // [修正] 保存ボタンの表示/非表示
    commentSaveBtn.style.display = s.isEditable ? 'block' : 'none';
    document.getElementById('pagination').style.display = 'flex'; // ページネーション自体は表示
  }
  function handlePrevSlide(){ if(currentPage>1){currentPage--;renderSlide(currentPage);} }
  function handleNextSlide(){ if(currentPage<slidesData.length){currentPage++;renderSlide(currentPage);} }


  // 市区町村 (表のレイアウト調整)
  async function prepareAndShowMunicipalityReport() {
    console.log('Prepare municipality report');
    updateNavActiveState('municipality', null, null);
    showScreen('screen3');
    document.getElementById('report-title').textContent = `アンケート結果　ーご回答者さまの市町村ー`;
    document.getElementById('report-subtitle').textContent = 'ご出産された方の住所（市町村）について教えてください。';
    document.getElementById('report-subtitle').style.textAlign = 'center'; 
    document.getElementById('report-separator').style.display='block';
    
    const slideBody = document.getElementById('slide-body');
    slideBody.style.whiteSpace = 'normal';
    slideBody.innerHTML = '<p class="text-center text-gray-500 py-16">市区町村データを読み込み中...</p>';
    slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
    
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
    // ▼▼▼ [修正] 表のフォントサイズ/パディングを調整して収める ▼▼▼
    let tableHtml = `<div class="municipality-table-container w-full h-full border border-gray-200 rounded-lg"><table class="w-full divide-y divide-gray-200"><thead class="bg-gray-50 sticky top-0 z-10"><tr><th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">都道府県</th><th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">市区町村</th><th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">件数</th><th class="py-3 text-left font-medium text-gray-500 uppercase tracking-wider">割合</th></tr></thead><tbody class="bg-white divide-y divide-gray-200">`; 
    data.forEach(row => { 
        tableHtml += `<tr><td class="py-2 font-medium text-gray-900">${row.prefecture}</td><td class="py-2 text-gray-700">${row.municipality}</td><td class="py-2 text-gray-700 text-right">${row.count}</td><td class="py-2 text-gray-700 text-right">${row.percentage.toFixed(2)}%</td></tr>`; 
    }); 
    tableHtml += '</tbody></table></div>'; 
    slideBody.innerHTML = tableHtml; 
  }


  // おすすめ理由 (フォントサイズ修正)
  async function prepareAndShowRecommendationReport() {
      console.log('Prepare recommendation report');
      updateNavActiveState('recommendation', null, null);
      showScreen('screen3');
      document.getElementById('report-title').textContent = 'アンケート結果　ー本病院を選ぶ上で最も参考にしたものー';
      document.getElementById('report-subtitle').textContent = 'ご出産された産婦人科医院への本病院を選ぶ上で最も参考にしたものについて、教えてください。';
      document.getElementById('report-subtitle').style.textAlign = 'center';
      document.getElementById('report-separator').style.display='block';
      
      const slideBody = document.getElementById('slide-body');
      slideBody.style.whiteSpace = 'normal';
      slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');

      // (円グラフ2つのシェル)
      slideBody.innerHTML = `<div class="grid grid-cols-1 md:grid-cols-2 gap-8 h-full"><div class="flex flex-col items-center"><h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3><div id="clinic-pie-chart" class="w-full h-[400px]"></div></div><div class="flex flex-col items-center"><h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3><div id="average-pie-chart" class="w-full h-[400px]"></div></div></div>`;
      
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
          
          // --- ▼▼▼ [修正] フォントサイズを 12pt に修正 ▼▼▼ ---
          const opt = {is3D: true,chartArea: { left: '5%', top: '5%', width: '90%', height: '90%' },pieSliceText: 'percentage',pieSliceTextStyle: { color: 'black', fontSize: 12, bold: true },legend: { position: 'labeled', textStyle: { color: 'black', fontSize: 12 } },tooltip: { showColorCode: true, textStyle: { fontSize: 12 }, trigger: 'focus' },};
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


  // Word Cloud表示 (Screen 3) (変更なし)
  async function prepareAndShowAnalysis(columnType) {
    showLoading(true, `テキスト分析中(${getColumnName(columnType)})...`);
    showScreen('screen3');
    clearAnalysisCharts();
    updateNavActiveState(null, columnType, null);
    
    let tl = [], td = 0;
    
    document.getElementById('report-title').textContent = getAnalysisTitle(columnType, 0); 
    document.getElementById('report-subtitle').textContent = '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。\n通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。';
    document.getElementById('report-subtitle').style.textAlign = 'left'; // WCページのみ左詰
    document.getElementById('report-separator').style.display='block';

    const slideBody = document.getElementById('slide-body');
    slideBody.classList.remove('flex', 'items-center', 'justify-center', 'items-start', 'justify-start');
    
    try {
        // (getChartDataはキャッシュを使うので、WC表示のための元テキスト取得にも流用する)
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
    
    // WCの新しいシェルを直接 body に設定
    slideBody.innerHTML = `
        <div class="grid grid-cols-2 gap-4 h-full">
            <div class="space-y-2 h-full pr-2 flex flex-col">
                <div id="noun-chart-container" class="chart-container h-1/4">
                    <h3 class="font-bold text-center mb-2 text-blue-600">名詞</h3>
                    <div id="noun-chart" class="w-full h-full"></div>
                </div>
                <div id="verb-chart-container" class="chart-container h-1/4">
                    <h3 class="font-bold text-center mb-2 text-red-600">動詞</h3>
                    <div id="verb-chart" class="w-full h-full"></div>
                </div>
                <div id="adj-chart-container" class="chart-container h-1/4">
                    <h3 class="font-bold text-center mb-2 text-green-600">形容詞</h3>
                    <div id="adj-chart" class="w-full h-full"></div>
                </div>
                <div id="int-chart-container" class="chart-container h-1/4">
                    <h3 class="font-bold text-center mb-2 text-gray-600">感動詞</h3>
                    <div id="int-chart" class="w-full h-full"></div>
                </div>
            </div>
            <div class="space-y-4 flex flex-col h-full">
                <p class="text-sm text-gray-600">
                    スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。<br>
                    単語の色は品詞の種類で異なります。<br>
                    <span class="text-blue-600 font-semibold">青色=名詞</span>、
                    <span class="text-red-600 font-semibold">赤色=動詞</span>、
                    <span class="text-green-600 font-semibold">緑色=形容詞</span>、
                    <span class="text-gray-600 font-semibold">灰色=感動詞</span>
                </p>
                <div id="word-cloud-container" class="h-[calc(100%-80px)]">
                    <canvas id="word-cloud-canvas" class="!h-full !w-full"></canvas>
                </div>
                <div id="analysis-error" class="text-red-500 text-sm text-center hidden"></div>
            </div>
        </div>
    `;

    try{
        console.log(`Sending ${tl.length} texts to Kuromoji...`);
        const r=await fetch('/api/analyzeText',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({textList:tl})});
        if(!r.ok){const et=await r.text();throw new Error(`分析APIエラー(${r.status}): ${et}`);}
        const ad=await r.json();
        console.log("Received Kuromoji analysis:",ad);
        setTimeout(() => drawAnalysisCharts(ad.results), 50);
    } catch(error){
        console.error('!!! Analyze fail:',error);
        document.getElementById('analysis-error').textContent=`分析失敗: ${error.message}`;
        document.getElementById('analysis-error').classList.remove('hidden');
    } finally {
        showLoading(false);
    }
  }

  // --- ▼▼▼ [修正] Word Cloud描画 (フォントサイズ 12pt に修正) ▼▼▼ ---
  function drawAnalysisCharts(results) { if(!results||results.length===0){console.log("No analysis results.");document.getElementById('analysis-error').textContent='分析結果なし';document.getElementById('analysis-error').classList.remove('hidden');return;} const nouns=results.filter(r=>r.pos==='名詞'),verbs=results.filter(r=>r.pos==='動詞'),adjs=results.filter(r=>r.pos==='形容詞'),ints=results.filter(r=>r.pos==='感動詞');const barOpt={bars:'horizontal',legend:{position:'none'},hAxis:{title:'スコア(出現頻度)',minValue:0, textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}},vAxis:{title:'単語', textStyle:{fontSize:12}, titleTextStyle:{fontSize:12}}, chartArea:{height:'90%', width:'80%', left:'15%', top:'5%'}};drawSingleBarChart(nouns.slice(0,20),'noun-chart',{...barOpt,colors:['#3b82f6']});drawSingleBarChart(verbs.slice(0,20),'verb-chart',{...barOpt,colors:['#ef4444']});drawSingleBarChart(adjs.slice(0,20),'adj-chart',{...barOpt,colors:['#22c55e']});drawSingleBarChart(ints.slice(0,20),'int-chart',{...barOpt,colors:['#6b7280']});const wl=results.map(r=>[r.word,r.score]).slice(0,100);const pm=results.reduce((map,item)=>{map[item.word]=item.pos;return map;},{});console.log('WordCloud List:',wl);const cv=document.getElementById('word-cloud-canvas');if(WordCloud.isSupported&&cv){try{const options={list:wl,gridSize:Math.round(16*cv.width/1024),weightFactor:function(s){return Math.pow(s,0.8)*cv.width/250;},fontFamily:'Noto Sans JP,sans-serif',color:function(w,wt,fs,d,t){const p=pm[w]||'不明';switch(p){case'名詞':return'#3b82f6';case'動詞':return'#ef4444';case'形容詞':return'#22c55e';case'感動詞':return'#6b7280';default:return'#a8a29e';}},backgroundColor:'#f9fafb',clearCanvas:true};console.log('WordCloud Options:',options);WordCloud(cv,options);console.log("Word cloud drawn.");}catch(wcError){console.error("Error drawing WordCloud:",wcError);document.getElementById('word-cloud-container').innerHTML=`<p class="text-center text-red-500">ワードクラウド描画エラー:${wcError.message}</p>`;}}else{console.warn("WordCloud unsupported/canvas missing.");document.getElementById('word-cloud-container').innerHTML='<p class="text-center text-gray-500">ワードクラウド非対応</p>';} }
  function drawSingleBarChart(data, elementId, options) { const c=document.getElementById(elementId);if(!c){console.error(`Element not found: ${elementId}`);return;} if(!data||data.length===0){c.innerHTML='<p class="text-center text-gray-500 text-sm py-4">データなし</p>';return;} const cd=[['単語','スコア',{role:'style'}]];const color=options.colors&&options.colors.length>0?options.colors[0]:'#a8a29e';data.slice().reverse().forEach(item=>{cd.push([item.word,item.score,color]);});try{const dt=google.visualization.arrayToDataTable(cd);const chart=new google.visualization.BarChart(c);chart.draw(dt,options);}catch(chartError){console.error(`Error drawing bar chart for ${elementId}:`,chartError);c.innerHTML=`<p class="text-center text-red-500 text-sm py-4">グラフ描画エラー<br>${chartError.message}</p>`;} }
  function clearAnalysisCharts() { document.getElementById('noun-chart').innerHTML='';document.getElementById('verb-chart').innerHTML='';document.getElementById('adj-chart').innerHTML='';document.getElementById('int-chart').innerHTML='';const c=document.getElementById('word-cloud-canvas');const x=c.getContext('2d');x.clearRect(0,0,c.width,c.height);document.getElementById('analysis-error').classList.add('hidden');document.getElementById('analysis-error').textContent=''; }
  function getAnalysisTitle(columnType, count) { const bt=`アンケート結果　ー${getColumnName(columnType)}ー`;return`${bt}　※全回答数${count}件ー`; }
  // (getColumnName は pdfGenerator.js にしか無いので、ブラウザ側でも定義)
  function getColumnName(columnType) { 
    switch(columnType){
        case'L':return'NPS推奨度 理由';
        case'I':case'I_good':case'I_bad':return'良かった点や悪かった点など';
        case'J':return'印象に残ったスタッフへのコメント';
        case'M':return'お産にかかわるご意見・ご感想';
        default:return'不明';
    } 
  }


  // --- ▼▼▼ [修正] AI詳細分析 (Screen 5) 処理 (サブタイトル中央揃えを削除) ▼▼▼
  async function prepareAndShowDetailedAnalysis(analysisType) {
    console.log(`Prepare detailed analysis: ${analysisType}`);
    const clinicName = currentClinicForModal;
    
    showLoading(true, `AI分析結果を読み込み中...\n(${getDetailedAnalysisTitleBase(analysisType)})`);
    showScreen('screen5');
    updateNavActiveState(null, null, analysisType);
    toggleEditDetailedAnalysis(false); 
    
    const errorDiv = document.getElementById('detailed-analysis-error');
    errorDiv.classList.add('hidden');
    errorDiv.textContent = '';
    
    // AI分析のタイトル/サブタイトルを、新しい固定フォーマットに準拠させる
    document.getElementById('detailed-analysis-title').textContent = getDetailedAnalysisTitleFull(analysisType);
    document.getElementById('detailed-analysis-subtitle').textContent = getDetailedAnalysisSubtitleForUI(analysisType, 'analysis'); 
    // ▼▼▼ [修正] 中央揃えの指示を削除 ▼▼▼
    // document.getElementById('detailed-analysis-subtitle').style.textAlign = 'center'; 
    
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
      console.log(`[AI] Regenerating analysis for ${currentDetailedAnalysisType}...`);
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
          console.log("[AI] Generation successful, raw JSON:", analysisJson);
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

  // (AI表示 (共通) - 変更なし)
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
      
      document.getElementById('display-analysis').textContent = analysisText;
      document.getElementById('textarea-analysis').value = analysisText;
      document.getElementById('display-suggestions').textContent = suggestionsText;
      document.getElementById('textarea-suggestions').value = suggestionsText;
      document.getElementById('display-overall').textContent = overallText;
      document.getElementById('textarea-overall').value = overallText;
      
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

  // (タブ切り替え - サブタイトル更新ロジックを追加)
  function handleTabClick(event) { 
    const tabId = event.target.dataset.tabId; 
    if (tabId) { 
        switchTab(tabId); 
    } 
  }
  function switchTab(tabId) { 
      // サブタイトルを更新
      document.getElementById('detailed-analysis-subtitle').textContent = getDetailedAnalysisSubtitleForUI(currentDetailedAnalysisType, tabId);
      
      if (isEditingDetailedAnalysis) { 
          // (編集中のタブ切り替えは許可)
      } 
      document.querySelectorAll('#ai-tab-nav .tab-button').forEach(button => { if (button.dataset.tabId === tabId) { button.classList.add('active'); } else { button.classList.remove('active'); } }); 
      document.querySelectorAll('#screen5 .tab-content').forEach(content => { if (content.id === `content-${tabId}`) { content.classList.remove('hidden'); } else { content.classList.add('hidden'); } }); 
      if (isEditingDetailedAnalysis) {
          const activeTextarea = document.getElementById(`textarea-${tabId}`);
          if (activeTextarea) {
              activeTextarea.style.height = 'auto';
              activeTextarea.style.height = (activeTextarea.scrollHeight + 5) + 'px';
          }
      }
  }

  // (編集モード切り替え - 変更なし)
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
      }
  }


  // (AI詳細分析 (保存) - 変更なし)
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

  // ▼▼▼ [修正] AIタイトル・サブタイトル取得関数 (サブタイトルルールを反映) ▼▼▼
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
  
  function getDetailedAnalysisSubtitleForUI(analysisType, tabId) {
       const base = '※コメントでいただいたフィードバックを元に分析しています';
       const getBodyText = (type) => {
           switch (type) {
               case 'L': 
               case 'M': return '患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。';
               case 'I_bad': return 'フィードバックの中で挙げられた「悪かった点」を分析すると、\n患者にとって以下の要素が特に課題として感じられていることが分かります。';
               case 'I_good': return 'フィードバックの中で挙げられた「良かった点」を分析すると、\n以下の要素が患者にとって特に高く評価されていることが分かります。';
               case 'J': return '印象に残ったスタッフに対するコメントから、いくつかの重要なテーマが浮かび上がります。\nこれらのテーマは、スタッフの評価においても重要なポイントとなります。';
               default: return '';
           }
       };

       if (tabId === 'overall') {
           return base + '\n患者から寄せられたお産に関するご意見の分析と改善策を基にした、総評は以下のとおりです.';
       }
       // analysis と suggestions は同じサブタイトルを使用
       return base + '\n' + getBodyText(analysisType);
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
  // ▲▲▲


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
