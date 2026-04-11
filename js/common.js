/* ===================================
   共通ナビゲーション・ヘッダー挿入
   ファイル: assets/js/common.js
   =================================== */

(function() {
  // ページタイトルのマッピング
  var pageTitles = {
    'tutorial_haq_direct.html':       'HAQ-DI 直接エビデンス評価',
    'tutorial_indirect_acr50.html':   'ACR50 間接・NMAエビデンス評価',
    'tutorial_sae_direct.html':       'SAE 評価者間不一致の実例',
    'index.html':                     'ホーム'
  };

  // 現在のページ名を取得
  var currentFile = window.location.pathname.split('/').pop() || 'index.html';
  var currentTitle = pageTitles[currentFile] || 'チュートリアル';

  // ナビゲーションHTMLを生成
  var navHTML = '<nav class="site-nav">' +
    '<a href="index.html">🏠 ホーム</a>' +
    '<span style="opacity:0.6;">›</span>' +
    '<span id="pageTitle">' + currentTitle + '</span>' +
    '</nav>';

  // bodyの先頭に挿入
  document.body.insertAdjacentHTML('afterbegin', navHTML);
})();
