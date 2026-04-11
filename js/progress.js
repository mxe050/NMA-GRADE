/* ===================================
   学習進捗トラッキング
   ファイル: assets/js/progress.js
   =================================== */

var NMAProgress = {
  // 完了済みチュートリアルをlocalStorageから取得
  getCompleted: function() {
    try {
      return JSON.parse(localStorage.getItem('nmaGradeProgress') || '[]');
    } catch(e) {
      return [];
    }
  },

  // チュートリアルを完了済みとしてマーク
  markCompleted: function(tutorialId) {
    var completed = this.getCompleted();
    if (!completed.includes(tutorialId)) {
      completed.push(tutorialId);
      try {
        localStorage.setItem('nmaGradeProgress', JSON.stringify(completed));
      } catch(e) {
        console.warn('localStorage not available');
      }
    }
    this.updateCards();
  },

  // index.htmlのカードに完了マークを反映
  updateCards: function() {
    var completed = this.getCompleted();
    document.querySelectorAll('.tutorial-card[data-tutorial-id]').forEach(function(card) {
      var id = card.getAttribute('data-tutorial-id');
      if (completed.includes(id)) {
        card.classList.add('completed');
        var mark = card.querySelector('.completed-mark');
        if (mark) mark.textContent = '✅';
        var btn = card.querySelector('.start-btn');
        if (btn) btn.textContent = '復習する →';
      }
    });

    // 全体の進捗率を表示
    var total = document.querySelectorAll('.tutorial-card[data-tutorial-id]').length;
    if (total > 0) {
      var pct = Math.round(completed.length / total * 100);
      var bar = document.getElementById('overallProgressFill');
      var label = document.getElementById('overallProgressLabel');
      if (bar) bar.style.width = pct + '%';
      if (label) label.textContent = completed.length + ' / ' + total + ' 完了（' + pct + '%）';
    }
  }
};

// ページ読み込み時にカードを更新（index.html用）
document.addEventListener('DOMContentLoaded', function() {
  NMAProgress.updateCards();
});
