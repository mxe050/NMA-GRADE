# =============================================================================
# NMA-GRADE エビデンス確実性評価サポートアプリ v1.0
# 単一ファイル版 - app.R
# BMJ 2023;381:e074495 (Izcovich et al.) 準拠
# Phillips et al. (BMJ Open 2022) 色分け図対応
# =============================================================================

cat("パッケージ読込中...\n")

required_packages <- c(
  "shiny", "shinydashboard", "shinyWidgets", "shinyjs",
  "netmeta", "meta", "DT", "ggplot2", "dplyr", "tidyr",
  "openxlsx", "jsonlite"
)

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    cat(paste0("  インストール中: ", pkg, "\n"))
    install.packages(pkg, repos = "https://cran.r-project.org")
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

if (!requireNamespace("NMA", quietly = TRUE)) {
  cat("  NMAパッケージをインストール中...\n")
  install.packages("NMA", repos = "https://cran.r-project.org")
}
suppressPackageStartupMessages(library(NMA))

cat("全パッケージ読込完了\n")

# =============================================================================
# 定数
# =============================================================================

CERT_LEVELS <- c("Very Low" = 1L, "Low" = 2L, "Moderate" = 3L, "High" = 4L)
GRADE3 <- c("Not serious", "Serious", "Very serious")
GRADE4 <- c("Not assessed", "Not serious", "Serious", "Very serious", "Extremely serious")
PUB_CHOICES <- c("Undetected", "Strongly suspected")

# =============================================================================
# ヘルパー関数
# =============================================================================

cert2num <- function(x) {
  if (is.null(x) || is.na(x) || x == "" || x == "NA") return(NA_integer_)
  val <- CERT_LEVELS[x]; if (is.na(val)) return(NA_integer_); unname(val)
}
num2cert <- function(n) {
  if (is.na(n)) return(NA_character_)
  n <- max(1L, min(4L, as.integer(n)))
  names(CERT_LEVELS)[match(n, CERT_LEVELS)]
}
dg_amount <- function(j) {
  if (is.null(j) || is.na(j)) return(0L)
  switch(as.character(j),
    "Not serious"=0L, "Not assessed"=0L, "Undetected"=0L,
    "Serious"=1L, "Strongly suspected"=1L,
    "Very serious"=2L, "Extremely serious"=3L, 0L)
}
calc_a1_preliminary <- function(rob, inc, ind, pub) {
  lv <- 4L - dg_amount(rob) - dg_amount(inc) - dg_amount(ind) - dg_amount(pub)
  num2cert(lv)
}
calc_downgrade <- function(start_cert, ...) {
  dots <- list(...)
  lv <- cert2num(start_cert)
  if (is.na(lv)) return(NA_character_)
  for (d in dots) lv <- lv - dg_amount(d)
  num2cert(lv)
}

auto_imprecision <- function(lower, upper, mid_lower, mid_upper, is_ratio = TRUE, effect_measure = "OR") {
  if (is.na(lower) || is.na(upper))
    return(list(result = "Not assessed", reason = "95%CIが利用不可"))
  log_null <- 0
  if (is_ratio) { log_mid_lo <- log(mid_lower); log_mid_up <- log(mid_upper) }
  else { log_mid_lo <- mid_lower; log_mid_up <- mid_upper }
  crosses_null   <- (lower <= log_null & upper >= log_null)
  crosses_mid_lo <- (lower <= log_mid_lo & upper >= log_mid_lo)
  crosses_mid_up <- (lower <= log_mid_up & upper >= log_mid_up)
  n_cross <- sum(c(crosses_mid_lo, crosses_mid_up), na.rm = TRUE)
  ci_w <- upper - lower; mid_w <- log_mid_up - log_mid_lo
  # OIS簡易判定: CI上限/下限の比率 (GRADE guidelines 33)
  ois_note <- ""
  if (is_ratio && upper > lower) {
    ci_ratio <- exp(upper) / exp(lower)
    ois_thresh <- if (effect_measure == "OR") 2.5 else 3.0
    ois_note <- sprintf(" [CI ratio=%.1f", ci_ratio)
    if (ci_ratio > ois_thresh) ois_note <- paste0(ois_note, sprintf(", >%.1f: OIS未達の可能性]", ois_thresh))
    else ois_note <- paste0(ois_note, "]")
  }
  # Step 1: CIが閾値を跨ぐか
  if (n_cross == 0 && !crosses_null) {
    # Step 2: CIが閾値を跨がない → 効果量とOISを確認
    te_mid <- (lower + upper) / 2
    large_effect <- if (is_ratio) abs(te_mid) > log(2) else FALSE
    if (large_effect && is_ratio) {
      ci_ratio <- exp(upper) / exp(lower)
      ois_thresh <- if (effect_measure == "OR") 2.5 else 3.0
      if (ci_ratio > ois_thresh)
        return(list(result="Serious", reason=paste0("CIは閾値内だが効果量が大きくOIS未達", ois_note)))
    }
    list(result="Not serious", reason=paste0("CIがnull/閾値を跨がない", ois_note))
  }
  else if (mid_w > 0 && ci_w > mid_w * 3)
    list(result="Extremely serious", reason=paste0("CI幅が閾値幅の3倍超（3段階格下げ相当）", ois_note))
  else if (n_cross >= 2)
    list(result="Very serious", reason=paste0("CIが閾値の両側を跨ぐ（2段階格下げ）", ois_note))
  else
    list(result="Serious", reason=paste0("CIがnullまたは閾値片側を跨ぐ（1段階格下げ）", ois_note))
}

auto_c1 <- function(row) {
  has_d <- isTRUE(row$has_direct) && !is.na(row$direct_TE)
  has_i <- !is.na(row$indirect_TE)
  if (has_d && !has_i) return(list(dom="直接のみ", src="Direct", cert=row$a1_preliminary, reason="間接なし→直接"))
  if (!has_d && has_i) return(list(dom="間接のみ", src="Indirect", cert=row$b2_preliminary, reason="直接なし→間接"))
  if (!has_d && !has_i) return(list(dom="なし", src=NA_character_, cert=NA_character_, reason="両方なし"))
  d_w <- row$direct_upper - row$direct_lower; i_w <- row$indirect_upper - row$indirect_lower
  if (is.na(d_w)||is.na(i_w)||i_w==0) return(list(dom="直接",src="Direct",cert=row$a1_preliminary,reason="CI幅不能→直接"))
  ratio <- d_w/i_w; d_n <- cert2num(row$a1_preliminary); i_n <- cert2num(row$b2_preliminary)
  if (is.na(d_n)) d_n <- 0L; if (is.na(i_n)) i_n <- 0L
  if (ratio < 0.5) list(dom="直接支配",src="Direct",cert=row$a1_preliminary,reason=paste0("ratio=",sprintf("%.2f",ratio),"→直接"))
  else if (ratio > 2.0) list(dom="間接支配",src="Indirect",cert=row$b2_preliminary,reason=paste0("ratio=",sprintf("%.2f",ratio),"→間接"))
  else if (d_n >= i_n) list(dom="同程度",src="Direct",cert=row$a1_preliminary,reason=paste0("ratio=",sprintf("%.2f",ratio),"→直接"))
  else list(dom="同程度",src="Indirect",cert=row$b2_preliminary,reason=paste0("ratio=",sprintf("%.2f",ratio),"→間接"))
}

auto_c2 <- function(row) {
  p <- row$netsplit_p; d <- row$direct_TE; i <- row$indirect_TE
  if (is.na(p)||is.na(d)||is.na(i)) return(list(result="Not serious",reason="p値/推定値欠損"))
  same <- (sign(d)==sign(i))
  if (p>=0.10) list(result="Not serious",reason=paste0("p=",sprintf("%.3f",p),"≥0.10"))
  else if (p>=0.05) list(result="Not serious",reason=paste0("p=",sprintf("%.3f",p)," 境界的"))
  else if (!same) list(result="Very serious",reason=paste0("p=",sprintf("%.3f",p)," 方向不一致"))
  else list(result="Serious",reason=paste0("p=",sprintf("%.3f",p)," 方向一致"))
}

auto_d <- function(row) {
  inco <- row$c2_incoherence %in% c("Serious","Very serious")
  if (!inco) return(list(best="NMA",cert=row$c3_final_nma,needs_a2b3=FALSE,reason="非整合性なし→NMA"))
  d_c <- cert2num(row$a2_final_direct); i_c <- cert2num(row$b3_final_indirect); n_c <- cert2num(row$c3_final_nma)
  certs <- c(Direct=d_c,Indirect=i_c,NMA=n_c); cv <- certs[!is.na(certs)]
  if (length(cv)==0) return(list(best="NMA",cert=row$c3_final_nma,needs_a2b3=TRUE,reason="A2/B3未評価→暫定NMA"))
  bn <- names(which.max(cv)); bc <- num2cert(max(cv))
  list(best=bn,cert=bc,needs_a2b3=TRUE,reason=paste0(bn,"(",bc,")を採用"))
}

classify_color <- function(cert, te, lower, upper, is_ratio, small_good, large_thresh) {
  if (is.na(cert)||is.na(te)) return("uncertain")
  hm <- cert %in% c("High","Moderate"); null_v <- 0
  sig <- !(lower<=null_v & upper>=null_v)
  beneficial <- if(small_good) te<null_v else te>null_v
  large_ben <- if(is_ratio) { if(small_good) te<log(large_thresh)&sig else te>log(1/large_thresh)&sig }
               else abs(te)>abs(large_thresh)&sig&beneficial
  sfx <- if(hm) "_hm" else "_lv"
  if(large_ben) paste0("among_best",sfx) else if(beneficial&sig) paste0("intermediate",sfx) else paste0("among_worst",sfx)
}

gradient_fill <- function(cat) switch(cat, "among_best_hm"="#006400","intermediate_hm"="#228B22","among_worst_hm"="#C8E6C9","among_best_lv"="#E0E0E0","intermediate_lv"="#BDBDBD","among_worst_lv"="#757575","uncertain"="#FFFFFF","#FFFFFF")
stoplight_fill <- function(cat) switch(cat, "among_best_hm"="#28a745","intermediate_hm"="#ffc107","among_worst_hm"="#dc3545","among_best_lv"="#E0E0E0","intermediate_lv"="#BDBDBD","among_worst_lv"="#757575","uncertain"="#FFFFFF","#FFFFFF")
text_col_for <- function(bg) if(bg %in% c("#006400","#757575","#dc3545","#228B22")) "white" else "black"

cert_badge_html <- function(level) {
  if (is.null(level)||is.na(level)||level==""||level=="NA") return("<span style='color:#999;'>---</span>")
  cols <- list("High"=c("#006400","#fff"),"Moderate"=c("#4169E1","#fff"),"Low"=c("#DAA520","#000"),"Very Low"=c("#DC143C","#fff"))
  syms <- list("High"=paste0(rep("\u2295",4),collapse=""),"Moderate"=paste0(rep("\u2295",3),"\u25cb",collapse=""),"Low"=paste0(rep("\u2295",2),rep("\u25cb",2),collapse=""),"Very Low"=paste0("\u2295",rep("\u25cb",3),collapse=""))
  cc <- cols[[level]]; if(is.null(cc)) return(paste0("<span>",level,"</span>"))
  sym <- syms[[level]]; if(is.null(sym)) sym <- ""
  paste0('<span style="background:',cc[1],';color:',cc[2],';padding:2px 8px;border-radius:10px;font-weight:600;font-size:12px;">',level,' ',sym,'</span>')
}

find_column <- function(df, patterns) {
  for(p in patterns){ m<-grep(p,tolower(names(df))); if(length(m)>0) return(names(df)[m[1]]) }; ""
}

# =============================================================================
# build_grade_table
# =============================================================================

build_grade_table <- function(nma_obj, pw_data) {
  trts <- sort(nma_obj$trts)
  pairs <- combn(trts, 2, simplify = FALSE)
  n <- length(pairs)

  gt <- data.frame(
    comparison = sapply(pairs, function(x) paste(x, collapse = ":")),
    treat1 = sapply(pairs, `[`, 1),
    treat2 = sapply(pairs, `[`, 2),
    stringsAsFactors = FALSE)

  gt$nma_TE <- gt$nma_seTE <- gt$nma_lower <- gt$nma_upper <- NA_real_
  for (i in seq_len(n)) {
    t1 <- gt$treat1[i]; t2 <- gt$treat2[i]
    gt$nma_TE[i]    <- nma_obj$TE.random[t1, t2]
    gt$nma_seTE[i]  <- nma_obj$seTE.random[t1, t2]
    gt$nma_lower[i] <- nma_obj$lower.random[t1, t2]
    gt$nma_upper[i] <- nma_obj$upper.random[t1, t2]
  }

  gt$direct_k <- rep(0L, n)
  gt$direct_TE <- gt$direct_seTE <- gt$direct_lower <- gt$direct_upper <- NA_real_
  gt$indirect_TE <- gt$indirect_seTE <- gt$indirect_lower <- gt$indirect_upper <- NA_real_
  gt$netsplit_p <- NA_real_

  # pairwise metagen() による直接推定値を別途保持
  gt$pw_direct_TE <- gt$pw_direct_lower <- gt$pw_direct_upper <- NA_real_
  # netsplit Direct 推定値を別途保持
  gt$ns_direct_TE <- gt$ns_direct_lower <- gt$ns_direct_upper <- NA_real_

  pw <- pw_data
  pw$t1c <- as.character(pw$treat1)
  pw$t2c <- as.character(pw$treat2)
  pw$comp_key <- apply(pw[, c("t1c","t2c")], 1, function(r) paste(sort(r), collapse=":"))

  for (i in seq_len(n)) {
    comp <- gt$comparison[i]
    sub <- pw[pw$comp_key == comp, ]
    k <- nrow(sub)
    gt$direct_k[i] <- as.integer(k)

    if (k >= 1) {
      tryCatch({
        m <- metagen(TE = sub$TE, seTE = sub$seTE, studlab = sub$studlab,
                     sm = nma_obj$sm, random = TRUE, common = TRUE)
        gt$direct_TE[i]    <- m$TE.random
        gt$direct_seTE[i]  <- m$seTE.random
        gt$direct_lower[i] <- m$lower.random
        gt$direct_upper[i] <- m$upper.random

        if (nrow(sub) > 0) {
          pw_t1 <- as.character(sub$treat1[1])
          pw_t2 <- as.character(sub$treat2[1])
          gt_t1 <- gt$treat1[i]; gt_t2 <- gt$treat2[i]
          if (pw_t1 == gt_t2 && pw_t2 == gt_t1) {
            gt$direct_TE[i]    <- -gt$direct_TE[i]
            tmp_lo <- gt$direct_lower[i]
            gt$direct_lower[i] <- -gt$direct_upper[i]
            gt$direct_upper[i] <- -tmp_lo
          }
        }

        # pairwise の値を保存
        gt$pw_direct_TE[i]    <- gt$direct_TE[i]
        gt$pw_direct_lower[i] <- gt$direct_lower[i]
        gt$pw_direct_upper[i] <- gt$direct_upper[i]

      }, error = function(e) {
        cat("metagen error for", comp, ":", e$message, "\n")
      })
    }
  }

  gt$has_direct <- gt$direct_k > 0

  gt$a1_rob <- rep("Not serious", n)
  gt$a1_inconsistency <- rep("Not serious", n)
  gt$a1_indirectness <- rep("Not serious", n)
  gt$a1_pub_bias <- rep("Undetected", n)
  gt$a1_preliminary <- rep(NA_character_, n)
  gt$a2_imprecision <- rep("Not assessed", n)
  gt$a2_reason <- rep("", n)
  gt$a2_final_direct <- rep(NA_character_, n)
  gt$b1_common <- rep("", n)
  gt$b1_cert1 <- rep(NA_character_, n)
  gt$b1_cert2 <- rep(NA_character_, n)
  gt$b1_start <- rep(NA_character_, n)
  gt$b2_intransitivity <- rep("Not serious", n)
  gt$b2_preliminary <- rep(NA_character_, n)
  gt$b3_imprecision <- rep("Not assessed", n)
  gt$b3_reason <- rep("", n)
  gt$b3_final_indirect <- rep(NA_character_, n)
  gt$c1_dominant <- rep(NA_character_, n)
  gt$c1_source <- rep(NA_character_, n)
  gt$c1_certainty <- rep(NA_character_, n)
  gt$c1_reason <- rep("", n)
  gt$c2_incoherence <- rep("Not serious", n)
  gt$c2_reason <- rep("", n)
  gt$c3_imprecision <- rep("Not serious", n)
  gt$c3_reason <- rep("", n)
  gt$c3_final_nma <- rep(NA_character_, n)
  gt$d_best <- rep("NMA", n)
  gt$d_reason <- rep("", n)
  gt$d_final_certainty <- rep(NA_character_, n)
  gt$d_effect_TE <- rep(NA_real_, n)
  gt$d_effect_lower <- rep(NA_real_, n)
  gt$d_effect_upper <- rep(NA_real_, n)
  gt$color_category <- rep(NA_character_, n)

  for (i in seq_len(n)) {
    gt$a1_preliminary[i] <- calc_a1_preliminary(gt$a1_rob[i], gt$a1_inconsistency[i],
      gt$a1_indirectness[i], gt$a1_pub_bias[i])
  }

  return(gt)
}

# =============================================================================
# apply_netsplit — ns_obj の内部構造から直接取得する方式に修正
# =============================================================================

apply_netsplit <- function(gt, ns_obj) {
  ns_comps <- ns_obj$comparison
  n_ns <- length(ns_comps)

  if (n_ns == 0) {
    cat("netsplit: comparison数=0\n")
    return(gt)
  }

  # netsplit 内部構造から直接取得
  ns_dir_te  <- ns_obj$direct.random$TE
  ns_dir_lo  <- ns_obj$direct.random$lower
  ns_dir_up  <- ns_obj$direct.random$upper
  ns_dir_se  <- ns_obj$direct.random$seTE
  ns_ind_te  <- ns_obj$indirect.random$TE
  ns_ind_lo  <- ns_obj$indirect.random$lower
  ns_ind_up  <- ns_obj$indirect.random$upper
  ns_ind_se  <- ns_obj$indirect.random$seTE
  ns_p       <- ns_obj$random$p
  ns_k       <- ns_obj$k

  cat("apply_netsplit: n_ns=", n_ns, "\n")
  cat("  direct.random$TE 非NA:", sum(!is.na(ns_dir_te)), "/", length(ns_dir_te), "\n")
  cat("  indirect.random$TE 非NA:", sum(!is.na(ns_ind_te)), "/", length(ns_ind_te), "\n")

  # 各 netsplit comparison を t1, t2 に分割
  ns_split <- strsplit(ns_comps, "\\s*vs\\s*|\\s*:\\s*")
  ns_t1_raw <- sapply(ns_split, function(x) trimws(x[1]))
  ns_t2_raw <- sapply(ns_split, function(x) trimws(x[2]))

  # ソート済みキー
  ns_sorted_keys <- sapply(seq_len(n_ns), function(j) {
    s <- sort(c(ns_t1_raw[j], ns_t2_raw[j]))
    paste0(s[1], ":", s[2])
  })

  for (j in seq_len(n_ns)) {
    # gt 内のマッチング
    idx <- which(gt$comparison == ns_sorted_keys[j])
    if (length(idx) == 0) next
    i <- idx[1]

    # 方向判定: netsplit の t1→t2 が gt の t1→t2 と逆か?
    flip <- (ns_t1_raw[j] == gt$treat2[i] && ns_t2_raw[j] == gt$treat1[i])

    cat("  ns[", j, "] ", ns_comps[j], " → gt[", i, "] ", gt$comparison[i],
        " flip=", flip,
        " dir_te=", if(is.na(ns_dir_te[j])) "NA" else round(ns_dir_te[j],4),
        " ind_te=", if(is.na(ns_ind_te[j])) "NA" else round(ns_ind_te[j],4),
        "\n")

    # k
    if (!is.null(ns_k) && length(ns_k) >= j && !is.na(ns_k[j])) {
      gt$direct_k[i] <- as.integer(ns_k[j])
    }

    # Direct
    if (!is.na(ns_dir_te[j])) {
      if (flip) {
        gt$direct_TE[i]    <- -ns_dir_te[j]
        gt$direct_lower[i] <- -ns_dir_up[j]
        gt$direct_upper[i] <- -ns_dir_lo[j]
        gt$ns_direct_TE[i]    <- -ns_dir_te[j]
        gt$ns_direct_lower[i] <- -ns_dir_up[j]
        gt$ns_direct_upper[i] <- -ns_dir_lo[j]
      } else {
        gt$direct_TE[i]    <- ns_dir_te[j]
        gt$direct_lower[i] <- ns_dir_lo[j]
        gt$direct_upper[i] <- ns_dir_up[j]
        gt$ns_direct_TE[i]    <- ns_dir_te[j]
        gt$ns_direct_lower[i] <- ns_dir_lo[j]
        gt$ns_direct_upper[i] <- ns_dir_up[j]
      }
      if (!is.null(ns_dir_se) && length(ns_dir_se) >= j) {
        gt$direct_seTE[i] <- ns_dir_se[j]
      }
    }

    # Indirect
    if (!is.na(ns_ind_te[j])) {
      if (flip) {
        gt$indirect_TE[i]    <- -ns_ind_te[j]
        gt$indirect_lower[i] <- -ns_ind_up[j]
        gt$indirect_upper[i] <- -ns_ind_lo[j]
      } else {
        gt$indirect_TE[i]    <- ns_ind_te[j]
        gt$indirect_lower[i] <- ns_ind_lo[j]
        gt$indirect_upper[i] <- ns_ind_up[j]
      }
      if (!is.null(ns_ind_se) && length(ns_ind_se) >= j) {
        gt$indirect_seTE[i] <- ns_ind_se[j]
      }
    }

    # p
    if (!is.null(ns_p) && length(ns_p) >= j && !is.na(ns_p[j])) {
      gt$netsplit_p[i] <- ns_p[j]
    }
  }

  gt$has_direct <- !is.na(gt$direct_TE) & gt$direct_k > 0

  # デバッグ: 上書き結果確認
  cat("\napply_netsplit 結果確認:\n")
  for (i in seq_len(min(6, nrow(gt)))) {
    cat("  ", gt$comparison[i],
        ": ns_direct=", if(is.na(gt$ns_direct_TE[i])) "NA" else round(gt$ns_direct_TE[i], 4),
        " pw_direct=", if(is.na(gt$pw_direct_TE[i])) "NA" else round(gt$pw_direct_TE[i], 4),
        " direct_TE=", if(is.na(gt$direct_TE[i])) "NA" else round(gt$direct_TE[i], 4),
        "\n")
  }

  return(gt)
}

# =============================================================================
# TSVコピー用 JavaScript
# =============================================================================

tsv_copy_js <- tags$script(HTML('
  function stripHTML(html) {
    var tmp = document.createElement("DIV");
    tmp.innerHTML = html;
    return (tmp.textContent || tmp.innerText || "").replace(/\\t/g, " ").replace(/\\r?\\n/g, " ").trim();
  }
  function copyDTAsTSV(wrapperId) {
    var container = document.getElementById(wrapperId);
    if (!container) { alert("DT not found: " + wrapperId); return; }
    var table = $(container).find("table").DataTable();
    if (!table) { alert("DataTable instance not found in: " + wrapperId); return; }
    var tsv = [];
    var headers = [];
    table.columns().every(function() { headers.push(stripHTML($(this.header()).html())); });
    if (headers.length > 0) tsv.push(headers.join("\\t"));
    table.rows({search: "applied"}).every(function() {
      var d = this.data(); var rowData = [];
      for (var j = 0; j < d.length; j++) { rowData.push(stripHTML(String(d[j]))); }
      tsv.push(rowData.join("\\t"));
    });
    copyToClip(tsv.join("\\n"), wrapperId);
  }
  function copyTableAsTSV(containerId) {
    var container = document.getElementById(containerId);
    if (!container) { alert("Table not found: " + containerId); return; }
    var table = container.querySelector("table");
    if (!table) { alert("No <table> found inside: " + containerId); return; }
    var tsv = [];
    table.querySelectorAll("tr").forEach(function(row) {
      var cells = row.querySelectorAll("th, td"); var rowData = [];
      cells.forEach(function(cell) { rowData.push(stripHTML(cell.innerHTML)); });
      if (rowData.length > 0) tsv.push(rowData.join("\\t"));
    });
    copyToClip(tsv.join("\\n"), containerId);
  }
  function copyPreAsTSV(preId) {
    var el = document.getElementById(preId);
    if (!el) { alert("Element not found: " + preId); return; }
    copyToClip(el.innerText || el.textContent || "", preId);
  }
  function copyToClip(text, sourceId) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function() { showCopyFeedback(sourceId); }).catch(function() { fallbackCopy(text, sourceId); });
    } else { fallbackCopy(text, sourceId); }
  }
  function fallbackCopy(text, sourceId) {
    var ta = document.createElement("textarea"); ta.value = text; ta.style.position = "fixed"; ta.style.left = "-9999px";
    document.body.appendChild(ta); ta.select(); try { document.execCommand("copy"); } catch(e) {} document.body.removeChild(ta);
    showCopyFeedback(sourceId);
  }
  function showCopyFeedback(sourceId) {
    var btn = document.querySelector("[data-copy-target=\'" + sourceId + "\']");
    if (btn) { var orig = btn.innerHTML; var origBg = btn.style.backgroundColor;
      btn.innerHTML = "\\u2705 Copied!"; btn.style.backgroundColor = "#d4edda";
      setTimeout(function() { btn.innerHTML = orig; btn.style.backgroundColor = origBg; }, 1500); }
  }
'))

tsv_btn <- function(target_id, label = "\U0001F4CB TSV\u30B3\u30D4\u30FC") {
  tags$button(HTML(label), class="btn btn-default btn-sm", style="margin:4px 2px;font-size:11px;padding:2px 10px;cursor:pointer;border-radius:4px;border:1px solid #aaa;",
    `data-copy-target`=target_id, onclick=sprintf("copyTableAsTSV('%s')",target_id))
}
tsv_btn_dt <- function(target_id, label = "\U0001F4CB TSV\u30B3\u30D4\u30FC") {
  tags$button(HTML(label), class="btn btn-default btn-sm", style="margin:4px 2px;font-size:11px;padding:2px 10px;cursor:pointer;border-radius:4px;border:1px solid #aaa;",
    `data-copy-target`=target_id, onclick=sprintf("copyDTAsTSV('%s')",target_id))
}
tsv_btn_pre <- function(target_id, label = "\U0001F4CB \u30B3\u30D4\u30FC") {
  tags$button(HTML(label), class="btn btn-default btn-sm", style="margin:4px 2px;font-size:11px;padding:2px 10px;cursor:pointer;border-radius:4px;border:1px solid #aaa;",
    `data-copy-target`=target_id, onclick=sprintf("copyPreAsTSV('%s')",target_id))
}

# =============================================================================
# UI
# =============================================================================

ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(title = "NMA-GRADE v1.0", titleWidth = 280),
  dashboardSidebar(width = 280,
    sidebarMenu(id = "tabs",
      menuItem("0. プロジェクト設定", tabName = "tab_settings", icon = icon("cog")),
      menuItem("1. データ読込", tabName = "tab_data", icon = icon("upload")),
      menuItem("2. NMA実行", tabName = "tab_nma_run", icon = icon("play")),
      tags$hr(style="margin:5px 15px;border-color:#444;"),
      tags$div(style="padding:2px 15px;color:#8aa;font-size:11px;","解析結果"),
      menuItem("ネットワーク図", tabName = "tab_netgraph", icon = icon("project-diagram")),
      menuItem("Forest Plot", tabName = "tab_forest", icon = icon("tree")),
      menuItem("ノード分割 (netsplit)", tabName = "tab_netsplit", icon = icon("code-branch")),
      menuItem("リーグテーブル", tabName = "tab_league", icon = icon("table")),
      menuItem("ランキング", tabName = "tab_rank", icon = icon("sort-amount-down")),
      menuItem("ファンネルプロット", tabName = "tab_funnel", icon = icon("filter")),
      tags$hr(style="margin:5px 15px;border-color:#444;"),
      tags$div(style="padding:2px 15px;color:#8aa;font-size:11px;","GRADE評価"),
      menuItem("A1. 直接 4ドメイン", tabName = "tab_a1", icon = icon("shield-alt")),
      menuItem("A2. 直接 不精確さ", tabName = "tab_a2", icon = icon("arrows-alt-h")),
      menuItem("B1. ループ選択", tabName = "tab_b1", icon = icon("project-diagram")),
      menuItem("B2. 非推移性", tabName = "tab_b2", icon = icon("random")),
      menuItem("B3. 間接 不精確さ", tabName = "tab_b3", icon = icon("arrows-alt-h")),
      menuItem("C1. NMA出発点", tabName = "tab_c1", icon = icon("sign-in-alt")),
      menuItem("C2. 非整合性", tabName = "tab_c2", icon = icon("exclamation-triangle")),
      menuItem("C3. NMA不精確さ", tabName = "tab_c3", icon = icon("arrows-alt-h")),
      menuItem("D. 最良推定値", tabName = "tab_d", icon = icon("trophy")),
      tags$hr(style="margin:5px 15px;border-color:#444;"),
      tags$div(style="padding:2px 15px;color:#8aa;font-size:11px;","結果"),
      menuItem("SoFテーブル", tabName = "tab_sof", icon = icon("table")),
      menuItem("色分け図", tabName = "tab_color", icon = icon("palette")),
      menuItem("保存/読込", tabName = "tab_export", icon = icon("save"))
    )
  ),
  dashboardBody(
    useShinyjs(),
    tags$head(
      tsv_copy_js,
      tags$style(HTML("
        .info-bx{background:#eef6ff;border-left:4px solid #4169E1;padding:12px 16px;margin:10px 0;border-radius:0 8px 8px 0;font-size:13px;}
        .warn-bx{background:#fff8e1;border-left:4px solid #f39c12;padding:12px 16px;margin:10px 0;border-radius:0 8px 8px 0;font-size:13px;}
        .err-bx{background:#fde8e8;border-left:4px solid #dc3545;padding:12px 16px;margin:10px 0;border-radius:0 8px 8px 0;font-size:13px;}
        .gc{border:1px solid #dee2e6;border-radius:8px;padding:14px;margin-bottom:12px;background:#fff;}
        .gc h5{margin:0 0 10px;color:#2C3E50;font-size:14px;}
        .btn-run{font-size:14px;padding:8px 20px;}
        .reason-text{font-size:11px;color:#555;margin-top:4px;font-style:italic;background:#f8f9fa;padding:6px 10px;border-radius:4px;}
        .step-hd{background:#2C3E50;color:#fff;padding:10px 16px;border-radius:8px;margin-bottom:12px;}
        .step-hd h4{margin:0;}
        .a1-comp-hd{background:#34495E;color:#fff;padding:10px 14px;border-radius:6px;margin-bottom:8px;font-size:15px;font-weight:bold;}
        .a1-meta-box{background:#f8f9fa;border:1px solid #dee2e6;border-radius:6px;padding:10px;margin:8px 0;font-size:12px;}
        .ns-direct-note{background:#fff3e0;border:1px solid #ff9800;border-radius:6px;padding:8px 12px;margin:6px 0;font-size:12px;}
      "))
    ),
    tabItems(
      tabItem(tabName = "tab_settings",
        div(class="step-hd",h4("プロジェクト設定")),
        box(width=12,status="primary",solidHeader=TRUE,title="基本設定",
          fluidRow(column(4,textInput("proj_name","プロジェクト名","My NMA")),column(4,textInput("outcome_name","アウトカム名","Mortality")),column(4,selectInput("effect_measure","効果指標",c("OR","RR","RD","MD","SMD"),"OR"))),
          fluidRow(column(3,radioButtons("outcome_type","アウトカム",c("二値"="binary","連続"="continuous"),inline=TRUE)),column(3,checkboxInput("small_good","小さい値が良い結果",TRUE)))
        ),
        box(width=12,status="info",solidHeader=TRUE,title="効果の大きさの閾値（GRADE partially contextualised framework）",
          div(class="info-bx",HTML(paste0(
            "<b>GRADE NMA結論フレームワーク（BMJ 2020;371:m3907 Brignardello-Petersen et al.）</b><br>",
            "NMAの結論を導くために、介入を効果の大きさで分類します:<br>",
            "\u2022 <b>些細な効果 (Trivial)</b>: 点推定値が小さい効果の閾値内<br>",
            "\u2022 <b>小さいが重要な効果 (Small but important)</b>: 小さい閾値〜中程度の閾値<br>",
            "\u2022 <b>中程度の効果 (Moderate)</b>: 中程度の閾値〜大きい閾値<br>",
            "\u2022 <b>大きな効果 (Large)</b>: 大きい閾値を超える<br>",
            "<span style='color:#c00;'>閾値は<b>絶対効果</b>に基づいて設定することが推奨されています。</span><br>",
            "二値アウトカムの場合、相対効果をベースラインリスクで絶対リスク差に換算して判断します。"
          ))),
          h5("相対効果の閾値（RR/OR尺度）"),
          div(style="font-size:12px;color:#555;margin-bottom:8px;",HTML(
            "小さい効果の境界 = 些細な効果と小さい効果の境界線。不精確さの評価で95%CIがこの閾値を跨ぐかを判定します。"
          )),
          fluidRow(
            column(3,numericInput("mid_lower","小さい効果(benefit側)",0.80,min=0.1,max=0.99,step=0.05)),
            column(3,numericInput("mid_upper","小さい効果(harm側)",1.25,min=1.01,max=5,step=0.05)),
            column(3,numericInput("large_effect","大きな効果(benefit側)",0.50,min=0.05,max=0.99,step=0.05)),
            column(3,numericInput("baseline_risk","ベースラインリスク",0.10,min=0.001,max=1.0,step=0.01))
          ),
          h5("連続アウトカム用の閾値（自然単位）"),
          div(style="font-size:12px;color:#555;margin-bottom:8px;",HTML(
            "例: 下痢持続時間なら 小さい効果=3時間, 中程度=12時間, 大きい=24時間"
          )),
          fluidRow(
            column(3,numericInput("thresh_small_cont","小さい効果",3,step=1)),
            column(3,numericInput("thresh_moderate_cont","中程度の効果",12,step=1)),
            column(3,numericInput("thresh_large_cont","大きな効果",24,step=1))
          ),
          div(class="info-bx",HTML(paste0(
            "<b>絶対効果（Absolute estimate）について:</b><br>",
            "GRADEでは、効果の大きさの判断は<b>絶対効果</b>に基づくべきとされています。<br>",
            "\u2022 <b>二値アウトカム</b>: 絶対リスク差 \u2248 ベースラインリスク \u00D7 (RR \u2212 1)<br>",
            "&nbsp;&nbsp;例: RR=0.80, ベースラインリスク10% → 絶対差 = \u22122%（1000人あたり20人減少）<br>",
            "&nbsp;&nbsp;例: RR=0.80, ベースラインリスク2% → 絶対差 = \u22120.4%（1000人あたり4人減少）→些細<br>",
            "&nbsp;&nbsp;例: RR=0.80, ベースラインリスク40% → 絶対差 = \u22128%（1000人あたり80人減少）→中程度〜大きい<br>",
            "\u2022 <b>連続アウトカム</b>: 平均差(MD)がそのまま絶対効果。SMDの場合はSD単位。<br>",
            "→ <b>同じ相対効果でもベースラインリスクにより臨床的意義が大きく異なります。</b>"
          ))),
          div(class="info-bx",HTML(paste0(
            "<b>不精確さの評価手順（GRADE guidelines 33, JCE 2021）:</b><br>",
            "<b>Step 1:</b> 95%CIが閾値を跨ぐか確認<br>",
            "&nbsp;&nbsp;\u2022 1つの閾値を跨ぐ → Serious（1段階格下げ）<br>",
            "&nbsp;&nbsp;\u2022 2つ以上の閾値を跨ぐ/CIが非常に広い → Very serious（2段階格下げ）<br>",
            "&nbsp;&nbsp;\u2022 極端に広いCI（大きな利益〜大きな害を含む） → Extremely serious（3段階格下げ）<br>",
            "<b>Step 2:</b> CIが閾値を跨がない場合 → 効果量を確認<br>",
            "&nbsp;&nbsp;\u2022 効果量が中程度以下(例:RRR<30%)で妥当 → 格下げ不要<br>",
            "&nbsp;&nbsp;\u2022 効果量が大きい → OIS（最適情報量）を確認<br>",
            "<b>Step 3:</b> OIS簡易判定<br>",
            "&nbsp;&nbsp;\u2022 RRのCI上限/下限比 > 3.0 → OIS未達（計算不要で格下げ）<br>",
            "&nbsp;&nbsp;\u2022 ORのCI上限/下限比 > 2.5 → OIS未達の可能性大<br>",
            "<b>注意:</b> 非整合性で既に格下げしている場合、不精確さとの二重カウントを避ける"
          )))
        )
      ),
      tabItem(tabName = "tab_data",
        div(class="step-hd",h4("1. データ読込")),
        fluidRow(
          box(width=6,status="primary",solidHeader=TRUE,title="CSVアップロード",fileInput("csvFile","CSVファイル",accept=".csv"),fluidRow(column(4,checkboxInput("csv_header","ヘッダーあり",TRUE)),column(4,selectInput("csv_sep","区切り",c(","=",",";"=";","Tab"="\t"))),column(4,selectInput("csv_enc","文字コード",c("UTF-8","CP932"))))),
          box(width=6,status="info",solidHeader=TRUE,title="サンプルデータ",selectInput("sampleData","選択",c("--"="","smoking","heartfailure","diabetes","antidiabetic")),actionButton("loadSample","読込",class="btn-info"))
        ),
        fluidRow(box(width=12,status="success",solidHeader=TRUE,title="列マッピング",fluidRow(column(3,radioButtons("data_format","形式",c("アーム"="arm","コントラスト"="contrast"),inline=TRUE)),column(9,uiOutput("col_mapping_ui"))))),
        fluidRow(box(width=12,title="プレビュー",tsv_btn_dt("data_preview"),DT::dataTableOutput("data_preview")))
      ),
      tabItem(tabName = "tab_nma_run",
        div(class="step-hd",h4("2. NMA実行")),
        box(width=12,status="success",solidHeader=TRUE,
          fluidRow(column(3,uiOutput("ref_treat_ui")),column(3,selectInput("tau_method","推定法",c("REML","DL","ML"),"REML")),column(3,checkboxInput("random_eff","ランダム効果",TRUE)),column(3,br(),actionButton("runNMA","NMA実行",class="btn-success btn-run"))),
          uiOutput("nma_status"),tsv_btn_pre("nma_debug_info"),verbatimTextOutput("nma_debug_info"))
      ),
      tabItem(tabName="tab_netgraph",div(class="step-hd",h4("ネットワーク図")),box(width=12,fluidRow(column(3,numericInput("ng_cex","倍率",1.5,0.5,5,0.5)),column(3,actionButton("runNetgraph","描画",class="btn-primary"))),plotOutput("netgraph_plot",height="550px"))),
      tabItem(tabName="tab_forest",div(class="step-hd",h4("Forest Plot")),box(width=12,actionButton("runForest","描画",class="btn-primary"),plotOutput("forest_plot",height="600px"))),
      tabItem(tabName = "tab_netsplit",
        div(class="step-hd",h4("ノード分割 (netsplit)")),
        box(width=12,
          div(class="info-bx","全比較ペアの直接/間接/NMA推定値（直接のみ・間接のみ含む）。",tags$span(style="color:red;","p<0.05は赤表示。"),tags$br(),"各ペアについて順方向と逆方向（符号反転）の両方を記載。",tags$span(style="color:#4169E1;","薄青背景行 (\u25C0) が逆方向。"),tags$br(),tags$span(style="color:#4caf50;","\u25A0")," Direct  ",tags$span(style="color:#ff9800;","\u25A0")," Indirect  ",tags$span(style="color:#2196f3;","\u25A0")," Direct+Indirect  ",tags$span(style="color:#9e9e9e;","\u25A0")," NMA only"),
          tsv_btn_pre("netsplit_debug"),verbatimTextOutput("netsplit_debug"),
          tsv_btn_dt("netsplit_table"),DT::dataTableOutput("netsplit_table"),
          hr(),h4("netsplit Forest Plot"),plotOutput("netsplit_forest",height="700px"))
      ),
      tabItem(tabName="tab_league",div(class="step-hd",h4("リーグテーブル")),box(width=12,tsv_btn_pre("league_output"),verbatimTextOutput("league_output"))),
      tabItem(tabName="tab_rank",div(class="step-hd",h4("ランキング")),box(width=12,tsv_btn_pre("rank_output"),verbatimTextOutput("rank_output"))),
      tabItem(tabName="tab_funnel",div(class="step-hd",h4("ファンネルプロット")),box(width=12,actionButton("runFunnel","描画",class="btn-primary"),plotOutput("funnel_plot",height="500px"))),
      tabItem(tabName = "tab_a1",
        div(class="step-hd",h4("A1: 直接推定値の4ドメイン評価")),
        box(width=12,
          div(class="info-bx",HTML(paste0(
            "<b>A1: 直接推定値の4ドメイン評価</b><br>",
            "直接比較のある各ペアについて、通常のGRADE 4ドメインで確実性を評価します。<br>",
            "下のドロップダウンで比較ペアを選択 → フォレストプロットとI\u00B2等が表示 → 4ドメインを評価 → 「この比較の評価を保存」で記録。<br>",
            "全比較を評価後、下部の一覧表で全評価を確認できます。<br>",
            "<b>※注:</b> Direct (netsplit) はグラフ理論的分離による推定値です。フォレストプロットは metagen() の値を表示しています。"
          ))),
          div(class="info-bx",HTML(paste0(
            "<b>各ドメインの評価ガイド:</b><br>",
            "<b>1. バイアスのリスク (RoB):</b><br>",
            "&nbsp;&nbsp;\u2022 各研究のRoB評価（RoB 2.0等）を考慮<br>",
            "&nbsp;&nbsp;\u2022 高リスク研究の割合と効果への影響を検討<br>",
            "&nbsp;&nbsp;\u2022 Serious: 高リスク研究が多い、または結果に重大な影響がある場合<br>",
            "<b>2. 非一貫性 (Inconsistency):</b><br>",
            "&nbsp;&nbsp;\u2022 フォレストプロットの点推定値のばらつきを視覚的に確認<br>",
            "&nbsp;&nbsp;\u2022 I\u00B2 > 50%: 中程度の異質性（Serious候補）<br>",
            "&nbsp;&nbsp;\u2022 I\u00B2 > 75%: 高い異質性（Very serious候補）<br>",
            "&nbsp;&nbsp;\u2022 ただしI\u00B2だけでなく、CIの重なりと臨床的意義も考慮<br>",
            "<b>3. 非直接性 (Indirectness):</b><br>",
            "&nbsp;&nbsp;\u2022 研究の対象者、介入、比較対照、アウトカムがレビューの疑問と合致しているか<br>",
            "&nbsp;&nbsp;\u2022 Serious: 1つ以上の要素で重要な相違がある場合<br>",
            "<b>4. 出版バイアス (Publication bias):</b><br>",
            "&nbsp;&nbsp;\u2022 ファンネルプロットの非対称性を確認<br>",
            "&nbsp;&nbsp;\u2022 小規模研究効果、言語バイアス等を考慮<br>",
            "&nbsp;&nbsp;\u2022 Strongly suspected: 明確な非対称性がある場合"
          ))),
          div(class="warn-bx",tags$b("一括ルール:")," 全比較に同じ判定を一括適用後、個別に修正可能。"),
          fluidRow(column(3,selectInput("a1_bulk_rob","RoB一括",c("--","Not serious","Serious","Very serious"))),column(3,selectInput("a1_bulk_inc","非一貫性一括",c("--","Not serious","Serious","Very serious"))),column(3,selectInput("a1_bulk_ind","非直接性一括",c("--","Not serious","Serious","Very serious"))),column(3,selectInput("a1_bulk_pub","出版バイアス一括",c("--","Undetected","Strongly suspected")))),
          actionButton("a1_bulk_apply","一括適用",class="btn-warning btn-sm"),hr(),
          fluidRow(column(6,uiOutput("a1_comp_selector")),column(6,div(style="margin-top:25px;",uiOutput("a1_selected_info")))),hr(),
          fluidRow(column(12,actionButton("a1_flip_direction","\u21C4 方向を反転して再描画 (A vs B \u2194 B vs A)",class="btn-info btn-sm",style="margin-bottom:10px;"),uiOutput("a1_direction_label"))),hr(),
          uiOutput("a1_meta_display"),hr(),uiOutput("a1_grade_inputs"),hr(),
          actionButton("a1_save_current","この比較の評価を保存",class="btn-success btn-run"),hr(),
          h4("A1 評価一覧"),tsv_btn_dt("a1_summary_table"),DT::dataTableOutput("a1_summary_table"))
      ),
      tabItem(tabName="tab_a2",div(class="step-hd",h4("A2: 直接不精確さ（オプション）")),box(width=12,
        div(class="warn-bx","非整合性(C2)がSerious以上の場合のみ必要です。非整合性がNot seriousの場合は不精確さの評価はC3（NMA推定値の不精確さ）のみで行います。"),
        div(class="info-bx",HTML(paste0(
          "<b>直接推定値の不精確さ評価（GRADE guidelines 33準拠）</b><br>",
          "直接推定値の95%CIと事前設定した閾値の関係で判定します:<br><br>",
          "<b>Step 1: CIと閾値の関係を確認</b><br>",
          "&nbsp;&nbsp;\u2022 CIが閾値を跨ぐ → 格下げ<br>",
          "&nbsp;&nbsp;\u2022 CIが1つの閾値を跨ぐ → <b>Serious</b>（1段階格下げ）<br>",
          "&nbsp;&nbsp;\u2022 CIが2つ以上の閾値を跨ぐ → <b>Very serious</b>（2段階格下げ）<br>",
          "&nbsp;&nbsp;\u2022 CIが極端に広い（大きな利益〜大きな害を含む） → <b>Extremely serious</b>（3段階格下げ）<br><br>",
          "<b>Step 2: CIが閾値を跨がない場合</b><br>",
          "&nbsp;&nbsp;\u2022 効果量が中程度以下（例: RRR < 30%）→ 格下げ不要<br>",
          "&nbsp;&nbsp;\u2022 効果量が大きい → OIS（最適情報量）を確認<br><br>",
          "<b>Step 3: OIS簡易判定（二値アウトカム）</b><br>",
          "&nbsp;&nbsp;\u2022 RRのCI上限/下限比率 > 3.0 → OIS未達成（計算不要で格下げ）<br>",
          "&nbsp;&nbsp;\u2022 ORのCI上限/下限比率 > 2.5 → OIS未達成の可能性大<br>",
          "&nbsp;&nbsp;\u2022 比率が閾値付近の場合 → 実効サンプルサイズの計算を推奨<br><br>",
          "<b>注意:</b> 非整合性で既に格下げしている場合、不精確さとの二重カウントを避けてください。"
        ))),
        checkboxInput("a2_auto","自動判定",FALSE),conditionalPanel("input.a2_auto",actionButton("a2_auto_run","自動判定実行",class="btn-info btn-sm")),uiOutput("a2_cards"))),
      tabItem(tabName="tab_b1",div(class="step-hd",h4("B1: ループ選択")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>間接推定値の出発点（B1）</b><br>",
          "各比較ペア A vs B に対して、間接エビデンスを構成する「ループ」を評価します。<br>",
          "共通比較対象 C を選択すると、A vs C と C vs B の直接エビデンスの確実性のうち、",
          "<b>低い方</b>が間接推定値の出発点となります。<br>",
          "ループが複数ある場合は、最も情報量の多い（優勢な）ループを選択してください。<br>",
          "優勢ループの評価基準: (1) 研究数(k)が多い、(2) 症例数(N)が多い、(3) エビデンスの確実性が高い"
        ))),
        tags$div(style="margin:10px 0;",
          tags$a(href="https://mxe050.github.io/NMA/", target="_blank",
            class="btn btn-info btn-sm", style="font-size:13px;padding:6px 16px;",
            HTML("\U0001F517 優勢ループの評価については、以下のサイト参照")
          )
        ),
        uiOutput("b1_cards")
      )),
      tabItem(tabName="tab_b2",div(class="step-hd",h4("B2: 非推移性")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>B2: 非推移性 (Intransitivity) の評価</b><br>",
          "間接比較の妥当性は「推移性の仮定」に依存します。<br>",
          "A vs C の研究群と C vs B の研究群で、効果修飾因子の分布が類似しているかを評価します。<br><br>",
          "<b>評価のポイント:</b><br>",
          "\u2022 研究間で対象患者の特性（年齢、重症度、併存疾患等）に重要な差がないか<br>",
          "\u2022 介入の用量・期間・投与方法に重要な差がないか<br>",
          "\u2022 比較対照（共通比較対象 C）の定義が研究間で一致しているか<br>",
          "\u2022 アウトカムの測定方法・時点が一致しているか<br>",
          "\u2022 研究のセッティング（地域、医療水準）に重要な差がないか<br><br>",
          "<b>判定基準:</b><br>",
          "\u2022 <b>Not serious</b>: 効果修飾因子の分布に重要な差がない<br>",
          "\u2022 <b>Serious</b>: 1つ以上の効果修飾因子で重要な差がある<br>",
          "\u2022 <b>Very serious</b>: 複数の効果修飾因子で重要な差がある"
        ))),
        uiOutput("b2_cards"))),
      tabItem(tabName="tab_b3",div(class="step-hd",h4("B3: 間接不精確さ（オプション）")),box(width=12,
        div(class="warn-bx","非整合性(C2)がSerious以上の場合のみ必要です。"),
        div(class="info-bx",HTML(paste0(
          "<b>間接推定値の不精確さ評価</b><br>",
          "間接推定値の95%CIに対してA2と同様の基準で判定します。<br>",
          "間接推定値のCIは一般に直接推定値より広くなる傾向があります。<br>",
          "GRADE guidelines 33の手順（CI→閾値→効果量→OIS）を適用してください。"
        ))),
        checkboxInput("b3_auto","自動判定",FALSE),conditionalPanel("input.b3_auto",actionButton("b3_auto_run","自動判定実行",class="btn-info btn-sm")),uiOutput("b3_cards"))),
      tabItem(tabName="tab_c1",div(class="step-hd",h4("C1: NMA出発点")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>C1: NMA推定値の出発点</b><br>",
          "NMA推定値の確実性は、直接エビデンス(A1)と間接エビデンス(B2)のうち、",
          "<b>より支配的な方</b>の確実性を出発点とします。<br><br>",
          "<b>支配性の判定:</b><br>",
          "\u2022 直接推定値のみ存在 → 直接(A1)を出発点<br>",
          "\u2022 間接推定値のみ存在 → 間接(B2)を出発点<br>",
          "\u2022 両方存在 → CIの幅の比率で判定<br>",
          "&nbsp;&nbsp;- 直接CI幅/間接CI幅 < 0.5 → 直接が支配的<br>",
          "&nbsp;&nbsp;- 直接CI幅/間接CI幅 > 2.0 → 間接が支配的<br>",
          "&nbsp;&nbsp;- 0.5〜2.0 → 同程度（確実性の高い方を選択）"
        ))),
        actionButton("c1_auto_run","自動判定",class="btn-info"),uiOutput("c1_cards"))),
      tabItem(tabName="tab_c2",div(class="step-hd",h4("C2: 非整合性")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>C2: 非整合性 (Incoherence) の評価</b><br>",
          "直接推定値と間接推定値の間に統計的な不一致がないかを評価します。<br><br>",
          "<b>netsplit p値による判定:</b><br>",
          "\u2022 p \u2265 0.10 → <b>Not serious</b>（非整合性なし）<br>",
          "\u2022 p < 0.10 かつ効果の方向が一致 → <b>Serious</b>候補<br>",
          "\u2022 p < 0.05 かつ効果の方向が不一致 → <b>Very serious</b>候補<br><br>",
          "<b>追加の考慮事項:</b><br>",
          "\u2022 p値だけでなく、直接・間接推定値の点推定値と信頼区間の重なりも視覚的に確認<br>",
          "\u2022 全体的な非整合性（デザインバイデザインモデル等）の結果も参考に<br>",
          "\u2022 <b>重要:</b> 非整合性がSerious以上の場合、A2（直接不精確さ）とB3（間接不精確さ）の評価が必要になります<br>",
          "\u2022 不精確さの評価で非整合性と二重カウントしないよう注意（GRADE guidelines 33）"
        ))),
        fluidRow(column(4,actionButton("c2_auto_run","半自動判定",class="btn-info")),column(4,actionButton("c2_bulk_ns","p>=0.10一括NS",class="btn-warning btn-sm"))),uiOutput("c2_cards"))),
      tabItem(tabName="tab_c3",div(class="step-hd",h4("C3: NMA不精確さ")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>C3: NMA推定値の不精確さ（GRADE guidelines 33, JCE 2021 完全準拠）</b><br>",
          "NMA推定値（ネットワーク推定値）の95%CIに対して不精確さを評価します。<br>",
          "NMA推定値は直接+間接エビデンスを統合しているため、通常は個別の推定値よりCIが狭くなります。<br><br>",
          "<b>\u25B6 評価手順（Fig.2 フローチャート準拠）:</b><br><br>",
          "<b>Step 1: 95%CIが閾値を跨ぐか?</b><br>",
          "&nbsp;&nbsp;→ はい: 格下げが必要（Step 1a へ）<br>",
          "&nbsp;&nbsp;→ いいえ: 効果量を確認（Step 2 へ）<br><br>",
          "<b>Step 1a: 何段階格下げするか?</b><br>",
          "&nbsp;&nbsp;\u2022 CIが1つの閾値を跨ぐ → <b>Serious</b>（1段階）<br>",
          "&nbsp;&nbsp;\u2022 CIが2つの閾値を跨ぐ → <b>Very serious</b>（2段階）<br>",
          "&nbsp;&nbsp;\u2022 CIが極端に広い（例: 大きな利益〜大きな害を含む） → <b>Extremely serious</b>（3段階）<br>",
          "&nbsp;&nbsp;&nbsp;&nbsp;例: リスク差 329人減少（330人減少〜670人増加）→3段階格下げが妥当<br><br>",
          "<b>Step 2: CIが閾値を跨がない場合 → 効果量を確認</b><br>",
          "&nbsp;&nbsp;\u2022 効果量が中程度以下（例: 相対リスク減少<30%）で妥当 → <b>格下げ不要</b><br>",
          "&nbsp;&nbsp;\u2022 効果量が大きい（例: RRR>30%） → OISを確認（Step 3へ）<br><br>",
          "<b>Step 3: OIS（最適情報量）の確認</b><br>",
          "&nbsp;&nbsp;NMA推定値のCIから「実効サンプルサイズ」を逆算できます:<br>",
          "&nbsp;&nbsp;\u2022 SE<sub>NMA</sub> = (log(CI上限) - log(CI下限)) / 3.92<br>",
          "&nbsp;&nbsp;\u2022 この SE から単一試験に換算したサンプルサイズを計算<br>",
          "&nbsp;&nbsp;\u2022 <b>簡易判定:</b> CI上限/CI下限の比率（相対尺度）で判定可能<br>",
          "&nbsp;&nbsp;&nbsp;&nbsp;- RRの場合: 比率 > 3.0 → OIS未達成（計算不要で格下げ）<br>",
          "&nbsp;&nbsp;&nbsp;&nbsp;- ORの場合: 比率 > 2.5 → OIS未達成の可能性大<br>",
          "&nbsp;&nbsp;&nbsp;&nbsp;- 比率が閾値近傍(例: 2.7〜3.3) → 計算して確認推奨<br><br>",
          "<b>\u26A0 注意事項:</b><br>",
          "\u2022 非整合性(C2)で既に格下げしている場合、不精確さとの二重カウントを避ける<br>",
          "\u2022 スパースなネットワークでは極端に広いCIが出やすい → 3段階格下げも検討<br>",
          "\u2022 閾値の判断は<b>絶対効果</b>に基づいて行うべき（プロジェクト設定の閾値を参照）"
        ))),
        checkboxInput("c3_auto","自動判定",FALSE),conditionalPanel("input.c3_auto",actionButton("c3_auto_run","自動判定実行",class="btn-info btn-sm")),uiOutput("c3_cards"))),
      tabItem(tabName="tab_d",div(class="step-hd",h4("D: 最良推定値")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>D: 最良推定値の選択</b><br>",
          "非整合性(C2)の結果に基づいて、各比較の最良推定値を決定します。<br><br>",
          "<b>判定ルール:</b><br>",
          "\u2022 非整合性が<b>Not serious</b> → NMA推定値（C3の確実性）を採用<br>",
          "\u2022 非整合性が<b>Serious以上</b> → 直接(A2)・間接(B3)・NMA(C3)のうち最も確実性の高いものを採用<br><br>",
          "<b>絶対効果の換算（二値アウトカム）:</b><br>",
          "相対効果を絶対効果に変換する式:<br>",
          "&nbsp;&nbsp;\u2022 RR: 絶対リスク差 = ベースラインリスク \u00D7 (RR \u2212 1)<br>",
          "&nbsp;&nbsp;\u2022 OR: 介入群リスク = (ベースラインリスク \u00D7 OR) / (1 \u2212 ベースラインリスク + ベースラインリスク \u00D7 OR)<br>",
          "&nbsp;&nbsp;&nbsp;&nbsp;絶対リスク差 = 介入群リスク \u2212 ベースラインリスク<br>",
          "→ SoFテーブルで絶対効果も表示されます。"
        ))),
        actionButton("d_auto_run","自動判定",class="btn-success btn-run"),uiOutput("d_warn_a2b3"),uiOutput("d_cards"))),
      tabItem(tabName="tab_sof",div(class="step-hd",h4("SoFテーブル")),box(width=12,
        div(class="info-bx",HTML(paste0(
          "<b>Summary of Findings (SoF) テーブル</b><br>",
          "各比較ペアの最終的な効果推定値、確実性、色分けカテゴリを一覧表示します。<br>",
          "<b>二値アウトカムの場合</b>は、相対効果に加えて「1000人あたりの絶対効果」も表示されます（ベースラインリスクに基づく換算）。<br>",
          "GRADEでは効果の大きさの解釈は絶対効果に基づくことが推奨されています。"
        ))),
        checkboxInput("sof_ref_only","参照介入との比較のみ",TRUE),tsv_btn_dt("sof_table"),DT::dataTableOutput("sof_table"),br(),downloadButton("dl_sof_xlsx","Excel"))),
      tabItem(tabName="tab_color",div(class="step-hd",h4("色分け図")),box(width=12,fluidRow(column(4,radioButtons("color_version","版",c("Gradient"="gradient","Stoplight"="stoplight","両方"="both"),inline=TRUE)),column(4,actionButton("gen_color_fig","生成",class="btn-success btn-run")),column(4,downloadButton("dl_color_png","PNG"))),uiOutput("color_fig_output"))),
      tabItem(tabName="tab_export",div(class="step-hd",h4("保存/読込")),fluidRow(box(width=6,title="保存",downloadButton("save_json","JSON保存",class="btn-primary btn-run")),box(width=6,title="読込",fileInput("load_json","JSON",accept=".json"),actionButton("load_json_btn","読込",class="btn-info"))),box(width=12,downloadButton("dl_all_xlsx","Excel一括",class="btn-success btn-run")))
    )
  )
)

# =============================================================================
# SERVER
# =============================================================================

server <- function(input, output, session) {

  rv <- reactiveValues(raw_data=NULL,pairwise_data=NULL,nma_result=NULL,netsplit_result=NULL,net_rank=NULL,grade_table=NULL,treatments=NULL,color_plot_data=NULL,color_plot_func=NULL,pairwise_meta_results=NULL,pw_study_counts=NULL,a1_flipped=FALSE)

  is_ratio <- reactive({ input$effect_measure %in% c("OR","RR","HR") })
  fmt_est <- function(te, lo, up) {
    if(is.na(te)) return("---")
    if(is_ratio()) sprintf("%.2f [%.2f, %.2f]",exp(te),exp(lo),exp(up))
    else sprintf("%.2f [%.2f, %.2f]",te,lo,up)
  }

  observeEvent(input$loadSample, {req(input$sampleData!="");tryCatch({env<-new.env();data(list=input$sampleData,package="NMA",envir=env);rv$raw_data<-get(input$sampleData,envir=env);showNotification(paste0(input$sampleData," 読込完了"),type="message")},error=function(e)showNotification(e$message,type="error"))})
  observeEvent(input$csvFile, {req(input$csvFile);tryCatch({rv$raw_data<-read.csv(input$csvFile$datapath,header=input$csv_header,sep=input$csv_sep,fileEncoding=input$csv_enc,stringsAsFactors=FALSE);showNotification("CSV読込完了",type="message")},error=function(e)showNotification(e$message,type="error"))})
  output$data_preview <- DT::renderDataTable({req(rv$raw_data);DT::datatable(rv$raw_data,options=list(pageLength=10,scrollX=TRUE))})
  output$col_mapping_ui <- renderUI({
    req(rv$raw_data);cols<-c("--"="",names(rv$raw_data));fc<-function(pats) find_column(rv$raw_data,pats)
    if(input$data_format=="arm"){tagList(fluidRow(column(3,selectInput("col_study","study",cols,fc(c("^study$","^id$")))),column(3,selectInput("col_trt","trt",cols,fc(c("^trt$","^treatment$","^t$")))),column(3,selectInput("col_d","d",cols,fc(c("^d$","^event")))),column(3,selectInput("col_n","n",cols,fc(c("^n$"))))),conditionalPanel("input.outcome_type=='continuous'",fluidRow(column(4,selectInput("col_m","mean",cols,fc(c("^m$","^mean$")))),column(4,selectInput("col_s","sd",cols,fc(c("^s$","^sd$")))),column(4,selectInput("col_n2","n",cols,fc(c("^n$")))))))}
    else{fluidRow(column(2,selectInput("col_study_c","study",cols,fc(c("^study$")))),column(2,selectInput("col_t1","treat1",cols,fc(c("^treat1")))),column(2,selectInput("col_t2","treat2",cols,fc(c("^treat2")))),column(3,selectInput("col_te","TE",cols,fc(c("^te$","^logor")))),column(3,selectInput("col_se","seTE",cols,fc(c("^se$","^sete")))))}
  })
  output$ref_treat_ui <- renderUI({req(rv$raw_data);trts<-NULL;if(input$data_format=="arm"&&!is.null(input$col_trt)&&input$col_trt!="")trts<-sort(unique(as.character(rv$raw_data[[input$col_trt]])));if(is.null(trts)||length(trts)==0)trts<-c("1");selectInput("ref_trt","参照介入",trts,trts[1])})

  observeEvent(input$runNMA, {
    req(rv$raw_data, input$ref_trt)
    tryCatch({
      withProgress(message="NMA解析中...", {
        incProgress(0.1)
        if(input$data_format=="arm"){
          if(input$outcome_type=="binary") rv$pairwise_data<-pairwise(treat=rv$raw_data[[input$col_trt]],event=rv$raw_data[[input$col_d]],n=rv$raw_data[[input$col_n]],studlab=rv$raw_data[[input$col_study]],sm=input$effect_measure)
          else rv$pairwise_data<-pairwise(treat=rv$raw_data[[input$col_trt]],mean=rv$raw_data[[input$col_m]],sd=rv$raw_data[[input$col_s]],n=rv$raw_data[[input$col_n2]],studlab=rv$raw_data[[input$col_study]],sm=input$effect_measure)
        } else {
          rv$pairwise_data<-data.frame(TE=rv$raw_data[[input$col_te]],seTE=rv$raw_data[[input$col_se]],treat1=rv$raw_data[[input$col_t1]],treat2=rv$raw_data[[input$col_t2]],studlab=rv$raw_data[[input$col_study_c]],stringsAsFactors=FALSE)
        }
        incProgress(0.3)
        rv$nma_result<-netmeta(TE=rv$pairwise_data$TE,seTE=rv$pairwise_data$seTE,treat1=rv$pairwise_data$treat1,treat2=rv$pairwise_data$treat2,studlab=rv$pairwise_data$studlab,reference.group=input$ref_trt,sm=input$effect_measure,random=input$random_eff,common=!input$random_eff,method.tau=input$tau_method)
        rv$treatments<-sort(rv$nma_result$trts)
        incProgress(0.2)
        rv$netsplit_result<-netsplit(rv$nma_result)
        pw<-rv$pairwise_data;pw$t1c<-as.character(pw$treat1);pw$t2c<-as.character(pw$treat2)
        pw$comp_key<-apply(pw[,c("t1c","t2c")],1,function(r)paste(sort(r),collapse=":"))
        n1c<-NULL;n2c<-NULL;for(cn in names(pw)){if(grepl("^n1$",cn,ignore.case=TRUE)&&is.null(n1c))n1c<-cn;if(grepl("^n2$",cn,ignore.case=TRUE)&&is.null(n2c))n2c<-cn}
        ucs<-unique(pw$comp_key);mr<-list();sc<-list()
        for(comp in ucs){sub<-pw[pw$comp_key==comp,];k<-nrow(sub);tn<-NA;if(!is.null(n1c)&&!is.null(n2c))tn<-sum(as.numeric(sub[[n1c]]),na.rm=TRUE)+sum(as.numeric(sub[[n2c]]),na.rm=TRUE);sc[[comp]]<-list(k=k,total_n=tn);if(k>=1)tryCatch({mr[[comp]]<-metagen(TE=sub$TE,seTE=sub$seTE,studlab=sub$studlab,sm=input$effect_measure,random=TRUE,common=TRUE)},error=function(e){})}
        rv$pairwise_meta_results<-mr;rv$pw_study_counts<-sc
        incProgress(0.1)
        rv$net_rank<-netrank(rv$nma_result,small.values=if(input$small_good)"good" else "bad")
        incProgress(0.2)
        gt <- build_grade_table(rv$nma_result, rv$pairwise_data)
        gt <- apply_netsplit(gt, rv$netsplit_result)
        rv$grade_table <- gt
        incProgress(0.1)
      })
      showNotification("NMA完了!",type="message")
    },error=function(e){showNotification(paste("NMAエラー:",e$message),type="error",duration=15)})
  })

  output$nma_status <- renderUI({
    if(is.null(rv$nma_result)) return(div(class="warn-bx","NMA未実行"))
    nt<-length(rv$nma_result$trts);nc<-nrow(rv$grade_table);nd<-sum(rv$grade_table$has_direct,na.rm=TRUE);ni<-sum(!is.na(rv$grade_table$indirect_TE))
    div(style="background:#d4edda;border-left:4px solid #28a745;padding:12px;margin:10px 0;border-radius:0 8px 8px 0;",HTML(paste0("<b>NMA完了</b> | 治療:",nt," | ペア:",nc," | 直接比較あり:",nd," | 間接推定あり:",ni)))
  })

  output$nma_debug_info <- renderPrint({
    req(rv$grade_table);gt<-rv$grade_table
    cat("=== GRADE Table ===\n");cat("行数:",nrow(gt),"\n")
    for(i in seq_len(min(6,nrow(gt)))){
      cat(gt$comparison[i],": k=",gt$direct_k[i],
          " ns_d=",if(is.na(gt$ns_direct_TE[i]))"NA" else round(gt$ns_direct_TE[i],3),
          " pw_d=",if(is.na(gt$pw_direct_TE[i]))"NA" else round(gt$pw_direct_TE[i],3),
          " d=",if(is.na(gt$direct_TE[i]))"NA" else round(gt$direct_TE[i],3),
          " i=",if(is.na(gt$indirect_TE[i]))"NA" else round(gt$indirect_TE[i],3),
          " n=",if(is.na(gt$nma_TE[i]))"NA" else round(gt$nma_TE[i],3),
          " p=",if(is.na(gt$netsplit_p[i]))"NA" else round(gt$netsplit_p[i],4),"\n")
    }
  })

  observeEvent(input$runNetgraph, {req(rv$nma_result);output$netgraph_plot<-renderPlot({nma_obj<-rv$nma_result;trts<-nma_obj$trts;np<-rep(0,length(trts));names(np)<-trts;pw<-rv$pairwise_data;if(!is.null(pw)){n1c<-NULL;n2c<-NULL;for(cn in names(pw)){if(grepl("^n1$",cn,ignore.case=TRUE)&&is.null(n1c))n1c<-cn;if(grepl("^n2$",cn,ignore.case=TRUE)&&is.null(n2c))n2c<-cn};if(!is.null(n1c)&&!is.null(n2c))for(j in seq_len(nrow(pw))){t1<-as.character(pw$treat1[j]);t2<-as.character(pw$treat2[j]);if(t1%in%trts)np[t1]<-np[t1]+as.numeric(pw[[n1c]][j]);if(t2%in%trts)np[t2]<-np[t2]+as.numeric(pw[[n2c]][j])}};if(sum(np,na.rm=TRUE)>0){ns<-sqrt(np[trts]);ns<-ns/max(ns)*2.5+0.5;ns<-ns*input$ng_cex;netgraph(nma_obj,number.of.studies=TRUE,cex.points=ns,multiarm=TRUE,plastic=FALSE,col="steelblue")}else netgraph(nma_obj,number.of.studies=TRUE,cex.points=input$ng_cex,multiarm=TRUE,plastic=FALSE,col="steelblue")})})
  observeEvent(input$runForest, {req(rv$nma_result);output$forest_plot<-renderPlot({forest(rv$nma_result,reference.group=input$ref_trt,sortvar=TE,smlab=paste("vs",input$ref_trt))},height=function()max(500,length(rv$nma_result$trts)*60))})

  # === netsplit テーブル ===
  output$netsplit_debug <- renderPrint({
    req(rv$netsplit_result,rv$nma_result);ns_obj<-rv$netsplit_result;nma_obj<-rv$nma_result
    cat("=== netsplit 内部構造デバッグ ===\n");cat("netsplit$comparison 数:",length(ns_obj$comparison),"\n");cat("netsplit$comparison[1:5]:",head(ns_obj$comparison,5),"\n\n")
    ns_n<-length(ns_obj$comparison);ns_dir<-ns_obj$direct.random$TE;ns_ind<-ns_obj$indirect.random$TE
    cat("netsplit内 direct.random$TE: 長さ=",length(ns_dir)," NA=",sum(is.na(ns_dir))," 非NA=",sum(!is.na(ns_dir)),"\n")
    cat("netsplit内 indirect.random$TE: 長さ=",length(ns_ind)," NA=",sum(is.na(ns_ind))," 非NA=",sum(!is.na(ns_ind)),"\n\n")
    trts<-sort(nma_obj$trts);n_all<-choose(length(trts),2)
    cat("全治療数:",length(trts),"\n全ペア数 (combn):",n_all,"\nnetsplit ペア数:",ns_n,"\n→ netsplit に含まれないペア数:",n_all-ns_n,"\n\n")
    hd<-!is.na(ns_dir);hi<-!is.na(ns_ind)
    cat("netsplit 内の Type 内訳:\n  Direct+Indirect:",sum(hd&hi),"\n  Direct only:    ",sum(hd&!hi),"\n  Indirect only:  ",sum(!hd&hi),"\n  Neither (NMA):  ",sum(!hd&!hi),"\n\n")
    cat("netsplit に含まれない",n_all-ns_n,"ペアは:\n  → 間接経路があれば 'Indirect only'\n  → なければ 'NMA only'\n")
  })

  output$netsplit_table <- DT::renderDataTable({
    req(rv$netsplit_result,rv$nma_result);ns_obj<-rv$netsplit_result;nma_obj<-rv$nma_result;ir<-is_ratio()
    trts<-sort(nma_obj$trts);all_pairs<-combn(trts,2,simplify=FALSE);n_all<-length(all_pairs)
    all_t1<-sapply(all_pairs,`[`,1);all_t2<-sapply(all_pairs,`[`,2);all_key<-paste0(all_t1,":",all_t2)
    all_nma_est<-all_nma_low<-all_nma_up<-rep(NA_real_,n_all)
    for(j in seq_len(n_all)){all_nma_est[j]<-nma_obj$TE.random[all_t1[j],all_t2[j]];all_nma_low[j]<-nma_obj$lower.random[all_t1[j],all_t2[j]];all_nma_up[j]<-nma_obj$upper.random[all_t1[j],all_t2[j]]}
    all_dir_est<-all_dir_low<-all_dir_up<-rep(NA_real_,n_all);all_ind_est<-all_ind_low<-all_ind_up<-rep(NA_real_,n_all);all_p<-rep(NA_real_,n_all);all_k<-rep(0L,n_all)
    ns_comps<-ns_obj$comparison;n_ns<-length(ns_comps)
    if(n_ns>0){
      ns_dir_est<-ns_obj$direct.random$TE;ns_dir_low<-ns_obj$direct.random$lower;ns_dir_up<-ns_obj$direct.random$upper
      ns_ind_est<-ns_obj$indirect.random$TE;ns_ind_low<-ns_obj$indirect.random$lower;ns_ind_up<-ns_obj$indirect.random$upper
      ns_p<-ns_obj$random$p;ns_k<-ns_obj$k
      ns_split<-strsplit(ns_comps,"\\s*vs\\s*|\\s*:\\s*");ns_t1_raw<-sapply(ns_split,function(x)trimws(x[1]));ns_t2_raw<-sapply(ns_split,function(x)trimws(x[2]))
      ns_sorted_keys<-sapply(seq_len(n_ns),function(j){s<-sort(c(ns_t1_raw[j],ns_t2_raw[j]));paste0(s[1],":",s[2])})
      for(j in seq_len(n_ns)){idx<-which(all_key==ns_sorted_keys[j]);if(length(idx)==0)next;i<-idx[1]
        flip<-(ns_t1_raw[j]==all_t2[i]&&ns_t2_raw[j]==all_t1[i])
        if(!is.na(ns_dir_est[j])){if(flip){all_dir_est[i]<- -ns_dir_est[j];all_dir_low[i]<- -ns_dir_up[j];all_dir_up[i]<- -ns_dir_low[j]}else{all_dir_est[i]<-ns_dir_est[j];all_dir_low[i]<-ns_dir_low[j];all_dir_up[i]<-ns_dir_up[j]}}
        if(!is.na(ns_ind_est[j])){if(flip){all_ind_est[i]<- -ns_ind_est[j];all_ind_low[i]<- -ns_ind_up[j];all_ind_up[i]<- -ns_ind_low[j]}else{all_ind_est[i]<-ns_ind_est[j];all_ind_low[i]<-ns_ind_low[j];all_ind_up[i]<-ns_ind_up[j]}}
        if(!is.na(ns_p[j]))all_p[i]<-ns_p[j];if(!is.na(ns_k[j]))all_k[i]<-as.integer(ns_k[j])}
    }
    for(j in seq_len(n_all)){if(is.na(all_dir_est[j])&&is.na(all_ind_est[j])&&!is.na(all_nma_est[j])){all_ind_est[j]<-all_nma_est[j];all_ind_low[j]<-all_nma_low[j];all_ind_up[j]<-all_nma_up[j]}}
    has_dir<-!is.na(all_dir_est);has_ind<-!is.na(all_ind_est)
    type_label<-dplyr::case_when(has_dir&has_ind~"Direct+Indirect",has_dir&!has_ind~"Direct",!has_dir&has_ind~"Indirect",TRUE~"NMA only")
    type_bg<-dplyr::case_when(type_label=="Direct"~"#4caf50",type_label=="Indirect"~"#ff9800",type_label=="Direct+Indirect"~"#2196f3",type_label=="NMA only"~"#9e9e9e",TRUE~"#cccccc")
    type_html<-sprintf('<span style="background:%s;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.85em;">%s</span>',type_bg,type_label)
    fmt_val_v<-function(est,lo,hi){sapply(seq_along(est),function(j){if(is.na(est[j]))return("\u2014");if(ir)sprintf("%.2f [%.2f, %.2f]",exp(est[j]),exp(lo[j]),exp(hi[j]))else sprintf("%.2f [%.2f, %.2f]",est[j],lo[j],hi[j])})}
    fmt_p<-function(p)ifelse(is.na(p),"\u2014",ifelse(p<0.001,"< 0.001",sprintf("%.3f",p)))
    df_fwd<-data.frame(Comparison=paste0(all_t1," vs ",all_t2),Direction="\u25B6",Type=type_html,k=all_k,Direct=fmt_val_v(all_dir_est,all_dir_low,all_dir_up),Indirect=fmt_val_v(all_ind_est,all_ind_low,all_ind_up),NMA=fmt_val_v(all_nma_est,all_nma_low,all_nma_up),p_value=fmt_p(all_p),is_reverse=FALSE,type_raw=type_label,stringsAsFactors=FALSE)
    df_rev<-data.frame(Comparison=paste0("\u25C0 ",all_t2," vs ",all_t1),Direction="\u25C0",Type=type_html,k=all_k,Direct=fmt_val_v(-all_dir_est,-all_dir_up,-all_dir_low),Indirect=fmt_val_v(-all_ind_est,-all_ind_up,-all_ind_low),NMA=fmt_val_v(-all_nma_est,-all_nma_up,-all_nma_low),p_value=fmt_p(all_p),is_reverse=TRUE,type_raw=type_label,stringsAsFactors=FALSE)
    row_list<-vector("list",n_all*2);for(j in seq_len(n_all)){row_list[[(j-1)*2+1]]<-df_fwd[j,];row_list[[(j-1)*2+2]]<-df_rev[j,]};df_all<-do.call(rbind,row_list)
    row_bg<-sapply(seq_len(nrow(df_all)),function(i){if(df_all$is_reverse[i])"#f0f4ff" else switch(df_all$type_raw[i],"Direct"="#e8f5e9","Indirect"="#fff3e0","Direct+Indirect"="#e3f2fd","NMA only"="#f5f5f5","#ffffff")})
    disp<-df_all[,c("Comparison","Direction","Type","k","Direct","Indirect","NMA","p_value")]
    dt<-DT::datatable(disp,escape=FALSE,rownames=FALSE,options=list(pageLength=100,scrollX=TRUE,dom="ftip",order=list()),callback=DT::JS(sprintf("var bg=%s;table.rows().every(function(i){$(this.node()).css('background-color',bg[i]);});",jsonlite::toJSON(row_bg,auto_unbox=FALSE))))
    dt%>%DT::formatStyle("p_value",color=DT::styleInterval(c(0.0499999),c("red","black")),fontWeight=DT::styleInterval(c(0.0499999),c("bold","normal")))
  })

  output$netsplit_forest<-renderPlot({req(rv$netsplit_result);tryCatch(forest(rv$netsplit_result),error=function(e){plot.new();text(0.5,0.5,e$message)})},height=function(){if(!is.null(rv$grade_table))max(600,nrow(rv$grade_table)*80) else 600})
  output$league_output<-renderPrint({req(rv$nma_result);netleague(rv$nma_result,digits=2)})
  output$rank_output<-renderPrint({req(rv$net_rank);print(rv$net_rank)})
  observeEvent(input$runFunnel,{req(rv$nma_result);output$funnel_plot<-renderPlot({funnel(rv$nma_result)})})

  # === A1 ===
  observeEvent(input$a1_bulk_apply, {req(rv$grade_table);gt<-rv$grade_table;if(input$a1_bulk_rob!="--")gt$a1_rob<-input$a1_bulk_rob;if(input$a1_bulk_inc!="--")gt$a1_inconsistency<-input$a1_bulk_inc;if(input$a1_bulk_ind!="--")gt$a1_indirectness<-input$a1_bulk_ind;if(input$a1_bulk_pub!="--")gt$a1_pub_bias<-input$a1_bulk_pub;for(i in seq_len(nrow(gt)))gt$a1_preliminary[i]<-calc_a1_preliminary(gt$a1_rob[i],gt$a1_inconsistency[i],gt$a1_indirectness[i],gt$a1_pub_bias[i]);rv$grade_table<-gt;showNotification("一括適用完了",type="message")})

  output$a1_comp_selector<-renderUI({req(rv$grade_table);gt<-rv$grade_table;idx<-which(gt$has_direct);if(length(idx)==0)return(div(class="warn-bx","直接比較のあるペアがありません"));choices<-setNames(idx,paste0(gt$treat1[idx]," vs ",gt$treat2[idx]," (k=",gt$direct_k[idx],")"));selectInput("a1_selected_comp","比較ペアを選択",choices=choices)})
  observeEvent(input$a1_selected_comp, {rv$a1_flipped<-FALSE})
  observeEvent(input$a1_flip_direction, {rv$a1_flipped<-!rv$a1_flipped})

  output$a1_direction_label<-renderUI({req(input$a1_selected_comp,rv$grade_table);i<-as.integer(input$a1_selected_comp);gt<-rv$grade_table
    if(rv$a1_flipped)div(style="display:inline-block;background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:6px 14px;font-weight:bold;color:#856404;",HTML(paste0("\u21C4 現在の方向: <b>",gt$treat2[i]," vs ",gt$treat1[i],"</b> (逆方向)")))
    else div(style="display:inline-block;background:#d4edda;border:1px solid #28a745;border-radius:6px;padding:6px 14px;font-weight:bold;color:#155724;",HTML(paste0("\u25B6 現在の方向: <b>",gt$treat1[i]," vs ",gt$treat2[i],"</b> (元の方向)")))
  })

  # A1 選択情報: ns_direct_TE を使用
  output$a1_selected_info <- renderUI({
    req(input$a1_selected_comp, rv$grade_table)
    i <- as.integer(input$a1_selected_comp); gt <- rv$grade_table; flipped <- rv$a1_flipped
    if (flipped) {
      ns_te <- -gt$ns_direct_TE[i]; ns_lo <- -gt$ns_direct_upper[i]; ns_up <- -gt$ns_direct_lower[i]
      pw_te <- -gt$pw_direct_TE[i]; pw_lo <- -gt$pw_direct_upper[i]; pw_up <- -gt$pw_direct_lower[i]
      lbl <- paste0("<b>", gt$treat2[i], " vs ", gt$treat1[i], "</b> (逆方向)")
    } else {
      ns_te <- gt$ns_direct_TE[i]; ns_lo <- gt$ns_direct_lower[i]; ns_up <- gt$ns_direct_upper[i]
      pw_te <- gt$pw_direct_TE[i]; pw_lo <- gt$pw_direct_lower[i]; pw_up <- gt$pw_direct_upper[i]
      lbl <- paste0("<b>", gt$treat1[i], " vs ", gt$treat2[i], "</b>")
    }
    tagList(
      HTML(paste0(lbl, " | k=", gt$direct_k[i])),
      div(style="margin-top:4px;font-size:13px;", HTML(paste0("<b>Direct (netsplit): ", fmt_est(ns_te, ns_lo, ns_up), "</b>"))),
      div(style="font-size:11px;color:#666;", HTML(paste0("Pairwise metagen: ", fmt_est(pw_te, pw_lo, pw_up))))
    )
  })

  # A1 メタ解析表示: ns_direct_TE を注記に使用
  output$a1_meta_display <- renderUI({
    req(input$a1_selected_comp, rv$grade_table)
    i <- as.integer(input$a1_selected_comp); gt <- rv$grade_table
    t1 <- gt$treat1[i]; t2 <- gt$treat2[i]; comp_key <- paste(sort(c(t1,t2)),collapse=":"); flipped <- rv$a1_flipped
    mr <- rv$pairwise_meta_results
    if(is.null(mr)||!comp_key%in%names(mr)||is.null(mr[[comp_key]])) return(div(class="warn-bx","このペアのメタ解析結果がありません"))
    m <- mr[[comp_key]]
    i2_val<-if(!is.null(m$I2))sprintf("%.1f%%",m$I2*100)else"N/A";tau2_val<-if(!is.null(m$tau2))sprintf("%.4f",m$tau2)else"N/A"
    q_val<-if(!is.null(m$Q))sprintf("%.2f",m$Q)else"N/A";q_p<-if(!is.null(m$pval.Q))sprintf("%.4f",m$pval.Q)else"N/A"
    if(flipped){
      re_te<- -m$TE.random;re_lo<- -m$upper.random;re_up<- -m$lower.random
      if(is_ratio())re<-sprintf("%.2f [%.2f, %.2f]",exp(re_te),exp(re_lo),exp(re_up))else re<-sprintf("%.2f [%.2f, %.2f]",re_te,re_lo,re_up)
      label_t1<-t2;label_t2<-t1
      ns_te<- -gt$ns_direct_TE[i];ns_lo<- -gt$ns_direct_upper[i];ns_up<- -gt$ns_direct_lower[i]
    } else {
      if(is_ratio())re<-sprintf("%.2f [%.2f, %.2f]",exp(m$TE.random),exp(m$lower.random),exp(m$upper.random))else re<-sprintf("%.2f [%.2f, %.2f]",m$TE.random,m$lower.random,m$upper.random)
      label_t1<-t1;label_t2<-t2
      ns_te<-gt$ns_direct_TE[i];ns_lo<-gt$ns_direct_lower[i];ns_up<-gt$ns_direct_upper[i]
    }
    ns_str <- fmt_est(ns_te, ns_lo, ns_up)
    dir_badge<-if(flipped)'<span style="background:#ffc107;color:#000;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:bold;">\u21C4 逆方向</span>' else ""
    tagList(
      div(class="a1-comp-hd",HTML(paste0("\u25B6 ",label_t1," vs ",label_t2," (k=",m$k,") ",dir_badge))),
      div(class="a1-meta-box",tags$table(style="width:100%;",tags$tr(tags$td(paste0("k=",m$k)),tags$td(HTML(paste0("I&sup2;=",i2_val))),tags$td(paste0("Q=",q_val," (p=",q_p,")")),tags$td(HTML(paste0("&tau;&sup2;=",tau2_val)))),tags$tr(tags$td(colspan="4",paste0("Pairwise RE推定値: ",re))))),
      div(class="ns-direct-note",HTML(paste0("\u203B <b>netsplit Direct推定値: ",ns_str,"</b>","<br><span style='font-size:11px;color:#666;'>（グラフ理論的分離。ネットワーク全体の\u03C4\u00B2・multi-arm補正を反映。上のフォレストプロットは通常のペアワイズ metagen() による値）</span>"))),
      plotOutput("a1_forest_current",height=paste0(max(250,m$k*40+120),"px"))
    )
  })

  output$a1_forest_current<-renderPlot({
    req(input$a1_selected_comp,rv$grade_table);i<-as.integer(input$a1_selected_comp);gt<-rv$grade_table;t1<-gt$treat1[i];t2<-gt$treat2[i];comp_key<-paste(sort(c(t1,t2)),collapse=":");flipped<-rv$a1_flipped
    mr<-rv$pairwise_meta_results;pw<-rv$pairwise_data;if(is.null(pw)){plot.new();text(0.5,0.5,"データなし");return()}
    pw$t1c<-as.character(pw$treat1);pw$t2c<-as.character(pw$treat2);pw$comp_key_sorted<-apply(pw[,c("t1c","t2c")],1,function(r)paste(sort(r),collapse=":"))
    sub<-pw[pw$comp_key_sorted==comp_key,];if(nrow(sub)==0){plot.new();text(0.5,0.5,"メタ解析結果なし");return()}
    if(flipped){sub_flip<-sub;sub_flip$TE<- -sub_flip$TE;label<-paste0(t2," vs ",t1," (\u21C4 逆方向)");tryCatch({m_flip<-metagen(TE=sub_flip$TE,seTE=sub_flip$seTE,studlab=sub_flip$studlab,sm=input$effect_measure,random=TRUE,common=TRUE);forest(m_flip,sortvar=TE,smlab=label,print.tau2=TRUE,print.I2=TRUE,print.pval.Q=TRUE)},error=function(e){plot.new();text(0.5,0.5,e$message)})}
    else{if(!is.null(mr)&&comp_key%in%names(mr)&&!is.null(mr[[comp_key]])){m<-mr[[comp_key]];tryCatch({forest(m,sortvar=TE,smlab=paste0(t1," vs ",t2),print.tau2=TRUE,print.I2=TRUE,print.pval.Q=TRUE)},error=function(e){plot.new();text(0.5,0.5,e$message)})}else{plot.new();text(0.5,0.5,"メタ解析結果なし")}}
  })

  output$a1_grade_inputs<-renderUI({req(input$a1_selected_comp,rv$grade_table);i<-as.integer(input$a1_selected_comp);gt<-rv$grade_table;fluidRow(column(3,selectInput("a1_cur_rob","RoB",GRADE3,gt$a1_rob[i])),column(3,selectInput("a1_cur_inc","非一貫性",GRADE3,gt$a1_inconsistency[i])),column(3,selectInput("a1_cur_ind","非直接性",GRADE3,gt$a1_indirectness[i])),column(3,selectInput("a1_cur_pub","出版バイアス",PUB_CHOICES,gt$a1_pub_bias[i])))})
  observeEvent(input$a1_save_current, {req(input$a1_selected_comp,rv$grade_table);i<-as.integer(input$a1_selected_comp);rv$grade_table$a1_rob[i]<-input$a1_cur_rob;rv$grade_table$a1_inconsistency[i]<-input$a1_cur_inc;rv$grade_table$a1_indirectness[i]<-input$a1_cur_ind;rv$grade_table$a1_pub_bias[i]<-input$a1_cur_pub;rv$grade_table$a1_preliminary[i]<-calc_a1_preliminary(input$a1_cur_rob,input$a1_cur_inc,input$a1_cur_ind,input$a1_cur_pub);showNotification(paste0(rv$grade_table$comparison[i]," の評価を保存しました"),type="message")})

  # A1 評価一覧: ns_direct を使用
  output$a1_summary_table <- DT::renderDataTable({
    req(rv$grade_table); gt <- rv$grade_table; idx <- which(gt$has_direct)
    if (length(idx) == 0) return(DT::datatable(data.frame(Info="直接比較なし")))
    df <- data.frame(
      Comparison = paste0(gt$treat1[idx], " vs ", gt$treat2[idx]),
      k = gt$direct_k[idx],
      `Direct(netsplit)` = sapply(idx, function(i) fmt_est(gt$ns_direct_TE[i], gt$ns_direct_lower[i], gt$ns_direct_upper[i])),
      `Direct(pairwise)` = sapply(idx, function(i) fmt_est(gt$pw_direct_TE[i], gt$pw_direct_lower[i], gt$pw_direct_upper[i])),
      RoB = gt$a1_rob[idx], Inconsistency = gt$a1_inconsistency[idx], Indirectness = gt$a1_indirectness[idx],
      Pub_Bias = gt$a1_pub_bias[idx], Preliminary = gt$a1_preliminary[idx],
      stringsAsFactors = FALSE, check.names = FALSE)
    DT::datatable(df, options = list(pageLength=50, dom="t", scrollX=TRUE), rownames=FALSE) %>%
      DT::formatStyle("Preliminary", backgroundColor = DT::styleEqual(c("High","Moderate","Low","Very Low"), c("#c8e6c9","#bbdefb","#fff9c4","#ffcdd2")))
  })

  # === A2 ===
  observeEvent(input$a2_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in which(gt$has_direct)){res<-auto_imprecision(gt$direct_lower[i],gt$direct_upper[i],input$mid_lower,input$mid_upper,is_ratio(),input$effect_measure);gt$a2_imprecision[i]<-res$result;gt$a2_reason[i]<-res$reason;gt$a2_final_direct[i]<-calc_downgrade(gt$a1_preliminary[i],res$result)};rv$grade_table<-gt;showNotification("A2完了",type="message")})
  output$a2_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;ir<-is_ratio();idx<-which(gt$has_direct);if(length(idx)==0)return(NULL);lapply(idx,function(i){
    fin<-calc_downgrade(gt$a1_preliminary[i],gt$a2_imprecision[i]);rv$grade_table$a2_final_direct[i]<<-fin
    ois_html <- ""
    if (ir && !is.na(gt$direct_lower[i]) && !is.na(gt$direct_upper[i])) {
      ci_ratio <- exp(gt$direct_upper[i]) / exp(gt$direct_lower[i])
      ois_thresh <- if (input$effect_measure == "OR") 2.5 else 3.0
      ois_color <- if (ci_ratio > ois_thresh) "#dc3545" else if (ci_ratio > ois_thresh * 0.8) "#ff9800" else "#28a745"
      ois_html <- sprintf(" | <span style='font-size:11px;'>CI ratio: <b style='color:%s;'>%.2f</b> (OIS閾値: %.1f)</span>", ois_color, ci_ratio, ois_thresh)
    }
    div(class="gc",h5(HTML(paste0(gt$treat1[i]," vs ",gt$treat2[i],"  ",fmt_est(gt$direct_TE[i],gt$direct_lower[i],gt$direct_upper[i]),ois_html))),
      fluidRow(column(4,selectInput(paste0("a2_",i),"不精確さ",GRADE4,gt$a2_imprecision[i])),column(4,div(style="margin-top:25px;",HTML(cert_badge_html(fin)))),column(4,div(class="reason-text",gt$a2_reason[i]))))})})
  observe({req(rv$grade_table);for(i in which(rv$grade_table$has_direct)){local({ii<-i;observeEvent(input[[paste0("a2_",ii)]],{rv$grade_table$a2_imprecision[ii]<-input[[paste0("a2_",ii)]];rv$grade_table$a2_final_direct[ii]<-calc_downgrade(rv$grade_table$a1_preliminary[ii],rv$grade_table$a2_imprecision[ii])},ignoreInit=TRUE)})}})

  # === B1 ===
  output$b1_cards<-renderUI({req(rv$grade_table,rv$treatments);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){t1<-gt$treat1[i];t2<-gt$treat2[i];others<-setdiff(rv$treatments,c(t1,t2));div(class="gc",h5(paste0(t1," vs ",t2)),fluidRow(column(4,selectInput(paste0("b1c_",i),"共通比較対象",c("(選択)"="",others),gt$b1_common[i])),column(5,uiOutput(paste0("b1info_",i))),column(3,div(style="margin-top:25px;",HTML(paste0("開始: ",cert_badge_html(gt$b1_start[i])))))))})})
  observe({req(rv$grade_table,rv$treatments);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;output[[paste0("b1info_",ii)]]<-renderUI({common<-input[[paste0("b1c_",ii)]];if(is.null(common)||common=="")return(div("選択してください"));gt<-rv$grade_table;t1<-gt$treat1[ii];t2<-gt$treat2[ii];find_info<-function(ta,tb){idx<-which((gt$treat1==ta&gt$treat2==tb)|(gt$treat1==tb&gt$treat2==ta));cert<-if(length(idx)>0)gt$a1_preliminary[idx[1]]else NA_character_;k<-if(length(idx)>0)gt$direct_k[idx[1]]else 0L;ck<-paste(sort(c(ta,tb)),collapse=":");tn<-NA;if(!is.null(rv$pw_study_counts)&&ck%in%names(rv$pw_study_counts))tn<-rv$pw_study_counts[[ck]]$total_n;list(cert=cert,k=k,total_n=tn)};i1<-find_info(t1,common);i2<-find_info(common,t2);rv$grade_table$b1_common[ii]<<-common;rv$grade_table$b1_cert1[ii]<<-i1$cert;rv$grade_table$b1_cert2[ii]<<-i2$cert;c1n<-cert2num(i1$cert);c2n<-cert2num(i2$cert);if(!is.na(c1n)&&!is.na(c2n))rv$grade_table$b1_start[ii]<<-num2cert(min(c1n,c2n));n1s<-if(!is.na(i1$total_n)&&i1$total_n>0)paste0(", N=",i1$total_n)else"";n2s<-if(!is.na(i2$total_n)&&i2$total_n>0)paste0(", N=",i2$total_n)else"";div(div(paste0(t1,":",common," (k=",i1$k,n1s,") → "),HTML(cert_badge_html(i1$cert))),div(paste0(common,":",t2," (k=",i2$k,n2s,") → "),HTML(cert_badge_html(i2$cert))))})})}})

  # === B2 ===
  output$b2_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){prel<-calc_downgrade(gt$b1_start[i],gt$b2_intransitivity[i]);rv$grade_table$b2_preliminary[i]<<-prel;div(class="gc",h5(paste0(gt$treat1[i]," vs ",gt$treat2[i])),fluidRow(column(4,selectInput(paste0("b2_",i),"非推移性",GRADE3,gt$b2_intransitivity[i])),column(4,div(style="margin-top:25px;",HTML(cert_badge_html(prel))))))})})
  observe({req(rv$grade_table);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;observeEvent(input[[paste0("b2_",ii)]],{rv$grade_table$b2_intransitivity[ii]<-input[[paste0("b2_",ii)]];rv$grade_table$b2_preliminary[ii]<-calc_downgrade(rv$grade_table$b1_start[ii],rv$grade_table$b2_intransitivity[ii])},ignoreInit=TRUE)})}})

  # === B3 ===
  observeEvent(input$b3_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){if(!is.na(gt$indirect_TE[i])){res<-auto_imprecision(gt$indirect_lower[i],gt$indirect_upper[i],input$mid_lower,input$mid_upper,is_ratio(),input$effect_measure);gt$b3_imprecision[i]<-res$result;gt$b3_reason[i]<-res$reason;gt$b3_final_indirect[i]<-calc_downgrade(gt$b2_preliminary[i],res$result)}};rv$grade_table<-gt;showNotification("B3完了",type="message")})
  output$b3_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){if(is.na(gt$indirect_TE[i]))return(NULL);fin<-calc_downgrade(gt$b2_preliminary[i],gt$b3_imprecision[i]);rv$grade_table$b3_final_indirect[i]<<-fin;div(class="gc",h5(paste0(gt$treat1[i]," vs ",gt$treat2[i])),fluidRow(column(4,selectInput(paste0("b3_",i),"不精確さ",GRADE4,gt$b3_imprecision[i])),column(4,div(style="margin-top:25px;",HTML(cert_badge_html(fin)))),column(4,div(class="reason-text",gt$b3_reason[i]))))})})

  # === C1 ===
  observeEvent(input$c1_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){res<-auto_c1(gt[i,]);gt$c1_dominant[i]<-res$dom;gt$c1_source[i]<-res$src;gt$c1_certainty[i]<-res$cert;gt$c1_reason[i]<-res$reason};rv$grade_table<-gt;showNotification("C1完了",type="message")})
  output$c1_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){div(class="gc",h5(paste0(gt$treat1[i]," vs ",gt$treat2[i])),fluidRow(column(3,div("A1:",HTML(cert_badge_html(gt$a1_preliminary[i])))),column(3,div("B2:",HTML(cert_badge_html(gt$b2_preliminary[i])))),column(3,selectInput(paste0("c1s_",i),"出発点",c("Direct","Indirect"),ifelse(is.na(gt$c1_source[i]),"Direct",gt$c1_source[i]))),column(3,div(style="margin-top:25px;",HTML(cert_badge_html(gt$c1_certainty[i]))))),div(class="reason-text",gt$c1_reason[i]))})})
  observe({req(rv$grade_table);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;observeEvent(input[[paste0("c1s_",ii)]],{src<-input[[paste0("c1s_",ii)]];rv$grade_table$c1_source[ii]<-src;rv$grade_table$c1_certainty[ii]<-if(src=="Direct")rv$grade_table$a1_preliminary[ii]else rv$grade_table$b2_preliminary[ii]},ignoreInit=TRUE)})}})

  # === C2 ===
  observeEvent(input$c2_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){res<-auto_c2(gt[i,]);gt$c2_incoherence[i]<-res$result;gt$c2_reason[i]<-res$reason};rv$grade_table<-gt;showNotification("C2完了",type="message")})
  observeEvent(input$c2_bulk_ns, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){if(!is.na(gt$netsplit_p[i])&&gt$netsplit_p[i]>=0.10){gt$c2_incoherence[i]<-"Not serious";gt$c2_reason[i]<-"p>=0.10"}};rv$grade_table<-gt})
  output$c2_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){pt<-if(is.na(gt$netsplit_p[i]))"---"else sprintf("%.3f",gt$netsplit_p[i]);ps<-if(!is.na(gt$netsplit_p[i])&&gt$netsplit_p[i]<0.05)"color:red;font-weight:bold;"else"";div(class="gc",h5(HTML(paste0(gt$treat1[i]," vs ",gt$treat2[i]," | p=<span style='",ps,"'>",pt,"</span>"))),fluidRow(column(4,selectInput(paste0("c2_",i),"非整合性",GRADE3,gt$c2_incoherence[i])),column(8,div(class="reason-text",gt$c2_reason[i]))))})})
  observe({req(rv$grade_table);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;observeEvent(input[[paste0("c2_",ii)]],{rv$grade_table$c2_incoherence[ii]<-input[[paste0("c2_",ii)]]},ignoreInit=TRUE)})}})

  # === C3 ===
  observeEvent(input$c3_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){res<-auto_imprecision(gt$nma_lower[i],gt$nma_upper[i],input$mid_lower,input$mid_upper,is_ratio(),input$effect_measure);gt$c3_imprecision[i]<-res$result;gt$c3_reason[i]<-res$reason;gt$c3_final_nma[i]<-calc_downgrade(gt$c1_certainty[i],gt$c2_incoherence[i],res$result)};rv$grade_table<-gt;showNotification("C3完了",type="message")})
  output$c3_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;ir<-is_ratio();lapply(seq_len(nrow(gt)),function(i){
    fin<-calc_downgrade(gt$c1_certainty[i],gt$c2_incoherence[i],gt$c3_imprecision[i]);rv$grade_table$c3_final_nma[i]<<-fin
    ois_html <- ""
    if (ir && !is.na(gt$nma_lower[i]) && !is.na(gt$nma_upper[i])) {
      ci_ratio <- exp(gt$nma_upper[i]) / exp(gt$nma_lower[i])
      ois_thresh <- if (input$effect_measure == "OR") 2.5 else 3.0
      ois_color <- if (ci_ratio > ois_thresh) "#dc3545" else if (ci_ratio > ois_thresh * 0.8) "#ff9800" else "#28a745"
      ois_html <- sprintf("<span style='font-size:11px;'> | CI ratio: <b style='color:%s;'>%.2f</b> (OIS閾値: %.1f)</span>", ois_color, ci_ratio, ois_thresh)
    }
    abs_html <- ""
    if (ir && !is.na(gt$nma_TE[i]) && !is.na(input$baseline_risk) && input$baseline_risk > 0) {
      br <- input$baseline_risk; sm <- input$effect_measure
      if (sm %in% c("RR","HR")) { rd <- br * (exp(gt$nma_TE[i]) - 1)
      } else if (sm == "OR") { rd <- (br * exp(gt$nma_TE[i])) / (1 - br + br * exp(gt$nma_TE[i])) - br
      } else { rd <- NA }
      if (!is.na(rd)) abs_html <- sprintf(" | <b>絶対効果: %+d/1000</b>", round(rd * 1000))
    }
    div(class="gc",h5(HTML(paste0(gt$treat1[i]," vs ",gt$treat2[i],"  ",fmt_est(gt$nma_TE[i],gt$nma_lower[i],gt$nma_upper[i]),abs_html,ois_html))),
      fluidRow(column(3,selectInput(paste0("c3_",i),"不精確さ",GRADE4,gt$c3_imprecision[i])),
        column(9,div(style="margin-top:25px;",HTML(paste0(cert_badge_html(gt$c1_certainty[i])," → C2(",gt$c2_incoherence[i],") → C3(",gt$c3_imprecision[i],") → ",cert_badge_html(fin)))))),
      div(class="reason-text",gt$c3_reason[i]))})})
  observe({req(rv$grade_table);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;observeEvent(input[[paste0("c3_",ii)]],{rv$grade_table$c3_imprecision[ii]<-input[[paste0("c3_",ii)]];rv$grade_table$c3_final_nma[ii]<-calc_downgrade(rv$grade_table$c1_certainty[ii],rv$grade_table$c2_incoherence[ii],rv$grade_table$c3_imprecision[ii])},ignoreInit=TRUE)})}})

  # === D ===
  observeEvent(input$d_auto_run, {req(rv$grade_table);gt<-rv$grade_table;for(i in seq_len(nrow(gt))){res<-auto_d(gt[i,]);gt$d_best[i]<-res$best;gt$d_final_certainty[i]<-res$cert;gt$d_reason[i]<-res$reason;if(res$best=="Direct"){gt$d_effect_TE[i]<-gt$direct_TE[i];gt$d_effect_lower[i]<-gt$direct_lower[i];gt$d_effect_upper[i]<-gt$direct_upper[i]}else if(res$best=="Indirect"){gt$d_effect_TE[i]<-gt$indirect_TE[i];gt$d_effect_lower[i]<-gt$indirect_lower[i];gt$d_effect_upper[i]<-gt$indirect_upper[i]}else{gt$d_effect_TE[i]<-gt$nma_TE[i];gt$d_effect_lower[i]<-gt$nma_lower[i];gt$d_effect_upper[i]<-gt$nma_upper[i]};gt$color_category[i]<-classify_color(gt$d_final_certainty[i],gt$d_effect_TE[i],gt$d_effect_lower[i],gt$d_effect_upper[i],is_ratio(),input$small_good,input$large_effect)};rv$grade_table<-gt;showNotification("D完了",type="message")})
  output$d_warn_a2b3<-renderUI({req(rv$grade_table);gt<-rv$grade_table;needs<-which(gt$c2_incoherence%in%c("Serious","Very serious"));if(length(needs)==0)return(NULL);a2t<-sum(gt$a2_imprecision[needs]=="Not assessed");b3t<-sum(gt$b3_imprecision[needs]=="Not assessed");if(a2t==0&&b3t==0)return(NULL);div(class="err-bx",HTML(paste0("<b>注意:</b> ",length(needs),"件で非整合性あり。A2(",a2t,") B3(",b3t,")")),actionButton("goto_a2","A2へ",class="btn-warning btn-sm"),actionButton("goto_b3","B3へ",class="btn-warning btn-sm"))})
  observeEvent(input$goto_a2,{updateTabItems(session,"tabs","tab_a2")});observeEvent(input$goto_b3,{updateTabItems(session,"tabs","tab_b3")})
  output$d_cards<-renderUI({req(rv$grade_table);gt<-rv$grade_table;lapply(seq_len(nrow(gt)),function(i){div(class="gc",h5(paste0(gt$treat1[i]," vs ",gt$treat2[i])),fluidRow(column(3,selectInput(paste0("db_",i),"推定値",c("NMA","Direct","Indirect"),gt$d_best[i])),column(3,div(style="margin-top:25px;",HTML(cert_badge_html(gt$d_final_certainty[i])))),column(6,div(style="margin-top:25px;",fmt_est(gt$d_effect_TE[i],gt$d_effect_lower[i],gt$d_effect_upper[i])))),div(class="reason-text",gt$d_reason[i]))})})
  observe({req(rv$grade_table);for(i in seq_len(nrow(rv$grade_table))){local({ii<-i;observeEvent(input[[paste0("db_",ii)]],{b<-input[[paste0("db_",ii)]];gt<-rv$grade_table;rv$grade_table$d_best[ii]<-b;if(b=="Direct"){rv$grade_table$d_final_certainty[ii]<-gt$a2_final_direct[ii];rv$grade_table$d_effect_TE[ii]<-gt$direct_TE[ii];rv$grade_table$d_effect_lower[ii]<-gt$direct_lower[ii];rv$grade_table$d_effect_upper[ii]<-gt$direct_upper[ii]}else if(b=="Indirect"){rv$grade_table$d_final_certainty[ii]<-gt$b3_final_indirect[ii];rv$grade_table$d_effect_TE[ii]<-gt$indirect_TE[ii];rv$grade_table$d_effect_lower[ii]<-gt$indirect_lower[ii];rv$grade_table$d_effect_upper[ii]<-gt$indirect_upper[ii]}else{rv$grade_table$d_final_certainty[ii]<-gt$c3_final_nma[ii];rv$grade_table$d_effect_TE[ii]<-gt$nma_TE[ii];rv$grade_table$d_effect_lower[ii]<-gt$nma_lower[ii];rv$grade_table$d_effect_upper[ii]<-gt$nma_upper[ii]}},ignoreInit=TRUE)})}})

  # === SoF ===
  output$sof_table<-DT::renderDataTable({req(rv$grade_table);gt<-rv$grade_table;if(isTRUE(input$sof_ref_only)&&!is.null(input$ref_trt))gt<-gt[gt$treat1==input$ref_trt|gt$treat2==input$ref_trt,];if(nrow(gt)==0)return(DT::datatable(data.frame(Info="なし")))
    br <- input$baseline_risk; ir_val <- is_ratio()
    abs_effect <- sapply(seq_len(nrow(gt)), function(i) {
      te <- gt$d_effect_TE[i]; lo <- gt$d_effect_lower[i]; up <- gt$d_effect_upper[i]
      if (is.na(te)) return("---")
      if (ir_val && !is.na(br) && br > 0) {
        sm <- input$effect_measure
        if (sm %in% c("RR","HR")) {
          rd <- br * (exp(te) - 1); rd_lo <- br * (exp(lo) - 1); rd_up <- br * (exp(up) - 1)
        } else if (sm == "OR") {
          p_int <- function(or_v) (br * or_v) / (1 - br + br * or_v) - br
          rd <- p_int(exp(te)); rd_lo <- p_int(exp(lo)); rd_up <- p_int(exp(up))
        } else { return("---") }
        sprintf("%+d (%+d, %+d)/1000", round(rd*1000), round(rd_lo*1000), round(rd_up*1000))
      } else if (!ir_val) {
        sprintf("%.2f (%.2f, %.2f)", te, lo, up)
      } else { "---" }
    })
    df<-data.frame(Comparison=gt$comparison,Source=gt$d_best,
      `Relative Effect`=sapply(seq_len(nrow(gt)),function(i)fmt_est(gt$d_effect_TE[i],gt$d_effect_lower[i],gt$d_effect_upper[i])),
      `Absolute Effect`=abs_effect,
      Certainty=gt$d_final_certainty,Category=gt$color_category,stringsAsFactors=FALSE,check.names=FALSE)
    DT::datatable(df,escape=FALSE,rownames=FALSE,options=list(pageLength=50,dom="t"))%>%DT::formatStyle("Certainty",backgroundColor=DT::styleEqual(c("High","Moderate","Low","Very Low"),c("#c8e6c9","#bbdefb","#fff9c4","#ffcdd2")))})
  output$dl_sof_xlsx<-downloadHandler(filename=function()paste0("SoF_",Sys.Date(),".xlsx"),content=function(file){req(rv$grade_table);wb<-createWorkbook();addWorksheet(wb,"SoF");writeData(wb,"SoF",rv$grade_table);saveWorkbook(wb,file,overwrite=TRUE)})

  # === 色分け図 ===
  observeEvent(input$gen_color_fig, {req(rv$grade_table,input$ref_trt);gt<-rv$grade_table;ri<-which(gt$treat1==input$ref_trt|gt$treat2==input$ref_trt);if(length(ri)==0){showNotification("比較なし",type="warning");return()};gtr<-gt[ri,];gtr$treatment<-ifelse(gtr$treat1==input$ref_trt,gtr$treat2,gtr$treat1);for(i in seq_len(nrow(gtr)))gtr$color_category[i]<-classify_color(gtr$d_final_certainty[i],gtr$d_effect_TE[i],gtr$d_effect_lower[i],gtr$d_effect_upper[i],is_ratio(),input$small_good,input$large_effect);gtr$display<-sapply(seq_len(nrow(gtr)),function(i){if(is.na(gtr$d_effect_TE[i]))"---"else if(is_ratio())sprintf("%.2f\n[%.2f, %.2f]",exp(gtr$d_effect_TE[i]),exp(gtr$d_effect_lower[i]),exp(gtr$d_effect_upper[i]))else sprintf("%.2f\n[%.2f, %.2f]",gtr$d_effect_TE[i],gtr$d_effect_lower[i],gtr$d_effect_upper[i])});gtr$sig<-!(gtr$d_effect_lower<=0&gtr$d_effect_upper>=0);gtr$sig[is.na(gtr$sig)]<-FALSE;co<-c("among_best_hm"=1,"intermediate_hm"=2,"among_best_lv"=3,"intermediate_lv"=4,"among_worst_lv"=5,"among_worst_hm"=6,"uncertain"=7);gtr$ov<-co[gtr$color_category];gtr$ov[is.na(gtr$ov)]<-99;gtr<-gtr[order(gtr$ov),];gtr$treatment<-factor(gtr$treatment,levels=rev(gtr$treatment));mp<-function(cf){gtr$fc<-sapply(gtr$color_category,cf);gtr$tc<-sapply(gtr$fc,text_col_for);ggplot(gtr,aes(x=1,y=treatment))+geom_tile(aes(fill=fc),color="white",linewidth=1.5)+geom_text(aes(label=display,fontface=ifelse(sig,"bold","plain"),color=tc),size=3.2,lineheight=0.85)+scale_fill_identity()+scale_color_identity()+labs(x=input$outcome_name,y=NULL,title=paste0(input$outcome_name," vs ",input$ref_trt))+theme_minimal(base_size=12)+theme(panel.grid=element_blank(),axis.text.y=element_text(size=11,face="bold"),axis.text.x=element_blank(),axis.ticks.x=element_blank())};rv$color_plot_data<-gtr;rv$color_plot_func<-mp;output$color_fig_output<-renderUI({ver<-input$color_version;pl<-tagList();if(ver%in%c("gradient","both")){output$color_grad_plot<-renderPlot({mp(gradient_fill)},height=function()max(300,nrow(gtr)*55+100));pl<-tagList(pl,plotOutput("color_grad_plot"))};if(ver%in%c("stoplight","both")){output$color_stop_plot<-renderPlot({mp(stoplight_fill)},height=function()max(300,nrow(gtr)*55+100));pl<-tagList(pl,plotOutput("color_stop_plot"))};pl})})
  output$dl_color_png<-downloadHandler(filename=function()paste0("color_",Sys.Date(),".png"),content=function(file){req(rv$color_plot_func,rv$color_plot_data);ggsave(file,rv$color_plot_func(gradient_fill),width=8,height=max(4,nrow(rv$color_plot_data)*0.6+2),dpi=150)})

  # === 保存/読込 ===
  output$save_json<-downloadHandler(filename=function()paste0("nma_grade_",Sys.Date(),".json"),content=function(file){writeLines(toJSON(list(project_name=input$proj_name,outcome_name=input$outcome_name,effect_measure=input$effect_measure,mid_lower=input$mid_lower,mid_upper=input$mid_upper,large_effect=input$large_effect,baseline_risk=input$baseline_risk,grade_table=rv$grade_table),pretty=TRUE,auto_unbox=TRUE),file)})
  observeEvent(input$load_json_btn,{req(input$load_json);tryCatch({proj<-fromJSON(input$load_json$datapath);rv$grade_table<-as.data.frame(proj$grade_table,stringsAsFactors=FALSE);if(!is.null(proj$project_name))updateTextInput(session,"proj_name",value=proj$project_name);showNotification("読込完了",type="message")},error=function(e)showNotification(e$message,type="error"))})
  output$dl_all_xlsx<-downloadHandler(filename=function()paste0("NMA_GRADE_",Sys.Date(),".xlsx"),content=function(file){wb<-createWorkbook();addWorksheet(wb,"GRADE");if(!is.null(rv$grade_table))writeData(wb,"GRADE",rv$grade_table);addWorksheet(wb,"Settings");writeData(wb,"Settings",data.frame(S=c("Project","Outcome","Measure"),V=c(input$proj_name,input$outcome_name,input$effect_measure)));saveWorkbook(wb,file,overwrite=TRUE)})
}

cat("\n========================================\n")
cat("NMA-GRADE v1.0\n")
cat("========================================\n\n")
shinyApp(ui=ui, server=server)
