import React, { useState, useMemo, useRef } from 'react';
import { TrendingUp, AlertCircle, Award, Target, BarChart3, RefreshCw, ChevronRight, Info, ChevronDown, ChevronUp, Calculator, AlertTriangle, FileText, Eye, EyeOff, BookOpen, Shield, ExternalLink, Lightbulb, Star, Users, TrendingDown, Database, Upload, Download, Check, X } from 'lucide-react';
import * as XLSX from 'xlsx';

// ======================================================
// 38指標マスタ（CRD準拠）
// ======================================================
const INDICATORS_38 = [
  // ①総合収益性
  { code: 'totalCapitalOperatingProfitRatio', name: '総資本営業利益率', category: '①総合収益性', formula: '営業利益÷総資本×100', unit: '%', higherIsBetter: true },
  { code: 'totalCapitalOrdinaryProfitRatio', name: '総資本経常利益率', category: '①総合収益性', formula: '経常利益÷総資本×100', unit: '%', higherIsBetter: true },
  { code: 'roa', name: '総資本当期純利益率(ROA)', category: '①総合収益性', formula: '当期純利益÷総資本×100', unit: '%', higherIsBetter: true },
  { code: 'roe', name: '自己資本当期純利益率(ROE)', category: '①総合収益性', formula: '当期純利益÷自己資本×100', unit: '%', higherIsBetter: true },
  { code: 'laborProductivity', name: '労働生産性', category: '①総合収益性', formula: '付加価値額÷従業員数', unit: '千円', higherIsBetter: true },
  // ②売上高利益
  { code: 'grossProfitRatio', name: '売上高総利益率', category: '②売上高利益', formula: '売上総利益÷売上高×100', unit: '%', higherIsBetter: true },
  { code: 'operatingProfitRatio', name: '売上高営業利益率', category: '②売上高利益', formula: '営業利益÷売上高×100', unit: '%', higherIsBetter: true },
  { code: 'ordinaryProfitRatio', name: '売上高経常利益率', category: '②売上高利益', formula: '経常利益÷売上高×100', unit: '%', higherIsBetter: true },
  { code: 'netProfitRatio', name: '売上高当期純利益率', category: '②売上高利益', formula: '当期純利益÷売上高×100', unit: '%', higherIsBetter: true },
  { code: 'sgaRatio', name: '売上高販管費率', category: '②売上高利益', formula: '販管費÷売上高×100', unit: '%', higherIsBetter: false },
  { code: 'laborCostRatio', name: '売上高対人件費比率', category: '②売上高利益', formula: '人件費÷売上高×100', unit: '%', higherIsBetter: false },
  // ③回転率・回転期間
  { code: 'totalCapitalTurnover', name: '総資本回転率', category: '③回転率', formula: '売上高÷総資本', unit: '回', higherIsBetter: true },
  { code: 'receivableTurnoverDays', name: '売上債権回転日数', category: '③回転率', formula: '(受取手形+売掛金)÷売上高×365', unit: '日', higherIsBetter: false },
  { code: 'inventoryTurnoverDays', name: '棚卸資産回転日数', category: '③回転率', formula: '棚卸資産÷売上高×365', unit: '日', higherIsBetter: false },
  { code: 'fixedAssetsTurnover', name: '有形固定資産回転率', category: '③回転率', formula: '売上高÷有形固定資産', unit: '回', higherIsBetter: true },
  { code: 'payableTurnoverDays', name: '買入債務回転日数', category: '③回転率', formula: '(支払手形+買掛金)÷売上高×365', unit: '日', higherIsBetter: false },
  // ④短期支払能力
  { code: 'workingCapitalRatio', name: '運転資金月商倍率', category: '④短期支払能力', formula: '所要運転資金÷(売上高÷12)', unit: '月', higherIsBetter: false },
  { code: 'currentRatio', name: '流動比率', category: '④短期支払能力', formula: '流動資産÷流動負債×100', unit: '%', higherIsBetter: true },
  { code: 'quickRatio', name: '当座比率', category: '④短期支払能力', formula: '当座資産÷流動負債×100', unit: '%', higherIsBetter: true },
  { code: 'cashRatio', name: '支払準備率', category: '④短期支払能力', formula: '現金預金÷流動負債×100', unit: '%', higherIsBetter: true },
  { code: 'cashToSalesRatio', name: '現預金比率', category: '④短期支払能力', formula: '現金預金÷売上高×100', unit: '%', higherIsBetter: true },
  // ⑤資本の安定性
  { code: 'equityRatio', name: '自己資本比率', category: '⑤資本の安定性', formula: '自己資本÷総資本×100', unit: '%', higherIsBetter: true },
  { code: 'equityMultiplier', name: '純資産倍率', category: '⑤資本の安定性', formula: '自己資本÷資本金', unit: '倍', higherIsBetter: true },
  // ⑥調達と運用の適合性
  { code: 'fixedRatio', name: '固定比率', category: '⑥調達と運用の適合性', formula: '固定資産÷自己資本×100', unit: '%', higherIsBetter: false },
  { code: 'fixedLongTermRatio', name: '固定長期適合率', category: '⑥調達と運用の適合性', formula: '固定資産÷(自己資本+固定負債)×100', unit: '%', higherIsBetter: false },
  // ⑦借入状況
  { code: 'debtDependency', name: '借入金依存度', category: '⑦借入状況', formula: '有利子負債÷総資本×100', unit: '%', higherIsBetter: false },
  { code: 'debtCapacityRatio', name: 'デットキャパシティレシオ', category: '⑦借入状況', formula: '有利子負債÷(現金預金+有形固定資産)×100', unit: '%', higherIsBetter: false },
  { code: 'depositLoanRatio', name: '預借率', category: '⑦借入状況', formula: '現金預金÷有利子負債×100', unit: '%', higherIsBetter: true },
  { code: 'debtToMonthlySales', name: '借入金月商倍率', category: '⑦借入状況', formula: '有利子負債÷(売上高÷12)', unit: '倍', higherIsBetter: false },
  { code: 'salesInterestRatio', name: '売上高支払利息・割引料率', category: '⑦借入状況', formula: '支払利息÷売上高×100', unit: '%', higherIsBetter: false },
  // ⑧債務償還能力
  { code: 'cfDebtRatio', name: 'キャッシュフロー有利子負債比率', category: '⑧債務償還能力', formula: '(当期純利益+減価償却)÷有利子負債×100', unit: '%', higherIsBetter: true },
  { code: 'interestCoverageRatio', name: 'インタレストカバレッジレシオ', category: '⑧債務償還能力', formula: '(営業利益+受取利息)÷支払利息', unit: '倍', higherIsBetter: true },
  { code: 'realDebtRepaymentYears', name: '実質債務償還年数', category: '⑧債務償還能力', formula: '(有利子負債-現金預金)÷(営業利益+減価償却)', unit: '年', higherIsBetter: false },
  // ⑨資産の健全性
  { code: 'depreciationRatio', name: '減価償却率', category: '⑨資産の健全性', formula: '減価償却÷(有形固定資産-土地+減価償却)×100', unit: '%', higherIsBetter: true },
  { code: 'salesDepreciationRatio', name: '売上高減価償却費率', category: '⑨資産の健全性', formula: '減価償却÷売上高×100', unit: '%', higherIsBetter: true },
  { code: 'otherCurrentAssetsRatio', name: 'その他流動資産比率', category: '⑨資産の健全性', formula: 'その他流動資産÷流動資産×100', unit: '%', higherIsBetter: false },
  // ⑩成長性
  { code: 'salesGrowthRate', name: '前年比増収率', category: '⑩成長性', formula: '(当期売上÷前期売上-1)×100', unit: '%', higherIsBetter: true },
  { code: 'totalCapitalGrowthRate', name: '総資本増減率', category: '⑩成長性', formula: '(当期総資本÷前期総資本-1)×100', unit: '%', higherIsBetter: true },
];

// ======================================================
// 業種マスタ（中小企業実態基本調査 産業中分類準拠）
// ======================================================
const INDUSTRY_MASTER = [
  // 建設業
  { code: '06', name: '総合工事業', parent: '建設業' },
  { code: '07', name: '職別工事業', parent: '建設業' },
  { code: '08', name: '設備工事業', parent: '建設業' },
  // 製造業
  { code: '09', name: '食料品製造業', parent: '製造業' },
  { code: '10', name: '飲料・たばこ・飼料製造業', parent: '製造業' },
  { code: '11', name: '繊維工業', parent: '製造業' },
  { code: '12', name: '木材・木製品製造業', parent: '製造業' },
  { code: '13', name: '家具・装備品製造業', parent: '製造業' },
  { code: '14', name: 'パルプ・紙・紙加工品製造業', parent: '製造業' },
  { code: '15', name: '印刷・同関連業', parent: '製造業' },
  { code: '16', name: '化学工業', parent: '製造業' },
  { code: '17', name: '石油製品・石炭製品製造業', parent: '製造業' },
  { code: '18', name: 'プラスチック製品製造業', parent: '製造業' },
  { code: '19', name: 'ゴム製品製造業', parent: '製造業' },
  { code: '20', name: 'なめし革・同製品・毛皮製造業', parent: '製造業' },
  { code: '21', name: '窯業・土石製品製造業', parent: '製造業' },
  { code: '22', name: '鉄鋼業', parent: '製造業' },
  { code: '23', name: '非鉄金属製造業', parent: '製造業' },
  { code: '24', name: '金属製品製造業', parent: '製造業' },
  { code: '25', name: 'はん用機械器具製造業', parent: '製造業' },
  { code: '26', name: '生産用機械器具製造業', parent: '製造業' },
  { code: '27', name: '業務用機械器具製造業', parent: '製造業' },
  { code: '28', name: '電子部品・デバイス・電子回路製造業', parent: '製造業' },
  { code: '29', name: '電気機械器具製造業', parent: '製造業' },
  { code: '30', name: '情報通信機械器具製造業', parent: '製造業' },
  { code: '31', name: '輸送用機械器具製造業', parent: '製造業' },
  { code: '32', name: 'その他の製造業', parent: '製造業' },
  // 情報通信業
  { code: '37', name: '通信業', parent: '情報通信業' },
  { code: '38', name: '放送業', parent: '情報通信業' },
  { code: '39', name: '情報サービス業', parent: '情報通信業' },
  { code: '40', name: 'インターネット附随サービス業', parent: '情報通信業' },
  { code: '41', name: '映像・音声・文字情報制作業', parent: '情報通信業' },
  // 運輸業，郵便業
  { code: '43', name: '道路旅客運送業', parent: '運輸業' },
  { code: '44', name: '道路貨物運送業', parent: '運輸業' },
  { code: '45', name: '水運業', parent: '運輸業' },
  { code: '47', name: '倉庫業', parent: '運輸業' },
  { code: '48', name: '運輸に附帯するサービス業', parent: '運輸業' },
  { code: '49', name: '郵便業', parent: '運輸業' },
  // 卸売業
  { code: '50', name: '各種商品卸売業', parent: '卸売業' },
  { code: '51', name: '繊維・衣服等卸売業', parent: '卸売業' },
  { code: '52', name: '飲食料品卸売業', parent: '卸売業' },
  { code: '53', name: '建築材料，鉱物・金属材料等卸売業', parent: '卸売業' },
  { code: '54', name: '機械器具卸売業', parent: '卸売業' },
  { code: '55', name: 'その他の卸売業', parent: '卸売業' },
  // 小売業
  { code: '56', name: '各種商品小売業', parent: '小売業' },
  { code: '57', name: '織物・衣服・身の回り品小売業', parent: '小売業' },
  { code: '58', name: '飲食料品小売業', parent: '小売業' },
  { code: '59', name: '機械器具小売業', parent: '小売業' },
  { code: '60', name: 'その他の小売業', parent: '小売業' },
  { code: '61', name: '無店舗小売業', parent: '小売業' },
  // 不動産業
  { code: '68', name: '不動産取引業', parent: '不動産業' },
  { code: '69', name: '不動産賃貸業・管理業', parent: '不動産業' },
  { code: '70', name: '物品賃貸業', parent: '不動産業' },
  // 学術研究・専門技術
  { code: '72', name: '専門サービス業', parent: '学術研究・専門技術' },
  { code: '73', name: '広告業', parent: '学術研究・専門技術' },
  { code: '74', name: '技術サービス業', parent: '学術研究・専門技術' },
  // 宿泊・飲食
  { code: '75', name: '宿泊業', parent: '宿泊・飲食' },
  { code: '76', name: '飲食店', parent: '宿泊・飲食' },
  { code: '77', name: '持ち帰り・配達飲食サービス業', parent: '宿泊・飲食' },
  // 生活関連サービス・娯楽
  { code: '78', name: '洗濯・理容・美容・浴場業', parent: '生活関連・娯楽' },
  { code: '79', name: 'その他の生活関連サービス業', parent: '生活関連・娯楽' },
  { code: '80', name: '娯楽業', parent: '生活関連・娯楽' },
  // サービス業
  { code: '88', name: '廃棄物処理業', parent: 'サービス業' },
  { code: '89', name: '自動車整備業', parent: 'サービス業' },
  { code: '90', name: '機械等修理業', parent: 'サービス業' },
  { code: '91', name: '職業紹介・労働者派遣業', parent: 'サービス業' },
  { code: '92', name: 'その他の事業サービス業', parent: 'サービス業' },
];

// 規模マスタ
const SCALE_MASTER = [
  { code: 'small', name: '5人以下', equityMul: 0.85 },
  { code: 'medium', name: '6〜20人', equityMul: 1.00 },
  { code: 'large', name: '21〜50人', equityMul: 1.10 },
  { code: 'xlarge', name: '51人以上', equityMul: 1.20 },
];

// ======================================================
// デフォルト業界平均データ（公開資料からの代表値）
// 構造: { [industryCode]: { [indicatorCode]: { mean, top25, sd } } }
// ※運営者がCSV/Excelで上書き可能
// ======================================================
const DEFAULT_BENCHMARK_DATA = {
  // 製造業の代表値（食料品製造業を基準にプリセット）
  '09': {
    equityRatio: { mean: 38.5, top25: 60.0, sd: 22.0 },
    currentRatio: { mean: 195.0, top25: 295.0, sd: 105.0 },
    fixedLongTermRatio: { mean: 85.0, top25: 60.0, sd: 38.0 },
    salesOrdinaryProfitRatio: { mean: 3.5, top25: 7.5, sd: 5.5 },
    ordinaryProfitRatio: { mean: 3.5, top25: 7.5, sd: 5.5 },
    operatingProfitRatio: { mean: 2.8, top25: 6.0, sd: 5.0 },
    totalCapitalOrdinaryProfitRatio: { mean: 4.0, top25: 8.5, sd: 5.0 },
    debtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
    realDebtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
    interestCoverageRatio: { mean: 8.0, top25: 18.0, sd: 11.0 },
    debtDependency: { mean: 38.0, top25: 18.0, sd: 22.0 },
    salesInterestRatio: { mean: 0.8, top25: 0.3, sd: 0.8 },
  },
  // 他業種は同様の代表値を持つ（運営者がCSV更新する想定）
};

// 全業種に代表値を展開（製造業ベース）
INDUSTRY_MASTER.forEach(ind => {
  if (!DEFAULT_BENCHMARK_DATA[ind.code]) {
    const base = DEFAULT_BENCHMARK_DATA['09'];
    DEFAULT_BENCHMARK_DATA[ind.code] = { ...base };
  }
});

// 旧定数の互換性維持（既存ロジックが参照しているため残す）
const INDUSTRY_BENCHMARKS = {
  manufacturing: { name: '製造業', ...DEFAULT_BENCHMARK_DATA['09'] },
  wholesale: { name: '卸売業', ...DEFAULT_BENCHMARK_DATA['55'] },
  retail: { name: '小売業', ...DEFAULT_BENCHMARK_DATA['58'] },
  service: { name: 'サービス業', ...DEFAULT_BENCHMARK_DATA['92'] },
  construction: { name: '建設業', ...DEFAULT_BENCHMARK_DATA['06'] },
  it: { name: '情報通信業', ...DEFAULT_BENCHMARK_DATA['39'] },
  transport: { name: '運輸業', ...DEFAULT_BENCHMARK_DATA['44'] },
  hospitality: { name: '宿泊業・飲食サービス業', ...DEFAULT_BENCHMARK_DATA['76'] },
};

const SCALE_ADJUSTMENT = {
  small: { equityMul: 0.85, label: '5人以下（小規模）' },
  medium: { equityMul: 1.00, label: '6-20人' },
  large: { equityMul: 1.10, label: '21-50人' },
  xlarge: { equityMul: 1.20, label: '51人以上' },
};

// 根拠記事DB（安田経営診断事務所｜安田順氏の解説記事）
const REFERENCE_ARTICLES = {
  // Sランク：スコア計算・主要評価の根拠
  mcss_contribution: {
    rank: 'S', title: 'MCSSの寄与率 377件を検証してみてわかったこと',
    url: 'https://yasuda-keiei.com/point/post-4120/',
    summary: '寄与率をそのまま弱点としない、5点ルールで主因を絞る、短期支払能力は別扱いという処理が重要。',
    category: 'スコア計算'
  },
  black_red_rating: {
    rank: 'S', title: '黒字なのに格付けが下がる会社、赤字でも上がる会社',
    url: 'https://yasuda-keiei.com/point/post-4001/',
    summary: '損益だけでは格付けは決まらない。借入金依存度や債務償還年数など財務体質が決定的。',
    category: 'スコア計算'
  },
  mcss_vs_jiko: {
    rank: 'S', title: '経営自己診断システムとMcSSの分析結果にはどういう違いがあるのか？',
    url: 'https://yasuda-keiei.com/point/kaizen/post-1955/',
    summary: '中小機構の経営自己診断システムは評価が甘め。McSSのCRDランクの方が銀行評価に近い。',
    category: 'スコア計算'
  },
  // Aランク：財務指標別コメントの根拠
  debt_repayment: {
    rank: 'A', title: '債務償還年数とは何か？──銀行が最も重視する財務指標',
    url: 'https://yasuda-keiei.com/point/post-4006/',
    summary: '債務償還年数は銀行が最重視する指標。10年以内が目安、15年超は危険水準。',
    category: '財務指標'
  },
  interest_ratio: {
    rank: 'A', title: '支払利息は売上の何％であるべきか？',
    url: 'https://yasuda-keiei.com/point/kaizen/post-1224/',
    summary: '売上高支払利息率1%（小売・卸売0.7%）超で要警戒、1.5%以上は過剰債務、2.8%でデフォルト企業以下。',
    category: '財務指標'
  },
  cashflow: {
    rank: 'A', title: 'キャッシュフロー計算書をみないとどうなるか',
    url: 'https://yasuda-keiei.com/point/post-4065/',
    summary: '営業CFを見ないと本業の現金創出力が分からない。PL黒字でも営業CFがマイナスなら危険。',
    category: '財務指標'
  },
  current_ratio: {
    rank: 'A', title: '流動比率はアテにならない',
    url: 'https://yasuda-keiei.com/point/post-3969/',
    summary: '流動比率は不良在庫・回収困難売掛金の影響を受けやすく、短期支払能力の正確な指標ではない。',
    category: '財務指標'
  },
  excess_debt: {
    rank: 'A', title: '過剰債務から抜け出すには営業利益率5％以上が必要',
    url: 'https://yasuda-keiei.com/point/post-2631/',
    summary: '過剰債務脱出には営業利益率5%以上が必要。中小企業の平均（2-3%）では返済原資が不足。',
    category: '財務指標'
  },
  // Bランク：危険信号・改善コメントの根拠
  tax_saving: {
    rank: 'B', title: '銀行が認める節税、認めない節税',
    url: 'https://yasuda-keiei.com/point/post-4018/',
    summary: '過度な節税は自己資本を毀損し格付けを下げる。銀行に評価される節税と嫌われる節税の違い。',
    category: '改善アドバイス'
  },
  trial_balance: {
    rank: 'B', title: '銀行は試算表をどう見ているか',
    url: 'https://yasuda-keiei.com/point/post-3991/',
    summary: '試算表は経営の月次健康診断書。銀行が試算表で何をチェックしているかを理解する。',
    category: '改善アドバイス'
  },
  explanation_doc: {
    rank: 'B', title: '銀行に提出する決算説明資料の書き方',
    url: 'https://yasuda-keiei.com/point/ginkou/post-2490/',
    summary: '決算説明資料で銀行評価を改善する書き方。数字の背景と将来見通しを伝える。',
    category: '改善アドバイス'
  },
  basics: {
    rank: 'B', title: '金融機関が求めているのは基本の徹底',
    url: 'https://yasuda-keiei.com/point/post-4062/',
    summary: '小手先の対策より、毎期黒字を出し、月次試算表を期日通り提出する基本動作の徹底が重要。',
    category: '改善アドバイス'
  },
  accounting_dependency: {
    rank: 'B', title: '経理に依存し過ぎるとどうなるか',
    url: 'https://yasuda-keiei.com/point/post-4022/',
    summary: '経理任せにすると経営判断が遅れる。社長自身が試算表を読める力が必要。',
    category: '改善アドバイス'
  },
  // 危険信号系
  loan_stop: {
    rank: 'B', title: '融資が止まっても事業を継続できるか',
    url: 'https://yasuda-keiei.com/point/post-4056/',
    summary: '融資依存度を下げる経営、自力でキャッシュを回す体制づくり。',
    category: '危険信号'
  },
  bank_overloan: {
    rank: 'B', title: '銀行に貸し込まれるとはどういう状態か？',
    url: 'https://yasuda-keiei.com/point/post-3861/',
    summary: '銀行は貸せるうちに貸そうとする。貸し込まれた状態の見分け方と対処法。',
    category: '危険信号'
  },
  factoring: {
    rank: 'B', title: 'ファクタリングの利用は1回限り',
    url: 'https://yasuda-keiei.com/point/kaizen/post-3915/',
    summary: 'ファクタリング依存は資金繰りを更に悪化させる。1回限りの緊急避難として使う。',
    category: '危険信号'
  },
  bad_news: {
    rank: 'B', title: '銀行に悪い話をどう伝えるか',
    url: 'https://yasuda-keiei.com/point/ginkou/post-694/',
    summary: '業績悪化等の悪い話は隠さず早めに伝える。伝え方とタイミングが信頼を左右する。',
    category: '危険信号'
  },
};

function App() {
  return <BankRatingAdvisor />;
}

export default App;

function BankRatingAdvisor() {
  const [step, setStep] = useState('input');
  const [industry, setIndustry] = useState('manufacturing');
  const [industryCode, setIndustryCode] = useState('09'); // 中分類コード
  const [scale, setScale] = useState('medium');
  const [simMode, setSimMode] = useState(false);
  const [expandedMetric, setExpandedMetric] = useState(null);
  const [inputMode, setInputMode] = useState('detailed');
  const [useActualBS, setUseActualBS] = useState(true);
  const [expandedSection, setExpandedSection] = useState({ assets: true, liabilities: true, adjust: true, dataManage: false });
  
  // ========== データ管理状態 ==========
  const [benchmarkData, setBenchmarkData] = useState(DEFAULT_BENCHMARK_DATA);
  const [dataVersion, setDataVersion] = useState({
    version: '令和6年確報（令和5年度決算実績）',
    importedAt: '2026/01/01（初期値）',
    source: '中小企業実態基本調査（経済産業省・中小企業庁）',
    industryCount: INDUSTRY_MASTER.length,
    indicatorCount: INDICATORS_38.length,
  });
  const [importPreview, setImportPreview] = useState(null);
  const [importError, setImportError] = useState(null);
  const [importSuccess, setImportSuccess] = useState(null);
  const fileInputRef = useRef(null);

  // 詳細な勘定科目データ
  const [accounts, setAccounts] = useState({
    // 流動資産
    cash: 30000,                  // 現金預金
    receivablesNotes: 5000,       // 受取手形
    receivablesAccount: 35000,    // 売掛金
    inventory: 25000,             // 棚卸資産
    advancesToEmployees: 0,       // 仮払金・立替金
    shortTermLoansToExec: 0,      // 短期役員貸付金
    otherCurrentAssets: 5000,     // その他流動資産
    // 固定資産
    land: 40000,                  // 土地
    building: 35000,              // 建物
    machinery: 15000,             // 機械装置
    longTermLoansToExec: 0,       // 長期役員貸付金
    investmentSecurities: 5000,   // 投資有価証券
    securityDeposit: 8000,        // 入居保証金・敷金
    otherFixedAssets: 5000,       // その他固定資産
    deferredAssets: 0,            // 繰延資産
    // 流動負債
    payablesNotes: 5000,          // 支払手形
    payablesAccount: 20000,       // 買掛金
    shortTermBorrowings: 25000,   // 短期借入金
    otherCurrentLiabilities: 10000,
    // 固定負債
    longTermBorrowings: 70000,    // 長期借入金
    longTermBorrowingsFromExec: 0,// 役員借入金
    otherFixedLiabilities: 10000,
    // 純資産
    capitalStock: 30000,          // 資本金
    capitalSurplus: 5000,         // 資本剰余金
    retainedEarnings: 25000,      // 利益剰余金
    // BS実態修正項目
    badInventory: 5000,           // 不良在庫（推定）
    badReceivables: 2000,         // 回収困難売掛金（推定）
    badAdvances: 0,               // 不良仮払金
    depreciationShortage: 0,      // 減価償却不足
    landMarketValueAdjust: 0,     // 土地時価調整（マイナスは含み損）
    securitiesMarketValueAdjust: 0, // 有価証券時価調整
    // PL
    sales: 200000,
    operatingProfit: 8000,
    ordinaryProfit: 6000,
    prevOrdinaryProfit: 5500,
    depreciation: 4000,
    interestPayment: 1500,
    consecutiveBlackYears: 2,
  });

  const [simAdjust, setSimAdjust] = useState({
    netAssets: 0,
    interestBearingDebt: 0,
    ordinaryProfit: 0,
    currentAssets: 0,
    currentLiabilities: 0,
  });

  // 表面BS（書類通り）
  const surfaceBS = useMemo(() => {
    const currentAssets = accounts.cash + accounts.receivablesNotes + accounts.receivablesAccount + 
                          accounts.inventory + accounts.advancesToEmployees + 
                          accounts.shortTermLoansToExec + accounts.otherCurrentAssets;
    const fixedAssets = accounts.land + accounts.building + accounts.machinery + 
                        accounts.longTermLoansToExec + accounts.investmentSecurities + 
                        accounts.securityDeposit + accounts.otherFixedAssets + accounts.deferredAssets;
    const currentLiabilities = accounts.payablesNotes + accounts.payablesAccount + 
                                accounts.shortTermBorrowings + accounts.otherCurrentLiabilities;
    const fixedLiabilities = accounts.longTermBorrowings + accounts.longTermBorrowingsFromExec + 
                              accounts.otherFixedLiabilities;
    const netAssets = accounts.capitalStock + accounts.capitalSurplus + accounts.retainedEarnings;
    const interestBearingDebt = accounts.shortTermBorrowings + accounts.longTermBorrowings + 
                                 accounts.longTermBorrowingsFromExec;
    return { currentAssets, fixedAssets, currentLiabilities, fixedLiabilities, netAssets, interestBearingDebt };
  }, [accounts]);

  // 実態BS変換（銀行員視点）
  const actualBSAdjustments = useMemo(() => {
    // マイナス調整（不良資産）
    const inventoryDeduction = -accounts.badInventory;          // 不良在庫
    const receivablesDeduction = -accounts.badReceivables;      // 不良売掛金
    const advancesDeduction = -accounts.badAdvances;            // 不良仮払金
    const execLoansDeduction = -(accounts.shortTermLoansToExec + accounts.longTermLoansToExec); // 役員貸付は不良資産扱い
    const deferredDeduction = -accounts.deferredAssets;         // 繰延資産はゼロ評価
    const securityDepositDeduction = -accounts.securityDeposit * 0.7; // 入居保証金30%評価
    const depreciationDeduction = -accounts.depreciationShortage; // 減価償却不足
    const landAdjust = accounts.landMarketValueAdjust;          // 土地時価調整
    const securitiesAdjust = accounts.securitiesMarketValueAdjust; // 有価証券時価調整

    // プラス調整（役員借入金は実質自己資本）
    const execBorrowingsAddition = accounts.longTermBorrowingsFromExec;

    return {
      inventoryDeduction,
      receivablesDeduction,
      advancesDeduction,
      execLoansDeduction,
      deferredDeduction,
      securityDepositDeduction,
      depreciationDeduction,
      landAdjust,
      securitiesAdjust,
      execBorrowingsAddition,
      totalAssetAdjust: inventoryDeduction + receivablesDeduction + advancesDeduction + 
                         execLoansDeduction + deferredDeduction + securityDepositDeduction + 
                         depreciationDeduction + landAdjust + securitiesAdjust,
      totalEquityAdjust: inventoryDeduction + receivablesDeduction + advancesDeduction + 
                          execLoansDeduction + deferredDeduction + securityDepositDeduction + 
                          depreciationDeduction + landAdjust + securitiesAdjust + execBorrowingsAddition,
    };
  }, [accounts]);

  // 採用するBS（表面 or 実態）
  const baseBS = useMemo(() => {
    if (!useActualBS) return surfaceBS;
    const adj = actualBSAdjustments;
    // 資産側の調整（不良在庫、不良売掛金、繰延、入居保証金は流動・固定別に按分は簡略化のため流動側）
    const currentAssetAdjust = adj.inventoryDeduction + adj.receivablesDeduction + adj.advancesDeduction;
    const fixedAssetAdjust = adj.execLoansDeduction + adj.deferredDeduction + 
                              adj.securityDepositDeduction + adj.depreciationDeduction + 
                              adj.landAdjust + adj.securitiesAdjust;
    // 役員借入金を負債から純資産へ振替
    return {
      currentAssets: surfaceBS.currentAssets + currentAssetAdjust,
      fixedAssets: surfaceBS.fixedAssets + fixedAssetAdjust,
      currentLiabilities: surfaceBS.currentLiabilities,
      fixedLiabilities: surfaceBS.fixedLiabilities - adj.execBorrowingsAddition,
      netAssets: surfaceBS.netAssets + adj.totalEquityAdjust,
      // 有利子負債は役員借入金を除く
      interestBearingDebt: surfaceBS.interestBearingDebt - adj.execBorrowingsAddition,
    };
  }, [surfaceBS, actualBSAdjustments, useActualBS]);

  // シミュレーション反映後
  const adjustedData = useMemo(() => ({
    currentAssets: baseBS.currentAssets + simAdjust.currentAssets,
    fixedAssets: baseBS.fixedAssets,
    currentLiabilities: Math.max(0, baseBS.currentLiabilities + simAdjust.currentLiabilities),
    fixedLiabilities: baseBS.fixedLiabilities,
    netAssets: baseBS.netAssets + simAdjust.netAssets,
    interestBearingDebt: Math.max(0, baseBS.interestBearingDebt + simAdjust.interestBearingDebt),
    sales: accounts.sales,
    operatingProfit: accounts.operatingProfit,
    ordinaryProfit: accounts.ordinaryProfit + simAdjust.ordinaryProfit,
    prevOrdinaryProfit: accounts.prevOrdinaryProfit,
    depreciation: accounts.depreciation,
    interestPayment: accounts.interestPayment,
    consecutiveBlackYears: accounts.consecutiveBlackYears,
  }), [baseBS, accounts, simAdjust]);

  const metrics = useMemo(() => {
    const d = adjustedData;
    const totalAssets = d.currentAssets + d.fixedAssets;
    const totalLiabilities = d.currentLiabilities + d.fixedLiabilities;
    const cashFlow = d.operatingProfit + d.depreciation;
    return {
      equityRatio: totalAssets > 0 ? (d.netAssets / totalAssets) * 100 : 0,
      gearingRatio: d.netAssets > 0 ? (totalLiabilities / d.netAssets) * 100 : 999,
      fixedLongTermRatio: (d.fixedLiabilities + d.netAssets) > 0 ? (d.fixedAssets / (d.fixedLiabilities + d.netAssets)) * 100 : 999,
      currentRatio: d.currentLiabilities > 0 ? (d.currentAssets / d.currentLiabilities) * 100 : 999,
      salesOrdinaryProfitRatio: d.sales > 0 ? (d.ordinaryProfit / d.sales) * 100 : 0,
      totalCapitalOrdinaryProfitRatio: totalAssets > 0 ? (d.ordinaryProfit / totalAssets) * 100 : 0,
      consecutiveBlackYears: d.consecutiveBlackYears,
      ordinaryProfitGrowthRate: d.prevOrdinaryProfit > 0 ? ((d.ordinaryProfit - d.prevOrdinaryProfit) / d.prevOrdinaryProfit) * 100 : 0,
      equityAmount: d.netAssets,
      salesAmount: d.sales,
      debtRepaymentYears: cashFlow > 0 ? d.interestBearingDebt / cashFlow : 999,
      interestCoverageRatio: d.interestPayment > 0 ? d.operatingProfit / d.interestPayment : 999,
      cashFlowAmount: cashFlow,
    };
  }, [adjustedData]);

  const scoreMetric = (value, benchmarks) => {
    // benchmarksは降順（厳しい条件→緩い条件）で並んでいる前提
    // 順方向に走査して、最初にヒットした条件のスコアを返す
    for (let i = 0; i < benchmarks.length; i++) {
      if (benchmarks[i].compare(value)) return benchmarks[i].score;
    }
    return 0;
  };

  const scores = useMemo(() => {
    const m = metrics;
    return {
      equityRatio: scoreMetric(m.equityRatio, [
        { compare: v => v >= 50, score: 10 },
        { compare: v => v >= 40, score: 8 },
        { compare: v => v >= 25, score: 5 },
        { compare: v => v >= 15, score: 3 },
        { compare: v => v >= 5, score: 1 },
        { compare: v => v >= 0, score: 0 },
      ]),
      gearingRatio: scoreMetric(m.gearingRatio, [
        { compare: v => v <= 50, score: 10 },
        { compare: v => v <= 100, score: 8 },
        { compare: v => v <= 150, score: 6 },
        { compare: v => v <= 250, score: 4 },
        { compare: v => v <= 400, score: 2 },
        { compare: v => v <= 999, score: 0 },
      ]),
      fixedLongTermRatio: scoreMetric(m.fixedLongTermRatio, [
        { compare: v => v <= 40, score: 10 },
        { compare: v => v <= 60, score: 7 },
        { compare: v => v <= 80, score: 5 },
        { compare: v => v <= 100, score: 3 },
        { compare: v => v <= 999, score: 0 },
      ]),
      currentRatio: scoreMetric(m.currentRatio, [
        { compare: v => v >= 200, score: 10 },
        { compare: v => v >= 140, score: 7 },
        { compare: v => v >= 120, score: 5 },
        { compare: v => v >= 100, score: 3 },
        { compare: v => v >= 0, score: 0 },
      ]),
      salesOrdinaryProfitRatio: scoreMetric(m.salesOrdinaryProfitRatio, [
        { compare: v => v >= 8, score: 10 },
        { compare: v => v >= 5, score: 8 },
        { compare: v => v >= 2, score: 6 },
        { compare: v => v >= 1, score: 4 },
        { compare: v => v >= 0, score: 2 },
        { compare: v => v >= -999, score: 0 },
      ]),
      totalCapitalOrdinaryProfitRatio: scoreMetric(m.totalCapitalOrdinaryProfitRatio, [
        { compare: v => v >= 5, score: 10 },
        { compare: v => v >= 3, score: 8 },
        { compare: v => v >= 1, score: 6 },
        { compare: v => v >= 0, score: 3 },
        { compare: v => v >= -999, score: 0 },
      ]),
      consecutiveBlackYears: scoreMetric(m.consecutiveBlackYears, [
        { compare: v => v >= 5, score: 10 },
        { compare: v => v >= 3, score: 8 },
        { compare: v => v >= 2, score: 6 },
        { compare: v => v >= 1, score: 3 },
        { compare: v => v >= 0, score: 0 },
      ]),
      ordinaryProfitGrowthRate: scoreMetric(m.ordinaryProfitGrowthRate, [
        { compare: v => v >= 30, score: 10 },
        { compare: v => v >= 15, score: 6 },
        { compare: v => v >= 5, score: 4 },
        { compare: v => v >= 0, score: 2 },
        { compare: v => v >= -999, score: 0 },
      ]),
      equityAmount: scoreMetric(m.equityAmount, [
        { compare: v => v >= 1000000, score: 10 },
        { compare: v => v >= 300000, score: 6 },
        { compare: v => v >= 100000, score: 2 },
        { compare: v => v >= 30000, score: 1 },
        { compare: v => v >= 0, score: 0 },
      ]),
      salesAmount: scoreMetric(m.salesAmount, [
        { compare: v => v >= 1000000, score: 10 },
        { compare: v => v >= 300000, score: 6 },
        { compare: v => v >= 100000, score: 2 },
        { compare: v => v >= 30000, score: 1 },
        { compare: v => v >= 0, score: 0 },
      ]),
      debtRepaymentYears: scoreMetric(m.debtRepaymentYears, [
        { compare: v => v <= 3, score: 10 },
        { compare: v => v <= 5, score: 8 },
        { compare: v => v <= 7, score: 5 },
        { compare: v => v <= 10, score: 3 },
        { compare: v => v <= 15, score: 1 },
        { compare: v => v <= 999, score: 0 },
      ]),
      interestCoverageRatio: scoreMetric(m.interestCoverageRatio, [
        { compare: v => v >= 10, score: 10 },
        { compare: v => v >= 5, score: 7 },
        { compare: v => v >= 2.5, score: 3 },
        { compare: v => v >= 1, score: 1 },
        { compare: v => v >= 0, score: 0 },
      ]),
      cashFlowAmount: scoreMetric(m.cashFlowAmount, [
        { compare: v => v >= 1000000, score: 10 },
        { compare: v => v >= 300000, score: 6 },
        { compare: v => v >= 100000, score: 2 },
        { compare: v => v >= 30000, score: 1 },
        { compare: v => v >= 0, score: 0 },
      ]),
    };
  }, [metrics]);

  const categoryScores = useMemo(() => {
    const safety = scores.equityRatio + scores.gearingRatio + scores.fixedLongTermRatio + scores.currentRatio;
    const profitability = scores.salesOrdinaryProfitRatio + scores.totalCapitalOrdinaryProfitRatio + scores.consecutiveBlackYears;
    const growth = scores.ordinaryProfitGrowthRate + scores.equityAmount + scores.salesAmount;
    const repayment = scores.debtRepaymentYears + scores.interestCoverageRatio + scores.cashFlowAmount;
    const total = safety + profitability + growth + repayment;
    const total100 = Math.round((total / 130) * 100);
    return { safety, profitability, growth, repayment, total, total100 };
  }, [scores]);

  // 実質債務超過判定
  const isInsolvent = adjustedData.netAssets < 0;

  const rating = useMemo(() => {
    if (isInsolvent) return { level: 8, label: '警戒先（実質債務超過）', color: '#dc2626', desc: '実態純資産がマイナス。新規融資はほぼ不可。', zone: 'danger' };
    const s = categoryScores.total100;
    if (s >= 90) return { level: 1, label: 'リスク無し', color: '#1e40af', desc: '安全性は最高基準。融資先としてのリスクは最小限。', zone: 'top' };
    if (s >= 80) return { level: 2, label: 'ほとんどリスク無し', color: '#2563eb', desc: 'かなり優れていると判断される。返済の確実性は極めて高い。', zone: 'top' };
    if (s >= 65) return { level: 3, label: 'リスク些少', color: '#0284c7', desc: '貸出先としてかなり魅力的。返済の可能性はかなり高い。', zone: 'top' };
    if (s >= 50) return { level: 4, label: 'リスクあるが良好', color: '#0891b2', desc: '元利払いの確実性は高い。優良企業の目標水準。', zone: 'good' };
    if (s >= 40) return { level: 5, label: '平均的水準', color: '#059669', desc: '当面の返済能力には不安なし。標準的な企業。', zone: 'good' };
    if (s >= 25) return { level: 6, label: 'リスクやや高い', color: '#ca8a04', desc: '当面の返済能力はあるが、長期的安全性は低い。融資可能ライン。', zone: 'caution' };
    if (s >= 15) return { level: 7, label: 'リスク高く徹底管理', color: '#ea580c', desc: '返済の確実性が低く、安全性に欠ける。要注意先。', zone: 'warning' };
    if (s >= 10) return { level: 8, label: '警戒先', color: '#dc2626', desc: '企業内容はかなり悪化。経営が行き詰る可能性。', zone: 'danger' };
    return { level: 9, label: '延滞先', color: '#991b1b', desc: '貸出金の回収に重大な懸念。新規融資はほぼ不可。', zone: 'danger' };
  }, [categoryScores, isInsolvent]);

  const deviation = useMemo(() => {
    const dev = 50 + ((categoryScores.total100 - 50) / 15) * 10;
    return Math.max(25, Math.min(75, Math.round(dev)));
  }, [categoryScores]);

  const crdRank = useMemo(() => {
    if (deviation >= 60) return { rank: 'A', desc: '上位20%', color: '#1e40af' };
    if (deviation >= 53) return { rank: 'B', desc: '上位40%', color: '#0891b2' };
    if (deviation >= 47) return { rank: 'C', desc: '中位', color: '#059669' };
    if (deviation >= 40) return { rank: 'D', desc: '下位40%', color: '#ea580c' };
    return { rank: 'E', desc: '下位20%', color: '#dc2626' };
  }, [deviation]);

  const debtorClass = useMemo(() => {
    if (rating.level <= 5) return { name: '正常先', color: '#059669' };
    if (rating.level === 6) return { name: '正常先(下位)', color: '#ca8a04' };
    if (rating.level === 7) return { name: '要注意先', color: '#ea580c' };
    if (rating.level === 8) return { name: '要管理先', color: '#dc2626' };
    return { name: '破綻懸念先', color: '#991b1b' };
  }, [rating]);

  // ========================================
  // 業界比較ロジック（フェーズ2）
  // ========================================
  const benchmark = useMemo(() => {
    // 業種コード（中分類）ベースで取得
    const industryData = benchmarkData[industryCode];
    const industryMeta = INDUSTRY_MASTER.find(i => i.code === industryCode);
    const scaleMeta = SCALE_MASTER.find(s => s.code === scale) || SCALE_MASTER[1];
    
    // デフォルト指標値（必ず存在することを保証）
    const fallback = {
      equityRatio: { mean: 40, top25: 60, sd: 21 },
      currentRatio: { mean: 190, top25: 280, sd: 95 },
      fixedLongTermRatio: { mean: 75, top25: 55, sd: 35 },
      gearingRatio: { mean: 140, top25: 65, sd: 90 },
      salesOrdinaryProfitRatio: { mean: 3.5, top25: 7.0, sd: 5.0 },
      ordinaryProfitRatio: { mean: 3.5, top25: 7.0, sd: 5.0 },
      operatingProfitRatio: { mean: 2.8, top25: 6.0, sd: 5.0 },
      totalCapitalOrdinaryProfitRatio: { mean: 4.0, top25: 8.0, sd: 5.0 },
      debtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
      realDebtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
      interestCoverageRatio: { mean: 8.0, top25: 18.0, sd: 11.0 },
      debtDependency: { mean: 38, top25: 18, sd: 22 },
      salesInterestRatio: { mean: 0.8, top25: 0.3, sd: 0.8 },
    };
    
    const data = industryData || {};
    // fallbackとindustryDataをマージ（industryDataが優先）
    const merged = { ...fallback, ...data };
    
    return {
      name: industryMeta?.name || '不明',
      parent: industryMeta?.parent || '',
      ...merged,
      equityRatio: {
        ...merged.equityRatio,
        mean: merged.equityRatio.mean * scaleMeta.equityMul,
        top25: merged.equityRatio.top25 * scaleMeta.equityMul,
      },
    };
  }, [industry, industryCode, scale, benchmarkData]);

  // 各指標の偏差値計算（業界平均50、標準偏差10）
  // 「低い方が良い指標」（債務償還年数、ギアリング比率、固定長期適合率）は反転
  const calcZ = (value, mean, sd, inverse = false) => {
    if (sd === 0 || isNaN(value) || !isFinite(value)) return 50;
    const z = (value - mean) / sd;
    const dev = inverse ? 50 - z * 10 : 50 + z * 10;
    return Math.max(25, Math.min(75, Math.round(dev)));
  };

  // デフォルト値（指標が業界平均データに存在しない場合のフォールバック）
  const DEFAULT_BM = {
    equityRatio: { mean: 40, top25: 60, sd: 21 },
    currentRatio: { mean: 190, top25: 280, sd: 95 },
    fixedLongTermRatio: { mean: 75, top25: 55, sd: 35 },
    gearingRatio: { mean: 140, top25: 65, sd: 90 },
    salesOrdinaryProfitRatio: { mean: 3.5, top25: 7.0, sd: 5.0 },
    ordinaryProfitRatio: { mean: 3.5, top25: 7.0, sd: 5.0 },
    totalCapitalOrdinaryProfitRatio: { mean: 4.0, top25: 8.0, sd: 5.0 },
    debtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
    realDebtRepaymentYears: { mean: 8.0, top25: 4.0, sd: 7.0 },
    interestCoverageRatio: { mean: 8.0, top25: 18.0, sd: 11.0 },
  };

  const getBM = (key) => {
    return benchmark[key] || DEFAULT_BM[key] || { mean: 50, top25: 60, sd: 10 };
  };

  const industryComparison = useMemo(() => {
    const bm = (key) => benchmark[key] || DEFAULT_BM[key] || { mean: 50, top25: 60, sd: 10 };
    return {
      equityRatio: {
        myValue: metrics.equityRatio,
        mean: bm('equityRatio').mean,
        top25: bm('equityRatio').top25,
        deviation: calcZ(metrics.equityRatio, bm('equityRatio').mean, bm('equityRatio').sd, false),
        higherIsBetter: true,
      },
      currentRatio: {
        myValue: metrics.currentRatio,
        mean: bm('currentRatio').mean,
        top25: bm('currentRatio').top25,
        deviation: calcZ(metrics.currentRatio, bm('currentRatio').mean, bm('currentRatio').sd, false),
        higherIsBetter: true,
      },
      fixedLongTermRatio: {
        myValue: metrics.fixedLongTermRatio,
        mean: bm('fixedLongTermRatio').mean,
        top25: bm('fixedLongTermRatio').top25,
        deviation: calcZ(metrics.fixedLongTermRatio, bm('fixedLongTermRatio').mean, bm('fixedLongTermRatio').sd, true),
        higherIsBetter: false,
      },
      gearingRatio: {
        myValue: metrics.gearingRatio,
        mean: bm('gearingRatio').mean,
        top25: bm('gearingRatio').top25,
        deviation: calcZ(metrics.gearingRatio, bm('gearingRatio').mean, bm('gearingRatio').sd, true),
        higherIsBetter: false,
      },
      salesOrdinaryProfitRatio: {
        myValue: metrics.salesOrdinaryProfitRatio,
        mean: bm('salesOrdinaryProfitRatio').mean,
        top25: bm('salesOrdinaryProfitRatio').top25,
        deviation: calcZ(metrics.salesOrdinaryProfitRatio, bm('salesOrdinaryProfitRatio').mean, bm('salesOrdinaryProfitRatio').sd, false),
        higherIsBetter: true,
      },
      totalCapitalOrdinaryProfitRatio: {
        myValue: metrics.totalCapitalOrdinaryProfitRatio,
        mean: bm('totalCapitalOrdinaryProfitRatio').mean,
        top25: bm('totalCapitalOrdinaryProfitRatio').top25,
        deviation: calcZ(metrics.totalCapitalOrdinaryProfitRatio, bm('totalCapitalOrdinaryProfitRatio').mean, bm('totalCapitalOrdinaryProfitRatio').sd, false),
        higherIsBetter: true,
      },
      debtRepaymentYears: {
        myValue: metrics.debtRepaymentYears,
        mean: bm('debtRepaymentYears').mean,
        top25: bm('debtRepaymentYears').top25,
        deviation: calcZ(metrics.debtRepaymentYears, bm('debtRepaymentYears').mean, bm('debtRepaymentYears').sd, true),
        higherIsBetter: false,
      },
      interestCoverageRatio: {
        myValue: metrics.interestCoverageRatio,
        mean: bm('interestCoverageRatio').mean,
        top25: bm('interestCoverageRatio').top25,
        deviation: calcZ(metrics.interestCoverageRatio, bm('interestCoverageRatio').mean, bm('interestCoverageRatio').sd, false),
        higherIsBetter: true,
      },
    };
  }, [metrics, benchmark]);

  // 業界比較に基づく総合偏差値（McSS風）
  const industryDeviation = useMemo(() => {
    const items = Object.values(industryComparison);
    const avg = items.reduce((sum, item) => sum + item.deviation, 0) / items.length;
    return Math.round(avg);
  }, [industryComparison]);

  // 業界順位の推定（偏差値から）
  const industryPosition = useMemo(() => {
    // 正規分布近似で順位を推定
    const z = (industryDeviation - 50) / 10;
    // 標準正規分布の累積関数の近似
    const erf = (x) => {
      const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
      const a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
      const sign = x < 0 ? -1 : 1;
      x = Math.abs(x);
      const t = 1.0 / (1.0 + p * x);
      const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
      return sign * y;
    };
    const percentile = (1 + erf(z / Math.sqrt(2))) / 2;
    const topPercent = Math.round((1 - percentile) * 100);
    return {
      topPercent: Math.max(1, Math.min(99, topPercent)),
      bottomPercent: Math.max(1, Math.min(99, 100 - topPercent)),
    };
  }, [industryDeviation]);

  const improvementImpact = useMemo(() => {
    const items = [
      { key: 'equityRatio', name: '自己資本比率', current: metrics.equityRatio, score: scores.equityRatio, unit: '%', cat: '安全性' },
      { key: 'gearingRatio', name: 'ギアリング比率', current: metrics.gearingRatio, score: scores.gearingRatio, unit: '%', cat: '安全性' },
      { key: 'fixedLongTermRatio', name: '固定長期適合率', current: metrics.fixedLongTermRatio, score: scores.fixedLongTermRatio, unit: '%', cat: '安全性' },
      { key: 'currentRatio', name: '流動比率', current: metrics.currentRatio, score: scores.currentRatio, unit: '%', cat: '安全性' },
      { key: 'salesOrdinaryProfitRatio', name: '売上高経常利益率', current: metrics.salesOrdinaryProfitRatio, score: scores.salesOrdinaryProfitRatio, unit: '%', cat: '収益性' },
      { key: 'totalCapitalOrdinaryProfitRatio', name: '総資本経常利益率(ROA)', current: metrics.totalCapitalOrdinaryProfitRatio, score: scores.totalCapitalOrdinaryProfitRatio, unit: '%', cat: '収益性' },
      { key: 'ordinaryProfitGrowthRate', name: '経常利益増加率', current: metrics.ordinaryProfitGrowthRate, score: scores.ordinaryProfitGrowthRate, unit: '%', cat: '成長性' },
      { key: 'debtRepaymentYears', name: '債務償還年数', current: metrics.debtRepaymentYears, score: scores.debtRepaymentYears, unit: '年', cat: '返済能力' },
      { key: 'interestCoverageRatio', name: 'ICレシオ', current: metrics.interestCoverageRatio, score: scores.interestCoverageRatio, unit: '倍', cat: '返済能力' },
    ];
    return items.sort((a, b) => a.score - b.score).slice(0, 5);
  }, [metrics, scores]);

  const formatNumber = (n) => {
    if (Math.abs(n) >= 999) return '−';
    return n.toFixed(1);
  };

  // 危険信号の検出（安田経営診断事務所の解説に基づく）
  const dangerSignals = useMemo(() => {
    const signals = [];
    const salesInterestRate = adjustedData.sales > 0 ? (adjustedData.interestPayment / adjustedData.sales) * 100 : 0;
    
    if (isInsolvent) {
      signals.push({
        type: 'critical',
        title: '実質債務超過',
        desc: '実態BSで純資産がマイナス。新規融資はほぼ不可能な状態です。',
        article: 'bank_overloan',
      });
    }
    if (salesInterestRate >= 1.5) {
      signals.push({
        type: 'critical',
        title: `売上高支払利息率が${salesInterestRate.toFixed(2)}%（過剰債務水準）`,
        desc: '売上高支払利息率1.5%以上の会社のほとんどが過剰債務に陥っているとの指摘があります。2.8%でデフォルト企業以下と判定されます。',
        article: 'interest_ratio',
      });
    } else if (salesInterestRate >= 1.0) {
      signals.push({
        type: 'warning',
        title: `売上高支払利息率が${salesInterestRate.toFixed(2)}%（要警戒）`,
        desc: 'CRDのレポートでは売上高支払利息率1%（小売・卸売0.7%）超で要警戒水準とされています。',
        article: 'interest_ratio',
      });
    }
    if (metrics.debtRepaymentYears > 15 && metrics.debtRepaymentYears < 999) {
      signals.push({
        type: 'critical',
        title: `債務償還年数が${metrics.debtRepaymentYears.toFixed(1)}年（危険水準）`,
        desc: '15年超は銀行から見て返済可能性に重大な懸念がある水準。新規融資が困難になります。',
        article: 'debt_repayment',
      });
    } else if (metrics.debtRepaymentYears > 10 && metrics.debtRepaymentYears <= 15) {
      signals.push({
        type: 'warning',
        title: `債務償還年数が${metrics.debtRepaymentYears.toFixed(1)}年（要注意）`,
        desc: '10年超は過剰債務の目安。営業利益率5%以上を目指す必要があります。',
        article: 'excess_debt',
      });
    }
    if (metrics.salesOrdinaryProfitRatio < 2 && metrics.debtRepaymentYears > 10 && metrics.debtRepaymentYears < 999) {
      signals.push({
        type: 'warning',
        title: '低収益＋過剰債務の組み合わせ',
        desc: '営業利益率が低く、債務償還年数も長い。過剰債務脱出には営業利益率5%以上が必要です。',
        article: 'excess_debt',
      });
    }
    if (adjustedData.consecutiveBlackYears === 0) {
      signals.push({
        type: 'warning',
        title: '当期赤字',
        desc: '赤字決算は債務者区分の引き下げ要因。ただし金融検査マニュアル別冊では、中小・零細企業の赤字を直ちに要注意先以下とはしないとされています。決算説明資料での丁寧な説明が重要です。',
        article: 'bad_news',
      });
    }
    return signals;
  }, [adjustedData, metrics, isInsolvent]);

  const actionMap = {
    equityRatio: '【改善方法】①利益を計上して内部留保を厚くする ②不要資産（遊休固定資産・不良在庫）を処分し総資産を圧縮 ③役員借入金は実質自己資本としてカウントされる',
    gearingRatio: '【改善方法】①有利子負債の繰上返済 ②利益剰余金の積み上げで自己資本を強化 ③不要不急の借入を抑制',
    fixedLongTermRatio: '【改善方法】①過大な設備投資の見直し ②短期借入を長期借入に切り替え ③遊休固定資産の売却',
    currentRatio: '【改善方法】①短期借入を長期借入に組み替え ②売掛金の早期回収・在庫圧縮 ③手元現預金の積み増し',
    salesOrdinaryProfitRatio: '【改善方法】①粗利率の改善（値上げ・原価低減） ②販管費の見直し ③支払利息の削減（金利交渉・繰上返済）',
    totalCapitalOrdinaryProfitRatio: '【改善方法】①総資産のスリム化（不良資産処分） ②売上・利益の拡大 ③遊休資産の活用または売却',
    ordinaryProfitGrowthRate: '【改善方法】①売上拡大策の実行 ②コスト構造の見直しによる利益体質強化 ③新規事業・新規顧客開拓',
    debtRepaymentYears: '【改善方法】①営業利益の最大化（収益改善） ②適正な減価償却の計上でCF拡大 ③借入金の計画的返済（無理な繰上返済は逆効果）',
    interestCoverageRatio: '【改善方法】①営業利益の拡大 ②借入金利の引下げ交渉 ③高金利借入の借換',
  };

  // 指標ごとの推奨記事マッピング
  const articleMap = {
    equityRatio: ['black_red_rating', 'tax_saving'],
    gearingRatio: ['black_red_rating', 'excess_debt'],
    fixedLongTermRatio: ['debt_repayment'],
    currentRatio: ['current_ratio', 'cashflow'],
    salesOrdinaryProfitRatio: ['excess_debt', 'interest_ratio'],
    totalCapitalOrdinaryProfitRatio: ['black_red_rating', 'excess_debt'],
    consecutiveBlackYears: ['basics', 'black_red_rating'],
    ordinaryProfitGrowthRate: ['basics'],
    equityAmount: ['tax_saving', 'basics'],
    salesAmount: ['basics'],
    debtRepaymentYears: ['debt_repayment', 'excess_debt'],
    interestCoverageRatio: ['interest_ratio', 'excess_debt'],
    cashFlowAmount: ['cashflow', 'debt_repayment'],
  };

  const totalAssetsAdj = adjustedData.currentAssets + adjustedData.fixedAssets;
  const totalLiabilitiesAdj = adjustedData.currentLiabilities + adjustedData.fixedLiabilities;
  const cashFlowAdj = adjustedData.operatingProfit + adjustedData.depreciation;

  const formulaMap = {
    equityRatio: {
      formula: '自己資本比率（%） = 純資産 ÷ 総資産 × 100',
      numerator: { label: '純資産', value: adjustedData.netAssets },
      denominator: { label: '総資産', value: totalAssetsAdj },
      meaning: '会社の総資産に占める自己資本の割合。返済不要な資金で経営が成り立っている度合いを示す、銀行が最重視する安全性指標。',
      source: '中小機構「経営自己診断システム」の安全性10指標／経産省ローカルベンチマーク6指標に採用',
    },
    gearingRatio: {
      formula: 'ギアリング比率（%） = 負債合計 ÷ 自己資本 × 100',
      numerator: { label: '負債合計', value: totalLiabilitiesAdj },
      denominator: { label: '自己資本', value: adjustedData.netAssets },
      meaning: '自己資本に対する負債の割合（レバレッジ比率）。低いほど借金依存度が低く、財務が安定していると評価される。',
      source: '民間金融機関の財務スコアリングモデルにおける安全性指標',
    },
    fixedLongTermRatio: {
      formula: '固定長期適合率（%） = 固定資産 ÷ (固定負債＋自己資本) × 100',
      numerator: { label: '固定資産', value: adjustedData.fixedAssets },
      denominator: { label: '固定負債＋自己資本', value: adjustedData.fixedLiabilities + adjustedData.netAssets },
      meaning: '長期的に保有する固定資産を、返済不要な自己資本と長期借入金でまかなえているかを示す。100%超は短期資金で固定資産投資をしている危険な状態。',
      source: '中小機構「経営自己診断システム」の安全性10指標に含まれる',
    },
    currentRatio: {
      formula: '流動比率（%） = 流動資産 ÷ 流動負債 × 100',
      numerator: { label: '流動資産', value: adjustedData.currentAssets },
      denominator: { label: '流動負債', value: adjustedData.currentLiabilities },
      meaning: '1年以内に支払うべき負債に対して、1年以内に現金化できる資産がどれだけあるかを示す短期支払能力の指標。',
      source: '中小機構「経営自己診断システム」の安全性10指標に含まれる',
    },
    salesOrdinaryProfitRatio: {
      formula: '売上高経常利益率（%） = 経常利益 ÷ 売上高 × 100',
      numerator: { label: '経常利益', value: adjustedData.ordinaryProfit },
      denominator: { label: '売上高', value: adjustedData.sales },
      meaning: '売上に対してどれだけ経常利益を稼げているかを示す収益性の代表的指標。',
      source: '中小機構「経営自己診断システム」収益性指標／経産省ローカルベンチマーク6指標',
    },
    totalCapitalOrdinaryProfitRatio: {
      formula: '総資本経常利益率(ROA)（%） = 経常利益 ÷ 総資産 × 100',
      numerator: { label: '経常利益', value: adjustedData.ordinaryProfit },
      denominator: { label: '総資産', value: totalAssetsAdj },
      meaning: '投下した資本全体に対してどれだけ効率的に利益を生み出しているかを示す。資産活用の効率性を測る指標。',
      source: '中小機構「経営自己診断システム」収益性指標',
    },
    consecutiveBlackYears: {
      formula: '連続黒字年数 = 経常利益がプラスである連続した期数',
      numerator: { label: '連続黒字年数', value: adjustedData.consecutiveBlackYears },
      denominator: null,
      meaning: '安定的に黒字を継続できているかを示す。継続性が銀行から高く評価される。',
      source: '金融検査マニュアル別冊が示す「赤字の継続性」評価の考え方に対応',
    },
    ordinaryProfitGrowthRate: {
      formula: '経常利益増加率（%） = (当期経常利益 − 前期経常利益) ÷ 前期経常利益 × 100',
      numerator: { label: '当期 − 前期', value: adjustedData.ordinaryProfit - adjustedData.prevOrdinaryProfit },
      denominator: { label: '前期経常利益', value: adjustedData.prevOrdinaryProfit },
      meaning: '前期と比較した経常利益の成長率。企業の成長性を測る重要指標。',
      source: '経産省ローカルベンチマーク6指標の成長性指標',
    },
    equityAmount: {
      formula: '自己資本額 = 純資産（絶対額）',
      numerator: { label: '純資産', value: adjustedData.netAssets },
      denominator: null,
      meaning: '自己資本の絶対額。規模が大きいほど融資可能額も増える。',
      source: '民間金融機関の財務スコアリング規模指標',
    },
    salesAmount: {
      formula: '売上高 = 当期の売上高（絶対額）',
      numerator: { label: '売上高', value: adjustedData.sales },
      denominator: null,
      meaning: '企業の事業規模を示す指標。',
      source: '民間金融機関の財務スコアリング規模指標',
    },
    debtRepaymentYears: {
      formula: '債務償還年数（年） = 有利子負債 ÷ キャッシュフロー\nキャッシュフロー = 営業利益 ＋ 減価償却費',
      numerator: { label: '有利子負債', value: adjustedData.interestBearingDebt },
      denominator: { label: 'CF（営業利益＋減価償却費）', value: cashFlowAdj },
      meaning: '現在の借入金を毎年のCFで返済すると何年かかるかを示す。7年以内が目安、10年超は要注意。',
      source: '中小機構「経営自己診断システム」／経産省ローカルベンチマーク6指標の双方に採用',
    },
    interestCoverageRatio: {
      formula: 'インタレストカバレッジレシオ（倍） = 営業利益 ÷ 支払利息',
      numerator: { label: '営業利益', value: adjustedData.operatingProfit },
      denominator: { label: '支払利息', value: adjustedData.interestPayment },
      meaning: '営業利益で支払利息を何倍まかなえるかを示す利息支払能力の指標。',
      source: '中小機構「経営自己診断システム」／経産省ローカルベンチマーク6指標の双方に採用',
    },
    cashFlowAmount: {
      formula: 'キャッシュフロー額 = 営業利益 ＋ 減価償却費',
      numerator: { label: '営業利益＋減価償却費', value: cashFlowAdj },
      denominator: null,
      meaning: '本業から生み出される現金創出力。返済原資の規模を示す。',
      source: '民間金融機関の財務スコアリング返済能力指標',
    },
  };

  const handleAccChange = (key, value) => {
    setAccounts(prev => ({ ...prev, [key]: Number(value) || 0 }));
  };

  const resetSim = () => {
    setSimAdjust({ netAssets: 0, interestBearingDebt: 0, ordinaryProfit: 0, currentAssets: 0, currentLiabilities: 0 });
  };

  const toggleSection = (key) => {
    setExpandedSection(prev => ({ ...prev, [key]: !prev[key] }));
  };

  // ========================================
  // データ管理：CSV/Excel取込ロジック
  // ========================================
  
  // ロング形式CSVをパース：industryCode, indicatorCode, mean, top25, sd
  const parseLongFormatData = (rows) => {
    const errors = [];
    const parsed = {};
    let validCount = 0;
    
    rows.forEach((row, idx) => {
      if (idx === 0) return; // ヘッダー行スキップ
      if (!row || row.length < 5) return;
      
      const [industryCode, indicatorCode, meanStr, top25Str, sdStr] = row;
      if (!industryCode || !indicatorCode) return;
      
      // バリデーション
      const industryExists = INDUSTRY_MASTER.find(i => i.code === String(industryCode).trim());
      const indicatorExists = INDICATORS_38.find(i => i.code === String(indicatorCode).trim());
      
      if (!industryExists) {
        errors.push(`行${idx + 1}: 業種コード「${industryCode}」が見つかりません`);
        return;
      }
      if (!indicatorExists) {
        errors.push(`行${idx + 1}: 指標コード「${indicatorCode}」が見つかりません`);
        return;
      }
      
      const mean = parseFloat(meanStr);
      const top25 = parseFloat(top25Str);
      const sd = parseFloat(sdStr);
      
      if (isNaN(mean) || isNaN(top25) || isNaN(sd)) {
        errors.push(`行${idx + 1}: 数値が不正です`);
        return;
      }
      
      const code = String(industryCode).trim();
      if (!parsed[code]) parsed[code] = {};
      parsed[code][String(indicatorCode).trim()] = { mean, top25, sd };
      validCount++;
    });
    
    return { parsed, errors, validCount };
  };

  // CSV文字列をパース
  const parseCSV = (text) => {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    return lines.map(line => line.split(',').map(c => c.trim()));
  };

  // ファイル取込ハンドラ
  const handleFileImport = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setImportError(null);
    setImportSuccess(null);
    setImportPreview(null);
    
    try {
      const fileName = file.name.toLowerCase();
      let rows = [];
      
      if (fileName.endsWith('.csv')) {
        const text = await file.text();
        rows = parseCSV(text);
      } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      } else {
        throw new Error('CSV または Excel (.xlsx) ファイルを選択してください');
      }
      
      const { parsed, errors, validCount } = parseLongFormatData(rows);
      
      if (validCount === 0) {
        throw new Error('有効なデータが見つかりませんでした。フォーマットを確認してください。');
      }
      
      setImportPreview({
        data: parsed,
        validCount,
        errors,
        industryCount: Object.keys(parsed).length,
        fileName: file.name,
      });
    } catch (err) {
      setImportError(err.message);
    }
    
    // ファイル選択リセット（同じファイルを再度選択できるように）
    event.target.value = '';
  };

  // プレビューを反映
  const applyImport = () => {
    if (!importPreview) return;
    
    // 既存データとマージ（新データで上書き）
    const merged = { ...benchmarkData };
    Object.keys(importPreview.data).forEach(code => {
      merged[code] = { ...merged[code], ...importPreview.data[code] };
    });
    
    setBenchmarkData(merged);
    setDataVersion(prev => ({
      ...prev,
      importedAt: new Date().toLocaleString('ja-JP'),
      version: `カスタム取込（${importPreview.fileName}）`,
    }));
    setImportSuccess(`✅ ${importPreview.validCount}件のデータを反映しました（${importPreview.industryCount}業種）`);
    setImportPreview(null);
    
    // 3秒後にメッセージを消す
    setTimeout(() => setImportSuccess(null), 5000);
  };

  // プレビューキャンセル
  const cancelImport = () => {
    setImportPreview(null);
    setImportError(null);
  };

  // テンプレートCSVをダウンロード
  const downloadTemplate = () => {
    const header = ['業種コード', '指標コード', '平均値', '上位25%ライン', '標準偏差'];
    const rows = [header.join(',')];
    
    // 全業種×全指標の組み合わせを空欄で出力（一部にサンプル値）
    INDUSTRY_MASTER.forEach(ind => {
      INDICATORS_38.forEach(indicator => {
        const sample = DEFAULT_BENCHMARK_DATA[ind.code]?.[indicator.code];
        rows.push([
          ind.code,
          indicator.code,
          sample?.mean || '',
          sample?.top25 || '',
          sample?.sd || '',
        ].join(','));
      });
    });
    
    const csvContent = '\ufeff' + rows.join('\n'); // BOM付きでExcel対応
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `業界平均データ_テンプレート_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Excel形式のテンプレートをダウンロード
  const downloadExcelTemplate = () => {
    const data = [
      ['業種コード', '指標コード', '平均値', '上位25%ライン', '標準偏差', '業種名(参考)', '指標名(参考)']
    ];
    
    INDUSTRY_MASTER.forEach(ind => {
      INDICATORS_38.forEach(indicator => {
        const sample = DEFAULT_BENCHMARK_DATA[ind.code]?.[indicator.code];
        data.push([
          ind.code,
          indicator.code,
          sample?.mean || '',
          sample?.top25 || '',
          sample?.sd || '',
          ind.name,
          indicator.name,
        ]);
      });
    });
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    // 列幅設定
    ws['!cols'] = [
      { wch: 10 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 30 }, { wch: 30 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '業界平均データ');
    
    // 業種マスタシート
    const industrySheet = XLSX.utils.aoa_to_sheet([
      ['業種コード', '業種名', '大分類'],
      ...INDUSTRY_MASTER.map(i => [i.code, i.name, i.parent])
    ]);
    industrySheet['!cols'] = [{ wch: 10 }, { wch: 35 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, industrySheet, '業種マスタ');
    
    // 指標マスタシート
    const indicatorSheet = XLSX.utils.aoa_to_sheet([
      ['指標コード', '指標名', 'カテゴリ', '計算式', '単位', '高い方が良い'],
      ...INDICATORS_38.map(i => [i.code, i.name, i.category, i.formula, i.unit, i.higherIsBetter ? 'YES' : 'NO'])
    ]);
    indicatorSheet['!cols'] = [{ wch: 30 }, { wch: 30 }, { wch: 20 }, { wch: 40 }, { wch: 8 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, indicatorSheet, '指標マスタ');
    
    XLSX.writeFile(wb, `業界平均データ_テンプレート_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // ----- 入力画面 -----
  if (step === 'input') {
    const inputField = (key, label, hint) => (
      <div className="grid grid-cols-5 gap-2 items-center mb-2">
        <label className="text-sm text-slate-700 col-span-2 flex items-center gap-1">
          {label}
          {hint && <span className="text-xs text-slate-400" title={hint}>ⓘ</span>}
        </label>
        <input 
          type="number" 
          value={accounts[key]}
          onChange={(e) => handleAccChange(key, e.target.value)}
          className="col-span-3 px-2 py-1 border border-slate-300 rounded text-right text-sm focus:ring-1 focus:ring-blue-500 focus:border-transparent"
        />
      </div>
    );

    return (
      <div className="min-h-screen bg-slate-50 p-6">
        <div className="max-w-5xl mx-auto">
          <header className="mb-6">
            <div className="flex items-center gap-3 mb-2">
              <div className="w-10 h-10 bg-blue-700 rounded-lg flex items-center justify-center">
                <BarChart3 className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-2xl font-bold text-slate-800">銀行格付け改善アドバイザー</h1>
                <p className="text-xs text-slate-500">勘定科目レベルで実態BSを再構築し、銀行員視点で診断</p>
              </div>
            </div>
          </header>

          {/* 準拠する公的基準 */}
          <div className="bg-gradient-to-br from-slate-50 to-blue-50 rounded-xl shadow-sm border border-blue-200 p-5 mb-4">
            <div className="flex items-center gap-2 mb-3">
              <Shield className="w-5 h-5 text-blue-700" />
              <h2 className="text-base font-semibold text-slate-800">本ツールが準拠する公的基準・公開ロジック</h2>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              <a href="https://www.fsa.go.jp/manual/manualj/yokin_b.html" target="_blank" rel="noopener noreferrer"
                className="bg-white rounded-lg p-3 border border-slate-200 hover:border-blue-400 hover:shadow-md transition-all group">
                <div className="flex items-start gap-2 mb-2">
                  <span className="text-xs font-bold text-white bg-blue-700 px-2 py-0.5 rounded">金融庁</span>
                  <ExternalLink className="w-3 h-3 text-slate-400 group-hover:text-blue-700 ml-auto" />
                </div>
                <p className="text-sm font-semibold text-slate-800 mb-1">金融検査マニュアル別冊</p>
                <p className="text-xs text-slate-600">〔中小企業融資編〕に基づく債務者区分（正常先・要注意先・要管理先・破綻懸念先・破綻先）の判定ロジックを採用</p>
              </a>
              <a href="https://k-sindan.smrj.go.jp/" target="_blank" rel="noopener noreferrer"
                className="bg-white rounded-lg p-3 border border-slate-200 hover:border-blue-400 hover:shadow-md transition-all group">
                <div className="flex items-start gap-2 mb-2">
                  <span className="text-xs font-bold text-white bg-green-700 px-2 py-0.5 rounded">中小機構</span>
                  <ExternalLink className="w-3 h-3 text-slate-400 group-hover:text-blue-700 ml-auto" />
                </div>
                <p className="text-sm font-semibold text-slate-800 mb-1">経営自己診断システム</p>
                <p className="text-xs text-slate-600">CRD（中小企業信用リスク情報DB）200万社データに基づく5項目27指標の評価枠組みを参照</p>
              </a>
              <a href="https://www.meti.go.jp/policy/economy/keiei_innovation/sangyokinyu/locaben/" target="_blank" rel="noopener noreferrer"
                className="bg-white rounded-lg p-3 border border-slate-200 hover:border-blue-400 hover:shadow-md transition-all group">
                <div className="flex items-start gap-2 mb-2">
                  <span className="text-xs font-bold text-white bg-orange-600 px-2 py-0.5 rounded">経産省</span>
                  <ExternalLink className="w-3 h-3 text-slate-400 group-hover:text-blue-700 ml-auto" />
                </div>
                <p className="text-sm font-semibold text-slate-800 mb-1">ローカルベンチマーク</p>
                <p className="text-xs text-slate-600">経産省が公表する「企業の健康診断」ツールの6つの財務指標を補完的に参照</p>
              </a>
            </div>
            <p className="text-xs text-slate-500 mt-3 leading-relaxed">
              💡 上記の公的機関が公開している指標体系・債務者区分の考え方を基に、藤井会計事務所等が公開する民間金融機関のスコアリング配点を補完して構築しています。CRDモデルそのもののスコアリングロジックは機密情報として非公開のため、本ツールは公開情報に基づく目安としてご活用ください。
            </p>
          </div>

          {/* 基本情報 */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 mb-4">
            <h2 className="text-base font-semibold text-slate-800 mb-3 flex items-center gap-2">
              <span className="w-1 h-5 bg-blue-700 rounded-full"></span>
              基本情報
              <span className="text-xs font-normal text-blue-700 bg-blue-50 px-2 py-0.5 rounded">業界比較の精度を決める重要項目</span>
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">業種（産業中分類）</label>
                <select 
                  value={industryCode} 
                  onChange={(e) => setIndustryCode(e.target.value)}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                >
                  {Object.entries(
                    INDUSTRY_MASTER.reduce((acc, ind) => {
                      if (!acc[ind.parent]) acc[ind.parent] = [];
                      acc[ind.parent].push(ind);
                      return acc;
                    }, {})
                  ).map(([parent, industries]) => (
                    <optgroup key={parent} label={parent}>
                      {industries.map(ind => (
                        <option key={ind.code} value={ind.code}>
                          {ind.code} {ind.name}
                        </option>
                      ))}
                    </optgroup>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">従業者規模</label>
                <select 
                  value={scale} 
                  onChange={(e) => setScale(e.target.value)}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                >
                  {SCALE_MASTER.map(s => (
                    <option key={s.code} value={s.code}>{s.name}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">連続黒字年数</label>
                <input 
                  type="number" 
                  value={accounts.consecutiveBlackYears}
                  onChange={(e) => handleAccChange('consecutiveBlackYears', e.target.value)}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                />
              </div>
            </div>
            <p className="text-xs text-slate-500 mt-3">
              💡 業種は{INDUSTRY_MASTER.length}業種（日本標準産業分類の中分類）から選択。業種・規模により、中小企業実態基本調査の同業他社データと比較します
            </p>
          </div>

          {/* データ管理セクション */}
          <div className="bg-gradient-to-br from-slate-50 to-blue-50 rounded-xl shadow-sm border border-slate-200 p-5 mb-4">
            <button 
              onClick={() => toggleSection('dataManage')}
              className="w-full flex items-center justify-between mb-2"
            >
              <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                <Database className="w-5 h-5 text-indigo-700" />
                データ管理（業界平均データの更新）
                <span className="text-xs font-normal text-indigo-700 bg-indigo-100 px-2 py-0.5 rounded">運営者向け</span>
              </h2>
              {expandedSection.dataManage ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
            
            {/* 現在のバージョン情報（常時表示） */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-xs mb-2">
              <div className="bg-white rounded p-2 border border-slate-200">
                <p className="text-slate-500 mb-0.5">データバージョン</p>
                <p className="font-semibold text-slate-800 text-[11px]">{dataVersion.version}</p>
              </div>
              <div className="bg-white rounded p-2 border border-slate-200">
                <p className="text-slate-500 mb-0.5">取込日</p>
                <p className="font-semibold text-slate-800">{dataVersion.importedAt}</p>
              </div>
              <div className="bg-white rounded p-2 border border-slate-200">
                <p className="text-slate-500 mb-0.5">業種数</p>
                <p className="font-semibold text-slate-800">{dataVersion.industryCount}業種</p>
              </div>
              <div className="bg-white rounded p-2 border border-slate-200">
                <p className="text-slate-500 mb-0.5">指標数</p>
                <p className="font-semibold text-slate-800">{dataVersion.indicatorCount}指標</p>
              </div>
            </div>

            {expandedSection.dataManage && (
              <div className="mt-4">
                <div className="bg-white rounded-lg p-4 border border-slate-200 mb-3">
                  <h3 className="text-sm font-semibold text-slate-800 mb-2 flex items-center gap-2">
                    <Download className="w-4 h-4" />
                    1. テンプレートをダウンロード
                  </h3>
                  <p className="text-xs text-slate-600 mb-3">
                    まず空のテンプレートをダウンロードして、e-Statや業界統計から数値を転記してください。Excel版には業種マスタ・指標マスタも同梱されます。
                  </p>
                  <div className="flex gap-2">
                    <button 
                      onClick={downloadTemplate}
                      className="px-3 py-1.5 bg-blue-100 text-blue-800 hover:bg-blue-200 rounded text-xs font-medium flex items-center gap-1"
                    >
                      <Download className="w-3 h-3" /> CSV テンプレート
                    </button>
                    <button 
                      onClick={downloadExcelTemplate}
                      className="px-3 py-1.5 bg-green-100 text-green-800 hover:bg-green-200 rounded text-xs font-medium flex items-center gap-1"
                    >
                      <Download className="w-3 h-3" /> Excel テンプレート
                    </button>
                  </div>
                </div>

                <div className="bg-white rounded-lg p-4 border border-slate-200 mb-3">
                  <h3 className="text-sm font-semibold text-slate-800 mb-2 flex items-center gap-2">
                    <Upload className="w-4 h-4" />
                    2. データファイルを選択
                  </h3>
                  <p className="text-xs text-slate-600 mb-3">
                    フォーマット：ロング形式（業種コード, 指標コード, 平均値, 上位25%, 標準偏差）
                  </p>
                  <input 
                    ref={fileInputRef}
                    type="file"
                    accept=".csv,.xlsx,.xls"
                    onChange={handleFileImport}
                    className="text-xs"
                  />
                  
                  {importError && (
                    <div className="mt-3 bg-red-50 border border-red-200 rounded p-3 text-xs text-red-800">
                      <p className="font-semibold flex items-center gap-1"><X className="w-3 h-3" /> エラー</p>
                      <p>{importError}</p>
                    </div>
                  )}
                  
                  {importSuccess && (
                    <div className="mt-3 bg-green-50 border border-green-200 rounded p-3 text-xs text-green-800">
                      <p className="font-semibold">{importSuccess}</p>
                    </div>
                  )}
                </div>

                {importPreview && (
                  <div className="bg-amber-50 border-2 border-amber-300 rounded-lg p-4 mb-3">
                    <h3 className="text-sm font-semibold text-amber-900 mb-2">📋 取込プレビュー</h3>
                    <div className="grid grid-cols-3 gap-2 mb-3 text-xs">
                      <div className="bg-white rounded p-2 border border-amber-200">
                        <p className="text-slate-500">ファイル名</p>
                        <p className="font-semibold text-slate-800 text-[11px]">{importPreview.fileName}</p>
                      </div>
                      <div className="bg-white rounded p-2 border border-amber-200">
                        <p className="text-slate-500">有効データ数</p>
                        <p className="font-semibold text-green-700">{importPreview.validCount}件</p>
                      </div>
                      <div className="bg-white rounded p-2 border border-amber-200">
                        <p className="text-slate-500">業種数</p>
                        <p className="font-semibold text-slate-800">{importPreview.industryCount}業種</p>
                      </div>
                    </div>
                    
                    {importPreview.errors.length > 0 && (
                      <div className="bg-orange-50 border border-orange-200 rounded p-2 mb-3 text-xs">
                        <p className="font-semibold text-orange-800 mb-1">⚠️ 警告: {importPreview.errors.length}件</p>
                        <div className="max-h-32 overflow-y-auto">
                          {importPreview.errors.slice(0, 10).map((err, i) => (
                            <p key={i} className="text-orange-700">・{err}</p>
                          ))}
                          {importPreview.errors.length > 10 && (
                            <p className="text-orange-700 italic">...他 {importPreview.errors.length - 10}件</p>
                          )}
                        </div>
                      </div>
                    )}
                    
                    <p className="text-xs text-slate-700 mb-3">
                      このデータで業界平均を上書きします。よろしいですか？
                    </p>
                    
                    <div className="flex gap-2">
                      <button 
                        onClick={applyImport}
                        className="px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded font-medium text-sm flex items-center gap-1"
                      >
                        <Check className="w-4 h-4" /> 保存して反映
                      </button>
                      <button 
                        onClick={cancelImport}
                        className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded font-medium text-sm flex items-center gap-1"
                      >
                        <X className="w-4 h-4" /> キャンセル
                      </button>
                    </div>
                  </div>
                )}

                <div className="bg-blue-50 rounded-lg p-3 text-xs text-blue-900 border border-blue-200">
                  <p className="font-semibold mb-1">📖 データ取込フォーマット</p>
                  <pre className="font-mono text-[10px] bg-white p-2 rounded mt-1 overflow-x-auto">業種コード,指標コード,平均値,上位25%ライン,標準偏差
09,equityRatio,38.5,60.0,22.0
09,currentRatio,195.0,295.0,105.0
09,salesOrdinaryProfitRatio,3.5,7.5,5.5
...</pre>
                  <p className="mt-2">
                    ・業種コードは「業種マスタ」シート参照（{INDUSTRY_MASTER.length}業種）<br />
                    ・指標コードは「指標マスタ」シート参照（{INDICATORS_38.length}指標）<br />
                    ・データソース推奨：e-Stat → 中小企業実態基本調査 → 産業中分類別データ
                  </p>
                </div>
              </div>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            {/* 資産の部 */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
              <button 
                onClick={() => toggleSection('assets')}
                className="w-full flex items-center justify-between mb-3"
              >
                <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                  <span className="w-1 h-5 bg-blue-700 rounded-full"></span>
                  資産の部（千円）
                </h2>
                {expandedSection.assets ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {expandedSection.assets && (
                <>
                  <p className="text-xs font-semibold text-slate-500 mb-2 mt-3">▼ 流動資産</p>
                  {inputField('cash', '現金・預金')}
                  {inputField('receivablesNotes', '受取手形')}
                  {inputField('receivablesAccount', '売掛金')}
                  {inputField('inventory', '棚卸資産（在庫）')}
                  {inputField('advancesToEmployees', '仮払金・立替金')}
                  {inputField('shortTermLoansToExec', '短期役員貸付金')}
                  {inputField('otherCurrentAssets', 'その他流動資産')}
                  
                  <p className="text-xs font-semibold text-slate-500 mb-2 mt-4">▼ 固定資産</p>
                  {inputField('land', '土地')}
                  {inputField('building', '建物')}
                  {inputField('machinery', '機械装置')}
                  {inputField('longTermLoansToExec', '長期役員貸付金')}
                  {inputField('investmentSecurities', '投資有価証券')}
                  {inputField('securityDeposit', '入居保証金・敷金')}
                  {inputField('otherFixedAssets', 'その他固定資産')}
                  {inputField('deferredAssets', '繰延資産')}
                </>
              )}
            </div>

            {/* 負債・純資産の部 */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
              <button 
                onClick={() => toggleSection('liabilities')}
                className="w-full flex items-center justify-between mb-3"
              >
                <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                  <span className="w-1 h-5 bg-blue-700 rounded-full"></span>
                  負債・純資産の部（千円）
                </h2>
                {expandedSection.liabilities ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              </button>
              {expandedSection.liabilities && (
                <>
                  <p className="text-xs font-semibold text-slate-500 mb-2 mt-3">▼ 流動負債</p>
                  {inputField('payablesNotes', '支払手形')}
                  {inputField('payablesAccount', '買掛金')}
                  {inputField('shortTermBorrowings', '短期借入金')}
                  {inputField('otherCurrentLiabilities', 'その他流動負債')}
                  
                  <p className="text-xs font-semibold text-slate-500 mb-2 mt-4">▼ 固定負債</p>
                  {inputField('longTermBorrowings', '長期借入金')}
                  {inputField('longTermBorrowingsFromExec', '役員借入金 ⭐実質自己資本扱い')}
                  {inputField('otherFixedLiabilities', 'その他固定負債')}
                  
                  <p className="text-xs font-semibold text-slate-500 mb-2 mt-4">▼ 純資産</p>
                  {inputField('capitalStock', '資本金')}
                  {inputField('capitalSurplus', '資本剰余金')}
                  {inputField('retainedEarnings', '利益剰余金')}

                  <h3 className="text-base font-semibold text-slate-800 mt-6 mb-3 flex items-center gap-2">
                    <span className="w-1 h-5 bg-blue-700 rounded-full"></span>
                    損益計算書（千円）
                  </h3>
                  {inputField('sales', '売上高')}
                  {inputField('operatingProfit', '営業利益')}
                  {inputField('ordinaryProfit', '経常利益（当期）')}
                  {inputField('prevOrdinaryProfit', '経常利益（前期）')}
                  {inputField('depreciation', '減価償却費')}
                  {inputField('interestPayment', '支払利息')}
                </>
              )}
            </div>
          </div>

          {/* 実態BS調整項目 */}
          <div className="bg-gradient-to-br from-amber-50 to-orange-50 rounded-xl shadow-sm border border-amber-200 p-5 mb-4">
            <button 
              onClick={() => toggleSection('adjust')}
              className="w-full flex items-center justify-between mb-3"
            >
              <h2 className="text-base font-semibold text-slate-800 flex items-center gap-2">
                <AlertTriangle className="w-5 h-5 text-amber-600" />
                実態BS修正項目（銀行員視点）
              </h2>
              {expandedSection.adjust ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
            {expandedSection.adjust && (
              <>
                <p className="text-xs text-slate-600 mb-4 bg-white p-3 rounded">
                  💡 銀行員は決算書をそのまま信用せず、不良資産を控除した「実態BS」で評価します。回収困難な資産・不良在庫・含み損などを入力してください。
                </p>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <p className="text-xs font-semibold text-red-700 mb-2">▼ 控除すべき不良資産</p>
                    {inputField('badInventory', '不良在庫・長期滞留在庫')}
                    {inputField('badReceivables', '回収困難売掛金・滞留債権')}
                    {inputField('badAdvances', '不良仮払金（回収困難）')}
                    {inputField('depreciationShortage', '減価償却不足額')}
                  </div>
                  <div>
                    <p className="text-xs font-semibold text-blue-700 mb-2">▼ 時価修正（マイナスは含み損）</p>
                    {inputField('landMarketValueAdjust', '土地の時価調整')}
                    {inputField('securitiesMarketValueAdjust', '有価証券の時価調整')}
                  </div>
                </div>
                <div className="mt-4 bg-white rounded-lg p-3 border border-slate-200">
                  <p className="text-xs text-slate-600 mb-2">⚙️ 自動適用される修正（入力不要）</p>
                  <ul className="text-xs text-slate-700 space-y-1">
                    <li>・役員貸付金 → 全額不良資産として控除</li>
                    <li>・繰延資産 → ゼロ評価（全額控除）</li>
                    <li>・入居保証金 → 30%評価（70%控除）</li>
                    <li>・役員借入金 → 実質自己資本としてカウント</li>
                  </ul>
                </div>
              </>
            )}
          </div>

          <button 
            onClick={() => setStep('result')}
            className="w-full bg-blue-700 hover:bg-blue-800 text-white font-semibold py-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2"
          >
            診断結果を見る
            <ChevronRight className="w-5 h-5" />
          </button>

          <p className="text-xs text-slate-500 mt-4 text-center">
            ※本ツールは公開情報に基づく目安です
          </p>
        </div>
      </div>
    );
  }

  // ----- 結果画面 -----
  const adj = actualBSAdjustments;
  return (
    <div className="min-h-screen bg-slate-50 p-6">
      <div className="max-w-6xl mx-auto">
        <header className="mb-6 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-700 rounded-lg flex items-center justify-center">
              <BarChart3 className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-800">診断結果</h1>
              <p className="text-xs text-slate-500">銀行格付け改善アドバイザー</p>
            </div>
          </div>
          <button 
            onClick={() => setStep('input')}
            className="px-4 py-2 text-sm text-slate-600 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors"
          >
            ← 入力画面に戻る
          </button>
        </header>

        {/* BS切り替えトグル */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 mb-4">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-semibold text-slate-800">評価方式の切り替え</p>
              <p className="text-xs text-slate-500">表面BSは決算書のまま、実態BSは銀行員視点での修正後</p>
            </div>
            <div className="flex gap-2">
              <button 
                onClick={() => setUseActualBS(false)}
                className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 ${
                  !useActualBS ? 'bg-slate-200 text-slate-800' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                }`}
              >
                <FileText className="w-4 h-4" />
                表面BS
              </button>
              <button 
                onClick={() => setUseActualBS(true)}
                className={`px-4 py-2 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 ${
                  useActualBS ? 'bg-blue-700 text-white' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                }`}
              >
                <Eye className="w-4 h-4" />
                実態BS（銀行員視点）
              </button>
            </div>
          </div>
        </div>

        {/* 実態BS差異サマリ */}
        {useActualBS && (
          <div className="bg-gradient-to-br from-amber-50 to-orange-50 rounded-xl shadow-sm border border-amber-200 p-5 mb-6">
            <h2 className="text-base font-semibold text-slate-800 mb-3 flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-amber-600" />
              実態BS変換による影響
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-3">
              <div className="bg-white rounded-lg p-3">
                <p className="text-xs text-slate-500">表面上の純資産</p>
                <p className="text-xl font-bold text-slate-800">{surfaceBS.netAssets.toLocaleString()}<span className="text-sm">千円</span></p>
              </div>
              <div className="bg-white rounded-lg p-3">
                <p className="text-xs text-slate-500">修正調整額</p>
                <p className={`text-xl font-bold ${adj.totalEquityAdjust >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                  {adj.totalEquityAdjust >= 0 ? '+' : ''}{adj.totalEquityAdjust.toLocaleString()}<span className="text-sm">千円</span>
                </p>
              </div>
              <div className="bg-white rounded-lg p-3">
                <p className="text-xs text-slate-500">実態純資産</p>
                <p className={`text-xl font-bold ${baseBS.netAssets >= 0 ? 'text-blue-700' : 'text-red-600'}`}>
                  {baseBS.netAssets.toLocaleString()}<span className="text-sm">千円</span>
                </p>
              </div>
            </div>
            {isInsolvent && (
              <div className="bg-red-100 border border-red-300 rounded-lg p-3 mb-3">
                <p className="text-sm font-bold text-red-800 flex items-center gap-2">
                  <AlertCircle className="w-4 h-4" />
                  実質債務超過です。銀行内では即「格付8（警戒先）」相当となります。
                </p>
              </div>
            )}
            <details className="text-xs">
              <summary className="cursor-pointer font-semibold text-slate-700 mb-2">▼ 修正項目の内訳を見る</summary>
              <div className="mt-2 space-y-1 bg-white rounded p-3">
                {[
                  { label: '不良在庫の控除', value: adj.inventoryDeduction },
                  { label: '回収困難売掛金の控除', value: adj.receivablesDeduction },
                  { label: '不良仮払金の控除', value: adj.advancesDeduction },
                  { label: '役員貸付金の控除（自動）', value: adj.execLoansDeduction },
                  { label: '繰延資産のゼロ評価（自動）', value: adj.deferredDeduction },
                  { label: '入居保証金70%控除（自動）', value: adj.securityDepositDeduction },
                  { label: '減価償却不足の控除', value: adj.depreciationDeduction },
                  { label: '土地時価調整', value: adj.landAdjust },
                  { label: '有価証券時価調整', value: adj.securitiesAdjust },
                  { label: '役員借入金を自己資本へ振替（自動）', value: adj.execBorrowingsAddition },
                ].filter(item => item.value !== 0).map((item, i) => (
                  <div key={i} className="flex justify-between border-b border-slate-100 py-1">
                    <span className="text-slate-700">{item.label}</span>
                    <span className={`font-mono font-semibold ${item.value >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                      {item.value >= 0 ? '+' : ''}{item.value.toLocaleString()}
                    </span>
                  </div>
                ))}
              </div>
            </details>
          </div>
        )}

        {/* メインダッシュボード */}
        <div className="bg-gradient-to-br from-blue-700 to-blue-900 rounded-2xl shadow-lg p-6 mb-6 text-white">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <Award className="w-5 h-5" />
              <h2 className="text-lg font-semibold">貴社の現在地（{useActualBS ? '実態BS' : '表面BS'}ベース）</h2>
            </div>
            <div className="flex gap-2 text-xs">
              <span className="bg-white/20 backdrop-blur px-2 py-1 rounded">📜 金融庁マニュアル準拠</span>
              <span className="bg-white/20 backdrop-blur px-2 py-1 rounded">🏛️ 中小機構CRD参照</span>
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div className="bg-white/10 backdrop-blur rounded-xl p-4">
              <p className="text-xs text-blue-100 mb-1">総合スコア</p>
              <p className="text-3xl font-bold">{categoryScores.total100}<span className="text-base ml-1">/100</span></p>
              <p className="text-xs text-blue-100 mt-1">{categoryScores.total}/130点</p>
            </div>
            <div className="bg-white/10 backdrop-blur rounded-xl p-4">
              <p className="text-xs text-blue-100 mb-1">信用格付け</p>
              <p className="text-3xl font-bold">{rating.level}<span className="text-base ml-1">/10</span></p>
              <p className="text-xs text-blue-100 mt-1">{rating.label}</p>
            </div>
            <div className="bg-white/10 backdrop-blur rounded-xl p-4">
              <p className="text-xs text-blue-100 mb-1">偏差値</p>
              <p className="text-3xl font-bold">{deviation}</p>
              <p className="text-xs text-blue-100 mt-1">🏛️ CRDランク {crdRank.rank}（{crdRank.desc}）</p>
            </div>
            <div className="bg-white/10 backdrop-blur rounded-xl p-4">
              <p className="text-xs text-blue-100 mb-1">債務者区分</p>
              <p className="text-2xl font-bold">{debtorClass.name}</p>
              <p className="text-xs text-blue-100 mt-1">📜 金融検査マニュアル別冊準拠</p>
            </div>
          </div>
          <div className="mt-4 p-3 bg-white/10 backdrop-blur rounded-lg">
            <p className="text-sm">{rating.desc}</p>
          </div>
        </div>

        {/* 格付け位置 */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 mb-6">
          <p className="text-sm font-semibold text-slate-700 mb-3">格付けポジション</p>
          <div className="relative h-12 bg-gradient-to-r from-red-500 via-yellow-400 via-green-400 to-blue-600 rounded-lg overflow-hidden">
            <div className="absolute inset-0 flex items-center justify-between px-3 text-xs font-semibold text-white">
              <span>9-10<br/>危険</span>
              <span>7-8<br/>要注意</span>
              <span>5-6<br/>標準</span>
              <span>3-4<br/>良好</span>
              <span>1-2<br/>最上位</span>
            </div>
            <div 
              className="absolute top-0 h-full w-1 bg-slate-900 shadow-lg transition-all duration-500"
              style={{ left: `${Math.min(98, Math.max(2, ((10 - rating.level) / 9) * 100))}%` }}
            >
              <div className="absolute -top-2 left-1/2 -translate-x-1/2 w-0 h-0 border-l-4 border-r-4 border-t-4 border-transparent border-t-slate-900"></div>
            </div>
          </div>
          <div className="mt-2 text-xs text-slate-500 text-center">
            ◀ 危険ゾーン　　　　　　　　　　　　　　　目標：格付4以上 ▶
          </div>
        </div>

        {/* 業界内ポジション（フェーズ2新規） */}
        <div className="bg-gradient-to-br from-indigo-50 to-purple-50 rounded-xl shadow-sm border border-indigo-200 p-6 mb-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
              <Users className="w-5 h-5 text-indigo-700" />
              業界内ポジション
            </h2>
            <div className="text-xs">
              <span className="bg-slate-100 text-slate-700 px-2 py-1 rounded mr-1">{benchmark.parent || ''}</span>
              <span className="bg-indigo-100 text-indigo-800 px-2 py-1 rounded mr-1">{benchmark.name}</span>
              <span className="bg-purple-100 text-purple-800 px-2 py-1 rounded">{SCALE_MASTER.find(s => s.code === scale)?.name}</span>
            </div>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div className="bg-white rounded-xl p-4 border border-indigo-100">
              <p className="text-xs text-slate-500 mb-1">業界内総合偏差値</p>
              <p className="text-4xl font-bold text-indigo-700">{industryDeviation}</p>
              <p className="text-xs text-slate-500 mt-1">業界平均=50・上位=高</p>
            </div>
            <div className="bg-white rounded-xl p-4 border border-indigo-100">
              <p className="text-xs text-slate-500 mb-1">推定順位</p>
              <p className="text-3xl font-bold text-indigo-700">上位 {industryPosition.topPercent}%</p>
              <p className="text-xs text-slate-500 mt-1">同業同規模 中小企業内</p>
            </div>
            <div className="bg-white rounded-xl p-4 border border-indigo-100">
              <p className="text-xs text-slate-500 mb-1">評価</p>
              <p className={`text-2xl font-bold ${
                industryDeviation >= 60 ? 'text-blue-700' :
                industryDeviation >= 53 ? 'text-cyan-600' :
                industryDeviation >= 47 ? 'text-green-600' :
                industryDeviation >= 40 ? 'text-orange-600' : 'text-red-600'
              }`}>
                {industryDeviation >= 60 ? 'A：優良' :
                 industryDeviation >= 53 ? 'B：良好' :
                 industryDeviation >= 47 ? 'C：標準' :
                 industryDeviation >= 40 ? 'D：要改善' : 'E：劣後'}
              </p>
              <p className="text-xs text-slate-500 mt-1">業界内位置</p>
            </div>
          </div>
          <div className="bg-white rounded-lg p-3 text-xs text-slate-600 border border-indigo-100">
            📊 出典：中小企業実態基本調査（経済産業省・中小企業庁、令和6年確報）の業種別・規模別データに基づき、貴社の値と比較しています
          </div>
        </div>

        {/* 業界比較バー（指標別） */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-6">
          <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
            <BarChart3 className="w-5 h-5 text-blue-700" />
            業界平均との比較（主要8指標）
          </h2>
          <div className="space-y-4">
            {[
              { key: 'equityRatio', label: '自己資本比率', unit: '%' },
              { key: 'currentRatio', label: '流動比率', unit: '%' },
              { key: 'gearingRatio', label: 'ギアリング比率', unit: '%' },
              { key: 'fixedLongTermRatio', label: '固定長期適合率', unit: '%' },
              { key: 'salesOrdinaryProfitRatio', label: '売上高経常利益率', unit: '%' },
              { key: 'totalCapitalOrdinaryProfitRatio', label: '総資本経常利益率(ROA)', unit: '%' },
              { key: 'debtRepaymentYears', label: '債務償還年数', unit: '年' },
              { key: 'interestCoverageRatio', label: 'ICレシオ', unit: '倍' },
            ].map(item => {
              const comp = industryComparison[item.key];
              if (!comp) return null;
              const isBetter = comp.higherIsBetter ? comp.myValue >= comp.mean : comp.myValue <= comp.mean;
              const isTop = comp.higherIsBetter ? comp.myValue >= comp.top25 : comp.myValue <= comp.top25;
              return (
                <div key={item.key}>
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-sm font-medium text-slate-700">{item.label}</span>
                    <div className="flex items-center gap-3 text-xs">
                      <span className="text-slate-500">業界平均: <span className="font-mono">{comp.mean.toFixed(1)}{item.unit}</span></span>
                      <span className="text-slate-500">上位25%: <span className="font-mono">{comp.top25.toFixed(1)}{item.unit}</span></span>
                      <span className={`font-bold ${isTop ? 'text-blue-700' : isBetter ? 'text-green-700' : 'text-red-600'}`}>
                        貴社: {formatNumber(comp.myValue)}{item.unit}
                      </span>
                      <span className={`px-2 py-0.5 rounded text-xs ${
                        comp.deviation >= 60 ? 'bg-blue-100 text-blue-800' :
                        comp.deviation >= 53 ? 'bg-cyan-100 text-cyan-800' :
                        comp.deviation >= 47 ? 'bg-green-100 text-green-800' :
                        comp.deviation >= 40 ? 'bg-orange-100 text-orange-800' :
                        'bg-red-100 text-red-800'
                      }`}>
                        偏差値 {comp.deviation}
                      </span>
                    </div>
                  </div>
                  {/* 比較バー */}
                  <div className="relative h-6 bg-slate-100 rounded overflow-hidden">
                    {/* 業界平均ライン */}
                    <div className="absolute top-0 h-full w-0.5 bg-slate-400 z-10" style={{ left: '50%' }}>
                      <div className="absolute -top-1 left-1/2 -translate-x-1/2 text-[10px] text-slate-500 whitespace-nowrap" style={{ marginTop: '-14px' }}>平均</div>
                    </div>
                    {/* 上位25%ライン */}
                    <div 
                      className="absolute top-0 h-full w-0.5 bg-blue-500 z-10" 
                      style={{ left: comp.higherIsBetter ? '75%' : '25%' }}
                    >
                      <div className="absolute -top-1 left-1/2 -translate-x-1/2 text-[10px] text-blue-600 whitespace-nowrap" style={{ marginTop: '-14px' }}>上位25%</div>
                    </div>
                    {/* 貴社マーカー */}
                    <div 
                      className={`absolute top-0 h-full w-2 z-20 transition-all duration-500 ${
                        isTop ? 'bg-blue-700' : isBetter ? 'bg-green-600' : 'bg-red-500'
                      }`}
                      style={{ left: `${Math.max(2, Math.min(98, (comp.deviation - 25) / 50 * 100))}%` }}
                    >
                      <div className="absolute -top-1 left-1/2 -translate-x-1/2 text-[10px] font-bold text-slate-800 whitespace-nowrap" style={{ marginTop: '-14px' }}>貴社</div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
          <div className="mt-4 p-3 bg-slate-50 rounded text-xs text-slate-600">
            ▶ 業界平均（偏差値50）を中心に、貴社の位置を表示しています。青いマーカーが右側にあれば業界平均を上回り、左側にあれば下回ります。
            <br />
            ▶ 緑：業界平均を上回る｜赤：業界平均を下回る｜青：上位25%以内
          </div>
        </div>

        {/* カテゴリ別 */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          {[
            { name: '安全性', score: categoryScores.safety, max: 40, icon: '🛡️' },
            { name: '収益性', score: categoryScores.profitability, max: 30, icon: '💰' },
            { name: '成長性', score: categoryScores.growth, max: 30, icon: '📈' },
            { name: '返済能力', score: categoryScores.repayment, max: 30, icon: '🔄' },
          ].map(cat => {
            const pct = (cat.score / cat.max) * 100;
            return (
              <div key={cat.name} className="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-sm font-semibold text-slate-700">{cat.icon} {cat.name}</span>
                  <span className="text-xs text-slate-500">{cat.score}/{cat.max}</span>
                </div>
                <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div 
                    className={`h-full transition-all duration-500 ${
                      pct >= 70 ? 'bg-green-500' : pct >= 40 ? 'bg-yellow-500' : 'bg-red-500'
                    }`}
                    style={{ width: `${pct}%` }}
                  />
                </div>
                <p className="text-xs text-slate-500 mt-2">{Math.round(pct)}%達成</p>
              </div>
            );
          })}
        </div>

        {/* シミュレーター */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
              <Target className="w-5 h-5 text-blue-700" />
              改善シミュレーター
            </h2>
            <div className="flex gap-2">
              <button 
                onClick={() => setSimMode(!simMode)}
                className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-colors ${
                  simMode ? 'bg-blue-700 text-white' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                }`}
              >
                {simMode ? 'ON' : 'OFF'}
              </button>
              {simMode && (
                <button 
                  onClick={resetSim}
                  className="px-3 py-1.5 rounded-lg text-sm font-medium bg-slate-100 text-slate-700 hover:bg-slate-200 flex items-center gap-1"
                >
                  <RefreshCw className="w-3 h-3" />
                  リセット
                </button>
              )}
            </div>
          </div>
          
          {simMode ? (
            <div className="space-y-4">
              <p className="text-sm text-slate-600 bg-blue-50 p-3 rounded-lg flex items-start gap-2">
                <Info className="w-4 h-4 text-blue-700 mt-0.5 flex-shrink-0" />
                <span>各項目を増減させると、上のスコア・格付けがリアルタイムで変動します</span>
              </p>
              {[
                { key: 'netAssets', label: '純資産を増やす（増資・利益積上）', step: 5000, min: -50000, max: 200000 },
                { key: 'interestBearingDebt', label: '借入金を増減する', step: 5000, min: -100000, max: 50000 },
                { key: 'ordinaryProfit', label: '経常利益を増やす', step: 1000, min: -10000, max: 50000 },
                { key: 'currentAssets', label: '流動資産を増やす（現預金等）', step: 5000, min: -50000, max: 100000 },
                { key: 'currentLiabilities', label: '流動負債を増減する', step: 5000, min: -50000, max: 50000 },
              ].map(({ key, label, step, min, max }) => (
                <div key={key}>
                  <div className="flex justify-between mb-1">
                    <label className="text-sm font-medium text-slate-700">{label}</label>
                    <span className={`text-sm font-semibold ${
                      simAdjust[key] > 0 ? 'text-green-600' : simAdjust[key] < 0 ? 'text-red-600' : 'text-slate-500'
                    }`}>
                      {simAdjust[key] >= 0 ? '+' : ''}{simAdjust[key].toLocaleString()}千円
                    </span>
                  </div>
                  <input 
                    type="range"
                    min={min}
                    max={max}
                    step={step}
                    value={simAdjust[key]}
                    onChange={(e) => setSimAdjust(prev => ({ ...prev, [key]: Number(e.target.value) }))}
                    className="w-full accent-blue-700"
                  />
                </div>
              ))}
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              ONにすると、財務数値を仮想的に動かして格付けへの影響をシミュレーションできます。
            </p>
          )}
        </div>

        {/* 危険信号 */}
        {dangerSignals.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border-2 border-red-300 p-6 mb-6">
            <h2 className="text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-red-600" />
              検出された危険信号
              <span className="text-xs font-normal text-red-700 bg-red-50 px-2 py-0.5 rounded">{dangerSignals.length}件</span>
            </h2>
            <p className="text-sm text-slate-600 mb-4">財務スコアリングの結果から、特に注意すべき危険信号を抽出しました</p>
            <div className="space-y-3">
              {dangerSignals.map((signal, idx) => {
                const article = REFERENCE_ARTICLES[signal.article];
                return (
                  <div key={idx} className={`rounded-lg p-4 border ${
                    signal.type === 'critical' ? 'bg-red-50 border-red-300' : 'bg-orange-50 border-orange-300'
                  }`}>
                    <div className="flex items-start gap-3">
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${
                        signal.type === 'critical' ? 'bg-red-600' : 'bg-orange-500'
                      }`}>
                        <AlertCircle className="w-5 h-5 text-white" />
                      </div>
                      <div className="flex-1">
                        <p className={`font-semibold mb-1 ${
                          signal.type === 'critical' ? 'text-red-900' : 'text-orange-900'
                        }`}>{signal.title}</p>
                        <p className="text-sm text-slate-700 mb-2">{signal.desc}</p>
                        {article && (
                          <a href={article.url} target="_blank" rel="noopener noreferrer"
                            className="inline-flex items-center gap-2 text-xs bg-white hover:bg-slate-50 px-3 py-1.5 rounded border border-slate-200 transition-colors">
                            <Lightbulb className="w-3 h-3 text-amber-600" />
                            <span className="font-semibold text-slate-700">詳しく読む:</span>
                            <span className="text-slate-600">{article.title}</span>
                            <ExternalLink className="w-3 h-3 text-slate-400" />
                          </a>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* 改善優先度 */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-6">
          <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
            <TrendingUp className="w-5 h-5 text-blue-700" />
            優先改善ポイント TOP5
          </h2>
          <p className="text-sm text-slate-600 mb-4">スコアの低い指標から優先的に取り組むことで、効率的に格付けが上がります</p>
          <div className="space-y-3">
            {improvementImpact.map((item, idx) => (
              <div key={item.key} className="border border-slate-200 rounded-lg p-4 hover:border-blue-300 transition-colors">
                <div className="flex items-start gap-3">
                  <div className={`w-8 h-8 rounded-full flex items-center justify-center font-bold text-white flex-shrink-0 ${
                    idx === 0 ? 'bg-red-500' : idx === 1 ? 'bg-orange-500' : idx === 2 ? 'bg-yellow-500' : 'bg-slate-400'
                  }`}>
                    {idx + 1}
                  </div>
                  <div className="flex-1">
                    <div className="flex flex-wrap items-center gap-2 mb-1">
                      <span className="font-semibold text-slate-800">{item.name}</span>
                      <span className="text-xs px-2 py-0.5 bg-slate-100 text-slate-600 rounded">{item.cat}</span>
                      <span className="text-xs px-2 py-0.5 bg-red-50 text-red-700 rounded">スコア {item.score}/10</span>
                    </div>
                    <p className="text-sm text-slate-600 mb-2">
                      現在値: <span className="font-mono font-semibold">{formatNumber(item.current)}{item.unit}</span>
                    </p>
                    {formulaMap[item.key] && (
                      <p className="text-xs text-slate-600 bg-blue-50 p-2 rounded mb-2 font-mono whitespace-pre-line">
                        📐 {formulaMap[item.key].formula}
                      </p>
                    )}
                    <p className="text-xs text-slate-600 bg-slate-50 p-2 rounded mb-2">
                      {actionMap[item.key]}
                    </p>
                    {articleMap[item.key] && articleMap[item.key].length > 0 && (
                      <div className="space-y-1">
                        {articleMap[item.key].slice(0, 1).map(articleKey => {
                          const article = REFERENCE_ARTICLES[articleKey];
                          if (!article) return null;
                          return (
                            <a key={articleKey} href={article.url} target="_blank" rel="noopener noreferrer"
                              className="flex items-center gap-2 text-xs text-amber-800 bg-amber-50 hover:bg-amber-100 p-2 rounded transition-colors">
                              <Lightbulb className="w-3 h-3 flex-shrink-0" />
                              <span className="font-semibold">推奨記事:</span>
                              <span className="flex-1 truncate">{article.title}</span>
                              <ExternalLink className="w-3 h-3 flex-shrink-0" />
                            </a>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* 指標一覧 */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-6">
          <h2 className="text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
            <BarChart3 className="w-5 h-5 text-blue-700" />
            全指標スコアカード
          </h2>
          <p className="text-sm text-slate-500 mb-4">各行をクリックすると計算式と数値の内訳が確認できます</p>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-200">
                  <th className="text-left py-2 px-2 w-8"></th>
                  <th className="text-left py-2 px-2 font-semibold text-slate-700">カテゴリ</th>
                  <th className="text-left py-2 px-2 font-semibold text-slate-700">指標</th>
                  <th className="text-right py-2 px-2 font-semibold text-slate-700">現在値</th>
                  <th className="text-right py-2 px-2 font-semibold text-slate-700">目標水準</th>
                  <th className="text-right py-2 px-2 font-semibold text-slate-700">スコア</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { key: 'equityRatio', cat: '安全性', name: '自己資本比率', value: metrics.equityRatio, target: '25%以上', unit: '%', score: scores.equityRatio },
                  { key: 'gearingRatio', cat: '安全性', name: 'ギアリング比率', value: metrics.gearingRatio, target: '150%以内', unit: '%', score: scores.gearingRatio },
                  { key: 'fixedLongTermRatio', cat: '安全性', name: '固定長期適合率', value: metrics.fixedLongTermRatio, target: '60%以内', unit: '%', score: scores.fixedLongTermRatio },
                  { key: 'currentRatio', cat: '安全性', name: '流動比率', value: metrics.currentRatio, target: '140%以上', unit: '%', score: scores.currentRatio },
                  { key: 'salesOrdinaryProfitRatio', cat: '収益性', name: '売上高経常利益率', value: metrics.salesOrdinaryProfitRatio, target: '2%以上', unit: '%', score: scores.salesOrdinaryProfitRatio },
                  { key: 'totalCapitalOrdinaryProfitRatio', cat: '収益性', name: '総資本経常利益率(ROA)', value: metrics.totalCapitalOrdinaryProfitRatio, target: '1%以上', unit: '%', score: scores.totalCapitalOrdinaryProfitRatio },
                  { key: 'consecutiveBlackYears', cat: '収益性', name: '連続黒字年数', value: metrics.consecutiveBlackYears, target: '2期以上', unit: '年', score: scores.consecutiveBlackYears },
                  { key: 'ordinaryProfitGrowthRate', cat: '成長性', name: '経常利益増加率', value: metrics.ordinaryProfitGrowthRate, target: '15%以上', unit: '%', score: scores.ordinaryProfitGrowthRate },
                  { key: 'equityAmount', cat: '成長性', name: '自己資本額', value: metrics.equityAmount/1000, target: '100百万以上', unit: '百万円', score: scores.equityAmount },
                  { key: 'salesAmount', cat: '成長性', name: '売上高', value: metrics.salesAmount/1000, target: '100百万以上', unit: '百万円', score: scores.salesAmount },
                  { key: 'debtRepaymentYears', cat: '返済能力', name: '債務償還年数', value: metrics.debtRepaymentYears, target: '7年以内', unit: '年', score: scores.debtRepaymentYears },
                  { key: 'interestCoverageRatio', cat: '返済能力', name: 'インタレストカバレッジレシオ', value: metrics.interestCoverageRatio, target: '2.5倍以上', unit: '倍', score: scores.interestCoverageRatio },
                  { key: 'cashFlowAmount', cat: '返済能力', name: 'キャッシュフロー額', value: metrics.cashFlowAmount/1000, target: '100百万以上', unit: '百万円', score: scores.cashFlowAmount },
                ].map((row, idx) => {
                  const isExpanded = expandedMetric === row.key;
                  const formula = formulaMap[row.key];
                  return (
                    <React.Fragment key={idx}>
                      <tr 
                        className="border-b border-slate-100 hover:bg-slate-50 cursor-pointer transition-colors"
                        onClick={() => setExpandedMetric(isExpanded ? null : row.key)}
                      >
                        <td className="py-2 px-2">
                          {isExpanded ? <ChevronUp className="w-4 h-4 text-slate-400" /> : <ChevronDown className="w-4 h-4 text-slate-400" />}
                        </td>
                        <td className="py-2 px-2">
                          <span className={`text-xs px-2 py-0.5 rounded ${
                            row.cat === '安全性' ? 'bg-blue-50 text-blue-700' :
                            row.cat === '収益性' ? 'bg-green-50 text-green-700' :
                            row.cat === '成長性' ? 'bg-purple-50 text-purple-700' :
                            'bg-orange-50 text-orange-700'
                          }`}>{row.cat}</span>
                        </td>
                        <td className="py-2 px-2 text-slate-700 font-medium">{row.name}</td>
                        <td className="py-2 px-2 text-right font-mono">{formatNumber(row.value)}{row.unit}</td>
                        <td className="py-2 px-2 text-right text-slate-500">{row.target}</td>
                        <td className="py-2 px-2 text-right">
                          <span className={`inline-block px-2 py-0.5 rounded font-semibold text-xs ${
                            row.score >= 7 ? 'bg-green-100 text-green-800' :
                            row.score >= 4 ? 'bg-yellow-100 text-yellow-800' :
                            'bg-red-100 text-red-800'
                          }`}>
                            {row.score}/10
                          </span>
                        </td>
                      </tr>
                      {isExpanded && formula && (
                        <tr className="bg-slate-50">
                          <td colSpan={6} className="py-4 px-4">
                            <div className="space-y-3">
                              <div className="bg-white rounded-lg p-3 border border-slate-200">
                                <div className="flex items-center gap-2 mb-2">
                                  <Calculator className="w-4 h-4 text-blue-700" />
                                  <span className="text-xs font-semibold text-slate-700">計算式</span>
                                </div>
                                <p className="text-sm font-mono text-slate-800 whitespace-pre-line bg-blue-50 p-2 rounded">{formula.formula}</p>
                              </div>
                              <div className="bg-white rounded-lg p-3 border border-slate-200">
                                <p className="text-xs font-semibold text-slate-700 mb-2">数値の内訳</p>
                                {formula.denominator ? (
                                  <div className="text-sm">
                                    <div className="flex items-center gap-2 mb-1">
                                      <span className="text-slate-600 w-32">{formula.numerator.label}:</span>
                                      <span className="font-mono font-semibold text-slate-800">{formula.numerator.value.toLocaleString()} 千円</span>
                                    </div>
                                    <div className="flex items-center gap-2 mb-2">
                                      <span className="text-slate-600 w-32">{formula.denominator.label}:</span>
                                      <span className="font-mono font-semibold text-slate-800">{formula.denominator.value.toLocaleString()} 千円</span>
                                    </div>
                                    <div className="border-t border-slate-200 pt-2 mt-2">
                                      <span className="text-slate-600">計算結果: </span>
                                      <span className="font-mono font-bold text-blue-700">
                                        {formula.numerator.value.toLocaleString()} ÷ {formula.denominator.value.toLocaleString()}
                                        {row.unit === '%' && ' × 100'}
                                        {' = '}
                                        {formatNumber(row.value)}{row.unit}
                                      </span>
                                    </div>
                                  </div>
                                ) : (
                                  <div className="text-sm">
                                    <span className="text-slate-600 w-32 mr-2">{formula.numerator.label}:</span>
                                    <span className="font-mono font-bold text-blue-700">
                                      {formula.numerator.value.toLocaleString()} {row.unit === '年' ? '年' : '千円'}
                                    </span>
                                  </div>
                                )}
                              </div>
                              <div className="bg-white rounded-lg p-3 border border-slate-200">
                                <p className="text-xs font-semibold text-slate-700 mb-1">この指標の意味</p>
                                <p className="text-sm text-slate-600">{formula.meaning}</p>
                              </div>
                              {formula.source && (
                                <div className="bg-blue-50 rounded-lg p-3 border border-blue-200">
                                  <div className="flex items-center gap-2 mb-1">
                                    <Shield className="w-4 h-4 text-blue-700" />
                                    <p className="text-xs font-semibold text-blue-900">公的基準・参照元</p>
                                  </div>
                                  <p className="text-xs text-blue-800">{formula.source}</p>
                                </div>
                              )}
                              {articleMap[row.key] && articleMap[row.key].length > 0 && (
                                <div className="bg-amber-50 rounded-lg p-3 border border-amber-200">
                                  <div className="flex items-center gap-2 mb-2">
                                    <Lightbulb className="w-4 h-4 text-amber-700" />
                                    <p className="text-xs font-semibold text-amber-900">この指標を深く理解する記事</p>
                                  </div>
                                  <div className="space-y-2">
                                    {articleMap[row.key].slice(0, 2).map(articleKey => {
                                      const article = REFERENCE_ARTICLES[articleKey];
                                      if (!article) return null;
                                      return (
                                        <a key={articleKey} href={article.url} target="_blank" rel="noopener noreferrer"
                                          className="block bg-white rounded p-2 border border-amber-100 hover:border-amber-400 hover:shadow-sm transition-all group">
                                          <div className="flex items-start gap-2">
                                            <span className={`text-xs font-bold px-1.5 py-0.5 rounded flex-shrink-0 ${
                                              article.rank === 'S' ? 'bg-red-600 text-white' :
                                              article.rank === 'A' ? 'bg-orange-500 text-white' :
                                              'bg-slate-400 text-white'
                                            }`}>{article.rank}</span>
                                            <div className="flex-1 min-w-0">
                                              <p className="text-xs font-semibold text-slate-800 group-hover:text-amber-900 flex items-center gap-1">
                                                {article.title}
                                                <ExternalLink className="w-3 h-3 flex-shrink-0" />
                                              </p>
                                              <p className="text-xs text-slate-600 mt-0.5">{article.summary}</p>
                                            </div>
                                          </div>
                                        </a>
                                      );
                                    })}
                                  </div>
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* 公的基準・参考文献 */}
        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 mb-6">
          <h2 className="text-lg font-semibold text-slate-800 mb-4 flex items-center gap-2">
            <BookOpen className="w-5 h-5 text-blue-700" />
            本診断が準拠する公的基準・参考文献
          </h2>
          <div className="space-y-3">
            <div className="border-l-4 border-blue-700 bg-blue-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-blue-900">📜 金融庁「金融検査マニュアル別冊〔中小企業融資編〕」</p>
                <a href="https://www.fsa.go.jp/manual/manualj/yokin_b.html" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-blue-700 hover:underline flex items-center gap-1">
                  原文を見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">
                銀行が融資先を区分する5段階の債務者区分（正常先・要注意先・要管理先・破綻懸念先・破綻先）の基準。本ツールの債務者区分判定は、本マニュアルの考え方に基づいています。
                また、「中小・零細企業等の場合、赤字・債務超過が直ちに要注意先以下とならない」「役員借入金は自己資本相当とみなすことができる」など、中小企業特有の評価ルールも実態BS変換ロジックに反映しています。
              </p>
            </div>
            <div className="border-l-4 border-green-700 bg-green-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-green-900">🏛️ 中小機構「経営自己診断システム」</p>
                <a href="https://k-sindan.smrj.go.jp/" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-green-700 hover:underline flex items-center gap-1">
                  公式サイトを見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">
                独立行政法人中小企業基盤整備機構が、経済産業省・中小企業庁主導で開発されたCRD（中小企業信用リスク情報データベース、約200万社の財務データ）を活用して提供する無料診断システム。
                収益性・効率性・生産性・安全性・成長性の5項目27指標で構成され、本ツールの13指標体系の主要な参照元です。
                安全性の10指標（自己資本比率、流動比率、当座比率、固定長期適合率、減価償却率、手元現金預金比率、借入金月商倍率、借入金依存度、預借率、売上高支払利息割引料率）は本ツールに反映されています。
              </p>
            </div>
            <div className="border-l-4 border-orange-600 bg-orange-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-orange-900">🏢 経済産業省「ローカルベンチマーク（ロカベン）」</p>
                <a href="https://www.meti.go.jp/policy/economy/keiei_innovation/sangyokinyu/locaben/" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-orange-700 hover:underline flex items-center gap-1">
                  公式サイトを見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">
                経済産業省が公表する「企業の健康診断」ツール。各種補助金等の申請にも活用される公式評価フレームワーク。
                6つの財務指標（売上高増加率・営業利益率・労働生産性・EBITDA有利子負債倍率・営業運転資本回転期間・自己資本比率）と非財務分析で構成され、本ツールが補完的に参照しています。
              </p>
            </div>
            <div className="border-l-4 border-indigo-600 bg-indigo-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-indigo-900">📊 中小企業庁「中小企業実態基本調査」（業界比較データの源泉）</p>
                <a href="https://www.chusho.meti.go.jp/koukai/chousa/kihon/index.html" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-indigo-700 hover:underline flex items-center gap-1">
                  公式サイトを見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">
                中小企業基本法第10条の規定に基づき、中小企業庁が毎年実施している公式統計調査。全国約11万社を対象に、産業別・従業者規模別の財務指標（自己資本比率、売上高経常利益率、流動比率、ROA等）を集計・公表。
                本ツールの「業界内ポジション」「業界平均との比較」セクションは、本調査の令和6年確報（令和5年度決算実績）の数値を参照しています。
                ※完全無料・統計法に基づく公式データ。
              </p>
            </div>
            <div className="border-l-4 border-teal-600 bg-teal-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-teal-900">🏦 日本政策金融公庫「小企業の経営指標」</p>
                <a href="https://www.jfc.go.jp/n/findings/shihyou_kekka_m_index.html" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-teal-700 hover:underline flex items-center gap-1">
                  公式サイトを見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">
                日本政策金融公庫が小企業の決算データを業種ごとに隔年で集計している公式統計。製造業・卸売業・小売業・サービス業を細分化した業種別（数百業種）に、平均値・上方/下方信頼限界・標準偏差を提供しており、本ツールの偏差値計算（業界平均からのばらつき）の補完に使用しています。
              </p>
            </div>
            <div className="border-l-4 border-slate-400 bg-slate-50 p-4 rounded-r-lg">
              <p className="text-sm font-semibold text-slate-800 mb-1">📚 民間スコアリング配点ロジック</p>
              <p className="text-xs text-slate-700 leading-relaxed">
                各指標の10点満点での配点（自己資本比率25%以上で5点、債務償還年数7年以内で5点 など）と、合計スコアによる格付け1〜10判定は、複数の税理士法人・コンサルティングファームが公開する民間金融機関の財務スコアリングモデルの基準値を参考にしています。
                CRDモデルそのもののロジックは機密情報のため、公開情報による近似値である点にご留意ください。
              </p>
            </div>
            <div className="border-l-4 border-amber-600 bg-amber-50 p-4 rounded-r-lg">
              <div className="flex items-start justify-between mb-1">
                <p className="text-sm font-semibold text-amber-900">💡 実務専門家による解説（安田経営診断事務所｜安田順氏）</p>
                <a href="https://yasuda-keiei.com/" target="_blank" rel="noopener noreferrer"
                  className="text-xs text-amber-700 hover:underline flex items-center gap-1">
                  公式サイトを見る <ExternalLink className="w-3 h-3" />
                </a>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed mb-2">
                経済産業大臣登録の中小企業診断士、認定経営革新等支援機関である安田順氏（元金融機関融資担当）が公開する財務分析・銀行交渉の実務解説記事を、各指標のコメント・危険信号判定の根拠として参照しています。
                MCSSの寄与率分析、CRDランクごとの実証データ、過剰債務脱出に必要な営業利益率水準など、実務に裏打ちされた具体的な数値基準を採用しています。
              </p>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-2 mt-3">
                <div className="bg-white rounded p-2 border border-amber-200">
                  <p className="text-xs font-bold text-red-700 mb-1">⭐ Sランク（スコア計算根拠）</p>
                  <p className="text-xs text-slate-600">3記事：MCSSの寄与率検証、黒字でも格付け下がる会社、経営自己診断システムとMcSSの違い</p>
                </div>
                <div className="bg-white rounded p-2 border border-amber-200">
                  <p className="text-xs font-bold text-orange-600 mb-1">⭐ Aランク（指標コメント根拠）</p>
                  <p className="text-xs text-slate-600">5記事：債務償還年数、支払利息率、CF計算書、流動比率の限界、過剰債務と営業利益率5%</p>
                </div>
                <div className="bg-white rounded p-2 border border-amber-200">
                  <p className="text-xs font-bold text-slate-600 mb-1">⭐ Bランク（改善提案根拠）</p>
                  <p className="text-xs text-slate-600">9記事：節税、試算表、決算説明資料、融資依存、貸し込まれ、ファクタリング等</p>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-slate-100 rounded-xl p-4 text-xs text-slate-600 mb-6">
          <p className="flex items-start gap-2">
            <AlertCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />
            <span>
              <strong>免責事項：</strong>本ツールは上記の公的機関が公開する基準と、民間専門家が公開する財務スコアリングロジックを基に構築した目安です。実際の格付けは各金融機関の独自モデル・定性評価・実態評価を加味して決定されるため、本診断結果と完全一致するものではありません。金融機関との具体的な交渉にあたっては、税理士・中小企業診断士等の専門家への相談を推奨します。
            </span>
          </p>
        </div>
      </div>
    </div>
  );
}
