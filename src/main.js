// ======== 設定定数 ========
/**
 * 最適化アルゴリズムの設定パラメータ
 */
const CONFIG = {
  // 粗配分設定
  BATCH_SIZE: 10,           // 1回あたりのポイント配分数
  TOP_VARS: 3,              // 同時に考慮する上位変数の数

  // リバランス設定
  MAX_ITERATIONS: 30,       // 最大反復回数
  MAX_CANDIDATES: 10,       // 評価する移動候補の最大数
  THRESHOLD: 0.00001        // 改善と判定する最小ダメージ増加量
};

// ======== ヘルパー関数 ========

/**
 * セル範囲に値を書き込み、再計算を実行する
 * @param {GoogleAppsScript.Spreadsheet.Range} range - 書き込み先のセル範囲
 * @param {number[]} values - 書き込む値の配列
 * @returns {void}
 */
function updateSheet(range, values) {
  range.setValues(values.map(v => [v]));
  SpreadsheetApp.flush();
}

/**
 * 配列から条件に一致するインデックスを抽出する
 * @param {number[]} arr - 対象配列
 * @param {function(number): boolean} predicate - 条件関数
 * @returns {number[]} 条件に一致したインデックスの配列
 */
function filterIndices(arr, predicate) {
  return arr.reduce((indices, val, i) => {
    if (predicate(val)) indices.push(i);
    return indices;
  }, []);
}

// ======== UI ========

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎮 最適化ツール')
    .addItem('🚀 サブステ最適化計算', 'optimizeSubStats')
    .addSeparator()
    .addItem('⚙️ 設定変更', 'configureSettings')
    .addToUi();
}

/**
 * サブステータス最適化を実行
 */
function optimizeSubStats() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const varRange = props.getProperty('varRange');
  const calcCell = props.getProperty('calcCell');

  if (!varRange || !calcCell) {
    ui.alert('エラー', 'セル位置が設定されていません。\n先に「⚙️ 設定変更」を実行してください。', ui.ButtonSet.OK);
    return;
  }

  const totalPoints = promptForTotalPoints(ui);
  if (totalPoints === null) return;

  const startTime = Date.now();
  const result = runOptimization(varRange, calcCell, totalPoints);
  const executionTime = (Date.now() - startTime) / 1000;

  showResultDialog(ui, result, executionTime);
}

/**
 * 総ポイント数の入力を求める
 * @param {GoogleAppsScript.Base.Ui} ui - スプレッドシートのUIインスタンス
 * @returns {number|null} 入力されたポイント数、キャンセル時はnull
 */
function promptForTotalPoints(ui) {
  const response = ui.prompt(
    '総サブステヒット数を指定',
    '配分する総サブステヒット数を入力してください\n例: 40',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const points = parseInt(response.getResponseText());
  if (isNaN(points) || points <= 0) {
    ui.alert('エラー', '有効な数値を入力してください', ui.ButtonSet.OK);
    return null;
  }

  return points;
}

/**
 * 最適化結果をダイアログで表示
 * @param {GoogleAppsScript.Base.Ui} ui - スプレッドシートのUIインスタンス
 * @param {OptimizationResult} result - 最適化結果
 * @param {number} executionTime - 実行時間（秒）
 * @returns {void}
 */
function showResultDialog(ui, result, executionTime) {
  const increaseRate = ((result.final / result.initial - 1) * 100).toFixed(2);
  ui.alert(
    '最適化完了✅',
    `実行時間: ${executionTime.toFixed(1)}秒\n` +
    `計算回数: ${result.calcCount}回\n` +
    `初期ダメージ: ${result.initial.toFixed(2)}\n` +
    `粗配分後: ${result.rough.toFixed(2)}\n` +
    `最終ダメージ: ${result.final.toFixed(2)}\n` +
    `増加率: +${increaseRate}%\n` +
    `リバランス改善: ${result.improvements}回`,
    ui.ButtonSet.OK
  );
}

/**
 * セル範囲の設定ダイアログを表示
 */
function configureSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const varRange = promptForSetting(ui, props, 'varRange', 'サブステヒットセルの範囲を指定',
    '最適化したいサブステヒット数のセル範囲を入力してください\n例: B2:B11');
  if (varRange === null) return;

  const calcCell = promptForSetting(ui, props, 'calcCell', 'ダメージセルを指定',
    'ダメージ計算結果が表示されるセルを入力してください\n例: D2');
  if (calcCell === null) return;

  props.setProperties({ varRange, calcCell });
  ui.alert('設定完了✅', `変数範囲: ${varRange}\n計算セル: ${calcCell}`, ui.ButtonSet.OK);
}

/**
 * 設定値の入力を求める
 * @param {GoogleAppsScript.Base.Ui} ui - スプレッドシートのUIインスタンス
 * @param {GoogleAppsScript.Properties.Properties} props - ドキュメントプロパティ
 * @param {string} key - プロパティのキー名
 * @param {string} title - ダイアログのタイトル
 * @param {string} message - ダイアログのメッセージ
 * @returns {string|null} 入力値、キャンセル時はnull
 */
function promptForSetting(ui, props, key, title, message) {
  const current = props.getProperty(key) || 'なし';
  const response = ui.prompt(title, `${message}\n\n現在の設定: ${current}`, ui.ButtonSet.OK_CANCEL);
  return response.getSelectedButton() === ui.Button.OK ? response.getResponseText() : null;
}

// ======== コアロジック ========

/**
 * @typedef {Object} State
 * @property {number[]} values - 各サブステータスの配分ポイント
 * @property {number} calcCount - ダメージ計算の実行回数
 */

/**
 * @typedef {Object} OptimizationResult
 * @property {number} calcCount - 計算実行回数
 * @property {number} initial - 初期ダメージ
 * @property {number} rough - 粗配分後ダメージ
 * @property {number} final - 最終ダメージ
 * @property {number} improvements - リバランス改善回数
 */

/**
 * サブステータスの最適化を実行
 * @param {string} varRangeStr - 変数セル範囲（例: "B2:B11"）
 * @param {string} calcCellStr - ダメージ計算セル（例: "D2"）
 * @param {number} totalPoints - 配分する総ポイント数
 * @returns {OptimizationResult}
 */
function runOptimization(varRangeStr, calcCellStr, totalPoints) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const varRange = sheet.getRange(varRangeStr);
  const calcCell = sheet.getRange(calcCellStr);
  const numVars = varRange.getNumRows();

  let state = { values: new Array(numVars).fill(0), calcCount: 0 };

  // 初期化
  updateSheet(varRange, state.values);
  const initialDamage = calcCell.getValue();
  if (initialDamage <= 0) throw new Error('初期ダメージが0以下です');

  // Phase 1: 粗配分
  const roughResult = allocateByUtility(state, totalPoints, varRange, calcCell);
  state = roughResult.state;
  const roughDamage = calcCell.getValue();

  // Phase 2: リバランス
  const rebalanceResult = rebalance(state, roughResult.utilities, varRange, calcCell);

  return {
    calcCount: rebalanceResult.state.calcCount,
    initial: initialDamage,
    rough: roughDamage,
    final: calcCell.getValue(),
    improvements: rebalanceResult.improvements
  };
}

/**
 * 効用に基づきポイントを配分（Phase 1: 貪欲法）
 * 
 * アルゴリズム:
 * 1. BATCH_SIZE（デフォルト10）ポイントずつ配分を繰り返す
 * 2. 各バッチで:
 *    a. 全変数の効用（1ポイント追加時のダメージ増加量）を測定
 *    b. 効用が高い上位TOP_VARS個の変数を選択
 *    c. 効用比に応じてバッチ内のポイントを按分
 *       例: 会心率の効用10, 会心ダメの効用5, 攻撃力の効用5
 *           → 10ポイントを 5:2.5:2.5 の比率で配分（四捨五入）
 * 3. スプレッドシートに書き込み、次のバッチへ
 * 
 * バッチ処理の利点:
 *   - 1ポイントずつ配分するより効率的（I/O回数削減）
 *   - 配分が進むにつれ効用が変化することに対応
 *     （例: 会心率が高くなると会心ダメの効用が上がる）
 * 
 * 限界:
 *   - 貪欲法なので局所最適解に陥る可能性
 *   - 初期に効用が低い変数は完全に無視される
 *   → これらはPhase 2のrebalanceで補完
 * 
 * @param {State} state - 現在の状態
 * @param {number} totalPoints - 配分する総ポイント数
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {{state: State, utilities: number[]}} 更新された状態と効用配列
 */
function allocateByUtility(state, totalPoints, varRange, calcCell) {
  let currentState = { ...state, values: [...state.values] };
  let latestUtilities = new Array(state.values.length).fill(0);
  let allocated = 0;

  while (allocated < totalPoints) {
    const batch = Math.min(CONFIG.BATCH_SIZE, totalPoints - allocated);

    const measureResult = measureUtilities(currentState, varRange, calcCell);
    currentState = measureResult.state;
    latestUtilities = measureResult.utilities;

    const topVars = selectTopVars(latestUtilities);
    const newValues = distributePoints(currentState.values, topVars, batch);

    currentState = { ...currentState, values: newValues };
    allocated += batch;

    updateSheet(varRange, newValues);
  }

  return { state: currentState, utilities: latestUtilities };
}

/**
 * 効用が高い上位変数を取得
 * @param {number[]} utilities - 各変数の効用値配列
 * @returns {{index: number, utility: number}[]} 上位変数の配列（インデックスと効用値）
 */
function selectTopVars(utilities) {
  return utilities
    .map((u, i) => ({ index: i, utility: u }))
    .filter(item => item.utility > 0)
    .sort((a, b) => b.utility - a.utility)
    .slice(0, CONFIG.TOP_VARS);
}

/**
 * 上位変数にポイントを分配
 * @param {number[]} currentValues - 現在の配分値
 * @param {{index: number, utility: number}[]} topVars - 上位変数の配列
 * @param {number} batch - 配分するポイント数
 * @returns {number[]} 配分後の値の配列
 */
function distributePoints(currentValues, topVars, batch) {
  const newValues = [...currentValues];

  if (topVars.length === 0) {
    newValues[0] += batch;
    return newValues;
  }

  const totalUtil = topVars.reduce((sum, v) => sum + v.utility, 0);
  let remaining = batch;

  for (const item of topVars) {
    const points = Math.min(Math.round(batch * item.utility / totalUtil), remaining);
    newValues[item.index] += points;
    remaining -= points;
  }

  // 端数は最高効用の変数へ
  if (remaining > 0) {
    newValues[topVars[0].index] += remaining;
  }

  return newValues;
}

/**
 * 各変数の効用（1ポイント追加時のダメージ増加量）を計測
 * 
 * プロセス:
 * 1. 現在のダメージ値を基準点として記録
 * 2. 各変数について順番に:
 *    - 現在値+1をシートに書き込む
 *    - ダメージセルを再計算
 *    - 増加量を効用として記録
 * 3. 【重要】すべての計測後、必ず元の値に復元
 * 
 * なぜ復元が必要？
 *   この関数は「もし+1したら」を測定する仮想的な操作
 *   実際の配分は呼び出し側（allocateByUtility）が決定する
 *   → 測定のための変更を残すと、意図しない状態で次の処理が始まる
 * 
 * 計算コスト:
 *   変数N個の場合、N回のスプレッドシート評価が必要
 *   → このコストが全体のボトルネック
 * 
 * @param {State} state - 現在の状態
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {{state: State, utilities: number[]}} 更新された状態と効用配列
 */
function measureUtilities(state, varRange, calcCell) {
  const baseDamage = calcCell.getValue();
  const utilities = [];
  let calcCount = state.calcCount;

  for (let i = 0; i < state.values.length; i++) {
    const testValues = [...state.values];
    testValues[i]++;

    updateSheet(varRange, testValues);
    utilities[i] = Math.max(0, calcCell.getValue() - baseDamage);
    calcCount++;
  }

  // 元の状態に復元
  updateSheet(varRange, state.values);

  return {
    state: { ...state, calcCount },
    utilities
  };
}

/**
 * 配分済みポイントの局所的な再配分
 * 
 * 粗配分後の解を2段階で改善する:
 * 
 * 【Step 1: optimizeBySwap】
 *   既にポイントが割り当てられている変数間で1ポイントを移動させて改善を探す
 *   例: 会心率3 → 会心率2, 会心ダメ5 → 会心ダメ6
 * 
 * 【Step 2: tryZeroVars】
 *   0割当の変数が実は有効ではないかを再評価
 *   例: 攻撃力10, 元素熟知0 → 攻撃力9, 元素熟知1
 *   
 *   なぜ必要？
 *   - 粗配分時は各変数を独立に評価するため、初期効用が低い変数を見落とす
 *   - しかし他のステが揃った後では有効になるケースがある（閾値効果など）
 * 
 * 【実行順序の理由】
 *   先にlocalで既存配分を最適化してから、0割当変数を試す
 *   → 0割当の評価時点で、既に最適化された状態からの改善を測定できる
 * 
 * 注意: 
 *   - currentStateは参照渡しで各関数内で直接変更される
 *   - utilitiesは粗配分時の効用値を保持するが、変更されない（参照のみ）
 * 
 * @param {State} state - 現在の状態
 * @param {number[]} initialUtilities - 粗配分時に計測された効用値
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {{state: State, improvements: number}} 更新された状態と改善回数
 */
function rebalance(state, initialUtilities, varRange, calcCell) {
  let currentState = { ...state, values: [...state.values] };
  let utilities = [...initialUtilities];
  let improvements = 0;

  improvements += optimizeBySwap(currentState, utilities, varRange, calcCell);
  improvements += tryZeroVars(currentState, utilities, varRange, calcCell);

  return { state: currentState, improvements };
}

/**
 * 割当済み変数間でポイントを交換して改善を探索
 * 
 * アルゴリズム:
 * 1. ポイントが割り当てられている変数のペアをすべて列挙
 * 2. 各ペアについて「from → to」の移動候補を生成
 * 3. 効用差(utilities[to] - utilities[from])でソート
 *    → 効用が低い変数から高い変数へ移動する候補が優先される
 * 4. 上位MAX_CANDIDATES個を実際に試す
 *    → 全候補を試すと計算コストが高いため、有望な候補のみ評価
 * 5. 改善があれば適用し、次の反復へ
 *    → baselineDamageが更新されるため、同じ移動は再び改善しない
 * 6. 改善がなくなるまで反復（最大MAX_ITERATIONS回）
 * 
 * 注意: currentStateは直接変更される
 * 
 * @param {State} currentState - 現在の状態（valuesは直接変更される）
 * @param {number[]} utilities - 各変数の効用値配列（参照のみ、変更されない）
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {number} 改善が見つかった回数
 */
function optimizeBySwap(currentState, utilities, varRange, calcCell) {
  let improvements = 0;

  for (let iteration = 0; iteration < CONFIG.MAX_ITERATIONS; iteration++) {
    const activeVars = filterIndices(currentState.values, v => v > 0);
    if (activeVars.length <= 1) break;

    const candidates = createSwapCandidates(activeVars, utilities);
    if (candidates.length === 0 || candidates[0].priority <= 0) break;

    const improved = applyBestSwap(currentState, candidates, utilities, varRange, calcCell);
    if (!improved) break;

    improvements++;
  }

  return improvements;
}

/**
 * ポイント移動候補を生成
 * @param {number[]} activeVars - アクティブな変数のインデックス配列
 * @param {number[]} utilities - 各変数の効用値配列
 * @returns {{from: number, to: number, priority: number}[]} 移動候補の配列（優先度順にソート済み）
 */
function createSwapCandidates(activeVars, utilities) {
  const candidates = [];

  for (const from of activeVars) {
    for (const to of activeVars) {
      if (from !== to) {
        candidates.push({
          from,
          to,
          priority: utilities[to] - utilities[from]
        });
      }
    }
  }

  return candidates.sort((a, b) => b.priority - a.priority);
}

/**
 * 最良の移動を試行し、改善があれば適用する
 * @param {State} currentState - 現在の状態（valuesは直接変更される）
 * @param {{from: number, to: number, priority: number}[]} candidates - 移動候補の配列
 * @param {number[]} utilities - 各変数の効用値配列（未使用：将来の拡張用に保持）
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {boolean} 改善があればtrue
 */
function applyBestSwap(currentState, candidates, utilities, varRange, calcCell) {
  const baselineDamage = calcCell.getValue();
  const maxTries = Math.min(CONFIG.MAX_CANDIDATES, candidates.length);

  for (let i = 0; i < maxTries; i++) {
    const candidate = candidates[i];
    const testValues = [...currentState.values];
    testValues[candidate.from]--;
    testValues[candidate.to]++;

    updateSheet(varRange, testValues);

    if (calcCell.getValue() > baselineDamage + CONFIG.THRESHOLD) {
      currentState.values = testValues;
      return true;
    }
  }

  return false;
}

/**
 * 0割当の変数を試行する
 * 
 * 目的:
 *   粗配分で見落とされた変数が、実は有効ではないかを再評価する
 * 
 * なぜ必要？
 *   例: 元素熟知は単体では効用が低く見えるが、
 *       会心率・会心ダメが揃った後では突破的に効果が出るケース
 *   → 粗配分時の効用測定では捉えられない
 * 
 * アルゴリズム:
 * 1. 0割当の変数（zeroVars）をすべて取得
 * 2. 割当済み変数のうち、効用が低いもの（lowVars）を取得
 *    → なぜ効用が低い変数から削る？
 *      効用が低い = 削っても損失が少ない = スワップの成功確率が高い
 * 3. すべてのzero×lowの組み合わせでスワップを試す
 * 4. 最も改善量が大きいスワップがあれば適用
 * 
 * 計算コスト削減:
 *   全組み合わせではなく、効用が低い上位TOP_VARS個のみ評価
 *   → 変数が多い場合の計算時間を抑制
 * 
 * 注意: currentStateは直接変更される
 * 
 * @param {State} currentState - 現在の状態（valuesは直接変更される）
 * @param {number[]} utilities - 各変数の効用値配列
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {number} 改善があれば1、なければ0
 */
function tryZeroVars(currentState, utilities, varRange, calcCell) {
  const zeroVars = filterIndices(currentState.values, v => v === 0);
  const lowVars = selectLowVars(currentState.values, utilities);

  if (zeroVars.length === 0 || lowVars.length === 0) return 0;

  updateSheet(varRange, currentState.values);
  const baselineDamage = calcCell.getValue();

  const bestSwap = findBestSwap(currentState.values, zeroVars, lowVars, baselineDamage, varRange, calcCell);

  if (bestSwap && bestSwap.gain > CONFIG.THRESHOLD) {
    currentState.values[bestSwap.zero]++;
    currentState.values[bestSwap.hit]--;
    updateSheet(varRange, currentState.values);
    return 1;
  }

  updateSheet(varRange, currentState.values);
  return 0;
}

/**
 * 効用が低い割当済み変数を取得
 * @param {number[]} values - 現在の配分値
 * @param {number[]} utilities - 各変数の効用値配列
 * @returns {number[]} 効用が低い変数のインデックス配列（最大TOP_VARS個）
 */
function selectLowVars(values, utilities) {
  return values
    .map((v, i) => v > 0 ? { index: i, utility: utilities[i] } : null)
    .filter(x => x !== null)
    .sort((a, b) => a.utility - b.utility)
    .slice(0, CONFIG.TOP_VARS)
    .map(x => x.index);
}

/**
 * 最良のスワップを探索
 * @param {number[]} values - 現在の配分値
 * @param {number[]} zeroVars - 0割当変数のインデックス配列
 * @param {number[]} hitVars - 割当済み変数のインデックス配列
 * @param {number} baselineDamage - 基準となるダメージ値
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {{zero: number, hit: number, gain: number}|null} 最良のスワップ、なければnull
 */
function findBestSwap(values, zeroVars, hitVars, baselineDamage, varRange, calcCell) {
  let bestSwap = null;
  let bestGain = 0;

  for (const zero of zeroVars) {
    for (const hit of hitVars) {
      const testValues = [...values];
      testValues[zero]++;
      testValues[hit]--;

      updateSheet(varRange, testValues);
      const gain = calcCell.getValue() - baselineDamage;

      if (gain > bestGain) {
        bestGain = gain;
        bestSwap = { zero, hit, gain };
      }
    }
  }

  return bestSwap;
}