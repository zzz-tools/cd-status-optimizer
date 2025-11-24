// ======== 設定定数 ======
/**
 * 最適化アルゴリズムの設定パラメータ
 * @constant
 * @type {Object}
 * @property {number} BATCH_SIZE - 粗配分時の1回あたりのポイント配分数
 * @property {number} TOP_VARS - 粗配分時に同時に考慮する上位変数の数
 * @property {number} MAX_ITERATIONS - リバランス時の最大反復回数
 * @property {number} MAX_CANDIDATES - リバランス時に評価する移動候補の最大数
 * @property {number} THRESHOLD - 改善と判定する最小ダメージ増加量
 */
const CONFIG = {
  BATCH_SIZE: 10,
  TOP_VARS: 3,
  MAX_ITERATIONS: 30,
  MAX_CANDIDATES: 10,
  THRESHOLD: 0.00001
};

// ======== UI =========

/**
 * スプレッドシート起動時に実行され、カスタムメニューを追加する
 * @function
 * @returns {void}
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
 * サブステータス最適化を実行するUIダイアログを表示し、最適化処理を実行する
 * @function
 * @returns {void}
 * @description
 * ユーザーに総ポイント数の入力を求め、設定されたセル範囲に対して
 * 最適化を実行する。実行結果（実行時間、計算回数、ダメージ増加率など）を
 * ダイアログで表示する。
 * 
 * 事前に configureSettings() でセル範囲とダメージセルの設定が必要。
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

  const response = ui.prompt('総サブステヒット数を指定', '配分する総サブステヒット数を入力してください\n例: 40', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const totalPoints = parseInt(response.getResponseText());
  if (isNaN(totalPoints) || totalPoints <= 0) {
    ui.alert('エラー', '有効な数値を入力してください', ui.ButtonSet.OK);
    return;
  }

  const startTime = new Date();
  const result = runOptimization(varRange, calcCell, totalPoints);
  const executionTime = (new Date() - startTime) / 1000;

  ui.alert(
    '最適化完了✅',
    `実行時間: ${executionTime.toFixed(1)}秒\n` +
    `計算回数: ${result.calcCount}回\n` +
    `初期ダメージ: ${result.initial.toFixed(2)}\n` +
    `粗配分後: ${result.rough.toFixed(2)}\n` +
    `最終ダメージ: ${result.final.toFixed(2)}\n` +
    `増加率: +${((result.final / result.initial - 1) * 100).toFixed(2)}%\n` +
    `リバランス改善: ${result.improvements}回`,
    ui.ButtonSet.OK
  );
}

/**
 * 最適化対象のセル範囲とダメージ計算セルを設定するUIダイアログを表示する
 * @function
 * @returns {void}
 * @description
 * ユーザーに2つの入力を求める:
 * 1. サブステヒット数を書き込むセル範囲（例: B2:B11）
 * 2. ダメージ計算結果が表示されるセル（例: D2）
 * 
 * 設定はドキュメントプロパティに保存され、以降の最適化実行で使用される。
 */
function configureSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const varResponse = ui.prompt(
    'サブステヒットセルの範囲を指定',
    `最適化したいサブステヒット数のセル範囲を入力してください\n例: B2:B11\n\n現在の設定: ${props.getProperty('varRange') || 'なし'}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (varResponse.getSelectedButton() !== ui.Button.OK) return;

  const calcResponse = ui.prompt(
    'ダメージセルを指定',
    `ダメージ計算結果が表示されるセルを入力してください\n例: D2\n\n現在の設定: ${props.getProperty('calcCell') || 'なし'}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (calcResponse.getSelectedButton() !== ui.Button.OK) return;

  props.setProperties({
    'varRange': varResponse.getResponseText(),
    'calcCell': calcResponse.getResponseText()
  });

  ui.alert('設定完了✅', `変数範囲: ${varResponse.getResponseText()}\n計算セル: ${calcResponse.getResponseText()}`, ui.ButtonSet.OK);
}

// ======== コア =========

/**
 * 最適化実行時の内部状態のデータ構造
 * @typedef {Object} State
 * @property {number[]} values - 各サブステータスの配分ポイント
 * @property {number} calcCount - ダメージ計算セルの評価実行回数
 */

/**
 * 最適化結果の統計情報
 * @typedef {Object} OptimizationResult
 * @property {number} calcCount - ダメージ計算セルの評価実行回数
 * @property {number} initial - 初期ダメージ値
 * @property {number} rough - 粗配分後のダメージ値
 * @property {number} final - 最終的なダメージ値
 * @property {number} improvements - リバランスによる改善回数
 */

/**
 * サブステータスの最適化を実行する
 * @function
 * @param {string} varRangeStr - 最適化対象のセル範囲（例: "B2:B11"）
 * @param {string} calcCellStr - ダメージ計算セルのアドレス（例: "D2"）
 * @param {number} totalPoints - 配分する総ポイント数
 * @returns {OptimizationResult} 最適化結果の統計情報
 * @throws {Error} 初期ダメージが0以下の場合
 * @description
 * 指定された変数セル範囲に対し、以下の2段階で最適化を実行する:
 * 
 * Phase 1: 粗配分 (allocateRoughly)
 *   - 各変数の効用（1ポイント追加時のダメージ増加）を測定
 *   - 効用が高い変数に優先的にポイントを配分
 * 
 * Phase 2: リバランス (rebalance)
 *   - 局所探索により、ポイント移動で改善できる箇所を探す
 *   - 0割当変数の再評価により、見落としがないかチェック
 * 
 * @example
 * // B2:B11のセル範囲に40ポイントを配分し、D2のダメージを最大化
 * const result = runOptimization('B2:B11', 'D2', 40);
 * console.log(`最終ダメージ: ${result.final}`);
 */
function runOptimization(varRangeStr, calcCellStr, totalPoints) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const varRange = sheet.getRange(varRangeStr);
  const calcCell = sheet.getRange(calcCellStr);
  const numVars = varRange.getNumRows();

  // state は現在の変数配列と、外部で参照計算した回数（負荷指標）を保持する。
  let state = {
    values: new Array(numVars).fill(0),
    calcCount: 0
  };

  // --- 初期化: 変数を全て0にし、初期ダメージを取得 ---
  // 注意: ここで一度全ての値を書き込んでflushするため、シート上の既存値は上書きされる。
  varRange.setValues(state.values.map(v => [v]));
  SpreadsheetApp.flush();
  const initialDamage = calcCell.getValue();
  if (initialDamage <= 0) throw new Error('初期ダメージが0以下です');

  // Phase 1: 粗い配分。効用（1ポイントあたりのダメージ増加）に基づいて一括で配分する。
  const roughResult = allocateRoughly(state, totalPoints, varRange, calcCell);
  state = roughResult.state;
  const roughDamage = calcCell.getValue();

  // Phase 2: リバランス。配分後に局所的なポイント移動で改善できるか探索する。
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
 * 粗配分の結果
 * @typedef {Object} AllocateResult
 * @property {State} state - 更新された状態
 * @property {number[]} utilities - 各変数の効用値（1ポイント追加時のダメージ増加量）
 */

/**
 * 効用に基づきポイントを粗く配分する
 * @function
 * @param {State} state - 現在の状態（valuesとcalcCountを含む）
 * @param {number} totalPoints - 配分する総ポイント数
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - サブステータス値を書き込むセル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {AllocateResult} 更新された状態と効用配列
 * @description
 * BATCH_SIZE単位でポイントを配分する。各バッチでは:
 * 1. 各変数の1ポイント追加時の効用を取得
 * 2. 効用が高い上位サブステに対して効率性に応じたヒット数を割り当てる
 * 3. スプシの更新
 * この処理を総ポイント数に達するまで繰り返す。
 */
function allocateRoughly(state, totalPoints, varRange, calcCell) {
  const numVars = state.values.length;
  let currentState = { ...state, values: [...state.values] };
  let latestUtilities = new Array(numVars).fill(0);
  let allocated = 0;

  while (allocated < totalPoints) {
    const batch = Math.min(CONFIG.BATCH_SIZE, totalPoints - allocated);
    // 各変数の1ポイント追加時の効用を取得
    const measureResult = measureUtilities(currentState, varRange, calcCell);
    currentState = measureResult.state;
    latestUtilities = measureResult.utilities;

    // 効用が高い上位変数を取り出す（効用が0のものは除外）
    const sorted = latestUtilities
      .map((u, i) => ({ i, u }))
      .sort((a, b) => b.u - a.u)
      .slice(0, CONFIG.TOP_VARS)
      .filter(item => item.u > 0);

    const newValues = [...currentState.values];

    if (sorted.length > 0) {
      // 上位変数の効用比に応じてバッチ内で分配
      const totalUtil = sorted.reduce((sum, item) => sum + item.u, 0);
      let remaining = batch;

      for (const item of sorted) {
        // 小数点は四捨五入して整数ポイントにする
        const points = Math.min(
          Math.round(batch * item.u / totalUtil),
          remaining
        );
        newValues[item.i] += points;
        remaining -= points;
      }
      // 端数が残ったら最も効用の高い変数に追加
      if (remaining > 0) newValues[sorted[0].i] += remaining;
    } else {
      // 全て効用が0ならとりあえず最初の変数へ付与（戦略的ではないがほぼありえない状況のためテキトー）
      newValues[0] += batch;
    }

    currentState = { ...currentState, values: newValues };
    allocated += batch;

    // シートに書き戻して評価セルの再計算を促す
    varRange.setValues(newValues.map(v => [v]));
    SpreadsheetApp.flush();
  }

  return {
    state: currentState,
    utilities: latestUtilities
  };
}

/**
 * 効用測定の結果
 * @typedef {Object} MeasureResult
 * @property {State} state - 更新された状態（calcCountがインクリメントされる）
 * @property {number[]} utilities - 各変数の効用値
 */

/**
 * 各変数を1ポイント増加させた場合の効用を計測する
 * @function
 * @param {State} state - 現在の状態
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {MeasureResult} 更新された状態と効用配列
 * @description
 * 各変数に対して:
 * 1. 現在の値に+1した値をシートに書き込む
 * 2. ダメージセルを評価し、増加量を効用として記録
 * 3. 次の変数の評価に備えて元の値に戻す
 * すべての変数の評価後、必ず元の状態に復元する。
 */
function measureUtilities(state, varRange, calcCell) {
  const currentDamage = calcCell.getValue();
  const utilities = [];
  let newCalcCount = state.calcCount;

  for (let i = 0; i < state.values.length; i++) {
    const testValues = [...state.values];
    testValues[i]++;

    // シートに書き込み、計算セルを更新して効用を測る
    varRange.setValues(testValues.map(v => [v]));
    SpreadsheetApp.flush();
    utilities[i] = Math.max(0, calcCell.getValue() - currentDamage);
    newCalcCount++;
  }

  // この計測はあくまで各ステータスの効率性を図るだけなので、必ず最後に元のサブステヒット数に戻す必要がある
  varRange.setValues(state.values.map(v => [v]));
  SpreadsheetApp.flush();

  return {
    state: { ...state, calcCount: newCalcCount },
    utilities
  };
}

/**
 * リバランスの結果
 * @typedef {Object} RebalanceResult
 * @property {State} state - 更新された状態
 * @property {number} improvements - 改善が見つかった回数
 */

/**
 * 配分済みポイントを局所的に移動し、改善がある場合に再配分する
 * @function
 * @param {State} state - 現在の状態
 * @param {number[]} initialUtilities - 粗配分時に計測された効用値
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {RebalanceResult} 更新された状態と改善回数
 * @description
 * 2つの最適化手法を順次適用する:
 * 1. localOptimization: 既に割り当てられた変数間でポイントを移動
 * 2. zeroHitSwapOptimization: 0割当の変数と割当済み変数を入れ替え
 * それぞれの手法で改善が見つかった回数を合計して返す。
 */
function rebalance(state, initialUtilities, varRange, calcCell) {
  let currentState = { ...state, values: [...state.values] };
  let utilities = [...initialUtilities];
  let improvements = 0;

  improvements += localOptimization(currentState, utilities, varRange, calcCell);
  improvements += zeroHitSwapOptimization(currentState, varRange, calcCell);

  return { state: currentState, improvements };
}

/**
 * 既に割り当てられた変数間でポイントを移動し、局所的な改善を行う
 * @function
 * @param {State} currentState - 現在の状態（この関数内で直接更新される）
 * @param {number[]} utilities - 各変数の効用値（移動時に動的に更新される）
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {number} 改善が見つかった回数
 * @description
 * MAX_ITERATIONS回まで以下を繰り返す:
 * 1. ポイントが割り当てられている変数をすべて取得
 * 2. 任意の2変数間でポイントを移動する候補を生成
 * 3. 効用差（utilities[to] - utilities[from]）でソート
 * 4. 上位MAX_CANDIDATES個の候補を実際に試す
 * 5. ダメージが改善する移動があれば適用し、効用を調整
 * 6. 改善がなければ探索を終了
 */
function localOptimization(currentState, utilities, varRange, calcCell) {
  let improvements = 0;
  for (let iteration = 0; iteration < CONFIG.MAX_ITERATIONS; iteration++) {
    const activeVars = currentState.values.map((v, i) => v > 0 ? i : -1).filter(i => i >= 0);
    if (activeVars.length <= 1) break;

    const candidates = [];
    for (const from of activeVars) {
      for (const to of activeVars) {
        if (from !== to) candidates.push({ from, to, priority: utilities[to] - utilities[from] });
      }
    }
    if (candidates.length === 0) break;

    candidates.sort((a, b) => b.priority - a.priority);
    const baselineDamage = calcCell.getValue();
    let bestMove = null;

    for (let i = 0; i < Math.min(CONFIG.MAX_CANDIDATES, candidates.length); i++) {
      const c = candidates[i];
      const testValues = [...currentState.values];
      testValues[c.from]--;
      testValues[c.to]++;

      varRange.setValues(testValues.map(v => [v]));
      SpreadsheetApp.flush();

      if (calcCell.getValue() > baselineDamage + CONFIG.THRESHOLD) {
        bestMove = c;
        currentState.values = testValues;
        improvements++;
        utilities[c.from] *= 0.95;
        utilities[c.to] *= 1.05;
        break;
      }
    }

    if (!bestMove) break;
  }

  return improvements;
}

/**
 * 0割当のサブステに最適なものが含まれていないかを再評価する
 * @function
 * @param {State} currentState - 現在の状態（改善時に直接更新される）
 * @param {GoogleAppsScript.Spreadsheet.Range} varRange - 変数セル範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} calcCell - ダメージ計算セル
 * @returns {number} 改善があれば1、なければ0
 * @description
 * 局所探索後の状態をベースラインとし、以下を行う:
 * 1. 0割当の変数と、ポイントが割り当てられている変数のペアをすべて列挙
 * 2. 各ペアについて、1ポイントを移動した場合のダメージ増加量を計算
 * 3. 最も改善量が大きいペアを見つける
 * 4. 改善がTHRESHOLD以上なら適用、なければ局所探索後の状態に戻す
 */
function zeroHitSwapOptimization(currentState, varRange, calcCell) {
  const zeroVars = currentState.values.map((v, i) => v === 0 ? i : -1).filter(i => i >= 0);
  const hitVars = currentState.values.map((v, i) => v > 0 ? i : -1).filter(i => i >= 0);

  // 局所探索後の状態を確実に反映
  varRange.setValues(currentState.values.map(v => [v]));
  SpreadsheetApp.flush();
  const baselineDamage = calcCell.getValue();

  let bestSwap = null;
  let bestGain = 0;

  for (const zero of zeroVars) {
    for (const hit of hitVars) {
      const testValues = [...currentState.values];
      testValues[zero]++;
      testValues[hit]--;

      // 一時的に反映して評価
      varRange.setValues(testValues.map(v => [v]));
      SpreadsheetApp.flush();
      const gain = calcCell.getValue() - baselineDamage;

      if (gain > bestGain) {
        bestGain = gain;
        bestSwap = { zero, hit };
      }
    }
  }

  let improvements = 0;

  if (bestSwap !== null && bestGain > CONFIG.THRESHOLD) {
    currentState.values[bestSwap.zero]++;
    currentState.values[bestSwap.hit]--;
    improvements = 1;
    varRange.setValues(currentState.values.map(v => [v]));
    SpreadsheetApp.flush();
  } else {
    // 改善がなければ、必ず局所探索後の状態に戻す
    varRange.setValues(currentState.values.map(v => [v]));
    SpreadsheetApp.flush();
  }

  return improvements;
}
