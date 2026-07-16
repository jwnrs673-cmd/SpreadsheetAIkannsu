/**
 * 報告書リンク（別スクリプト）— クレームAI分類 統合版と共存できるように修正した版。
 * ------------------------------------------------------------------
 * 変更点は末尾の onOpen だけ：
 *   もとの function onOpen() を削除し、代わりに buildReportLinkMenu_() を定義。
 *   → 統合版の onOpen が起動時にこれを自動で呼び、両方のメニューが表示される。
 *   （メニュー登録以外のロジックは一切変更していません）
 * ------------------------------------------------------------------
 */

/**
 * 基本設定
 */
const SETTINGS = {
  SHEET_NAME: 'ご意見記録',
  FOLDER_ID: '1zNrC8BAddxZZ98IW0haesp2BrMEHWMbY',
  HEADER_ROW: 1,
  FILE_NAME_HEADER: '報告書リンク',
  OUTPUT_HEADER: '報告書URL'
};

/**
 * メイン処理
 */
function updateReportLinks() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SETTINGS.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      '対象シート「' +
      SETTINGS.SHEET_NAME +
      '」が見つかりません。'
    );
  }

  // ヘッダー名から列番号を取得
  const fileNameColumn = findHeaderColumn_(
    sheet,
    SETTINGS.HEADER_ROW,
    SETTINGS.FILE_NAME_HEADER
  );

  const outputColumn = findHeaderColumn_(
    sheet,
    SETTINGS.HEADER_ROW,
    SETTINGS.OUTPUT_HEADER
  );

  if (fileNameColumn === outputColumn) {
    throw new Error(
      'ファイル名の列とURL出力列が同じになっています。'
    );
  }

  const firstDataRow = SETTINGS.HEADER_ROW + 1;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) {
    spreadsheet.toast(
      '処理対象のデータがありません。',
      '報告書URL更新',
      5
    );
    return;
  }

  const rowCount = lastRow - firstDataRow + 1;

  // ファイル名が入力されている列
  const fileNames = sheet
    .getRange(
      firstDataRow,
      fileNameColumn,
      rowCount,
      1
    )
    .getDisplayValues();

  // URL出力列
  const outputRange = sheet.getRange(
    firstDataRow,
    outputColumn,
    rowCount,
    1
  );

  const currentOutputs =
    outputRange.getDisplayValues();

  /*
   * 以前設定した「ファイル名表示のリンク」を
   * 判定するためにリッチテキスト情報も取得
   */
  const currentRichTexts =
    outputRange.getRichTextValues();

  // 指定フォルダを取得
  const folder = DriveApp.getFolderById(
    SETTINGS.FOLDER_ID
  );

  // フォルダ内のExcelファイルを一覧化
  const fileIndex = buildExcelFileIndex_(folder);

  let urlCount = 0;
  let notFoundCount = 0;
  let noFileNameCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < rowCount; i++) {
    const rowNumber = firstDataRow + i;

    const fileName = String(
      fileNames[i][0]
    ).trim();

    const currentValue = String(
      currentOutputs[i][0]
    ).trim();

    const richText = currentRichTexts[i][0];

    /*
     * 以前のコードで設定した
     * 「ファイル名表示のリンク」があるか確認
     */
    const existingLinkUrl = richText
      ? richText.getLinkUrl()
      : null;

    const isOldFileNameLink =
      existingLinkUrl &&
      !isUrl_(currentValue);

    /*
     * 次の場合だけ処理
     *
     * ・URL列が空欄
     * ・URL列が「なし」
     * ・以前作成したファイル名表示のリンク
     *
     * すでにURLが入っている行は変更しません。
     */
    const shouldProcess =
      currentValue === '' ||
      currentValue === 'なし' ||
      isOldFileNameLink;

    if (!shouldProcess) {
      skippedCount++;
      continue;
    }

    // ファイル名が空欄の場合は処理しない
    if (fileName === '') {
      noFileNameCount++;
      continue;
    }

    // ファイル名からExcelファイルを検索
    const file = findFileFromIndex_(
      fileName,
      fileIndex
    );

    const outputCell = sheet.getRange(
      rowNumber,
      outputColumn
    );

    if (file) {
      /*
       * URLを文字列としてセルに保存
       *
       * IMPORTRANGEやQUERYで別シートに連携しても、
       * URL文字列が引き継がれます。
       */
      outputCell.setValue(file.getUrl());

      urlCount++;
    } else {
      outputCell.setValue('なし');

      notFoundCount++;
    }
  }

  spreadsheet.toast(
    'URL設定：' + urlCount + '件\n' +
    '該当なし：' + notFoundCount + '件\n' +
    'ファイル名空欄：' + noFileNameCount + '件\n' +
    '処理対象外：' + skippedCount + '件',
    '報告書URL更新完了',
    10
  );
}

/**
 * ヘッダー名から列番号を取得
 */
function findHeaderColumn_(
  sheet,
  headerRow,
  targetHeader
) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error(
      'シートにヘッダーがありません。'
    );
  }

  const headers = sheet
    .getRange(
      headerRow,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const matchingColumns = [];

  for (let i = 0; i < headers.length; i++) {
    const headerName = String(
      headers[i]
    ).trim();

    if (headerName === targetHeader) {
      matchingColumns.push(i + 1);
    }
  }

  if (matchingColumns.length === 0) {
    throw new Error(
      headerRow +
      '行目にヘッダー「' +
      targetHeader +
      '」が見つかりません。'
    );
  }

  if (matchingColumns.length > 1) {
    throw new Error(
      headerRow +
      '行目にヘッダー「' +
      targetHeader +
      '」が複数あります。'
    );
  }

  return matchingColumns[0];
}

/**
 * フォルダ内のExcelファイルを一覧化
 */
function buildExcelFileIndex_(folder) {
  const fileIndex = {
    fullName: Object.create(null),
    baseName: Object.create(null)
  };

  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (!isExcelFile_(file)) {
      continue;
    }

    const actualFileName = String(
      file.getName()
    ).trim();

    // 拡張子を含むファイル名
    const fullNameKey = normalizeFileName_(
      actualFileName
    );

    setNewestFile_(
      fileIndex.fullName,
      fullNameKey,
      file
    );

    // 拡張子を除いたファイル名
    const baseName = actualFileName.replace(
      /\.(xlsx|xls|xlsm|xlsb)$/i,
      ''
    );

    const baseNameKey = normalizeFileName_(
      baseName
    );

    setNewestFile_(
      fileIndex.baseName,
      baseNameKey,
      file
    );
  }

  return fileIndex;
}

/**
 * 入力されたファイル名からファイルを検索
 */
function findFileFromIndex_(
  inputFileName,
  fileIndex
) {
  const normalizedName = normalizeFileName_(
    inputFileName
  );

  const hasExcelExtension =
    /\.(xlsx|xls|xlsm|xlsb)$/i.test(
      inputFileName
    );

  // 拡張子が入力されている場合
  if (hasExcelExtension) {
    return (
      fileIndex.fullName[normalizedName] ||
      null
    );
  }

  // 拡張子が入力されていない場合
  return (
    fileIndex.fullName[normalizedName] ||
    fileIndex.baseName[normalizedName] ||
    null
  );
}

/**
 * ファイル名を検索用に統一
 */
function normalizeFileName_(fileName) {
  return String(fileName)
    .trim()
    .toLowerCase();
}

/**
 * 同名ファイルが複数ある場合、
 * 更新日時が新しいファイルを採用
 */
function setNewestFile_(
  targetIndex,
  key,
  file
) {
  if (!key) {
    return;
  }

  const currentFile = targetIndex[key];

  if (
    !currentFile ||
    file.getLastUpdated().getTime() >
      currentFile.getLastUpdated().getTime()
  ) {
    targetIndex[key] = file;
  }
}

/**
 * Excelファイルかどうかを判定
 */
function isExcelFile_(file) {
  const excelMimeTypes = [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel.sheet.macroEnabled.12',
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
  ];

  const mimeType = file.getMimeType();
  const fileName = file.getName();

  return (
    excelMimeTypes.includes(mimeType) ||
    /\.(xlsx|xls|xlsm|xlsb)$/i.test(fileName)
  );
}

/**
 * URLかどうかを判定
 */
function isUrl_(value) {
  return /^https?:\/\//i.test(
    String(value).trim()
  );
}

/**
 * 「報告書リンク」メニューを構築。
 * ------------------------------------------------------------------
 * ※ もとの onOpen() はここに置き換えました（削除しました）。
 *   統合版 Code.gs の onOpen が起動時にこの関数を自動で呼び出します。
 *   この関数名（buildReportLinkMenu_）は変更しないでください。
 * ------------------------------------------------------------------
 */
function buildReportLinkMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('報告書リンク')
    .addItem(
      'URLを更新する',
      'updateReportLinks'
    )
    .addToUi();
}
