/** ASINの処理単位 */
const CHUNCK_ASINS = 5;
/** データシート名 */
const RESEARCH_SHEET = 'リサーチ';
const RESEARCH_TABLE_HEAD_ROW = 1;
/** プロパティサービス */
const SCRIPT_PROP = PropertiesService.getScriptProperties();

/** 永続的に保持する値 */
const SETTING = {
  get keepa_api_key() {
    return SpreadsheetApp.getActive().getRangeByName('keepa_key')?.getValue() || '';
  },
  get update_asin_len() {
    return parseInt(SpreadsheetApp.getActive().getRangeByName('update_asin_len')?.getValue() || '0');
  },
  /** 前回の更新位置 */
  get last_update_row() {
    return JSON.parse(SCRIPT_PROP.getProperty('last_update_row') || '{}');
  },
  set last_update_row(row: { [asin: string]: number }) {
    SCRIPT_PROP.setProperty('last_update_row', JSON.stringify(row));
  },
  get last_update_row_str() {
    const last = SETTING.last_update_row;
    const asin = Object.keys(last)[0];
    return asin ? `最終更新位置 ${last[asin]}行(${asin})` : 'なし';
  },
};

/** ユーティリティクラス */
class Util {
  /** １次元配列を、指定した要素数で分割した２次元配列に変換 */
  static bunch<T>(arr: Array<T>, chunk: number) {
    return [...Array(Math.ceil(arr.length / chunk))].map((_, i) =>
      arr.slice(i * chunk, (i + 1) * chunk)
    );
  }

  /** 二次元のテーブルデータを見出しをキーとしたオブジェクト配列に変換する */
  static namingCellValues(df: CellValue[][]) {
    const [head, ...values] = df;
    const records = (values as CellValue[][]).map((r) =>
      (head as CellValue[]).reduce((acc, value, i) => {
        acc[value.toString()] = r[i]?.toString() || '';
        return acc;
      }, {} as { [key: string]: string })
    );
    return records;
  }
}

/** KeepaAPI用クラス */
class Keepa {
  static KEEPA_API_EP = 'https://api.keepa.com/';
  /** ASINの商品情報を取得 */
  static requestProducts(asins: string[]) {
    const option: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      payload: {
        domain: '5',
        stats: '90',
        rating: '0',
        buybox: '1',
        key: SETTING.keepa_api_key,
        asin: asins.join(',')
      },
      muteHttpExceptions: true,
    };
    const resp = UrlFetchApp.fetch(this.KEEPA_API_EP + 'product', option);

    // Error handling
    const code = resp.getResponseCode();
    if (code === 429)
      throw new Error('Keepaのトークンが不足しています');
    if (code !== 200)
      throw new Error('failed to fetch keepa product api');
    const { error, products }: { error?: any, products: KeepaProduct[] } = JSON.parse(resp.getContentText());
    if (error)
      throw new Error(`failed to fetch keepa product api ${error.message}`);
    return { products, code };
  }
}

/** スクリプト実行用メニューの追加 */
function onOpen() {
  const menu = [{
    name: '🔄 シート更新',
    functionName: 'beginUpdateSheet',
  }];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('メニュー', menu);
}

/** 前回の更新位置から、今回の更新対象ASINを取得 */
function getTargetAsins(records: { [key: string]: string }[]) {
  // 現在の全ASIN
  const curr_asins = records.map((r) => r['ASIN']);

  // 前回の更新位置を取得
  const last_update = SETTING.last_update_row;
  const last_asin = Object.keys(last_update)?.[0];
  const last_row = last_update[last_asin] || 0;

  // 前回の更新位置よりも下のASINを抽出
  const filtered_asins = records.filter((r, i) =>
    (r['ASIN'] && last_row < (RESEARCH_TABLE_HEAD_ROW + 1 + i))
  ).map((r) => r['ASIN']);

  // 今回更新する順序でASINを並べ替え
  const target_asins = [
    ...filtered_asins,
    ...curr_asins,
  ].slice(0, SETTING.update_asin_len);

  return target_asins;
}

/** シートにデータを書き込み */
function writeItem(records: { [key: string]: string }[], items: Item[]) {
  // 最終更新位置の保持
  const last = { row: 0, asin: '' };
  // 書き込み用データの生成
  const data: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];
  items.map((item) => {
    records.map((record, i) => {
      if (record['ASIN'] === item.asin) {
        const row = RESEARCH_TABLE_HEAD_ROW + 1 + i;
        data.push({
          range: `${RESEARCH_SHEET}!A${row}:ZZ${row}`,
          values: [[
            item.asin,
            item.title,
            item.buybox_price,
            item.url
          ]],
        });
        last.row = row;
        last.asin = item.asin;
      }
    });
  });

  // シートに書き込み
  const ss = SpreadsheetApp.getActive();
  Sheets.Spreadsheets?.Values?.batchUpdate({
    valueInputOption: 'USER_ENTERED',
    data,
  }, ss.getId());

  // 書き込み位置を保存
  SETTING.last_update_row = { [last.asin]: last.row };
}

/** シートの全データ更新 */
function beginUpdateSheet() {
  // シートの全データを取得
  const ss = SpreadsheetApp.getActive();
  const resp = Sheets.Spreadsheets?.Values?.batchGet(ss.getId(), {
    ranges: [`${RESEARCH_SHEET}!A${RESEARCH_TABLE_HEAD_ROW}:ZZ`],
  });
  const df: CellValue[][] = resp?.valueRanges ? (resp.valueRanges[0].values || [[]]) : [[]];
  const records = Util.namingCellValues(df);

  try {
    ss.toast('', 'Keepa情報取得...');

    // 対象のASINを更新順で取り出し
    Util.bunch(getTargetAsins(records), CHUNCK_ASINS).map((asins) => {
      // Keepaのデータから必要な情報を抽出
      const { products } = Keepa.requestProducts(asins);
      const items = products.map((product) => {
        const { asin, title, stats } = product;
        const { buyBoxPrice } = stats;
        const item: Item = {
          asin,
          title,
          buybox_price: buyBoxPrice,
          url: 'https://www.amazon.co.jp/dp/' + asin,
        };
        return item;
      });

      // シートに書き込み
      writeItem(records, items);
      ss.toast(SETTING.last_update_row_str, 'Keepa情報取得...');
    });
    ss.toast(SETTING.last_update_row_str, 'Keepa情報取得完了');
  } catch (e) {
    // Error
    ss.toast([SETTING.last_update_row_str, (e as Error).message].join('\n'), 'Error');
  }
}
