/** シートのセル値 */
type CellValue = (string | number | boolean | Date);

/** KeepaAPI商品情報 */
interface KeepaProduct {
  imagesCSV: string,
  asin: string,
  title: string,
  brand: string,
  rootCategory: string,
  model: string,
  eanList: string[],
  csv: number[][],
  stats: {
    avg30: number[]
    salesRankDrops30: number,
    salesRankDrops90: number,
    buyBoxPrice: number,
    buyBoxShipping: number,
  },
  liveOffersOrder?: number[],
  offers?: {
    lastSeen: number,
    sellerId: string,
    offerCSV: number[],
    condition: number,
    isPrime: boolean,
    isMAP: boolean,
    isShippable: boolean,
    isAddonItem: boolean,
    isPreorder: boolean,
    isWarehouseDeal: boolean,
    isScam: boolean,
    isAmazon: boolean,
    isPrimeExcl: boolean,
    offerId: 5,
    isFBA: boolean,
    shipsFromChina: boolean
  }[],
}

/** シートの見出し名 */
type ItemHeader =
  'ASIN' |
  '商品名' |
  'カート価格' |
  'リンク';

/** 商品の行 */
type Item = { [key in ItemHeader]: string };

