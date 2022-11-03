/** シートのセル値 */
type CellValue = (string | number | boolean | Date);

/** 商品情報 */
interface Item {
  asin: string,
  title: string
  buybox_price: number,
  url: string,
}

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
