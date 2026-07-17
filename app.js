const STORAGE_KEY = "overseas-inventory-board-v2";
const TRANSLATE_CACHE_STORAGE_KEY = `${STORAGE_KEY}:google-translate-cache`;
const GOOGLE_TRANSLATE_ENDPOINT = "https://translate.googleapis.com/translate_a/single";
const TRANSLATE_SOURCE_LANGUAGE = "zh-CN";
const SUPPORTED_TRANSLATION_LANGUAGES = new Set([
  "zh-CN",
  "en",
  "de",
  "fr",
  "es",
  "it",
  "pt",
  "ja",
  "ko",
  "ru",
  "vi",
  "th",
  "ms",
  "id",
  "tr",
]);

const warehouses = [
  {
    id: "frankfurt",
    code: "DE",
    name: "德国法兰克福仓",
    sourceName: "德国法兰克福仓（Germany-UEX）",
    city: "德国法兰克福",
    role: "欧洲售后与整机库存",
    accent: "#315f96",
    soft: "#e5eef9",
    x: "54%",
    y: "40%",
    mobileY: "18%",
  },
  {
    id: "netherlands",
    code: "NL",
    name: "荷兰仓",
    sourceName: "荷兰仓",
    city: "荷兰",
    role: "欧洲补充周转仓",
    accent: "#25808a",
    soft: "#e0f3f5",
    x: "43%",
    y: "24%",
    mobileY: "38%",
  },
  {
    id: "chicago",
    code: "US",
    name: "美国芝加哥仓",
    sourceName: "美国芝加哥仓",
    city: "美国芝加哥",
    role: "北美备件库存",
    accent: "#92651f",
    soft: "#fff2d6",
    x: "24%",
    y: "42%",
    mobileY: "58%",
  },
  {
    id: "sydney",
    code: "AU",
    name: "澳大利亚悉尼仓",
    sourceName: "悉尼仓",
    city: "澳大利亚悉尼",
    role: "澳洲备件库存",
    accent: "#b84a62",
    soft: "#fde8ed",
    x: "82%",
    y: "72%",
    mobileY: "78%",
  },
];

const domesticWarehouses = [
  {
    id: "fuzhou",
    code: "FZ",
    name: "福州仓",
    sourceName: "FJDYN 福州成品仓 福州原材料仓 福州半成品仓",
    city: "中国福州",
    role: "国内备货与发运仓",
    accent: "#5f6f7e",
    soft: "#eef2f6",
    x: "70%",
    y: "43%",
    mobileY: "88%",
  },
];

const inventoryProjectScope = new Set(["R1916", "R1917", "R2404"]);
const syncLocationScope = ["SVEA", "售后备件"];

const excelRows = [
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "372200100J0260A",
        "materialName":  "车控盒总成",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "美国芝加哥仓",
        "project":  "R1916",
        "materialCode":  "S050160000502",
        "materialName":  "大垫圈 A级",
        "category":  "零部件及主材_标准件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "SVEA HK",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 96.1 5-达克罗96h",
        "model":  "GB/T 96.1 5-达克罗96h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0100080J0260B",
        "materialName":  "锁扣",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "YXL-XS-003033A",
        "materialName":  "线束",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  20,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "BZB-KZ-003095A",
        "materialName":  "控制类",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  19,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "控制类-R1911_HW-MCU-CONTROLER-40A-V4.1-大割1500W电机控制板",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "900280200J0080D",
        "materialName":  "LoRaWAN智能网关",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  1,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "UG67-868",
        "model":  "UG67-L04EU-868M"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "QCG-CS-003003A",
        "materialName":  "超声波测距传感器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  8,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "超声波测距传感器-3.3~24V-RS485-50~500cm--25~65℃-IP67",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R2312",
        "materialCode":  "J0300018",
        "materialName":  "Sveaverken FeedLogger Pro",
        "category":  "成品_畜牧",
        "productLine":  "畜牧-FJ05",
        "location":  "借用还回",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "BZB-CG-003042C",
        "materialName":  "传感器板",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  6,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "传感器板-R1916-HW-CG-PSO-V1-光敏检测板",
        "model":  "传感器板-R1916-HW-CG-PSO-V1-光敏检测板"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "J0080009",
        "materialName":  "智能项圈",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "展会用品",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "868",
        "model":  "SCN100"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "BZB-QD-003115B",
        "materialName":  "驱动类",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  20,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "驱动类-R1916-HW-MCU-DRIVER-40A-V3-驱动板",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "214200000J0260D",
        "materialName":  "MCU总成",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  7,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "荷兰仓",
        "project":  "R1916",
        "materialCode":  "1H0000000J0260A",
        "materialName":  "底部配重装箱",
        "category":  "半成品_实件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  13,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "310600010J0020D",
        "materialName":  "驱动轮",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "Φ250*80",
        "model":  "Φ250*80"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0100064J0030B",
        "materialName":  "补光灯",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R2312",
        "materialCode":  "J0300018",
        "materialName":  "Sveaverken FeedLogger Pro",
        "category":  "成品_畜牧",
        "productLine":  "畜牧-FJ05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "280000362J0260A",
        "materialName":  "电源转换器支架",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  5,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "S050160000502",
        "materialName":  "大垫圈 A级",
        "category":  "零部件及主材_标准件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 96.1 5-达克罗96h",
        "model":  "GB/T 96.1 5-达克罗96h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0112000J0260B",
        "materialName":  "超声波测距总成",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "373600014J0260D",
        "materialName":  "急停开关",
        "category":  "零部件及主材_电器件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  7,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "无",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R2312",
        "materialCode":  "J0300018",
        "materialName":  "Sveaverken FeedLogger Pro",
        "category":  "成品_畜牧",
        "productLine":  "畜牧-FJ05",
        "location":  "退货",
        "unit":  "个",
        "onHandQty":  2,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "L0056390",
        "materialName":  "USB转485/422转换器通讯线模块",
        "category":  "零部件及主材_电器件",
        "productLine":  "畜牧-FJ05",
        "location":  "退货",
        "unit":  "个",
        "onHandQty":  2,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "摩可灵USB转485/422转换器通讯线模块"
    },
    {
        "warehouseName":  "美国芝加哥仓",
        "project":  "R1916",
        "materialCode":  "S010200501001",
        "materialName":  "内六角圆柱头螺钉",
        "category":  "零部件及主材_标准件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "SVEA HK",
        "unit":  "个",
        "onHandQty":  50,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 70.1 M5×10-A2-70-钝化72h",
        "model":  "GB/T 70.1 M5×10-A2-70-钝化72h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1H0000000J0260A",
        "materialName":  "底部配重装箱",
        "category":  "半成品_实件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  2,
        "reservedQty":  2,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "211100024J0030B",
        "materialName":  "磁导航传感器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "J0080005",
        "materialName":  "智能项圈",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "868",
        "model":  "SCN100"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "210310004J0020C",
        "materialName":  "行走驱动电机",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0100072J0030D",
        "materialName":  "补光灯支架",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "无",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "211100014J0030B",
        "materialName":  "陀螺仪",
        "category":  "零部件及主材_电器件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  12,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "无",
        "model":  "无"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "370000124J0260A",
        "materialName":  "直流变换器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  5,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "BZB-CG-003042JC",
        "materialName":  "传感器板",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "传感器板-R1916-HW-CG-PSO-V1-光敏检测板",
        "model":  "传感器板-R1916-HW-CG-PSO-V1-光敏检测板"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "L0056390",
        "materialName":  "USB转485/422转换器通讯线模块",
        "category":  "零部件及主材_电器件",
        "productLine":  "畜牧-FJ05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  4,
        "reservedQty":  1,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "摩可灵USB转485/422转换器通讯线模块"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0114004J0260A",
        "materialName":  "视觉算法板",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  9,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "悉尼仓",
        "project":  "R1916",
        "materialCode":  "S010470502501",
        "materialName":  "带弹平垫盘头螺栓",
        "category":  "零部件及主材_标准件",
        "productLine":  "智能割草机配件",
        "location":  "全新品",
        "unit":  "个",
        "onHandQty":  100,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 9074.4 M5×25-A2-70-钝化72H",
        "model":  "GB/T 9074.4 M5×25-A2-70-钝化72h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1R0000011J0260B",
        "materialName":  "胶皮",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  36,
        "reservedQty":  2,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R2312",
        "materialCode":  "J0300014",
        "materialName":  "Sveaverken FeedLogger",
        "category":  "成品_畜牧",
        "productLine":  "畜牧-FJ05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  2,
        "reservedQty":  1,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "BZB-SF-003016G",
        "materialName":  "算法类",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "算法类-A311D机器人算法板-R2003-SFB-V1-处理机器人各个功能模块运行，数据收发运算",
        "model":  "无"
    },
    {
        "warehouseName":  "悉尼仓",
        "project":  "R1916",
        "materialCode":  "S050160000502",
        "materialName":  "大垫圈 A级",
        "category":  "零部件及主材_标准件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "全新品",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 96.1 5-达克罗96h",
        "model":  "GB/T 96.1 5-达克罗96h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "280000012J0260D",
        "materialName":  "电池扣板",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  3,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "210310004J0020C",
        "materialName":  "行走驱动电机",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  6,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "900280200J0080D",
        "materialName":  "LoRaWAN智能网关",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "展会用品",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "UG67-868",
        "model":  "UG67-L04EU-868M"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "J0260013",
        "materialName":  "mini推料机",
        "category":  "成品_畜牧",
        "productLine":  "推料机器人-05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  8,
        "reservedQty":  3,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "9WZ-1.05A(SCP300)"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "L0056390",
        "materialName":  "USB转485/422转换器通讯线模块",
        "category":  "零部件及主材_电器件",
        "productLine":  "畜牧-FJ05",
        "location":  "借用还回",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "摩可灵USB转485/422转换器通讯线模块"
    },
    {
        "warehouseName":  "悉尼仓",
        "project":  "R1916",
        "materialCode":  "QXL-XS-003310B",
        "materialName":  "主线束",
        "category":  "零部件及主材_电器件",
        "productLine":  "其他-FJ05",
        "location":  "全新品",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "线束-平地机-2300mm-1",
        "model":  "线束-平地机-2300mm-1"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "S010200401202",
        "materialName":  "内六角圆柱头螺钉",
        "category":  "零部件及主材_标准件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  30,
        "reservedQty":  30,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 70.1 M4×12-A2-70-钝化72h",
        "model":  "GB/T 70.1 M4×12-A2-70-钝化72h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "J0360013",
        "materialName":  "推料机器人",
        "category":  "成品_畜牧",
        "productLine":  "推料机器人-05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  4,
        "reservedQty":  4,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "9WZ-1.05A(SCP300)-X"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "213412004J0020D",
        "materialName":  "接触器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  10,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "DH50",
        "model":  ""
    },
    {
        "warehouseName":  "荷兰仓",
        "project":  "R1916",
        "materialCode":  "J0350001",
        "materialName":  "自动推料机",
        "category":  "成品_畜牧",
        "productLine":  "推料机器人-05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  2,
        "reservedQty":  2,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "LF02"
    },
    {
        "warehouseName":  "悉尼仓",
        "project":  "R1916",
        "materialCode":  "S010200501001",
        "materialName":  "内六角圆柱头螺钉",
        "category":  "零部件及主材_标准件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "全新品",
        "unit":  "个",
        "onHandQty":  50,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 70.1 M5×10-A2-70-钝化72h",
        "model":  "GB/T 70.1 M5×10-A2-70-钝化72h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "372200100J0260A",
        "materialName":  "车控盒总成",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "Q0039067",
        "materialName":  "按键板总成",
        "category":  "成品_智能套件",
        "productLine":  "其他-FJ02-SZ",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "AT-EC02"
    },
    {
        "warehouseName":  "荷兰仓",
        "project":  "R1916",
        "materialCode":  "J0260013",
        "materialName":  "mini推料机",
        "category":  "成品_畜牧",
        "productLine":  "推料机器人-05",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  7,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  "9WZ-1.05A(SCP300)"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "372220004J0020A",
        "materialName":  "熔断器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  6,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "80A",
        "model":  "AET-80"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "213412004J0020D",
        "materialName":  "接触器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  1,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "DH50",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "372210004J0020A",
        "materialName":  "熔断器",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  6,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "63A",
        "model":  "AET-63"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1S0140000J0260H",
        "materialName":  "视觉盒子",
        "category":  "半成品_畜牧",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  4,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "J0160003",
        "materialName":  "牛体刷总成 2.0（欧洲版）",
        "category":  "成品_畜牧",
        "productLine":  "智能牛体刷",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  12,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "SCEB200",
        "model":  "SCEB200"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "400000014J0320A",
        "materialName":  "主线束",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "QCD-DJ-003004A",
        "materialName":  "充电机",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "SVEA",
        "unit":  "个",
        "onHandQty":  7,
        "reservedQty":  5,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "充电机-100Vac to 245Vac-58.3V-58.8V-190 X 140 X 75mm-欧规 SB50A-600V IP65",
        "model":  "H82-48V10A"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "370000114J0260B",
        "materialName":  "雷达线",
        "category":  "零部件及主材_电器件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  5,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "美国芝加哥仓",
        "project":  "R1916",
        "materialCode":  "S010470502501",
        "materialName":  "带弹平垫盘头螺栓",
        "category":  "零部件及主材_标准件",
        "productLine":  "智能割草机配件",
        "location":  "SVEA HK",
        "unit":  "个",
        "onHandQty":  100,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "GB/T 9074.4 M5×25-A2-70-钝化72H",
        "model":  "GB/T 9074.4 M5×25-A2-70-钝化72h"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "J0080009",
        "materialName":  "智能项圈",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "库存差异",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "868",
        "model":  "SCN100"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1917",
        "materialCode":  "900280200J0080A",
        "materialName":  "LoRaWAN智能网关",
        "category":  "成品_畜牧",
        "productLine":  "智能项圈",
        "location":  "展会用品",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "ND67_868",
        "model":  "ND67"
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R1916",
        "materialCode":  "1R0000011J0260B",
        "materialName":  "胶皮",
        "category":  "零部件及主材_金属件",
        "productLine":  "推料机器人配件",
        "location":  "售后备件",
        "unit":  "个",
        "onHandQty":  10,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    },
    {
        "warehouseName":  "德国法兰克福仓（Germany-UEX）",
        "project":  "R2312",
        "materialCode":  "J0300014",
        "materialName":  "Sveaverken FeedLogger",
        "category":  "成品_畜牧",
        "productLine":  "畜牧-FJ05",
        "location":  "退货",
        "unit":  "个",
        "onHandQty":  1,
        "reservedQty":  0,
        "frozenQty":  0,
        "supplierOwnedQty":  0,
        "spec":  "",
        "model":  ""
    }
];

const views = ["overview", "material", "replenishment", "rma"];
const inventoryData = window.inventoryData || {};
const inventoryRows = Array.isArray(window.inventoryRows) ? window.inventoryRows : [];
const hasSyncedInventoryRows = inventoryRows.length > 0;
const afterSalesData = window.afterSalesData || { replenishmentOrders: [], rmaOrders: [] };
const dashboardRefreshConfig = window.dashboardRefreshConfig || {};
const afterSalesRefreshEndpoint = String(dashboardRefreshConfig.afterSalesEndpoint || "").trim();
let afterSalesGeneratedAt = String(afterSalesData.generatedAt || "");
let afterSalesRefreshMessage = "";
let afterSalesRefreshBusy = false;
const baseReplenishmentOrders = (afterSalesData.replenishmentOrders || []).map(normalizeReplenishmentOrder);
const baseRmaOrders = (afterSalesData.rmaOrders || []).map(normalizeRmaOrder);
const replenishmentOrders = [...baseReplenishmentOrders];
const rmaOrders = [...baseRmaOrders];
const baseRows = (inventoryRows.length ? inventoryRows : excelRows).map(normalizeRow);

let rows = [];
let dataOverrides = {
  inventory: false,
  replenishment: false,
  rma: false,
};
let state = {
  view: "overview",
  language: "zh-CN",
  keyword: "",
  warehouse: "all",
  project: "all",
  status: "all",
  sort: "availableAsc",
  overviewQuickFilter: "all",
  replenishmentKeyword: "",
  replenishmentWarehouse: "all",
  replenishmentStatus: "all",
  replenishmentTransport: "all",
  rmaKeyword: "",
  rmaStatus: "all",
  rmaCountry: "all",
  rmaProject: "all",
  rmaQuickFilter: "all",
  rmaSelectedNo: "",
};

const elements = {
  sourceNote: document.querySelector("#sourceNote"),
  languageSelect: document.querySelector("#languageSelect"),
  languagePicker: document.querySelector(".language-picker"),
  translationStatus: document.querySelector("#translationStatus"),
  viewTabs: document.querySelectorAll("[data-view]"),
  viewPanels: document.querySelectorAll("[data-view-panel]"),
  summaryGrid: document.querySelector("#summaryGrid"),
  keywordInput: document.querySelector("#keywordInput"),
  warehouseSelect: document.querySelector("#warehouseSelect"),
  projectSelect: document.querySelector("#projectSelect"),
  statusSelect: document.querySelector("#statusSelect"),
  sortSelect: document.querySelector("#sortSelect"),
  sampleButton: document.querySelector("#sampleButton"),
  exportButton: document.querySelector("#exportButton"),
  resetButton: document.querySelector("#resetButton"),
  materialLookupSection: document.querySelector("#materialLookupSection"),
  materialLookupTitle: document.querySelector("#materialLookupTitle"),
  materialLookupMeta: document.querySelector("#materialLookupMeta"),
  materialLookupCount: document.querySelector("#materialLookupCount"),
  materialLookupGrid: document.querySelector("#materialLookupGrid"),
  warehouseCount: document.querySelector("#warehouseCount"),
  warehouseList: document.querySelector("#warehouseList"),
  overviewDetailTitle: document.querySelector("#overviewDetailTitle"),
  overviewDetailMeta: document.querySelector("#overviewDetailMeta"),
  overviewDetailCount: document.querySelector("#overviewDetailCount"),
  overviewTableBody: document.querySelector("#overviewTableBody"),
  overviewEmptyState: document.querySelector("#overviewEmptyState"),
  networkStatus: document.querySelector("#networkStatus"),
  warehouseMap: document.querySelector("#warehouseMap"),
  alertCount: document.querySelector("#alertCount"),
  lowStockList: document.querySelector("#lowStockList"),
  transferList: document.querySelector("#transferList"),
  rowCount: document.querySelector("#rowCount"),
  stockTableBody: document.querySelector("#stockTableBody"),
  emptyState: document.querySelector("#emptyState"),
  importStatus: document.querySelector("#importStatus"),
  importArea: document.querySelector("#importArea"),
  importButton: document.querySelector("#importButton"),
  fillTemplateButton: document.querySelector("#fillTemplateButton"),
  overviewImportStatus: document.querySelector("#overviewImportStatus"),
  overviewImportArea: document.querySelector("#overviewImportArea"),
  overviewImportFile: document.querySelector("#overviewImportFile"),
  overviewImportButton: document.querySelector("#overviewImportButton"),
  overviewTemplateButton: document.querySelector("#overviewTemplateButton"),
  overviewRestoreButton: document.querySelector("#overviewRestoreButton"),
  odooCommandButton: document.querySelector("#odooCommandButton"),
  replenishmentCount: document.querySelector("#replenishmentCount"),
  replenishmentRefreshButton: document.querySelector("#replenishmentRefreshButton"),
  replenishmentStatus: document.querySelector("#replenishmentStatus"),
  replenishmentKeywordInput: document.querySelector("#replenishmentKeywordInput"),
  replenishmentWarehouseSelect: document.querySelector("#replenishmentWarehouseSelect"),
  replenishmentStatusSelect: document.querySelector("#replenishmentStatusSelect"),
  replenishmentTransportSelect: document.querySelector("#replenishmentTransportSelect"),
  replenishmentResetButton: document.querySelector("#replenishmentResetButton"),
  replenishmentWarehouseGrid: document.querySelector("#replenishmentWarehouseGrid"),
  replenishmentTableBody: document.querySelector("#replenishmentTableBody"),
  replenishmentEmptyState: document.querySelector("#replenishmentEmptyState"),
  replenishmentImportStatus: document.querySelector("#replenishmentImportStatus"),
  replenishmentImportArea: document.querySelector("#replenishmentImportArea"),
  replenishmentImportFile: document.querySelector("#replenishmentImportFile"),
  replenishmentImportButton: document.querySelector("#replenishmentImportButton"),
  replenishmentTemplateButton: document.querySelector("#replenishmentTemplateButton"),
  replenishmentRestoreButton: document.querySelector("#replenishmentRestoreButton"),
  rmaCount: document.querySelector("#rmaCount"),
  rmaSummaryGrid: document.querySelector("#rmaSummaryGrid"),
  rmaKeywordInput: document.querySelector("#rmaKeywordInput"),
  rmaStatusSelect: document.querySelector("#rmaStatusSelect"),
  rmaCountrySelect: document.querySelector("#rmaCountrySelect"),
  rmaProjectSelect: document.querySelector("#rmaProjectSelect"),
  rmaResetButton: document.querySelector("#rmaResetButton"),
  rmaStatus: document.querySelector("#rmaStatus"),
  rmaOrderTable: document.querySelector("#rmaOrderTable"),
  rmaTableBody: document.querySelector("#rmaTableBody"),
  rmaEmptyState: document.querySelector("#rmaEmptyState"),
  rmaDetailTable: document.querySelector("#rmaDetailTable"),
  rmaDetailTitle: document.querySelector("#rmaDetailTitle"),
  rmaDetailMeta: document.querySelector("#rmaDetailMeta"),
  rmaDetailCount: document.querySelector("#rmaDetailCount"),
  rmaDetailTableBody: document.querySelector("#rmaDetailTableBody"),
  rmaDetailEmptyState: document.querySelector("#rmaDetailEmptyState"),
  rmaImportStatus: document.querySelector("#rmaImportStatus"),
  rmaImportArea: document.querySelector("#rmaImportArea"),
  rmaImportFile: document.querySelector("#rmaImportFile"),
  rmaImportButton: document.querySelector("#rmaImportButton"),
  rmaTemplateButton: document.querySelector("#rmaTemplateButton"),
  rmaRestoreButton: document.querySelector("#rmaRestoreButton"),
  rmaRefreshButton: document.querySelector("#rmaRefreshButton"),
};

let translationRunId = 0;
let translationTimer = null;
const sourceTextByNode = new WeakMap();
const sourceAttributesByElement = new WeakMap();
const translationCache = loadTranslationCache();

function loadTranslationCache() {
  try {
    const parsed = JSON.parse(localStorage.getItem(TRANSLATE_CACHE_STORAGE_KEY) || "{}");
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function saveTranslationCache() {
  const entries = Object.entries(translationCache);
  const capped = entries.slice(Math.max(0, entries.length - 3000));
  localStorage.setItem(TRANSLATE_CACHE_STORAGE_KEY, JSON.stringify(Object.fromEntries(capped)));
}

function setTranslationStatus(message, type = "normal") {
  if (elements.translationStatus) elements.translationStatus.textContent = message;
  elements.languagePicker?.classList.toggle("is-busy", type === "busy");
  elements.languagePicker?.classList.toggle("is-error", type === "error");
}

function normalizeGoogleTargetLanguage(language) {
  return language === "zh-CN" ? "zh-CN" : language;
}

function restoreOriginalLanguage() {
  translationRunId += 1;
  collectTranslatableEntries(document.querySelector(".app-shell")).forEach((entry) => entry.restore());
  document.documentElement.lang = "zh-CN";
  setTranslationStatus("Google翻译");
}

function scheduleLanguageRender() {
  window.clearTimeout(translationTimer);
  translationTimer = window.setTimeout(() => {
    applyPageLanguage(state.language);
  }, 60);
}

function isTranslatableText(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed || trimmed.length < 2) return false;
  if (!/[\u4e00-\u9fff]/.test(trimmed)) return false;
  return !/^[\d\s.,:;+\-/()%#A-Za-z]+$/.test(trimmed);
}

function getTranslationRoots() {
  return [
    document.querySelector(".topbar"),
    document.querySelector(".view-tabs"),
    document.querySelector(`[data-view-panel="${state.view}"]`),
  ].filter(Boolean);
}

function collectTranslatableEntries(root = null) {
  const roots = root ? [root] : getTranslationRoots();
  return roots.flatMap((targetRoot) => [
    ...collectTranslatableTextNodes(targetRoot),
    ...collectTranslatableAttributeEntries(targetRoot),
  ]);
}

function collectTranslatableTextNodes(root = document.querySelector(".app-shell")) {
  if (!root) return [];
  const nodes = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      if (parent.closest("script, style, noscript, textarea, input, #languageSelect")) return NodeFilter.FILTER_REJECT;
      if (parent.closest("[data-no-translate]")) return NodeFilter.FILTER_REJECT;
      const text = node.nodeValue || "";
      const source = sourceTextByNode.get(node) || text;
      const trimmed = source.trim();
      if (!isTranslatableText(trimmed)) return NodeFilter.FILTER_REJECT;
      sourceTextByNode.set(node, source);
      nodes.push({
        source,
        text: trimmed,
        apply: (target) => applyTranslatedText(node, source, target),
        restore: () => {
          node.nodeValue = source;
        },
      });
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  while (walker.nextNode()) {
    // The acceptNode callback records accepted nodes.
  }
  return nodes;
}

function collectTranslatableAttributeEntries(root = document.querySelector(".app-shell")) {
  if (!root) return [];
  const entries = [];
  const attributes = ["placeholder", "aria-label", "alt"];
  const selector = attributes.map((attribute) => `[${attribute}]`).join(",");
  root.querySelectorAll(selector).forEach((element) => {
    if (element.closest("[data-no-translate], #languageSelect")) return;
    attributes.forEach((attribute) => {
      if (!element.hasAttribute(attribute)) return;
      const source = getSourceAttribute(element, attribute);
      const text = source.trim();
      if (!isTranslatableText(text)) return;
      entries.push({
        source,
        text,
        apply: (target) => element.setAttribute(attribute, target),
        restore: () => element.setAttribute(attribute, source),
      });
    });
  });
  return entries;
}

function getSourceAttribute(element, attribute) {
  let sourceMap = sourceAttributesByElement.get(element);
  if (!sourceMap) {
    sourceMap = new Map();
    sourceAttributesByElement.set(element, sourceMap);
  }
  if (!sourceMap.has(attribute)) {
    sourceMap.set(attribute, element.getAttribute(attribute) || "");
  }
  return sourceMap.get(attribute) || "";
}

function applyTranslatedText(node, source, target) {
  const prefix = (source.match(/^\s*/) || [""])[0];
  const suffix = (source.match(/\s*$/) || [""])[0];
  node.nodeValue = `${prefix}${target}${suffix}`;
}

function getTranslationCacheKey(language, text) {
  return `${language}\n${text}`;
}

async function translateTextBatch(texts, language) {
  if (!texts.length) return [];
  if (texts.length === 1) return [await translateTextWithGoogle(texts[0], language)];

  const mergedText = texts.join("\n");
  const mergedTranslation = await translateTextWithGoogle(mergedText, language);
  const parts = mergedTranslation
    .split(/\r?\n/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length === texts.length) return parts;

  const results = [];
  for (const text of texts) {
    results.push(await translateTextWithGoogle(text, language));
  }
  return results;
}

function createTranslationBatches(texts) {
  const batches = [];
  let current = [];
  let currentLength = 0;
  const maxBatchCount = 32;
  const maxBatchLength = 1600;

  texts.forEach((text) => {
    const length = text.length + 1;
    if (current.length && (current.length >= maxBatchCount || currentLength + length > maxBatchLength)) {
      batches.push(current);
      current = [];
      currentLength = 0;
    }
    current.push(text);
    currentLength += length;
  });

  if (current.length) batches.push(current);
  return batches;
}

async function translateTextWithGoogle(text, language) {
  const target = normalizeGoogleTargetLanguage(language);
  const params = new URLSearchParams({
    client: "gtx",
    sl: TRANSLATE_SOURCE_LANGUAGE,
    tl: target,
    dt: "t",
    q: text,
  });
  const response = await fetch(`${GOOGLE_TRANSLATE_ENDPOINT}?${params.toString()}`, {
    method: "GET",
  });
  if (!response.ok) {
    throw new Error(response.status === 429 ? "Google 翻译请求过多，请稍后再试" : `Google 翻译返回 ${response.status}`);
  }
  const payload = await response.json();
  return parseGoogleTranslateResponse(payload) || text;
}

function parseGoogleTranslateResponse(payload) {
  if (!Array.isArray(payload?.[0])) throw new Error("Google 翻译返回格式不正确");
  return payload[0]
    .map((segment) => (Array.isArray(segment) ? segment[0] || "" : ""))
    .join("")
    .trim();
}

async function applyPageLanguage(language) {
  if (!SUPPORTED_TRANSLATION_LANGUAGES.has(language)) return;
  if (elements.languageSelect && elements.languageSelect.value !== language) elements.languageSelect.value = language;
  if (language === "zh-CN") {
    restoreOriginalLanguage();
    return;
  }

  const runId = ++translationRunId;
  const entries = collectTranslatableEntries();
  const uniqueTexts = Array.from(new Set(entries.map((item) => item.text)));
  const missingTexts = uniqueTexts.filter((text) => !translationCache[getTranslationCacheKey(language, text)]);

  try {
    entries.forEach((entry) => {
      const translated = translationCache[getTranslationCacheKey(language, entry.text)];
      if (translated) entry.apply(translated);
    });

    setTranslationStatus(missingTexts.length ? "翻译中..." : "已缓存", missingTexts.length ? "busy" : "normal");
    const batches = createTranslationBatches(missingTexts);
    let translatedCount = 0;
    for (const batch of batches) {
      setTranslationStatus(`翻译中 ${Math.min(translatedCount + batch.length, missingTexts.length)}/${missingTexts.length}`, "busy");
      const translations = await translateTextBatch(batch, language);
      batch.forEach((text, batchIndex) => {
        translationCache[getTranslationCacheKey(language, text)] = translations[batchIndex] || text;
      });
      saveTranslationCache();
      if (runId !== translationRunId) return;
      translatedCount += batch.length;

      entries.forEach((entry) => {
        if (!batch.includes(entry.text)) return;
        const translated = translationCache[getTranslationCacheKey(language, entry.text)];
        if (translated) entry.apply(translated);
      });
    }

    entries.forEach((entry) => {
      const translated = translationCache[getTranslationCacheKey(language, entry.text)];
      if (translated) entry.apply(translated);
    });
    document.documentElement.lang = language;
    setTranslationStatus("Google已翻译");
  } catch (error) {
    console.error(error);
    setTranslationStatus("翻译失败", "error");
  }
}

function getSavedAfterSalesGeneratedAt(saved) {
  const sourceGeneratedAt = saved && saved.sourceGeneratedAt;
  return sourceGeneratedAt ? String(sourceGeneratedAt.afterSales || "") : "";
}

function shouldUseSavedAfterSalesData(saved) {
  const overrides = (saved && saved.dataOverrides) || {};
  if (!overrides.replenishment && !overrides.rma) return true;
  if (!afterSalesGeneratedAt) return true;
  return getSavedAfterSalesGeneratedAt(saved) === afterSalesGeneratedAt;
}

function resetAfterSalesFilters() {
  state.replenishmentKeyword = "";
  state.replenishmentWarehouse = "all";
  state.replenishmentStatus = "all";
  state.replenishmentTransport = "all";
  state.rmaKeyword = "";
  state.rmaStatus = "all";
  state.rmaCountry = "all";
  state.rmaProject = "all";
  state.rmaQuickFilter = "all";
  state.rmaSelectedNo = "";
}

function load() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    const useSavedAfterSalesData = shouldUseSavedAfterSalesData(saved);
    dataOverrides = {
      ...dataOverrides,
      ...(saved.dataOverrides || {}),
    };
    if (!useSavedAfterSalesData) {
      dataOverrides.replenishment = false;
      dataOverrides.rma = false;
    }
    rows = dataOverrides.inventory && Array.isArray(saved.rows) && saved.rows.length ? saved.rows.map(normalizeRow) : baseRows;
    if (dataOverrides.replenishment && Array.isArray(saved.replenishmentOrders)) {
      replaceCollection(replenishmentOrders, saved.replenishmentOrders.map(normalizeReplenishmentOrder));
    }
    if (dataOverrides.rma && Array.isArray(saved.rmaOrders)) {
      replaceCollection(rmaOrders, saved.rmaOrders.map(normalizeRmaOrder));
    }
    state = { ...state, ...(saved.state || {}) };
    if (!useSavedAfterSalesData) resetAfterSalesFilters();
    if (!views.includes(state.view)) state.view = "overview";
    if (!SUPPORTED_TRANSLATION_LANGUAGES.has(state.language)) state.language = "zh-CN";
    if (!["all", "lowStock"].includes(state.overviewQuickFilter)) state.overviewQuickFilter = "all";
    if (!getRmaQuickFilters([]).some((filter) => filter.id === state.rmaQuickFilter)) state.rmaQuickFilter = "all";
  } catch {
    rows = baseRows;
  }
}

function save() {
  const payload = {
    state,
    dataOverrides,
    sourceGeneratedAt: {
      afterSales: afterSalesGeneratedAt,
    },
  };
  if (dataOverrides.inventory) payload.rows = rows;
  if (dataOverrides.replenishment) payload.replenishmentOrders = replenishmentOrders;
  if (dataOverrides.rma) payload.rmaOrders = rmaOrders;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function replaceCollection(target, incomingRows) {
  target.splice(0, target.length, ...incomingRows);
}

function getAfterSalesDataUrl() {
  const url = new URL("data-orders.js", window.location.href);
  url.searchParams.set("refresh", String(Date.now()));
  return url.toString();
}

function parseAfterSalesDataScript(scriptText) {
  const jsonText = String(scriptText || "")
    .replace(/^\s*window\.afterSalesData\s*=\s*/, "")
    .replace(/;\s*$/, "");
  return JSON.parse(jsonText);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function triggerAfterSalesRefreshEndpoint() {
  if (!afterSalesRefreshEndpoint) return false;
  const response = await fetch(afterSalesRefreshEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ scope: "afterSales", requestedAt: new Date().toISOString() }),
  });
  if (!response.ok) throw new Error(`同步接口返回 ${response.status}`);
  return true;
}

async function fetchLatestAfterSalesData() {
  const response = await fetch(getAfterSalesDataUrl(), { cache: "no-store" });
  if (!response.ok) throw new Error(`数据文件返回 ${response.status}`);
  return parseAfterSalesDataScript(await response.text());
}

async function fetchLatestAfterSalesDataAfterSync(previousGeneratedAt) {
  let latestData = await fetchLatestAfterSalesData();
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const latestGeneratedAt = String(latestData.generatedAt || "");
    if (!previousGeneratedAt || latestGeneratedAt !== previousGeneratedAt) return latestData;
    await sleep(3000);
    latestData = await fetchLatestAfterSalesData();
  }
  return latestData;
}

function applyLatestAfterSalesData(data) {
  const previousGeneratedAt = afterSalesGeneratedAt;
  const nextReplenishmentOrders = (data.replenishmentOrders || []).map(normalizeReplenishmentOrder);
  const nextRmaOrders = (data.rmaOrders || []).map(normalizeRmaOrder);
  afterSalesGeneratedAt = String(data.generatedAt || "");
  window.afterSalesData = data;
  replaceCollection(baseReplenishmentOrders, nextReplenishmentOrders);
  replaceCollection(baseRmaOrders, nextRmaOrders);
  replaceCollection(replenishmentOrders, nextReplenishmentOrders);
  replaceCollection(rmaOrders, nextRmaOrders);
  dataOverrides.replenishment = false;
  dataOverrides.rma = false;
  return {
    changed: previousGeneratedAt !== afterSalesGeneratedAt,
    replenishmentRows: nextReplenishmentOrders.length,
    rmaRows: nextRmaOrders.length,
  };
}

function renderAfterSalesRefreshControls() {
  const buttons = [elements.replenishmentRefreshButton, elements.rmaRefreshButton].filter(Boolean);
  buttons.forEach((button) => {
    button.disabled = afterSalesRefreshBusy;
    button.textContent = afterSalesRefreshBusy ? "刷新中..." : "刷新数据";
    button.classList.toggle("is-loading", afterSalesRefreshBusy);
    button.title = afterSalesRefreshEndpoint ? "从多维表格同步最新数据" : "刷新已发布的多维表格数据";
  });
}

async function refreshAfterSalesFromSource() {
  if (afterSalesRefreshBusy) return;
  const previousGeneratedAt = afterSalesGeneratedAt;
  const hadManualOverrides = dataOverrides.replenishment || dataOverrides.rma;
  afterSalesRefreshBusy = true;
  afterSalesRefreshMessage = "正在刷新多维表格数据...";
  render();
  try {
    const triggeredRemoteSync = await triggerAfterSalesRefreshEndpoint();
    const latestData = triggeredRemoteSync
      ? await fetchLatestAfterSalesDataAfterSync(previousGeneratedAt)
      : await fetchLatestAfterSalesData();
    const result = applyLatestAfterSalesData(latestData);
    resetAfterSalesFilters();
    const sourceLabel = triggeredRemoteSync ? "已同步多维表格" : "已刷新线上发布数据";
    if (result.changed || hadManualOverrides) {
      afterSalesRefreshMessage = `${sourceLabel}：${afterSalesGeneratedAt || "-"} / 备货 ${result.replenishmentRows} 条 / RMA ${result.rmaRows} 条`;
    } else {
      afterSalesRefreshMessage = `已是最新发布数据：${afterSalesGeneratedAt || "-"} / 备货 ${result.replenishmentRows} 条 / RMA ${result.rmaRows} 条`;
    }
  } catch (error) {
    console.error(error);
    afterSalesRefreshMessage = "刷新失败，请稍后重试或等待定时同步";
  } finally {
    afterSalesRefreshBusy = false;
    save();
    render();
  }
}

function getWarehouseLookupValue(row) {
  const warehouseId = String(row.warehouseId || "").trim();
  if (warehouseId && warehouseId !== "unknown") return warehouseId;
  return row.warehouse || row.warehouseName || warehouseId;
}

function normalizeProject(value) {
  const project = String(value || "").trim();
  if (!project) return "未标注";
  return project.toUpperCase() === "MRP" ? "R1916" : project;
}

function normalizeRow(row) {
  const warehouse = findWarehouse(getWarehouseLookupValue(row));

  return {
    warehouseId: warehouse.id,
    warehouseName: String(row.warehouseName || row.warehouse || warehouse.name).trim(),
    project: normalizeProject(row.project),
    materialCode: String(row.materialCode || "").trim(),
    materialName: String(row.materialName || "").trim(),
    category: String(row.category || "").trim(),
    productLine: String(row.productLine || "").trim(),
    location: String(row.location || "").trim(),
    unit: String(row.unit || "").trim() || "个",
    onHandQty: parseNumber(row.onHandQty),
    reservedQty: parseNumber(row.reservedQty),
    frozenQty: parseNumber(row.frozenQty),
    supplierOwnedQty: parseNumber(row.supplierOwnedQty),
    unitCost: parseNumber(row.unitCost || row.standardPrice || row.cost || row.price),
    inventoryAmount: parseNumber(row.inventoryAmount || row.amount || row.value),
    priceSource: String(row.priceSource || "").trim(),
    priceStartPlace: String(row.priceStartPlace || "").trim(),
    spec: String(row.spec || "").trim(),
    model: String(row.model || "").trim(),
    productId: String(row.productId || "").trim(),
    imageUrl: String(row.imageUrl || row.image || "").trim(),
    updatedAt: row.updatedAt || new Date().toISOString(),
  };
}

function normalizeReplenishmentOrder(order) {
  const warehouseLookup = order.warehouseId && order.warehouseId !== "unknown" ? order.warehouseId : order.warehouseText || order.warehouse;
  const warehouse = findWarehouse(warehouseLookup);
  return {
    orderNo: String(order.orderNo || "").trim(),
    materialCode: String(order.materialCode || "").trim(),
    materialName: String(order.materialName || "").trim(),
    qty: parseNumber(order.qty),
    warehouseId: warehouse.id,
    warehouseText: String(order.warehouseText || warehouse.name).trim(),
    unitPrice: String(order.unitPrice || "").trim(),
    totalPrice: String(order.totalPrice || "").trim(),
    shippedDate: String(order.shippedDate || "").trim(),
    expectedArrival: String(order.expectedArrival || "").trim(),
    leadTime: String(order.leadTime || "").trim(),
    transport: String(order.transport || "未标注").trim(),
    status: String(order.status || "待确认").trim(),
    inboundFlag: String(order.inboundFlag || "").trim(),
  };
}

function normalizeRmaOrder(order) {
  return {
    rmaNo: String(order.rmaNo || "").trim(),
    serviceNo: String(order.serviceNo || "").trim(),
    sn: String(order.sn || "").trim(),
    customer: String(order.customer || "").trim(),
    productLine: String(order.productLine || "").trim(),
    country: String(order.country || "").trim(),
    model: String(order.model || "").trim(),
    project: String(order.project || "").trim(),
    materialCode: String(order.materialCode || "").trim(),
    materialName: String(order.materialName || "").trim(),
    qty: parseNumber(order.qty),
    unitPrice: String(order.unitPrice || "").trim(),
    totalUsd: String(order.totalUsd || "").trim(),
    origin: String(order.origin || "").trim(),
    shippedDate: String(order.shippedDate || "").trim(),
    status: String(order.status || "待确认").trim(),
    afterSalesType: String(order.afterSalesType || "").trim(),
    specialNote: String(order.specialNote || "").trim(),
    reason: String(order.reason || "").trim(),
    fae: String(order.fae || "").trim(),
    badPartSuggestion: String(order.badPartSuggestion || "").trim(),
    trackingNo: String(order.trackingNo || "").trim(),
    logisticsDays: String(order.logisticsDays || "").trim(),
    parentRecord: String(order.parentRecord || "").trim(),
  };
}

function findWarehouse(value) {
  const rawText = String(value || "").trim();
  const text = rawText.toLowerCase();
  const matched = getKnownWarehouses().find((warehouse) => {
    const haystack = [warehouse.id, warehouse.code, warehouse.name, warehouse.sourceName, warehouse.city].join(" ").toLowerCase();
    return haystack.includes(text) || text.includes(warehouse.name.toLowerCase()) || text.includes(warehouse.sourceName.toLowerCase());
  });

  return matched || createDynamicWarehouse(makeWarehouseId(rawText), rawText || "未匹配仓库");
}

function getKnownWarehouses() {
  return [...warehouses, ...domesticWarehouses];
}

function makeWarehouseId(value) {
  const text = String(value || "未匹配仓库").trim().toLowerCase();
  const ascii = text.replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  if (ascii) return `odoo-${ascii}`;

  let hash = 0;
  for (let index = 0; index < text.length; index += 1) {
    hash = (hash * 31 + text.charCodeAt(index)) | 0;
  }
  return `odoo-${Math.abs(hash).toString(36)}`;
}

function createDynamicWarehouse(id, name, index = 0) {
  const positions = [
    { x: "38%", y: "70%", mobileY: "84%" },
    { x: "54%", y: "72%", mobileY: "88%" },
    { x: "70%", y: "72%", mobileY: "92%" },
  ];
  const position = positions[index % positions.length];
  const label = String(name || id || "未匹配仓库").trim();

  return {
    id,
    code: "OD",
    name: label,
    sourceName: label,
    city: "Odoo 库位",
    role: "未匹配到固定海外仓",
    accent: "#5f6f7e",
    soft: "#eef2f6",
    ...position,
  };
}

function getActiveWarehouses(sourceRows = rows, options = {}) {
  const knownWarehouses = options.includeDomestic === false ? warehouses : getKnownWarehouses();
  const map = new Map(knownWarehouses.map((warehouse) => [warehouse.id, warehouse]));

  sourceRows.forEach((row) => {
    if (!row.warehouseId || map.has(row.warehouseId)) return;
    map.set(row.warehouseId, createDynamicWarehouse(row.warehouseId, row.warehouseName || row.warehouseId, map.size));
  });

  return Array.from(map.values());
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const normalized = String(value).replace(/,/g, "").replace(/[^\d.-]/g, "");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function availableQty(row) {
  return row.onHandQty - row.reservedQty - row.frozenQty;
}

function inventoryAmount(row) {
  return row.inventoryAmount || row.onHandQty * row.unitCost;
}

function isInInventoryProjectScope(row) {
  return inventoryProjectScope.has(String(row.project || "").trim());
}

function isInInventoryLocationScope(row) {
  const location = String(row.location || "").trim();
  if (!location) return false;
  const segments = location
    .split("/")
    .map((part) => part.trim())
    .filter(Boolean);
  const values = segments.length ? segments : [location];

  return syncLocationScope.some((allowed) => {
    const value = String(allowed || "").trim();
    if (!value) return false;
    return values.some((segment) => {
      if (segment.toLowerCase() === value.toLowerCase()) return true;
      return value.toLowerCase() === "svea" && segment.toLowerCase().startsWith("svea");
    });
  });
}

function getVisibleInventoryRows() {
  return getOverviewInventoryRows();
}

function hasInventory(row) {
  return row.onHandQty > 0 || row.reservedQty > 0 || row.frozenQty > 0 || availableQty(row) > 0;
}

function isOverseasWarehouse(row) {
  return warehouses.some((warehouse) => warehouse.id === row.warehouseId);
}

function getOverviewInventoryRows() {
  return rows.filter((row) => isInInventoryProjectScope(row) && isOverseasWarehouse(row) && isInInventoryLocationScope(row) && hasInventory(row));
}

function getMaterialInventoryRows() {
  return rows.filter((row) => isInInventoryProjectScope(row) && hasInventory(row));
}

function rowStatus(row) {
  const available = availableQty(row);
  if (available <= 0) return "shortage";
  if (row.frozenQty > 0) return "frozen";
  if (row.reservedQty > 0) return "reserved";
  return "normal";
}

function statusLabel(status) {
  return {
    normal: "正常",
    reserved: "有预留",
    frozen: "有冻结",
    shortage: "无可用",
  }[status];
}

function formatQty(value) {
  return Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

function formatMoney(value) {
  return Number(value || 0).toLocaleString("zh-CN", {
    maximumFractionDigits: 2,
    minimumFractionDigits: 0,
  });
}

function getWarehouseName(id) {
  return getActiveWarehouses().find((warehouse) => warehouse.id === id)?.name || id;
}

function getProjects() {
  return Array.from(new Set(getMaterialInventoryRows().map((row) => row.project).filter(Boolean))).sort((a, b) => a.localeCompare(b, "zh-CN"));
}

function getFilteredRows() {
  const keyword = state.keyword.trim().toLowerCase();

  return getMaterialInventoryRows()
    .filter((row) => {
      const warehouse = getWarehouseName(row.warehouseId);
      const text = [
        warehouse,
        row.project,
        row.materialCode,
        row.materialName,
        row.category,
        row.productLine,
        row.location,
        row.spec,
        row.model,
      ]
        .join(" ")
        .toLowerCase();
      const matchesKeyword = !keyword || text.includes(keyword);
      const matchesWarehouse = state.warehouse === "all" || row.warehouseId === state.warehouse;
      const matchesProject = state.project === "all" || row.project === state.project;
      const matchesStatus = state.status === "all" || rowStatus(row) === state.status;
      return matchesKeyword && matchesWarehouse && matchesProject && matchesStatus;
    })
    .sort(sortRows);
}

function sortRows(a, b) {
  if (state.sort === "qtyDesc") return b.onHandQty - a.onHandQty;
  if (state.sort === "warehouseAsc") {
    const warehouseDiff = warehouseIndex(a.warehouseId) - warehouseIndex(b.warehouseId);
    return warehouseDiff || a.materialCode.localeCompare(b.materialCode, "zh-CN");
  }
  if (state.sort === "materialAsc") return a.materialCode.localeCompare(b.materialCode, "zh-CN");
  return availableQty(a) - availableQty(b) || b.onHandQty - a.onHandQty;
}

function warehouseIndex(id) {
  return getActiveWarehouses().findIndex((warehouse) => warehouse.id === id);
}

function aggregateRows(sourceRows, options = {}) {
  const warehouseList = options.warehouses || getActiveWarehouses(sourceRows);
  const includeEmpty = Boolean(options.includeEmpty);
  const base = new Map(
    warehouseList.map((warehouse) => [
      warehouse.id,
      {
        warehouse,
        materialCodes: new Set(),
        projects: new Set(),
        productLines: new Set(),
        onHandQty: 0,
        reservedQty: 0,
        frozenQty: 0,
        availableQty: 0,
        inventoryAmount: 0,
        supplierOwnedQty: 0,
        shortageCount: 0,
        lowStockCount: 0,
        reservedCount: 0,
      },
    ]),
  );

  sourceRows.forEach((row) => {
    const item = base.get(row.warehouseId);
    if (!item) return;
    item.materialCodes.add(row.materialCode);
    item.projects.add(row.project);
    item.productLines.add(row.productLine);
    item.onHandQty += row.onHandQty;
    item.reservedQty += row.reservedQty;
    item.frozenQty += row.frozenQty;
    item.availableQty += availableQty(row);
    item.inventoryAmount += inventoryAmount(row);
    item.supplierOwnedQty += row.supplierOwnedQty;
    if (rowStatus(row) === "shortage") item.shortageCount += 1;
    if (row.onHandQty < 5) item.lowStockCount += 1;
    if (rowStatus(row) === "reserved") item.reservedCount += 1;
  });

  return Array.from(base.values()).filter((item) => includeEmpty || item.onHandQty || item.materialCodes.size);
}

function render() {
  renderView();
  renderControls();
  renderReplenishmentControls();
  renderRmaControls();
  renderAfterSalesRefreshControls();
  const visibleInventoryRows = getVisibleInventoryRows();
  const filteredRows = getFilteredRows();
  const allWarehouseStats = aggregateRows(visibleInventoryRows, { warehouses, includeEmpty: true });
  const overviewRows = getOverviewRows();
  renderSummary(overviewRows);
  renderMaterialLookup();
  renderWarehouses(allWarehouseStats);
  renderOverviewDetail(overviewRows);
  renderNetwork(allWarehouseStats);
  renderInsights();
  renderTable(filteredRows);
  renderReplenishment();
  renderRma();
  if (elements.sourceNote) {
    elements.sourceNote.textContent = "";
    elements.sourceNote.hidden = true;
  }
  if (elements.languageSelect) elements.languageSelect.value = state.language;
  scheduleLanguageRender();
}

function renderView() {
  const activeView = views.includes(state.view) ? state.view : "overview";
  state.view = activeView;

  elements.viewTabs.forEach((tab) => {
    const isActive = tab.dataset.view === activeView;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", String(isActive));
  });

  elements.viewPanels.forEach((panel) => {
    const isActive = panel.dataset.viewPanel === activeView;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });
}

function getOverviewRows() {
  const sourceRows = getVisibleInventoryRows();
  const selectedOverseasWarehouse = warehouses.some((warehouse) => warehouse.id === state.warehouse);
  const warehouseRows = state.warehouse === "all" || !selectedOverseasWarehouse ? sourceRows : sourceRows.filter((row) => row.warehouseId === state.warehouse);
  const filteredRows = state.overviewQuickFilter === "lowStock" ? warehouseRows.filter((row) => row.onHandQty < 5) : warehouseRows;
  return [...filteredRows].sort(sortRows);
}

function normalizeSearchCode(value) {
  return String(value || "").trim().replace(/\s+/g, "").toLowerCase();
}

function getMaterialLookupMatches() {
  const keyword = normalizeSearchCode(state.keyword);
  if (!keyword) return [];

  return getMaterialInventoryRows().filter((row) => row.materialCode.toLowerCase().includes(keyword));
}

function renderMaterialLookup() {
  const matches = getMaterialLookupMatches();
  if (!matches.length) {
    elements.materialLookupSection.hidden = true;
    elements.materialLookupGrid.innerHTML = "";
    return;
  }

  const keyword = normalizeSearchCode(state.keyword);
  const exactRows = matches.filter((row) => row.materialCode.toLowerCase() === keyword);
  const lookupRows = exactRows.length ? exactRows : matches;
  const codes = Array.from(new Set(lookupRows.map((row) => row.materialCode))).sort((a, b) => a.localeCompare(b, "zh-CN"));
  const names = Array.from(new Set(lookupRows.map((row) => row.materialName).filter(Boolean))).slice(0, 3);
  const coveredWarehouseCount = new Set(lookupRows.map((row) => row.warehouseId)).size;
  const totalOnHand = lookupRows.reduce((sum, row) => sum + row.onHandQty, 0);
  const totalAvailable = lookupRows.reduce((sum, row) => sum + availableQty(row), 0);
  const totalAmount = lookupRows.reduce((sum, row) => sum + inventoryAmount(row), 0);

  elements.materialLookupSection.hidden = false;
  elements.materialLookupTitle.textContent = exactRows.length ? `物料编码查询：${codes[0]}` : `物料编码查询：${state.keyword.trim()}`;
  elements.materialLookupMeta.textContent = `${names.join(" / ") || "匹配物料"}；匹配 ${lookupRows.length} 条明细，覆盖 ${coveredWarehouseCount} 个仓库。下方按全部仓库展示，未存放的仓库显示为 0。`;
  elements.materialLookupCount.textContent = `现存 ${formatQty(totalOnHand)} / 可用 ${formatQty(totalAvailable)} / 金额 ${formatMoney(totalAmount)}`;

  elements.materialLookupGrid.innerHTML = [
    buildMaterialImageCard(lookupRows, codes[0], names[0]),
    ...getActiveWarehouses(lookupRows)
    .map((warehouse) => {
      const warehouseRows = lookupRows.filter((row) => row.warehouseId === warehouse.id);
      const onHand = warehouseRows.reduce((sum, row) => sum + row.onHandQty, 0);
      const reserved = warehouseRows.reduce((sum, row) => sum + row.reservedQty, 0);
      const frozen = warehouseRows.reduce((sum, row) => sum + row.frozenQty, 0);
      const available = onHand - reserved - frozen;
      const amount = warehouseRows.reduce((sum, row) => sum + inventoryAmount(row), 0);
      const locations = Array.from(new Set(warehouseRows.map((row) => row.location).filter(Boolean))).join("、") || "无库存";
      const projects = Array.from(new Set(warehouseRows.map((row) => row.project).filter(Boolean))).join(" / ") || "-";
      const hasStockClass = onHand > 0 || reserved > 0 || frozen > 0 ? " has-stock" : "";
      const transitInfo = warehouse.id === "fuzhou" ? null : getMaterialTransitInfo(codes, warehouse.id);

      return `
        <article class="material-warehouse-card${hasStockClass}" style="--accent: ${warehouse.accent}; --soft: ${warehouse.soft}">
          <div class="material-warehouse-head">
            <span class="warehouse-code">${escapeHtml(warehouse.code)}</span>
            <div>
              <strong>${escapeHtml(warehouse.name)}</strong>
              <small>${escapeHtml(projects)}</small>
            </div>
          </div>
          <div class="material-qty-grid">
            <span><b>${formatQty(onHand)}</b><small>现存</small></span>
            <span><b>${formatQty(reserved)}</b><small>预留</small></span>
            <span><b>${formatQty(frozen)}</b><small>冻结</small></span>
            <span><b>${formatQty(available)}</b><small>可用</small></span>
          </div>
          <p class="amount-line">库存金额 ${formatMoney(amount)}</p>
          ${
            transitInfo
              ? `<div class="transit-line">
                  <strong>在途 ${formatQty(transitInfo.qty)}</strong>
                  <span>${escapeHtml(transitInfo.summary)}</span>
                </div>`
              : warehouse.id !== "fuzhou"
                ? '<div class="transit-line muted-transit"><strong>在途 0</strong><span>暂无未入库备货</span></div>'
                : ""
          }
          <p>${escapeHtml(locations)}</p>
        </article>
      `;
    }),
  ].join("");
}

function getMaterialTransitInfo(materialCodes, warehouseId) {
  const codeSet = new Set(materialCodes.map((code) => String(code || "").trim().toLowerCase()).filter(Boolean));
  if (!codeSet.size) return null;

  const openOrders = replenishmentOrders.filter((order) => {
    if (order.warehouseId !== warehouseId) return false;
    if (!codeSet.has(String(order.materialCode || "").trim().toLowerCase())) return false;
    const statusText = `${order.status || ""} ${inboundLabel(order.inboundFlag)}`;
    if (String(order.inboundFlag || "").trim() === "1") return false;
    if (/已入库|已完成|完成/.test(statusText)) return false;
    return true;
  });
  if (!openOrders.length) return null;

  const qty = openOrders.reduce((sum, order) => sum + parseNumber(order.qty), 0);
  const orderNos = Array.from(new Set(openOrders.map((order) => order.orderNo).filter(Boolean)));
  const statuses = Array.from(new Set(openOrders.map((order) => order.status).filter(Boolean)));
  const arrivals = Array.from(new Set(openOrders.map((order) => order.expectedArrival).filter(Boolean))).slice(0, 2);
  const summaryParts = [];
  summaryParts.push(`${orderNos.length || openOrders.length} 个订单`);
  if (statuses.length) summaryParts.push(statuses.slice(0, 2).join(" / "));
  if (arrivals.length) summaryParts.push(`预计 ${arrivals.join(" / ")}`);
  return {
    qty,
    summary: summaryParts.join("，"),
  };
}

function buildMaterialImageCard(lookupRows, materialCode, materialName) {
  const imageUrl = lookupRows.map((row) => row.imageUrl).find(Boolean);
  const initials = String(materialName || materialCode || "SKU").slice(0, 2).toUpperCase();
  return `
    <article class="material-image-card">
      <div class="material-image-frame">
        ${
          imageUrl
            ? `<img src="${escapeHtml(imageUrl)}" alt="${escapeHtml(materialName || materialCode || "产品实物图")}" loading="lazy">`
            : `<span>${escapeHtml(initials)}</span>`
        }
      </div>
      <strong>${escapeHtml(materialCode || "-")}</strong>
      <small>${escapeHtml(materialName || "暂无产品图片")}</small>
    </article>
  `;
}

function renderControls() {
  elements.keywordInput.value = state.keyword;
  elements.projectSelect.innerHTML = [
    '<option value="all">全部项目</option>',
    ...getProjects().map((project) => `<option value="${escapeHtml(project)}">${escapeHtml(project)}</option>`),
  ].join("");
  elements.projectSelect.value = getProjects().includes(state.project) ? state.project : "all";
  elements.statusSelect.value = state.status;
  elements.sortSelect.value = state.sort;
  elements.warehouseSelect.innerHTML = [
    '<option value="all">全部仓库</option>',
    ...getActiveWarehouses(getMaterialInventoryRows()).map((warehouse) => `<option value="${escapeHtml(warehouse.id)}">${escapeHtml(warehouse.name)}</option>`),
  ].join("");
  elements.warehouseSelect.value = getActiveWarehouses(getMaterialInventoryRows()).some((warehouse) => warehouse.id === state.warehouse) ? state.warehouse : "all";
}

function renderReplenishmentControls() {
  elements.replenishmentKeywordInput.value = state.replenishmentKeyword;
  elements.replenishmentWarehouseSelect.innerHTML = [
    '<option value="all">全部海外仓</option>',
    ...warehouses.map((warehouse) => `<option value="${warehouse.id}">${escapeHtml(warehouse.name)}</option>`),
  ].join("");
  elements.replenishmentWarehouseSelect.value = warehouses.some((warehouse) => warehouse.id === state.replenishmentWarehouse)
    ? state.replenishmentWarehouse
    : "all";

  elements.replenishmentStatusSelect.innerHTML = buildOptions("全部状态", uniqueValues(replenishmentOrders, "status"));
  elements.replenishmentStatusSelect.value = uniqueValues(replenishmentOrders, "status").includes(state.replenishmentStatus)
    ? state.replenishmentStatus
    : "all";
  elements.replenishmentTransportSelect.innerHTML = buildOptions("全部运输方式", uniqueValues(replenishmentOrders, "transport"));
  elements.replenishmentTransportSelect.value = uniqueValues(replenishmentOrders, "transport").includes(state.replenishmentTransport)
    ? state.replenishmentTransport
    : "all";
}

function renderRmaControls() {
  elements.rmaKeywordInput.value = state.rmaKeyword;
  elements.rmaStatusSelect.innerHTML = buildOptions("全部状态", uniqueValues(rmaOrders, "status"));
  elements.rmaStatusSelect.value = uniqueValues(rmaOrders, "status").includes(state.rmaStatus) ? state.rmaStatus : "all";
  elements.rmaCountrySelect.innerHTML = buildOptions("全部国家", uniqueValues(rmaOrders, "country"));
  elements.rmaCountrySelect.value = uniqueValues(rmaOrders, "country").includes(state.rmaCountry) ? state.rmaCountry : "all";
  elements.rmaProjectSelect.innerHTML = buildOptions("全部项目", uniqueValues(rmaOrders, "project"));
  elements.rmaProjectSelect.value = uniqueValues(rmaOrders, "project").includes(state.rmaProject) ? state.rmaProject : "all";
}

function buildOptions(allLabel, values) {
  return [
    `<option value="all">${escapeHtml(allLabel)}</option>`,
    ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`),
  ].join("");
}

function uniqueValues(sourceRows, key) {
  return Array.from(new Set(sourceRows.map((row) => row[key]).filter(Boolean))).sort((a, b) => a.localeCompare(b, "zh-CN"));
}

function renderSummary(filteredRows) {
  const activeWarehouseTotal = warehouses.length;
  const warehouseCount = activeWarehouseTotal;
  const skuCount = new Set(filteredRows.map((row) => row.materialCode)).size;
  const totalOnHand = filteredRows.reduce((sum, row) => sum + row.onHandQty, 0);
  const totalAvailable = filteredRows.reduce((sum, row) => sum + availableQty(row), 0);
  const totalAmount = filteredRows.reduce((sum, row) => sum + inventoryAmount(row), 0);
  const lowStockRows = filteredRows.filter((row) => row.onHandQty < 5).length;

  const cards = [
    { label: "覆盖海外仓", value: `${warehouseCount} / ${activeWarehouseTotal}`, note: "固定展示 4 个海外仓" },
    { label: "物料 SKU", value: `${skuCount}`, note: getProjects().join(" / ") },
    { label: "现存总量", value: formatQty(totalOnHand), note: "包含预留数量" },
    { label: "可用库存", value: formatQty(totalAvailable), note: "现存减预留减冻结" },
    { label: "库存金额", value: formatMoney(totalAmount), note: "按库存数量与单价计算" },
    { label: "预留关注", value: `${lowStockRows}`, note: "现存量低于 5 的备件，点击查看", filter: "lowStock" },
  ];

  elements.summaryGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="summary-card${card.filter ? ` summary-action-card${state.overviewQuickFilter === card.filter ? " active" : ""}` : ""}"${card.filter ? ` data-overview-filter="${escapeHtml(card.filter)}" role="button" tabindex="0"` : ""}>
          <span>${escapeHtml(card.label)}</span>
          <strong>${escapeHtml(card.value)}</strong>
          <small>${escapeHtml(card.note)}</small>
        </article>
      `,
    )
    .join("");
}

function renderWarehouses(stats) {
  const maxQty = Math.max(...stats.map((item) => item.onHandQty), 1);
  elements.warehouseCount.textContent = `${stats.length} 个海外仓`;
  elements.warehouseList.innerHTML = stats
    .map((item) => {
      const warehouse = item.warehouse;
      const width = Math.max(2, Math.round((item.onHandQty / maxQty) * 100));
      const activeClass = state.warehouse === warehouse.id ? " active" : "";
      const emptyClass = item.materialCodes.size ? "" : " empty";
      const projectText = Array.from(item.projects).sort().join(" / ") || "-";
      return `
        <article class="warehouse-card${activeClass}${emptyClass}" data-warehouse="${warehouse.id}" style="--accent: ${warehouse.accent}; --soft: ${warehouse.soft}; --width: ${width}%">
          <div class="warehouse-head">
            <div>
              <div class="warehouse-name">${escapeHtml(warehouse.name)}</div>
              <div class="warehouse-location">${escapeHtml(warehouse.city)} / ${escapeHtml(warehouse.role)}</div>
            </div>
            <span class="warehouse-code">${escapeHtml(warehouse.code)}</span>
          </div>
          <div class="warehouse-meta">
            <span>现存 ${formatQty(item.onHandQty)}</span>
            <span>可用 ${formatQty(item.availableQty)}</span>
            <span>金额 ${formatMoney(item.inventoryAmount)}</span>
            <span>SKU ${item.materialCodes.size}</span>
            <span>项目 ${escapeHtml(projectText)}</span>
          </div>
          <div class="bar-track"><span class="bar"></span></div>
          <div class="warehouse-meta">
            <span>预留 ${formatQty(item.reservedQty)}</span>
            <span>冻结 ${formatQty(item.frozenQty)}</span>
            <span>无可用 ${item.shortageCount}</span>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderOverviewDetail(overviewRows) {
  const activeWarehouse = warehouses.find((warehouse) => warehouse.id === state.warehouse);
  const warehouseNames = new Set(overviewRows.map((row) => row.warehouseId));
  const skuCount = new Set(overviewRows.map((row) => row.materialCode)).size;
  const available = overviewRows.reduce((sum, row) => sum + availableQty(row), 0);

  const quickFilterLabel = state.overviewQuickFilter === "lowStock" ? "低库存备件" : "库存明细";
  elements.overviewDetailTitle.textContent = activeWarehouse ? `${activeWarehouse.name}${quickFilterLabel}` : `全部海外仓${quickFilterLabel}`;
  elements.overviewDetailMeta.textContent = activeWarehouse
    ? `${activeWarehouse.city} / ${activeWarehouse.role}`
    : `覆盖 ${warehouseNames.size} 个海外仓，按当前库存数据展示`;
  elements.overviewDetailCount.textContent = `${overviewRows.length} 条 / SKU ${skuCount} / 可用 ${formatQty(available)}`;
  elements.overviewEmptyState.hidden = overviewRows.length > 0;
  elements.overviewTableBody.innerHTML = buildOverviewRowsHtml(overviewRows);
}

function renderNetwork(stats) {
  const alertWarehouseCount = stats.filter((item) => item.lowStockCount > 0).length;
  elements.networkStatus.textContent = `${alertWarehouseCount} 个仓有低库存或关注项`;
  elements.warehouseMap.innerHTML = [
    '<span class="map-route" aria-hidden="true"></span>',
    ...stats.map((item) => {
      const warehouse = item.warehouse;
      const risk = item.materialCodes.size === 0 ? "暂无库存" : item.lowStockCount ? `低库存 ${item.lowStockCount} 项` : "库存正常";
      return `
        <article class="map-node" style="--x: ${warehouse.x}; --y: ${warehouse.y}; --mobile-y: ${warehouse.mobileY}; --accent: ${warehouse.accent}">
          <span class="map-dot" aria-hidden="true"></span>
          <strong>${escapeHtml(warehouse.name)}</strong>
          <span>${escapeHtml(risk)} / 可用 ${formatQty(item.availableQty)}</span>
        </article>
      `;
    }),
  ].join("");
}

function renderInsights() {
  const alertRows = getVisibleInventoryRows()
    .filter((row) => row.onHandQty < 5)
    .sort((a, b) => a.onHandQty - b.onHandQty || availableQty(a) - availableQty(b))
    .slice(0, 6);

  elements.alertCount.textContent = `${alertRows.length} 条重点`;
  elements.lowStockList.innerHTML = alertRows.length
    ? alertRows
        .map(
          (row) => `
            <article class="compact-item clickable-compact-item" data-overview-filter="lowStock" title="点击查看低库存明细">
              <strong>${escapeHtml(row.project)} / ${escapeHtml(row.materialCode)} / ${escapeHtml(row.materialName)}</strong>
              <span>${escapeHtml(getWarehouseName(row.warehouseId))} / 现存 ${formatQty(row.onHandQty)} / 预留 ${formatQty(row.reservedQty)} / 可用 ${formatQty(availableQty(row))}</span>
            </article>
          `,
        )
        .join("")
    : '<p class="muted">暂无现存量低于 5 的备件。</p>';

  const sharedItems = buildSharedStockItems().slice(0, 5);
  elements.transferList.innerHTML = sharedItems.length
    ? sharedItems
        .map(
          (item) => `
            <article class="compact-item">
              <strong>${escapeHtml(item.materialCode)} / ${escapeHtml(item.materialName)}</strong>
              <span>${escapeHtml(item.warehouses)} / 合计可用 ${formatQty(item.availableQty)}</span>
            </article>
          `,
        )
        .join("")
    : '<p class="muted">暂无同一物料跨多个海外仓存放的记录。</p>';
}

function buildSharedStockItems() {
  const byMaterial = new Map();
  getVisibleInventoryRows().forEach((row) => {
    const list = byMaterial.get(row.materialCode) || [];
    list.push(row);
    byMaterial.set(row.materialCode, list);
  });

  const items = [];
  byMaterial.forEach((materialRows, materialCode) => {
    const warehouseNames = Array.from(new Set(materialRows.map((row) => getWarehouseName(row.warehouseId))));
    if (warehouseNames.length < 2) return;
    items.push({
      materialCode,
      materialName: materialRows[0].materialName,
      warehouses: warehouseNames.join("、"),
      availableQty: materialRows.reduce((sum, row) => sum + availableQty(row), 0),
    });
  });
  return items.sort((a, b) => b.availableQty - a.availableQty);
}

function renderTable(filteredRows) {
  elements.rowCount.textContent = `${filteredRows.length} 条`;
  elements.emptyState.hidden = filteredRows.length > 0;
  elements.stockTableBody.innerHTML = buildMaterialRowsHtml(filteredRows);
}

function buildOverviewRowsHtml(sourceRows) {
  return sourceRows
    .map((row) => {
      const status = rowStatus(row);
      return `
        <tr>
          <td><span class="project-pill">${escapeHtml(row.project)}</span></td>
          <td class="code-cell">${escapeHtml(row.materialCode)}</td>
          <td>${escapeHtml(row.materialName || "-")}</td>
          <td>${escapeHtml(row.category || "-")}</td>
          <td>${escapeHtml(row.location || "-")}</td>
          <td class="num">${formatQty(row.onHandQty)}</td>
          <td class="num">${formatQty(row.reservedQty)}</td>
          <td class="num">${formatQty(row.frozenQty)}</td>
          <td class="num">${formatQty(availableQty(row))}</td>
          <td class="num">${formatMoney(inventoryAmount(row))}</td>
          <td><span class="status-tag tag-${status}">${statusLabel(status)}</span></td>
        </tr>
      `;
    })
    .join("");
}

function buildMaterialRowsHtml(sourceRows) {
  return sourceRows
    .map((row) => {
      const status = rowStatus(row);
      return `
        <tr>
          <td><span class="project-pill">${escapeHtml(row.project)}</span></td>
          <td>${escapeHtml(getWarehouseName(row.warehouseId))}</td>
          <td class="code-cell">${escapeHtml(row.materialCode)}</td>
          <td>${escapeHtml(row.materialName || "-")}</td>
          <td>${escapeHtml(row.location || "-")}</td>
          <td class="num">${formatQty(row.onHandQty)}</td>
          <td class="num">${formatQty(row.reservedQty)}</td>
          <td class="num">${formatQty(row.frozenQty)}</td>
          <td class="num">${formatQty(availableQty(row))}</td>
          <td><span class="status-tag tag-${status}">${statusLabel(status)}</span></td>
        </tr>
      `;
    })
    .join("");
}

function renderReplenishment() {
  const filteredOrders = getReplenishmentFilteredRows();
  const totalQty = filteredOrders.reduce((sum, order) => sum + order.qty, 0);
  const orderCount = new Set(filteredOrders.map((order) => order.orderNo)).size;
  elements.replenishmentCount.textContent = `${filteredOrders.length} 条 / ${orderCount} 个订单 / 数量 ${formatQty(totalQty)}`;
  elements.replenishmentStatus.textContent = afterSalesRefreshMessage || (replenishmentOrders.length ? `源数据 ${replenishmentOrders.length} 条` : "待接入表格数据");

  elements.replenishmentWarehouseGrid.innerHTML = warehouses
    .map((warehouse) => {
      const warehouseOrders = filteredOrders.filter((order) => order.warehouseId === warehouse.id);
      const warehouseOrderCount = new Set(warehouseOrders.map((order) => order.orderNo)).size;
      const skuCount = new Set(warehouseOrders.map((order) => order.materialCode)).size;
      const qty = warehouseOrders.reduce((sum, order) => sum + parseNumber(order.qty), 0);
      const statuses = uniqueValues(warehouseOrders, "status").join(" / ") || "无记录";
      const transports = uniqueValues(warehouseOrders, "transport").join(" / ") || "-";
      const activeClass = state.replenishmentWarehouse === warehouse.id ? " active" : "";
      return `
        <article class="replenishment-card${activeClass}" data-replenishment-warehouse="${warehouse.id}" style="--accent: ${warehouse.accent}; --soft: ${warehouse.soft}">
          <div class="material-warehouse-head">
            <span class="warehouse-code">${escapeHtml(warehouse.code)}</span>
            <div>
              <strong>${escapeHtml(warehouse.name)}</strong>
              <small>${escapeHtml(transports)}</small>
            </div>
          </div>
          <div class="warehouse-meta">
            <span>订单 ${warehouseOrderCount}</span>
            <span>SKU ${skuCount}</span>
            <span>数量 ${formatQty(qty)}</span>
          </div>
          <p class="card-note">${escapeHtml(statuses)}</p>
        </article>
      `;
    })
    .join("");

  elements.replenishmentEmptyState.hidden = filteredOrders.length > 0;
  elements.replenishmentTableBody.innerHTML = filteredOrders
    .map((order) => {
      const warehouse = findWarehouse(order.warehouseId);
      return `
        <tr>
          <td>${escapeHtml(warehouse.name)}</td>
          <td class="code-cell">${escapeHtml(order.orderNo || "-")}</td>
          <td>${escapeHtml(order.shippedDate || "-")}</td>
          <td>${escapeHtml(order.expectedArrival || "-")}</td>
          <td>${escapeHtml(order.transport || "-")}</td>
          <td><span class="status-tag ${statusTagClass(order.status)}">${escapeHtml(order.status || "待确认")}</span></td>
          <td>${escapeHtml(inboundLabel(order.inboundFlag))}</td>
          <td class="code-cell">${escapeHtml(order.materialCode || "-")}</td>
          <td>${escapeHtml(order.materialName || "-")}</td>
          <td class="num">${formatQty(order.qty)}</td>
          <td>${escapeHtml(order.leadTime || "-")}</td>
        </tr>
      `;
    })
    .join("");
}

function getReplenishmentFilteredRows() {
  const keyword = state.replenishmentKeyword.trim().toLowerCase();
  return replenishmentOrders.filter((order) => {
    const text = [
      getWarehouseName(order.warehouseId),
      order.warehouseText,
      order.orderNo,
      order.materialCode,
      order.materialName,
      order.transport,
      order.status,
      order.shippedDate,
      order.expectedArrival,
      order.leadTime,
    ]
      .join(" ")
      .toLowerCase();
    const matchesKeyword = !keyword || text.includes(keyword);
    const matchesWarehouse = state.replenishmentWarehouse === "all" || order.warehouseId === state.replenishmentWarehouse;
    const matchesStatus = state.replenishmentStatus === "all" || order.status === state.replenishmentStatus;
    const matchesTransport = state.replenishmentTransport === "all" || order.transport === state.replenishmentTransport;
    return matchesKeyword && matchesWarehouse && matchesStatus && matchesTransport;
  });
}

function renderRma() {
  const baseGroups = getRmaFilteredGroups({ includeQuickFilter: false });
  const filteredGroups = applyRmaQuickFilter(baseGroups, state.rmaQuickFilter);
  const filteredRows = filteredGroups.flatMap((group) => group.lines);
  const selectedGroup = getSelectedRmaGroup(filteredGroups);
  const activeFilter = getRmaQuickFilter(state.rmaQuickFilter);

  elements.rmaCount.textContent = `${filteredGroups.length} 个 RM 订单 / ${filteredRows.length} 条明细`;
  elements.rmaStatus.textContent = afterSalesRefreshMessage || (rmaOrders.length ? `当前分组：${activeFilter.label} / 源数据 ${rmaOrders.length} 条` : "待接入 RMA 数据");
  elements.rmaSummaryGrid.innerHTML = getRmaQuickFilters(baseGroups)
    .map(
      (filter) => `
        <article class="summary-card summary-action-card${filter.id === state.rmaQuickFilter ? " active" : ""}" data-rma-filter="${escapeHtml(filter.id)}" role="button" tabindex="0">
          <span>${escapeHtml(filter.label)}</span>
          <strong>${escapeHtml(formatQty(filter.count))}</strong>
          <small>${escapeHtml(filter.note)}</small>
        </article>
      `,
    )
    .join("");

  elements.rmaEmptyState.hidden = filteredGroups.length > 0;
  elements.rmaTableBody.innerHTML = filteredGroups
    .map(
      (group) => `
        <tr class="clickable-row${selectedGroup && selectedGroup.rmaNo === group.rmaNo ? " active" : ""}" data-rma-no="${escapeHtml(group.rmaNo)}">
          <td class="code-cell" title="${escapeHtml(group.rmaNo || "-")}"><span class="clip-text">${escapeHtml(group.rmaNo || "-")}</span></td>
          <td title="${escapeHtml(group.serviceNos.join("、") || "-")}"><span class="clip-text two-line">${escapeHtml(group.serviceNos[0] || "-")}</span></td>
          <td title="${escapeHtml(group.countries.join("、") || "-")}"><span class="clip-text">${escapeHtml(group.countries.join("、") || "-")}</span></td>
          <td title="${escapeHtml(group.customers.join("、") || "-")}"><span class="clip-text two-line">${escapeHtml(group.customers.join("、") || "-")}</span></td>
          <td title="${escapeHtml(group.shippedDates.join("、") || "-")}"><span class="clip-text">${escapeHtml(group.shippedDates.join("、") || "-")}</span></td>
          <td>${group.statuses.map((status) => `<span class="status-tag ${statusTagClass(status)}">${escapeHtml(status)}</span>`).join("")}</td>
          <td>${buildTrackingLinks(group.trackingNos)}</td>
          <td title="${escapeHtml(group.faes.join("、") || "-")}"><span class="clip-text">${escapeHtml(group.faes.join("、") || "-")}</span></td>
        </tr>
      `,
    )
    .join("");

  renderRmaDetail(selectedGroup);
}

function getRmaFilteredGroups(options = {}) {
  const groups = buildRmaGroups(rmaOrders);
  const keyword = state.rmaKeyword.trim().toLowerCase();
  const filtered = groups.filter((group) => {
    const matchesKeyword = !keyword || group.searchText.includes(keyword);
    const matchesStatus = state.rmaStatus === "all" || group.statuses.includes(state.rmaStatus);
    const matchesCountry = state.rmaCountry === "all" || group.countries.includes(state.rmaCountry);
    const matchesProject = state.rmaProject === "all" || group.projects.includes(state.rmaProject);
    return matchesKeyword && matchesStatus && matchesCountry && matchesProject;
  });

  return options.includeQuickFilter === false ? filtered : applyRmaQuickFilter(filtered, state.rmaQuickFilter);
}

function getRmaQuickFilters(groups) {
  const allGroups = groups || [];
  const filters = [
    { id: "all", label: "RMA订单", note: "点击查看全部订单" },
    { id: "pendingStock", label: "待备货", note: "点击查看待备货订单" },
    { id: "pendingInbound", label: "待生产入库", note: "点击查看待入库订单" },
    { id: "shipped", label: "已出库（包含已完成）", note: "点击查看已出库订单" },
  ];

  return filters.map((filter) => ({
    ...filter,
    count: applyRmaQuickFilter(allGroups, filter.id).length,
  }));
}

function getRmaQuickFilter(filterId) {
  return getRmaQuickFilters([]).find((filter) => filter.id === filterId) || getRmaQuickFilters([])[0];
}

function applyRmaQuickFilter(groups, filterId) {
  const filter = filterId || "all";
  if (filter === "pendingStock") return groups.filter((group) => groupHasAnyStatus(group, ["待备货"]));
  if (filter === "pendingInbound") return groups.filter((group) => groupHasAnyStatus(group, ["待生产入库"]));
  if (filter === "shipped") return groups.filter((group) => groupHasAnyStatus(group, ["已出库", "已完成"]));
  return groups;
}

function groupHasAnyStatus(group, statuses) {
  return group.statuses.some((status) => statuses.includes(status));
}

function buildRmaGroups(sourceRows) {
  const groupMap = new Map();
  sourceRows.forEach((row, index) => {
    const key = row.rmaNo || `未标注-${index + 1}`;
    const group =
      groupMap.get(key) ||
      {
        rmaNo: row.rmaNo || key,
        serviceNos: [],
        countries: [],
        customers: [],
        models: [],
        projects: [],
        statuses: [],
        trackingNos: [],
        shippedDates: [],
        faes: [],
        totalQty: 0,
        skuCount: 0,
        lines: [],
        searchText: "",
      };

    group.lines.push(row);
    group.totalQty += row.qty;
    pushUnique(group.serviceNos, row.serviceNo);
    pushUnique(group.countries, row.country);
    pushUnique(group.customers, row.customer);
    pushUnique(group.models, row.model);
    pushUnique(group.projects, row.project);
    pushUnique(group.statuses, row.status);
    pushUnique(group.trackingNos, row.trackingNo);
    pushUnique(group.shippedDates, row.shippedDate);
    pushUnique(group.faes, row.fae);
    groupMap.set(key, group);
  });

  return Array.from(groupMap.values()).map((group) => {
    group.skuCount = new Set(group.lines.map((line) => line.materialCode).filter(Boolean)).size;
    group.searchText = [
      group.rmaNo,
      group.serviceNos.join(" "),
      group.countries.join(" "),
      group.customers.join(" "),
      group.models.join(" "),
      group.projects.join(" "),
      group.statuses.join(" "),
      group.trackingNos.join(" "),
      group.shippedDates.join(" "),
      group.faes.join(" "),
      ...group.lines.flatMap((line) => [
        line.sn,
        line.materialCode,
        line.materialName,
        line.afterSalesType,
        line.specialNote,
        line.reason,
        line.badPartSuggestion,
      ]),
    ]
      .join(" ")
      .toLowerCase();
    return group;
  });
}

function getSelectedRmaGroup(groups) {
  if (!groups.length) {
    state.rmaSelectedNo = "";
    return null;
  }
  const selected = groups.find((group) => group.rmaNo === state.rmaSelectedNo) || groups[0];
  state.rmaSelectedNo = selected.rmaNo;
  return selected;
}

function renderRmaDetail(group) {
  if (!group) {
    elements.rmaDetailTitle.textContent = "RMA产品明细";
    elements.rmaDetailMeta.textContent = "";
    elements.rmaDetailCount.textContent = "0 条";
    elements.rmaDetailEmptyState.hidden = false;
    elements.rmaDetailTableBody.innerHTML = "";
    return;
  }

  elements.rmaDetailTitle.textContent = `${group.rmaNo} 产品明细`;
  elements.rmaDetailMeta.textContent = `${group.customers.join("、") || "-"} / ${group.countries.join("、") || "-"} / ${group.models.join("、") || "-"}`;
  elements.rmaDetailCount.textContent = `${group.lines.length} 条 / 品种 ${formatQty(group.skuCount)} / 数量 ${formatQty(group.totalQty)}`;
  elements.rmaDetailEmptyState.hidden = group.lines.length > 0;
  elements.rmaDetailTableBody.innerHTML = group.lines
    .map(
      (row) => `
        <tr>
          <td class="code-cell">${escapeHtml(row.materialCode || "-")}</td>
          <td>${escapeHtml(row.materialName || "-")}</td>
          <td class="num">${formatQty(row.qty)}</td>
          <td><span class="status-tag ${statusTagClass(row.status)}">${escapeHtml(row.status || "待确认")}</span></td>
          <td>${buildTrackingLinks(row.trackingNo)}</td>
          <td>${escapeHtml(row.unitPrice || "-")}</td>
          <td>${escapeHtml(row.totalUsd || "-")}</td>
          <td class="reason-cell">${escapeHtml(row.specialNote || "-")}</td>
          <td class="reason-cell">${escapeHtml(row.reason || "-")}</td>
        </tr>
      `,
    )
    .join("");
}

function setupResizableColumns(table, storageKey, options = {}) {
  if (!table) return;
  const minWidth = options.minWidth || 64;
  const columns = Array.from(table.querySelectorAll("col"));
  const headers = Array.from(table.querySelectorAll("thead th"));
  if (!columns.length || columns.length !== headers.length) return;

  try {
    const savedWidths = JSON.parse(localStorage.getItem(storageKey) || "[]");
    savedWidths.forEach((width, index) => {
      if (Number.isFinite(width) && columns[index]) columns[index].style.width = `${width}px`;
    });
  } catch {
    localStorage.removeItem(storageKey);
  }

  headers.forEach((header, index) => {
    header.classList.add("resizable-heading");
    const handle = document.createElement("span");
    handle.className = "column-resize-handle";
    handle.setAttribute("aria-hidden", "true");

    handle.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const startX = event.clientX;
      const startWidth = columns[index].getBoundingClientRect().width || header.getBoundingClientRect().width;
      const pointerId = event.pointerId;
      handle.setPointerCapture(pointerId);
      document.body.classList.add("is-resizing-column");

      const handleMove = (moveEvent) => {
        const nextWidth = Math.max(minWidth, Math.round(startWidth + moveEvent.clientX - startX));
        columns[index].style.width = `${nextWidth}px`;
      };

      const finishResize = () => {
        if (handle.hasPointerCapture(pointerId)) handle.releasePointerCapture(pointerId);
        handle.removeEventListener("pointermove", handleMove);
        handle.removeEventListener("pointerup", finishResize);
        handle.removeEventListener("pointercancel", finishResize);
        document.body.classList.remove("is-resizing-column");
        const widths = columns.map((column) => Math.round(column.getBoundingClientRect().width));
        localStorage.setItem(storageKey, JSON.stringify(widths));
      };

      handle.addEventListener("pointermove", handleMove);
      handle.addEventListener("pointerup", finishResize);
      handle.addEventListener("pointercancel", finishResize);
    });

    header.appendChild(handle);
  });
}

function pushUnique(list, value) {
  const text = String(value || "").trim();
  if (text && !list.includes(text)) list.push(text);
}

function buildTrackingLinks(trackingNos) {
  const values = Array.isArray(trackingNos) ? trackingNos : [trackingNos];
  const links = values
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .map((trackingNo) => {
      const queryValue = getTrackingQueryValue(trackingNo);
      const url = `https://t.17track.net/zh-cn#nums=${encodeURIComponent(queryValue)}`;
      return `<a class="track-link" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" title="在 17TRACK 查询 ${escapeHtml(queryValue)}">${escapeHtml(trackingNo)}</a>`;
    });

  return links.length ? `<span class="track-link-list">${links.join("")}</span>` : "-";
}

function getTrackingQueryValue(trackingNo) {
  const text = String(trackingNo || "").trim();
  const candidates = text.match(/[A-Z0-9][A-Z0-9-]{5,}/gi) || [];
  const carrierWords = new Set(["DHL", "PAKET", "EXPRESS", "FEDEX", "UPS"]);
  return (
    candidates
      .reverse()
      .find((candidate) => /\d/.test(candidate) && !carrierWords.has(candidate.toUpperCase())) || text
  );
}

function inboundLabel(value) {
  if (String(value).trim() === "1") return "已入库";
  if (String(value).trim() === "0") return "未入库";
  return value || "-";
}

function statusTagClass(status) {
  const text = String(status || "");
  if (text.includes("完成") || text.includes("已入库") || text.includes("已出库")) return "tag-normal";
  if (text.includes("运输") || text.includes("签收") || text.includes("快递")) return "tag-reserved";
  return "tag-shortage";
}

function parseImportedRows(text) {
  const table = parseDelimitedRows(text);
  const indexes = getHeaderIndexes(table.headers);

  return table.dataRows.map((cells) =>
    normalizeRow({
      warehouse: cell(cells, indexes.warehouse),
      project: cell(cells, indexes.project),
      materialCode: cell(cells, indexes.materialCode),
      materialName: cell(cells, indexes.materialName),
      category: cell(cells, indexes.category),
      productLine: cell(cells, indexes.productLine),
      location: cell(cells, indexes.location),
      spec: cell(cells, indexes.spec),
      model: cell(cells, indexes.model),
      onHandQty: cell(cells, indexes.onHandQty),
      reservedQty: cell(cells, indexes.reservedQty),
      frozenQty: cell(cells, indexes.frozenQty),
      unitCost: cell(cells, indexes.unitCost),
      inventoryAmount: cell(cells, indexes.inventoryAmount),
      unit: cell(cells, indexes.unit),
    }),
  );
}

function parseImportedReplenishmentRows(text) {
  const table = parseDelimitedRows(text);
  const indexes = getImportIndexes(table.headers, {
    orderNo: ["订单编号", "订单号", "备货订单", "order"],
    materialCode: ["物料编码", "物料编号", "料号", "material"],
    materialName: ["物料名称", "产品名称", "名称", "name"],
    qty: ["备货数量", "数量", "qty"],
    warehouse: ["海外仓", "仓库", "warehouse"],
    unitPrice: ["单价", "unitprice"],
    totalPrice: ["总价", "金额", "total"],
    shippedDate: ["发货时间", "发货日期", "发出日期"],
    expectedArrival: ["预计到仓", "预计到货", "预计到达"],
    leadTime: ["时效", "物流时效"],
    transport: ["运输方式", "运输"],
    status: ["状态", "status"],
    inboundFlag: ["是否入库", "入库"],
  });

  return table.dataRows
    .map((cells) =>
      normalizeReplenishmentOrder({
        orderNo: cell(cells, indexes.orderNo),
        materialCode: cell(cells, indexes.materialCode),
        materialName: cell(cells, indexes.materialName),
        qty: cell(cells, indexes.qty),
        warehouseText: cell(cells, indexes.warehouse),
        unitPrice: cell(cells, indexes.unitPrice),
        totalPrice: cell(cells, indexes.totalPrice),
        shippedDate: cell(cells, indexes.shippedDate),
        expectedArrival: cell(cells, indexes.expectedArrival),
        leadTime: cell(cells, indexes.leadTime),
        transport: cell(cells, indexes.transport),
        status: cell(cells, indexes.status),
        inboundFlag: cell(cells, indexes.inboundFlag),
      }),
    )
    .filter((order) => order.orderNo && order.materialCode);
}

function parseImportedRmaRows(text) {
  const table = parseDelimitedRows(text);
  const indexes = getImportIndexes(table.headers, {
    rmaNo: ["售后单号", "RM订单号", "RM单号", "RMA订单", "RMA"],
    serviceNo: ["服务单号"],
    sn: ["SN号", "SN"],
    customer: ["客户名称", "客户"],
    productLine: ["产品线"],
    country: ["国家"],
    model: ["机型", "型号"],
    project: ["项目"],
    materialCode: ["物料编码", "物料编号", "料号", "material"],
    materialName: ["产品名称", "物料名称", "名称"],
    qty: ["需求数量", "需求总数", "数量", "qty"],
    unitPrice: ["单价"],
    totalUsd: ["总价 USD", "总价|USD", "总价USD", "总价", "USD"],
    origin: ["启运地", "发货地"],
    shippedDate: ["发出日期", "发货日期"],
    status: ["状态"],
    afterSalesType: ["售后类型"],
    specialNote: ["特殊备注", "备注"],
    reason: ["原因描述", "原因"],
    fae: ["FAE人员", "FAE"],
    badPartSuggestion: ["坏件处置建议"],
    trackingNo: ["物流单号", "物流订单", "快递单号"],
    logisticsDays: ["物流时效/天", "物流时效", "时效"],
    parentRecord: ["父记录"],
  });

  return table.dataRows
    .map((cells) =>
      normalizeRmaOrder({
        rmaNo: cell(cells, indexes.rmaNo),
        serviceNo: cell(cells, indexes.serviceNo),
        sn: cell(cells, indexes.sn),
        customer: cell(cells, indexes.customer),
        productLine: cell(cells, indexes.productLine),
        country: cell(cells, indexes.country),
        model: cell(cells, indexes.model),
        project: cell(cells, indexes.project),
        materialCode: cell(cells, indexes.materialCode),
        materialName: cell(cells, indexes.materialName),
        qty: cell(cells, indexes.qty),
        unitPrice: cell(cells, indexes.unitPrice),
        totalUsd: cell(cells, indexes.totalUsd),
        origin: cell(cells, indexes.origin),
        shippedDate: cell(cells, indexes.shippedDate),
        status: cell(cells, indexes.status),
        afterSalesType: cell(cells, indexes.afterSalesType),
        specialNote: cell(cells, indexes.specialNote),
        reason: cell(cells, indexes.reason),
        fae: cell(cells, indexes.fae),
        badPartSuggestion: cell(cells, indexes.badPartSuggestion),
        trackingNo: cell(cells, indexes.trackingNo),
        logisticsDays: cell(cells, indexes.logisticsDays),
        parentRecord: cell(cells, indexes.parentRecord),
      }),
    )
    .filter((order) => order.rmaNo && order.materialCode);
}

function parseDelimitedRows(text) {
  const raw = text.trim();
  if (!raw) return { headers: [], dataRows: [] };
  const delimiter = raw.includes("\t") ? "\t" : ",";
  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const headers = splitLine(lines[0], delimiter).map((header) => header.replace(/^\uFEFF/, ""));

  return {
    headers,
    dataRows: lines.slice(1).map((line) => splitLine(line, delimiter)),
  };
}

function splitLine(line, delimiter) {
  if (delimiter === "\t") return line.split("\t").map((item) => item.trim());
  return line.split(",").map((item) => item.replace(/^"|"$/g, "").trim());
}

function getImportIndexes(headers, aliases) {
  const normalizedHeaders = headers.map(normalizeHeaderLabel);
  const genericAliases = new Set(["名称", "name"]);

  return Object.fromEntries(
    Object.entries(aliases).map(([key, names]) => {
      const normalizedNames = names.map(normalizeHeaderLabel);
      const exactIndex = normalizedHeaders.findIndex((header) => normalizedNames.includes(header));
      if (exactIndex >= 0) return [key, exactIndex];

      const fuzzyNames = names
        .map((name) => ({ raw: String(name || ""), normalized: normalizeHeaderLabel(name) }))
        .filter((name) => name.normalized && !genericAliases.has(name.raw.toLowerCase()));
      return [
        key,
        normalizedHeaders.findIndex((header) =>
          fuzzyNames.some((name) => header.includes(name.normalized) || name.normalized.includes(header)),
        ),
      ];
    }),
  );
}

function normalizeHeaderLabel(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[\s_/-]+/g, "");
}

function getHeaderIndexes(headers) {
  const aliases = {
    warehouse: ["仓库", "warehouse"],
    project: ["项目", "主采购项目", "project"],
    materialCode: ["物料编码", "物料编号", "料号", "material"],
    materialName: ["物料名称", "名称", "name"],
    category: ["分类", "物料基本分类", "category"],
    productLine: ["产品线", "product"],
    location: ["货位", "库位", "location"],
    spec: ["规格", "spec"],
    model: ["型号", "model"],
    onHandQty: ["现存", "结存", "库存", "onhand"],
    reservedQty: ["预留", "reserved"],
    frozenQty: ["冻结", "frozen"],
    unitCost: ["成本", "成本价", "单价", "unitcost", "standardprice"],
    inventoryAmount: ["库存金额", "金额", "amount", "value"],
    unit: ["单位", "主单位", "unit"],
  };

  return getImportIndexes(headers, aliases);
}

function cell(cells, index) {
  return index >= 0 ? cells[index] || "" : "";
}

function upsertRows(incomingRows) {
  const map = new Map(rows.map((row) => [rowKey(row), row]));
  incomingRows.forEach((row) => {
    if (!row.materialCode) return;
    map.set(rowKey(row), { ...(map.get(rowKey(row)) || {}), ...row });
  });
  rows = Array.from(map.values());
  dataOverrides.inventory = true;
  save();
}

function replaceInventoryRows(incomingRows) {
  const metaByMaterial = new Map();
  [...baseRows, ...rows].forEach((row) => {
    const key = String(row.materialCode || "").trim().toLowerCase();
    if (!key) return;
    const existing = metaByMaterial.get(key) || {};
    metaByMaterial.set(key, {
      ...existing,
      productId: row.productId || existing.productId,
      productTemplateId: row.productTemplateId || existing.productTemplateId,
      imageUrl: row.imageUrl || existing.imageUrl,
      priceSource: row.priceSource || existing.priceSource,
      priceStartPlace: row.priceStartPlace || existing.priceStartPlace,
      unitCost: row.unitCost > 0 ? row.unitCost : existing.unitCost || 0,
    });
  });

  rows = incomingRows
    .map((row) => {
      const normalized = normalizeRow(row);
      const meta = metaByMaterial.get(normalized.materialCode.toLowerCase());
      if (!meta) return normalized;
      const unitCost = normalized.unitCost > 0 ? normalized.unitCost : meta.unitCost || 0;
      return {
        ...normalized,
        productId: normalized.productId || meta.productId || "",
        productTemplateId: normalized.productTemplateId || meta.productTemplateId || "",
        imageUrl: normalized.imageUrl || meta.imageUrl || "",
        priceSource: normalized.priceSource || meta.priceSource || "",
        priceStartPlace: normalized.priceStartPlace || meta.priceStartPlace || "",
        unitCost,
        inventoryAmount: normalized.inventoryAmount > 0 ? normalized.inventoryAmount : normalized.onHandQty * unitCost,
      };
    })
    .filter((row) => row.materialCode);
  dataOverrides.inventory = true;
  save();
}

function rowKey(row) {
  return [row.warehouseId, row.project, row.materialCode, row.location].join("|");
}

function exportCsv() {
  const filteredRows = getFilteredRows();
  const headers = ["仓库", "项目", "物料编码", "物料名称", "分类", "产品线", "货位", "规格", "型号", "现存", "预留", "冻结", "可用", "库存金额", "单位", "状态"];
  const body = filteredRows.map((row) => [
    getWarehouseName(row.warehouseId),
    row.project,
    row.materialCode,
    row.materialName,
    row.category,
    row.productLine,
    row.location,
    row.spec,
    row.model,
    row.onHandQty,
    row.reservedQty,
    row.frozenQty,
    availableQty(row),
    inventoryAmount(row),
    row.unit,
    statusLabel(rowStatus(row)),
  ]);
  const csv = [headers, ...body]
    .map((line) => line.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `海外库存看板-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function templateText() {
  return [
    "仓库,项目,物料编码,物料名称,分类,产品线,货位,规格,型号,现存,预留,冻结,库存金额,单位",
    "德国法兰克福仓,R1916,DEMO-001,示例物料,零部件及主材_电器件,推料机器人配件,SVEA,示例规格,示例型号,10,2,0,1000,个",
    "美国芝加哥仓,R1916,DEMO-001,示例物料,零部件及主材_电器件,推料机器人配件,SVEA HK,示例规格,示例型号,5,0,0,500,个",
  ].join("\n");
}

function replenishmentTemplateText() {
  return [
    ["订单编号", "物料编码", "物料名称", "数量", "海外仓", "单价", "总价", "发货时间", "预计到仓", "时效", "运输方式", "状态", "是否入库"].join("\t"),
    ["BH260000000001", "1S0114004J0260A", "视觉算法板", "9", "德国法兰克福仓", "", "", "2026-07-02", "2026-08-02", "31天", "海运", "运输中", "0"].join("\t"),
  ].join("\n");
}

function rmaTemplateText() {
  return [
    ["售后单号", "服务单号", "SN号", "客户名称", "产品线", "国家", "机型", "项目", "物料编码", "产品名称", "需求数量", "单价", "总价 USD", "启运地", "发出日期", "状态", "售后类型", "特殊备注", "原因描述", "FAE人员", "坏件处置建议", "物流单号", "物流时效/天", "父记录"].join("\t"),
    ["RM260000001", "CS20260702000001", "SVEA-DEMO-SN", "示例客户", "畜牧", "德国", "畜牧-Nimbo", "R1916", "1S0114004J0260A", "视觉算法板", "1", "141", "141", "法兰克福", "2026-07-02", "待备货", "", "", "示例原因", "Anna.Jiang", "", "DHL 1234567890", "", ""].join("\t"),
  ].join("\n");
}

function setImportStatus(element, message) {
  if (element) element.textContent = message;
}

function readImportFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("文件读取失败，请重新选择表格。"));
    reader.readAsText(file, "utf-8");
  });
}

function readImportFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("文件读取失败，请重新选择表格。"));
    reader.readAsArrayBuffer(file);
  });
}

function isPlainTextImportFile(file) {
  return /\.(tsv|txt)$/i.test(file.name || "");
}

function formatImportedCell(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  return String(value ?? "").replace(/\r?\n/g, " ").trim();
}

const importHeaderHints = [
  "仓库",
  "海外仓",
  "物料编码",
  "物料编号",
  "料号",
  "订单编号",
  "售后单号",
  "RM订单号",
  "服务单号",
  "产品名称",
  "物料名称",
  "需求数量",
  "总价",
  "状态",
];

function normalizeWorksheetRows(rows) {
  return rows
    .map((row) => row.map(formatImportedCell))
    .filter((row) => row.some(Boolean));
}

function countHeaderHits(row) {
  const normalizedHints = importHeaderHints.map(normalizeHeaderLabel);
  const labels = row.map(normalizeHeaderLabel).filter(Boolean);
  return normalizedHints.filter((hint) => labels.some((label) => label.includes(hint) || hint.includes(label))).length;
}

function findLikelyHeaderRowIndex(rows) {
  return rows.findIndex((row) => countHeaderHits(row) >= 2);
}

function worksheetRowsToText(rows) {
  const normalizedRows = normalizeWorksheetRows(rows);
  const headerIndex = findLikelyHeaderRowIndex(normalizedRows);
  const tableRows = headerIndex > 0 ? normalizedRows.slice(headerIndex) : normalizedRows;
  return tableRows.map((row) => row.join("\t")).join("\n");
}

function scoreWorksheetRows(rows) {
  const normalizedRows = normalizeWorksheetRows(rows);
  if (!normalizedRows.length) return 0;

  const headerIndex = findLikelyHeaderRowIndex(normalizedRows);
  if (headerIndex < 0) return 1;

  const headerHits = countHeaderHits(normalizedRows[headerIndex]);
  const dataRowCount = Math.max(0, normalizedRows.length - headerIndex - 1);
  return headerHits * 100 + dataRowCount;
}

function getBestWorksheetRows(workbook) {
  const candidates = workbook.SheetNames.map((sheetName) => {
    const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });

    return {
      sheetName,
      rows,
      score: scoreWorksheetRows(rows),
    };
  }).filter((candidate) => candidate.rows.length);

  return candidates.sort((a, b) => b.score - a.score)[0] || null;
}

async function readImportFile(file) {
  if (isPlainTextImportFile(file)) return readImportFileAsText(file);
  if (!window.XLSX) throw new Error("Excel 解析库没有加载成功，请刷新页面后重试。");

  const isCsv = /\.(csv)$/i.test(file.name || "");
  const workbook = isCsv
    ? window.XLSX.read(await readImportFileAsText(file), { type: "string", raw: false })
    : window.XLSX.read(await readImportFileAsArrayBuffer(file), { type: "array", cellDates: true });
  const bestSheet = getBestWorksheetRows(workbook);
  if (!bestSheet) throw new Error("这个文件没有可读取的工作表。");

  const text = worksheetRowsToText(bestSheet.rows);
  if (!text.trim()) throw new Error("这个文件没有可读取的数据。");
  return text;
}

async function importFromFile(input, area, statusElement, importAction, label) {
  const file = input?.files?.[0];
  if (!file) return;

  setImportStatus(statusElement, `正在读取 ${file.name}...`);
  try {
    area.value = await readImportFile(file);
    const result = importAction(area, statusElement);
    if (result?.ok) setImportStatus(statusElement, `${result.message}（${file.name}）`);
  } catch (error) {
    setImportStatus(statusElement, `${label}导入失败：${error.message}`);
  } finally {
    input.value = "";
  }
}

function importInventoryFromArea(area, statusElement) {
  try {
    const imported = parseImportedRows(area.value).filter((row) => row.materialCode);
    if (!imported.length) throw new Error("没有识别到库存明细，请保留表头后再粘贴。");
    replaceInventoryRows(imported);
    resetFilters();
    area.value = "";
    const message = `已导入 ${imported.length} 条库存`;
    setImportStatus(statusElement, message);
    render();
    return { ok: true, count: imported.length, message };
  } catch (error) {
    setImportStatus(statusElement, error.message);
    return { ok: false, message: error.message };
  }
}

function restoreInventoryFromSource(statusElement) {
  rows = baseRows;
  dataOverrides.inventory = false;
  resetFilters();
  save();
  setImportStatus(statusElement, hasSyncedInventoryRows ? "已恢复 Odoo 同步库存" : "已恢复系统库存");
  render();
}

async function copyOdooUpdateCommand() {
  const command =
    'powershell -NoProfile -ExecutionPolicy Bypass -File .\\sync-dashboard-data.ps1 -OdooLogin anna.jiang@sveav.com -PromptOdooPassword';
  try {
    await navigator.clipboard.writeText(command);
    setImportStatus(elements.overviewImportStatus, "Odoo 更新命令已复制");
  } catch {
    setImportStatus(elements.overviewImportStatus, command);
  }
}

function importReplenishmentFromArea(area, statusElement) {
  try {
    const imported = parseImportedReplenishmentRows(area.value);
    if (!imported.length) throw new Error("没有识别到备货订单，请保留表头后再粘贴。");
    replaceCollection(replenishmentOrders, imported);
    dataOverrides.replenishment = true;
    afterSalesRefreshMessage = "";
    resetReplenishmentFilters();
    area.value = "";
    save();
    const message = `已导入 ${imported.length} 条备货明细`;
    setImportStatus(statusElement, message);
    render();
    return { ok: true, count: imported.length, message };
  } catch (error) {
    setImportStatus(statusElement, error.message);
    return { ok: false, message: error.message };
  }
}

function restoreReplenishmentFromSource(statusElement) {
  replaceCollection(replenishmentOrders, baseReplenishmentOrders);
  dataOverrides.replenishment = false;
  afterSalesRefreshMessage = "";
  resetReplenishmentFilters();
  save();
  setImportStatus(statusElement, "已恢复系统备货订单");
  render();
}

function importRmaFromArea(area, statusElement) {
  try {
    const imported = parseImportedRmaRows(area.value);
    if (!imported.length) throw new Error("没有识别到 RMA 明细，请保留表头后再粘贴。");
    replaceCollection(rmaOrders, imported);
    dataOverrides.rma = true;
    afterSalesRefreshMessage = "";
    resetRmaFilters();
    area.value = "";
    save();
    const message = `已导入 ${imported.length} 条 RMA 明细`;
    setImportStatus(statusElement, message);
    render();
    return { ok: true, count: imported.length, message };
  } catch (error) {
    setImportStatus(statusElement, error.message);
    return { ok: false, message: error.message };
  }
}

function restoreRmaFromSource(statusElement) {
  replaceCollection(rmaOrders, baseRmaOrders);
  dataOverrides.rma = false;
  afterSalesRefreshMessage = "";
  resetRmaFilters();
  save();
  setImportStatus(statusElement, "已恢复系统 RMA 订单");
  render();
}

function resetFilters() {
  state = { ...state, keyword: "", warehouse: "all", project: "all", status: "all", sort: "availableAsc", overviewQuickFilter: "all" };
}

function resetReplenishmentFilters() {
  state.replenishmentKeyword = "";
  state.replenishmentWarehouse = "all";
  state.replenishmentStatus = "all";
  state.replenishmentTransport = "all";
}

function resetRmaFilters() {
  state.rmaKeyword = "";
  state.rmaStatus = "all";
  state.rmaCountry = "all";
  state.rmaProject = "all";
  state.rmaQuickFilter = "all";
  state.rmaSelectedNo = "";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

elements.viewTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    state.view = tab.dataset.view;
    save();
    render();
    if (state.view === "material") elements.keywordInput.focus();
    if (state.view === "replenishment") elements.replenishmentKeywordInput.focus();
    if (state.view === "rma") elements.rmaKeywordInput.focus();
  });
});

elements.keywordInput.addEventListener("input", (event) => {
  state.keyword = event.target.value;
  save();
  render();
});

elements.warehouseSelect.addEventListener("change", (event) => {
  state.warehouse = event.target.value;
  save();
  render();
});

elements.projectSelect.addEventListener("change", (event) => {
  state.project = event.target.value;
  save();
  render();
});

elements.statusSelect.addEventListener("change", (event) => {
  state.status = event.target.value;
  save();
  render();
});

elements.sortSelect.addEventListener("change", (event) => {
  state.sort = event.target.value;
  save();
  render();
});

elements.sampleButton.addEventListener("click", () => {
  restoreInventoryFromSource(elements.importStatus);
});

elements.resetButton.addEventListener("click", () => {
  resetFilters();
  elements.importStatus.textContent = "已重置筛选";
  save();
  render();
});

elements.exportButton.addEventListener("click", exportCsv);

elements.languageSelect?.addEventListener("change", (event) => {
  state.language = event.target.value;
  save();
  applyPageLanguage(state.language);
});

elements.warehouseList.addEventListener("click", (event) => {
  const card = event.target.closest("[data-warehouse]");
  if (!card) return;
  state.warehouse = state.warehouse === card.dataset.warehouse ? "all" : card.dataset.warehouse;
  save();
  render();
});

function applyOverviewQuickFilter(filter) {
  state.view = "overview";
  state.overviewQuickFilter = state.overviewQuickFilter === filter ? "all" : filter;
  save();
  render();
}

elements.summaryGrid.addEventListener("click", (event) => {
  const card = event.target.closest("[data-overview-filter]");
  if (!card) return;
  applyOverviewQuickFilter(card.dataset.overviewFilter || "all");
});

elements.summaryGrid.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const card = event.target.closest("[data-overview-filter]");
  if (!card) return;
  event.preventDefault();
  applyOverviewQuickFilter(card.dataset.overviewFilter || "all");
});

elements.lowStockList.addEventListener("click", (event) => {
  const item = event.target.closest("[data-overview-filter]");
  if (!item) return;
  applyOverviewQuickFilter(item.dataset.overviewFilter || "all");
});

elements.replenishmentKeywordInput.addEventListener("input", (event) => {
  state.replenishmentKeyword = event.target.value;
  save();
  render();
});

elements.replenishmentWarehouseSelect.addEventListener("change", (event) => {
  state.replenishmentWarehouse = event.target.value;
  save();
  render();
});

elements.replenishmentStatusSelect.addEventListener("change", (event) => {
  state.replenishmentStatus = event.target.value;
  save();
  render();
});

elements.replenishmentTransportSelect.addEventListener("change", (event) => {
  state.replenishmentTransport = event.target.value;
  save();
  render();
});

elements.replenishmentResetButton.addEventListener("click", () => {
  resetReplenishmentFilters();
  save();
  render();
});

elements.replenishmentRefreshButton?.addEventListener("click", refreshAfterSalesFromSource);

elements.replenishmentWarehouseGrid.addEventListener("click", (event) => {
  const card = event.target.closest("[data-replenishment-warehouse]");
  if (!card) return;
  state.replenishmentWarehouse = state.replenishmentWarehouse === card.dataset.replenishmentWarehouse ? "all" : card.dataset.replenishmentWarehouse;
  save();
  render();
});

elements.rmaKeywordInput.addEventListener("input", (event) => {
  state.rmaKeyword = event.target.value;
  save();
  render();
});

elements.rmaStatusSelect.addEventListener("change", (event) => {
  state.rmaStatus = event.target.value;
  state.rmaQuickFilter = "all";
  state.rmaSelectedNo = "";
  save();
  render();
});

elements.rmaCountrySelect.addEventListener("change", (event) => {
  state.rmaCountry = event.target.value;
  save();
  render();
});

elements.rmaProjectSelect.addEventListener("change", (event) => {
  state.rmaProject = event.target.value;
  save();
  render();
});

elements.rmaSummaryGrid.addEventListener("click", (event) => {
  const card = event.target.closest("[data-rma-filter]");
  if (!card) return;
  state.rmaQuickFilter = card.dataset.rmaFilter || "all";
  state.rmaStatus = "all";
  state.rmaSelectedNo = "";
  save();
  render();
});

elements.rmaSummaryGrid.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  const card = event.target.closest("[data-rma-filter]");
  if (!card) return;
  event.preventDefault();
  state.rmaQuickFilter = card.dataset.rmaFilter || "all";
  state.rmaStatus = "all";
  state.rmaSelectedNo = "";
  save();
  render();
});

elements.rmaTableBody.addEventListener("click", (event) => {
  if (event.target.closest(".track-link")) return;
  const row = event.target.closest("[data-rma-no]");
  if (!row) return;
  state.rmaSelectedNo = row.dataset.rmaNo;
  save();
  render();
});

elements.rmaResetButton.addEventListener("click", () => {
  resetRmaFilters();
  save();
  render();
});

elements.rmaRefreshButton?.addEventListener("click", refreshAfterSalesFromSource);

elements.importButton.addEventListener("click", () => {
  importInventoryFromArea(elements.importArea, elements.importStatus);
});

elements.fillTemplateButton.addEventListener("click", () => {
  elements.importArea.value = templateText();
  elements.importStatus.textContent = "模板已填入";
});

elements.overviewImportButton.addEventListener("click", () => {
  importInventoryFromArea(elements.overviewImportArea, elements.overviewImportStatus);
});

elements.overviewImportFile?.addEventListener("change", () => {
  importFromFile(elements.overviewImportFile, elements.overviewImportArea, elements.overviewImportStatus, importInventoryFromArea, "库存表格");
});

elements.overviewTemplateButton.addEventListener("click", () => {
  elements.overviewImportArea.value = templateText();
  elements.overviewImportStatus.textContent = "模板已填入";
});

elements.overviewRestoreButton.addEventListener("click", () => {
  restoreInventoryFromSource(elements.overviewImportStatus);
});
elements.odooCommandButton?.addEventListener("click", copyOdooUpdateCommand);

elements.replenishmentImportButton.addEventListener("click", () => {
  importReplenishmentFromArea(elements.replenishmentImportArea, elements.replenishmentImportStatus);
});

elements.replenishmentImportFile?.addEventListener("change", () => {
  importFromFile(
    elements.replenishmentImportFile,
    elements.replenishmentImportArea,
    elements.replenishmentImportStatus,
    importReplenishmentFromArea,
    "备货订单表格",
  );
});

elements.replenishmentTemplateButton.addEventListener("click", () => {
  elements.replenishmentImportArea.value = replenishmentTemplateText();
  elements.replenishmentImportStatus.textContent = "模板已填入";
});

elements.replenishmentRestoreButton.addEventListener("click", () => {
  restoreReplenishmentFromSource(elements.replenishmentImportStatus);
});

elements.rmaImportButton.addEventListener("click", () => {
  importRmaFromArea(elements.rmaImportArea, elements.rmaImportStatus);
});

elements.rmaImportFile?.addEventListener("change", () => {
  importFromFile(elements.rmaImportFile, elements.rmaImportArea, elements.rmaImportStatus, importRmaFromArea, "RMA订单表格");
});

elements.rmaTemplateButton.addEventListener("click", () => {
  elements.rmaImportArea.value = rmaTemplateText();
  elements.rmaImportStatus.textContent = "模板已填入";
});

elements.rmaRestoreButton.addEventListener("click", () => {
  restoreRmaFromSource(elements.rmaImportStatus);
});

load();
setupResizableColumns(elements.rmaOrderTable, `${STORAGE_KEY}-rma-order-columns`, { minWidth: 16 });
setupResizableColumns(elements.rmaDetailTable, `${STORAGE_KEY}-rma-detail-columns`);
render();
