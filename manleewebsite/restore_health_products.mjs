import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, getDocs, collection } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAgA4ki31tp775zoGESBB3inQVL9nieDEc",
  authDomain: "online-store-99126206.firebaseapp.com",
  projectId: "online-store-99126206",
  storageBucket: "online-store-99126206.firebasestorage.app",
  messagingSenderId: "154746962773",
  appId: "1:154746962773:web:7bd1c5d0d1b163e4d3e5dc"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// All products to restore/add (health supplements + new skincare)
// VIP prices corrected from user's master list
const allProducts = [
  // ── Monbélac ──
  { id:"p1",  code:"1000354", name:"法國 Monbélac 夢貝朗配方奶粉 1 號",     price:456, wholesale:48,  vip_price:48, material:"奶粉",     spec:"900G",            category:"Monbélac",  stock:99, image:"", desc:"" },
  { id:"p2",  code:"1000356", name:"法國 Monbélac 夢貝朗配方奶粉 3 號",     price:456, wholesale:48,  vip_price:48, material:"奶粉",     spec:"900G",            category:"Monbélac",  stock:99, image:"", desc:"" },
  // ── Activitae ──
  { id:"p3",  code:"1000043", name:"Activitae 瓜拿那",                       price:209, wholesale:22,  vip_price:22, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p4",  code:"1000044", name:"Activitae 蟻木",                         price:209, wholesale:22,  vip_price:22, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p5",  code:"1000046", name:"Activitae 山楂果",                       price:38,  wholesale:30,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p6",  code:"1000048", name:"Activitae 雄風寶加強精華版",             price:60,  wholesale:48,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p7",  code:"1000458", name:"Activitae 螺旋藻",                       price:247, wholesale:26,  vip_price:26, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p8",  code:"1084024", name:"Activitae 人参灵芝",                     price:228, wholesale:24,  vip_price:24, material:"口服液",   spec:"15包x每包15ml",   category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p9",  code:"1084041", name:"Activitae 祛濕膏",                       price:114, wholesale:12,  vip_price:12, material:"啫喱",     spec:"15條/每條10G",    category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p12", code:"1084059", name:"銀杏舞茸第二代",                         price:50,  wholesale:40,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p13", code:"1084063", name:"Activitae關節寶",                        price:219, wholesale:23,  vip_price:23, material:"沖泡粉",   spec:"30包x每包5g",     category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p15", code:"1084065", name:"Activitae女士寶第三代升級版",            price:228, wholesale:24,  vip_price:24, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p16", code:"1084066", name:"Activitae嬌妍寶",                        price:171, wholesale:18,  vip_price:18, material:"膠囊",     spec:"每盒30粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p17", code:"1084067", name:"Activitae雄風寶加強版",                  price:266, wholesale:28,  vip_price:28, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p18", code:"1084068", name:"Activitae酵益寶",                        price:40,  wholesale:32,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p19", code:"1084069", name:"Activitae 元氣寶",                       price:209, wholesale:22,  vip_price:22, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p49", code:"1084070", name:"Activitae 活腦營養素",                   price:200, wholesale:21,  vip_price:21, material:"沖泡粉",   spec:"30包x每包3G",     category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p29", code:"1084080", name:"Activitae 益生菌第二代升級版",           price:219, wholesale:23,  vip_price:23, material:"沖泡粉",   spec:"30包x每包3g",     category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p30", code:"1084081", name:"Activitae血糖安",                        price:171, wholesale:18,  vip_price:18, material:"口服液",   spec:"6支x30ML",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p31", code:"1084082", name:"Activitae 納豆紅麴+Q10",                 price:266, wholesale:28,  vip_price:28, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p32", code:"1084083", name:"Activitae 健肝寶加強版",                 price:228, wholesale:24,  vip_price:24, material:"口服液",   spec:"20mlx10瓶",       category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p33", code:"1084084", name:"Activitae 潤肺清喉咽",                   price:171, wholesale:18,  vip_price:18, material:"口服液",   spec:"30mlx6瓶",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p34", code:"1084085", name:"Activitae 長青護心高鈣脫脂奶粉",        price:304, wholesale:32,  vip_price:32, material:"沖泡粉",   spec:"30包x每包25g",    category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p35", code:"1084086", name:"Activitae 視力寶加強版",                 price:171, wholesale:18,  vip_price:18, material:"啫喱",     spec:"15包x每包15g",    category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p37", code:"1084088", name:"Activitae 美白凝凍加強版",               price:219, wholesale:23,  vip_price:23, material:"啫喱",     spec:"15條x每條15克",   category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p38", code:"1084089", name:"Activitae 美莓飲",                       price:95,  wholesale:10,  vip_price:10, material:"沖泡粉",   spec:"30包x每包3克",    category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p39", code:"1084090", name:"Activitae 尤加利茶",                     price:105, wholesale:11,  vip_price:11, material:"茶包",     spec:"2克x20包",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p40", code:"1084091", name:"Activitae 美莓粉包",                     price:19,  wholesale:15,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p41", code:"1084092", name:"Activitae 賦活寶加強版",                 price:456, wholesale:48,  vip_price:48, material:"沖泡粉",   spec:"3克x30包",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p42", code:"1084093", name:"Activitae 膠原蛋白",                     price:209, wholesale:22,  vip_price:22, material:"口服液",   spec:"6瓶x30毫升",      category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p43", code:"1084094", name:"Activitae 睡得寶",                       price:209, wholesale:22,  vip_price:22, material:"膠囊",     spec:"每盒60粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p44", code:"1084095", name:"Activitae 腸道寶",                       price:285, wholesale:30,  vip_price:30, material:"膠囊",     spec:"每盒30粒",        category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p45", code:"1084096", name:"Activitae 健胃寶",                       price:38,  wholesale:30,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p46", code:"1084097", name:"Activitae 活髮寶",                       price:32,  wholesale:26,  vip_price:0,  material:"",         spec:"",               category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p48", code:"1084099", name:"Activitae雄風寶勁能包",                  price:304, wholesale:32,  vip_price:32, material:"膠囊",     spec:"15包",            category:"Activitae", stock:99, image:"", desc:"" },
  { id:"p50", code:"1084100", name:"Activitae 御參元",                       price:257, wholesale:27,  vip_price:27, material:"沖泡粉",   spec:"15包x10ml",       category:"Activitae", stock:99, image:"", desc:"" },
  // ── ENERLAB ── (1084074 vip_price corrected to 29)
  { id:"p20", code:"1084071", name:"Enerlab超燃膠囊",                        price:266, wholesale:28,  vip_price:28, material:"膠囊",     spec:"每盒30粒",        category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p21", code:"1084072", name:"Enerlab超耐力沖飲包",                    price:266, wholesale:28,  vip_price:28, material:"沖泡粉",   spec:"30包x每包3G",     category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p22", code:"1084073", name:"Enerlab超極限能量粉",                    price:70,  wholesale:56,  vip_price:0,  material:"",         spec:"",               category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p23", code:"1084074", name:"Enerlab超護心納麴+Q10沖泡粉",            price:276, wholesale:29,  vip_price:29, material:"沖泡粉",   spec:"30包x每包3G",     category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p24", code:"1084075", name:"Enerlab超平衡祛濕粉",                    price:114, wholesale:12,  vip_price:12, material:"沖泡粉",   spec:"15包x每包10g",    category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p25", code:"1084076", name:"Enerlab超防禦保肝液",                    price:48,  wholesale:38,  vip_price:0,  material:"",         spec:"",               category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p26", code:"1084077", name:"Enerlab超益菌沖泡粉",                    price:171, wholesale:18,  vip_price:18, material:"沖泡粉",   spec:"30包x每包3g",     category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p27", code:"1084078", name:"Enerlab超關節沖泡粉",                    price:209, wholesale:22,  vip_price:22, material:"沖泡粉",   spec:"30包x每包5g",     category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p28", code:"1084079", name:"Enerlab超活妍膠原蛋白",                  price:48,  wholesale:38,  vip_price:0,  material:"",         spec:"",               category:"ENERLAB",   stock:99, image:"", desc:"" },
  { id:"p47", code:"1084098", name:"Enerlab 超活力冲飲",                     price:48,  wholesale:38,  vip_price:0,  material:"",         spec:"",               category:"ENERLAB",   stock:99, image:"", desc:"" },
  // ── BELSANTE ── (spec corrected for p10)
  { id:"p10", code:"1084043", name:"BELSANTE 健肝寶粉",                      price:190, wholesale:20,  vip_price:20, material:"沖泡粉",   spec:"30包x每包6g",     category:"BELSANTE",  stock:99, image:"", desc:"" },
  { id:"p11", code:"1084044", name:"BELSANTE 紅蔘寶粉",                      price:152, wholesale:16,  vip_price:16, material:"沖泡粉",   spec:"30包x每包5g",     category:"BELSANTE",  stock:99, image:"", desc:"" },
  { id:"p14", code:"1084064", name:"BELSANTE 增強免疫力之寶第二代",          price:247, wholesale:26,  vip_price:26, material:"沖泡粉",   spec:"30包x每包3g",     category:"BELSANTE",  stock:99, image:"", desc:"" },
  { id:"p36", code:"1084087", name:"Belsante 膠原蛋白美肌粉包第二代",       price:209, wholesale:22,  vip_price:22, material:"沖泡粉",   spec:"30包x每包8g",     category:"BELSANTE",  stock:99, image:"", desc:"" },
  // ── 護膚品 (new products) ──
  { id:"p61", code:"0300530", name:"SENBEL 柔和泡沫潔面乳",                  price:171, wholesale:18,  vip_price:18, material:"潔面乳",   spec:"150ml",           category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p62", code:"0300423", name:"SENBEL 淨化爽膚水",                      price:152, wholesale:16,  vip_price:16, material:"爽膚水",   spec:"100ml",           category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p63", code:"0300427", name:"SENBEL 淨化排毒面膜",                    price:171, wholesale:18,  vip_price:18, material:"面膜",     spec:"100ml",           category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p64", code:"0100345", name:"ESTEBEL 第六靈感 插牆式薰香器第二代 黃色", price:266, wholesale:28, vip_price:28, material:"薰香器",   spec:"0100345",         category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p65", code:"0100346", name:"ESTEBEL 第六靈感 插牆式薰香器第二代 紫色", price:266, wholesale:28, vip_price:28, material:"薰香器",   spec:"0100346",         category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p66", code:"0300405", name:"SENBEL 保濕精華面膜",                    price:95,  wholesale:10,  vip_price:10, material:"面膜",     spec:"每盒5片",         category:"護膚品",    stock:99, image:"", desc:"" },
  { id:"p67", code:"0300442", name:"SENBEL 再注氧潔面啫喱",                  price:114, wholesale:12,  vip_price:12, material:"啫喱",     spec:"150ml",           category:"護膚品",    stock:99, image:"", desc:"" },
];

// Get existing product codes to avoid overwriting
const snap = await getDocs(collection(db, "products"));
const existingIds = new Set(snap.docs.map(d => d.id));
const existingCodes = new Set(snap.docs.map(d => d.data().code).filter(Boolean));

console.log(`現有 Firestore 產品：${snap.size} 個\n`);

let added = 0, skipped = 0;
for (const p of allProducts) {
  if (existingIds.has(p.id)) {
    console.log(`  ⬜ ${p.id} (${p.code}) 已存在，跳過`);
    skipped++;
    continue;
  }
  if (existingCodes.has(p.code)) {
    console.log(`  ⬜ 貨號 ${p.code} 已存在（不同ID），跳過`);
    skipped++;
    continue;
  }
  await setDoc(doc(db, "products", p.id), p);
  console.log(`  ✅ 已新增 ${p.id} (${p.code}) — ${p.name}`);
  added++;
}

console.log(`\n完成：新增 ${added} 個，跳過 ${skipped} 個`);
process.exit(0);
