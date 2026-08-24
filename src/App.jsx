import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  LayoutDashboard, Package, ShoppingCart, Users, BarChart3,
  Plus, Trash2, Pencil, X, Search, Store, Globe,
  TrendingUp, AlertTriangle, Loader2, ChevronDown, ChevronRight,
  ArrowDownToLine, ArrowUpFromLine, Barcode, ImagePlus, ImageOff, Check, Printer, RotateCcw
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, PieChart, Pie, Cell
} from "recharts";

/* ---------------------------------------------------------------
   Sổ Bán Hàng — phỏng theo cấu trúc sheet "HILITEK"
   Giai đoạn 1: Sản phẩm & Tồn kho (Mã VT, tồn đầu/nhập/xuất/cuối kỳ,
   giá xuất bình quân gia quyền) + quản lý Series gắn theo sản phẩm.
   Palette: ink navy / brass / parchment-grey / forest / rust
--------------------------------------------------------------- */

const STORAGE_KEY = "solbh-data-v2";

const INK = "#1F2A44";
const BRASS = "#B4863F";
const PAPER = "#F1F0EA";
const FOREST = "#3F6B52";
const RUST = "#B0462F";
const BLUE = "#3E6FA6";
const LINE = "#D8D3C4";

const CHANNELS = [
  { id: "store", label: "Tại cửa hàng", icon: Store },
  { id: "online", label: "Online", icon: Globe },
];
const STATUSES = [
  { id: "pending", label: "Chờ xử lý", color: BRASS },
  { id: "shipping", label: "Đang giao", color: BLUE },
  { id: "done", label: "Hoàn tất", color: FOREST },
  { id: "cancelled", label: "Đã huỷ", color: RUST },
];
const UNITS = ["Bộ", "Cái", "Hộp", "Thùng", "Chiếc"];
const VAT_OPTIONS = [
  { id: "KCT", label: "KCT" },
  { id: "VAT0", label: "VAT 0%" },
  { id: "VAT8", label: "VAT 8%" },
  { id: "VAT10", label: "VAT 10%" },
];
const SKU_PREFIX = "HI";
const PO_STATUSES = [
  { id: "pending", label: "Chờ giao", color: BRASS },
  { id: "received", label: "Đã nhập", color: FOREST },
];
const PO_PREFIX = "POH";
const BRANCHES = ["Kho tổng"];
const EMPLOYEES = ["Chủ cửa hàng"];
const PAYMENT_METHODS = [
  { id: "cash", label: "Tiền mặt" },
  { id: "credit", label: "Công nợ" },
];
const NCC_PREFIX = "NCC";
const KH_PREFIX = "KH";
function nextCustomerCode(customers) {
  let max = 0;
  customers.forEach((c) => { const m = /^KH(\d+)$/.exec(c.code || ""); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return KH_PREFIX + String(max + 1).padStart(3, "0");
}
// 34 tỉnh/thành phố theo địa giới hành chính mới (sau sáp nhập, hiệu lực từ 1/7/2025; Đồng Nai lên thành phố 30/4/2026)
const VN_PROVINCES = [
  "Thành phố Hà Nội", "Thành phố Hải Phòng", "Thành phố Đà Nẵng", "Thành phố Đồng Nai", "Thành phố Huế",
  "Thành phố Hồ Chí Minh", "Thành phố Cần Thơ",
  "Tỉnh Cao Bằng", "Tỉnh Điện Biên", "Tỉnh Hà Tĩnh", "Tỉnh Lai Châu", "Tỉnh Lạng Sơn", "Tỉnh Nghệ An",
  "Tỉnh Quảng Ninh", "Tỉnh Thanh Hóa", "Tỉnh Sơn La", "Tỉnh Tuyên Quang", "Tỉnh Lào Cai", "Tỉnh Thái Nguyên",
  "Tỉnh Phú Thọ", "Tỉnh Bắc Ninh", "Tỉnh Hưng Yên", "Tỉnh Ninh Bình", "Tỉnh Quảng Trị", "Tỉnh Quảng Ngãi",
  "Tỉnh Gia Lai", "Tỉnh Khánh Hòa", "Tỉnh Lâm Đồng", "Tỉnh Đắk Lắk", "Tỉnh Tây Ninh", "Tỉnh Vĩnh Long",
  "Tỉnh Đồng Tháp", "Tỉnh Cà Mau", "Tỉnh An Giang",
];

// Thông tin công ty in trên hoá đơn — chỉnh lại tại đây nếu công ty đổi thông tin
const COMPANY_INFO = {
  name: "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ HILI",
  address: "6/27A Đường Số 3, C/x Lữ Gia, Phường Phú Thọ, TP Hồ Chí Minh, Việt Nam",
  taxCode: "0316296138",
  bankAccount: "19551097 - Ngân Hàng Á Châu ACB – phòng giao dịch Lý Thường Kiệt",
  phone: "0939206865",
};

const CHU_SO_VN = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
function docBaChuSoVN(so, daydu) {
  const tram = Math.floor(so / 100), chuc = Math.floor((so % 100) / 10), donvi = so % 10;
  let s = "";
  if (tram > 0 || daydu) {
    s += CHU_SO_VN[tram] + " trăm";
    if (chuc === 0 && donvi > 0) s += " linh";
  }
  if (chuc > 1) {
    s += " " + CHU_SO_VN[chuc] + " mươi";
    if (donvi === 1) s += " mốt"; else if (donvi === 5) s += " lăm"; else if (donvi > 0) s += " " + CHU_SO_VN[donvi];
  } else if (chuc === 1) {
    s += " mười";
    if (donvi === 1) s += " một"; else if (donvi === 5) s += " lăm"; else if (donvi > 0) s += " " + CHU_SO_VN[donvi];
  } else if (chuc === 0 && donvi > 0) {
    s += (tram > 0 || daydu ? " " : "") + CHU_SO_VN[donvi];
  }
  return s.trim();
}
function soTienBangChu(num) {
  num = Math.round(Math.abs(num || 0));
  if (num === 0) return "Không đồng";
  const donVi = ["", "nghìn", "triệu", "tỷ"];
  const nhom = [];
  let n = num;
  while (n > 0) { nhom.unshift(n % 1000); n = Math.floor(n / 1000); }
  const parts = [];
  nhom.forEach((g, i) => {
    if (g === 0) return;
    const isFirst = i === 0;
    const words = docBaChuSoVN(g, !isFirst);
    const dv = donVi[nhom.length - 1 - i];
    parts.push(words + (dv ? " " + dv : ""));
  });
  let result = parts.join(" ").replace(/\s+/g, " ").trim();
  result = result.charAt(0).toUpperCase() + result.slice(1);
  return result + " đồng";
}

const SUPPLIER_PAYMENT_TERMS = [
  { id: "cash", label: "TM (Tiền mặt)" },
  { id: "credit", label: "Công nợ" },
];
function nextSupplierCode(suppliers) {
  let max = 0;
  suppliers.forEach((s) => {
    const m = /^NCC(\d+)$/.exec(s.code || "");
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return NCC_PREFIX + String(max + 1).padStart(3, "0");
}

function nextSKU(products) {
  let max = 0;
  products.forEach((p) => {
    const m = /^HI(\d+)$/.exec(p.sku || "");
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return SKU_PREFIX + String(max + 1).padStart(3, "0");
}
function nextPOCode(purchaseOrders) {
  let max = 0;
  purchaseOrders.forEach((po) => {
    const m = /^POH(\d+)$/.exec(po.code || "");
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return PO_PREFIX + String(max + 1).padStart(3, "0");
}
function formatDateTime(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso || "";
  return d.toLocaleDateString("vi-VN") + " " + d.toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" });
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const vnd = (n) => (Math.round(Number(n)) || 0).toLocaleString("vi-VN") + "đ";
const todayISO = () => new Date().toISOString().slice(0, 10);
const parseSeries = (text) =>
    text.split(/[\s,]+/).map((s) => s.trim()).filter(Boolean);

/* ---------------- derived stock helpers (Tồn đầu / Nhập / Xuất / Tồn cuối) ---------------- */
function productStats(p) {
  const inMoves = p.movements.filter((m) => m.type === "in");
  const outMoves = p.movements.filter((m) => m.type === "out");
  const importedQty = inMoves.reduce((s, m) => s + m.qty, 0);
  const importedValue = inMoves.reduce((s, m) => s + m.qty * m.price, 0);
  const exportedQty = outMoves.reduce((s, m) => s + m.qty, 0);
  const exportedValue = outMoves.reduce((s, m) => s + m.qty * m.price, 0);
  const closingQty = p.openingQty + importedQty - exportedQty;
  // Giá nhập là số bạn tự nhập trực tiếp trên sản phẩm (product.costPrice) — không tự tính bình quân.
  const avgCost = p.costPrice || 0;
  return { importedQty, importedValue, exportedQty, exportedValue, closingQty, avgCost };
}

// flatten every product's movements into a series list: {serial, code, name, importDoc, importDate, exportDoc, exportDate, status}
function seriesList(p) {
  const rows = [];
  const exported = new Set();
  p.movements
      .filter((m) => m.type === "out")
      .forEach((m) => (m.series || []).forEach((s) => exported.add(s)));
  p.movements
      .filter((m) => m.type === "in")
      .forEach((m) => {
        (m.series || []).forEach((s) => {
          const outMove = p.movements.find((mo) => mo.type === "out" && (mo.series || []).includes(s));
          rows.push({
            serial: s,
            importDoc: m.docNo, importDate: m.date,
            exportDoc: outMove?.docNo || "", exportDate: outMove?.date || "",
            status: outMove ? "Đã xuất" : "Còn tồn",
          });
        });
      });
  return rows;
}

function seedData() {
  const win11 = uid(), office = uid(), khungTivi = uid();
  return {
    products: [
      {
        id: win11, code: "Win11P", name: "Phần mềm Win Pro 11 64Bit Eng Intl 1pk DSP OEI (FQC-10528)",
        unit: "Bộ", category: "Phần mềm", hasSeries: true, retailPrice: 1650000, wholesalePrice: 1500000, costPrice: 1100000,
        sku: "HI001", vat: "VAT10", barcode: "", image: null,
        openingQty: 0,
        movements: [],
      },
      {
        id: office, code: "Office21PP", name: "Phần mềm Office Professional Plus 2021 English APAC EM Medialess",
        unit: "Bộ", category: "Phần mềm", hasSeries: true, retailPrice: 2650000, wholesalePrice: 2450000, costPrice: 2200000,
        sku: "HI002", vat: "VAT10", barcode: "", image: null,
        openingQty: 0,
        movements: [],
      },
      {
        id: khungTivi, code: "E2600", name: "Khung treo Tivi di động E2600",
        unit: "Cái", category: "Gia dụng", hasSeries: false, retailPrice: 990000, wholesalePrice: 890000, costPrice: 800000,
        sku: "HI003", vat: "VAT10", barcode: "", image: null,
        openingQty: 0,
        movements: [],
      },
    ],
    customers: [
      { id: uid(), name: "Nguyễn Thị Lan", phone: "0901 234 567", note: "Khách quen" },
      { id: uid(), name: "Trần Văn Minh", phone: "0912 345 678", note: "" },
    ],
    orders: [],
  };
}

async function loadData() {
  try {
    const res = await window.storage.get(STORAGE_KEY, false);
    if (res && res.value) return JSON.parse(res.value);
  } catch (e) { /* chưa có dữ liệu */ }
  return null;
}
async function saveData(data) {
  try { await window.storage.set(STORAGE_KEY, JSON.stringify(data), false); }
  catch (e) { console.error("Lỗi lưu dữ liệu:", e); }
}

// Đảm bảo mọi sản phẩm tải từ bộ nhớ đều có đủ field cần thiết,
// tránh lỗi trắng trang khi dữ liệu cũ (trước khi có SKU/VAT/ảnh/giá sỉ...) được nạp lại.
function normalizeProduct(p) {
  return {
    id: p.id || uid(),
    code: p.code || "",
    name: p.name || "",
    unit: p.unit || UNITS[0],
    category: p.category || "",
    hasSeries: !!p.hasSeries,
    retailPrice: Number(p.retailPrice ?? p.sellPrice ?? 0) || 0,
    wholesalePrice: Number(p.wholesalePrice ?? p.sellPrice ?? 0) || 0,
    costPrice: Number(p.costPrice ?? 0) || 0,
    openingQty: Number(p.openingQty ?? 0) || 0,
    sku: p.sku || "",
    vat: p.vat || "VAT10",
    barcode: p.barcode || "",
    image: p.image || null,
    movements: Array.isArray(p.movements) ? p.movements.map((m) => ({
      id: m.id || uid(), type: m.type === "out" ? "out" : "in",
      docNo: m.docNo ?? "", date: m.date || todayISO(),
      qty: Number(m.qty) || 0, price: Number(m.price) || 0,
      series: Array.isArray(m.series) ? m.series : [],
    })) : [],
  };
}
function normalizeOrder(o) {
  const createdAt = o.createdAt || new Date().toISOString();
  const status = STATUSES.some((s) => s.id === o.status) ? o.status : "pending";
  const paidAmount = Number(o.paidAmount) || 0;
  const items = Array.isArray(o.items) ? o.items.map((it) => ({
    productId: it.productId, qty: Number(it.qty) || 1, price: Number(it.price) || 0,
    series: Array.isArray(it.series) ? it.series : [],
  })) : [];
  const subtotal = items.reduce((s, it) => s + orderLineTotal(it), 0);
  const payable = subtotal - (Number(o.orderDiscount) || 0) + (Number(o.shippingFee) || 0);
  const isFullyPaid = paidAmount >= payable && payable > 0;
  return {
    id: o.id || uid(), code: o.code || "", createdAt,
    customerId: o.customerId || "", channel: o.channel === "online" ? "online" : "store",
    branch: o.branch || BRANCHES[0], seller: o.seller || EMPLOYEES[0], deliveryDate: o.deliveryDate || "",
    tags: Array.isArray(o.tags) ? o.tags : [], notes: o.notes || "",
    status, date: o.date || todayISO(),
    shippingAt: o.shippingAt || ((status === "shipping" || status === "done") ? createdAt : null),
    deliveredAt: o.deliveredAt || (status === "done" ? createdAt : null),
    paidCompleteAt: o.paidCompleteAt || (isFullyPaid ? createdAt : null),
    cancelledAt: o.cancelledAt || (status === "cancelled" ? createdAt : null),
    items,
    vat: o.vat || "VAT10", orderDiscount: Number(o.orderDiscount) || 0, shippingFee: Number(o.shippingFee) || 0, paidAmount,
    returns: Array.isArray(o.returns) ? o.returns.map((r) => ({
      id: r.id || uid(), code: r.code || "", createdAt: r.createdAt || createdAt, type: r.type === "exchange" ? "exchange" : "refund", note: r.note || "",
      returnedItems: Array.isArray(r.returnedItems) ? r.returnedItems : [],
      exchangeItems: Array.isArray(r.exchangeItems) ? r.exchangeItems : [],
    })) : [],
  };
}
function vatPercent(vatId) { return { KCT: 0, VAT0: 0, VAT8: 8, VAT10: 10 }[vatId] ?? 0; }
function orderLineTotal(it) { return it.qty * it.price; }
function returnLineTotal(it) { return it.qty * it.price; }
function orderCalc(o) {
  const subtotal = o.items.reduce((s, it) => s + orderLineTotal(it), 0);
  const vp = vatPercent(o.vat);
  const vatTotal = (subtotal * vp) / (100 + vp);
  const returns = o.returns || [];
  const returnedValue = returns.reduce((s, r) => s + r.returnedItems.reduce((s2, it) => s2 + returnLineTotal(it), 0), 0);
  const exchangeValue = returns.reduce((s, r) => s + r.exchangeItems.reduce((s2, it) => s2 + returnLineTotal(it), 0), 0);
  const payable = subtotal - (o.orderDiscount || 0) + (o.shippingFee || 0) - returnedValue + exchangeValue;
  const remaining = payable - (o.paidAmount || 0);
  return { subtotal, vatTotal, returnedValue, exchangeValue, payable, remaining };
}
// SL đã trả của 1 sản phẩm trong đơn (cộng dồn qua các phiếu đổi trả đã tạo)
function returnedQtyOf(order, productId) {
  return (order.returns || []).reduce((s, r) => s + r.returnedItems.filter((it) => it.productId === productId).reduce((s2, it) => s2 + it.qty, 0), 0);
}
function returnedSeriesOf(order, productId) {
  const out = [];
  (order.returns || []).forEach((r) => r.returnedItems.filter((it) => it.productId === productId).forEach((it) => out.push(...it.series)));
  return out;
}
function nextOrderCode(orders) {
  let max = 0;
  orders.forEach((o) => { const m = /^DH(\d+)$/.exec(o.code || ""); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return "DH" + String(max + 1).padStart(3, "0");
}
function nextReturnCode(order) {
  return `${order.code}-RT${String((order.returns || []).length + 1).padStart(2, "0")}`;
}
const CUSTOMER_GROUPS = [
  { id: "retail", label: "KH Lẻ" },
  { id: "b2b", label: "B2B" },
  { id: "enterprise", label: "Doanh nghiệp" },
];
function normalizeCustomer(c) {
  return {
    id: c.id || uid(), code: c.code || "", name: c.name || "", phone: c.phone || "", note: c.note || "",
    email: c.email || "", taxCode: c.taxCode || "", province: c.province || "", ward: c.ward || "", addressDetail: c.addressDetail || "",
    group: CUSTOMER_GROUPS.some((g) => g.id === c.group) ? c.group : "retail",
  };
}
function normalizeSupplier(s) {
  return {
    id: s.id || uid(), code: s.code || "", name: s.name || "", taxCode: s.taxCode || "",
    address: s.address || "", contactPerson: s.contactPerson || "", phone: s.phone || "", email: s.email || "",
    paymentTerm: SUPPLIER_PAYMENT_TERMS.some((t) => t.id === s.paymentTerm) ? s.paymentTerm : "cash",
    creditDays: Number(s.creditDays) || 0,
  };
}
function normalizePO(po) {
  const createdAt = po.createdAt || new Date().toISOString();
  const status = PO_STATUSES.some((s) => s.id === po.status) ? po.status : "pending";
  const paid = !!po.paid;
  return {
    id: po.id || uid(), code: po.code || "", createdAt,
    status,
    receivedAt: po.receivedAt || (status === "received" ? createdAt : null),
    branch: po.branch || BRANCHES[0], supplier: po.supplier || "", createdBy: po.createdBy || "",
    invoiceNo: po.invoiceNo || "", notes: po.notes || "", tags: Array.isArray(po.tags) ? po.tags : [],
    paid, paidAt: po.paidAt || (paid ? createdAt : null), paymentMethod: po.paymentMethod || (paid ? "cash" : "credit"),
    items: Array.isArray(po.items) ? po.items.map((it) => ({
      productId: it.productId, qty: Number(it.qty) || 0, price: Number(it.price) || 0,
      vat: it.vat || "VAT10",
      series: Array.isArray(it.series) ? it.series : [],
    })) : [],
  };
}

// Ghi nhận hàng của 1 đơn nhập vào tồn kho sản phẩm (tạo movement "in" cho từng dòng hàng).
function applyPOToStock(po, setProducts) {
  setProducts((prev) => prev.map((p) => {
    const it = po.items.find((i) => i.productId === p.id);
    if (!it) return p;
    return { ...p, movements: [...p.movements, { id: uid(), type: "in", docNo: po.code, date: po.createdAt.slice(0, 10), qty: it.qty, price: it.price, series: it.series }] };
  }));
}

class AppErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  componentDidCatch(error, info) { console.error("Lỗi hiển thị:", error, info); }
  render() {
    if (this.state.error) {
      return (
          <div className="p-6 text-sm" style={{ color: RUST, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "pre-wrap" }}>
            Đã xảy ra lỗi khi hiển thị: {String(this.state.error?.message || this.state.error)}
          </div>
      );
    }
    return this.props.children;
  }
}

/* ---------------- shared UI bits ---------------- */

function Stamp({ status }) {
  const s = STATUSES.find((x) => x.id === status) || STATUSES[0];
  return (
      <span style={{ borderColor: s.color, color: s.color, fontFamily: "'IBM Plex Mono', monospace" }}
            className="inline-block border-2 rounded px-2 py-0.5 text-[11px] uppercase tracking-wider -rotate-2 select-none">
      {s.label}
    </span>
  );
}
function Field({ label, children, hint }) {
  return (
      <label className="block mb-3">
        <span className="block text-xs uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.6 }}>{label}</span>
        {children}
        {hint && <span className="block text-[11px] mt-1 opacity-50">{hint}</span>}
      </label>
  );
}
const inputCls = "w-full bg-transparent border-b-2 outline-none py-1.5 px-1 text-[15px] focus:border-opacity-100";

function TagsNotesCompact({ tags, setTags, notes, setNotes }) {
  return (
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-4 pt-3" style={{ borderTop: `1px dashed ${LINE}` }}>
        <div>
          <span className="block text-[10px] uppercase tracking-wider mb-1 opacity-45">Tags</span>
          <SeriesTagInput series={tags} setSeries={setTags} placeholder="Gõ rồi cách khoảng trắng…" />
        </div>
        <div>
          <span className="block text-[10px] uppercase tracking-wider mb-1 opacity-45">Ghi chú</span>
          <textarea rows={1} className="w-full border rounded-sm p-2 text-xs" style={{ borderColor: LINE }} value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Ghi chú thêm…" />
        </div>
      </div>
  );
}

function Modal({ title, onClose, children, wide, size }) {
  const sizeClass = { md: "max-w-md", lg: "max-w-lg", xl: "max-w-3xl", "2xl": "max-w-6xl" }[size] || (wide ? "max-w-lg" : "max-w-md");
  return (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4" style={{ background: "rgba(31,42,68,0.45)" }} onClick={onClose}>
        <div onClick={(e) => e.stopPropagation()}
             className={`w-full ${sizeClass} rounded-sm shadow-2xl relative flex flex-col`}
             style={{ background: PAPER, border: `1px solid ${LINE}`, maxHeight: "88vh" }}>
          <div className="flex items-center justify-between px-4 sm:px-6 pt-4 sm:pt-5 pb-3 shrink-0">
            <h3 className="text-lg sm:text-xl pr-6" style={{ fontFamily: "'Fraunces', serif", color: INK }}>{title}</h3>
            <button onClick={onClose} className="absolute top-4 right-4 opacity-60 hover:opacity-100" style={{ color: INK }}><X size={18} /></button>
          </div>
          <div className="px-4 sm:px-6 pb-4 sm:pb-6 overflow-y-auto min-w-0" style={{ flex: "1 1 auto" }}>
            {children}
          </div>
        </div>
      </div>
  );
}

/* ---------------- Dashboard ---------------- */

function Dashboard({ products, orders, customers }) {
  const [period, setPeriod] = useState("month");
  const [customFrom, setCustomFrom] = useState(todayISO());
  const [customTo, setCustomTo] = useState(todayISO());

  const range = useMemo(() => {
    const now = new Date();
    const startOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const endOfDay = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23,59,59);
    let from, to;
    if (period === "today") { from = startOfDay(now); to = endOfDay(now); }
    else if (period === "week") { const d = new Date(now); d.setDate(now.getDate()-6); from = startOfDay(d); to = endOfDay(now); }
    else if (period === "month") { from = new Date(now.getFullYear(), now.getMonth(), 1); to = endOfDay(now); }
    else if (period === "year") { from = new Date(now.getFullYear(), 0, 1); to = endOfDay(now); }
    else if (period === "custom") { from = new Date(customFrom); to = new Date(customTo); to.setHours(23,59,59); }
    else { from = new Date(2000,0,1); to = endOfDay(now); }
    return { from, to };
  }, [period, customFrom, customTo]);

  const filteredOrders = useMemo(() => {
    return orders.filter(o => {
      const d = new Date(o.createdAt);
      return d >= range.from && d <= range.to && o.status !== "cancelled";
    });
  }, [orders, range]);

  const stockValue = products.reduce((s, p) => s + productStats(p).closingQty * productStats(p).avgCost, 0);
  const totalInAll = products.reduce((s, p) => s + productStats(p).importedValue, 0);
  const totalOutAll = products.reduce((s, p) => s + productStats(p).exportedValue, 0);

  const revenue = filteredOrders.reduce((s,o)=> s + orderCalc(o).payable, 0);
  const costOfGoods = filteredOrders.reduce((s,o)=> {
    return s + o.items.reduce((s2,it)=>{
      const p = products.find(x=>x.id===it.productId);
      return s2 + it.qty * (p?.costPrice||0);
    },0);
  },0);
  const profit = revenue - costOfGoods;
  const debt = orders.filter(o=>o.status!=="cancelled").reduce((s,o)=> s+ Math.max(0, orderCalc(o).remaining),0);
  const orderCount = filteredOrders.length;
  const customerCount = customers?.length || 0;

  const topProducts = useMemo(() => {
    const map={};
    filteredOrders.forEach(o=>{
      o.items.forEach(it=>{
        const p = products.find(x=>x.id===it.productId);
        if(!p) return;
        map[p.id] = map[p.id]||{ name:p.name, code:p.code, total:0, qty:0 };
        map[p.id].total += it.qty*it.price;
        map[p.id].qty += it.qty;
      });
    });
    return Object.values(map).sort((a,b)=>b.total-a.total).slice(0,5);
  }, [filteredOrders, products]);

  const chartData = useMemo(() => {
    const fmt = (d) => {
      if(period==="year") return `${d.getMonth()+1}/${d.getFullYear()}`;
      return d.toLocaleDateString("vi-VN", { day:"2-digit", month:"2-digit" });
    };
    const buckets={};
    filteredOrders.forEach(o=>{
      const d = new Date(o.createdAt);
      const key = period==="year" ? `${d.getFullYear()}-${d.getMonth()}` : d.toISOString().slice(0,10);
      const label = fmt(d);
      if(!buckets[key]) buckets[key]={ label, revenue:0, profit:0, orders:0 };
      const oc = orderCalc(o);
      buckets[key].revenue += oc.payable;
      const c = o.items.reduce((s2,it)=>{ const p=products.find(x=>x.id===it.productId); return s2+it.qty*(p?.costPrice||0); },0);
      buckets[key].profit += oc.payable - c;
      buckets[key].orders +=1;
    });
    return Object.values(buckets).sort((a,b)=> a.label.localeCompare(b.label)).slice(-14);
  }, [filteredOrders, period, products]);

  const lowStock = products.filter((p) => productStats(p).closingQty <= 5);
  const pieColors = [INK, BRASS, FOREST, BLUE, RUST];

  return (
      <div>
        <div className="flex flex-wrap items-center gap-2 mb-6 p-3 rounded-sm" style={{ background:"#fff", border:`1px solid ${LINE}` }}>
          <span className="text-xs uppercase tracking-wider opacity-60 mr-1">Xem doanh thu:</span>
          {[
            { id:"today", label:"Hôm nay" },
            { id:"week", label:"7 ngày" },
            { id:"month", label:"Tháng này" },
            { id:"year", label:"Năm nay" },
            { id:"custom", label:"Tùy chọn" },
          ].map(p=>(
              <button key={p.id} onClick={()=>setPeriod(p.id)} className="text-xs px-3 py-1.5 rounded-full border"
                      style={{ borderColor: period===p.id?INK:LINE, background: period===p.id?INK:"transparent", color: period===p.id?"#fff":INK }}>{p.label}</button>
          ))}
          {period==="custom" && (
              <div className="flex items-center gap-2 ml-2">
                <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)} className="border rounded-sm px-2 py-1 text-xs" style={{ borderColor:LINE }} />
                <span className="text-xs opacity-50">đến</span>
                <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)} className="border rounded-sm px-2 py-1 text-xs" style={{ borderColor:LINE }} />
              </div>
          )}
          <span className="ml-auto text-[11px] opacity-50" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{range.from.toLocaleDateString("vi-VN")} - {range.to.toLocaleDateString("vi-VN")}</span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
          <StatCard label={`Doanh thu (${period==="today"?"hôm nay":period==="week"?"7 ngày":period==="month"?"tháng này":period==="year"?"năm nay":"tùy chọn"})`} value={vnd(revenue)} sub={`${orderCount} đơn`} icon={TrendingUp} accent={FOREST} />
          <StatCard label="Lợi nhuận gộp (ước tính)" value={vnd(profit)} sub={`Giá vốn ${vnd(costOfGoods)}`} icon={BarChart3} accent={BRASS} />
          <StatCard label="Công nợ khách hàng" value={vnd(debt)} sub={`${customerCount} khách hàng`} icon={Users} accent={RUST} />
          <StatCard label="Giá trị tồn kho hiện tại" value={vnd(stockValue)} sub={`${products.length} SKU`} icon={Package} accent={BLUE} />
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-5 gap-6 mb-6">
          <div className="lg:col-span-3 p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6 }}>Doanh thu theo {period==="year"?"tháng":"ngày"} - {vnd(revenue)}</h4>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
                <XAxis dataKey="label" tick={{ fontSize: 11, fill: INK }} />
                <YAxis tick={{ fontSize: 11, fill: INK }} width={45} tickFormatter={(v)=>`${(v/1000).toFixed(0)}k`} />
                <Tooltip formatter={(v)=>vnd(v)} contentStyle={{ fontFamily: "Inter", fontSize: 13 }} />
                <Bar dataKey="revenue" name="Doanh thu" fill={FOREST} radius={[4,4,0,0]} />
                <Bar dataKey="profit" name="Lợi nhuận" fill={BRASS} radius={[4,4,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="lg:col-span-2 p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6 }}>Top sản phẩm bán chạy (kỳ đã chọn)</h4>
            {topProducts.length === 0 ? <p className="text-sm opacity-60 py-8 text-center">Chưa có dữ liệu trong kỳ.</p> : (
                <div>
                  <ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie data={topProducts} dataKey="total" nameKey="name" innerRadius={45} outerRadius={75}>
                        {topProducts.map((_, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
                      </Pie>
                      <Tooltip formatter={(v) => vnd(v)} contentStyle={{ fontFamily: "Inter", fontSize: 13 }} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="mt-2 space-y-1.5">
                    {topProducts.map((p,i)=>(
                        <div key={p.code} className="flex items-center justify-between text-xs py-1" style={{ borderBottom:`1px dashed ${LINE}` }}>
                          <span className="flex items-center gap-2"><span className="w-2.5 h-2.5 rounded-full" style={{ background:pieColors[i%pieColors.length] }} />{p.code} · {p.name.slice(0,28)}</span>
                          <span style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{p.qty} · {vnd(p.total)}</span>
                        </div>
                    ))}
                  </div>
                </div>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6 }}>Tồn kho theo sản phẩm (toàn hệ thống)</h4>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={products.map((p) => ({ name: p.code, ton: productStats(p).closingQty }))}>
                <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
                <XAxis dataKey="name" tick={{ fontSize: 11, fill: INK }} />
                <YAxis tick={{ fontSize: 11, fill: INK }} width={30} />
                <Tooltip contentStyle={{ fontFamily: "Inter", fontSize: 13 }} />
                <Bar dataKey="ton" fill={INK} radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-2" style={{ color: INK, opacity: 0.6 }}>Tổng quan kho</h4>
            <div className="grid grid-cols-3 gap-3 mb-4">
              <div className="p-3 rounded-sm text-center" style={{ background:PAPER, border:`1px solid ${LINE}` }}><p className="text-[10px] uppercase opacity-50">Tổng nhập (all)</p><p className="text-sm font-bold mt-1" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(totalInAll)}</p></div>
              <div className="p-3 rounded-sm text-center" style={{ background:PAPER, border:`1px solid ${LINE}` }}><p className="text-[10px] uppercase opacity-50">Tổng xuất (all)</p><p className="text-sm font-bold mt-1" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(totalOutAll)}</p></div>
              <div className="p-3 rounded-sm text-center" style={{ background:PAPER, border:`1px solid ${LINE}` }}><p className="text-[10px] uppercase opacity-50">Tồn kho</p><p className="text-sm font-bold mt-1" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(stockValue)}</p></div>
            </div>
            <div className="text-xs opacity-60">Kỳ lọc chỉ áp dụng cho doanh thu/đơn hàng. Tồn kho luôn tính toàn bộ hệ thống.</div>
          </div>
        </div>

        {lowStock.length > 0 && (
            <div className="mt-6 p-4 rounded-sm flex items-start gap-3" style={{ background: "#FBF0EC", border: `1px solid ${RUST}55` }}>
              <AlertTriangle size={18} style={{ color: RUST }} className="mt-0.5 shrink-0" />
              <div className="text-sm" style={{ color: INK }}>
                <span className="font-medium">Tồn kho thấp (≤5): </span>
                {lowStock.map((p) => `${p.name} (còn ${productStats(p).closingQty})`).join(", ")}
              </div>
            </div>
        )}
      </div>
  );
}
function StatCard({ label, value, sub, icon: Icon, accent }) {
  return (
      <div className="p-5 rounded-sm flex items-center justify-between" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
        <div>
          <p className="text-xs uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55 }}>{label}</p>
          <p className="text-xl sm:text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{value}</p>
          {sub && <p className="text-[11px] opacity-50 mt-1">{sub}</p>}
        </div>
        <div className="w-10 h-10 rounded-full flex items-center justify-center shrink-0" style={{ background: `${accent}1A` }}>
          <Icon size={18} style={{ color: accent }} />
        </div>
      </div>
  );
}

/* ---------------- Products & Inventory (Sản phẩm & Tồn kho) ---------------- */

function ProductsInventory({ products, setProducts }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null); // sản phẩm đang thêm/sửa thông tin
  const [form, setForm] = useState({});
  const [ioModal, setIoModal] = useState(null); // { product, type: 'in'|'out' }
  const [ioForm, setIoForm] = useState({});
  const [viewingId, setViewingId] = useState(null); // id sản phẩm đang xem chi tiết

  const filtered = products.filter(
      (p) => p.name.toLowerCase().includes(query.toLowerCase()) || p.code.toLowerCase().includes(query.toLowerCase())
  );

  const openNew = () => { setForm({ code: "", name: "", unit: UNITS[0], category: "", hasSeries: false, retailPrice: "", wholesalePrice: "", costPrice: "", openingQty: 0, sku: nextSKU(products), vat: "VAT10", barcode: "", image: null }); setEditing({}); };
  const openEdit = (p) => { setForm({ ...p }); setEditing(p); };
  const submitInfo = () => {
    if (!form.code || !form.name) return;
    if (editing.id) {
      setProducts((prev) => prev.map((p) => (p.id === editing.id ? {
        ...p, code: form.code, name: form.name, unit: form.unit, category: form.category || "", hasSeries: !!form.hasSeries,
        retailPrice: Number(form.retailPrice) || 0, wholesalePrice: Number(form.wholesalePrice) || 0, costPrice: Number(form.costPrice) || 0, openingQty: Number(form.openingQty) || 0,
        sku: form.sku || p.sku, vat: form.vat, barcode: form.barcode || "", image: form.image || null,
      } : p)));
    } else {
      setProducts((prev) => [...prev, {
        id: uid(), code: form.code, name: form.name, unit: form.unit, category: form.category || "",
        hasSeries: !!form.hasSeries, retailPrice: Number(form.retailPrice) || 0, wholesalePrice: Number(form.wholesalePrice) || 0, costPrice: Number(form.costPrice) || 0,
        openingQty: Number(form.openingQty) || 0,
        sku: form.sku || nextSKU(products), vat: form.vat || "VAT10", barcode: form.barcode || "", image: form.image || null,
        movements: [],
      }]);
    }
    setEditing(null);
  };
  const onImagePick = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 1.5 * 1024 * 1024) { alert("Ảnh quá lớn — vui lòng chọn ảnh dưới 1.5MB."); return; }
    const dataUrl = await fileToDataUrl(file);
    setForm((f) => ({ ...f, image: dataUrl }));
  };
  const removeProduct = (id) => setProducts((prev) => prev.filter((p) => p.id !== id));

  const openIO = (product, type) => {
    setIoForm({ docNo: "", date: todayISO(), qty: "", price: type === "out" ? product.retailPrice : product.costPrice || "", priceLevel: "retail", series: [], selectedSeries: [] });
    setIoModal({ product, type });
  };

  const availableSeries = ioModal?.type === "out" ? seriesList(ioModal.product).filter((s) => s.status === "Còn tồn") : [];

  const submitIO = () => {
    const { product, type } = ioModal;
    const qty = Number(ioForm.qty);
    const price = Number(ioForm.price);
    if (!ioForm.docNo || !qty || qty <= 0 || !price) return;

    let series = [];
    if (product.hasSeries) {
      if (type === "in") {
        series = ioForm.series;
        if (series.length !== qty) {
          alert(`Sản phẩm này quản lý theo series — cần nhập đúng ${qty} số series (đang có ${series.length}).`);
          return;
        }
      } else {
        series = ioForm.selectedSeries;
        if (series.length !== qty) {
          alert(`Sản phẩm này quản lý theo series — cần chọn đúng ${qty} số series còn tồn (đang chọn ${series.length}).`);
          return;
        }
      }
    } else if (type === "out") {
      const { closingQty } = productStats(product);
      if (qty > closingQty) {
        alert(`Chỉ còn tồn ${closingQty} — không thể xuất ${qty}.`);
        return;
      }
    }

    const movement = { id: uid(), type, docNo: ioForm.docNo, date: ioForm.date, qty, price, series };
    setProducts((prev) => prev.map((p) => (p.id === product.id ? { ...p, movements: [...p.movements, movement] } : p)));
    setIoModal(null);
  };

  const viewingProduct = products.find((p) => p.id === viewingId) || null;
  const openIOFromDetail = (product, type) => { setViewingId(null); openIO(product, type); };
  const openEditFromDetail = (product) => { setViewingId(null); openEdit(product); };

  return (
      <div>
        <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
          <div className="relative flex-1 max-w-xs">
            <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã VT hoặc tên…"
                   className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
          </div>
          <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
            <Plus size={15} /> Thêm sản phẩm
          </button>
        </div>

        <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <table className="w-full text-sm">
            <thead>
            <tr style={{ borderBottom: `2px solid ${INK}` }}>
              {["", "", "Mã VT", "SKU", "Tên vật tư", "ĐVT", "Tồn đầu kỳ", "Nhập trong kỳ", "Xuất trong kỳ", "Tồn cuối kỳ", "Giá nhập", ""].map((h, hi) => (
                  <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {filtered.map((p) => {
              const stats = productStats(p);
              return (
                  <tr key={p.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                    <td className="px-2 py-3">
                      <button onClick={() => setViewingId(p.id)} className="opacity-50 hover:opacity-100" title="Xem chi tiết"><ChevronRight size={15} /></button>
                    </td>
                    <td className="px-2 py-3">
                      {p.image ? (
                          <img src={p.image} alt={p.name} className="w-9 h-9 object-cover rounded-sm" style={{ border: `1px solid ${LINE}` }} />
                      ) : (
                          <div className="w-9 h-9 rounded-sm flex items-center justify-center" style={{ background: PAPER, border: `1px dashed ${LINE}` }}>
                            <ImageOff size={13} className="opacity-30" />
                          </div>
                      )}
                    </td>
                    <td className="px-3 py-3 font-medium">
                      <button onClick={() => setViewingId(p.id)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{p.code}</button>
                    </td>
                    <td className="px-3 py-3 opacity-70 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.sku || "—"}</td>
                    <td className="px-3 py-3" style={{ color: INK }}>
                      <button onClick={() => setViewingId(p.id)} className="text-left hover:underline">{p.name}</button>
                      <div className="flex gap-1.5 mt-1">
                        {p.hasSeries && <span className="inline-flex items-center gap-1 text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BLUE}1A`, color: BLUE }}><Barcode size={10} /> Series</span>}
                        {p.vat && <span className="inline-block text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BRASS}1A`, color: BRASS }}>{VAT_OPTIONS.find((v) => v.id === p.vat)?.label || p.vat}</span>}
                      </div>
                    </td>
                    <td className="px-3 py-3 opacity-70 whitespace-nowrap">{p.unit}</td>
                    <td className="px-3 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.openingQty}</td>
                    <td className="px-3 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>+{stats.importedQty}</td>
                    <td className="px-3 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>-{stats.exportedQty}</td>
                    <td className="px-3 py-3 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: stats.closingQty <= 5 ? RUST : INK }}>{stats.closingQty}</td>
                    <td className="px-3 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(stats.avgCost)}</td>
                    <td className="px-3 py-3">
                      <div className="flex gap-1.5 justify-end whitespace-nowrap">
                        <button onClick={() => openIO(p, "in")} title="Nhập kho" className="p-1.5 rounded-sm hover:bg-black/5" style={{ color: FOREST }}><ArrowDownToLine size={14} /></button>
                        <button onClick={() => openIO(p, "out")} title="Xuất kho" className="p-1.5 rounded-sm hover:bg-black/5" style={{ color: RUST }}><ArrowUpFromLine size={14} /></button>
                        <button onClick={() => openEdit(p)} title="Sửa" className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                        <button onClick={() => removeProduct(p.id)} title="Xoá" className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
              );
            })}
            {filtered.length === 0 && <tr><td colSpan={12} className="text-center py-8 opacity-50">Không có sản phẩm nào.</td></tr>}
            </tbody>
          </table>
        </div>

        {/* Modal chi tiết sản phẩm — thông tin đầy đủ + danh sách series + lịch sử nhập/xuất */}
        {viewingProduct && (() => {
          const stats = productStats(viewingProduct);
          const rows = viewingProduct.hasSeries ? seriesList(viewingProduct) : [];
          return (
              <Modal title="Chi tiết sản phẩm" onClose={() => setViewingId(null)} size="2xl">
                <div className="flex items-start gap-4 mb-5">
                  {viewingProduct.image ? (
                      <img src={viewingProduct.image} alt={viewingProduct.name} className="w-24 h-24 object-cover rounded-sm shrink-0" style={{ border: `1px solid ${LINE}` }} />
                  ) : (
                      <div className="w-24 h-24 rounded-sm flex items-center justify-center shrink-0" style={{ background: PAPER, border: `1px dashed ${LINE}` }}>
                        <ImageOff size={22} className="opacity-30" />
                      </div>
                  )}
                  <div className="min-w-0">
                    <h4 style={{ fontFamily: "'Fraunces', serif", color: INK }} className="text-lg leading-snug">{viewingProduct.name}</h4>
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: PAPER, color: INK, fontFamily: "'IBM Plex Mono', monospace", border: `1px solid ${LINE}` }}>{viewingProduct.code}</span>
                      <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: PAPER, color: INK, fontFamily: "'IBM Plex Mono', monospace", border: `1px solid ${LINE}` }}>SKU {viewingProduct.sku || "—"}</span>
                      {viewingProduct.hasSeries && <span className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${BLUE}1A`, color: BLUE }}><Barcode size={11} /> Series</span>}
                      <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${BRASS}1A`, color: BRASS }}>{VAT_OPTIONS.find((v) => v.id === viewingProduct.vat)?.label || viewingProduct.vat}</span>
                    </div>
                    <p className="text-xs opacity-50 mt-2">{viewingProduct.category || "Chưa phân nhóm"} · {viewingProduct.unit}{viewingProduct.barcode ? ` · Barcode: ${viewingProduct.barcode}` : ""}</p>
                  </div>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2 mb-5">
                  {[
                    ["Tồn đầu kỳ", viewingProduct.openingQty, INK],
                    ["Nhập trong kỳ", `+${stats.importedQty}`, FOREST],
                    ["Xuất trong kỳ", `-${stats.exportedQty}`, RUST],
                    ["Tồn cuối kỳ", stats.closingQty, stats.closingQty <= 5 ? RUST : INK],
                    ["Giá nhập", vnd(viewingProduct.costPrice), INK],
                    ["Giá bán lẻ", vnd(viewingProduct.retailPrice), INK],
                  ].map(([label, val, color], i) => (
                      <div key={i} className="p-2.5 rounded-sm text-center" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                        <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">{label}</p>
                        <p className="text-sm font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color }}>{val}</p>
                      </div>
                  ))}
                </div>

                <div className="flex gap-2 mb-5 flex-wrap">
                  <button onClick={() => openIOFromDetail(viewingProduct, "in")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm text-white" style={{ background: FOREST }}><ArrowDownToLine size={14} /> Nhập kho</button>
                  <button onClick={() => openIOFromDetail(viewingProduct, "out")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm text-white" style={{ background: RUST }}><ArrowUpFromLine size={14} /> Xuất kho</button>
                  <button onClick={() => openEditFromDetail(viewingProduct)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}><Pencil size={14} /> Sửa thông tin</button>
                </div>

                <div className="mb-5">
                  <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Lịch sử Nhập - Xuất - Trả (chi tiết theo chứng từ)</p>
                  <div className="rounded-sm overflow-hidden" style={{ border:`1px solid ${LINE}` }}>
                    <table className="w-full text-xs">
                      <thead style={{ background: PAPER }}><tr className="opacity-60">
                        <th className="text-left py-2 px-2">Ngày</th><th className="text-left py-2 px-2">Loại</th><th className="text-left py-2 px-2">Chứng từ</th><th className="text-right py-2 px-2">SL</th><th className="text-right py-2 px-2">Đơn giá</th><th className="text-right py-2 px-2">Tồn sau</th><th className="text-left py-2 px-2">Series</th>
                      </tr></thead>
                      <tbody>
                      {(() => {
                        const moves = [...viewingProduct.movements].sort((a,b)=> new Date(a.date)-new Date(b.date));
                        let running = viewingProduct.openingQty;
                        return moves.map((m,i)=>{
                          running += m.type==="in"? m.qty : -m.qty;
                          const isReturn = m.docNo.includes("-RT");
                          return (
                              <tr key={m.id} style={{ borderTop:`1px dashed ${LINE}`, background: isReturn? `${BRASS}0D` : "transparent" }}>
                                <td className="py-1.5 px-2">{m.date}</td>
                                <td className="py-1.5 px-2">
                              <span className="px-1.5 py-0.5 rounded-full text-[10px]" style={{ background: m.type==="in"? `${FOREST}15`:`${RUST}15`, color: m.type==="in"?FOREST:RUST }}>
                                {m.type==="in" ? (isReturn? "Trả hàng nhập" : "Nhập kho") : (isReturn? "Xuất đổi trả" : "Xuất kho")}
                              </span>
                                </td>
                                <td className="py-1.5 px-2" style={{ fontFamily:"'IBM Plex Mono', monospace", color:BLUE }}>{m.docNo}</td>
                                <td className="py-1.5 px-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace", color: m.type==="in"?FOREST:RUST }}>{m.type==="in"?`+${m.qty}`:`-${m.qty}`}</td>
                                <td className="py-1.5 px-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(m.price)}</td>
                                <td className="py-1.5 px-2 text-right font-medium" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{running}</td>
                                <td className="py-1.5 px-2" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{m.series?.length? m.series.join(", "): "—"}</td>
                              </tr>
                          );
                        });
                      })()}
                      {viewingProduct.movements.length===0 && <tr><td colSpan={7} className="text-center py-6 opacity-40">Chưa có phát sinh nhập/xuất/trả.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>

                {viewingProduct.hasSeries && (
                    <div>
                      <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Danh sách series — {rows.filter((r) => r.status === "Còn tồn").length} còn tồn / {rows.length} tổng</p>
                      {rows.length === 0 ? <p className="text-sm opacity-50">Chưa có series nào — nhập kho để thêm.</p> : (
                          <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
                            <table className="w-full text-xs">
                              <thead style={{ background: PAPER }}><tr className="opacity-60">
                                <th className="text-left py-2 px-2">Số Series</th><th className="text-left py-2 px-2">Phiếu nhập</th><th className="text-left py-2 px-2">Ngày nhập</th>
                                <th className="text-left py-2 px-2">Phiếu xuất</th><th className="text-left py-2 px-2">Ngày xuất</th><th className="text-left py-2 px-2">Trạng thái</th><th className="text-left py-2 px-2">Có thể bán</th>
                              </tr></thead>
                              <tbody>
                              {rows.map((r) => (
                                  <tr key={r.serial} style={{ borderTop: `1px dashed ${LINE}` }}>
                                    <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{r.serial}</td>
                                    <td className="py-1.5 px-2">{r.importDoc}</td>
                                    <td className="py-1.5 px-2">{r.importDate}</td>
                                    <td className="py-1.5 px-2">{r.exportDoc || "—"}</td>
                                    <td className="py-1.5 px-2">{r.exportDate || "—"}</td>
                                    <td className="py-1.5 px-2"><span style={{ color: r.status === "Còn tồn" ? FOREST : RUST }}>{r.status}</span></td>
                                    <td className="py-1.5 px-2">{r.status === "Còn tồn" ? <span style={{ color: FOREST }}>Có</span> : <span className="opacity-40">Không</span>}</td>
                                  </tr>
                              ))}
                              </tbody>
                            </table>
                          </div>
                      )}
                    </div>
                )}
              </Modal>
          );
        })()}
        {editing !== null && (
            <Modal title={editing.id ? "Sửa thông tin sản phẩm" : "Thêm sản phẩm"} onClose={() => setEditing(null)}>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Mã VT"><input className={inputCls} style={{ borderColor: LINE }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} disabled={!!editing.id} /></Field>
                <Field label="Đơn vị tính">
                  <select className={inputCls} style={{ borderColor: LINE }} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })}>
                    {UNITS.map((u) => <option key={u} value={u}>{u}</option>)}
                  </select>
                </Field>
              </div>
              <Field label="Tên vật tư"><input className={inputCls} style={{ borderColor: LINE }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
              <Field label="Nhóm hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} /></Field>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Giá nhập (đ)"><input type="number" className={inputCls} style={{ borderColor: LINE }} value={form.costPrice} onChange={(e) => setForm({ ...form, costPrice: e.target.value })} /></Field>
                <Field label="Giá bán lẻ (đ)"><input type="number" className={inputCls} style={{ borderColor: LINE }} value={form.retailPrice} onChange={(e) => setForm({ ...form, retailPrice: e.target.value })} /></Field>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Giá bán sỉ (đ)"><input type="number" className={inputCls} style={{ borderColor: LINE }} value={form.wholesalePrice} onChange={(e) => setForm({ ...form, wholesalePrice: e.target.value })} /></Field>
                <Field label="Tồn đầu kỳ (SL)"><input type="number" className={inputCls} style={{ borderColor: LINE }} value={form.openingQty} onChange={(e) => setForm({ ...form, openingQty: e.target.value })} /></Field>
              </div>
              <label className="flex items-center gap-2 mb-4 mt-1 text-sm" style={{ color: INK }}>
                <input type="checkbox" checked={!!form.hasSeries} onChange={(e) => setForm({ ...form, hasSeries: e.target.checked })} />
                Quản lý theo số series
              </label>

              <div className="my-4" style={{ borderTop: `1px dashed ${LINE}` }} />

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Mã SKU" hint="Tự động sinh — có thể sửa lại">
                  <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} />
                </Field>
                <Field label="Barcode" hint="Bỏ trống nếu không có">
                  <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.barcode} onChange={(e) => setForm({ ...form, barcode: e.target.value })} placeholder="893xxxxxxxxxx" />
                </Field>
              </div>

              <Field label="Thuế giá trị gia tăng (VAT)">
                <div className="flex gap-2 flex-wrap">
                  {VAT_OPTIONS.map((v) => (
                      <button key={v.id} type="button" onClick={() => setForm({ ...form, vat: v.id })}
                              className="px-3 py-1.5 rounded-sm text-sm border"
                              style={{ borderColor: form.vat === v.id ? INK : LINE, background: form.vat === v.id ? INK : "transparent", color: form.vat === v.id ? "#fff" : INK }}>
                        {v.label}
                      </button>
                  ))}
                </div>
              </Field>

              <Field label="Hình ảnh sản phẩm (ảnh chính)" hint="Không bắt buộc — tối đa 1.5MB">
                <div className="flex items-center gap-3">
                  {form.image ? (
                      <div className="relative">
                        <img src={form.image} alt="preview" className="w-16 h-16 object-cover rounded-sm" style={{ border: `1px solid ${LINE}` }} />
                        <button type="button" onClick={() => setForm({ ...form, image: null })}
                                className="absolute -top-2 -right-2 w-5 h-5 rounded-full flex items-center justify-center text-white" style={{ background: RUST }}>
                          <X size={11} />
                        </button>
                      </div>
                  ) : (
                      <label className="w-16 h-16 rounded-sm flex flex-col items-center justify-center cursor-pointer gap-0.5"
                             style={{ background: PAPER, border: `1px dashed ${LINE}`, color: INK }}>
                        <ImagePlus size={16} className="opacity-50" />
                        <input type="file" accept="image/*" className="hidden" onChange={onImagePick} />
                      </label>
                  )}
                  <span className="text-xs opacity-50">{form.image ? "Bấm ✕ để xoá / đổi ảnh khác" : "Bấm để tải ảnh lên"}</span>
                </div>
              </Field>

              <button onClick={submitInfo} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>
                {editing.id ? "Lưu thay đổi" : "Thêm sản phẩm"}
              </button>
            </Modal>
        )}

        {/* Modal nhập kho / xuất kho */}
        {ioModal && (
            <Modal title={`${ioModal.type === "in" ? "Nhập kho" : "Xuất kho"} — ${ioModal.product.code}`} onClose={() => setIoModal(null)} wide>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label={ioModal.type === "in" ? "Số phiếu nhập" : "Số phiếu xuất"}>
                  <input className={inputCls} style={{ borderColor: LINE }} value={ioForm.docNo} onChange={(e) => setIoForm({ ...ioForm, docNo: e.target.value })} />
                </Field>
                <Field label="Ngày">
                  <input type="date" className={inputCls} style={{ borderColor: LINE }} value={ioForm.date} onChange={(e) => setIoForm({ ...ioForm, date: e.target.value })} />
                </Field>
              </div>
              {ioModal.type === "out" && (
                  <Field label="Áp dụng mức giá">
                    <div className="flex gap-2">
                      <button type="button" onClick={() => setIoForm({ ...ioForm, priceLevel: "retail", price: ioModal.product.retailPrice })}
                              className="px-3 py-1.5 rounded-sm text-sm border" style={{ borderColor: ioForm.priceLevel === "retail" ? INK : LINE, background: ioForm.priceLevel === "retail" ? INK : "transparent", color: ioForm.priceLevel === "retail" ? "#fff" : INK }}>
                        Giá lẻ · {vnd(ioModal.product.retailPrice)}
                      </button>
                      <button type="button" onClick={() => setIoForm({ ...ioForm, priceLevel: "wholesale", price: ioModal.product.wholesalePrice })}
                              className="px-3 py-1.5 rounded-sm text-sm border" style={{ borderColor: ioForm.priceLevel === "wholesale" ? INK : LINE, background: ioForm.priceLevel === "wholesale" ? INK : "transparent", color: ioForm.priceLevel === "wholesale" ? "#fff" : INK }}>
                        Giá sỉ · {vnd(ioModal.product.wholesalePrice)}
                      </button>
                    </div>
                  </Field>
              )}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Số lượng">
                  <input type="number" className={inputCls} style={{ borderColor: LINE }} value={ioForm.qty} onChange={(e) => setIoForm({ ...ioForm, qty: e.target.value })} />
                </Field>
                <Field label={ioModal.type === "in" ? "Đơn giá nhập (đ)" : "Đơn giá xuất (đ)"}>
                  <input type="number" className={inputCls} style={{ borderColor: LINE }} value={ioForm.price} onChange={(e) => setIoForm({ ...ioForm, price: e.target.value })} />
                </Field>
              </div>

              {ioModal.product.hasSeries && ioModal.type === "in" && (
                  <Field label="Số series (gõ rồi nhấn dấu cách hoặc Enter để tạo thẻ)" hint={`Đã nhập ${ioForm.series.length} / cần ${Number(ioForm.qty) || 0}`}>
                    <SeriesTagInput series={ioForm.series} setSeries={(arr) => setIoForm({ ...ioForm, series: arr })} placeholder="VD: 03305377170620" />
                  </Field>
              )}

              {ioModal.product.hasSeries && ioModal.type === "out" && (
                  <Field label="Chọn series xuất (bắt buộc, phải khớp số lượng)" hint={`Đã chọn ${ioForm.selectedSeries.length} / cần ${Number(ioForm.qty) || 0} — còn tồn ${availableSeries.length}`}>
                    <SeriesPicker available={availableSeries} selected={ioForm.selectedSeries} setSelected={(arr) => setIoForm({ ...ioForm, selectedSeries: arr })} need={Number(ioForm.qty) || 0} />
                  </Field>
              )}

              <button onClick={submitIO} disabled={ioModal.product.hasSeries && (ioModal.type === "in" ? ioForm.series.length !== (Number(ioForm.qty) || 0) : ioForm.selectedSeries.length !== (Number(ioForm.qty) || 0))}
                      className="w-full py-2.5 rounded-sm text-white text-sm mt-2 disabled:opacity-40" style={{ background: ioModal.type === "in" ? FOREST : RUST }}>
                {ioModal.type === "in" ? "Xác nhận nhập kho" : "Xác nhận xuất kho"}
              </button>
            </Modal>
        )}
      </div>
  );
}

/* ---------------- Nhập hàng (Purchase Orders / Đơn nhập hàng) ---------------- */

function SeriesTagInput({ series, setSeries, placeholder }) {
  const [input, setInput] = useState("");
  const commitTokens = (raw) => {
    const tokens = parseSeries(raw);
    if (tokens.length) setSeries([...series, ...tokens]);
  };
  const handleChange = (e) => {
    const val = e.target.value;
    if (/[\s,]$/.test(val)) { commitTokens(val); setInput(""); }
    else setInput(val);
  };
  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      if (input.trim()) { commitTokens(input); setInput(""); }
    } else if (e.key === "Backspace" && input === "" && series.length > 0) {
      setSeries(series.slice(0, -1));
    }
  };
  const handleBlur = () => { if (input.trim()) { commitTokens(input); setInput(""); } };
  const removeAt = (idx) => setSeries(series.filter((_, i) => i !== idx));

  return (
      <div className="w-full border rounded-sm p-2 flex flex-wrap gap-1.5 items-center" style={{ borderColor: LINE, background: "#fff" }}>
        {series.map((s, i) => (
            <span key={i} className="inline-flex items-center gap-1 pl-2.5 pr-1.5 py-1 rounded-full text-xs" style={{ background: `${BLUE}17`, color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>
          {s}
              <button type="button" onClick={() => removeAt(i)} className="hover:opacity-60 rounded-full" style={{ padding: 2 }}><X size={11} /></button>
        </span>
        ))}
        <input
            value={input} onChange={handleChange} onKeyDown={handleKeyDown} onBlur={handleBlur}
            placeholder={series.length === 0 ? (placeholder || "Nhập số series rồi cách khoảng trắng…") : ""}
            className="flex-1 min-w-[140px] outline-none text-sm bg-transparent py-0.5"
        />
      </div>
  );
}

function ProductPicker({ products, onPick }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const q = query.trim().toLowerCase();
  const matches = q
      ? products.filter((p) => p.name.toLowerCase().includes(q) || (p.code || "").toLowerCase().includes(q) || (p.sku || "").toLowerCase().includes(q)).slice(0, 60)
      : products.slice(0, 60);

  return (
      <div className="relative">
        <div className="relative">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 opacity-40" />
          <input
              value={query}
              onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
              onFocus={() => setOpen(true)}
              onBlur={() => setTimeout(() => setOpen(false), 150)}
              placeholder="Gõ tên hoặc mã sản phẩm để tìm…"
              className="w-full pl-8 pr-2 py-2 text-sm rounded-sm border outline-none"
              style={{ borderColor: LINE, background: "#fff" }}
          />
        </div>
        {open && (
            <div className="absolute z-20 mt-1 w-full max-h-60 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
              {matches.length === 0 ? (
                  <p className="text-sm opacity-50 p-3">Không tìm thấy sản phẩm.</p>
              ) : matches.map((p) => (
                  <button key={p.id} onMouseDown={() => { onPick(p.id); setQuery(""); setOpen(false); }}
                          className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex items-center justify-between gap-3"
                          style={{ borderBottom: `1px dashed ${LINE}` }}>
                    <span style={{ color: INK }}>{p.name}</span>
                    <span className="opacity-50 text-xs whitespace-nowrap shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.code}{p.hasSeries ? " · series" : ""}</span>
                  </button>
              ))}
            </div>
        )}
      </div>
  );
}

function ItemsTable({ items, products, onUpdate, onRemove, lockQtyPrice }) {
  const colCount = onRemove ? 7 : 6;
  return (
      <div className="rounded-sm overflow-x-auto mb-3" style={{ border: `1px solid ${LINE}` }}>
        <table className="w-full text-sm table-fixed" style={{ minWidth: 620 }}>
          <colgroup>
            <col style={{ width: 28 }} />
            <col style={{ width: 90 }} />
            <col />
            <col style={{ width: 76 }} />
            <col style={{ width: 110 }} />
            <col style={{ width: 62 }} />
            <col style={{ width: 110 }} />
            {onRemove && <col style={{ width: 32 }} />}
          </colgroup>
          <thead>
          <tr style={{ borderBottom: `2px solid ${INK}`, background: PAPER }}>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-50">#</th>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">SKU</th>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Sản phẩm</th>
            <th className="text-center px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">SL</th>
            <th className="text-right px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Đơn giá</th>
            <th className="text-center px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-40">VAT</th>
            <th className="text-right px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Thành tiền</th>
            {onRemove && <th></th>}
          </tr>
          </thead>
          <tbody>
          {items.map((it, idx) => {
            const p = products.find((x) => x.id === it.productId);
            return (
                <React.Fragment key={it.productId}>
                  <tr style={{ borderBottom: p?.hasSeries ? "none" : `1px dashed ${LINE}` }}>
                    <td className="px-2 py-3 text-xs opacity-40 align-top">{idx + 1}</td>
                    <td className="px-2 py-3 align-top truncate" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK, opacity: 0.7 }}>{p?.sku || p?.code}</td>
                    <td className="px-2 py-3 align-top">
                      <div className="font-semibold text-base leading-snug" style={{ color: INK }}>{p?.name}</div>
                      {p?.hasSeries && <span className="mt-1 inline-flex items-center gap-1 text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BLUE}1A`, color: BLUE }}><Barcode size={10} /> Series</span>}
                    </td>
                    <td className="px-1 py-3 align-top">
                      {lockQtyPrice ? (
                          <span className="block text-center text-[15px] font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{it.qty}</span>
                      ) : (
                          <input type="number" min={1} value={it.qty} onChange={(e) => onUpdate(it.productId, { qty: Math.max(1, Number(e.target.value)) })}
                                 className="w-full border rounded-sm py-2 px-1 text-center text-[15px] font-medium" style={{ borderColor: LINE }} />
                      )}
                    </td>
                    <td className="px-1 py-3 align-top">
                      {lockQtyPrice ? (
                          <span className="block text-right text-[15px] font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(it.price)}</span>
                      ) : (
                          <input type="number" value={it.price} onChange={(e) => onUpdate(it.productId, { price: Number(e.target.value) })}
                                 className="w-full border rounded-sm py-2 px-2 text-right text-[15px] font-medium" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                      )}
                    </td>
                    <td className="px-1 py-3 align-top">
                      <select value={it.vat} onChange={(e) => onUpdate(it.productId, { vat: e.target.value })}
                              className="text-[11px] border rounded-sm py-1 px-1 opacity-70 w-full" style={{ borderColor: LINE }}>
                        {VAT_OPTIONS.map((v) => <option key={v.id} value={v.id}>{v.label}</option>)}
                      </select>
                    </td>
                    <td className="px-2 py-3 text-right font-bold text-base whitespace-nowrap align-top" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(it.qty * it.price)}</td>
                    {onRemove && <td className="px-1 py-3 align-top"><button onClick={() => onRemove(it.productId)} style={{ color: RUST }}><X size={14} /></button></td>}
                  </tr>
                  {p?.hasSeries && (
                      <tr style={{ borderBottom: `1px dashed ${LINE}` }}>
                        <td></td>
                        <td colSpan={colCount - 1} className="px-2 pb-3">
                          <span className="block opacity-60 mb-1 text-xs">Số series — cần {it.qty}</span>
                          <SeriesTagInput series={it.series} setSeries={(arr) => onUpdate(it.productId, { series: arr })} placeholder="VD: 03305377170620" />
                        </td>
                      </tr>
                  )}
                </React.Fragment>
            );
          })}
          {items.length === 0 && <tr><td colSpan={colCount} className="text-center py-6 opacity-40 text-sm">Chưa có sản phẩm nào.</td></tr>}
          </tbody>
        </table>
      </div>
  );
}

function POProgressStepper({ po }) {
  const steps = [
    { label: "Đặt hàng", done: true, at: po.createdAt },
    { label: "Đã giao", done: po.status === "received", at: po.receivedAt },
    { label: "Đã thanh toán", done: po.paid, at: po.paidAt },
    { label: po.paid ? "Hoàn thành" : "Đang giao dịch", done: po.paid, at: po.paidAt },
  ];
  return (
      <div className="flex items-start mb-5">
        {steps.map((s, i) => (
            <React.Fragment key={i}>
              <div className="flex flex-col items-center text-center shrink-0" style={{ width: 96 }}>
                <div className="w-7 h-7 rounded-full flex items-center justify-center mb-1.5" style={{ background: s.done ? FOREST : "#fff", border: `2px solid ${s.done ? FOREST : LINE}` }}>
                  {s.done ? <Check size={14} color="#fff" /> : <span className="text-[11px] opacity-40">{i + 1}</span>}
                </div>
                <span className="text-[11px] font-medium leading-tight" style={{ color: INK, opacity: s.done ? 1 : 0.45 }}>{s.label}</span>
                {s.done && s.at && <span className="text-[10px] opacity-45 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(s.at)}</span>}
              </div>
              {i < steps.length - 1 && <div className="flex-1 h-0.5 mt-3.5" style={{ background: steps[i + 1].done ? FOREST : LINE }} />}
            </React.Fragment>
        ))}
      </div>
  );
}

function PurchaseOrders({ purchaseOrders, setPurchaseOrders, products, setProducts, suppliers }) {
  const [creating, setCreating] = useState(false);
  const [form, setForm] = useState({});
  const [expanded, setExpanded] = useState(null);
  const [viewingId, setViewingId] = useState(null); // id của đơn đang xem/sửa
  const [editForm, setEditForm] = useState(null);

  const viewingPO = purchaseOrders.find((x) => x.id === viewingId) || null;

  const openNew = () => { setForm({ supplier: "", branch: BRANCHES[0], createdBy: EMPLOYEES[0], paymentMethod: "credit", invoiceNo: "", notes: "", tags: [], items: [] }); setCreating(true); };

  const addItem = (productId) => {
    if (!productId) return;
    setForm((f) => {
      if (f.items.some((it) => it.productId === productId)) return f;
      const p = products.find((x) => x.id === productId);
      return { ...f, items: [...f.items, { productId, qty: 1, price: p.costPrice || 0, vat: p.vat || "VAT10", series: [] }] };
    });
  };
  const updateItem = (productId, patch) => setForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
  const removeItem = (productId) => setForm((f) => ({ ...f, items: f.items.filter((it) => it.productId !== productId) }));
  const total = form.items?.reduce((s, it) => s + it.qty * it.price, 0) || 0;

  const validateItems = (itemsRaw) => {
    if (!itemsRaw || itemsRaw.length === 0) { alert("Chưa có sản phẩm nào trong đơn."); return null; }
    const built = [];
    for (const it of itemsRaw) {
      const p = products.find((x) => x.id === it.productId);
      if (p.hasSeries && it.series.length !== it.qty) {
        alert(`"${p.name}" quản lý theo series — cần đúng ${it.qty} số series (đang có ${it.series.length}). Vui lòng điền đủ số series trước khi nhập hàng.`);
        return null;
      }
      built.push({ productId: it.productId, qty: it.qty, price: it.price, vat: it.vat, series: p.hasSeries ? it.series : [] });
    }
    return built;
  };
  const itemsInvalid = (itemsRaw) => {
    if (!itemsRaw || itemsRaw.length === 0) return true;
    return itemsRaw.some((it) => {
      const p = products.find((x) => x.id === it.productId);
      return p?.hasSeries && it.series.length !== it.qty;
    });
  };

  const submit = (status) => {
    const items = validateItems(form.items);
    if (!items) return;
    const now = new Date().toISOString();
    const paid = form.paymentMethod === "cash";
    const po = {
      id: uid(), code: nextPOCode(purchaseOrders), createdAt: now,
      status, receivedAt: status === "received" ? now : null,
      branch: form.branch, supplier: form.supplier, createdBy: form.createdBy,
      invoiceNo: form.invoiceNo || "", notes: form.notes || "", tags: form.tags || [],
      paymentMethod: form.paymentMethod, paid, paidAt: paid ? now : null, items,
    };
    setPurchaseOrders((prev) => [po, ...prev]);
    if (status === "received") applyPOToStock(po, setProducts);
    setCreating(false);
  };

  const confirmReceive = (po) => {
    const receivedAt = new Date().toISOString();
    setPurchaseOrders((prev) => prev.map((x) => (x.id === po.id ? { ...x, status: "received", receivedAt } : x)));
    applyPOToStock(po, setProducts);
  };
  const markPaid = (po) => {
    const paidAt = new Date().toISOString();
    setPurchaseOrders((prev) => prev.map((x) => (x.id === po.id ? { ...x, paid: true, paidAt } : x)));
  };

  // ----- Xem / sửa đơn đã tạo -----
  const openView = (po) => {
    setEditForm({
      supplier: po.supplier, branch: po.branch, createdBy: po.createdBy, paymentMethod: po.paymentMethod,
      invoiceNo: po.invoiceNo || "", notes: po.notes || "", tags: [...(po.tags || [])],
      items: po.items.map((it) => ({ ...it, series: [...it.series] })),
    });
    setViewingId(po.id);
  };
  const closeView = () => { setViewingId(null); setEditForm(null); };
  const editAddItem = (productId) => {
    if (!productId) return;
    setEditForm((f) => {
      if (f.items.some((it) => it.productId === productId)) return f;
      const p = products.find((x) => x.id === productId);
      return { ...f, items: [...f.items, { productId, qty: 1, price: p.costPrice || 0, vat: p.vat || "VAT10", series: [] }] };
    });
  };
  const editUpdateItem = (productId, patch) => setEditForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
  const editRemoveItem = (productId) => setEditForm((f) => ({ ...f, items: f.items.filter((it) => it.productId !== productId) }));
  const editTotal = editForm?.items?.reduce((s, it) => s + it.qty * it.price, 0) || 0;

  // Đơn đã "Đã nhập": SL/đơn giá bị khoá (đã cộng vào tồn kho), nhưng vẫn cho sửa số series & VAT của từng dòng,
  // đồng thời đồng bộ số series mới vào đúng bút toán nhập kho đã tạo cho đơn này.
  const saveReceivedEdits = () => {
    for (const it of editForm.items) {
      const p = products.find((x) => x.id === it.productId);
      if (p?.hasSeries && it.series.length !== it.qty) {
        alert(`"${p.name}" cần đúng ${it.qty} số series (đang có ${it.series.length}).`);
        return;
      }
    }
    const items = editForm.items;
    setPurchaseOrders((prev) => prev.map((x) => {
      if (x.id !== viewingId) return x;
      return { ...x, supplier: editForm.supplier, branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, items };
    }));
    setProducts((prev) => prev.map((p) => {
      const it = items.find((i) => i.productId === p.id);
      if (!it) return p;
      return { ...p, movements: p.movements.map((m) => (m.type === "in" && m.docNo === viewingPO.code ? { ...m, series: it.series } : m)) };
    }));
    closeView();
  };
  const savePendingItems = () => {
    const items = validateItems(editForm.items);
    if (!items) return;
    const paid = editForm.paymentMethod === "cash";
    setPurchaseOrders((prev) => prev.map((x) => (x.id === viewingId ? { ...x, supplier: editForm.supplier, branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, paid, paidAt: paid ? (x.paidAt || new Date().toISOString()) : null, items } : x)));
    closeView();
  };
  const confirmReceiveFromEdit = () => {
    const items = validateItems(editForm.items);
    if (!items) return;
    const receivedAt = new Date().toISOString();
    const paid = editForm.paymentMethod === "cash";
    const updated = { ...viewingPO, supplier: editForm.supplier, branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, paid, paidAt: paid ? (viewingPO.paidAt || receivedAt) : null, items, status: "received", receivedAt };
    setPurchaseOrders((prev) => prev.map((x) => (x.id === viewingId ? updated : x)));
    applyPOToStock(updated, setProducts);
    closeView();
  };

  return (
      <div>
        <datalist id="ncc-list">
          {suppliers.map((s) => <option key={s.id} value={`${s.code} - ${s.name}`} />)}
        </datalist>
        <div className="flex items-center justify-between mb-5">
          <p className="text-sm opacity-60">{purchaseOrders.length} đơn nhập hàng</p>
          <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
            <Plus size={15} /> Tạo đơn nhập hàng
          </button>
        </div>

        <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <table className="w-full text-sm">
            <thead>
            <tr style={{ borderBottom: `2px solid ${INK}` }}>
              {["", "Mã đơn nhập", "Ngày nhập", "Trạng thái nhập", "Chi nhánh nhập", "Nhà cung cấp", "Nhân viên tạo", "Thanh toán", "Giá trị đơn", ""].map((h, hi) => (
                  <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {purchaseOrders.map((po) => {
              const s = PO_STATUSES.find((x) => x.id === po.status);
              const total = po.items.reduce((sum, it) => sum + it.qty * it.price, 0);
              const isOpen = expanded === po.id;
              return (
                  <React.Fragment key={po.id}>
                    <tr style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                      <td className="px-2 py-3">
                        <button onClick={() => setExpanded(isOpen ? null : po.id)} className="opacity-60 hover:opacity-100">
                          {isOpen ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                        </button>
                      </td>
                      <td className="px-3 py-3 font-medium">
                        <button onClick={() => openView(po)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{po.code}</button>
                      </td>
                      <td className="px-3 py-3 whitespace-nowrap opacity-80" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(po.createdAt)}</td>
                      <td className="px-3 py-3">
                        <span style={{ borderColor: s.color, color: s.color, fontFamily: "'IBM Plex Mono', monospace" }} className="inline-block border rounded-full px-2.5 py-0.5 text-[11px]">{s.label}</span>
                      </td>
                      <td className="px-3 py-3 opacity-70 whitespace-nowrap">{po.branch}</td>
                      <td className="px-3 py-3 opacity-70 whitespace-nowrap">{po.supplier || "—"}</td>
                      <td className="px-3 py-3 opacity-70 whitespace-nowrap">{po.createdBy || "—"}</td>
                      <td className="px-3 py-3">
                      <span className="inline-block rounded-full px-2.5 py-0.5 text-[11px] whitespace-nowrap"
                            style={{ background: po.paid ? `${FOREST}1A` : `${BRASS}1A`, color: po.paid ? FOREST : BRASS }}>
                        {po.paid ? "Hoàn thành" : "Đang giao dịch"}
                      </span>
                      </td>
                      <td className="px-3 py-3 text-right font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(total)}</td>
                      <td className="px-3 py-3">
                        {po.status === "pending" && (
                            <button onClick={() => confirmReceive(po)} className="text-[11px] px-2.5 py-1.5 rounded-sm text-white whitespace-nowrap" style={{ background: FOREST }}>
                              Xác nhận nhập hàng
                            </button>
                        )}
                      </td>
                    </tr>
                    {isOpen && (
                        <tr>
                          <td colSpan={10} className="px-6 py-3" style={{ background: PAPER, borderBottom: `1px dashed ${LINE}` }}>
                            <table className="w-full text-xs">
                              <thead><tr className="opacity-60">
                                <th className="text-left py-1 pr-2">#</th><th className="text-left py-1">Sản phẩm</th><th className="text-left py-1">SL</th><th className="text-left py-1">Đơn giá</th><th className="text-left py-1">VAT</th><th className="text-left py-1">Thành tiền</th><th className="text-left py-1">Series</th>
                              </tr></thead>
                              <tbody>
                              {po.items.map((it, i) => {
                                const p = products.find((x) => x.id === it.productId);
                                return (
                                    <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                      <td className="py-1.5 pr-2 opacity-50">{i + 1}</td>
                                      <td className="py-1.5">{p?.name || "?"}</td>
                                      <td className="py-1.5">{it.qty}</td>
                                      <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td>
                                      <td className="py-1.5">{VAT_OPTIONS.find((v) => v.id === it.vat)?.label || it.vat}</td>
                                      <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(it.qty * it.price)}</td>
                                      <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series?.length ? it.series.join(", ") : "—"}</td>
                                    </tr>
                                );
                              })}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                    )}
                  </React.Fragment>
              );
            })}
            {purchaseOrders.length === 0 && <tr><td colSpan={10} className="text-center py-8 opacity-50">Chưa có đơn nhập hàng nào.</td></tr>}
            </tbody>
          </table>
        </div>

        {/* Modal tạo đơn mới */}
        {creating && (
            <Modal title="Tạo đơn nhập hàng" onClose={() => setCreating(false)} size="2xl">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <Field label="Nhà cung cấp" hint="Chọn từ danh sách NCC hoặc gõ tự do">
                  <input list="ncc-list" className={inputCls} style={{ borderColor: LINE }} value={form.supplier} onChange={(e) => setForm({ ...form, supplier: e.target.value })} placeholder="Gõ để tìm NCC…" />
                </Field>
                <Field label="Chi nhánh nhập">
                  <select className={inputCls} style={{ borderColor: LINE }} value={form.branch} onChange={(e) => setForm({ ...form, branch: e.target.value })}>
                    {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                  </select>
                </Field>
                <Field label="Nhân viên tạo" hint="Phân quyền theo tài khoản sẽ làm ở giai đoạn sau">
                  <select className={inputCls} style={{ borderColor: LINE }} value={form.createdBy} onChange={(e) => setForm({ ...form, createdBy: e.target.value })}>
                    {EMPLOYEES.map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                  </select>
                </Field>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Số hóa đơn" hint="Hóa đơn từ nhà cung cấp — có thể điền sau">
                  <input className={inputCls} style={{ borderColor: LINE }} value={form.invoiceNo} onChange={(e) => setForm({ ...form, invoiceNo: e.target.value })} placeholder="VD: HD0004521" />
                </Field>
                <Field label="Hình thức thanh toán" hint="Công nợ nhà cung cấp sẽ nối vào đơn này ở giai đoạn sau">
                  <div className="flex gap-2">
                    {PAYMENT_METHODS.map((m) => (
                        <button key={m.id} type="button" onClick={() => setForm({ ...form, paymentMethod: m.id })}
                                className="px-3.5 py-1.5 rounded-sm text-sm border"
                                style={{ borderColor: form.paymentMethod === m.id ? INK : LINE, background: form.paymentMethod === m.id ? INK : "transparent", color: form.paymentMethod === m.id ? "#fff" : INK }}>
                          {m.label}
                        </button>
                    ))}
                  </div>
                </Field>
              </div>

              <div className="my-4" style={{ borderTop: `1px dashed ${LINE}` }} />

              <Field label="Thêm sản phẩm vào đơn">
                <ProductPicker products={products} onPick={addItem} />
              </Field>

              <ItemsTable items={form.items || []} products={products} onUpdate={updateItem} onRemove={removeItem} />

              <div className="flex justify-between items-center py-2 mb-1" style={{ borderTop: `2px solid ${INK}` }}>
                <span className="text-sm uppercase tracking-wider opacity-60">Tổng giá trị đơn</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }} className="text-lg">{vnd(total)}</span>
              </div>

              <TagsNotesCompact tags={form.tags} setTags={(arr) => setForm({ ...form, tags: arr })} notes={form.notes} setNotes={(v) => setForm({ ...form, notes: v })} />

              <div className="flex flex-col sm:flex-row gap-3 mt-4">
                <button onClick={() => submit("pending")} disabled={itemsInvalid(form.items)} className="flex-1 py-2.5 rounded-sm text-sm border disabled:opacity-40" style={{ borderColor: BRASS, color: BRASS }}>Chờ giao</button>
                <button onClick={() => submit("received")} disabled={itemsInvalid(form.items)} className="flex-1 py-2.5 rounded-sm text-white text-sm disabled:opacity-40" style={{ background: FOREST }}>Nhập hàng</button>
              </div>
            </Modal>
        )}

        {/* Modal xem / sửa đơn đã tạo */}
        {viewingPO && editForm && (
            <Modal title={`Đơn nhập hàng ${viewingPO.code}`} onClose={closeView} size="2xl">
              <POProgressStepper po={viewingPO} />

              {!viewingPO.paid && (
                  <button onClick={() => markPaid(viewingPO)} className="mb-5 px-3.5 py-1.5 rounded-sm text-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                    Đánh dấu đã thanh toán cho nhà cung cấp
                  </button>
              )}

              <div className="my-4" style={{ borderTop: `1px dashed ${LINE}` }} />

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <Field label="Nhà cung cấp">
                  <input list="ncc-list" className={inputCls} style={{ borderColor: LINE }} value={editForm.supplier} onChange={(e) => setEditForm({ ...editForm, supplier: e.target.value })} />
                </Field>
                <Field label="Chi nhánh nhập">
                  <select className={inputCls} style={{ borderColor: LINE }} value={editForm.branch} onChange={(e) => setEditForm({ ...editForm, branch: e.target.value })}>
                    {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                  </select>
                </Field>
                <Field label="Nhân viên tạo">
                  <select className={inputCls} style={{ borderColor: LINE }} value={editForm.createdBy} onChange={(e) => setEditForm({ ...editForm, createdBy: e.target.value })}>
                    {EMPLOYEES.map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                  </select>
                </Field>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Số hóa đơn" hint="Hóa đơn từ nhà cung cấp — có thể điền sau">
                  <input className={inputCls} style={{ borderColor: LINE }} value={editForm.invoiceNo} onChange={(e) => setEditForm({ ...editForm, invoiceNo: e.target.value })} placeholder="VD: HD0004521" />
                </Field>
                <Field label="Hình thức thanh toán">
                  <div className="flex gap-2">
                    {PAYMENT_METHODS.map((m) => (
                        <button key={m.id} type="button" onClick={() => setEditForm({ ...editForm, paymentMethod: m.id })}
                                className="px-3.5 py-1.5 rounded-sm text-sm border"
                                style={{ borderColor: editForm.paymentMethod === m.id ? INK : LINE, background: editForm.paymentMethod === m.id ? INK : "transparent", color: editForm.paymentMethod === m.id ? "#fff" : INK }}>
                          {m.label}
                        </button>
                    ))}
                  </div>
                </Field>
              </div>

              <div className="my-4" style={{ borderTop: `1px dashed ${LINE}` }} />

              {viewingPO.status === "pending" ? (
                  <>
                    <Field label="Thêm sản phẩm vào đơn">
                      <ProductPicker products={products} onPick={editAddItem} />
                    </Field>
                    {editForm.items.length > 0 && (
                        <ItemsTable items={editForm.items} products={products} onUpdate={editUpdateItem} onRemove={editRemoveItem} />
                    )}
                    <div className="flex justify-between items-center py-2 mb-1" style={{ borderTop: `2px solid ${INK}` }}>
                      <span className="text-sm uppercase tracking-wider opacity-60">Tổng giá trị đơn</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }} className="text-lg">{vnd(editTotal)}</span>
                    </div>
                    <TagsNotesCompact tags={editForm.tags} setTags={(arr) => setEditForm({ ...editForm, tags: arr })} notes={editForm.notes} setNotes={(v) => setEditForm({ ...editForm, notes: v })} />
                    <div className="flex flex-col sm:flex-row gap-3 mt-4">
                      <button onClick={savePendingItems} disabled={itemsInvalid(editForm.items)} className="flex-1 py-2.5 rounded-sm text-sm border disabled:opacity-40" style={{ borderColor: LINE, color: INK }}>Lưu thay đổi</button>
                      <button onClick={confirmReceiveFromEdit} disabled={itemsInvalid(editForm.items)} className="flex-1 py-2.5 rounded-sm text-white text-sm disabled:opacity-40" style={{ background: FOREST }}>Xác nhận nhập hàng</button>
                    </div>
                  </>
              ) : (
                  <>
                    <p className="text-xs mb-3 px-3 py-2 rounded-sm" style={{ background: `${BLUE}10`, color: INK }}>Đơn đã nhập hàng — số lượng và đơn giá đã cộng vào tồn kho nên khoá lại. Vẫn có thể sửa <b>số series</b> hoặc <b>VAT</b> nếu nhập nhầm.</p>
                    <ItemsTable items={editForm.items} products={products} onUpdate={editUpdateItem} lockQtyPrice />
                    <div className="flex justify-between items-center py-2 mb-1" style={{ borderTop: `2px solid ${INK}` }}>
                      <span className="text-sm uppercase tracking-wider opacity-60">Tổng giá trị đơn</span>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }} className="text-lg">{vnd(editTotal)}</span>
                    </div>
                    <TagsNotesCompact tags={editForm.tags} setTags={(arr) => setEditForm({ ...editForm, tags: arr })} notes={editForm.notes} setNotes={(v) => setEditForm({ ...editForm, notes: v })} />
                    <button onClick={saveReceivedEdits} disabled={itemsInvalid(editForm.items)} className="w-full py-2.5 rounded-sm text-white text-sm disabled:opacity-40 mt-4" style={{ background: INK }}>Lưu thay đổi</button>
                  </>
              )}
            </Modal>
        )}
      </div>
  );
}



/* ---------------- Sản phẩm & Tồn kho — bao gồm 2 menu con ---------------- */

function Suppliers({ suppliers, setSuppliers }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({});

  const filtered = suppliers.filter((s) =>
      s.name.toLowerCase().includes(query.toLowerCase()) || s.code.toLowerCase().includes(query.toLowerCase()) || (s.contactPerson || "").toLowerCase().includes(query.toLowerCase())
  );

  const openNew = () => { setForm({ code: nextSupplierCode(suppliers), name: "", taxCode: "", address: "", contactPerson: "", phone: "", email: "", paymentTerm: "cash", creditDays: 30 }); setEditing({}); };
  const openEdit = (s) => { setForm({ ...s }); setEditing(s); };
  const submit = () => {
    if (!form.name) return;
    if (editing.id) {
      setSuppliers((prev) => prev.map((s) => (s.id === editing.id ? { ...s, ...form, creditDays: Number(form.creditDays) || 0 } : s)));
    } else {
      setSuppliers((prev) => [...prev, { ...form, id: uid(), code: form.code || nextSupplierCode(suppliers), creditDays: Number(form.creditDays) || 0 }]);
    }
    setEditing(null);
  };
  const remove = (id) => setSuppliers((prev) => prev.filter((s) => s.id !== id));

  return (
      <div>
        <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
          <div className="relative flex-1 max-w-xs">
            <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã, tên hoặc người liên hệ…"
                   className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
          </div>
          <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
            <Plus size={15} /> Thêm nhà cung cấp
          </button>
        </div>

        <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <table className="w-full text-sm" style={{ minWidth: 720 }}>
            <thead>
            <tr style={{ borderBottom: `2px solid ${INK}` }}>
              {["Mã NCC", "Tên NCC", "Mã số thuế", "Người liên hệ", "SĐT", "Địa chỉ", "Công nợ cấp", ""].map((h, hi) => (
                  <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {filtered.map((s) => (
                <tr key={s.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                  <td className="px-3 py-3 font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{s.code}</td>
                  <td className="px-3 py-3" style={{ color: INK }}>{s.name}</td>
                  <td className="px-3 py-3 opacity-70 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.taxCode || "—"}</td>
                  <td className="px-3 py-3 opacity-70 whitespace-nowrap">{s.contactPerson || "—"}</td>
                  <td className="px-3 py-3 opacity-70 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.phone || "—"}</td>
                  <td className="px-3 py-3 opacity-70 max-w-[220px] truncate">{s.address || "—"}</td>
                  <td className="px-3 py-3 whitespace-nowrap">
                  <span className="inline-block text-[11px] px-2 py-0.5 rounded-full" style={{ background: s.paymentTerm === "cash" ? `${FOREST}1A` : `${BRASS}1A`, color: s.paymentTerm === "cash" ? FOREST : BRASS }}>
                    {s.paymentTerm === "cash" ? "TM (Tiền mặt)" : `Công nợ ${s.creditDays} ngày`}
                  </span>
                  </td>
                  <td className="px-3 py-3">
                    <div className="flex gap-1.5 justify-end whitespace-nowrap">
                      <button onClick={() => openEdit(s)} title="Sửa" className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                      <button onClick={() => remove(s.id)} title="Xoá" className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={8} className="text-center py-8 opacity-50">Chưa có nhà cung cấp nào.</td></tr>}
            </tbody>
          </table>
        </div>

        {editing !== null && (
            <Modal title={editing.id ? "Sửa nhà cung cấp" : "Thêm nhà cung cấp"} onClose={() => setEditing(null)}>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Mã NCC" hint="Tự động sinh — có thể sửa lại">
                  <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} />
                </Field>
                <Field label="Mã số thuế">
                  <input className={inputCls} style={{ borderColor: LINE }} value={form.taxCode} onChange={(e) => setForm({ ...form, taxCode: e.target.value })} />
                </Field>
              </div>
              <Field label="Tên nhà cung cấp">
                <input className={inputCls} style={{ borderColor: LINE }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
              </Field>
              <Field label="Địa chỉ">
                <input className={inputCls} style={{ borderColor: LINE }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
              </Field>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Người liên hệ">
                  <input className={inputCls} style={{ borderColor: LINE }} value={form.contactPerson} onChange={(e) => setForm({ ...form, contactPerson: e.target.value })} />
                </Field>
                <Field label="Số điện thoại">
                  <input className={inputCls} style={{ borderColor: LINE }} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} />
                </Field>
              </div>
              <Field label="Email">
                <input className={inputCls} style={{ borderColor: LINE }} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} />
              </Field>

              <Field label="Công nợ cấp">
                <div className="flex gap-2 mb-2">
                  {SUPPLIER_PAYMENT_TERMS.map((t) => (
                      <button key={t.id} type="button" onClick={() => setForm({ ...form, paymentTerm: t.id })}
                              className="px-3.5 py-1.5 rounded-sm text-sm border"
                              style={{ borderColor: form.paymentTerm === t.id ? INK : LINE, background: form.paymentTerm === t.id ? INK : "transparent", color: form.paymentTerm === t.id ? "#fff" : INK }}>
                        {t.label}
                      </button>
                  ))}
                </div>
                {form.paymentTerm === "credit" && (
                    <label className="text-xs block max-w-[180px]">
                      <span className="block opacity-60 mb-1">Số ngày công nợ</span>
                      <input type="number" min={1} value={form.creditDays} onChange={(e) => setForm({ ...form, creditDays: e.target.value })} className="w-full border rounded-sm py-1.5 px-2" style={{ borderColor: LINE }} />
                    </label>
                )}
              </Field>

              <button onClick={submit} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>
                {editing.id ? "Lưu thay đổi" : "Thêm nhà cung cấp"}
              </button>
            </Modal>
        )}
      </div>
  );
}

function ProductsSection({ products, setProducts, purchaseOrders, setPurchaseOrders, suppliers, setSuppliers }) {
  const [sub, setSub] = useState("list");
  return (
      <div>
        <div className="flex gap-2 mb-5 flex-wrap">
          <button onClick={() => setSub("list")} className="px-3.5 py-1.5 rounded-full text-sm border"
                  style={{ borderColor: sub === "list" ? INK : LINE, background: sub === "list" ? INK : "transparent", color: sub === "list" ? "#fff" : INK }}>
            Danh sách sản phẩm
          </button>
          <button onClick={() => setSub("purchase")} className="px-3.5 py-1.5 rounded-full text-sm border"
                  style={{ borderColor: sub === "purchase" ? INK : LINE, background: sub === "purchase" ? INK : "transparent", color: sub === "purchase" ? "#fff" : INK }}>
            Nhập hàng
          </button>
          <button onClick={() => setSub("suppliers")} className="px-3.5 py-1.5 rounded-full text-sm border"
                  style={{ borderColor: sub === "suppliers" ? INK : LINE, background: sub === "suppliers" ? INK : "transparent", color: sub === "suppliers" ? "#fff" : INK }}>
            Nhà cung cấp
          </button>
        </div>
        {sub === "list" && <ProductsInventory products={products} setProducts={setProducts} />}
        {sub === "purchase" && <PurchaseOrders purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} products={products} setProducts={setProducts} suppliers={suppliers} />}
        {sub === "suppliers" && <Suppliers suppliers={suppliers} setSuppliers={setSuppliers} />}
      </div>
  );
}

/* ---------------- Customers ---------------- */

function ProvinceSelect({ value, onChange }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const q = query.trim().toLowerCase();
  const matches = q ? VN_PROVINCES.filter((p) => p.toLowerCase().includes(q)) : VN_PROVINCES;
  return (
      <div className="relative">
        <button type="button" onClick={() => setOpen((o) => !o)} className="w-full text-left border-b-2 outline-none py-1.5 px-1 text-[15px]" style={{ borderColor: LINE, color: value ? INK : "#999" }}>
          {value || "Chọn Tỉnh/Thành phố"}
        </button>
        {open && (
            <div className="absolute z-20 mt-1 w-full max-h-64 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
              <div className="p-2 sticky top-0" style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
                <div className="relative">
                  <Search size={13} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-40" />
                  <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm kiếm tỉnh/thành…"
                         className="w-full pl-7 pr-2 py-1.5 text-sm rounded-sm border outline-none" style={{ borderColor: LINE }} />
                </div>
              </div>
              {matches.length === 0 ? <p className="text-sm opacity-50 p-3">Không tìm thấy.</p> : matches.map((p) => (
                  <button key={p} onClick={() => { onChange(p); setQuery(""); setOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5" style={{ borderBottom: `1px dashed ${LINE}`, color: INK }}>{p}</button>
              ))}
            </div>
        )}
      </div>
  );
}

function Customers({ customers, setCustomers, orders }) {
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({});
  const [viewingId, setViewingId] = useState(null);
  const [query, setQuery] = useState("");

  const openNew = () => { setForm({ code: nextCustomerCode(customers), name: "", phone: "", email: "", taxCode: "", province: "", ward: "", addressDetail: "", note: "", group: "retail" }); setEditing({}); };
  const openEdit = (c) => { setForm({ ...c }); setEditing(c); };
  const submit = () => {
    if (!form.name) return;
    if (editing.id) setCustomers((prev) => prev.map((c) => (c.id === editing.id ? { ...c, ...form, code: form.code || c.code } : c)));
    else setCustomers((prev) => [...prev, { ...form, id: uid(), code: form.code || nextCustomerCode(customers) }]);
    setEditing(null);
  };
  const remove = (id) => setCustomers((prev) => prev.filter((c) => c.id !== id));
  const fullAddress = (c) => [c.addressDetail, c.ward, c.province].filter(Boolean).join(", ");

  const customerStats = (id) => {
    const custOrders = orders.filter((o) => o.customerId === id);
    const active = custOrders.filter((o) => o.status !== "cancelled");
    const totalSpent = active.reduce((s, o) => s + orderCalc(o).payable, 0);
    const debt = active.reduce((s, o) => s + Math.max(0, orderCalc(o).remaining), 0);
    return { orderCount: custOrders.length, totalSpent, debt, custOrders };
  };

  const filtered = customers.filter((c) =>
      c.name.toLowerCase().includes(query.toLowerCase()) || c.code.toLowerCase().includes(query.toLowerCase()) || (c.phone || "").includes(query)
  );

  const viewingCustomer = customers.find((c) => c.id === viewingId) || null;

  return (
      <div>
        <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
          <div className="relative flex-1 max-w-xs">
            <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã, tên, SĐT…"
                   className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
          </div>
          <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}><Plus size={15} /> Thêm khách hàng</button>
        </div>

        <div className="rounded-sm overflow-x-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <table className="w-full text-sm" style={{ minWidth: 760 }}>
            <thead>
            <tr style={{ borderBottom: `2px solid ${INK}` }}>
              {["Mã KH", "Tên khách hàng", "SĐT", "Nhóm khách hàng", "Công nợ", "Tổng chi tiêu", "Tổng SL đơn", ""].map((h, hi) => (
                  <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {filtered.map((c) => {
              const s = customerStats(c.id);
              const g = CUSTOMER_GROUPS.find((x) => x.id === c.group) || CUSTOMER_GROUPS[0];
              return (
                  <tr key={c.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                    <td className="px-3 py-3 font-medium whitespace-nowrap">
                      <button onClick={() => setViewingId(c.id)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{c.code}</button>
                    </td>
                    <td className="px-3 py-3 whitespace-nowrap" style={{ color: INK }}>{c.name}</td>
                    <td className="px-3 py-3 whitespace-nowrap opacity-70" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{c.phone || "—"}</td>
                    <td className="px-3 py-3 whitespace-nowrap">
                      <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: `${BLUE}15`, color: BLUE }}>{g.label}</span>
                    </td>
                    <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: s.debt > 0 ? RUST : "inherit", opacity: s.debt > 0 ? 1 : 0.4 }}>{vnd(s.debt)}</td>
                    <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(s.totalSpent)}</td>
                    <td className="px-3 py-3 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.orderCount}</td>
                    <td className="px-3 py-3">
                      <div className="flex gap-1.5 justify-end whitespace-nowrap">
                        <button onClick={() => openEdit(c)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                        <button onClick={() => remove(c.id)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
              );
            })}
            {filtered.length === 0 && <tr><td colSpan={8} className="text-center py-10 opacity-50">Không có khách hàng nào.</td></tr>}
            </tbody>
          </table>
        </div>

        {/* Modal chi tiết khách hàng */}
        {viewingCustomer && (() => {
          const s = customerStats(viewingCustomer.id);
          const g = CUSTOMER_GROUPS.find((x) => x.id === viewingCustomer.group) || CUSTOMER_GROUPS[0];
          return (
              <Modal title={viewingCustomer.name} onClose={() => setViewingId(null)} size="xl">
                <div className="flex items-center gap-2 mb-4 flex-wrap">
                  <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${BLUE}15`, color: BLUE }}>{g.label}</span>
                  <span className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{viewingCustomer.code}</span>
                  <button onClick={() => { setViewingId(null); openEdit(viewingCustomer); }} className="ml-auto text-xs px-3 py-1.5 rounded-sm border flex items-center gap-1" style={{ borderColor: LINE, color: INK }}><Pencil size={12} /> Sửa thông tin</button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1.5 text-sm mb-4 p-3 rounded-sm" style={{ background: PAPER }}>
                  <p><span className="opacity-50">SĐT: </span>{viewingCustomer.phone || "—"}</p>
                  <p><span className="opacity-50">Email: </span>{viewingCustomer.email || "—"}</p>
                  <p><span className="opacity-50">MST: </span>{viewingCustomer.taxCode || "—"}</p>
                  <p><span className="opacity-50">Địa chỉ: </span>{fullAddress(viewingCustomer) || "—"}</p>
                  {viewingCustomer.note && <p className="sm:col-span-2"><span className="opacity-50">Ghi chú: </span>{viewingCustomer.note}</p>}
                </div>

                <div className="grid grid-cols-3 gap-2 mb-5">
                  <div className="p-3 rounded-sm text-center" style={{ border: `1px solid ${LINE}` }}>
                    <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Tổng chi tiêu</p>
                    <p className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(s.totalSpent)}</p>
                  </div>
                  <div className="p-3 rounded-sm text-center" style={{ border: `1px solid ${LINE}` }}>
                    <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Tổng SL đơn</p>
                    <p className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{s.orderCount}</p>
                  </div>
                  <div className="p-3 rounded-sm text-center" style={{ border: `1px solid ${LINE}` }}>
                    <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Công nợ hiện tại</p>
                    <p className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: s.debt > 0 ? RUST : FOREST }}>{vnd(s.debt)}</p>
                  </div>
                </div>

                <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Lịch sử mua hàng</p>
                <div className="rounded-sm overflow-x-auto min-w-0" style={{ border: `1px solid ${LINE}` }}>
                  <table className="w-full text-xs" style={{ minWidth: 480 }}>
                    <thead style={{ background: PAPER }}><tr className="opacity-60">
                      <th className="text-left py-2 px-2">Mã đơn</th><th className="text-left py-2 px-2">Trạng thái</th><th className="text-right py-2 px-2">Giá trị</th><th className="text-right py-2 px-2">Còn phải trả</th><th className="text-left py-2 px-2">Ngày</th>
                    </tr></thead>
                    <tbody>
                    {s.custOrders.map((o) => {
                      const oc = orderCalc(o);
                      return (
                          <tr key={o.id} style={{ borderTop: `1px dashed ${LINE}` }}>
                            <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{o.code}</td>
                            <td className="py-1.5 px-2"><Stamp status={o.status} /></td>
                            <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(oc.payable)}</td>
                            <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: oc.remaining > 0 ? RUST : "inherit" }}>{vnd(oc.remaining)}</td>
                            <td className="py-1.5 px-2 whitespace-nowrap">{formatDateTime(o.createdAt)}</td>
                          </tr>
                      );
                    })}
                    {s.custOrders.length === 0 && <tr><td colSpan={5} className="text-center py-6 opacity-40">Chưa có đơn hàng nào.</td></tr>}
                    </tbody>
                  </table>
                </div>
              </Modal>
          );
        })()}

        {editing !== null && (
            <Modal title={editing.id ? "Sửa khách hàng" : "Thêm khách hàng"} onClose={() => setEditing(null)} size="lg">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Tên khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
                <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Mã khách hàng" hint="Tự động sinh — có thể sửa lại">
                  <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} />
                </Field>
                <Field label="Mã số thuế"><input className={inputCls} style={{ borderColor: LINE }} value={form.taxCode} onChange={(e) => setForm({ ...form, taxCode: e.target.value })} /></Field>
              </div>
              <Field label="Email"><input className={inputCls} style={{ borderColor: LINE }} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
              <Field label="Nhóm khách hàng">
                <div className="flex gap-2 flex-wrap">
                  {CUSTOMER_GROUPS.map((g) => (
                      <button key={g.id} type="button" onClick={() => setForm({ ...form, group: g.id })}
                              className="px-3.5 py-1.5 rounded-sm text-sm border"
                              style={{ borderColor: form.group === g.id ? INK : LINE, background: form.group === g.id ? INK : "transparent", color: form.group === g.id ? "#fff" : INK }}>
                        {g.label}
                      </button>
                  ))}
                </div>
              </Field>

              <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
              <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Địa chỉ</p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Field label="Tỉnh/Thành phố">
                  <ProvinceSelect value={form.province} onChange={(v) => setForm({ ...form, province: v })} />
                </Field>
                <Field label="Phường/Xã" hint="Nhập tên phường/xã theo địa danh mới">
                  <input className={inputCls} style={{ borderColor: LINE }} value={form.ward} onChange={(e) => setForm({ ...form, ward: e.target.value })} placeholder="VD: Xã Kim Anh" />
                </Field>
              </div>
              <Field label="Địa chỉ cụ thể"><input className={inputCls} style={{ borderColor: LINE }} value={form.addressDetail} onChange={(e) => setForm({ ...form, addressDetail: e.target.value })} placeholder="Số nhà, tên đường, thôn/khu vực…" /></Field>

              <Field label="Ghi chú"><input className={inputCls} style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
              <button onClick={submit} className="mt-2 w-full py-2.5 rounded-sm text-white text-sm" style={{ background: INK }}>{editing.id ? "Lưu thay đổi" : "Thêm khách hàng"}</button>
            </Modal>
        )}
      </div>
  );
}

/* ---------------- Orders (bán hàng — sẽ đổi thành Phiếu xuất ở giai đoạn sau) ---------------- */

function CustomerPicker({ customers, setCustomers, onPick, placeholder }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [creatingNew, setCreatingNew] = useState(false);
  const [newForm, setNewForm] = useState({ name: "", phone: "" });
  const q = query.trim().toLowerCase();
  const matches = q
      ? customers.filter((c) => c.name.toLowerCase().includes(q) || (c.phone || "").toLowerCase().includes(q)).slice(0, 40)
      : customers.slice(0, 40);

  const startCreate = () => { setCreatingNew(true); setNewForm({ name: query, phone: "" }); };
  const cancelNew = () => setCreatingNew(false);
  const saveNew = () => {
    if (!newForm.name.trim()) return;
    const nc = {
      id: uid(), code: nextCustomerCode(customers), name: newForm.name.trim(), phone: newForm.phone.trim(),
      note: "", email: "", taxCode: "", province: "", ward: "", addressDetail: "", group: "retail",
    };
    setCustomers((prev) => [...prev, nc]);
    onPick(nc.id);
    setCreatingNew(false); setQuery(""); setOpen(false);
  };

  return (
      <div className="relative">
        <div className="relative">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 opacity-40" />
          <input
              value={query} onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
              onFocus={() => setOpen(true)} onBlur={() => setTimeout(() => { if (!creatingNew) setOpen(false); }, 150)}
              placeholder={placeholder || "Tìm theo tên, SĐT khách hàng…"}
              className="w-full pl-8 pr-2 py-2.5 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }}
          />
        </div>
        {(open || creatingNew) && (
            <div className="absolute z-20 mt-1 w-full max-h-72 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
              {creatingNew ? (
                  <div className="p-3">
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Tạo khách hàng mới</p>
                    <input autoFocus value={newForm.name} onChange={(e) => setNewForm({ ...newForm, name: e.target.value })} placeholder="Tên khách hàng"
                           className="w-full border rounded-sm py-2 px-2.5 text-sm mb-2" style={{ borderColor: LINE }} />
                    <input value={newForm.phone} onChange={(e) => setNewForm({ ...newForm, phone: e.target.value })} placeholder="Số điện thoại"
                           className="w-full border rounded-sm py-2 px-2.5 text-sm mb-3" style={{ borderColor: LINE }} />
                    <div className="flex gap-2">
                      <button onMouseDown={cancelNew} className="flex-1 py-1.5 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}>Huỷ</button>
                      <button onMouseDown={saveNew} className="flex-1 py-1.5 rounded-sm text-sm text-white" style={{ background: INK }}>Lưu &amp; chọn</button>
                    </div>
                  </div>
              ) : (
                  <>
                    <button onMouseDown={startCreate} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex items-center gap-1.5 font-medium" style={{ borderBottom: `1px dashed ${LINE}`, color: BLUE }}>
                      <Plus size={14} /> Tạo khách hàng mới
                    </button>
                    <button onMouseDown={() => { onPick(null); setQuery(""); setOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 opacity-70" style={{ borderBottom: `1px dashed ${LINE}` }}>
                      Khách lẻ (không chọn khách hàng)
                    </button>
                    {matches.length === 0 ? (
                        <p className="text-sm opacity-50 p-3">Không tìm thấy khách hàng.</p>
                    ) : matches.map((c) => (
                        <button key={c.id} onMouseDown={() => { onPick(c.id); setQuery(""); setOpen(false); }}
                                className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex items-center justify-between gap-3" style={{ borderBottom: `1px dashed ${LINE}` }}>
                          <span style={{ color: INK }}>{c.name}</span>
                          <span className="opacity-50 text-xs whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{c.phone}</span>
                        </button>
                    ))}
                  </>
              )}
            </div>
        )}
      </div>
  );
}

function SeriesPicker({ available, selected, setSelected, need }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const remaining = available.filter((s) => !selected.includes(s.serial));
  const q = query.trim().toLowerCase();
  const matches = (q ? remaining.filter((s) => s.serial.toLowerCase().includes(q)) : remaining).slice(0, 30);
  const pick = (serial) => { setSelected([...selected, serial]); setQuery(""); };
  const removeAt = (serial) => setSelected(selected.filter((s) => s !== serial));

  return (
      <div>
        {selected.length > 0 && (
            <div className="flex flex-wrap gap-1.5 mb-1.5">
              {selected.map((s) => (
                  <span key={s} className="inline-flex items-center gap-1 pl-2.5 pr-1.5 py-1 rounded-full text-xs" style={{ background: `${BLUE}17`, color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>
              {s}
                    <button type="button" onMouseDown={() => removeAt(s)} className="hover:opacity-60 rounded-full" style={{ padding: 2 }}><X size={11} /></button>
            </span>
              ))}
            </div>
        )}
        <div className="relative">
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 opacity-40" />
            <input
                value={query} onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
                onFocus={() => setOpen(true)} onBlur={() => setTimeout(() => setOpen(false), 150)}
                placeholder={selected.length >= need ? "Đã đủ series" : "Gõ để tìm số series còn tồn…"}
                className="w-full pl-8 pr-2 py-2 text-xs rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }}
            />
          </div>
          {open && (
              <div className="absolute z-20 mt-1 w-full max-h-40 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                {matches.length === 0 ? (
                    <p className="text-xs opacity-50 p-2">Không tìm thấy series phù hợp.</p>
                ) : matches.map((s) => (
                    <button key={s.serial} onMouseDown={() => pick(s.serial)} className="w-full text-left px-3 py-1.5 text-xs hover:bg-black/5"
                            style={{ fontFamily: "'IBM Plex Mono', monospace", borderBottom: `1px dashed ${LINE}` }}>{s.serial}</button>
                ))}
              </div>
          )}
        </div>
      </div>
  );
}

function SalesItemsTable({ items, products, onUpdate, onRemove }) {
  return (
      <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
        <table className="w-full text-sm table-fixed" style={{ minWidth: 520 }}>
          <colgroup>
            <col style={{ width: 28 }} />
            <col style={{ width: 46 }} />
            <col />
            <col style={{ width: 76 }} />
            <col style={{ width: 120 }} />
            <col style={{ width: 120 }} />
            <col style={{ width: 32 }} />
          </colgroup>
          <thead>
          <tr style={{ borderBottom: `2px solid ${INK}`, background: PAPER }}>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-50">STT</th>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Ảnh</th>
            <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Sản phẩm</th>
            <th className="text-center px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">SL</th>
            <th className="text-right px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Đơn giá</th>
            <th className="text-right px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-60">Thành tiền</th>
            <th></th>
          </tr>
          </thead>
          <tbody>
          {items.map((it, idx) => {
            const p = products.find((x) => x.id === it.productId);
            const available = p?.hasSeries ? seriesList(p).filter((s) => s.status === "Còn tồn" || it.series.includes(s.serial)) : [];
            return (
                <React.Fragment key={it.productId}>
                  <tr style={{ borderBottom: p?.hasSeries ? "none" : `1px dashed ${LINE}` }}>
                    <td className="px-2 py-3 text-xs opacity-40 align-top">{idx + 1}</td>
                    <td className="px-2 py-3 align-top">
                      {p?.image ? <img src={p.image} alt={p.name} className="w-9 h-9 object-cover rounded-sm" style={{ border: `1px solid ${LINE}` }} /> : (
                          <div className="w-9 h-9 rounded-sm flex items-center justify-center" style={{ background: PAPER, border: `1px dashed ${LINE}` }}><ImageOff size={12} className="opacity-30" /></div>
                      )}
                    </td>
                    <td className="px-2 py-3 align-top">
                      <div className="font-semibold text-base leading-snug" style={{ color: INK }}>{p?.name}</div>
                      <div className="text-xs opacity-50 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p?.sku || p?.code}</div>
                    </td>
                    <td className="px-1 py-3 align-top">
                      <input type="number" min={1} value={it.qty} onChange={(e) => onUpdate(it.productId, { qty: Math.max(1, Number(e.target.value)) })}
                             className="w-full border rounded-sm py-2 px-1 text-center text-[15px] font-medium" style={{ borderColor: LINE }} />
                    </td>
                    <td className="px-1 py-3 align-top">
                      <input type="number" value={it.price} onChange={(e) => onUpdate(it.productId, { price: Number(e.target.value) })}
                             className="w-full border rounded-sm py-2 px-2 text-right text-[15px] font-medium" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                    </td>
                    <td className="px-2 py-3 text-right font-bold text-base whitespace-nowrap align-top" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(orderLineTotal(it))}</td>
                    <td className="px-1 py-3 align-top"><button onClick={() => onRemove(it.productId)} style={{ color: RUST }}><X size={14} /></button></td>
                  </tr>
                  {p?.hasSeries && (
                      <tr style={{ borderBottom: `1px dashed ${LINE}` }}>
                        <td></td>
                        <td colSpan={6} className="px-2 pb-3">
                          <span className="block opacity-60 mb-1 text-xs">Chọn series xuất bán — cần {it.qty} (còn tồn {available.length})</span>
                          <SeriesPicker available={available} selected={it.series} setSelected={(arr) => onUpdate(it.productId, { series: arr })} need={it.qty} />
                        </td>
                      </tr>
                  )}
                </React.Fragment>
            );
          })}
          {items.length === 0 && <tr><td colSpan={7} className="text-center py-8 opacity-40 text-sm">Chưa có thông tin sản phẩm.</td></tr>}
          </tbody>
        </table>
      </div>
  );
}

function OrderProgressStepper({ order }) {
  if (order.status === "cancelled") {
    return (
        <div className="flex items-center gap-2 mb-5 px-3 py-2.5 rounded-sm" style={{ background: `${RUST}10`, border: `1px solid ${RUST}33` }}>
          <span className="w-6 h-6 rounded-full flex items-center justify-center shrink-0" style={{ background: RUST }}><X size={13} color="#fff" /></span>
          <div>
            <p className="text-sm font-medium" style={{ color: RUST }}>Đã huỷ</p>
            {order.cancelledAt && <p className="text-[11px] opacity-60" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(order.cancelledAt)}</p>}
          </div>
        </div>
    );
  }
  const steps = [
    { label: "Đặt hàng", done: true, at: order.createdAt },
    { label: "Đang giao", done: !!order.shippingAt, at: order.shippingAt },
    { label: "Đã giao", done: !!order.deliveredAt, at: order.deliveredAt },
    { label: "Đã thanh toán", done: !!order.paidCompleteAt, at: order.paidCompleteAt },
    { label: "Hoàn thành", done: !!order.deliveredAt && !!order.paidCompleteAt, at: order.deliveredAt && order.paidCompleteAt ? (order.deliveredAt > order.paidCompleteAt ? order.deliveredAt : order.paidCompleteAt) : null },
  ];
  return (
      <div className="flex items-start mb-5 overflow-x-auto">
        {steps.map((s, i) => (
            <React.Fragment key={i}>
              <div className="flex flex-col items-center text-center shrink-0" style={{ width: 86 }}>
                <div className="w-7 h-7 rounded-full flex items-center justify-center mb-1.5" style={{ background: s.done ? FOREST : "#fff", border: `2px solid ${s.done ? FOREST : LINE}` }}>
                  {s.done ? <Check size={14} color="#fff" /> : <span className="text-[11px] opacity-40">{i + 1}</span>}
                </div>
                <span className="text-[11px] font-medium leading-tight" style={{ color: INK, opacity: s.done ? 1 : 0.45 }}>{s.label}</span>
                {s.done && s.at && <span className="text-[10px] opacity-45 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(s.at)}</span>}
              </div>
              {i < steps.length - 1 && <div className="flex-1 h-0.5 mt-3.5" style={{ background: steps[i + 1].done ? FOREST : LINE, minWidth: 16 }} />}
            </React.Fragment>
        ))}
      </div>
  );
}

function OrderInvoicePrint({ order, products, customer }) {
  if (!order) return null;
  const c = orderCalc(order);
  const vp = vatPercent(order.vat);
  return (
      <div id="invoice-print-area">
        <div style={{ fontFamily: "'Times New Roman', Times, serif", color: "#111", fontSize: 13, padding: "10mm 12mm", maxWidth: "210mm", margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: "2px solid #111", paddingBottom: 10, marginBottom: 14 }}>
            <div style={{ lineHeight: 1.6 }}>
              <div><b>Tên đơn vị: </b>{COMPANY_INFO.name}</div>
              <div><b>Địa chỉ: </b>{COMPANY_INFO.address}</div>
              <div><b>MST: </b>{COMPANY_INFO.taxCode}</div>
              <div><b>Tài khoản số: </b>{COMPANY_INFO.bankAccount}</div>
              <div><b>Liên hệ: </b>{COMPANY_INFO.phone}</div>
            </div>
          </div>

          <h1 style={{ textAlign: "center", fontSize: 22, letterSpacing: 1, margin: "4px 0 10px" }}>HÓA ĐƠN BÁN HÀNG</h1>
          <div style={{ display: "flex", justifyContent: "center", gap: 24, marginBottom: 14 }}>
            <span>Ngày {new Date(order.createdAt).getDate()} tháng {new Date(order.createdAt).getMonth() + 1} năm {new Date(order.createdAt).getFullYear()}</span>
            <span><b>Số:</b> {order.code}</span>
          </div>

          <div style={{ marginBottom: 10 }}><b>Khách hàng: </b>{customer?.name || "Khách lẻ"}</div>
          {customer && (customer.addressDetail || customer.ward || customer.province) && (
              <div style={{ marginBottom: 10 }}><b>Địa chỉ: </b>{[customer.addressDetail, customer.ward, customer.province].filter(Boolean).join(", ")}</div>
          )}
          {customer?.phone && <div style={{ marginBottom: 14 }}><b>Điện thoại: </b>{customer.phone}</div>}

          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 14 }}>
            <thead>
            <tr>
              {["STT", "Tên SP", "Mã", "Đơn vị tính", "Số lượng", "Đơn giá", "Thành tiền", "Ghi chú"].map((h, i) => (
                  <th key={i} style={{ border: "1px solid #111", padding: "6px 8px", fontSize: 12 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {order.items.map((it, i) => {
              const p = products.find((x) => x.id === it.productId);
              return (
                  <tr key={i}>
                    <td style={{ border: "1px solid #111", padding: "6px 8px", textAlign: "center" }}>{i + 1}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px" }}>{p?.name}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px" }}>{p?.code}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px", textAlign: "center" }}>{p?.unit}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px", textAlign: "center" }}>{it.qty}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px", textAlign: "right" }}>{it.price.toLocaleString("vi-VN")}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px", textAlign: "right" }}>{orderLineTotal(it).toLocaleString("vi-VN")}</td>
                    <td style={{ border: "1px solid #111", padding: "6px 8px" }}></td>
                  </tr>
              );
            })}
            </tbody>
          </table>

          <table style={{ width: "45%", marginLeft: "auto", borderCollapse: "collapse", marginBottom: 14 }}>
            <tbody>
            <tr><td style={{ padding: "3px 8px" }}>Cộng:</td><td style={{ padding: "3px 8px", textAlign: "right" }}>{c.subtotal.toLocaleString("vi-VN")}</td></tr>
            <tr><td style={{ padding: "3px 8px" }}>Thuế{vp > 0 ? ` (${vp}%)` : ""}:</td><td style={{ padding: "3px 8px", textAlign: "right" }}>{VAT_OPTIONS.find((v) => v.id === order.vat)?.label === "KCT" ? "KCT" : c.vatTotal.toLocaleString("vi-VN")}</td></tr>
            {order.orderDiscount > 0 && <tr><td style={{ padding: "3px 8px" }}>Chiết khấu:</td><td style={{ padding: "3px 8px", textAlign: "right" }}>{order.orderDiscount.toLocaleString("vi-VN")}</td></tr>}
            {order.shippingFee > 0 && <tr><td style={{ padding: "3px 8px" }}>Phí giao hàng:</td><td style={{ padding: "3px 8px", textAlign: "right" }}>{order.shippingFee.toLocaleString("vi-VN")}</td></tr>}
            <tr style={{ fontWeight: "bold", borderTop: "1px solid #111" }}><td style={{ padding: "5px 8px" }}>Tổng cộng:</td><td style={{ padding: "5px 8px", textAlign: "right" }}>{c.payable.toLocaleString("vi-VN")}</td></tr>
            </tbody>
          </table>

          <p style={{ marginBottom: 18 }}><i>Tổng số tiền bằng chữ: {soTienBangChu(c.payable)}.</i></p>
          <p style={{ fontWeight: "bold", marginBottom: 28 }}>VUI LÒNG KIỂM TRA HÀNG NGAY KHI NHẬN</p>

          <div style={{ display: "flex", justifyContent: "space-between", textAlign: "center" }}>
            <div style={{ width: "30%" }}><b>Người lập phiếu</b><br /><span style={{ fontSize: 11 }}>(Ký, họ tên)</span><br /><br /><br />{order.seller}</div>
            <div style={{ width: "30%" }}><b>Người nhận hàng</b><br /><span style={{ fontSize: 11 }}>(Ký, họ tên)</span></div>
            <div style={{ width: "30%" }}><b>Giám đốc</b><br /><span style={{ fontSize: 11 }}>(Ký, họ tên)</span></div>
          </div>
        </div>
      </div>
  );
}

function Orders({ orders, setOrders, products, setProducts, customers, setCustomers }) {
  const [creating, setCreating] = useState(false);
  const [form, setForm] = useState({});
  const [filterStatus, setFilterStatus] = useState("all");
  const [viewingId, setViewingId] = useState(null);
  const [payInput, setPayInput] = useState("");
  const [printingId, setPrintingId] = useState(null);
  const [returning, setReturning] = useState(false);
  const [returnForm, setReturnForm] = useState(null);
  const [returnViewId, setReturnViewId] = useState(null);
  const [returnEditForm, setReturnEditForm] = useState(null);

  const viewingOrder = orders.find((o) => o.id === viewingId) || null;
  const printingOrder = orders.find((o) => o.id === printingId) || null;

  useEffect(() => {
    if (!printingId) return;
    const t = setTimeout(() => window.print(), 150);
    return () => clearTimeout(t);
  }, [printingId]);
  useEffect(() => {
    const clear = () => setPrintingId(null);
    window.addEventListener("afterprint", clear);
    return () => window.removeEventListener("afterprint", clear);
  }, []);

  const openNew = () => {
    setForm({
      customerId: "", channel: "store", branch: BRANCHES[0], seller: EMPLOYEES[0], deliveryDate: "",
      tags: [], notes: "", items: [], vat: "VAT10", orderDiscount: 0, shippingFee: 0, paidAmount: 0,
    });
    setCreating(true);
  };

  const addItem = (productId) => {
    if (!productId) return;
    setForm((f) => {
      if (f.items.some((it) => it.productId === productId)) return f;
      const p = products.find((x) => x.id === productId);
      return { ...f, items: [...f.items, { productId, qty: 1, price: p.retailPrice, series: [] }] };
    });
  };
  const updateItem = (productId, patch) => setForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
  const removeItem = (productId) => setForm((f) => ({ ...f, items: f.items.filter((it) => it.productId !== productId) }));

  const calc = orderCalc({ items: form.items || [], vat: form.vat, orderDiscount: form.orderDiscount, shippingFee: form.shippingFee, paidAmount: form.paidAmount });

  const itemsInvalid = () => {
    if (!form.items || form.items.length === 0) return true;
    return form.items.some((it) => {
      const p = products.find((x) => x.id === it.productId);
      if (!p) return true;
      if (p.hasSeries && it.series.length !== it.qty) return true;
      const stock = productStats(p).closingQty;
      if (!p.hasSeries && it.qty > stock) return true;
      return false;
    });
  };

  const submit = () => {
    if (itemsInvalid()) { alert("Vui lòng kiểm tra lại số lượng / số series của sản phẩm trong đơn."); return; }
    const code = nextOrderCode(orders);
    const now = new Date().toISOString();
    const paidAmount = Number(form.paidAmount) || 0;
    const payable = calc.payable;
    const newOrder = {
      id: uid(), code, createdAt: now, date: todayISO(),
      customerId: form.customerId, channel: form.channel, branch: form.branch, seller: form.seller, deliveryDate: form.deliveryDate,
      tags: form.tags, notes: form.notes, status: "pending", items: form.items, vat: form.vat,
      shippingAt: null, deliveredAt: null, paidCompleteAt: paidAmount >= payable && payable > 0 ? now : null, cancelledAt: null,
      orderDiscount: Number(form.orderDiscount) || 0, shippingFee: Number(form.shippingFee) || 0, paidAmount,
    };
    setOrders((prev) => [newOrder, ...prev]);
    setProducts((prev) => prev.map((p) => {
      const it = form.items.find((i) => i.productId === p.id);
      if (!it) return p;
      return { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: code, date: todayISO(), qty: it.qty, price: it.price, series: it.series }] };
    }));
    setCreating(false);
  };

  const setStatus = (id, status) => {
    const now = new Date().toISOString();
    setOrders((prev) => prev.map((o) => {
      if (o.id !== id) return o;
      const next = { ...o, status };
      if (status === "shipping" && !o.shippingAt) next.shippingAt = now;
      if (status === "done") { if (!o.shippingAt) next.shippingAt = now; if (!o.deliveredAt) next.deliveredAt = now; }
      if (status === "cancelled" && !o.cancelledAt) next.cancelledAt = now;
      return next;
    }));
  };
  const remove = (id) => setOrders((prev) => prev.filter((o) => o.id !== id));
  const addPayment = () => {
    const amt = Number(payInput);
    if (!amt || amt <= 0) return;
    const now = new Date().toISOString();
    setOrders((prev) => prev.map((o) => {
      if (o.id !== viewingId) return o;
      const newPaid = (o.paidAmount || 0) + amt;
      const c = orderCalc({ ...o, paidAmount: newPaid });
      return { ...o, paidAmount: newPaid, paidCompleteAt: o.paidCompleteAt || (c.remaining <= 0 ? now : null) };
    }));
    setPayInput("");
  };

  // ----- Đổi trả hàng -----
  const openReturn = (order) => {
    setReturnForm({
      type: "refund",
      returnedItems: order.items.map((it) => ({ productId: it.productId, price: it.price, qty: 0, series: [] })),
      exchangeItems: [], note: "",
    });
    setReturning(true);
  };
  const updateReturnedItem = (productId, patch) => setReturnForm((f) => ({ ...f, returnedItems: f.returnedItems.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
  const addExchangeItem = (productId) => {
    if (!productId) return;
    setReturnForm((f) => {
      if (f.exchangeItems.some((it) => it.productId === productId)) return f;
      const p = products.find((x) => x.id === productId);
      return { ...f, exchangeItems: [...f.exchangeItems, { productId, qty: 1, price: p.retailPrice, series: [] }] };
    });
  };
  const updateExchangeItem = (productId, patch) => setReturnForm((f) => ({ ...f, exchangeItems: f.exchangeItems.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
  const removeExchangeItem = (productId) => setReturnForm((f) => ({ ...f, exchangeItems: f.exchangeItems.filter((it) => it.productId !== productId) }));

  const returnInvalid = () => {
    if (!returnForm) return true;
    const active = returnForm.returnedItems.filter((it) => it.qty > 0);
    if (active.length === 0) return true;
    for (const it of active) {
      const p = products.find((x) => x.id === it.productId);
      if (p?.hasSeries && it.series.length !== it.qty) return true;
    }
    if (returnForm.type === "exchange") {
      if (returnForm.exchangeItems.length === 0) return true;
      for (const it of returnForm.exchangeItems) {
        const p = products.find((x) => x.id === it.productId);
        if (!p) return true;
        if (p.hasSeries && it.series.length !== it.qty) return true;
        if (!p.hasSeries && it.qty > productStats(p).closingQty) return true;
      }
    }
    return false;
  };

  const submitReturn = () => {
    if (returnInvalid() || !viewingOrder) return;
    const code = nextReturnCode(viewingOrder);
    const now = new Date().toISOString();
    const returnedItems = returnForm.returnedItems.filter((it) => it.qty > 0);
    const exchangeItems = returnForm.type === "exchange" ? returnForm.exchangeItems : [];
    const rec = { id: uid(), code, createdAt: now, type: returnForm.type, note: returnForm.note, returnedItems, exchangeItems };

    setOrders((prev) => prev.map((o) => (o.id === viewingId ? { ...o, returns: [...(o.returns || []), rec] } : o)));
    setProducts((prev) => prev.map((p) => {
      let movs = [];
      const ret = returnedItems.find((i) => i.productId === p.id);
      if (ret) movs.push({ id: uid(), type: "in", docNo: code, date: todayISO(), qty: ret.qty, price: ret.price, series: ret.series });
      const exc = exchangeItems.find((i) => i.productId === p.id);
      if (exc) movs.push({ id: uid(), type: "out", docNo: code, date: todayISO(), qty: exc.qty, price: exc.price, series: exc.series });
      if (movs.length === 0) return p;
      return { ...p, movements: [...p.movements, ...movs] };
    }));
    setReturning(false);
    setReturnForm(null);
  };

  const visible = filterStatus === "all" ? orders : orders.filter((o) => o.status === filterStatus);

  return (
      <div>
        <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
          <div className="flex gap-2 flex-wrap">
            <FilterChip active={filterStatus === "all"} onClick={() => setFilterStatus("all")}>Tất cả</FilterChip>
            {STATUSES.map((s) => <FilterChip key={s.id} active={filterStatus === s.id} onClick={() => setFilterStatus(s.id)} color={s.color}>{s.label}</FilterChip>)}
          </div>
          <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}><Plus size={15} /> Tạo đơn và giao hàng</button>
        </div>

        <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <table className="w-full text-sm" style={{ minWidth: 760 }}>
            <thead>
            <tr style={{ borderBottom: `2px solid ${INK}` }}>
              {["Mã đơn", "Ngày tạo", "Khách hàng", "Bán tại", "Trạng thái giao", "Tổng tiền", "Còn phải trả", ""].map((h, hi) => (
                  <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
              ))}
            </tr>
            </thead>
            <tbody>
            {visible.map((o) => {
              const cust = customers.find((c) => c.id === o.customerId);
              const c = orderCalc(o);
              return (
                  <tr key={o.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                    <td className="px-3 py-3 font-medium whitespace-nowrap">
                      <button onClick={() => { setViewingId(o.id); setPayInput(""); }} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{o.code}</button>
                    </td>
                    <td className="px-3 py-3 whitespace-nowrap opacity-80" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(o.createdAt)}</td>
                    <td className="px-3 py-3 whitespace-nowrap">{cust?.name || "Khách lẻ"}</td>
                    <td className="px-3 py-3 whitespace-nowrap opacity-70">{o.branch}</td>
                    <td className="px-3 py-3"><Stamp status={o.status} /></td>
                    <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(c.payable)}</td>
                    <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: c.remaining > 0 ? RUST : FOREST }}>{vnd(c.remaining)}</td>
                    <td className="px-3 py-3">
                      <button onClick={() => remove(o.id)} className="opacity-50 hover:opacity-100" style={{ color: RUST }}><Trash2 size={14} /></button>
                    </td>
                  </tr>
              );
            })}
            {visible.length === 0 && <tr><td colSpan={8} className="text-center py-10 opacity-50">Không có đơn hàng nào.</td></tr>}
            </tbody>
          </table>
        </div>

        {/* Modal tạo đơn và giao hàng */}
        {creating && (
            <Modal title="Tạo đơn và giao hàng" onClose={() => setCreating(false)} size="2xl">
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
                <div className="lg:col-span-2 space-y-5 min-w-0">
                  <div>
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Thông tin khách hàng</p>
                    <CustomerPicker customers={customers} setCustomers={setCustomers} onPick={(id) => setForm({ ...form, customerId: id })} />
                    {form.customerId && (() => {
                      const cust = customers.find((c) => c.id === form.customerId);
                      return (
                          <div className="mt-2.5 p-3 rounded-sm flex items-center justify-between" style={{ background: `${FOREST}0F`, border: `1px solid ${FOREST}44` }}>
                            <div>
                              <p className="text-[10px] uppercase tracking-wider opacity-60 mb-0.5">Khách hàng</p>
                              <p className="text-lg font-bold uppercase" style={{ color: INK, letterSpacing: "0.02em" }}>{cust?.name}</p>
                              {cust?.phone && <p className="text-xs opacity-60 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{cust.phone}</p>}
                            </div>
                            <button onClick={() => setForm({ ...form, customerId: "" })} className="text-xs opacity-50 hover:opacity-100 underline shrink-0">Bỏ chọn</button>
                          </div>
                      );
                    })()}
                    {!form.customerId && <p className="text-sm mt-2 opacity-50">Khách lẻ (chưa chọn khách hàng)</p>}
                  </div>

                  <div>
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Thông tin sản phẩm</p>
                    <ProductPicker products={products} onPick={addItem} />
                    <div className="mt-3">
                      <SalesItemsTable items={form.items || []} products={products} onUpdate={updateItem} onRemove={removeItem} />
                    </div>
                  </div>
                </div>

                <div className="space-y-4 min-w-0">
                  <div className="p-4 rounded-sm" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Thông tin bổ sung</p>
                    <Field label="Bán tại">
                      <select className={inputCls} style={{ borderColor: LINE }} value={form.branch} onChange={(e) => setForm({ ...form, branch: e.target.value })}>
                        {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                      </select>
                    </Field>
                    <Field label="Bán bởi">
                      <select className={inputCls} style={{ borderColor: LINE }} value={form.seller} onChange={(e) => setForm({ ...form, seller: e.target.value })}>
                        {EMPLOYEES.map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                      </select>
                    </Field>
                    <Field label="Nguồn">
                      <div className="flex gap-2">
                        {CHANNELS.map((c) => (
                            <button key={c.id} type="button" onClick={() => setForm({ ...form, channel: c.id })}
                                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-sm text-sm border"
                                    style={{ borderColor: form.channel === c.id ? INK : LINE, background: form.channel === c.id ? INK : "transparent", color: form.channel === c.id ? "#fff" : INK }}>
                              <c.icon size={14} /> {c.label}
                            </button>
                        ))}
                      </div>
                    </Field>
                    <Field label="Hẹn giao">
                      <input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.deliveryDate} onChange={(e) => setForm({ ...form, deliveryDate: e.target.value })} />
                    </Field>
                    <Field label="Mã đơn"><input disabled className={inputCls} style={{ borderColor: LINE, opacity: 0.5 }} value={nextOrderCode(orders)} /></Field>
                  </div>

                  <div className="p-4 rounded-sm space-y-2.5" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                    <div className="flex justify-between text-sm"><span className="opacity-60">Tổng tiền ({form.items?.length || 0} sản phẩm)</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(calc.subtotal)}</span></div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="opacity-60">VAT</span>
                      <div className="flex gap-1">
                        {VAT_OPTIONS.map((v) => (
                            <button key={v.id} type="button" onClick={() => setForm({ ...form, vat: v.id })}
                                    className="px-2 py-1 rounded-sm text-[11px] border"
                                    style={{ borderColor: form.vat === v.id ? INK : LINE, background: form.vat === v.id ? INK : "transparent", color: form.vat === v.id ? "#fff" : INK }}>
                              {v.label}
                            </button>
                        ))}
                      </div>
                    </div>
                    <div className="flex justify-between text-sm opacity-60"><span>Trong đó VAT</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.vatTotal)}</span></div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="opacity-60">Chiết khấu</span>
                      <input type="number" value={form.orderDiscount} onChange={(e) => setForm({ ...form, orderDiscount: Number(e.target.value) })} className="w-28 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                    </div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="opacity-60">Phí giao hàng</span>
                      <input type="number" value={form.shippingFee} onChange={(e) => setForm({ ...form, shippingFee: Number(e.target.value) })} className="w-28 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                    </div>
                    <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: INK }}><span>Khách phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.payable)}</span></div>
                    <div className="flex justify-between items-center text-sm">
                      <span className="opacity-60">Khách đã trả</span>
                      <input type="number" value={form.paidAmount} onChange={(e) => setForm({ ...form, paidAmount: Number(e.target.value) })} className="w-28 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                    </div>
                    <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: calc.remaining > 0 ? RUST : FOREST }}><span>Còn phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.remaining)}</span></div>
                  </div>

                  <button onClick={submit} disabled={itemsInvalid()} className="w-full py-3 rounded-sm text-white text-sm font-medium disabled:opacity-40" style={{ background: INK }}>Tạo đơn hàng</button>
                </div>
              </div>

              <TagsNotesCompact tags={form.tags} setTags={(arr) => setForm({ ...form, tags: arr })} notes={form.notes} setNotes={(v) => setForm({ ...form, notes: v })} />
            </Modal>
        )}

        {/* Modal xem chi tiết đơn */}
        {viewingOrder && (() => {
          const cust = customers.find((c) => c.id === viewingOrder.customerId);
          const c = orderCalc(viewingOrder);
          const ch = CHANNELS.find((x) => x.id === viewingOrder.channel);
          return (
              <Modal title={`Đơn hàng ${viewingOrder.code}`} onClose={() => setViewingId(null)} size="2xl">
                <OrderProgressStepper order={viewingOrder} />

                <div className="flex items-center gap-2 mb-4 flex-wrap">
                  <Stamp status={viewingOrder.status} />
                  <span className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(viewingOrder.createdAt)}</span>
                  <span className="text-xs opacity-50 flex items-center gap-1"><ch.icon size={12} /> {ch.label}</span>
                </div>

                <div className="flex flex-wrap gap-1.5 mb-4">
                  {STATUSES.map((s) => (
                      <button key={s.id} onClick={() => setStatus(viewingOrder.id, s.id)} className="text-[11px] px-2.5 py-1.5 rounded-sm"
                              style={{ background: s.id === viewingOrder.status ? `${s.color}22` : PAPER, color: s.color }}>{s.label}</button>
                  ))}
                </div>

                <div className="flex flex-wrap gap-2 mb-4">
                  <button onClick={() => setPrintingId(viewingOrder.id)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                    <Printer size={13} /> In đơn hàng
                  </button>
                  <button onClick={() => openReturn(viewingOrder)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: BRASS, color: BRASS }}>
                    <RotateCcw size={13} /> Đổi trả hàng
                  </button>
                </div>

                <p className="text-[10px] uppercase tracking-wider opacity-50 mb-0.5">Khách hàng</p>
                <p className="text-lg font-bold uppercase mb-1" style={{ color: INK, letterSpacing: "0.02em" }}>{cust?.name || "Khách lẻ"}</p>
                <p className="text-sm mb-1 opacity-70">Bán tại {viewingOrder.branch} · Bán bởi {viewingOrder.seller}{viewingOrder.deliveryDate ? ` · Hẹn giao ${viewingOrder.deliveryDate}` : ""}</p>
                {viewingOrder.notes && <p className="text-sm mb-3 opacity-70">Ghi chú: {viewingOrder.notes}</p>}

                <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />

                <table className="w-full text-xs mb-4">
                  <thead><tr className="opacity-60">
                    <th className="text-left py-1">Sản phẩm</th><th className="text-left py-1">SL</th><th className="text-left py-1">Đơn giá</th><th className="text-left py-1">Thành tiền</th><th className="text-left py-1">Series</th>
                  </tr></thead>
                  <tbody>
                  {viewingOrder.items.map((it, i) => {
                    const p = products.find((x) => x.id === it.productId);
                    const returnedQty = returnedQtyOf(viewingOrder, it.productId);
                    return (
                        <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                          <td className="py-1.5">{p?.name || "?"}{returnedQty > 0 && <span className="ml-1.5 text-[10px]" style={{ color: RUST }}>(đã trả {returnedQty})</span>}</td>
                          <td className="py-1.5">{it.qty}</td>
                          <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td>
                          <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(orderLineTotal(it))}</td>
                          <td className="py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series?.length ? it.series.join(", ") : "—"}</td>
                        </tr>
                    );
                  })}
                  </tbody>
                </table>

                {viewingOrder.returns?.length > 0 && (
                    <div className="mb-5">
                      <p className="text-sm font-semibold uppercase tracking-wider mb-3" style={{ color:INK }}>Lịch sử đổi trả — {viewingOrder.returns.length} phiếu (bấm để xem/sửa)</p>
                      <div className="grid gap-3">
                        {viewingOrder.returns.map((r) => {
                          const retTotal = r.returnedItems.reduce((s,it)=>s+it.qty*it.price,0);
                          const excTotal = r.exchangeItems.reduce((s,it)=>s+it.qty*it.price,0);
                          return (
                              <div key={r.id} onClick={()=> setReturnViewId(r.id)} className="p-3.5 rounded-sm cursor-pointer hover:shadow-md transition-all" style={{ background:"#fff", border:`1px solid ${r.type==="exchange"? BLUE+"55": BRASS+"55"}`, borderLeft:`4px solid ${r.type==="exchange"? BLUE: BRASS}` }}>
                                <div className="flex items-center justify-between mb-2">
                                  <div className="flex items-center gap-2">
                                    <span className="text-sm font-bold" style={{ fontFamily:"'IBM Plex Mono', monospace", color:BLUE }}>{r.code}</span>
                                    <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: r.type==="exchange"? `${BLUE}15`:`${BRASS}15`, color: r.type==="exchange"?BLUE:BRASS }}>{r.type==="exchange"?"Đổi hàng":"Hoàn tiền"}</span>
                                  </div>
                                  <span className="text-xs opacity-60">{formatDateTime(r.createdAt)}</span>
                                </div>
                                <div className="grid grid-cols-2 gap-2 text-xs">
                                  <div><span className="opacity-60">Trả:</span> {r.returnedItems.map(it=>{ const p=products.find(x=>x.id===it.productId); return `${p?.code||""} x${it.qty}`; }).join(", ")} <span style={{ fontFamily:"'IBM Plex Mono', monospace", color:RUST }}> -{vnd(retTotal)}</span></div>
                                  {r.exchangeItems.length>0 && <div><span className="opacity-60">Đổi:</span> {r.exchangeItems.map(it=>{ const p=products.find(x=>x.id===it.productId); return `${p?.code||""} x${it.qty}`; }).join(", ")} <span style={{ fontFamily:"'IBM Plex Mono', monospace", color:FOREST }}> +{vnd(excTotal)}</span></div>}
                                </div>
                                {r.note && <div className="text-xs mt-1.5 opacity-60 italic">Ghi chú: {r.note}</div>}
                                <div className="text-[11px] mt-2 opacity-50">Bấm để xem chi tiết, sửa hoặc xóa phiếu này →</div>
                              </div>
                          );
                        })}
                      </div>
                    </div>
                )}

                <div className="p-4 rounded-sm space-y-2.5 mb-4" style={{ border: `1px solid ${LINE}`, background: PAPER }}>
                  <div className="flex justify-between text-sm"><span className="opacity-60">Tổng tiền</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.subtotal)}</span></div>
                  <div className="flex justify-between text-sm opacity-60"><span>Trong đó VAT ({VAT_OPTIONS.find((v) => v.id === viewingOrder.vat)?.label})</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.vatTotal)}</span></div>
                  <div className="flex justify-between text-sm"><span className="opacity-60">Chiết khấu</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingOrder.orderDiscount)}</span></div>
                  <div className="flex justify-between text-sm"><span className="opacity-60">Phí giao hàng</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingOrder.shippingFee)}</span></div>
                  {c.returnedValue > 0 && <div className="flex justify-between text-sm"><span className="opacity-60">Giá trị hàng trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>-{vnd(c.returnedValue)}</span></div>}
                  {c.exchangeValue > 0 && <div className="flex justify-between text-sm"><span className="opacity-60">Giá trị hàng đổi</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>+{vnd(c.exchangeValue)}</span></div>}
                  <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: INK }}><span>Khách phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.payable)}</span></div>
                  <div className="flex justify-between text-sm"><span className="opacity-60">Khách đã trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingOrder.paidAmount)}</span></div>
                  <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: c.remaining > 0 ? RUST : FOREST }}>
                    <span>{c.remaining >= 0 ? "Còn phải trả" : "Cần hoàn lại cho khách"}</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(Math.abs(c.remaining))}</span>
                  </div>
                </div>

                {c.remaining > 0 && (
                    <div className="flex gap-2">
                      <input type="number" value={payInput} onChange={(e) => setPayInput(e.target.value)} placeholder="Số tiền khách trả thêm"
                             className="flex-1 border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }} />
                      <button onClick={addPayment} className="px-4 py-2 rounded-sm text-white text-sm" style={{ background: FOREST }}>Ghi nhận thanh toán</button>
                    </div>
                )}
              </Modal>
          );
        })()}

        {/* Modal tạo phiếu đổi trả hàng */}
        {returning && returnForm && viewingOrder && (
            <Modal title={`Đổi trả hàng — Đơn ${viewingOrder.code}`} onClose={() => { setReturning(false); setReturnForm(null); }} size="2xl">
              <Field label="Hình thức">
                <div className="flex gap-2">
                  <button type="button" onClick={() => setReturnForm({ ...returnForm, type: "refund" })} className="px-3.5 py-1.5 rounded-sm text-sm border"
                          style={{ borderColor: returnForm.type === "refund" ? INK : LINE, background: returnForm.type === "refund" ? INK : "transparent", color: returnForm.type === "refund" ? "#fff" : INK }}>
                    Hoàn tiền
                  </button>
                  <button type="button" onClick={() => setReturnForm({ ...returnForm, type: "exchange" })} className="px-3.5 py-1.5 rounded-sm text-sm border"
                          style={{ borderColor: returnForm.type === "exchange" ? INK : LINE, background: returnForm.type === "exchange" ? INK : "transparent", color: returnForm.type === "exchange" ? "#fff" : INK }}>
                    Đổi hàng
                  </button>
                </div>
              </Field>

              <p className="text-xs uppercase tracking-wider mb-2 mt-3 opacity-60">Sản phẩm khách trả lại</p>
              <div className="space-y-3 mb-4">
                {viewingOrder.items.map((orig) => {
                  const p = products.find((x) => x.id === orig.productId);
                  const already = returnedQtyOf(viewingOrder, orig.productId);
                  const maxReturnable = orig.qty - already;
                  const it = returnForm.returnedItems.find((x) => x.productId === orig.productId);
                  const allowedSeries = orig.series.filter((s) => !returnedSeriesOf(viewingOrder, orig.productId).includes(s));
                  if (maxReturnable <= 0) return null;
                  return (
                      <div key={orig.productId} className="p-3 rounded-sm" style={{ border: `1px solid ${LINE}` }}>
                        <div className="flex items-center justify-between text-sm mb-2">
                          <span style={{ color: INK }} className="font-medium">{p?.name}</span>
                          <span className="text-xs opacity-50">Đã mua {orig.qty}{already > 0 ? ` · đã trả ${already}` : ""} · tối đa {maxReturnable}</span>
                        </div>
                        <label className="text-xs">
                          <span className="block opacity-60 mb-1">Số lượng trả</span>
                          <input type="number" min={0} max={maxReturnable} value={it.qty}
                                 onChange={(e) => {
                                   const qty = Math.min(maxReturnable, Math.max(0, Number(e.target.value)));
                                   updateReturnedItem(orig.productId, { qty, series: qty === 0 ? [] : it.series.slice(0, qty) });
                                 }}
                                 className="w-20 border rounded-sm py-1.5 px-2 text-center" style={{ borderColor: LINE }} />
                        </label>
                        {p?.hasSeries && it.qty > 0 && (
                            <div className="mt-2">
                              <span className="block opacity-60 mb-1 text-xs">Chọn series trả — cần {it.qty}</span>
                              <div className="flex flex-wrap gap-1.5">
                                {allowedSeries.map((s) => {
                                  const picked = it.series.includes(s);
                                  return (
                                      <button key={s} type="button" onClick={() => {
                                        const newSeries = picked ? it.series.filter((x) => x !== s) : (it.series.length < it.qty ? [...it.series, s] : it.series);
                                        updateReturnedItem(orig.productId, { series: newSeries });
                                      }} className="text-xs px-2 py-1 rounded-full border" style={{ borderColor: picked ? BLUE : LINE, background: picked ? `${BLUE}17` : "transparent", color: picked ? BLUE : INK, fontFamily: "'IBM Plex Mono', monospace" }}>
                                        {s}
                                      </button>
                                  );
                                })}
                              </div>
                            </div>
                        )}
                      </div>
                  );
                })}
              </div>

              {returnForm.type === "exchange" && (
                  <>
                    <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Sản phẩm đổi lấy</p>
                    <ProductPicker products={products} onPick={addExchangeItem} />
                    <div className="mt-3">
                      <SalesItemsTable items={returnForm.exchangeItems} products={products} onUpdate={updateExchangeItem} onRemove={removeExchangeItem} />
                    </div>
                  </>
              )}

              <Field label="Ghi chú" hint="VD: lý do đổi trả">
                <textarea rows={2} className="w-full border rounded-sm p-2 text-sm mt-1" style={{ borderColor: LINE }} value={returnForm.note} onChange={(e) => setReturnForm({ ...returnForm, note: e.target.value })} />
              </Field>

              <button onClick={submitReturn} disabled={returnInvalid()} className="w-full py-2.5 rounded-sm text-white text-sm disabled:opacity-40 mt-2" style={{ background: INK }}>
                {returnForm.type === "exchange" ? "Tạo phiếu đổi hàng" : "Tạo phiếu hoàn tiền"}
              </button>
            </Modal>
        )}

        {returnViewId && viewingOrder && (() => {
          const r = (viewingOrder.returns||[]).find(x=>x.id===returnViewId);
          if(!r) return null;
          const isEditing = !!returnEditForm;
          const formR = isEditing ? returnEditForm : r;
          const openEdit = () => setReturnEditForm({ ...r, returnedItems: r.returnedItems.map(it=>({...it})), exchangeItems: r.exchangeItems.map(it=>({...it})) });
          const saveEdit = () => {
            const retItems = formR.returnedItems.filter(it=>it.qty>0);
            if(retItems.length===0){ alert("Phải có ít nhất 1 sản phẩm trả"); return; }
            for(const it of retItems){ const p=products.find(x=>x.id===it.productId); if(p?.hasSeries && it.series.length!==it.qty){ alert(`Sản phẩm ${p.name} cần đúng ${it.qty} series`); return; } }
            if(formR.type==="exchange"){
              for(const it of formR.exchangeItems){ const p=products.find(x=>x.id===it.productId); if(p?.hasSeries && it.series.length!==it.qty){ alert(`Đổi ${p.name} thiếu series`); return; } }
            }
            // remove old movements of this return code
            setProducts(prev=> prev.map(p=>{
              const filtered = p.movements.filter(m=> m.docNo!==r.code);
              return { ...p, movements: filtered };
            }));
            // then add new
            setTimeout(()=>{
              setProducts(prev=> prev.map(p=>{
                let movs=[];
                const ret = retItems.find(i=>i.productId===p.id);
                if(ret) movs.push({ id: uid(), type:"in", docNo:r.code, date: todayISO(), qty: ret.qty, price: ret.price, series: ret.series });
                const exc = formR.type==="exchange" ? formR.exchangeItems.find(i=>i.productId===p.id) : null;
                if(exc) movs.push({ id: uid(), type:"out", docNo:r.code, date: todayISO(), qty: exc.qty, price: exc.price, series: exc.series });
                if(movs.length===0) return p;
                return { ...p, movements: [...p.movements, ...movs] };
              }));
            },0);
            setOrders(prev=> prev.map(o=>{
              if(o.id!==viewingOrder.id) return o;
              return { ...o, returns: o.returns.map(rr=> rr.id===r.id ? { ...rr, type: formR.type, note: formR.note, returnedItems: retItems, exchangeItems: formR.type==="exchange"?formR.exchangeItems:[] } : rr) };
            }));
            setReturnEditForm(null);
            setReturnViewId(null);
          };
          const deleteReturn = () => {
            if(!confirm(`Xóa phiếu ${r.code}? Tồn kho sẽ được hoàn lại.`)) return;
            setProducts(prev=> prev.map(p=> ({ ...p, movements: p.movements.filter(m=> m.docNo!==r.code) })));
            setOrders(prev=> prev.map(o=> o.id===viewingOrder.id ? { ...o, returns: o.returns.filter(rr=>rr.id!==r.id) } : o));
            setReturnViewId(null);
            setReturnEditForm(null);
          };
          const updateRetItem = (pid, patch) => setReturnEditForm(f=> ({ ...f, returnedItems: f.returnedItems.map(it=> it.productId===pid?{...it,...patch}:it) }));
          const updateExcItem = (pid, patch) => setReturnEditForm(f=> ({ ...f, exchangeItems: f.exchangeItems.map(it=> it.productId===pid?{...it,...patch}:it) }));
          const removeExc = (pid) => setReturnEditForm(f=> ({ ...f, exchangeItems: f.exchangeItems.filter(it=>it.productId!==pid) }));
          const addExc = (pid) => {
            if(!pid) return;
            setReturnEditForm(f=>{
              if(f.exchangeItems.some(it=>it.productId===pid)) return f;
              const p=products.find(x=>x.id===pid);
              return { ...f, exchangeItems:[...f.exchangeItems, { productId:pid, qty:1, price:p.retailPrice, series:[] }] };
            });
          };
          const retTotal = formR.returnedItems.reduce((s,it)=>s+it.qty*it.price,0);
          const excTotal = formR.exchangeItems.reduce((s,it)=>s+it.qty*it.price,0);
          return (
              <Modal title={`Phiếu đổi trả ${r.code}`} onClose={()=>{ setReturnViewId(null); setReturnEditForm(null); }} size="xl">
                {!isEditing ? (
                    <div>
                      <div className="flex items-center gap-2 mb-4">
                        <span className="text-xs px-2.5 py-1 rounded-full" style={{ background: r.type==="exchange"? `${BLUE}15`:`${BRASS}15`, color: r.type==="exchange"?BLUE:BRASS }}>{r.type==="exchange"?"Đổi hàng":"Hoàn tiền"}</span>
                        <span className="text-xs opacity-50" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{formatDateTime(r.createdAt)}</span>
                        <div className="ml-auto flex gap-2">
                          <button onClick={openEdit} className="text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor:LINE, color:INK }}><Pencil size={12} className="inline mr-1" />Sửa</button>
                          <button onClick={deleteReturn} className="text-xs px-3 py-1.5 rounded-sm text-white" style={{ background:RUST }}><Trash2 size={12} className="inline mr-1" />Xóa phiếu</button>
                        </div>
                      </div>
                      {r.note && <div className="mb-3 p-2 rounded-sm text-sm" style={{ background:PAPER, border:`1px dashed ${LINE}` }}>Ghi chú: {r.note}</div>}
                      <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Hàng khách trả</p>
                      <table className="w-full text-xs mb-4" style={{ border:`1px solid ${LINE}` }}>
                        <thead style={{ background:PAPER }}><tr><th className="text-left p-2">SP</th><th className="text-center p-2">SL</th><th className="text-right p-2">Đơn giá</th><th className="text-right p-2">Thành tiền</th><th className="text-left p-2">Series</th></tr></thead>
                        <tbody>{r.returnedItems.map(it=>{ const p=products.find(x=>x.id===it.productId); return <tr key={it.productId} style={{ borderTop:`1px dashed ${LINE}` }}><td className="p-2">{p?.name}</td><td className="p-2 text-center">{it.qty}</td><td className="p-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td><td className="p-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(it.qty*it.price)}</td><td className="p-2" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{it.series?.join(", ")||"—"}</td></tr>; })}</tbody>
                      </table>
                      {r.exchangeItems.length>0 && (<>
                        <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Hàng đổi lấy</p>
                        <table className="w-full text-xs mb-4" style={{ border:`1px solid ${LINE}` }}>
                          <thead style={{ background:PAPER }}><tr><th className="text-left p-2">SP</th><th className="text-center p-2">SL</th><th className="text-right p-2">Đơn giá</th><th className="text-right p-2">Thành tiền</th><th className="text-left p-2">Series</th></tr></thead>
                          <tbody>{r.exchangeItems.map(it=>{ const p=products.find(x=>x.id===it.productId); return <tr key={it.productId} style={{ borderTop:`1px dashed ${LINE}` }}><td className="p-2">{p?.name}</td><td className="p-2 text-center">{it.qty}</td><td className="p-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td><td className="p-2 text-right" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{vnd(it.qty*it.price)}</td><td className="p-2" style={{ fontFamily:"'IBM Plex Mono', monospace" }}>{it.series?.join(", ")||"—"}</td></tr>; })}</tbody>
                        </table>
                      </>)}
                      <div className="p-3 rounded-sm flex justify-between" style={{ background:PAPER, border:`1px solid ${LINE}` }}>
                        <span className="text-sm">Chênh lệch</span><span className="text-sm font-bold" style={{ fontFamily:"'IBM Plex Mono', monospace", color: excTotal-retTotal>=0?FOREST:RUST }}>{vnd(excTotal-retTotal)}</span>
                      </div>
                    </div>
                ) : (
                    <div>
                      <Field label="Hình thức">
                        <div className="flex gap-2">
                          <button onClick={()=>setReturnEditForm(f=>({...f, type:"refund"}))} className="px-3 py-1.5 rounded-sm text-sm border" style={{ borderColor:formR.type==="refund"?INK:LINE, background:formR.type==="refund"?INK:"transparent", color:formR.type==="refund"?"#fff":INK }}>Hoàn tiền</button>
                          <button onClick={()=>setReturnEditForm(f=>({...f, type:"exchange"}))} className="px-3 py-1.5 rounded-sm text-sm border" style={{ borderColor:formR.type==="exchange"?INK:LINE, background:formR.type==="exchange"?INK:"transparent", color:formR.type==="exchange"?"#fff":INK }}>Đổi hàng</button>
                        </div>
                      </Field>
                      <p className="text-xs uppercase tracking-wider mt-3 mb-2 opacity-60">Sản phẩm trả</p>
                      <div className="space-y-2 mb-3">
                        {formR.returnedItems.map(it=>{
                          const p=products.find(x=>x.id===it.productId);
                          return <div key={it.productId} className="flex items-center gap-2 p-2 rounded-sm" style={{ border:`1px solid ${LINE}` }}>
                            <span className="text-xs flex-1">{p?.name}</span>
                            <input type="number" min={0} value={it.qty} onChange={e=> updateRetItem(it.productId, { qty: Math.max(0, Number(e.target.value)) })} className="w-16 border rounded-sm p-1 text-center text-xs" style={{ borderColor:LINE }} />
                            <input type="number" value={it.price} onChange={e=> updateRetItem(it.productId, { price: Number(e.target.value) })} className="w-24 border rounded-sm p-1 text-right text-xs" style={{ borderColor:LINE, fontFamily:"'IBM Plex Mono', monospace" }} />
                          </div>;
                        })}
                      </div>
                      {formR.type==="exchange" && <>
                        <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Sản phẩm đổi</p>
                        <ProductPicker products={products} onPick={addExc} />
                        <div className="mt-2 space-y-2 mb-3">
                          {formR.exchangeItems.map(it=>{
                            const p=products.find(x=>x.id===it.productId);
                            return <div key={it.productId} className="flex items-center gap-2 p-2 rounded-sm" style={{ border:`1px solid ${LINE}` }}>
                              <span className="text-xs flex-1">{p?.name}</span>
                              <input type="number" min={1} value={it.qty} onChange={e=> updateExcItem(it.productId, { qty: Math.max(1, Number(e.target.value)) })} className="w-16 border rounded-sm p-1 text-center text-xs" style={{ borderColor:LINE }} />
                              <input type="number" value={it.price} onChange={e=> updateExcItem(it.productId, { price: Number(e.target.value) })} className="w-24 border rounded-sm p-1 text-right text-xs" style={{ borderColor:LINE, fontFamily:"'IBM Plex Mono', monospace" }} />
                              <button onClick={()=>removeExc(it.productId)} className="opacity-50"><X size={12} /></button>
                            </div>;
                          })}
                        </div>
                      </>}
                      <Field label="Ghi chú"><textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor:LINE }} value={formR.note} onChange={e=> setReturnEditForm(f=>({...f, note:e.target.value}))} /></Field>
                      <div className="flex gap-2 mt-4">
                        <button onClick={()=>setReturnEditForm(null)} className="flex-1 py-2 rounded-sm border text-sm" style={{ borderColor:LINE }}>Hủy sửa</button>
                        <button onClick={saveEdit} className="flex-1 py-2 rounded-sm text-white text-sm" style={{ background:INK }}>Lưu thay đổi</button>
                      </div>
                    </div>
                )}
              </Modal>
          );
        })()}

        {printingOrder && (
            <OrderInvoicePrint order={printingOrder} products={products} customer={customers.find((c) => c.id === printingOrder.customerId)} />
        )}
      </div>
  );
}
function FilterChip({ active, onClick, children, color }) {
  return (
      <button onClick={onClick} className="text-xs px-3 py-1.5 rounded-full border"
              style={{ borderColor: active ? (color || INK) : LINE, background: active ? (color || INK) : "transparent", color: active ? "#fff" : INK }}>
        {children}
      </button>
  );
}

/* ---------------- Reports ---------------- */

function Reports({ orders, products, customers }) {
  const byCategory = useMemo(() => {
    const map = {};
    products.forEach((p) => { const cat = p.category || "Khác"; map[cat] = (map[cat] || 0) + productStats(p).exportedValue; });
    return Object.entries(map).map(([name, total]) => ({ name, total }));
  }, [products]);

  const stockAging = products.map((p) => ({ code: p.code, ton: productStats(p).closingQty, giaTri: Math.round(productStats(p).closingQty * productStats(p).avgCost) }));

  return (
      <div className="grid lg:grid-cols-2 gap-6">
        <div className="p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
          <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6 }}>Doanh số xuất theo nhóm hàng</h4>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={byCategory}>
              <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
              <XAxis dataKey="name" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => `${v / 1000}k`} />
              <Tooltip formatter={(v) => vnd(v)} />
              <Bar dataKey="total" fill={BRASS} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="p-5 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
          <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6 }}>Giá trị tồn kho theo mã VT</h4>
          <div className="space-y-2 max-h-56 overflow-y-auto">
            {stockAging.map((s) => (
                <div key={s.code} className="flex items-center justify-between text-sm py-1.5" style={{ borderBottom: `1px dashed ${LINE}` }}>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{s.code} · {s.ton} đvt</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(s.giaTri)}</span>
                </div>
            ))}
          </div>
        </div>
      </div>
  );
}

/* ---------------- App shell ---------------- */

const TABS = [
  { id: "dashboard", label: "Tổng quan", icon: LayoutDashboard },
  { id: "products", label: "Sản phẩm & Tồn kho", icon: Package },
  { id: "orders", label: "Bán hàng", icon: ShoppingCart },
  { id: "customers", label: "Khách hàng", icon: Users },
  { id: "reports", label: "Báo cáo", icon: BarChart3 },
];

export default function SalesManager() {
  const [tab, setTab] = useState("dashboard");
  const [products, setProducts] = useState([]);
  const [orders, setOrders] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [purchaseOrders, setPurchaseOrders] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    (async () => {
      let data = null;
      try { data = await loadData(); } catch (e) { console.error("Lỗi tải dữ liệu:", e); }
      try {
        if (data) {
          setProducts((data.products || []).map(normalizeProduct));
          setOrders((data.orders || []).map(normalizeOrder));
          setCustomers((data.customers || []).map(normalizeCustomer));
          setPurchaseOrders((data.purchaseOrders || []).map(normalizePO));
          setSuppliers((data.suppliers || []).map(normalizeSupplier));
        } else {
          const seed = seedData();
          setProducts(seed.products.map(normalizeProduct));
          setOrders(seed.orders.map(normalizeOrder));
          setCustomers(seed.customers.map(normalizeCustomer));
          setPurchaseOrders([]);
          setSuppliers([]);
        }
      } catch (e) {
        console.error("Lỗi chuẩn hoá dữ liệu, dùng dữ liệu mẫu:", e);
        const seed = seedData();
        setProducts(seed.products);
        setOrders(seed.orders);
        setCustomers(seed.customers);
        setPurchaseOrders([]);
        setSuppliers([]);
      }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => { saveData({ products, orders, customers, purchaseOrders, suppliers }); }, 400);
    return () => clearTimeout(t);
  }, [products, orders, customers, purchaseOrders, suppliers, loaded]);

  if (!loaded) {
    return <div className="flex items-center justify-center h-96" style={{ color: INK }}><Loader2 className="animate-spin mr-2" size={18} /> Đang tải dữ liệu…</div>;
  }

  return (
      <div style={{ background: PAPER, minHeight: "100%", fontFamily: "'Inter', sans-serif" }} className="w-full overflow-x-hidden">
        <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        select { appearance: none; }
        #invoice-print-area { display: none; }
        @media print {
          #app-shell { display: none !important; }
          #invoice-print-area { display: block !important; }
          @page { size: A4; margin: 12mm; }
        }
      `}</style>
        <div id="app-shell" className="flex flex-col md:flex-row">
          <div className="md:w-60 shrink-0 p-5 md:min-h-screen" style={{ background: INK }}>
            <div className="mb-8">
              <h1 style={{ fontFamily: "'Fraunces', serif", color: "#fff" }} className="text-xl leading-tight">Sổ Bán Hàng</h1>
              <p className="text-[11px] uppercase tracking-widest mt-1" style={{ color: BRASS }}>quản lý sản phẩm &amp; tồn kho</p>
            </div>
            <nav className="flex flex-col gap-1">
              {TABS.map((t) => (
                  <button key={t.id} onClick={() => setTab(t.id)} className="flex items-center gap-2.5 px-3 py-2.5 rounded-sm text-sm transition-colors"
                          style={{ background: tab === t.id ? "rgba(255,255,255,0.1)" : "transparent", color: tab === t.id ? "#fff" : "rgba(255,255,255,0.55)", borderLeft: tab === t.id ? `3px solid ${BRASS}` : "3px solid transparent" }}>
                    <t.icon size={16} /> {t.label}
                  </button>
              ))}
            </nav>
          </div>
          <div className="flex-1 p-5 md:p-8 min-w-0">
            <div className="flex items-center justify-between mb-6">
              <h2 style={{ fontFamily: "'Fraunces', serif", color: INK }} className="text-2xl">{TABS.find((t) => t.id === tab)?.label}</h2>
              <span className="text-xs opacity-40" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{todayISO()}</span>
            </div>
            <AppErrorBoundary key={tab}>
              {tab === "dashboard" && <Dashboard products={products} orders={orders} customers={customers} />}
              {tab === "products" && <ProductsSection products={products} setProducts={setProducts} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} suppliers={suppliers} setSuppliers={setSuppliers} />}
              {tab === "orders" && <Orders orders={orders} setOrders={setOrders} products={products} setProducts={setProducts} customers={customers} setCustomers={setCustomers} />}
              {tab === "customers" && <Customers customers={customers} setCustomers={setCustomers} orders={orders} />}
              {tab === "reports" && <Reports orders={orders} products={products} customers={customers} />}
            </AppErrorBoundary>
          </div>
        </div>
      </div>
  );
}
