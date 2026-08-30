import React, { useState, useEffect, useMemo, useRef } from "react";
import {
    LayoutDashboard, Package, ShoppingCart, Users, BarChart3,
    Plus, Trash2, Pencil, X, Search, Store, Globe,
    TrendingUp, AlertTriangle, Loader2, ChevronDown, ChevronRight, ChevronLeft,
    ArrowDownToLine, ArrowUpFromLine, Barcode, ImagePlus, ImageOff, Check, Printer, RotateCcw, KeyRound, LogOut, Eye, EyeOff, Filter, Target, History, ShieldCheck, XCircle, Wallet, PackageCheck, Truck, Clock, Bell, FileSpreadsheet, FileText, MapPin
} from "lucide-react";
import {
    BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
    ResponsiveContainer, PieChart, Pie, Cell, ComposedChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";

// Xuất 1 hoặc nhiều bảng dữ liệu ra 1 file Excel (.xlsx), mỗi bảng là 1 sheet riêng.
function exportExcel(filename, sheets) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(({ name, rows }) => {
        const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ "Không có dữ liệu": "" }]);
        XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    });
    XLSX.writeFile(wb, filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`);
}

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
const PURPLE = "#6B4FA0";
const STATUSES = [
    { id: "pending", label: "Chờ xử lý", color: BRASS },
    { id: "shipping", label: "Đang giao", color: BLUE },
    { id: "delivered", label: "Đã giao", color: FOREST },
    { id: "done", label: "Hoàn thành", color: PURPLE },
    { id: "cancelled", label: "Đã huỷ", color: RUST },
];
const UNITS = ["Bộ", "Cái", "Hộp", "Thùng", "Chiếc"];
const WARRANTY_OPTIONS = [3, 6, 12, 18, 24, 36, 60]; // số tháng bảo hành — 0 nghĩa là không bảo hành/chưa khai báo
const WARRANTY_LIFETIME = -1; // bảo hành vĩnh viễn
function warrantyLabel(months) { return months === WARRANTY_LIFETIME ? "Vĩnh viễn" : months > 0 ? `${months} tháng` : "Không bảo hành"; }
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
const ACCOUNT_ROLES = [
    { id: "admin", label: "Quản trị viên" },
    { id: "staff", label: "Nhân viên" },
    { id: "ctv", label: "Cộng tác viên" },
];
function normalizeAccount(a) {
    return {
        id: a.id || uid(), username: (a.username || "").trim().toLowerCase(),
        passwordHash: a.passwordHash || "", passwordSalt: a.passwordSalt || "",
        // Cờ tạm để nhận diện tài khoản còn lưu mật khẩu dạng thường (dữ liệu cũ) — sẽ được mã hoá lại ngay khi tải app.
        _legacyPassword: a.passwordHash ? undefined : (a.password || undefined),
        fullName: a.fullName || "", role: ACCOUNT_ROLES.some((r) => r.id === a.role) ? a.role : "staff",
        active: a.active !== false,
    };
}
function seedAccounts() {
    return [normalizeAccount({ username: "admin", password: "admin123", fullName: "Chủ cửa hàng", role: "admin", active: true })];
}
// Mã hoá mật khẩu bằng SHA-256 (Web Crypto API có sẵn trên trình duyệt) kèm salt ngẫu nhiên cho từng tài khoản.
function randomSalt() {
    const arr = crypto.getRandomValues(new Uint8Array(8));
    return Array.from(arr).map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function hashPassword(password, salt) {
    const enc = new TextEncoder();
    const data = enc.encode(`${salt}:${password}`);
    const buf = await crypto.subtle.digest("SHA-256", data);
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function verifyPassword(password, salt, hash) {
    if (!hash) return false;
    const h = await hashPassword(password, salt);
    return h === hash;
}
// Chuyển các tài khoản còn lưu mật khẩu dạng thường (từ dữ liệu cũ) sang dạng đã mã hoá.
async function migrateAccountPasswords(accs) {
    const out = [];
    for (const a of accs) {
        if (a._legacyPassword) {
            const salt = randomSalt();
            const hash = await hashPassword(a._legacyPassword, salt);
            const { _legacyPassword, ...rest } = a;
            out.push({ ...rest, passwordHash: hash, passwordSalt: salt });
        } else {
            const { _legacyPassword, ...rest } = a;
            out.push(rest);
        }
    }
    return out;
}
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
    "Thành phố Cần Thơ", "Thành phố Đà Nẵng", "Thành phố Hà Nội", "Thành phố Hải Phòng",
    "Thành phố Hồ Chí Minh", "Thành phố Huế", "Tỉnh An Giang", "Tỉnh Bắc Ninh",
    "Tỉnh Cà Mau", "Tỉnh Cao Bằng", "Tỉnh Đắk Lắk", "Tỉnh Điện Biên",
    "Tỉnh Đồng Nai", "Tỉnh Đồng Tháp", "Tỉnh Gia Lai", "Tỉnh Hà Tĩnh",
    "Tỉnh Hưng Yên", "Tỉnh Khánh Hòa", "Tỉnh Lai Châu", "Tỉnh Lâm Đồng",
    "Tỉnh Lạng Sơn", "Tỉnh Lào Cai", "Tỉnh Nghệ An", "Tỉnh Ninh Bình",
    "Tỉnh Phú Thọ", "Tỉnh Quảng Ngãi", "Tỉnh Quảng Ninh", "Tỉnh Quảng Trị",
    "Tỉnh Sơn La", "Tỉnh Tây Ninh", "Tỉnh Thái Nguyên", "Tỉnh Thanh Hóa",
    "Tỉnh Tuyên Quang", "Tỉnh Vĩnh Long",
];

// Danh sách phường/xã theo từng tỉnh/thành sau sáp nhập (hiệu lực 1/7/2025) — nguồn dữ liệu do người dùng cung cấp.
const WARDS_BY_PROVINCE = {
    "Thành phố Cần Thơ": ["Phường An Bình", "Phường Bình Thủy", "Phường Cái Khế", "Phường Cái Răng", "Phường Đại Thành", "Phường Hưng Phú", "Phường Khánh Hòa", "Phường Long Bình", "Phường Long Mỹ", "Phường Long Phú 1", "Phường Long Tuyền", "Phường Mỹ Quới", "Phường Mỹ Xuyên", "Phường Ngã Bảy", "Phường Ngã Năm", "Phường Ninh Kiều", "Phường Ô Môn", "Phường Phú Lợi", "Phường Phước Thới", "Phường Sóc Trăng", "Phường Tân An", "Phường Tân Lộc", "Phường Thới An Đông", "Phường Thới Long", "Phường Thốt Nốt", "Phường Thuận Hưng", "Phường Trung Nhứt", "Phường Vị Tân", "Phường Vị Thanh", "Phường Vĩnh Châu", "Phường Vĩnh Phước", "Xã An Lạc Thôn", "Xã An Ninh", "Xã An Thạnh", "Xã Châu Thành", "Xã Cờ Đỏ", "Xã Cù Lao Dung", "Xã Đại Hải", "Xã Đại Ngãi", "Xã Đông Hiệp", "Xã Đông Phước", "Xã Đông Thuận", "Xã Gia Hòa", "Xã Hiệp Hưng", "Xã Hồ Đắc Kiện", "Xã Hòa An", "Xã Hỏa Lựu", "Xã Hòa Tú", "Xã Kế Sách", "Xã Lai Hòa", "Xã Lâm Tân", "Xã Lịch Hội Thượng", "Xã Liêu Tú", "Xã Long Hưng", "Xã Long Phú", "Xã Lương Tâm", "Xã Mỹ Hương", "Xã Mỹ Phước", "Xã Mỹ Tú", "Xã Ngọc Tố", "Xã Nhơn Ái", "Xã Nhơn Mỹ", "Xã Nhu Gia", "Xã Phong Điền", "Xã Phong Nẫm", "Xã Phú Hữu", "Xã Phú Lộc", "Xã Phú Tâm", "Xã Phụng Hiệp", "Xã Phương Bình", "Xã Tài Văn", "Xã Tân Bình", "Xã Tân Hòa", "Xã Tân Long", "Xã Tân Phước Hưng", "Xã Tân Thạnh", "Xã Thạnh An", "Xã Thạnh Hòa", "Xã Thạnh Phú", "Xã Thạnh Quới", "Xã Thạnh Thới An", "Xã Thạnh Xuân", "Xã Thới An Hội", "Xã Thới Hưng", "Xã Thới Lai", "Xã Thuận Hòa", "Xã Trần Đề", "Xã Trung Hưng", "Xã Trường Khánh", "Xã Trường Long", "Xã Trường Long Tây", "Xã Trường Thành", "Xã Trường Xuân", "Xã Vị Thanh 1", "Xã Vị Thủy", "Xã Vĩnh Hải", "Xã Vĩnh Lợi", "Xã Vĩnh Thạnh", "Xã Vĩnh Thuận Đông", "Xã Vĩnh Trinh", "Xã Vĩnh Tường", "Xã Vĩnh Viễn", "Xã Xà Phiên"],
    "Thành phố Đà Nẵng": ["Đặc khu Hoàng Sa", "Phường An Hải", "Phường An Khê", "Phường An Thắng", "Phường Bàn Thạch", "Phường Cẩm Lệ", "Phường Điện Bàn", "Phường Điện Bàn Bắc", "Phường Điện Bàn Đông", "Phường Hải Châu", "Phường Hải Vân", "Phường Hòa Cường", "Phường Hòa Khánh", "Phường Hòa Xuân", "Phường Hội An", "Phường Hội An Đông", "Phường Hội An Tây", "Phường Hương Trà", "Phường Liên Chiểu", "Phường Ngũ Hành Sơn", "Phường Quảng Phú", "Phường Sơn Trà", "Phường Tam Kỳ", "Phường Thanh Khê", "Xã Avương", "Xã Bà Nà", "Xã Bến Giằng", "Xã Bến Hiên", "Xã Chiên Đàn", "Xã Đắc Pring", "Xã Đại Lộc", "Xã Điện Bàn Tây", "Xã Đồng Dương", "Xã Đông Giang", "Xã Đức Phú", "Xã Duy Nghĩa", "Xã Duy Xuyên", "Xã Gò Nổi", "Xã Hà Nha", "Xã Hiệp Đức", "Xã Hòa Tiến", "Xã Hòa Vang", "Xã Hùng Sơn", "Xã Khâm Đức", "Xã La Dêê", "Xã La Êê", "Xã Lãnh Ngọc", "Xã Nam Giang", "Xã Nam Phước", "Xã Nam Trà My", "Xã Nông Sơn", "Xã Núi Thành", "Xã Phú Ninh", "Xã Phú Thuận", "Xã Phước Chánh", "Xã Phước Hiệp", "Xã Phước Năng", "Xã Phước Thành", "Xã Phước Trà", "Xã Quế Phước", "Xã Quế Sơn", "Xã Quế Sơn Trung", "Xã Sơn Cẩm Hà", "Xã Sông Kôn", "Xã Sông Vàng", "Xã Tam Anh", "Xã Tam Hải", "Xã Tam Mỹ", "Xã Tam Xuân", "Xã Tân Hiệp", "Xã Tây Giang", "Xã Tây Hồ", "Xã Thăng An", "Xã Thăng Bình", "Xã Thăng Điền", "Xã Thăng Phú", "Xã Thăng Trường", "Xã Thạnh Bình", "Xã Thạnh Mỹ", "Xã Thu Bồn", "Xã Thượng Đức", "Xã Tiên Phước", "Xã Trà Đốc", "Xã Trà Giáp", "Xã Trà Leng", "Xã Trà Liên", "Xã Trà Linh", "Xã Trà My", "Xã Trà Tân", "Xã Trà Tập", "Xã Trà Vân", "Xã Việt An", "Xã Vu Gia", "Xã Xuân Phú"],
    "Thành phố Hà Nội": ["Phường Ba Đình", "Phường Bạch Mai", "Phường Bồ Đề", "Phường Cầu Giấy", "Phường Chương Mỹ", "Phường Cửa Nam", "Phường Đại Mỗ", "Phường Định Công", "Phường Đống Đa", "Phường Đông Ngạc", "Phường Dương Nội", "Phường Giảng Võ", "Phường Hà Đông", "Phường Hai Bà Trưng", "Phường Hoàn Kiếm", "Phường Hoàng Liệt", "Phường Hoàng Mai", "Phường Hồng Hà", "Phường Khương Đình", "Phường Kiến Hưng", "Phường Kim Liên", "Phường Láng", "Phường Lĩnh Nam", "Phường Long Biên", "Phường Nghĩa Đô", "Phường Ngọc Hà", "Phường Ô Chợ Dừa", "Phường Phú Diễn", "Phường Phú Lương", "Phường Phú Thượng", "Phường Phúc Lợi", "Phường Phương Liệt", "Phường Sơn Tây", "Phường Tây Hồ", "Phường Tây Mỗ", "Phường Tây Tựu", "Phường Thanh Liệt", "Phường Thanh Xuân", "Phường Thượng Cát", "Phường Từ Liêm", "Phường Tùng Thiện", "Phường Tương Mai", "Phường Văn Miếu - Quốc Tử Giám", "Phường Việt Hưng", "Phường Vĩnh Hưng", "Phường Vĩnh Tuy", "Phường Xuân Đỉnh", "Phường Xuân Phương", "Phường Yên Hòa", "Phường Yên Nghĩa", "Phường Yên Sở", "Xã An Khánh", "Xã Ba Vì", "Xã Bất Bạt", "Xã Bát Tràng", "Xã Bình Minh", "Xã Chương Dương", "Xã Chuyên Mỹ", "Xã Cổ Đô", "Xã Đa Phúc", "Xã Đại Thanh", "Xã Đại Xuyên", "Xã Dân Hòa", "Xã Đan Phượng", "Xã Đoài Phương", "Xã Đông Anh", "Xã Dương Hòa", "Xã Gia Lâm", "Xã Hạ Bằng", "Xã Hát Môn", "Xã Hòa Lạc", "Xã Hòa Phú", "Xã Hòa Xá", "Xã Hoài Đức", "Xã Hồng Sơn", "Xã Hồng Vân", "Xã Hưng Đạo", "Xã Hương Sơn", "Xã Kiều Phú", "Xã Kim Anh", "Xã Liên Minh", "Xã Mê Linh", "Xã Minh Châu", "Xã Mỹ Đức", "Xã Nam Phù", "Xã Ngọc Hồi", "Xã Nội Bài", "Xã Ô Diên", "Xã Phú Cát", "Xã Phù Đổng", "Xã Phú Nghĩa", "Xã Phú Xuyên", "Xã Phúc Lộc", "Xã Phúc Sơn", "Xã Phúc Thịnh", "Xã Phúc Thọ", "Xã Phượng Dực", "Xã Quảng Bị", "Xã Quang Minh", "Xã Quảng Oai", "Xã Quốc Oai", "Xã Sóc Sơn", "Xã Sơn Đồng", "Xã Suối Hai", "Xã Tam Hưng", "Xã Tây Phương", "Xã Thạch Thất", "Xã Thanh Oai", "Xã Thanh Trì", "Xã Thiên Lộc", "Xã Thư Lâm", "Xã Thuận An", "Xã Thượng Phúc", "Xã Thường Tín", "Xã Tiến Thắng", "Xã Trần Phú", "Xã Trung Giã", "Xã Ứng Hòa", "Xã Ứng Thiên", "Xã Vân Đình", "Xã Vật Lại", "Xã Vĩnh Thanh", "Xã Xuân Mai", "Xã Yên Bài", "Xã Yên Lãng", "Xã Yên Xuân"],
    "Thành phố Hải Phòng": ["Đặc khu Bạch Long Vĩ", "Đặc khu Cát Hải", "Phường Ái Quốc", "Phường An Biên", "Phường An Dương", "Phường An Hải", "Phường An Phong", "Phường Bắc An Phụ", "Phường Bạch Đằng", "Phường Chí Linh", "Phường Chu Văn An", "Phường Đồ Sơn", "Phường Đông Hải", "Phường Dương Kinh", "Phường Gia Viên", "Phường Hải An", "Phường Hải Dương", "Phường Hòa Bình", "Phường Hồng An", "Phường Hồng Bàng", "Phường Hưng Đạo", "Phường Kiến An", "Phường Kinh Môn", "Phường Lê Chân", "Phường Lê Đại Hành", "Phường Lê Ích Mộc", "Phường Lê Thanh Nghị", "Phường Lưu Kiếm", "Phường Nam Đồ Sơn", "Phường Nam Đồng", "Phường Nam Triệu", "Phường Ngô Quyền", "Phường Nguyễn Đại Năng", "Phường Nguyễn Trãi", "Phường Nhị Chiểu", "Phường Phạm Sư Mạnh", "Phường Phù Liễn", "Phường Tân Hưng", "Phường Thạch Khôi", "Phường Thành Đông", "Phường Thiên Hương", "Phường Thủy Nguyên", "Phường Trần Hưng Đạo", "Phường Trần Liễu", "Phường Trần Nhân Tông", "Phường Tứ Minh", "Phường Việt Hòa", "Xã An Hưng", "Xã An Khánh", "Xã An Lão", "Xã An Phú", "Xã An Quang", "Xã An Thành", "Xã An Trường", "Xã Bắc Thanh Miện", "Xã Bình Giang", "Xã Cẩm Giàng", "Xã Cẩm Giang", "Xã Chấn Hưng", "Xã Chí Minh", "Xã Đại Sơn", "Xã Đường An", "Xã Gia Lộc", "Xã Gia Phúc", "Xã Hà Bắc", "Xã Hà Đông", "Xã Hà Nam", "Xã Hà Tây", "Xã Hải Hưng", "Xã Hồng Châu", "Xã Hợp Tiến", "Xã Hùng Thắng", "Xã Kẻ Sặt", "Xã Khúc Thừa Dụ", "Xã Kiến Hải", "Xã Kiến Hưng", "Xã Kiến Minh", "Xã Kiến Thụy", "Xã Kim Thành", "Xã Lạc Phượng", "Xã Lai Khê", "Xã Mao Điền", "Xã Nam An Phụ", "Xã Nam Sách", "Xã Nam Thanh Miện", "Xã Nghi Dương", "Xã Nguyễn Bỉnh Khiêm", "Xã Nguyên Giáp", "Xã Nguyễn Lương Bằng", "Xã Ninh Giang", "Xã Phú Thái", "Xã Quyết Thắng", "Xã Tân An", "Xã Tân Kỳ", "Xã Tân Minh", "Xã Thái Tân", "Xã Thanh Hà", "Xã Thanh Miện", "Xã Thượng Hồng", "Xã Tiên Lãng", "Xã Tiên Minh", "Xã Trần Phú", "Xã Trường Tân", "Xã Tứ Kỳ", "Xã Tuệ Tĩnh", "Xã Việt Khê", "Xã Vĩnh Am", "Xã Vĩnh Bảo", "Xã Vĩnh Hải", "Xã Vĩnh Hòa", "Xã Vĩnh Lại", "Xã Vĩnh Thịnh", "Xã Vĩnh Thuận", "Xã Yết Kiêu"],
    "Thành phố Hồ Chí Minh": ["Đặc khu Côn Đảo", "Phường An Đông", "Phường An Hội Đông", "Phường An Hội Tây", "Phường An Khánh", "Phường An Lạc", "Phường An Nhơn", "Phường An Phú", "Phường An Phú Đông", "Phường Bà Rịa", "Phường Bàn Cờ", "Phường Bảy Hiền", "Phường Bến Cát", "Phường Bến Thành", "Phường Bình Cơ", "Phường Bình Đông", "Phường Bình Dương", "Phường Bình Hòa", "Phường Bình Hưng Hòa", "Phường Bình Lợi Trung", "Phường Bình Phú", "Phường Bình Quới", "Phường Bình Tân", "Phường Bình Tây", "Phường Bình Thạnh", "Phường Bình Thới", "Phường Bình Tiên", "Phường Bình Trị Đông", "Phường Bình Trưng", "Phường Cát Lái", "Phường Cầu Kiệu", "Phường Cầu Ông Lãnh", "Phường Chánh Hiệp", "Phường Chánh Hưng", "Phường Chánh Phú Hòa", "Phường Chợ Lớn", "Phường Chợ Quán", "Phường Dĩ An", "Phường Diên Hồng", "Phường Đông Hòa", "Phường Đông Hưng Thuận", "Phường Đức Nhuận", "Phường Gia Định", "Phường Gò Vấp", "Phường Hạnh Thông", "Phường Hiệp Bình", "Phường Hòa Bình", "Phường Hòa Hưng", "Phường Hòa Lợi", "Phường Khánh Hội", "Phường Lái Thiêu", "Phường Linh Xuân", "Phường Long Bình", "Phường Long Hương", "Phường Long Nguyên", "Phường Long Phước", "Phường Long Trường", "Phường Minh Phụng", "Phường Nhiêu Lộc", "Phường Phú An", "Phường Phú Định", "Phường Phú Lâm", "Phường Phú Lợi", "Phường Phú Mỹ", "Phường Phú Nhuận", "Phường Phú Thạnh", "Phường Phú Thọ", "Phường Phú Thọ Hòa", "Phường Phú Thuận", "Phường Phước Long", "Phường Phước Thắng", "Phường Rạch Dừa", "Phường Sài Gòn", "Phường Tam Bình", "Phường Tam Long", "Phường Tam Thắng", "Phường Tân Bình", "Phường Tân Định", "Phường Tân Đông Hiệp", "Phường Tân Hải", "Phường Tân Hiệp", "Phường Tân Hòa", "Phường Tân Hưng", "Phường Tân Khánh", "Phường Tân Mỹ", "Phường Tân Phú", "Phường Tân Phước", "Phường Tân Sơn", "Phường Tân Sơn Hòa", "Phường Tân Sơn Nhất", "Phường Tân Sơn Nhì", "Phường Tân Tạo", "Phường Tân Thành", "Phường Tân Thới Hiệp", "Phường Tân Thuận", "Phường Tân Uyên", "Phường Tăng Nhơn Phú", "Phường Tây Nam", "Phường Tây Thạnh", "Phường Thạnh Mỹ Tây", "Phường Thới An", "Phường Thới Hòa", "Phường Thông Tây Hội", "Phường Thủ Dầu Một", "Phường Thủ Đức", "Phường Thuận An", "Phường Thuận Giao", "Phường Trung Mỹ Tây", "Phường Vĩnh Hội", "Phường Vĩnh Tân", "Phường Vũng Tàu", "Phường Vườn Lài", "Phường Xóm Chiếu", "Phường Xuân Hòa", "Xã An Long", "Xã An Nhơn Tây", "Xã An Thới Đông", "Xã Bà Điểm", "Xã Bắc Tân Uyên", "Xã Bàu Bàng", "Xã Bàu Lâm", "Xã Bình Chánh", "Xã Bình Châu", "Xã Bình Giã", "Xã Bình Hưng", "Xã Bình Khánh", "Xã Bình Lợi", "Xã Bình Mỹ", "Xã Cần Giờ", "Xã Châu Đức", "Xã Châu Pha", "Xã Củ Chi", "Xã Đất Đỏ", "Xã Dầu Tiếng", "Xã Đông Thạnh", "Xã Hiệp Phước", "Xã Hồ Tràm", "Xã Hòa Hiệp", "Xã Hòa Hội", "Xã Hóc Môn", "Xã Hưng Long", "Xã Kim Long", "Xã Long Điền", "Xã Long Hải", "Xã Long Hòa", "Xã Long Sơn", "Xã Minh Thạnh", "Xã Ngãi Giao", "Xã Nghĩa Thành", "Xã Nhà Bè", "Xã Nhuận Đức", "Xã Phú Giáo", "Xã Phú Hòa Đông", "Xã Phước Hải", "Xã Phước Hòa", "Xã Phước Thành", "Xã Tân An Hội", "Xã Tân Nhựt", "Xã Tân Vĩnh Lộc", "Xã Thái Mỹ", "Xã Thạnh An", "Xã Thanh An", "Xã Thường Tân", "Xã Trừ Văn Thố", "Xã Vĩnh Lộc", "Xã Xuân Sơn", "Xã Xuân Thới Sơn", "Xã Xuyên Mộc"],
    "Thành phố Huế": ["Phường An Cựu", "Phường Dương Nỗ", "Phường Hóa Châu", "Phường Hương An", "Phường Hương Thủy", "Phường Hương Trà", "Phường Kim Long", "Phường Kim Trà", "Phường Mỹ Thượng", "Phường Phong Điền", "Phường Phong Dinh", "Phường Phong Phú", "Phường Phong Quảng", "Phường Phong Thái", "Phường Phú Bài", "Phường Phú Xuân", "Phường Thanh Thủy", "Phường Thuận An", "Phường Thuận Hóa", "Phường Thủy Xuân", "Phường Vỹ Dạ", "Xã A Lưới 1", "Xã A Lưới 2", "Xã A Lưới 3", "Xã A Lưới 4", "Xã A Lưới 5", "Xã Bình Điền", "Xã Chân Mây - Lăng Cô", "Xã Đan Điền", "Xã Hưng Lộc", "Xã Khe Tre", "Xã Lộc An", "Xã Long Quảng", "Xã Nam Đông", "Xã Phú Hồ", "Xã Phú Lộc", "Xã Phú Vang", "Xã Phú Vinh", "Xã Quảng Điền", "Xã Vinh Lộc"],
    "Tỉnh An Giang": ["Đặc khu Kiên Hải", "Đặc khu Phú Quốc", "Đặc khu Thổ Châu", "Phường Bình Đức", "Phường Châu Đốc", "Phường Chi Lăng", "Phường Hà Tiên", "Phường Long Phú", "Phường Long Xuyên", "Phường Mỹ Thới", "Phường Rạch Giá", "Phường Tân Châu", "Phường Thới Sơn", "Phường Tịnh Biên", "Phường Tô Châu", "Phường Vĩnh Tế", "Phường Vĩnh Thông", "Xã An Biên", "Xã An Châu", "Xã An Cư", "Xã An Minh", "Xã An Phú", "Xã Ba Chúc", "Xã Bình An", "Xã Bình Giang", "Xã Bình Hòa", "Xã Bình Mỹ", "Xã Bình Sơn", "Xã Bình Thạnh Đông", "Xã Cần Đăng", "Xã Châu Phong", "Xã Châu Phú", "Xã Châu Thành", "Xã Chợ Mới", "Xã Chợ Vàm", "Xã Cô Tô", "Xã Cù Lao Giêng", "Xã Định Hòa", "Xã Định Mỹ", "Xã Đông Hòa", "Xã Đông Hưng", "Xã Đông Thái", "Xã Giang Thành", "Xã Giồng Riềng", "Xã Gò Quao", "Xã Hòa Điền", "Xã Hòa Hưng", "Xã Hòa Lạc", "Xã Hòa Thuận", "Xã Hội An", "Xã Hòn Đất", "Xã Hòn Nghệ", "Xã Khánh Bình", "Xã Kiên Lương", "Xã Long Điền", "Xã Long Kiến", "Xã Long Thạnh", "Xã Mỹ Đức", "Xã Mỹ Hòa Hưng", "Xã Mỹ Thuận", "Xã Ngọc Chúc", "Xã Nhơn Hội", "Xã Nhơn Mỹ", "Xã Núi Cấm", "Xã Ô Lâm", "Xã Óc Eo", "Xã Phú An", "Xã Phú Hòa", "Xã Phú Hữu", "Xã Phú Lâm", "Xã Phú Tân", "Xã Sơn Hải", "Xã Sơn Kiên", "Xã Tân An", "Xã Tân Hiệp", "Xã Tân Hội", "Xã Tân Thạnh", "Xã Tây Phú", "Xã Tây Yên", "Xã Thạnh Đông", "Xã Thạnh Hưng", "Xã Thạnh Lộc", "Xã Thạnh Mỹ Tây", "Xã Thoại Sơn", "Xã Tiên Hải", "Xã Tri Tôn", "Xã U Minh Thượng", "Xã Vân Khánh", "Xã Vĩnh An", "Xã Vĩnh Bình", "Xã Vĩnh Điều", "Xã Vĩnh Gia", "Xã Vĩnh Hanh", "Xã Vĩnh Hậu", "Xã Vĩnh Hòa", "Xã Vĩnh Hòa Hưng", "Xã Vĩnh Phong", "Xã Vĩnh Thạnh Trung", "Xã Vĩnh Thuận", "Xã Vĩnh Trạch", "Xã Vĩnh Tuy", "Xã Vĩnh Xương"],
    "Tỉnh Bắc Ninh": ["Phường Bắc Giang", "Phường Bồng Lai", "Phường Cảnh Thụy", "Phường Chũ", "Phường Đa Mai", "Phường Đào Viên", "Phường Đồng Nguyên", "Phường Hạp Lĩnh", "Phường Kinh Bắc", "Phường Mão Điền", "Phường Nam Sơn", "Phường Nếnh", "Phường Nhân Hòa", "Phường Ninh Xá", "Phường Phù Khê", "Phường Phương Liễu", "Phường Phượng Sơn", "Phường Quế Võ", "Phường Song Liễu", "Phường Tam Sơn", "Phường Tân An", "Phường Tân Tiến", "Phường Thuận Thành", "Phường Tiền Phong", "Phường Trạm Lộ", "Phường Trí Quả", "Phường Tự Lạn", "Phường Từ Sơn", "Phường Vân Hà", "Phường Việt Yên", "Phường Võ Cường", "Phường Vũ Ninh", "Phường Yên Dũng", "Xã An Lạc", "Xã Bắc Lũng", "Xã Bảo Đài", "Xã Biển Động", "Xã Biên Sơn", "Xã Bố Hạ", "Xã Cẩm Lý", "Xã Cao Đức", "Xã Chi Lăng", "Xã Đại Đồng", "Xã Đại Lai", "Xã Đại Sơn", "Xã Đèo Gia", "Xã Đông Cứu", "Xã Đồng Kỳ", "Xã Đông Phú", "Xã Đồng Việt", "Xã Dương Hưu", "Xã Gia Bình", "Xã Hiệp Hòa", "Xã Hoàng Vân", "Xã Hợp Thịnh", "Xã Kép", "Xã Kiên Lao", "Xã Lâm Thao", "Xã Lạng Giang", "Xã Liên Bão", "Xã Lục Nam", "Xã Lục Ngạn", "Xã Lục Sơn", "Xã Lương Tài", "Xã Mỹ Thái", "Xã Nam Dương", "Xã Nghĩa Phương", "Xã Ngọc Thiện", "Xã Nhã Nam", "Xã Nhân Thắng", "Xã Phật Tích", "Xã Phù Lãng", "Xã Phúc Hòa", "Xã Quang Trung", "Xã Sa Lý", "Xã Sơn Động", "Xã Sơn Hải", "Xã Tam Đa", "Xã Tam Giang", "Xã Tam Tiến", "Xã Tân Chi", "Xã Tân Dĩnh", "Xã Tân Sơn", "Xã Tân Yên", "Xã Tây Yên Tử", "Xã Tiên Du", "Xã Tiên Lục", "Xã Trung Chính", "Xã Trung Kênh", "Xã Trường Sơn", "Xã Tuấn Đạo", "Xã Văn Môn", "Xã Vân Sơn", "Xã Xuân Cẩm", "Xã Xuân Lương", "Xã Yên Định", "Xã Yên Phong", "Xã Yên Thế", "Xã Yên Trung"],
    "Tỉnh Cà Mau": ["Phường An Xuyên", "Phường Bạc Liêu", "Phường Giá Rai", "Phường Hiệp Thành", "Phường Hòa Thành", "Phường Láng Tròn", "Phường Lý Văn Lâm", "Phường Tân Thành", "Phường Vĩnh Trạch", "Xã An Trạch", "Xã Biển Bạch", "Xã Cái Đôi Vàm", "Xã Cái Nước", "Xã Châu Thới", "Xã Đá Bạc", "Xã Đầm Dơi", "Xã Đất Mới", "Xã Đất Mũi", "Xã Định Thành", "Xã Đông Hải", "Xã Gành Hào", "Xã Hồ Thị Kỷ", "Xã Hòa Bình", "Xã Hồng Dân", "Xã Hưng Hội", "Xã Hưng Mỹ", "Xã Khánh An", "Xã Khánh Bình", "Xã Khánh Hưng", "Xã Khánh Lâm", "Xã Long Điền", "Xã Lương Thế Trân", "Xã Năm Căn", "Xã Nguyễn Phích", "Xã Nguyễn Việt Khái", "Xã Ninh Quới", "Xã Ninh Thạnh Lợi", "Xã Phan Ngọc Hiển", "Xã Phong Hiệp", "Xã Phong Thạnh", "Xã Phú Mỹ", "Xã Phú Tân", "Xã Phước Long", "Xã Quách Phẩm", "Xã Sông Đốc", "Xã Tạ An Khương", "Xã Tam Giang", "Xã Tân Ân", "Xã Tân Hưng", "Xã Tân Lộc", "Xã Tân Thuận", "Xã Tân Tiến", "Xã Thanh Tùng", "Xã Thới Bình", "Xã Trần Phán", "Xã Trần Văn Thời", "Xã Trí Phải", "Xã U Minh", "Xã Vĩnh Hậu", "Xã Vĩnh Lộc", "Xã Vĩnh Lợi", "Xã Vĩnh Mỹ", "Xã Vĩnh Phước", "Xã Vĩnh Thanh"],
    "Tỉnh Cao Bằng": ["Phường Nùng Trí Cao", "Phường Tân Giang", "Phường Thục Phán", "Xã Bạch Đằng", "Xã Bảo Lạc", "Xã Bảo Lâm", "Xã Bế Văn Đàn", "Xã Ca Thành", "Xã Cần Yên", "Xã Canh Tân", "Xã Cô Ba", "Xã Cốc Pàng", "Xã Đàm Thủy", "Xã Đình Phong", "Xã Đoài Dương", "Xã Độc Lập", "Xã Đông Khê", "Xã Đức Long", "Xã Hạ Lang", "Xã Hà Quảng", "Xã Hạnh Phúc", "Xã Hòa An", "Xã Hưng Đạo", "Xã Huy Giáp", "Xã Khánh Xuân", "Xã Kim Đồng", "Xã Lũng Nặm", "Xã Lý Bôn", "Xã Lý Quốc", "Xã Minh Khai", "Xã Minh Tâm", "Xã Nam Quang", "Xã Nam Tuấn", "Xã Nguyên Bình", "Xã Nguyễn Huệ", "Xã Phan Thanh", "Xã Phục Hòa", "Xã Quang Hán", "Xã Quảng Lâm", "Xã Quang Long", "Xã Quang Trung", "Xã Quảng Uyên", "Xã Sơn Lộ", "Xã Tam Kim", "Xã Thạch An", "Xã Thành Công", "Xã Thanh Long", "Xã Thông Nông", "Xã Tĩnh Túc", "Xã Tổng Cọt", "Xã Trà Lĩnh", "Xã Trùng Khánh", "Xã Trường Hà", "Xã Vinh Quý", "Xã Xuân Trường", "Xã Yên Thổ"],
    "Tỉnh Đắk Lắk": ["Phường Bình Kiến", "Phường Buôn Hồ", "Phường Buôn Ma Thuột", "Phường Cư Bao", "Phường Đông Hòa", "Phường Ea Kao", "Phường Hòa Hiệp", "Phường Phú Yên", "Phường Sông Cầu", "Phường Tân An", "Phường Tân Lập", "Phường Thành Nhất", "Phường Tuy Hòa", "Phường Xuân Đài", "Xã Buôn Đôn", "Xã Cư M'gar", "Xã Cư M'ta", "Xã Cư Pơng", "Xã Cư Prao", "Xã Cư Pui", "Xã Cư Yang", "Xã Cuôr Đăng", "Xã Đắk Liêng", "Xã Đắk Phơi", "Xã Dang Kang", "Xã Dliê Ya", "Xã Đồng Xuân", "Xã Dray Bhăng", "Xã Đức Bình", "Xã Dur Kmăl", "Xã Ea Bá", "Xã Ea Bung", "Xã Ea Drăng", "Xã Ea Drông", "Xã Ea H'Leo", "Xã Ea Hiao", "Xã Ea Kar", "Xã Ea Khăl", "Xã Ea Kiết", "Xã Ea Kly", "Xã Ea Knốp", "Xã Ea Knuếc", "Xã Ea Ktur", "Xã Ea Ly", "Xã Ea M'Droh", "Xã Ea Na", "Xã Ea Ning", "Xã Ea Nuôl", "Xã Ea Ô", "Xã Ea Păl", "Xã Ea Phê", "Xã Ea Riêng", "Xã Ea Rốk", "Xã Ea Súp", "Xã Ea Trang", "Xã Ea Tul", "Xã Ea Wer", "Xã Ea Wy", "Xã Hòa Mỹ", "Xã Hòa Phú", "Xã Hòa Sơn", "Xã Hòa Thịnh", "Xã Hòa Xuân", "Xã Ia Lốp", "Xã Ia Rvê", "Xã Krông Á", "Xã Krông Ana", "Xã Krông Bông", "Xã Krông Búk", "Xã Krông Năng", "Xã Krông Nô", "Xã Krông Pắc", "Xã Liên Sơn Lắk", "Xã M'Drắk", "Xã Nam Ka", "Xã Ô Loan", "Xã Phú Hòa 1", "Xã Phú Hòa 2", "Xã Phú Mỡ", "Xã Phú Xuân", "Xã Pơng Drang", "Xã Quảng Phú", "Xã Sơn Hòa", "Xã Sơn Thành", "Xã Sông Hinh", "Xã Suối Trai", "Xã Tam Giang", "Xã Tân Tiến", "Xã Tây Hòa", "Xã Tây Sơn", "Xã Tuy An Bắc", "Xã Tuy An Đông", "Xã Tuy An Nam", "Xã Tuy An Tây", "Xã Vân Hòa", "Xã Vụ Bổn", "Xã Xuân Cảnh", "Xã Xuân Lãnh", "Xã Xuân Lộc", "Xã Xuân Phước", "Xã Xuân Thọ", "Xã Yang Mao"],
    "Tỉnh Điện Biên": ["Phường Điện Biên Phủ", "Phường Mường Lay", "Phường Mường Thanh", "Xã Búng Lao", "Xã Chà Tở", "Xã Chiềng Sinh", "Xã Mường Ảng", "Xã Mường Chà", "Xã Mường Lạn", "Xã Mường Luân", "Xã Mường Mùn", "Xã Mường Nhà", "Xã Mường Nhé", "Xã Mường Phăng", "Xã Mường Pồn", "Xã Mường Toong", "Xã Mường Tùng", "Xã Nà Bủng", "Xã Nà Hỳ", "Xã Na Sang", "Xã Na Son", "Xã Nà Tấu", "Xã Nậm Kè", "Xã Nậm Nèn", "Xã Núa Ngam", "Xã Pa Ham", "Xã Phình Giàng", "Xã Pu Nhi", "Xã Pú Nhung", "Xã Quài Tở", "Xã Quảng Lâm", "Xã Sam Mứn", "Xã Sáng Nhè", "Xã Si Pa Phìn", "Xã Sín Chải", "Xã Sín Thầu", "Xã Sính Phình", "Xã Thanh An", "Xã Thanh Nưa", "Xã Thanh Yên", "Xã Tìa Dình", "Xã Tủa Chùa", "Xã Tủa Thàng", "Xã Tuần Giáo", "Xã Xa Dung"],
    "Tỉnh Đồng Nai": ["Phường An Lộc", "Phường Bảo Vinh", "Phường Biên Hòa", "Phường Bình Lộc", "Phường Bình Long", "Phường Bình Phước", "Phường Chơn Thành", "Phường Đồng Xoài", "Phường Hàng Gòn", "Phường Hố Nai", "Phường Long Bình", "Phường Long Hưng", "Phường Long Khánh", "Phường Minh Hưng", "Phường Phước Bình", "Phường Phước Long", "Phường Phước Tân", "Phường Tam Hiệp", "Phường Tam Phước", "Phường Tân Triều", "Phường Trấn Biên", "Phường Trảng Dài", "Phường Xuân Lập", "Xã An Phước", "Xã An Viễn", "Xã Bàu Hàm", "Xã Bình An", "Xã Bình Minh", "Xã Bình Tân", "Xã Bom Bo", "Xã Bù Đăng", "Xã Bù Gia Mập", "Xã Cẩm Mỹ", "Xã Đa Kia", "Xã Đại Phước", "Xã Đak Lua", "Xã Đak Nhau", "Xã Đăk Ơ", "Xã Dầu Giây", "Xã Định Quán", "Xã Đồng Phú", "Xã Đồng Tâm", "Xã Gia Kiệm", "Xã Hưng Phước", "Xã Hưng Thịnh", "Xã La Ngà", "Xã Lộc Hưng", "Xã Lộc Ninh", "Xã Lộc Quang", "Xã Lộc Tấn", "Xã Lộc Thạnh", "Xã Lộc Thành", "Xã Long Hà", "Xã Long Phước", "Xã Long Thành", "Xã Minh Đức", "Xã Nam Cát Tiên", "Xã Nghĩa Trung", "Xã Nha Bích", "Xã Nhơn Trạch", "Xã Phú Hòa", "Xã Phú Lâm", "Xã Phú Lý", "Xã Phú Nghĩa", "Xã Phú Riềng", "Xã Phú Trung", "Xã Phú Vinh", "Xã Phước An", "Xã Phước Sơn", "Xã Phước Thái", "Xã Sông Ray", "Xã Tà Lài", "Xã Tân An", "Xã Tân Hưng", "Xã Tân Khai", "Xã Tân Lợi", "Xã Tân Phú", "Xã Tân Quan", "Xã Tân Tiến", "Xã Thanh Sơn", "Xã Thiện Hưng", "Xã Thọ Sơn", "Xã Thống Nhất", "Xã Thuận Lợi", "Xã Trảng Bom", "Xã Trị An", "Xã Xuân Bắc", "Xã Xuân Định", "Xã Xuân Đông", "Xã Xuân Đường", "Xã Xuân Hòa", "Xã Xuân Lộc", "Xã Xuân Phú", "Xã Xuân Quế", "Xã Xuân Thành"],
    "Tỉnh Đồng Tháp": ["Phường An Bình", "Phường Bình Xuân", "Phường Cai Lậy", "Phường Cao Lãnh", "Phường Đạo Thạnh", "Phường Gò Công", "Phường Hồng Ngự", "Phường Long Thuận", "Phường Mỹ Ngãi", "Phường Mỹ Phong", "Phường Mỹ Phước Tây", "Phường Mỹ Tho", "Phường Mỹ Trà", "Phường Nhị Quý", "Phường Sa Đéc", "Phường Sơn Qui", "Phường Thanh Hòa", "Phường Thới Sơn", "Phường Thường Lạc", "Phường Trung An", "Xã An Hòa", "Xã An Hữu", "Xã An Long", "Xã An Phước", "Xã An Thạnh Thủy", "Xã Ba Sao", "Xã Bình Hàng Trung", "Xã Bình Ninh", "Xã Bình Phú", "Xã Bình Thành", "Xã Bình Trưng", "Xã Cái Bè", "Xã Châu Thành", "Xã Chợ Gạo", "Xã Đốc Binh Kiều", "Xã Đồng Sơn", "Xã Gia Thuận", "Xã Gò Công Đông", "Xã Hậu Mỹ", "Xã Hiệp Đức", "Xã Hòa Long", "Xã Hội Cư", "Xã Hưng Thạnh", "Xã Kim Sơn", "Xã Lai Vung", "Xã Lấp Vò", "Xã Long Bình", "Xã Long Định", "Xã Long Hưng", "Xã Long Khánh", "Xã Long Phú Thuận", "Xã Long Tiên", "Xã Lương Hòa Lạc", "Xã Mỹ An Hưng", "Xã Mỹ Đức Tây", "Xã Mỹ Hiệp", "Xã Mỹ Lợi", "Xã Mỹ Quí", "Xã Mỹ Thành", "Xã Mỹ Thiện", "Xã Mỹ Thọ", "Xã Mỹ Tịnh An", "Xã Ngũ Hiệp", "Xã Phong Hòa", "Xã Phong Mỹ", "Xã Phú Cường", "Xã Phú Hựu", "Xã Phú Thành", "Xã Phú Thọ", "Xã Phương Thịnh", "Xã Tam Nông", "Xã Tân Điền", "Xã Tân Đông", "Xã Tân Dương", "Xã Tân Hộ Cơ", "Xã Tân Hòa", "Xã Tân Hồng", "Xã Tân Hương", "Xã Tân Khánh Trung", "Xã Tân Long", "Xã Tân Nhuận Đông", "Xã Tân Phú", "Xã Tân Phú Đông", "Xã Tân Phú Trung", "Xã Tân Phước 1", "Xã Tân Phước 2", "Xã Tân Phước 3", "Xã Tân Thạnh", "Xã Tân Thành", "Xã Tân Thới", "Xã Tân Thuận Bình", "Xã Thanh Bình", "Xã Thanh Hưng", "Xã Thanh Mỹ", "Xã Thạnh Phú", "Xã Tháp Mười", "Xã Thường Phước", "Xã Tràm Chim", "Xã Trường Xuân", "Xã Vĩnh Bình", "Xã Vĩnh Hựu", "Xã Vĩnh Kim"],
    "Tỉnh Gia Lai": ["Phường An Bình", "Phường An Khê", "Phường An Nhơn", "Phường An Nhơn Bắc", "Phường An Nhơn Đông", "Phường An Nhơn Nam", "Phường An Phú", "Phường Ayun Pa", "Phường Bình Định", "Phường Bồng Sơn", "Phường Diên Hồng", "Phường Hoài Nhơn", "Phường Hoài Nhơn Bắc", "Phường Hoài Nhơn Đông", "Phường Hoài Nhơn Nam", "Phường Hoài Nhơn Tây", "Phường Hội Phú", "Phường Pleiku", "Phường Quy Nhơn", "Phường Quy Nhơn Bắc", "Phường Quy Nhơn Đông", "Phường Quy Nhơn Nam", "Phường Quy Nhơn Tây", "Phường Tam Quan", "Phường Thống Nhất", "Xã Al Bá", "Xã Ân Hảo", "Xã An Hòa", "Xã An Lão", "Xã An Lương", "Xã An Nhơn Tây", "Xã An Toàn", "Xã Ân Tường", "Xã An Vinh", "Xã Ayun", "Xã Bàu Cạn", "Xã Biển Hồ", "Xã Bình An", "Xã Bình Dương", "Xã Bình Hiệp", "Xã Bình Khê", "Xã Bình Phú", "Xã Bờ Ngoong", "Xã Canh Liên", "Xã Canh Vinh", "Xã Cát Tiến", "Xã Chơ Long", "Xã Chư A Thai", "Xã Chư Krey", "Xã Chư Păh", "Xã Chư Prông", "Xã Chư Pưh", "Xã Chư Sê", "Xã Cửu An", "Xã Đak Đoa", "Xã Đak Pơ", "Xã Đak Rong", "Xã Đak Sơmei", "Xã Đăk Song", "Xã Đề Gi", "Xã Đức Cơ", "Xã Gào", "Xã Hòa Hội", "Xã Hoài Ân", "Xã Hội Sơn", "Xã Hra", "Xã Ia Băng", "Xã Ia Boòng", "Xã Ia Chia", "Xã Ia Dơk", "Xã Ia Dom", "Xã Ia Dreh", "Xã Ia Grai", "Xã Ia Hiao", "Xã Ia Hrú", "Xã Ia Hrung", "Xã Ia Khươl", "Xã Ia Ko", "Xã Ia Krái", "Xã Ia Krêl", "Xã Ia Lâu", "Xã Ia Le", "Xã Ia Ly", "Xã Ia Mơ", "Xã Ia Nan", "Xã Ia O", "Xã Ia Pa", "Xã Ia Phí", "Xã Ia Pia", "Xã Ia Pnôn", "Xã Ia Púch", "Xã Ia Rbol", "Xã Ia Rsai", "Xã Ia Sao", "Xã Ia Tôr", "Xã Ia Tul", "Xã Kbang", "Xã KDang", "Xã Kim Sơn", "Xã Kon Chiêng", "Xã Kon Gang", "Xã Kông Bơ La", "Xã Kông Chro", "Xã Krong", "Xã Lơ Pang", "Xã Mang Yang", "Xã Ngô Mây", "Xã Nhơn Châu", "Xã Phù Cát", "Xã Phù Mỹ", "Xã Phù Mỹ Bắc", "Xã Phù Mỹ Đông", "Xã Phù Mỹ Nam", "Xã Phù Mỹ Tây", "Xã Phú Thiện", "Xã Phú Túc", "Xã Pờ Tó", "Xã Sơn Lang", "Xã SRó", "Xã Tây Sơn", "Xã Tơ Tung", "Xã Tuy Phước", "Xã Tuy Phước Bắc", "Xã Tuy Phước Đông", "Xã Tuy Phước Tây", "Xã Uar", "Xã Vân Canh", "Xã Vạn Đức", "Xã Vĩnh Quang", "Xã Vĩnh Sơn", "Xã Vĩnh Thạnh", "Xã Vĩnh Thịnh", "Xã Xuân An", "Xã Ya Hội", "Xã Ya Ma"],
    "Tỉnh Hà Tĩnh": ["Phường Bắc Hồng Lĩnh", "Phường Hà Huy Tập", "Phường Hải Ninh", "Phường Hoành Sơn", "Phường Nam Hồng Lĩnh", "Phường Sông Trí", "Phường Thành Sen", "Phường Trần Phú", "Phường Vũng Áng", "Xã Cẩm Bình", "Xã Cẩm Duệ", "Xã Cẩm Hưng", "Xã Cẩm Lạc", "Xã Cẩm Trung", "Xã Cẩm Xuyên", "Xã Can Lộc", "Xã Cổ Đạm", "Xã Đan Hải", "Xã Đông Kinh", "Xã Đồng Lộc", "Xã Đồng Tiến", "Xã Đức Đồng", "Xã Đức Minh", "Xã Đức Quang", "Xã Đức Thịnh", "Xã Đức Thọ", "Xã Gia Hanh", "Xã Hà Linh", "Xã Hồng Lộc", "Xã Hương Bình", "Xã Hương Đô", "Xã Hương Khê", "Xã Hương Phố", "Xã Hương Sơn", "Xã Hương Xuân", "Xã Kim Hoa", "Xã Kỳ Anh", "Xã Kỳ Hoa", "Xã Kỳ Khang", "Xã Kỳ Lạc", "Xã Kỳ Thượng", "Xã Kỳ Văn", "Xã Kỳ Xuân", "Xã Lộc Hà", "Xã Mai Hoa", "Xã Mai Phụ", "Xã Nghi Xuân", "Xã Phúc Trạch", "Xã Sơn Giang", "Xã Sơn Hồng", "Xã Sơn Kim 1", "Xã Sơn Kim 2", "Xã Sơn Tây", "Xã Sơn Tiến", "Xã Thạch Hà", "Xã Thạch Khê", "Xã Thạch Lạc", "Xã Thạch Xuân", "Xã Thiên Cầm", "Xã Thượng Đức", "Xã Tiên Điền", "Xã Toàn Lưu", "Xã Trường Lưu", "Xã Tứ Mỹ", "Xã Tùng Lộc", "Xã Việt Xuyên", "Xã Vũ Quang", "Xã Xuân Lộc", "Xã Yên Hòa"],
    "Tỉnh Hưng Yên": ["Phường Đường Hào", "Phường Hồng Châu", "Phường Mỹ Hào", "Phường Phố Hiến", "Phường Sơn Nam", "Phường Thái Bình", "Phường Thượng Hồng", "Phường Trà Lý", "Phường Trần Hưng Đạo", "Phường Trần Lãm", "Phường Vũ Phúc", "Xã A Sào", "Xã Ái Quốc", "Xã Ân Thi", "Xã Bắc Đông Hưng", "Xã Bắc Đông Quan", "Xã Bắc Thái Ninh", "Xã Bắc Thụy Anh", "Xã Bắc Tiên Hưng", "Xã Bình Định", "Xã Bình Nguyên", "Xã Bình Thanh", "Xã Châu Ninh", "Xã Chí Minh", "Xã Đại Đồng", "Xã Diên Hà", "Xã Đoàn Đào", "Xã Đồng Bằng", "Xã Đồng Châu", "Xã Đông Hưng", "Xã Đông Quan", "Xã Đông Thái Ninh", "Xã Đông Thụy Anh", "Xã Đông Tiền Hải", "Xã Đông Tiên Hưng", "Xã Đức Hợp", "Xã Hiệp Cường", "Xã Hoàn Long", "Xã Hoàng Hoa Thám", "Xã Hồng Minh", "Xã Hồng Quang", "Xã Hồng Vũ", "Xã Hưng Hà", "Xã Hưng Phú", "Xã Khoái Châu", "Xã Kiến Xương", "Xã Lạc Đạo", "Xã Lê Lợi", "Xã Lê Quý Đôn", "Xã Long Hưng", "Xã Lương Bằng", "Xã Mễ Sở", "Xã Minh Thọ", "Xã Nam Cường", "Xã Nam Đông Hưng", "Xã Nam Thái Ninh", "Xã Nam Thụy Anh", "Xã Nam Tiền Hải", "Xã Nam Tiên Hưng", "Xã Nghĩa Dân", "Xã Nghĩa Trụ", "Xã Ngọc Lâm", "Xã Ngự Thiên", "Xã Nguyễn Du", "Xã Nguyễn Trãi", "Xã Nguyễn Văn Linh", "Xã Như Quỳnh", "Xã Phạm Ngũ Lão", "Xã Phụ Dực", "Xã Phụng Công", "Xã Quang Hưng", "Xã Quang Lịch", "Xã Quỳnh An", "Xã Quỳnh Phụ", "Xã Tân Hưng", "Xã Tân Thuận", "Xã Tân Tiến", "Xã Tây Thái Ninh", "Xã Tây Thụy Anh", "Xã Tây Tiền Hải", "Xã Thái Ninh", "Xã Thái Thụy", "Xã Thần Khê", "Xã Thư Trì", "Xã Thư Vũ", "Xã Thụy Anh", "Xã Tiền Hải", "Xã Tiên Hoa", "Xã Tiên Hưng", "Xã Tiên La", "Xã Tiên Lữ", "Xã Tiên Tiến", "Xã Tống Trân", "Xã Trà Giang", "Xã Triệu Việt Vương", "Xã Văn Giang", "Xã Vạn Xuân", "Xã Việt Tiến", "Xã Việt Yên", "Xã Vũ Quý", "Xã Vũ Thư", "Xã Vũ Tiên", "Xã Xuân Trúc", "Xã Yên Mỹ"],
    "Tỉnh Khánh Hòa": ["Đặc khu Trường Sa", "Phường Ba Ngòi", "Phường Bắc Cam Ranh", "Phường Bắc Nha Trang", "Phường Bảo An", "Phường Cam Linh", "Phường Cam Ranh", "Phường Đô Vinh", "Phường Đông Hải", "Phường Đông Ninh Hòa", "Phường Hòa Thắng", "Phường Nam Nha Trang", "Phường Nha Trang", "Phường Ninh Chử", "Phường Ninh Hòa", "Phường Phan Rang", "Phường Tây Nha Trang", "Xã Anh Dũng", "Xã Bác Ái", "Xã Bác Ái Đông", "Xã Bác Ái Tây", "Xã Bắc Khánh Vĩnh", "Xã Bắc Ninh Hòa", "Xã Cà Ná", "Xã Cam An", "Xã Cam Hiệp", "Xã Cam Lâm", "Xã Công Hải", "Xã Đại Lãnh", "Xã Diên Điền", "Xã Diên Khánh", "Xã Diên Lạc", "Xã Diên Lâm", "Xã Diên Thọ", "Xã Đông Khánh Sơn", "Xã Hòa Trí", "Xã Khánh Sơn", "Xã Khánh Vĩnh", "Xã Lâm Sơn", "Xã Mỹ Sơn", "Xã Nam Cam Ranh", "Xã Nam Khánh Vĩnh", "Xã Nam Ninh Hòa", "Xã Ninh Hải", "Xã Ninh Phước", "Xã Ninh Sơn", "Xã Phước Dinh", "Xã Phước Hà", "Xã Phước Hậu", "Xã Phước Hữu", "Xã Suối Dầu", "Xã Suối Hiệp", "Xã Tân Định", "Xã Tây Khánh Sơn", "Xã Tây Khánh Vĩnh", "Xã Tây Ninh Hòa", "Xã Thuận Bắc", "Xã Thuận Nam", "Xã Trung Khánh Vĩnh", "Xã Tu Bông", "Xã Vạn Hưng", "Xã Vạn Ninh", "Xã Vạn Thắng", "Xã Vĩnh Hải", "Xã Xuân Hải"],
    "Tỉnh Lai Châu": ["Phường Đoàn Kết", "Phường Tân Phong", "Xã Bản Bo", "Xã Bình Lư", "Xã Bum Nưa", "Xã Bum Tở", "Xã Dào San", "Xã Hồng Thu", "Xã Hua Bum", "Xã Khoen On", "Xã Khổng Lào", "Xã Khun Há", "Xã Lê Lợi", "Xã Mù Cả", "Xã Mường Khoa", "Xã Mường Kim", "Xã Mường Mô", "Xã Mường Tè", "Xã Mường Than", "Xã Nậm Cuổi", "Xã Nậm Hàng", "Xã Nậm Mạ", "Xã Nậm Sỏ", "Xã Nậm Tăm", "Xã Pa Tần", "Xã Pa Ủ", "Xã Pắc Ta", "Xã Phong Thổ", "Xã Pu Sam Cáp", "Xã Sì Lở Lầu", "Xã Sìn Hồ", "Xã Sin Suối Hồ", "Xã Tả Lèng", "Xã Tà Tổng", "Xã Tân Uyên", "Xã Than Uyên", "Xã Thu Lũm", "Xã Tủa Sín Chải"],
    "Tỉnh Lâm Đồng": ["Đặc khu Phú Quý", "Phường 1 Bảo Lộc", "Phường 2 Bảo Lộc", "Phường 3 Bảo Lộc", "Phường B'Lao", "Phường Bắc Gia Nghĩa", "Phường Bình Thuận", "Phường Cam Ly - Đà Lạt", "Phường Đông Gia Nghĩa", "Phường Hàm Thắng", "Phường La Gi", "Phường Lâm Viên - Đà Lạt", "Phường Lang Biang - Đà Lạt", "Phường Mũi Né", "Phường Nam Gia Nghĩa", "Phường Phan Thiết", "Phường Phú Thủy", "Phường Phước Hội", "Phường Tiến Thành", "Phường Xuân Hương - Đà Lạt", "Phường Xuân Trường - Đà Lạt", "Xã Bắc Bình", "Xã Bắc Ruộng", "Xã Bảo Lâm 1", "Xã Bảo Lâm 2", "Xã Bảo Lâm 3", "Xã Bảo Lâm 4", "Xã Bảo Lâm 5", "Xã Bảo Thuận", "Xã Cát Tiên", "Xã Cát Tiên 2", "Xã Cát Tiên 3", "Xã Cư Jút", "Xã D'Ran", "Xã Đạ Huoai", "Xã Đạ Huoai 2", "Xã Đạ Huoai 3", "Xã Đạ Tẻh", "Xã Đạ Tẻh 2", "Xã Đạ Tẻh 3", "Xã Đắk Mil", "Xã Đắk Sắk", "Xã Đắk Song", "Xã Đắk Wil", "Xã Đam Rông 1", "Xã Đam Rông 2", "Xã Đam Rông 3", "Xã Đam Rông 4", "Xã Di Linh", "Xã Đinh Trang Thượng", "Xã Đinh Văn Lâm Hà", "Xã Đơn Dương", "Xã Đông Giang", "Xã Đồng Kho", "Xã Đức An", "Xã Đức Lập", "Xã Đức Linh", "Xã Đức Trọng", "Xã Gia Hiệp", "Xã Hải Ninh", "Xã Hàm Kiệm", "Xã Hàm Liêm", "Xã Hàm Tân", "Xã Hàm Thạnh", "Xã Hàm Thuận", "Xã Hàm Thuận Bắc", "Xã Hàm Thuận Nam", "Xã Hiệp Thạnh", "Xã Hòa Bắc", "Xã Hòa Ninh", "Xã Hòa Thắng", "Xã Hoài Đức", "Xã Hồng Sơn", "Xã Hồng Thái", "Xã Ka Đô", "Xã Kiến Đức", "Xã Krông Nô", "Xã La Dạ", "Xã Lạc Dương", "Xã Liên Hương", "Xã Lương Sơn", "Xã Nam Ban Lâm Hà", "Xã Nam Đà", "Xã Nam Dong", "Xã Nam Hà Lâm Hà", "Xã Nâm Nung", "Xã Nam Thành", "Xã Nghị Đức", "Xã Nhân Cơ", "Xã Ninh Gia", "Xã Phan Rí Cửa", "Xã Phan Sơn", "Xã Phú Sơn Lâm Hà", "Xã Phúc Thọ Lâm Hà", "Xã Quảng Hòa", "Xã Quảng Khê", "Xã Quảng Lập", "Xã Quảng Phú", "Xã Quảng Sơn", "Xã Quảng Tân", "Xã Quảng Tín", "Xã Quảng Trực", "Xã Sơn Điền", "Xã Sơn Mỹ", "Xã Sông Lũy", "Xã Suối Kiết", "Xã Tà Đùng", "Xã Tà Hine", "Xã Tà Năng", "Xã Tân Hà Lâm Hà", "Xã Tân Hải", "Xã Tân Hội", "Xã Tân Lập", "Xã Tân Minh", "Xã Tân Thành", "Xã Tánh Linh", "Xã Thuận An", "Xã Thuận Hạnh", "Xã Trà Tân", "Xã Trường Xuân", "Xã Tuy Đức", "Xã Tuy Phong", "Xã Tuyên Quang", "Xã Vĩnh Hảo"],
    "Tỉnh Lạng Sơn": ["Phường Đông Kinh", "Phường Kỳ Lừa", "Phường Lương Văn Tri", "Phường Tam Thanh", "Xã Ba Sơn", "Xã Bắc Sơn", "Xã Bằng Mạc", "Xã Bình Gia", "Xã Cai Kinh", "Xã Cao Lộc", "Xã Châu Sơn", "Xã Chi Lăng", "Xã Chiến Thắng", "Xã Công Sơn", "Xã Điềm He", "Xã Đình Lập", "Xã Đoàn Kết", "Xã Đồng Đăng", "Xã Hoa Thám", "Xã Hoàng Văn Thụ", "Xã Hội Hoan", "Xã Hồng Phong", "Xã Hưng Vũ", "Xã Hữu Liên", "Xã Hữu Lũng", "Xã Kháng Chiến", "Xã Khánh Khê", "Xã Khuất Xá", "Xã Kiên Mộc", "Xã Lộc Bình", "Xã Lợi Bác", "Xã Mẫu Sơn", "Xã Na Dương", "Xã Na Sầm", "Xã Nhân Lý", "Xã Nhất Hòa", "Xã Quan Sơn", "Xã Quốc Khánh", "Xã Quốc Việt", "Xã Quý Hòa", "Xã Tân Đoàn", "Xã Tân Thành", "Xã Tân Tiến", "Xã Tân Tri", "Xã Tân Văn", "Xã Thái Bình", "Xã Thất Khê", "Xã Thiện Hòa", "Xã Thiện Long", "Xã Thiện Tân", "Xã Thiện Thuật", "Xã Thống Nhất", "Xã Thụy Hùng", "Xã Tràng Định", "Xã Tri Lễ", "Xã Tuấn Sơn", "Xã Văn Lãng", "Xã Vạn Linh", "Xã Vân Nham", "Xã Văn Quan", "Xã Vũ Lăng", "Xã Vũ Lễ", "Xã Xuân Dương", "Xã Yên Bình", "Xã Yên Phúc"],
    "Tỉnh Lào Cai": ["Phường Âu Lâu", "Phường Cam Đường", "Phường Cầu Thia", "Phường Lào Cai", "Phường Nam Cường", "Phường Nghĩa Lộ", "Phường Sa Pa", "Phường Trung Tâm", "Phường Văn Phú", "Phường Yên Bái", "Xã A Mú Sung", "Xã Bắc Hà", "Xã Bản Hồ", "Xã Bản Lầu", "Xã Bản Liền", "Xã Bản Xèo", "Xã Bảo Ái", "Xã Bảo Hà", "Xã Bảo Nhai", "Xã Bảo Thắng", "Xã Bảo Yên", "Xã Bát Xát", "Xã Cảm Nhân", "Xã Cao Sơn", "Xã Cát Thịnh", "Xã Chấn Thịnh", "Xã Châu Quế", "Xã Chế Tạo", "Xã Chiềng Ken", "Xã Cốc Lầu", "Xã Cốc San", "Xã Dền Sáng", "Xã Đông Cuông", "Xã Dương Quỳ", "Xã Gia Hội", "Xã Gia Phú", "Xã Hạnh Phúc", "Xã Hợp Thành", "Xã Hưng Khánh", "Xã Khánh Hòa", "Xã Khánh Yên", "Xã Khao Mang", "Xã Lâm Giang", "Xã Lâm Thượng", "Xã Lao Chải", "Xã Liên Sơn", "Xã Lục Yên", "Xã Lùng Phình", "Xã Lương Thịnh", "Xã Mậu A", "Xã Minh Lương", "Xã Mỏ Vàng", "Xã Mù Cang Chải", "Xã Mường Bo", "Xã Mường Hum", "Xã Mường Khương", "Xã Mường Lai", "Xã Nậm Chày", "Xã Nậm Có", "Xã Nậm Xé", "Xã Nghĩa Đô", "Xã Nghĩa Tâm", "Xã Ngũ Chỉ Sơn", "Xã Pha Long", "Xã Phình Hồ", "Xã Phong Dụ Hạ", "Xã Phong Dụ Thượng", "Xã Phong Hải", "Xã Phúc Khánh", "Xã Phúc Lợi", "Xã Púng Luông", "Xã Quy Mông", "Xã Si Ma Cai", "Xã Sín Chéng", "Xã Sơn Lương", "Xã Tả Củ Tỷ", "Xã Tả Phìn", "Xã Tả Van", "Xã Tà Xi Láng", "Xã Tân Hợp", "Xã Tân Lĩnh", "Xã Tằng Loỏng", "Xã Thác Bà", "Xã Thượng Bằng La", "Xã Thượng Hà", "Xã Trạm Tấu", "Xã Trấn Yên", "Xã Trịnh Tường", "Xã Tú Lệ", "Xã Văn Bàn", "Xã Văn Chấn", "Xã Việt Hồng", "Xã Võ Lao", "Xã Xuân Ái", "Xã Xuân Hòa", "Xã Xuân Quang", "Xã Y Tý", "Xã Yên Bình", "Xã Yên Thành"],
    "Tỉnh Nghệ An": ["Phường Cửa Lò", "Phường Hoàng Mai", "Phường Quỳnh Mai", "Phường Tân Mai", "Phường Tây Hiếu", "Phường Thái Hòa", "Phường Thành Vinh", "Phường Trường Vinh", "Phường Vinh Hưng", "Phường Vinh Lộc", "Phường Vinh Phú", "Xã An Châu", "Xã Anh Sơn", "Xã Anh Sơn Đông", "Xã Bắc Lý", "Xã Bạch Hà", "Xã Bạch Ngọc", "Xã Bích Hào", "Xã Bình Chuẩn", "Xã Bình Minh", "Xã Cam Phục", "Xã Cát Ngạn", "Xã Châu Bình", "Xã Châu Hồng", "Xã Châu Khê", "Xã Châu Lộc", "Xã Châu Tiến", "Xã Chiêu Lưu", "Xã Con Cuông", "Xã Đại Đồng", "Xã Đại Huệ", "Xã Diễn Châu", "Xã Đô Lương", "Xã Đông Hiếu", "Xã Đông Lộc", "Xã Đông Thành", "Xã Đức Châu", "Xã Giai Lạc", "Xã Giai Xuân", "Xã Hải Châu", "Xã Hải Lộc", "Xã Hạnh Lâm", "Xã Hoa Quân", "Xã Hợp Minh", "Xã Hùng Chân", "Xã Hùng Châu", "Xã Hưng Nguyên", "Xã Hưng Nguyên Nam", "Xã Huồi Tụ", "Xã Hữu Khuông", "Xã Hữu Kiệm", "Xã Keng Đu", "Xã Kim Bảng", "Xã Kim Liên", "Xã Lam Thành", "Xã Lượng Minh", "Xã Lương Sơn", "Xã Mậu Thạch", "Xã Minh Châu", "Xã Minh Hợp", "Xã Môn Sơn", "Xã Mường Chọng", "Xã Mường Ham", "Xã Mường Lống", "Xã Mường Quàng", "Xã Mường Típ", "Xã Mường Xén", "Xã Mỹ Lý", "Xã Na Loi", "Xã Na Ngoi", "Xã Nậm Cắn", "Xã Nam Đàn", "Xã Nga My", "Xã Nghi Lộc", "Xã Nghĩa Đàn", "Xã Nghĩa Đồng", "Xã Nghĩa Hành", "Xã Nghĩa Hưng", "Xã Nghĩa Khánh", "Xã Nghĩa Lâm", "Xã Nghĩa Lộc", "Xã Nghĩa Mai", "Xã Nghĩa Thọ", "Xã Nhân Hòa", "Xã Nhôn Mai", "Xã Phúc Lộc", "Xã Quan Thành", "Xã Quảng Châu", "Xã Quang Đồng", "Xã Quế Phong", "Xã Quỳ Châu", "Xã Quỳ Hợp", "Xã Quỳnh Anh", "Xã Quỳnh Lưu", "Xã Quỳnh Phú", "Xã Quỳnh Sơn", "Xã Quỳnh Tam", "Xã Quỳnh Thắng", "Xã Quỳnh Văn", "Xã Sơn Lâm", "Xã Tam Đồng", "Xã Tam Hợp", "Xã Tam Quang", "Xã Tam Thái", "Xã Tân An", "Xã Tân Châu", "Xã Tân Kỳ", "Xã Tân Phú", "Xã Thần Lĩnh", "Xã Thành Bình Thọ", "Xã Thiên Nhẫn", "Xã Thông Thụ", "Xã Thuần Trung", "Xã Tiên Đồng", "Xã Tiền Phong", "Xã Tri Lễ", "Xã Trung Lộc", "Xã Tương Dương", "Xã Vạn An", "Xã Vân Du", "Xã Văn Hiến", "Xã Văn Kiều", "Xã Vân Tụ", "Xã Vĩnh Tường", "Xã Xuân Lâm", "Xã Yên Hòa", "Xã Yên Na", "Xã Yên Thành", "Xã Yên Trung", "Xã Yên Xuân"],
    "Tỉnh Ninh Bình": ["Phường Châu Sơn", "Phường Đông A", "Phường Đông Hoa Lư", "Phường Đồng Văn", "Phường Duy Hà", "Phường Duy Tân", "Phường Duy Tiên", "Phường Hà Nam", "Phường Hoa Lư", "Phường Hồng Quang", "Phường Kim Bảng", "Phường Kim Thanh", "Phường Lê Hồ", "Phường Liêm Tuyền", "Phường Lý Thường Kiệt", "Phường Mỹ Lộc", "Phường Nam Định", "Phường Nam Hoa Lư", "Phường Nguyễn Úy", "Phường Phủ Lý", "Phường Phù Vân", "Phường Tam Chúc", "Phường Tam Điệp", "Phường Tây Hoa Lư", "Phường Thành Nam", "Phường Thiên Trường", "Phường Tiên Sơn", "Phường Trung Sơn", "Phường Trường Thi", "Phường Vị Khê", "Phường Yên Sơn", "Phường Yên Thắng", "Xã Bắc Lý", "Xã Bình An", "Xã Bình Giang", "Xã Bình Lục", "Xã Bình Minh", "Xã Bình Mỹ", "Xã Bình Sơn", "Xã Cát Thành", "Xã Chất Bình", "Xã Cổ Lễ", "Xã Cúc Phương", "Xã Đại Hoàng", "Xã Định Hóa", "Xã Đồng Thái", "Xã Đồng Thịnh", "Xã Gia Hưng", "Xã Gia Lâm", "Xã Gia Phong", "Xã Gia Trấn", "Xã Gia Tường", "Xã Gia Vân", "Xã Gia Viễn", "Xã Giao Bình", "Xã Giao Hòa", "Xã Giao Hưng", "Xã Giao Minh", "Xã Giao Ninh", "Xã Giao Phúc", "Xã Giao Thủy", "Xã Hải An", "Xã Hải Anh", "Xã Hải Hậu", "Xã Hải Hưng", "Xã Hải Quang", "Xã Hải Thịnh", "Xã Hải Tiến", "Xã Hải Xuân", "Xã Hiển Khánh", "Xã Hồng Phong", "Xã Khánh Hội", "Xã Khánh Nhạc", "Xã Khánh Thiện", "Xã Khánh Trung", "Xã Kim Đông", "Xã Kim Sơn", "Xã Lai Thành", "Xã Liêm Hà", "Xã Liên Minh", "Xã Lý Nhân", "Xã Minh Tân", "Xã Minh Thái", "Xã Nam Đồng", "Xã Nam Hồng", "Xã Nam Lý", "Xã Nam Minh", "Xã Nam Ninh", "Xã Nam Trực", "Xã Nam Xang", "Xã Nghĩa Hưng", "Xã Nghĩa Lâm", "Xã Nghĩa Sơn", "Xã Nhân Hà", "Xã Nho Quan", "Xã Ninh Cường", "Xã Ninh Giang", "Xã Phát Diệm", "Xã Phong Doanh", "Xã Phú Long", "Xã Phú Sơn", "Xã Quang Hưng", "Xã Quang Thiện", "Xã Quỹ Nhất", "Xã Quỳnh Lưu", "Xã Rạng Đông", "Xã Tân Minh", "Xã Tân Thanh", "Xã Thanh Bình", "Xã Thanh Lâm", "Xã Thanh Liêm", "Xã Thanh Sơn", "Xã Trần Thương", "Xã Trực Ninh", "Xã Vạn Thắng", "Xã Vĩnh Trụ", "Xã Vụ Bản", "Xã Vũ Dương", "Xã Xuân Giang", "Xã Xuân Hồng", "Xã Xuân Hưng", "Xã Xuân Trường", "Xã Ý Yên", "Xã Yên Cường", "Xã Yên Đồng", "Xã Yên Khánh", "Xã Yên Mạc", "Xã Yên Mô", "Xã Yên Từ"],
    "Tỉnh Phú Thọ": ["Phường Âu Cơ", "Phường Hòa Bình", "Phường Kỳ Sơn", "Phường Nông Trang", "Phường Phong Châu", "Phường Phú Thọ", "Phường Phúc Yên", "Phường Tân Hòa", "Phường Thanh Miếu", "Phường Thống Nhất", "Phường Vân Phú", "Phường Việt Trì", "Phường Vĩnh Phúc", "Phường Vĩnh Yên", "Phường Xuân Hòa", "Xã An Bình", "Xã An Nghĩa", "Xã Bản Nguyên", "Xã Bằng Luân", "Xã Bao La", "Xã Bình Nguyên", "Xã Bình Phú", "Xã Bình Tuyền", "Xã Bình Xuyên", "Xã Cẩm Khê", "Xã Cao Dương", "Xã Cao Phong", "Xã Cao Sơn", "Xã Chân Mộng", "Xã Chí Đám", "Xã Chí Tiên", "Xã Cự Đồng", "Xã Đà Bắc", "Xã Đại Đình", "Xã Đại Đồng", "Xã Dân Chủ", "Xã Đan Thượng", "Xã Đạo Trù", "Xã Đào Xá", "Xã Đoan Hùng", "Xã Đồng Lương", "Xã Đông Thành", "Xã Đức Nhàn", "Xã Dũng Tiến", "Xã Hạ Hòa", "Xã Hải Lựu", "Xã Hiền Lương", "Xã Hiền Quan", "Xã Hoàng An", "Xã Hoàng Cương", "Xã Hội Thịnh", "Xã Hợp Kim", "Xã Hợp Lý", "Xã Hùng Việt", "Xã Hương Cần", "Xã Hy Cương", "Xã Khả Cửu", "Xã Kim Bôi", "Xã Lạc Lương", "Xã Lạc Sơn", "Xã Lạc Thủy", "Xã Lai Đồng", "Xã Lâm Thao", "Xã Lập Thạch", "Xã Liên Châu", "Xã Liên Hòa", "Xã Liên Minh", "Xã Liên Sơn", "Xã Long Cốc", "Xã Lương Sơn", "Xã Mai Châu", "Xã Mai Hạ", "Xã Minh Đài", "Xã Minh Hòa", "Xã Mường Bi", "Xã Mường Động", "Xã Mường Hoa", "Xã Mường Thàng", "Xã Mường Vang", "Xã Nật Sơn", "Xã Ngọc Sơn", "Xã Nguyệt Đức", "Xã Nhân Nghĩa", "Xã Pà Cò", "Xã Phú Khê", "Xã Phú Mỹ", "Xã Phù Ninh", "Xã Phùng Nguyên", "Xã Quảng Yên", "Xã Quy Đức", "Xã Quyết Thắng", "Xã Sơn Đông", "Xã Sơn Lương", "Xã Sông Lô", "Xã Tam Đảo", "Xã Tam Dương", "Xã Tam Dương Bắc", "Xã Tam Hồng", "Xã Tam Nông", "Xã Tam Sơn", "Xã Tân Lạc", "Xã Tân Mai", "Xã Tân Pheo", "Xã Tân Sơn", "Xã Tây Cốc", "Xã Tề Lỗ", "Xã Thái Hòa", "Xã Thanh Ba", "Xã Thanh Sơn", "Xã Thanh Thủy", "Xã Thịnh Minh", "Xã Thổ Tang", "Xã Thọ Văn", "Xã Thu Cúc", "Xã Thung Nai", "Xã Thượng Cốc", "Xã Thượng Long", "Xã Tiên Lữ", "Xã Tiên Lương", "Xã Tiền Phong", "Xã Toàn Thắng", "Xã Trạm Thản", "Xã Trung Sơn", "Xã Tu Vũ", "Xã Vân Bán", "Xã Văn Lang", "Xã Văn Miếu", "Xã Vân Sơn", "Xã Vạn Xuân", "Xã Vĩnh An", "Xã Vĩnh Chân", "Xã Vĩnh Hưng", "Xã Vĩnh Phú", "Xã Vĩnh Thành", "Xã Vĩnh Tường", "Xã Võ Miếu", "Xã Xuân Đài", "Xã Xuân Lãng", "Xã Xuân Lũng", "Xã Xuân Viên", "Xã Yên Kỳ", "Xã Yên Lạc", "Xã Yên Lãng", "Xã Yên Lập", "Xã Yên Phú", "Xã Yên Sơn", "Xã Yên Thủy", "Xã Yên Trị"],
    "Tỉnh Quảng Ngãi": ["Đặc khu Lý Sơn", "Phường Cẩm Thành", "Phường Đăk Bla", "Phường Đăk Cấm", "Phường Đức Phổ", "Phường Kon Tum", "Phường Nghĩa Lộ", "Phường Sa Huỳnh", "Phường Trà Câu", "Phường Trương Quang Trọng", "Xã An Phú", "Xã Ba Dinh", "Xã Ba Động", "Xã Ba Gia", "Xã Ba Tô", "Xã Ba Tơ", "Xã Ba Vì", "Xã Ba Vinh", "Xã Ba Xa", "Xã Bình Chương", "Xã Bình Minh", "Xã Bình Sơn", "Xã Bờ Y", "Xã Cà Đam", "Xã Đăk Hà", "Xã Đăk Kôi", "Xã Đăk Long", "Xã Đăk Mar", "Xã Đăk Môn", "Xã Đăk Pék", "Xã Đăk Plô", "Xã Đăk Pxi", "Xã Đăk Rơ Wa", "Xã Đăk Rve", "Xã Đăk Sao", "Xã Đăk Tô", "Xã Đăk Tờ Kan", "Xã Đăk Ui", "Xã Đặng Thùy Trâm", "Xã Đình Cương", "Xã Đông Sơn", "Xã Đông Trà Bồng", "Xã Dục Nông", "Xã Ia Chim", "Xã Ia Đal", "Xã Ia Tơi", "Xã Khánh Cường", "Xã Kon Braih", "Xã Kon Đào", "Xã Kon Plông", "Xã Lân Phong", "Xã Long Phụng", "Xã Măng Bút", "Xã Măng Đen", "Xã Măng Ri", "Xã Minh Long", "Xã Mỏ Cày", "Xã Mộ Đức", "Xã Mô Rai", "Xã Nghĩa Giang", "Xã Nghĩa Hành", "Xã Ngọc Linh", "Xã Ngọk Bay", "Xã Ngọk Réo", "Xã Ngọk Tụ", "Xã Nguyễn Nghiêm", "Xã Phước Giang", "Xã Rờ Kơi", "Xã Sa Bình", "Xã Sa Loong", "Xã Sa Thầy", "Xã Sơn Hạ", "Xã Sơn Hà", "Xã Sơn Kỳ", "Xã Sơn Linh", "Xã Sơn Mai", "Xã Sơn Tây", "Xã Sơn Tây Hạ", "Xã Sơn Tây Thượng", "Xã Sơn Thủy", "Xã Sơn Tịnh", "Xã Tây Trà", "Xã Tây Trà Bồng", "Xã Thanh Bồng", "Xã Thiện Tín", "Xã Thọ Phong", "Xã Tịnh Khê", "Xã Trà Bồng", "Xã Trà Giang", "Xã Trường Giang", "Xã Tu Mơ Rông", "Xã Tư Nghĩa", "Xã Vạn Tường", "Xã Vệ Giang", "Xã Xốp", "Xã Ya Ly"],
    "Tỉnh Quảng Ninh": ["Đặc khu Cô Tô", "Đặc khu Vân Đồn", "Phường An Sinh", "Phường Bãi Cháy", "Phường Bình Khê", "Phường Cẩm Phả", "Phường Cao Xanh", "Phường Cửa Ông", "Phường Đông Mai", "Phường Đông Triều", "Phường Hà An", "Phường Hà Lầm", "Phường Hạ Long", "Phường Hà Tu", "Phường Hiệp Hòa", "Phường Hoàng Quế", "Phường Hoành Bồ", "Phường Hồng Gai", "Phường Liên Hòa", "Phường Mạo Khê", "Phường Móng Cái 1", "Phường Móng Cái 2", "Phường Móng Cái 3", "Phường Mông Dương", "Phường Phong Cốc", "Phường Quang Hanh", "Phường Quảng Yên", "Phường Tuần Châu", "Phường Uông Bí", "Phường Vàng Danh", "Phường Việt Hưng", "Phường Yên Tử", "Xã Ba Chẽ", "Xã Bình Liêu", "Xã Cái Chiên", "Xã Đầm Hà", "Xã Điền Xá", "Xã Đông Ngũ", "Xã Đường Hoa", "Xã Hải Hòa", "Xã Hải Lạng", "Xã Hải Ninh", "Xã Hải Sơn", "Xã Hoành Mô", "Xã Kỳ Thượng", "Xã Lục Hồn", "Xã Lương Minh", "Xã Quảng Đức", "Xã Quảng Hà", "Xã Quảng La", "Xã Quảng Tân", "Xã Thống Nhất", "Xã Tiên Yên", "Xã Vĩnh Thực"],
    "Tỉnh Quảng Trị": ["Đặc khu Cồn Cỏ", "Phường Ba Đồn", "Phường Bắc Gianh", "Phường Đông Hà", "Phường Đồng Hới", "Phường Đồng Sơn", "Phường Đồng Thuận", "Phường Nam Đông Hà", "Phường Quảng Trị", "Xã A Dơi", "Xã Ái Tử", "Xã Ba Lòng", "Xã Bắc Trạch", "Xã Bến Hải", "Xã Bến Quan", "Xã Bố Trạch", "Xã Cam Hồng", "Xã Cam Lộ", "Xã Cồn Tiên", "Xã Cửa Tùng", "Xã Cửa Việt", "Xã Đakrông", "Xã Dân Hóa", "Xã Diên Sanh", "Xã Đồng Lê", "Xã Đông Trạch", "Xã Gio Linh", "Xã Hải Lăng", "Xã Hiếu Giang", "Xã Hòa Trạch", "Xã Hoàn Lão", "Xã Hướng Hiệp", "Xã Hướng Lập", "Xã Hướng Phùng", "Xã Khe Sanh", "Xã Kim Điền", "Xã Kim Ngân", "Xã Kim Phú", "Xã La Lay", "Xã Lao Bảo", "Xã Lệ Ninh", "Xã Lệ Thủy", "Xã Lìa", "Xã Minh Hóa", "Xã Mỹ Thủy", "Xã Nam Ba Đồn", "Xã Nam Cửa Việt", "Xã Nam Gianh", "Xã Nam Hải Lăng", "Xã Nam Trạch", "Xã Ninh Châu", "Xã Phong Nha", "Xã Phú Trạch", "Xã Quảng Ninh", "Xã Quảng Trạch", "Xã Sen Ngư", "Xã Tà Rụt", "Xã Tân Gianh", "Xã Tân Lập", "Xã Tân Mỹ", "Xã Tân Thành", "Xã Thượng Trạch", "Xã Triệu Bình", "Xã Triệu Cơ", "Xã Triệu Phong", "Xã Trung Thuần", "Xã Trường Ninh", "Xã Trường Phú", "Xã Trường Sơn", "Xã Tuyên Bình", "Xã Tuyên Hóa", "Xã Tuyên Lâm", "Xã Tuyên Phú", "Xã Tuyên Sơn", "Xã Vĩnh Định", "Xã Vĩnh Hoàng", "Xã Vĩnh Linh", "Xã Vĩnh Thủy"],
    "Tỉnh Sơn La": ["Phường Chiềng An", "Phường Chiềng Cơi", "Phường Chiềng Sinh", "Phường Mộc Châu", "Phường Mộc Sơn", "Phường Thảo Nguyên", "Phường Tô Hiệu", "Phường Vân Sơn", "Xã Bắc Yên", "Xã Bình Thuận", "Xã Bó Sinh", "Xã Chiềng Hặc", "Xã Chiềng Hoa", "Xã Chiềng Khoong", "Xã Chiềng Khương", "Xã Chiềng La", "Xã Chiềng Lao", "Xã Chiềng Mai", "Xã Chiềng Mung", "Xã Chiềng Sại", "Xã Chiềng Sơ", "Xã Chiềng Sơn", "Xã Chiềng Sung", "Xã Co Mạ", "Xã Đoàn Kết", "Xã Gia Phù", "Xã Huổi Một", "Xã Kim Bon", "Xã Long Hẹ", "Xã Lóng Phiêng", "Xã Lóng Sập", "Xã Mai Sơn", "Xã Muổi Nọi", "Xã Mường Bám", "Xã Mường Bang", "Xã Mường Bú", "Xã Mường Chanh", "Xã Mường Chiên", "Xã Mường Cơi", "Xã Mường É", "Xã Mường Giôn", "Xã Mường Hung", "Xã Mường Khiêng", "Xã Mường La", "Xã Mường Lầm", "Xã Mường Lạn", "Xã Mường Lèo", "Xã Mường Sại", "Xã Nậm Lầu", "Xã Nậm Ty", "Xã Ngọc Chiến", "Xã Pắc Ngà", "Xã Phiêng Cằm", "Xã Phiêng Khoài", "Xã Phiêng Pằn", "Xã Phù Yên", "Xã Púng Bánh", "Xã Quỳnh Nhai", "Xã Song Khủa", "Xã Sông Mã", "Xã Sốp Cộp", "Xã Suối Tọ", "Xã Tà Hộc", "Xã Tạ Khoa", "Xã Tà Xùa", "Xã Tân Phong", "Xã Tân Yên", "Xã Thuận Châu", "Xã Tô Múa", "Xã Tường Hạ", "Xã Vân Hồ", "Xã Xím Vàng", "Xã Xuân Nha", "Xã Yên Châu", "Xã Yên Sơn"],
    "Tỉnh Tây Ninh": ["Phường An Tịnh", "Phường Bình Minh", "Phường Gia Lộc", "Phường Gò Dầu", "Phường Hòa Thành", "Phường Khánh Hậu", "Phường Kiến Tường", "Phường Long An", "Phường Long Hoa", "Phường Ninh Thạnh", "Phường Tân An", "Phường Tân Ninh", "Phường Thanh Điền", "Phường Trảng Bàng", "Xã An Lục Long", "Xã An Ninh", "Xã Bến Cầu", "Xã Bến Lức", "Xã Bình Đức", "Xã Bình Hiệp", "Xã Bình Hòa", "Xã Bình Thành", "Xã Cần Đước", "Xã Cần Giuộc", "Xã Cầu Khởi", "Xã Châu Thành", "Xã Đông Thành", "Xã Đức Hòa", "Xã Đức Huệ", "Xã Đức Lập", "Xã Dương Minh Châu", "Xã Hảo Đước", "Xã Hậu Nghĩa", "Xã Hậu Thạnh", "Xã Hiệp Hòa", "Xã Hòa Hội", "Xã Hòa Khánh", "Xã Hưng Điền", "Xã Hưng Thuận", "Xã Khánh Hưng", "Xã Lộc Ninh", "Xã Long Cang", "Xã Long Chữ", "Xã Long Hựu", "Xã Long Thuận", "Xã Lương Hòa", "Xã Mộc Hóa", "Xã Mỹ An", "Xã Mỹ Hạnh", "Xã Mỹ Lệ", "Xã Mỹ Lộc", "Xã Mỹ Quý", "Xã Mỹ Thạnh", "Xã Mỹ Yên", "Xã Nhơn Hòa Lập", "Xã Nhơn Ninh", "Xã Nhựt Tảo", "Xã Ninh Điền", "Xã Phước Chỉ", "Xã Phước Lý", "Xã Phước Thạnh", "Xã Phước Vinh", "Xã Phước Vĩnh Tây", "Xã Rạch Kiến", "Xã Tầm Vu", "Xã Tân Biên", "Xã Tân Châu", "Xã Tân Đông", "Xã Tân Hòa", "Xã Tân Hội", "Xã Tân Hưng", "Xã Tân Lân", "Xã Tân Lập", "Xã Tân Long", "Xã Tân Phú", "Xã Tân Tập", "Xã Tân Tây", "Xã Tân Thạnh", "Xã Tân Thành", "Xã Tân Trụ", "Xã Thạnh Bình", "Xã Thạnh Đức", "Xã Thạnh Hóa", "Xã Thạnh Lợi", "Xã Thạnh Phước", "Xã Thủ Thừa", "Xã Thuận Mỹ", "Xã Trà Vong", "Xã Truông Mít", "Xã Tuyên Bình", "Xã Tuyên Thạnh", "Xã Vàm Cỏ", "Xã Vĩnh Châu", "Xã Vĩnh Công", "Xã Vĩnh Hưng", "Xã Vĩnh Thạnh"],
    "Tỉnh Thái Nguyên": ["Phường Bá Xuyên", "Phường Bắc Kạn", "Phường Bách Quang", "Phường Đức Xuân", "Phường Gia Sàng", "Phường Linh Sơn", "Phường Phan Đình Phùng", "Phường Phổ Yên", "Phường Phúc Thuận", "Phường Quan Triều", "Phường Quyết Thắng", "Phường Sông Công", "Phường Tích Lương", "Phường Trung Thành", "Phường Vạn Xuân", "Xã An Khánh", "Xã Ba Bể", "Xã Bạch Thông", "Xã Bằng Thành", "Xã Bằng Vân", "Xã Bình Thành", "Xã Bình Yên", "Xã Cẩm Giàng", "Xã Cao Minh", "Xã Chợ Đồn", "Xã Chợ Mới", "Xã Chợ Rã", "Xã Côn Minh", "Xã Cường Lợi", "Xã Đại Phúc", "Xã Đại Từ", "Xã Dân Tiến", "Xã Điềm Thụy", "Xã Định Hóa", "Xã Đồng Hỷ", "Xã Đồng Phúc", "Xã Đức Lương", "Xã Hiệp Lực", "Xã Hợp Thành", "Xã Kha Sơn", "Xã Kim Phượng", "Xã La Bằng", "Xã La Hiên", "Xã Lam Vỹ", "Xã Nà Phặc", "Xã Na Rì", "Xã Nam Cường", "Xã Nam Hòa", "Xã Ngân Sơn", "Xã Nghĩa Tá", "Xã Nghiên Loan", "Xã Nghinh Tường", "Xã Phong Quang", "Xã Phú Bình", "Xã Phú Đình", "Xã Phú Lạc", "Xã Phú Lương", "Xã Phú Thịnh", "Xã Phủ Thông", "Xã Phú Xuyên", "Xã Phúc Lộc", "Xã Phượng Tiến", "Xã Quân Chu", "Xã Quảng Bạch", "Xã Quang Sơn", "Xã Sảng Mộc", "Xã Tân Cương", "Xã Tân Khánh", "Xã Tân Kỳ", "Xã Tân Thành", "Xã Thần Sa", "Xã Thành Công", "Xã Thanh Mai", "Xã Thanh Thịnh", "Xã Thượng Minh", "Xã Thượng Quan", "Xã Trại Cau", "Xã Trần Phú", "Xã Tràng Xá", "Xã Trung Hội", "Xã Văn Hán", "Xã Văn Lang", "Xã Văn Lăng", "Xã Vạn Phú", "Xã Vĩnh Thông", "Xã Võ Nhai", "Xã Vô Tranh", "Xã Xuân Dương", "Xã Yên Bình", "Xã Yên Phong", "Xã Yên Thịnh", "Xã Yên Trạch"],
    "Tỉnh Thanh Hóa": ["Phường Bỉm Sơn", "Phường Đào Duy Từ", "Phường Đông Quang", "Phường Đông Sơn", "Phường Đông Tiến", "Phường Hạc Thành", "Phường Hải Bình", "Phường Hải Lĩnh", "Phường Hàm Rồng", "Phường Nam Sầm Sơn", "Phường Nghi Sơn", "Phường Ngọc Sơn", "Phường Nguyệt Viên", "Phường Quảng Phú", "Phường Quang Trung", "Phường Sầm Sơn", "Phường Tân Dân", "Phường Tĩnh Gia", "Phường Trúc Lâm", "Xã An Nông", "Xã Ba Đình", "Xã Bá Thước", "Xã Bát Mọt", "Xã Biện Thượng", "Xã Các Sơn", "Xã Cẩm Tân", "Xã Cẩm Thạch", "Xã Cẩm Thủy", "Xã Cẩm Tú", "Xã Cẩm Vân", "Xã Cổ Lũng", "Xã Công Chính", "Xã Điền Lư", "Xã Điền Quang", "Xã Định Hòa", "Xã Định Tân", "Xã Đồng Lương", "Xã Đông Thành", "Xã Đồng Tiến", "Xã Giao An", "Xã Hà Long", "Xã Hà Trung", "Xã Hậu Lộc", "Xã Hiền Kiệt", "Xã Hồ Vương", "Xã Hoa Lộc", "Xã Hóa Quỳ", "Xã Hoằng Châu", "Xã Hoằng Giang", "Xã Hoằng Hóa", "Xã Hoằng Lộc", "Xã Hoằng Phú", "Xã Hoằng Sơn", "Xã Hoằng Thanh", "Xã Hoằng Tiến", "Xã Hoạt Giang", "Xã Hồi Xuân", "Xã Hợp Tiến", "Xã Kiên Thọ", "Xã Kim Tân", "Xã Lam Sơn", "Xã Linh Sơn", "Xã Lĩnh Toại", "Xã Luận Thành", "Xã Lương Sơn", "Xã Lưu Vệ", "Xã Mậu Lâm", "Xã Minh Sơn", "Xã Mường Chanh", "Xã Mường Lát", "Xã Mường Lý", "Xã Mường Mìn", "Xã Na Mèo", "Xã Nam Xuân", "Xã Nga An", "Xã Nga Sơn", "Xã Nga Thắng", "Xã Ngọc Lặc", "Xã Ngọc Liên", "Xã Ngọc Trạo", "Xã Nguyệt Ấn", "Xã Nhi Sơn", "Xã Như Thanh", "Xã Như Xuân", "Xã Nông Cống", "Xã Phú Lệ", "Xã Phú Xuân", "Xã Pù Luông", "Xã Pù Nhi", "Xã Quan Sơn", "Xã Quảng Bình", "Xã Quang Chiểu", "Xã Quảng Chính", "Xã Quảng Ngọc", "Xã Quảng Ninh", "Xã Quảng Yên", "Xã Quý Lộc", "Xã Quý Lương", "Xã Sao Vàng", "Xã Sơn Điện", "Xã Sơn Thủy", "Xã Tam Chung", "Xã Tam Lư", "Xã Tam Thanh", "Xã Tân Ninh", "Xã Tân Thành", "Xã Tân Tiến", "Xã Tây Đô", "Xã Thạch Bình", "Xã Thạch Lập", "Xã Thạch Quảng", "Xã Thăng Bình", "Xã Thắng Lộc", "Xã Thắng Lợi", "Xã Thanh Kỳ", "Xã Thanh Phong", "Xã Thanh Quân", "Xã Thành Vinh", "Xã Thiên Phủ", "Xã Thiết Ống", "Xã Thiệu Hóa", "Xã Thiệu Quang", "Xã Thiệu Tiến", "Xã Thiệu Toán", "Xã Thiệu Trung", "Xã Thọ Bình", "Xã Thọ Lập", "Xã Thọ Long", "Xã Thọ Ngọc", "Xã Thọ Phú", "Xã Thọ Xuân", "Xã Thượng Ninh", "Xã Thường Xuân", "Xã Tiên Trang", "Xã Tống Sơn", "Xã Triệu Lộc", "Xã Triệu Sơn", "Xã Trung Chính", "Xã Trung Hạ", "Xã Trung Lý", "Xã Trung Sơn", "Xã Trung Thành", "Xã Trường Lâm", "Xã Trường Văn", "Xã Tượng Lĩnh", "Xã Vân Du", "Xã Vạn Lộc", "Xã Văn Nho", "Xã Văn Phú", "Xã Vạn Xuân", "Xã Vĩnh Lộc", "Xã Xuân Bình", "Xã Xuân Chinh", "Xã Xuân Du", "Xã Xuân Hòa", "Xã Xuân Lập", "Xã Xuân Thái", "Xã Xuân Tín", "Xã Yên Định", "Xã Yên Khương", "Xã Yên Nhân", "Xã Yên Ninh", "Xã Yên Phú", "Xã Yên Thắng", "Xã Yên Thọ", "Xã Yên Trường"],
    "Tỉnh Tuyên Quang": ["Phường An Tường", "Phường Bình Thuận", "Phường Hà Giang 1", "Phường Hà Giang 2", "Phường Minh Xuân", "Phường Mỹ Lâm", "Phường Nông Tiến", "Xã Bắc Mê", "Xã Bắc Quang", "Xã Bạch Đích", "Xã Bạch Ngọc", "Xã Bạch Xa", "Xã Bản Máy", "Xã Bằng Hành", "Xã Bằng Lang", "Xã Bình An", "Xã Bình Ca", "Xã Bình Xa", "Xã Cán Tỷ", "Xã Cao Bồ", "Xã Chiêm Hóa", "Xã Côn Lôn", "Xã Đồng Tâm", "Xã Đông Thọ", "Xã Đồng Văn", "Xã Đồng Yên", "Xã Du Già", "Xã Đường Hồng", "Xã Đường Thượng", "Xã Giáp Trung", "Xã Hàm Yên", "Xã Hồ Thầu", "Xã Hòa An", "Xã Hoàng Su Phì", "Xã Hồng Sơn", "Xã Hồng Thái", "Xã Hùng An", "Xã Hùng Đức", "Xã Hùng Lợi", "Xã Khâu Vai", "Xã Khuôn Lùng", "Xã Kiên Đài", "Xã Kiến Thiết", "Xã Kim Bình", "Xã Lâm Bình", "Xã Lao Chải", "Xã Liên Hiệp", "Xã Linh Hồ", "Xã Lực Hành", "Xã Lũng Cú", "Xã Lũng Phìn", "Xã Lùng Tám", "Xã Mậu Duệ", "Xã Mèo Vạc", "Xã Minh Ngọc", "Xã Minh Quang", "Xã Minh Sơn", "Xã Minh Tân", "Xã Minh Thanh", "Xã Nà Hang", "Xã Nấm Dẩn", "Xã Nậm Dịch", "Xã Nghĩa Thuận", "Xã Ngọc Đường", "Xã Ngọc Long", "Xã Nhữ Khê", "Xã Niêm Sơn", "Xã Pà Vầy Sủ", "Xã Phố Bảng", "Xã Phú Linh", "Xã Phú Lương", "Xã Phù Lưu", "Xã Pờ Ly Ngài", "Xã Quản Bạ", "Xã Quang Bình", "Xã Quảng Nguyên", "Xã Sà Phìn", "Xã Sơn Dương", "Xã Sơn Thủy", "Xã Sơn Vĩ", "Xã Sủng Máng", "Xã Tân An", "Xã Tân Long", "Xã Tân Mỹ", "Xã Tân Quang", "Xã Tân Thanh", "Xã Tân Tiến", "Xã Tân Trào", "Xã Tân Trịnh", "Xã Tát Ngà", "Xã Thái Bình", "Xã Thái Hòa", "Xã Thái Sơn", "Xã Thắng Mố", "Xã Thàng Tín", "Xã Thanh Thủy", "Xã Thông Nguyên", "Xã Thuận Hòa", "Xã Thượng Lâm", "Xã Thượng Nông", "Xã Thượng Sơn", "Xã Tiên Nguyên", "Xã Tiên Yên", "Xã Tri Phú", "Xã Trung Hà", "Xã Trung Sơn", "Xã Trung Thịnh", "Xã Trường Sinh", "Xã Tùng Bá", "Xã Tùng Vài", "Xã Vị Xuyên", "Xã Việt Lâm", "Xã Vĩnh Tuy", "Xã Xín Mần", "Xã Xuân Giang", "Xã Xuân Vân", "Xã Yên Cường", "Xã Yên Hoa", "Xã Yên Lập", "Xã Yên Minh", "Xã Yên Nguyên", "Xã Yên Phú", "Xã Yên Sơn", "Xã Yên Thành"],
    "Tỉnh Vĩnh Long": ["Phường An Hội", "Phường Bến Tre", "Phường Bình Minh", "Phường Cái Vồn", "Phường Đông Thành", "Phường Duyên Hải", "Phường Hòa Thuận", "Phường Long Châu", "Phường Long Đức", "Phường Nguyệt Hóa", "Phường Phú Khương", "Phường Phú Tân", "Phường Phước Hậu", "Phường Sơn Đông", "Phường Tân Hạnh", "Phường Tân Ngãi", "Phường Thanh Đức", "Phường Trà Vinh", "Phường Trường Long Hòa", "Xã An Bình", "Xã An Định", "Xã An Hiệp", "Xã An Ngãi Trung", "Xã An Phú Tân", "Xã An Qui", "Xã An Trường", "Xã Ba Tri", "Xã Bảo Thạnh", "Xã Bình Đại", "Xã Bình Phú", "Xã Bình Phước", "Xã Cái Ngang", "Xã Cái Nhum", "Xã Càng Long", "Xã Cầu Kè", "Xã Cầu Ngang", "Xã Châu Hòa", "Xã Châu Hưng", "Xã Châu Thành", "Xã Chợ Lách", "Xã Đại An", "Xã Đại Điền", "Xã Đôn Châu", "Xã Đông Hải", "Xã Đồng Khởi", "Xã Giao Long", "Xã Giồng Trôm", "Xã Hàm Giang", "Xã Hiệp Mỹ", "Xã Hiếu Phụng", "Xã Hiếu Thành", "Xã Hòa Bình", "Xã Hòa Hiệp", "Xã Hòa Minh", "Xã Hùng Hòa", "Xã Hưng Khánh Trung", "Xã Hưng Mỹ", "Xã Hưng Nhượng", "Xã Hương Mỹ", "Xã Lộc Thuận", "Xã Long Hiệp", "Xã Long Hồ", "Xã Long Hòa", "Xã Long Hữu", "Xã Long Thành", "Xã Long Vĩnh", "Xã Lục Sĩ Thành", "Xã Lương Hòa", "Xã Lương Phú", "Xã Lưu Nghiệp Anh", "Xã Mỏ Cày", "Xã Mỹ Chánh Hòa", "Xã Mỹ Long", "Xã Mỹ Thuận", "Xã Ngãi Tứ", "Xã Ngũ Lạc", "Xã Nhị Long", "Xã Nhị Trường", "Xã Nhơn Phú", "Xã Nhuận Phú Tân", "Xã Phong Thạnh", "Xã Phú Phụng", "Xã Phú Quới", "Xã Phú Thuận", "Xã Phú Túc", "Xã Phước Long", "Xã Phước Mỹ Trung", "Xã Quới An", "Xã Quới Điền", "Xã Quới Thiện", "Xã Song Lộc", "Xã Song Phú", "Xã Tam Bình", "Xã Tam Ngãi", "Xã Tân An", "Xã Tân Hào", "Xã Tân Hòa", "Xã Tân Long Hội", "Xã Tân Lược", "Xã Tân Phú", "Xã Tân Quới", "Xã Tân Thành Bình", "Xã Tân Thủy", "Xã Tân Xuân", "Xã Tập Ngãi", "Xã Tập Sơn", "Xã Thạnh Hải", "Xã Thạnh Phong", "Xã Thạnh Phú", "Xã Thạnh Phước", "Xã Thành Thới", "Xã Thạnh Trị", "Xã Thới Thuận", "Xã Tiên Thủy", "Xã Tiểu Cần", "Xã Trà Côn", "Xã Trà Cú", "Xã Trà Ôn", "Xã Trung Hiệp", "Xã Trung Ngãi", "Xã Trung Thành", "Xã Vinh Kim", "Xã Vĩnh Thành", "Xã Vĩnh Xuân"],
};

// Thông tin công ty in trên hoá đơn — chỉnh lại tại đây nếu công ty đổi thông tin
const COMPANY_INFO = {
    name: "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ HILI",
    address: "6/27A Đường Số 3, C/x Lữ Gia, Phường Phú Thọ, TP Hồ Chí Minh, Việt Nam",
    taxCode: "0316296138",
    bankAccount: "19551097 - Ngân Hàng Á Châu ACB – phòng giao dịch Lý Thường Kiệt",
    phone: "0939206865",
    email: "Hilitek@gmail.com",
    city: "TP. Hồ Chí Minh",
    representativeName: "Nguyễn Anh Liêm",
    representativeTitle: "Giám đốc",
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
const THU_VN = ["Chủ Nhật", "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"];
function thuNgayThangNam(iso) {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    return `${THU_VN[d.getDay()]}, ngày ${d.getDate()} tháng ${d.getMonth() + 1} năm ${d.getFullYear()}`;
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
    // Tách nguồn nhập: từ đơn nhập hàng NCC (mã POH...) hay nhập lại kho do đổi trả (mã chứa -RT)
    const supplierInMoves = inMoves.filter((m) => (m.docNo || "").includes("-RT") === false && (m.docNo || "").startsWith("POH"));
    const returnInMoves = inMoves.filter((m) => (m.docNo || "").includes("-RT"));
    const importedFromSupplierQty = supplierInMoves.reduce((s, m) => s + m.qty, 0);
    const importedFromSupplierValue = supplierInMoves.reduce((s, m) => s + m.qty * m.price, 0);
    const importedFromReturnQty = returnInMoves.reduce((s, m) => s + m.qty, 0);
    const importedFromReturnValue = returnInMoves.reduce((s, m) => s + m.qty * m.price, 0);
    const closingQty = p.openingQty + importedQty - exportedQty;
    // Giá nhập là số bạn tự nhập trực tiếp trên sản phẩm (product.costPrice) — không tự tính bình quân.
    const avgCost = p.costPrice || 0;
    return {
        importedQty, importedValue, exportedQty, exportedValue, closingQty, avgCost,
        importedFromSupplierQty, importedFromSupplierValue, importedFromReturnQty, importedFromReturnValue,
    };
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
// Tìm sản phẩm + thông tin series khi chỉ có số series (dùng khi quét series lúc kiểm kho — chưa biết trước thuộc sản phẩm nào).
function findSeriesOwner(products, serial) {
    for (const p of products) {
        if (!p.hasSeries) continue;
        const row = seriesList(p).find((s) => s.serial === serial);
        if (row) return { product: p, row };
    }
    return null;
}
// Số lượng sản phẩm đang được giữ tại "Kho bảo hành" (đã nhận từ khách, chờ xử lý với NCC) — tính từ các phiếu bảo hành
// còn đang "Đang xử lý" hoặc "Xác nhận" (chưa trả khách/chưa từ chối, tức món hàng lỗi vẫn còn nằm ở kho).
function warrantyStockQtyOf(warrantyTickets, productId) {
    return warrantyTickets.reduce((sum, t) => {
        if (t.status !== "pending" && t.status !== "confirmed") return sum;
        const fromCustomer = t.items.filter((it) => it.productId === productId).reduce((s, it) => s + it.qty, 0);
        const fromSupplier = (t.replacementReceived || []).filter((r) => r.productId === productId).reduce((s, r) => s + r.qty, 0);
        return sum + fromCustomer + fromSupplier;
    }, 0);
}
function nextStocktakeCode(stocktakes) {
    let max = 0;
    stocktakes.forEach((s) => {
        const m = /^KK(\d+)$/.exec(s.code || "");
        if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    return "KK" + String(max + 1).padStart(3, "0");
}
function normalizeStocktake(s) {
    return {
        id: s.id || uid(), code: s.code || "", createdAt: s.createdAt || new Date().toISOString(), createdBy: s.createdBy || "",
        note: s.note || "",
        lines: Array.isArray(s.lines) ? s.lines.map((l) => ({
            productId: l.productId, productName: l.productName || "", productCode: l.productCode || "",
            systemQty: Number(l.systemQty) || 0, countedQty: Number(l.countedQty) || 0, diff: Number(l.diff) || 0,
        })) : [],
        warnings: Array.isArray(s.warnings) ? s.warnings : [],
    };
}

// Mã phiếu bảo hành dạng PBH + tháng(2 số) + năm(2 số) + số thứ tự trong tháng đó — vd PBH082601, PBH082602.
function nextWarrantyCode(tickets) {
    const now = new Date();
    const prefix = `PBH${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getFullYear()).slice(-2)}`;
    let max = 0;
    tickets.forEach((t) => {
        if (t.code && t.code.startsWith(prefix)) {
            const seq = parseInt(t.code.slice(prefix.length), 10);
            if (!isNaN(seq)) max = Math.max(max, seq);
        }
    });
    return `${prefix}${String(max + 1).padStart(2, "0")}`;
}
const WARRANTY_TICKET_STATUSES = [
    { id: "pending", label: "Đang xử lý", color: BRASS },
    { id: "confirmed", label: "Xác nhận", color: BLUE },
    { id: "completed", label: "Đã trả khách", color: FOREST },
    { id: "rejected", label: "Từ chối BH", color: RUST },
];
function normalizeWarrantyTicket(t) {
    return {
        id: t.id || uid(), code: t.code || "", createdAt: t.createdAt || new Date().toISOString(), createdBy: t.createdBy || "",
        customerName: t.customerName || "", customerPhone: t.customerPhone || "", customerAddress: t.customerAddress || "",
        receivedDate: t.receivedDate || todayISO(), returnDate: t.returnDate || "",
        status: WARRANTY_TICKET_STATUSES.some((s) => s.id === t.status) ? t.status : "pending",
        items: Array.isArray(t.items) ? t.items.map((it) => ({
            orderId: it.orderId || "", orderCode: it.orderCode || "", productId: it.productId || "", productName: it.productName || "",
            productCode: it.productCode || "", series: Array.isArray(it.series) ? it.series : [], qty: Number(it.qty) || 1,
            warrantyMonths: Number(it.warrantyMonths) || 0, condition: it.condition || "",
        })) : [],
        note: t.note || "",
        // Thông tin xử lý bảo hành — điền khi chuyển sang trạng thái "Xác nhận".
        resolutionType: ["exchange", "refund"].includes(t.resolutionType) ? t.resolutionType : "",
        exchangeSource: ["main", "warranty"].includes(t.exchangeSource) ? t.exchangeSource : "main", // main = xuất từ kho chính, warranty = hàng NCC trả lại (kho bảo hành)
        exchangeProductId: t.exchangeProductId || "", exchangeProductName: t.exchangeProductName || "", exchangeProductCode: t.exchangeProductCode || "",
        exchangeSeries: Array.isArray(t.exchangeSeries) ? t.exchangeSeries : [], exchangeQty: Number(t.exchangeQty) || 1,
        refundAmount: Number(t.refundAmount) || 0,
        // Hàng thay thế NCC đã trả lại — nhập vào Kho bảo hành, sẵn sàng giao khách (khác với hàng lỗi khách mang tới).
        replacementReceived: Array.isArray(t.replacementReceived) ? t.replacementReceived.map((r) => ({
            id: r.id || uid(), productId: r.productId || "", productName: r.productName || "", productCode: r.productCode || "",
            series: Array.isArray(r.series) ? r.series : [], qty: Number(r.qty) || 1, receivedDate: r.receivedDate || todayISO(),
        })) : [],
        // Lý do từ chối bảo hành — điền khi chuyển sang trạng thái "Từ chối BH".
        rejectReason: t.rejectReason || "",
        // Mã phiếu Xuất trả bảo hành — tự sinh khi chuyển sang "Đã trả khách" (định dạng XTBH-{mã phiếu bảo hành gốc}).
        xtbhCode: t.xtbhCode || "",
        stockDeducted: !!t.stockDeducted, // đã trừ kho chính cho lần đổi SP từ kho chính chưa (tránh trừ 2 lần)
    };
}

/* ---------------- Phiếu dịch vụ: Sửa chữa & IT Helpdesk (2 mảng dữ liệu tách riêng) ---------------- */
const REPAIR_STATUSES = [
    { id: "received", label: "Tiếp nhận", color: BRASS },
    { id: "repairing", label: "Đang sửa", color: BLUE },
    { id: "done", label: "Sửa xong", color: FOREST },
    { id: "returned", label: "Đã trả khách", color: FOREST },
    { id: "unrepairable", label: "Không sửa được", color: RUST },
];
function nextRepairCode(tickets) {
    const now = new Date();
    const prefix = `SC${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getFullYear()).slice(-2)}`;
    let max = 0;
    tickets.forEach((t) => { if (t.code && t.code.startsWith(prefix)) { const seq = parseInt(t.code.slice(prefix.length), 10); if (!isNaN(seq)) max = Math.max(max, seq); } });
    return `${prefix}${String(max + 1).padStart(2, "0")}`;
}
function normalizeRepairTicket(t) {
    return {
        id: t.id || uid(), code: t.code || "", createdAt: t.createdAt || new Date().toISOString(), createdBy: t.createdBy || "",
        customerName: t.customerName || "", customerPhone: t.customerPhone || "", customerAddress: t.customerAddress || "",
        deviceName: t.deviceName || "", deviceBrand: t.deviceBrand || "", serial: t.serial || "", issueDescription: t.issueDescription || "",
        receivedDate: t.receivedDate || todayISO(), returnDate: t.returnDate || "",
        estimatedCost: Number(t.estimatedCost) || 0, actualCost: Number(t.actualCost) || 0, vat: VAT_OPTIONS.some((v) => v.id === t.vat) ? t.vat : "VAT8",
        status: REPAIR_STATUSES.some((s) => s.id === t.status) ? t.status : "received",
        note: t.note || "",
    };
}

const HELPDESK_STATUSES = [
    { id: "new", label: "Mới tiếp nhận", color: BRASS },
    { id: "processing", label: "Đang xử lý", color: BLUE },
    { id: "done", label: "Hoàn thành", color: FOREST },
    { id: "cancelled", label: "Đã huỷ", color: RUST },
];
const HELPDESK_TYPES = ["Cài đặt phần mềm", "Sự cố phần cứng", "Mạng - Internet", "Khác"];
function nextHelpdeskCode(tickets) {
    const now = new Date();
    const prefix = `HD${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getFullYear()).slice(-2)}`;
    let max = 0;
    tickets.forEach((t) => { if (t.code && t.code.startsWith(prefix)) { const seq = parseInt(t.code.slice(prefix.length), 10); if (!isNaN(seq)) max = Math.max(max, seq); } });
    return `${prefix}${String(max + 1).padStart(2, "0")}`;
}
function normalizeHelpdeskTicket(t) {
    return {
        id: t.id || uid(), code: t.code || "", createdAt: t.createdAt || new Date().toISOString(), createdBy: t.createdBy || "",
        customerName: t.customerName || "", customerPhone: t.customerPhone || "", customerAddress: t.customerAddress || "",
        requestType: HELPDESK_TYPES.includes(t.requestType) ? t.requestType : HELPDESK_TYPES[0],
        description: t.description || "", assignee: t.assignee || "",
        receivedDate: t.receivedDate || todayISO(), completedDate: t.completedDate || "",
        status: HELPDESK_STATUSES.some((s) => s.id === t.status) ? t.status : "new",
        solution: t.solution || "", note: t.note || "",
    };
}

/* ---------------- Vận chuyển — quản lý vận đơn thủ công, liên kết với đơn hàng ---------------- */
const SHIPPING_CARRIERS = ["J&T Express", "Viettel Post", "Giao Hàng Nhanh (GHN)", "Giao Hàng Tiết Kiệm (GHTK)", "Ninja Van", "Grab Express", "Shopee Express", "Đối tác khác"];
const SHIPPING_STATUSES = [
    { id: "packing", label: "Chờ đóng gói", color: BRASS },
    { id: "picked", label: "Chờ lấy hàng", color: BRASS },
    { id: "shipping", label: "Đang giao", color: BLUE },
    { id: "delivered", label: "Đã giao", color: FOREST },
    { id: "failed", label: "Giao thất bại", color: RUST },
    { id: "returned", label: "Hoàn hàng", color: RUST },
];
function nextShippingCode(tickets) {
    const now = new Date();
    const prefix = `VC${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getFullYear()).slice(-2)}`;
    let max = 0;
    tickets.forEach((t) => { if (t.code && t.code.startsWith(prefix)) { const seq = parseInt(t.code.slice(prefix.length), 10); if (!isNaN(seq)) max = Math.max(max, seq); } });
    return `${prefix}${String(max + 1).padStart(2, "0")}`;
}
function normalizeShippingTicket(t) {
    return {
        id: t.id || uid(), code: t.code || "", createdAt: t.createdAt || new Date().toISOString(), createdBy: t.createdBy || "",
        orderId: t.orderId || "", orderCode: t.orderCode || "",
        carrier: t.carrier || SHIPPING_CARRIERS[0], trackingCode: t.trackingCode || "",
        recipientName: t.recipientName || "", recipientPhone: t.recipientPhone || "", recipientAddress: t.recipientAddress || "",
        packDate: t.packDate || todayISO(), pickupDate: t.pickupDate || "", deliveredDate: t.deliveredDate || "",
        shippingFee: Number(t.shippingFee) || 0, codAmount: Number(t.codAmount) || 0,
        status: SHIPPING_STATUSES.some((s) => s.id === t.status) ? t.status : "packing",
        note: t.note || "",
    };
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

const MIGRATION_MARKER_KEY = STORAGE_KEY + ":shared-migrated";

async function loadData() {
    try {
        // Đã từng chuyển sang kho dùng chung rồi — chỉ đọc từ kho chung từ nay về sau.
        const marker = await window.storage.get(MIGRATION_MARKER_KEY, true).catch(() => null);
        if (marker && marker.value === "1") {
            const shared = await window.storage.get(STORAGE_KEY, true);
            if (shared && shared.value) return JSON.parse(shared.value);
            return null;
        }
        // Lần đầu tiên sau khi bật chế độ dùng chung: đánh dấu ngay (tránh 2 người cùng mở app trong lúc
        // chuyển đổi làm chạy migrate 2 lần đè lên nhau), sau đó thử lấy dữ liệu cá nhân (personal) cũ của
        // đúng người đang mở app này để làm dữ liệu khởi tạo cho kho chung — không mất dữ liệu đã thiết lập lúc test.
        await window.storage.set(MIGRATION_MARKER_KEY, "1", true).catch(() => {});
        const personal = await window.storage.get(STORAGE_KEY, false).catch(() => null);
        if (personal && personal.value) {
            await window.storage.set(STORAGE_KEY, personal.value, true).catch(() => {});
            return JSON.parse(personal.value);
        }
        return null;
    } catch (e) { /* chưa có dữ liệu */ }
    return null;
}
async function saveData(data) {
    try { await window.storage.set(STORAGE_KEY, JSON.stringify(data), true); }
    catch (e) { console.error("Lỗi lưu dữ liệu:", e); }
}

// Bảng mã hậu tố ngắn cho các giá trị thuộc tính phiên bản thường gặp (màu sắc, kích cỡ...) — dùng để tự sinh SKU/mã VT theo phiên bản.
// Giá trị không có trong bảng sẽ tự suy ra 2-3 ký tự đầu (bỏ dấu, viết hoa).
const VARIANT_CODE_MAP = {
    "đen": "BK", "trắng": "WH", "xanh dương": "BL", "xanh biển": "BL", "xanh navy": "NV", "xanh lá": "GRE", "xanh lá cây": "GRE",
    "xám": "GRA", "tím": "PU", "hồng": "PI", "bạc": "SI", "đỏ": "RD", "vàng": "YE", "cam": "OR", "nâu": "BR", "be": "BE", "gold": "GD", "vàng gold": "GD",
    "s": "S", "m": "M", "l": "L", "xl": "XL", "xxl": "XXL", "xs": "XS", "freesize": "FS", "free size": "FS",
};
function stripDiacriticsVN(str) {
    return String(str).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}
function variantValueCode(value) {
    const key = String(value).trim().toLowerCase();
    if (VARIANT_CODE_MAP[key]) return VARIANT_CODE_MAP[key];
    const clean = stripDiacriticsVN(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
    return clean.slice(0, 3) || "XX";
}
// Tích Descartes giữa các danh sách giá trị thuộc tính — vd [["Đen","Trắng"],["S","M"]] -> [["Đen","S"],["Đen","M"],["Trắng","S"],["Trắng","M"]]
function cartesianProduct(arrays) {
    return arrays.reduce((acc, curr) => acc.flatMap((a) => curr.map((c) => [...a, c])), [[]]);
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
        brand: p.brand || "",
        hasSeries: !!p.hasSeries,
        retailPrice: Number(p.retailPrice ?? p.sellPrice ?? 0) || 0,
        wholesalePrice: Number(p.wholesalePrice ?? p.sellPrice ?? 0) || 0,
        costPrice: Number(p.costPrice ?? 0) || 0,
        openingQty: Number(p.openingQty ?? 0) || 0,
        minStockLevel: p.minStockLevel !== undefined ? Number(p.minStockLevel) || 0 : 5,
        sku: p.sku || "",
        vat: p.vat || "VAT10",
        barcode: p.barcode || "",
        supplierId: p.supplierId || "",
        warrantyMonths: (WARRANTY_OPTIONS.includes(Number(p.warrantyMonths)) || Number(p.warrantyMonths) === WARRANTY_LIFETIME) ? Number(p.warrantyMonths) : 0,
        // Phiên bản sản phẩm (màu sắc, kích cỡ...) — các phiên bản cùng 1 sản phẩm gốc chia sẻ chung variantGroupId.
        variantGroupId: p.variantGroupId || null,
        variantAttrs: (p.variantAttrs && typeof p.variantAttrs === "object" && !Array.isArray(p.variantAttrs)) ? p.variantAttrs : null,
        // Sản phẩm dịch vụ (vd Phí sửa chữa, Phí IT Helpdesk) — không quản lý tồn kho, không trừ kho khi bán.
        isService: !!p.isService,
        // Lịch sử thay đổi giá (bán lẻ/bán sỉ/giá nhập) — mỗi lần sửa sản phẩm mà giá đổi sẽ tự ghi lại ai đổi, lúc nào, từ bao nhiêu thành bao nhiêu.
        priceHistory: Array.isArray(p.priceHistory) ? p.priceHistory.map((h) => ({ id: h.id || uid(), date: h.date || new Date().toISOString(), changedBy: h.changedBy || "", field: h.field || "", oldValue: Number(h.oldValue) || 0, newValue: Number(h.newValue) || 0 })) : [],
        image: p.image || null,
        images: Array.isArray(p.images) ? p.images.filter(Boolean).slice(0, 3) : [],
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
        fulfilled: it.fulfilled !== undefined ? !!it.fulfilled : true,
    })) : [];
    const subtotal = items.reduce((s, it) => s + orderLineTotal(it), 0);
    const discountType = o.discountType === "percent" ? "percent" : "amount";
    const discountAmount = discountType === "percent" ? (subtotal * (Number(o.orderDiscount) || 0)) / 100 : (Number(o.orderDiscount) || 0);
    const payable = subtotal - discountAmount + (Number(o.shippingFee) || 0);
    const isFullyPaid = paidAmount >= payable && payable > 0;
    return {
        id: o.id || uid(), code: o.code || "", createdAt,
        customerId: o.customerId || "", channel: o.channel === "online" ? "online" : "store",
        branch: o.branch || BRANCHES[0], seller: o.seller || EMPLOYEES[0], deliveryDate: o.deliveryDate || "",
        shippingAddress: o.shippingAddress ? {
            recipientName: o.shippingAddress.recipientName || "", recipientPhone: o.shippingAddress.recipientPhone || "",
            province: o.shippingAddress.province || "", ward: o.shippingAddress.ward || "", addressDetail: o.shippingAddress.addressDetail || "",
        } : null,
        tags: Array.isArray(o.tags) ? o.tags : [], notes: o.notes || "",
        status, date: o.date || todayISO(),
        creditDays: Number(o.creditDays) || 0,
        shippingAt: o.shippingAt || ((status === "shipping" || status === "delivered" || status === "done") ? createdAt : null),
        deliveredAt: o.deliveredAt || ((status === "delivered" || status === "done") ? createdAt : null),
        paidCompleteAt: o.paidCompleteAt || (isFullyPaid ? createdAt : null),
        cancelledAt: o.cancelledAt || (status === "cancelled" ? createdAt : null), cancelReason: o.cancelReason || "",
        cancelledByRole: o.cancelledByRole || "", cancelledByName: o.cancelledByName || "",
        cancelRequest: o.cancelRequest ? {
            reason: o.cancelRequest.reason || "", requestedAt: o.cancelRequest.requestedAt || createdAt,
            requestedByRole: o.cancelRequest.requestedByRole || "", requestedByName: o.cancelRequest.requestedByName || "",
        } : null,
        returnRequest: o.returnRequest ? {
            type: o.returnRequest.type === "exchange" ? "exchange" : "refund", note: o.returnRequest.note || "",
            returnedItems: Array.isArray(o.returnRequest.returnedItems) ? o.returnRequest.returnedItems : [],
            exchangeItems: Array.isArray(o.returnRequest.exchangeItems) ? o.returnRequest.exchangeItems : [],
            requestedAt: o.returnRequest.requestedAt || createdAt,
            requestedByRole: o.returnRequest.requestedByRole || "", requestedByName: o.returnRequest.requestedByName || "",
        } : null,
        items,
        vat: o.vat || "VAT10", orderDiscount: Number(o.orderDiscount) || 0, discountType, shippingFee: Number(o.shippingFee) || 0, paidAmount,
        // Lịch sử từng lần thu/hoàn tiền theo ngày (dùng cho báo cáo dòng tiền) — dữ liệu cũ chưa có lịch sử chi tiết thì tự suy ra 1 bản ghi từ paidAmount hiện có.
        payments: Array.isArray(o.payments) ? o.payments.map((p) => ({ id: p.id || uid(), date: p.date || createdAt, amount: Number(p.amount) || 0, type: p.type === "hoan" ? "hoan" : "thu" }))
            : (paidAmount > 0 ? [{ id: uid(), date: o.paidCompleteAt || createdAt, amount: paidAmount, type: "thu" }] : []),
        invoiceStatus: o.invoiceStatus === "issued" ? "issued" : "pending", invoiceNo: o.invoiceNo || "",
        approvalStatus: o.approvalStatus === "pending" ? "pending" : "approved", approvalReason: o.approvalReason || "", createdByRole: o.createdByRole || "",
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
function orderDiscountAmount(o) {
    const subtotal = o.items.reduce((s, it) => s + orderLineTotal(it), 0);
    return o.discountType === "percent" ? (subtotal * (o.orderDiscount || 0)) / 100 : (o.orderDiscount || 0);
}
function orderCalc(o) {
    const subtotal = o.items.reduce((s, it) => s + orderLineTotal(it), 0);
    const vp = vatPercent(o.vat);
    const vatTotal = Math.round((subtotal * vp) / (100 + vp));
    const returns = o.returns || [];
    const returnedValue = returns.reduce((s, r) => s + r.returnedItems.reduce((s2, it) => s2 + returnLineTotal(it), 0), 0);
    const exchangeValue = returns.reduce((s, r) => s + r.exchangeItems.reduce((s2, it) => s2 + returnLineTotal(it), 0), 0);
    const discAmt = orderDiscountAmount(o);
    const payable = subtotal - discAmt + (o.shippingFee || 0) - returnedValue + exchangeValue;
    const remaining = payable - (o.paidAmount || 0);
    // Đơn đã huỷ: coi như không có công nợ / giá trị phải thu (đã hoàn kho + hoàn tiền khi huỷ).
    if (o.status === "cancelled") {
        return { subtotal, vatTotal, returnedValue, exchangeValue, discountAmount: discAmt, payable: 0, remaining: 0 };
    }
    return { subtotal, vatTotal, returnedValue, exchangeValue, discountAmount: discAmt, payable, remaining };
}
// Đơn hàng B2B bán trả chậm — hạn thanh toán = ngày tạo + số ngày công nợ. Cảnh báo khi đã quá hạn ≥3 ngày.
function orderDueInfo(o) {
    if (!o.creditDays || o.status === "cancelled") return null;
    const c = orderCalc(o);
    if (c.remaining <= 0) return null;
    const due = new Date(o.createdAt);
    due.setDate(due.getDate() + o.creditDays);
    due.setHours(0, 0, 0, 0);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const daysLeft = Math.round((due.getTime() - today.getTime()) / 86400000);
    return { dueDate: due, daysLeft, overdue: daysLeft < 0, overdue3: daysLeft <= -3 };
}
// Giá sàn được phép bán theo vai trò — không dùng giá vốn/giá nhập (thông tin nội bộ, không cho Kinh doanh/CTV xem).
// Kinh doanh (nhân viên): tối thiểu = Giá bán sỉ. CTV: tối thiểu = Giá bán lẻ - 10%. Admin: không giới hạn.
function minSellPrice(product, role) {
    if (role === "admin") return 0;
    if (role === "ctv") return Math.round((product.retailPrice || 0) * 0.9);
    return product.wholesalePrice || 0; // staff / kinh doanh
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
function nextSupplierReturnCode(po) {
    return `${po.code}-TR${String((po.returns || []).length + 1).padStart(2, "0")}`;
}
function nextQuoteCode(quotations) {
    let max = 0;
    quotations.forEach((q) => { const m = /^BG(\d+)$/.exec(q.code || ""); if (m) max = Math.max(max, parseInt(m[1], 10)); });
    return "BG" + String(max + 1).padStart(3, "0");
}
// Phiếu báo giá — không trừ kho, không ràng buộc, có hạn hiệu lực (mặc định 7 ngày). Chuyển thành đơn hàng thật bằng 1 nút khi khách đồng ý.
function normalizeQuote(q) {
    const createdAt = q.createdAt || new Date().toISOString();
    const defaultExpiry = new Date(createdAt); defaultExpiry.setDate(defaultExpiry.getDate() + 7);
    return {
        id: q.id || uid(), code: q.code || "", createdAt,
        expiryDate: q.expiryDate || defaultExpiry.toISOString(),
        customerId: q.customerId || "", customerName: q.customerName || "", customerPhone: q.customerPhone || "",
        customerAddress: q.customerAddress || "", customerTaxCode: q.customerTaxCode || "",
        branch: q.branch || BRANCHES[0], seller: q.seller || EMPLOYEES[0],
        items: Array.isArray(q.items) ? q.items.map((it) => ({ productId: it.productId, qty: Number(it.qty) || 1, price: Number(it.price) || 0 })) : [],
        vat: q.vat || "VAT10", orderDiscount: Number(q.orderDiscount) || 0, discountType: q.discountType === "percent" ? "percent" : "amount",
        shippingFee: Number(q.shippingFee) || 0, notes: q.notes || "", tags: Array.isArray(q.tags) ? q.tags : [],
        status: q.status || "active", convertedOrderId: q.convertedOrderId || null,
    };
}
function quoteCalc(q) {
    const subtotal = q.items.reduce((s, it) => s + it.qty * it.price, 0);
    const vp = vatPercent(q.vat);
    const vatTotal = Math.round((subtotal * vp) / (100 + vp));
    const discAmt = q.discountType === "percent" ? (subtotal * (q.orderDiscount || 0)) / 100 : (q.orderDiscount || 0);
    const total = subtotal - discAmt + (q.shippingFee || 0);
    return { subtotal, vatTotal, discountAmount: discAmt, total };
}
function quoteExpiryInfo(q) {
    const exp = new Date(q.expiryDate); exp.setHours(23, 59, 59, 999);
    const today = new Date();
    const daysLeft = Math.ceil((exp.getTime() - today.getTime()) / 86400000);
    return { expired: daysLeft < 0, daysLeft };
}
// Tính giá trị thực tế đạt được của 1 kế hoạch — dùng chung cho màn Kế hoạch, cảnh báo KPI, và liên kết với Xếp hạng bán hàng.
// scope: company (toàn công ty) | category (theo nhóm hàng) | product (theo sản phẩm cụ thể); metric: revenue (giá trị) | qty (số lượng)
function planActual(plan, orders, purchaseOrders, products) {
    if (plan.type === "sales") {
        const relevant = orders.filter((o) => o.status !== "cancelled" && o.createdAt.slice(0, 7) === plan.month && (!plan.sellerName || o.seller === plan.sellerName));
        if (plan.scope === "company") {
            return plan.metric === "qty"
                ? relevant.reduce((s, o) => s + o.items.reduce((s2, it) => s2 + it.qty, 0), 0)
                : relevant.reduce((s, o) => s + orderCalc(o).payable, 0);
        }
        let total = 0;
        relevant.forEach((o) => {
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                const match = plan.scope === "product" ? it.productId === plan.targetProductId : !!(p && p.category === plan.targetCategory);
                if (match) total += plan.metric === "qty" ? it.qty : orderLineTotal(it);
            });
        });
        return total;
    }
    const relevantPOs = purchaseOrders.filter((po) => po.status === "received" && po.createdAt.slice(0, 7) === plan.month);
    if (plan.scope === "company") {
        return plan.metric === "qty"
            ? relevantPOs.reduce((s, po) => s + po.items.reduce((s2, it) => s2 + it.qty, 0), 0)
            : relevantPOs.reduce((s, po) => s + po.items.reduce((s2, it) => s2 + it.qty * it.price, 0), 0);
    }
    let total = 0;
    relevantPOs.forEach((po) => {
        po.items.forEach((it) => {
            const p = products.find((x) => x.id === it.productId);
            const match = plan.scope === "product" ? it.productId === plan.targetProductId : !!(p && p.category === plan.targetCategory);
            if (match) total += plan.metric === "qty" ? it.qty : it.qty * it.price;
        });
    });
    return total;
}
const CUSTOMER_GROUPS = [
    { id: "retail", label: "KH Lẻ" },
    { id: "b2b", label: "B2B" },
    { id: "enterprise", label: "Doanh nghiệp" },
];
function normalizeCustomer(c) {
    const base = {
        id: c.id || uid(), code: c.code || "", name: c.name || "", phone: c.phone || "", contactPerson: c.contactPerson || "", note: c.note || "",
        email: c.email || "", taxCode: c.taxCode || "", province: c.province || "", ward: c.ward || "", addressDetail: c.addressDetail || "",
        group: CUSTOMER_GROUPS.some((g) => g.id === c.group) ? c.group : "retail",
        representativeName: c.representativeName || "", representativeTitle: c.representativeTitle || "",
        assignedTo: c.assignedTo || "",
    };
    let addresses = Array.isArray(c.addresses) ? c.addresses.map((a) => ({
        id: a.id || uid(), label: a.label || "Địa chỉ giao hàng", recipientName: a.recipientName || "", recipientPhone: a.recipientPhone || "",
        province: a.province || "", ward: a.ward || "", addressDetail: a.addressDetail || "", isDefault: !!a.isDefault,
    })) : [];
    // Dọn dữ liệu cũ: loại địa chỉ trong sổ trùng khớp hoàn toàn với địa chỉ gốc của khách hàng (được phiên bản trước tự sinh ra),
    // tránh hiển thị trùng lặp với địa chỉ #1 (ảo, luôn lấy trực tiếp từ hồ sơ khách hàng).
    if (base.addressDetail) {
        addresses = addresses.filter((a) => !(a.addressDetail === base.addressDetail && a.province === base.province && a.ward === base.ward && (a.recipientName || "") === base.name && (a.recipientPhone || "") === base.phone));
    }
    return { ...base, addresses };
}
// Sổ địa chỉ đầy đủ của khách hàng: địa chỉ #1 luôn là địa chỉ gốc (nhập khi tạo/sửa hồ sơ khách hàng — tên, SĐT, tỉnh/phường/địa chỉ cụ thể),
// mặc định trừ khi khách có địa chỉ khác trong sổ được đặt làm mặc định riêng. Các địa chỉ thêm qua "+ Thêm địa chỉ mới" xếp từ #2 trở đi.
function customerAddressBook(customer) {
    if (!customer) return [];
    const list = Array.isArray(customer.addresses) ? customer.addresses : [];
    const hasExplicitDefault = list.some((a) => a.isDefault);
    const hasPrimaryInfo = customer.addressDetail || customer.province || customer.ward || customer.phone || customer.name;
    if (!hasPrimaryInfo) return list;
    const primary = {
        id: "primary", label: "Địa chỉ khách hàng", recipientName: customer.name, recipientPhone: customer.phone,
        province: customer.province, ward: customer.ward, addressDetail: customer.addressDetail,
        isDefault: !hasExplicitDefault, isPrimary: true,
    };
    return [primary, ...list];
}
function customerDefaultAddress(customer) {
    const book = customerAddressBook(customer);
    return book.find((a) => a.isDefault) || book[0] || null;
}
// Nhãn hiệu giờ thuộc về 1 nhóm hàng cụ thể (vd nhóm "Phần mềm" có nhãn MS, Adobe...) — hỗ trợ chuyển dữ liệu cũ (chuỗi đơn) sang dạng có nhóm hàng.
function normalizeBrandEntry(b) {
    if (typeof b === "string") return { id: uid(), name: b, category: "" };
    return { id: b.id || uid(), name: b.name || "", category: b.category || "" };
}
function normalizeSupplier(s) {
    return {
        id: s.id || uid(), code: s.code || "", name: s.name || "", taxCode: s.taxCode || "",
        address: s.address || "", contactPerson: s.contactPerson || "", phone: s.phone || "", email: s.email || "",
        paymentTerm: SUPPLIER_PAYMENT_TERMS.some((t) => t.id === s.paymentTerm) ? s.paymentTerm : "cash",
        creditDays: Number(s.creditDays) || 0,
    };
}
function normalizePlan(p) {
    return {
        id: p.id || uid(), month: p.month || todayISO().slice(0, 7), type: p.type === "purchase" ? "purchase" : "sales",
        targetValue: Number(p.targetValue) || 0, note: p.note || "", sellerName: p.sellerName || "",
        scope: ["company", "category", "product"].includes(p.scope) ? p.scope : "company",
        targetCategory: p.targetCategory || "", targetProductId: p.targetProductId || "",
        metric: p.metric === "qty" ? "qty" : "revenue",
    };
}
function normalizeLog(l) {
    return { id: l.id || uid(), at: l.at || new Date().toISOString(), userId: l.userId || "", userName: l.userName || "", role: l.role || "", action: l.action || "", detail: l.detail || "" };
}
function normalizeNotif(n) {
    return {
        id: n.id || uid(), key: n.key || "", category: n.category || "", title: n.title || "", detail: n.detail || "",
        createdAt: n.createdAt || new Date().toISOString(), read: !!n.read, readAt: n.readAt || null,
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
        branch: po.branch || BRANCHES[0], supplier: po.supplier || "", supplierId: po.supplierId || "", createdBy: po.createdBy || "",
        invoiceNo: po.invoiceNo || "", notes: po.notes || "", tags: Array.isArray(po.tags) ? po.tags : [],
        paid, paidAt: po.paidAt || (paid ? createdAt : null), paymentMethod: po.paymentMethod || (paid ? "cash" : "credit"),
        creditDays: Number(po.creditDays) || 0,
        items: Array.isArray(po.items) ? po.items.map((it) => ({
            productId: it.productId, qty: Number(it.qty) || 0, price: Number(it.price) || 0,
            vat: it.vat || "VAT10",
            series: Array.isArray(it.series) ? it.series : [],
        })) : [],
        // Lịch sử trả hàng lại cho NCC (nhập sai/hàng lỗi...) — mỗi phiếu gồm danh sách sản phẩm trả kèm số lượng/series và lý do.
        returns: Array.isArray(po.returns) ? po.returns.map((r) => ({
            id: r.id || uid(), code: r.code || "", createdAt: r.createdAt || createdAt, note: r.note || "",
            items: Array.isArray(r.items) ? r.items.map((it) => ({ productId: it.productId, qty: Number(it.qty) || 0, price: Number(it.price) || 0, series: Array.isArray(it.series) ? it.series : [] })) : [],
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
function poTotal(po) { return po.items.reduce((s, it) => s + it.qty * it.price, 0); }
// Giá trị hàng đã trả lại cho NCC (nhập sai/hàng lỗi...) — trừ vào công nợ vì NCC không còn được nhận số tiền này nữa.
function poReturnedValue(po) { return (po.returns || []).reduce((s, r) => s + r.items.reduce((s2, it) => s2 + it.qty * it.price, 0), 0); }
// Giá trị ròng thực tế còn phải thanh toán cho NCC = giá trị đơn nhập gốc - giá trị đã trả lại.
function poNetTotal(po) { return poTotal(po) - poReturnedValue(po); }

// Thông tin hạn công nợ NCC cho đơn nhập trả chậm — dùng để cảnh báo khi gần/đã quá hạn.
function poDueInfo(po) {
    if (po.paymentMethod !== "credit" || po.paid || !po.creditDays) return null;
    const due = new Date(po.createdAt);
    due.setDate(due.getDate() + po.creditDays);
    due.setHours(0, 0, 0, 0);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const daysLeft = Math.round((due.getTime() - today.getTime()) / 86400000);
    return { dueDate: due, daysLeft, overdue: daysLeft < 0, nearDue: daysLeft <= 3 };
}
// Khớp 1 đơn nhập hàng với đúng nhà cung cấp (ô "Nhà cung cấp" trên đơn nhập là free-text/autocomplete nên khớp theo tên hoặc mã).
// Khớp 1 đơn nhập hàng với đúng nhà cung cấp — ưu tiên khớp chính xác theo supplierId; các đơn cũ chưa có supplierId thì khớp theo tên/mã (fuzzy).
function poMatchesSupplier(po, supplier) {
    if (po.supplierId) return po.supplierId === supplier.id;
    if (!po.supplier) return false;
    const s = po.supplier.toLowerCase();
    return s.includes(supplier.name.toLowerCase()) || s.includes(supplier.code.toLowerCase());
}

// Kiểm tra số series trùng: 1 số series không được nhập kho 2 lần cho cùng 1 mã sản phẩm (dù là phiếu nhập cũ hay mới).
// excludeDocNo: bỏ qua chính phiếu đang sửa (để sửa lại 1 đơn nhập cũ không tự báo trùng với chính nó).
function findDuplicateSeries(product, newSeries, excludeDocNo) {
    const existing = new Set();
    (product.movements || []).forEach((m) => {
        if (m.type === "in" && m.docNo !== excludeDocNo) (m.series || []).forEach((sn) => existing.add(sn.trim().toLowerCase()));
    });
    const seenInBatch = new Set();
    const dups = [];
    newSeries.forEach((sn) => {
        const key = sn.trim().toLowerCase();
        if (!key) return;
        if (existing.has(key) || seenInBatch.has(key)) dups.push(sn);
        seenInBatch.add(key);
    });
    return dups;
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

// Input số tiền — tự hiển thị dấu chấm phân cách hàng nghìn khi gõ (vd 1.650.000), giữ nguyên vị trí con trỏ.
// value: number | "" ; onChange nhận về number | "".
function MoneyInput({ value, onChange, className, style, placeholder }) {
    const ref = useRef(null);
    const display = (value === "" || value === null || value === undefined) ? "" : Number(value).toLocaleString("vi-VN");
    const handleChange = (e) => {
        const el = e.target;
        const cursorPos = el.selectionStart;
        const digitsBeforeCursor = el.value.slice(0, cursorPos).replace(/\D/g, "").length;
        const cleaned = el.value.replace(/\D/g, "");
        const num = cleaned === "" ? "" : Number(cleaned);
        onChange(num);
        requestAnimationFrame(() => {
            if (!ref.current) return;
            const newDisplay = num === "" ? "" : Number(num).toLocaleString("vi-VN");
            let pos = newDisplay.length, digitCount = 0;
            if (digitsBeforeCursor === 0) { pos = 0; }
            else {
                for (let i = 0; i < newDisplay.length; i++) {
                    if (/\d/.test(newDisplay[i])) digitCount++;
                    if (digitCount === digitsBeforeCursor) { pos = i + 1; break; }
                }
            }
            try { ref.current.setSelectionRange(pos, pos); } catch (err) {}
        });
    };
    return (
        <input ref={ref} type="text" inputMode="numeric" placeholder={placeholder}
               value={display} className={className} style={style} onChange={handleChange} />
    );
}

function TagsNotesCompact({ tags, setTags, notes, setNotes }) {
    return (
        <div className="grid grid-cols-2 gap-3 mt-4 pt-3" style={{ borderTop: `1px dashed ${LINE}` }}>
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

/* ---------------- Xem ảnh phóng to (lightbox) — dùng chung cho ảnh sản phẩm, hỗ trợ chuyển đổi giữa nhiều ảnh ---------------- */
function ImageLightbox({ images, startIndex, onClose }) {
    const list = (images || []).filter((im) => im && im.src);
    const [idx, setIdx] = useState(Math.min(Math.max(startIndex || 0, 0), Math.max(list.length - 1, 0)));
    const [copyState, setCopyState] = useState("idle"); // idle | ok | fail
    if (list.length === 0) return null;
    const current = list[idx];
    const goPrev = (e) => { e.stopPropagation(); setIdx((i) => (i - 1 + list.length) % list.length); setCopyState("idle"); };
    const goNext = (e) => { e.stopPropagation(); setIdx((i) => (i + 1) % list.length); setCopyState("idle"); };
    const copyImage = async () => {
        try {
            const res = await fetch(current.src);
            const blob = await res.blob();
            await navigator.clipboard.write([new window.ClipboardItem({ [blob.type]: blob })]);
            setCopyState("ok");
        } catch (e) {
            setCopyState("fail");
        }
        setTimeout(() => setCopyState("idle"), 2200);
    };
    return (
        <div className="fixed inset-0 flex items-center justify-center p-4" style={{ zIndex: 9999, background: "rgba(15,20,32,0.9)" }} onClick={onClose}>
            <div onClick={(e) => e.stopPropagation()} className="relative max-w-[92vw] max-h-[92vh] flex flex-col items-center">
                <div className="relative flex items-center justify-center">
                    {list.length > 1 && (
                        <button onClick={goPrev} className="absolute left-0 -translate-x-12 sm:-translate-x-14 w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "rgba(255,255,255,0.15)", color: "#fff" }}>
                            <ChevronLeft size={20} />
                        </button>
                    )}
                    <img src={current.src} alt={current.alt || "Ảnh"} className="max-w-full object-contain rounded-sm shadow-2xl" style={{ maxHeight: "68vh", background: "#fff" }} />
                    {list.length > 1 && (
                        <button onClick={goNext} className="absolute right-0 translate-x-12 sm:translate-x-14 w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "rgba(255,255,255,0.15)", color: "#fff" }}>
                            <ChevronRight size={20} />
                        </button>
                    )}
                </div>

                {list.length > 1 && (
                    <div className="flex items-center gap-2 mt-4 flex-wrap justify-center max-w-full">
                        {list.map((img, i) => (
                            <img key={i} src={img.src} alt={img.alt || ""} onClick={() => setIdx(i)}
                                 className="w-12 h-12 object-cover rounded-sm cursor-pointer shrink-0"
                                 style={{ border: i === idx ? `2px solid ${BLUE}` : "2px solid rgba(255,255,255,0.25)", opacity: i === idx ? 1 : 0.55 }} />
                        ))}
                    </div>
                )}

                <div className="flex items-center gap-2 mt-4 flex-wrap justify-center">
                    <button onClick={copyImage} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
                        {copyState === "ok" ? <><Check size={14} /> Đã sao chép!</> : copyState === "fail" ? "Không sao chép được — bấm Tải ảnh" : "Sao chép ảnh"}
                    </button>
                    <a href={current.src} download={`${(current.alt || "anh-san-pham").replace(/[^\w\-]+/g, "_")}.png`} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: "rgba(255,255,255,0.4)", color: "#fff" }}>
                        Tải ảnh
                    </a>
                    <button onClick={onClose} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: "rgba(255,255,255,0.4)", color: "#fff" }}>
                        <X size={14} /> Đóng
                    </button>
                </div>
            </div>
        </div>
    );
}

function Modal({ title, onClose, children, wide, size }) {
    const sizeClass = { md: "max-w-md", lg: "max-w-lg", xl: "max-w-3xl", "2xl": "max-w-6xl", "3xl": "max-w-[92rem]" }[size] || (wide ? "max-w-lg" : "max-w-md");
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4" style={{ background: "rgba(31,42,68,0.45)" }}>
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

function MiniStat({ label, value, icon: Icon, accent, onClick }) {
    const [hover, setHover] = useState(false);
    const Tag = onClick ? "button" : "div";
    return (
        <Tag onClick={onClick} onMouseEnter={() => onClick && setHover(true)} onMouseLeave={() => setHover(false)}
             className="flex items-center gap-3 text-left" style={onClick ? { cursor: "pointer" } : undefined}>
            <div className="w-11 h-11 rounded-full flex items-center justify-center shrink-0" style={{ background: `${accent}1A`, boxShadow: onClick && hover ? `0 0 0 2px ${accent}` : "0 0 0 2px transparent", transition: "box-shadow 0.15s ease" }}>
                <Icon size={18} style={{ color: accent }} />
            </div>
            <div>
                <p className="text-sm opacity-60">{label}</p>
                <p className="text-lg font-semibold" style={{ color: INK }}>{value}</p>
            </div>
        </Tag>
    );
}

function Dashboard({ products, orders, goToOrdersFilter }) {
    const today = todayISO();
    const isToday = (dateStr) => (dateStr || "").slice(0, 10) === today;

    const stockValue = products.reduce((s, p) => s + productStats(p).closingQty * productStats(p).avgCost, 0);
    const stockUnits = products.reduce((s, p) => s + productStats(p).closingQty, 0);
    const totalIn = products.reduce((s, p) => s + productStats(p).importedValue, 0);
    const totalOut = products.reduce((s, p) => s + productStats(p).exportedValue, 0);
    const lowStock = products.filter((p) => productStats(p).closingQty <= (p.minStockLevel ?? 5) && productStats(p).closingQty >= 0);

    // ---- Kết quả kinh doanh trong ngày ----
    const todayOrders = orders.filter((o) => isToday(o.createdAt));
    const todayRevenue = todayOrders.filter((o) => o.status !== "cancelled").reduce((s, o) => s + orderCalc(o).payable, 0);
    const todayNewOrders = todayOrders.length;
    const todayReturns = orders.reduce((s, o) => s + (o.returns || []).filter((r) => isToday(r.createdAt)).length, 0);
    const todayCancelled = orders.filter((o) => o.status === "cancelled" && isToday(o.cancelledAt)).length;

    // ---- Biểu đồ doanh thu theo ngày (có chọn khoảng ngày) ----
    const [chartDays, setChartDays] = useState(7);
    const dailyRevenue = useMemo(() => {
        const days = [];
        for (let i = chartDays - 1; i >= 0; i--) {
            const d = new Date(); d.setDate(d.getDate() - i);
            const key = d.toISOString().slice(0, 10);
            days.push({ key, label: `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`, total: 0 });
        }
        orders.filter((o) => o.status !== "cancelled").forEach((o) => {
            const key = o.createdAt.slice(0, 10);
            const d = days.find((x) => x.key === key);
            if (d) d.total += orderCalc(o).payable;
        });
        return days;
    }, [orders, chartDays]);
    const chartTotal = dailyRevenue.reduce((s, d) => s + d.total, 0);

    // ---- Đơn hàng chờ xử lý ----
    const pendingApproval = orders.filter((o) => o.status !== "cancelled" && o.approvalStatus === "pending").length;
    const pendingPayment = orders.filter((o) => o.status !== "cancelled" && orderCalc(o).remaining > 0).length;
    const pendingProcess = orders.filter((o) => o.status === "pending").length;
    const shippingNow = orders.filter((o) => o.status === "shipping").length;
    const deliveredNotDone = orders.filter((o) => o.status === "delivered").length;
    const pendingRequests = orders.filter((o) => o.status !== "cancelled" && (o.cancelRequest || o.returnRequest)).length;

    // ---- Top sản phẩm theo số lượng bán (7 ngày qua) ----
    const [topDays, setTopDays] = useState(7);
    const topProductsByQty = useMemo(() => {
        const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - topDays); cutoff.setHours(0, 0, 0, 0);
        const map = {};
        products.forEach((p) => {
            p.movements.filter((m) => m.type === "out" && new Date(m.date) >= cutoff).forEach((m) => {
                map[p.id] = map[p.id] || { name: p.name, code: p.code, qty: 0 };
                map[p.id].qty += m.qty;
            });
        });
        return Object.values(map).sort((a, b) => b.qty - a.qty).slice(0, 5);
    }, [products, topDays]);

    const pieColors = [INK, BRASS, FOREST, BLUE, RUST];
    const topProductsByValue = useMemo(() => {
        return products
            .map((p) => ({ name: p.name, total: productStats(p).exportedValue }))
            .filter((p) => p.total > 0)
            .sort((a, b) => b.total - a.total)
            .slice(0, 5);
    }, [products]);

    return (
        <div>
            {/* Kết quả kinh doanh trong ngày */}
            <div className="p-5 rounded-sm mb-6" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6, letterSpacing: "0.06em" }}>Kết quả kinh doanh trong ngày</h4>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-5">
                    <MiniStat label="Doanh thu" value={vnd(todayRevenue)} icon={Wallet} accent={BLUE} />
                    <MiniStat label="Đơn hàng mới" value={todayNewOrders} icon={Package} accent={FOREST} />
                    <MiniStat label="Đơn trả hàng" value={todayReturns} icon={RotateCcw} accent={BRASS} />
                    <MiniStat label="Đơn huỷ" value={todayCancelled} icon={XCircle} accent={RUST} />
                </div>
            </div>

            {/* Doanh thu bán hàng theo ngày */}
            <div className="p-5 rounded-sm mb-6" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                    <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.6, letterSpacing: "0.06em" }}>Doanh thu bán hàng</h4>
                    <div className="flex gap-1.5">
                        {[7, 14, 30].map((d) => (
                            <button key={d} onClick={() => setChartDays(d)} className="text-xs px-3 py-1.5 rounded-full border"
                                    style={{ borderColor: chartDays === d ? INK : LINE, background: chartDays === d ? INK : "transparent", color: chartDays === d ? "#fff" : INK }}>
                                {d} ngày qua
                            </button>
                        ))}
                    </div>
                </div>
                <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={dailyRevenue} barCategoryGap="30%">
                        <CartesianGrid vertical={false} strokeDasharray="3 3" stroke={LINE} />
                        <XAxis dataKey="label" tick={{ fontSize: 11, fill: INK }} axisLine={{ stroke: LINE }} tickLine={false} interval={chartDays > 14 ? Math.ceil(chartDays / 12) : 0} />
                        <YAxis tick={{ fontSize: 11, fill: INK }} axisLine={false} tickLine={false} width={56} tickFormatter={(v) => (v >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                        <Tooltip formatter={(v) => vnd(v)} contentStyle={{ fontFamily: "Inter", fontSize: 13, borderRadius: 8, border: `1px solid ${LINE}` }} cursor={{ fill: PAPER }} />
                        <Bar dataKey="total" fill={BLUE} radius={[6, 6, 0, 0]} maxBarSize={40} />
                    </BarChart>
                </ResponsiveContainer>
                <p className="text-sm mt-2" style={{ color: INK }}>Tổng doanh thu: <span className="font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(chartTotal)}</span></p>
            </div>

            {/* Đơn hàng chờ xử lý */}
            <div className="p-5 rounded-sm mb-6" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6, letterSpacing: "0.06em" }}>Đơn hàng chờ xử lý</h4>
                <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-5">
                    <MiniStat label="Chờ duyệt" value={pendingApproval} icon={ShieldCheck} accent={RUST} onClick={() => goToOrdersFilter && goToOrdersFilter("approval_pending")} />
                    <MiniStat label="Chờ thanh toán" value={pendingPayment} icon={Wallet} accent={BRASS} onClick={() => goToOrdersFilter && goToOrdersFilter("delivered")} />
                    <MiniStat label="Chờ xử lý" value={pendingProcess} icon={Clock} accent={BRASS} onClick={() => goToOrdersFilter && goToOrdersFilter("pending")} />
                    <MiniStat label="Đang giao hàng" value={shippingNow} icon={Truck} accent={BLUE} onClick={() => goToOrdersFilter && goToOrdersFilter("shipping")} />
                    <MiniStat label="Đã giao" value={deliveredNotDone} icon={PackageCheck} accent={FOREST} onClick={() => goToOrdersFilter && goToOrdersFilter("delivered")} />
                    <MiniStat label="Yêu cầu huỷ/đổi trả" value={pendingRequests} icon={RotateCcw} accent={RUST} onClick={() => goToOrdersFilter && goToOrdersFilter("return_request")} />
                </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
                <StatCard label="Tổng giá trị xuất (doanh thu)" value={vnd(totalOut)} icon={TrendingUp} accent={FOREST} />
                <StatCard label="Tổng giá trị nhập" value={vnd(totalIn)} icon={ArrowDownToLine} accent={BRASS} />
                <StatCard label="Giá trị tồn kho hiện tại" value={vnd(stockValue)} icon={Package} accent={BLUE} />
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="p-5 rounded-sm min-w-0" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                        <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.6, letterSpacing: "0.06em" }}>Top sản phẩm bán chạy</h4>
                        <div className="flex gap-1.5">
                            {[7, 30].map((d) => (
                                <button key={d} onClick={() => setTopDays(d)} className="text-xs px-2.5 py-1 rounded-full border"
                                        style={{ borderColor: topDays === d ? INK : LINE, background: topDays === d ? INK : "transparent", color: topDays === d ? "#fff" : INK }}>
                                    {d} ngày
                                </button>
                            ))}
                        </div>
                    </div>
                    {topProductsByQty.length === 0 ? <p className="text-sm opacity-50 text-center py-10">Chưa có dữ liệu xuất kho trong khoảng thời gian này.</p> : (
                        <div className="space-y-3">
                            {topProductsByQty.map((p, i) => (
                                <div key={p.code} className="flex items-center gap-3">
                                    <span className="w-6 h-6 rounded-full flex items-center justify-center text-[11px] font-semibold shrink-0" style={{ background: `${pieColors[i % pieColors.length]}1A`, color: pieColors[i % pieColors.length] }}>{String(i + 1).padStart(2, "0")}</span>
                                    <div className="min-w-0 flex-1">
                                        <p className="text-sm truncate" style={{ color: INK }}>{p.name}</p>
                                        <p className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.code}</p>
                                    </div>
                                    <span className="text-sm font-semibold shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{p.qty}</span>
                                </div>
                            ))}
                        </div>
                    )}
                </div>

                <div className="p-5 rounded-sm min-w-0" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.6, letterSpacing: "0.06em" }}>Thông tin kho</h4>
                    <div className="space-y-1">
                        <div className="flex items-center justify-between py-3" style={{ borderBottom: `1px dashed ${LINE}` }}>
              <span className="text-sm flex items-center gap-1.5" style={{ color: lowStock.length > 0 ? RUST : INK }}>
                {lowStock.length > 0 && <AlertTriangle size={14} />} Sản phẩm dưới định mức (≤5)
              </span>
                            <span className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: lowStock.length > 0 ? RUST : INK }}>{lowStock.length}</span>
                        </div>
                        <div className="flex items-center justify-between py-3" style={{ borderBottom: `1px dashed ${LINE}` }}>
                            <span className="text-sm opacity-70">Số tồn kho (tổng SL)</span>
                            <span className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{stockUnits}</span>
                        </div>
                        <div className="flex items-center justify-between py-3">
                            <span className="text-sm opacity-70">Giá trị tồn kho</span>
                            <span className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(stockValue)}</span>
                        </div>
                    </div>
                    {lowStock.length > 0 && (
                        <div className="mt-3 p-3 rounded-sm text-xs" style={{ background: "#FBF0EC", color: INK }}>
                            {lowStock.slice(0, 6).map((p) => `${p.name} (còn ${productStats(p).closingQty})`).join(", ")}{lowStock.length > 6 ? "…" : ""}
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
}
function StatCard({ label, value, icon: Icon, accent }) {
    return (
        <div className="p-5 rounded-sm flex items-center justify-between" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div>
                <p className="text-xs uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55 }}>{label}</p>
                <p className="text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{value}</p>
            </div>
            <div className="w-10 h-10 rounded-full flex items-center justify-center shrink-0" style={{ background: `${accent}1A` }}>
                <Icon size={18} style={{ color: accent }} />
            </div>
        </div>
    );
}

/* ---------------- Products & Inventory (Sản phẩm & Tồn kho) ---------------- */

function ProductsInventory({ products, setProducts, addLog, currentUser, focusProductId, onFocusHandled, goToDoc, suppliers, goToSupplier, categories, setCategories, brands, setBrands }) {
    const isAdmin = currentUser.role === "admin";
    const isCtv = currentUser.role === "ctv";
    const [query, setQuery] = useState("");
    const [editing, setEditing] = useState(null); // sản phẩm đang thêm/sửa thông tin
    const [form, setForm] = useState({});
    const [ioModal, setIoModal] = useState(null); // { product, type: 'in'|'out' }
    const [ioForm, setIoForm] = useState({});
    const [viewingId, setViewingId] = useState(null); // id sản phẩm đang xem chi tiết
    const [historyPage, setHistoryPage] = useState(1);
    const [historySearch, setHistorySearch] = useState("");
    const [historyFrom, setHistoryFrom] = useState("");
    const [historyTo, setHistoryTo] = useState("");
    const HISTORY_PAGE_SIZE = 20;
    const [filterCategory, setFilterCategory] = useState("");
    const [filterBrand, setFilterBrand] = useState("");
    const [filterVat, setFilterVat] = useState("");
    const [selectedIds, setSelectedIds] = useState(new Set());
    const [zoomImage, setZoomImage] = useState(null); // { src, alt } — ảnh đang phóng to
    const [managingCategories, setManagingCategories] = useState(false);
    const [newCategoryInput, setNewCategoryInput] = useState("");
    const [managingBrands, setManagingBrands] = useState(false);
    const [showPriceHistory, setShowPriceHistory] = useState(false);
    const [newBrandInputByCat, setNewBrandInputByCat] = useState({}); // { [tênNhómHàng]: text đang gõ }

    const categoryOptions = [...(categories || [])].sort();
    // Nhãn hiệu giờ thuộc về 1 nhóm hàng cụ thể — brandOptions (phẳng, dùng cho bộ lọc) và brandOptionsOf(category) (dùng cho form sản phẩm theo đúng nhóm hàng đã chọn).
    const brandOptions = [...new Set((brands || []).map((b) => b.name))].sort();
    const brandOptionsOf = (cat) => [...(brands || [])].filter((b) => b.category === cat).map((b) => b.name).sort();

    const filtered = products.filter(
        (p) => (p.name.toLowerCase().includes(query.toLowerCase()) || p.code.toLowerCase().includes(query.toLowerCase()))
            && (!filterCategory || p.category === filterCategory)
            && (!filterBrand || p.brand === filterBrand)
            && (!filterVat || p.vat === filterVat)
    );
    const toggleSelect = (id) => setSelectedIds((prev) => { const next = new Set(prev); next.has(id) ? next.delete(id) : next.add(id); return next; });
    const toggleSelectAll = () => setSelectedIds((prev) => (prev.size === filtered.length ? new Set() : new Set(filtered.map((p) => p.id))));

    // Xuất Excel: theo sản phẩm đã chọn, hoặc theo bộ lọc hiện tại (nhóm hàng/nhãn hiệu/VAT/tìm kiếm) nếu chưa chọn dòng nào.
    const exportProducts = () => {
        const list = selectedIds.size > 0 ? products.filter((p) => selectedIds.has(p.id)) : filtered;
        if (list.length === 0) { alert("Không có sản phẩm nào để xuất."); return; }
        const rows = list.map((p) => {
            const s = productStats(p);
            const row = {
                "Mã VT": p.code, "SKU": p.sku, "Tên vật tư": p.name, "Phiên bản": p.variantAttrs ? Object.values(p.variantAttrs).join(" / ") : "", "Nhóm hàng": p.category, "Nhãn hiệu": p.brand, "ĐVT": p.unit,
                "Quản lý series": p.hasSeries ? "Có" : "Không", "VAT": VAT_OPTIONS.find((v) => v.id === p.vat)?.label || p.vat, "Bảo hành": warrantyLabel(p.warrantyMonths || 0),
                "Tồn đầu kỳ": p.openingQty, "Nhập từ NCC": s.importedFromSupplierQty, "Nhập lại (đổi trả)": s.importedFromReturnQty,
                "Xuất trong kỳ": s.exportedQty, "Tồn cuối kỳ": s.closingQty, "Giá bán sỉ": p.wholesalePrice, "Giá bán lẻ": p.retailPrice,
            };
            if (isAdmin) row["Giá nhập"] = p.costPrice;
            return row;
        });
        exportExcel(`SanPham_${todayISO()}`, [{ name: "Sản phẩm", rows }]);
        addLog("Xuất Excel sản phẩm", `${list.length} sản phẩm`);
    };

    // Nhập sản phẩm hàng loạt từ file Excel — dùng đúng định dạng file "Xuất Excel": tải file xuất ra, thêm dòng mới rồi tải lên lại.
    // Mã VT hoặc SKU đã có thì bỏ qua (không ghi đè), mã mới thì thêm vào.
    const importFileRef = useRef(null);
    const [importResult, setImportResult] = useState(null); // { added, skipped: [] } — hiện sau khi nhập xong

    const parseVatCell = (v) => {
        const s = String(v ?? "").trim().toLowerCase();
        if (!s) return "VAT10";
        if (s.includes("kct")) return "KCT";
        if (s.includes("10")) return "VAT10";
        if (s.includes("8")) return "VAT8";
        if (s.includes("0")) return "VAT0";
        return "VAT10";
    };
    const parseWarrantyCell = (v) => {
        const s = String(v ?? "").trim().toLowerCase();
        if (!s || s === "0") return 0;
        if (s.includes("vĩnh viễn") || s.includes("vinh vien") || s === "-1") return WARRANTY_LIFETIME;
        const n = parseInt(s, 10);
        return WARRANTY_OPTIONS.includes(n) ? n : 0;
    };
    const triggerImportFile = () => importFileRef.current && importFileRef.current.click();
    const handleImportFile = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: "array" });
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                const existingCodes = new Set(products.map((p) => p.code.trim().toLowerCase()));
                const existingSkus = new Set(products.filter((p) => p.sku).map((p) => p.sku.trim().toLowerCase()));
                const newProducts = [];
                const skipped = [];
                const newCatsSet = new Set(categories || []);
                const newBrandsList = [...(brands || [])];
                rows.forEach((row) => {
                    const code = String(row["Mã VT"] ?? "").trim();
                    const name = String(row["Tên vật tư"] ?? "").trim();
                    if (!code || !name) return; // bỏ qua dòng thiếu thông tin bắt buộc (mã VT / tên)
                    const codeKey = code.toLowerCase();
                    const sku = String(row["SKU"] ?? "").trim();
                    const skuKey = sku.toLowerCase();
                    if (existingCodes.has(codeKey) || (skuKey && existingSkus.has(skuKey))) { skipped.push(code); return; }
                    existingCodes.add(codeKey);
                    if (skuKey) existingSkus.add(skuKey);
                    const category = String(row["Nhóm hàng"] ?? "").trim();
                    const brand = String(row["Nhãn hiệu"] ?? "").trim();
                    if (category && !newCatsSet.has(category)) newCatsSet.add(category);
                    if (brand && !newBrandsList.some((b) => b.name === brand && b.category === category)) newBrandsList.push({ id: uid(), name: brand, category });
                    newProducts.push(normalizeProduct({
                        id: uid(), code, name, sku: sku || nextSKU([...products, ...newProducts]),
                        category, brand, unit: String(row["ĐVT"] ?? "").trim() || UNITS[0],
                        hasSeries: /^(có|co|yes|true|1|x)$/i.test(String(row["Quản lý series"] ?? "").trim()),
                        vat: parseVatCell(row["VAT"]), warrantyMonths: parseWarrantyCell(row["Bảo hành"]),
                        retailPrice: Number(row["Giá bán lẻ"]) || 0, wholesalePrice: Number(row["Giá bán sỉ"]) || 0, costPrice: Number(row["Giá nhập"]) || 0,
                        openingQty: Number(row["Tồn đầu kỳ"]) || 0,
                        minStockLevel: row["Định mức tồn tối thiểu"] !== undefined && row["Định mức tồn tối thiểu"] !== "" ? Number(row["Định mức tồn tối thiểu"]) : 5,
                        barcode: String(row["Mã vạch"] ?? "").trim(), movements: [],
                    }));
                });
                if (newProducts.length > 0) setProducts((prev) => [...prev, ...newProducts]);
                if (newCatsSet.size !== (categories || []).length) setCategories([...newCatsSet]);
                if (newBrandsList.length !== (brands || []).length) setBrands(newBrandsList);
                addLog("Nhập sản phẩm từ Excel", `${newProducts.length} sản phẩm mới${skipped.length > 0 ? ` · Bỏ qua ${skipped.length} mã đã tồn tại` : ""}`);
                setImportResult({ added: newProducts.length, skipped });
            } catch (err) {
                alert("Không đọc được file Excel — vui lòng dùng đúng file đã tải từ nút \"Xuất Excel\" (chỉ thêm dòng mới, không đổi tên cột) rồi thử lại.");
            }
            if (importFileRef.current) importFileRef.current.value = "";
        };
        reader.readAsArrayBuffer(file);
    };

    const openNew = () => { setForm({ code: "", name: "", unit: UNITS[0], category: "", brand: "", hasSeries: false, isService: false, retailPrice: "", wholesalePrice: "", costPrice: "", openingQty: 0, minStockLevel: 5, sku: nextSKU(products), vat: "VAT10", barcode: "", supplierId: "", warrantyMonths: 0, image: null, images: [], hasVariants: false, variantAttr1Name: "Màu sắc", variantAttr1Values: [], variantAttr2Enabled: false, variantAttr2Name: "Kích cỡ", variantAttr2Values: [] }); setEditing({}); };
    const openEdit = (p) => { setForm({ ...p }); setEditing(p); };
    const submitInfo = () => {
        if (!form.code || !form.name) return;
        if (editing.id) {
            setProducts((prev) => prev.map((p) => {
                if (p.id !== editing.id) return p;
                const newRetail = Number(form.retailPrice) || 0, newWholesale = Number(form.wholesalePrice) || 0, newCost = Number(form.costPrice) || 0;
                const changes = [];
                if (newRetail !== p.retailPrice) changes.push({ field: "Giá bán lẻ", oldValue: p.retailPrice, newValue: newRetail });
                if (newWholesale !== p.wholesalePrice) changes.push({ field: "Giá bán sỉ", oldValue: p.wholesalePrice, newValue: newWholesale });
                if (newCost !== p.costPrice) changes.push({ field: "Giá nhập", oldValue: p.costPrice, newValue: newCost });
                const now = new Date().toISOString();
                const newHistoryEntries = changes.map((c) => ({ id: uid(), date: now, changedBy: currentUser.fullName, field: c.field, oldValue: c.oldValue, newValue: c.newValue }));
                return {
                    ...p, code: form.code, name: form.name, unit: form.unit, category: form.category || "", brand: form.brand || "", hasSeries: !!form.hasSeries, isService: !!form.isService,
                    retailPrice: newRetail, wholesalePrice: newWholesale, costPrice: newCost, openingQty: Number(form.openingQty) || 0,
                    minStockLevel: Number(form.minStockLevel) || 0,
                    sku: form.sku || p.sku, vat: form.vat, barcode: form.barcode || "", supplierId: form.supplierId || "", warrantyMonths: Number(form.warrantyMonths) || 0, image: form.image || null, images: Array.isArray(form.images) ? form.images.filter(Boolean).slice(0, 3) : [],
                    priceHistory: newHistoryEntries.length > 0 ? [...newHistoryEntries, ...(p.priceHistory || [])] : (p.priceHistory || []),
                };
            }));
            addLog("Sửa sản phẩm", `${form.code} · ${form.name}`);
        } else if (form.hasVariants) {
            // Tạo hàng loạt phiên bản (màu sắc/kích cỡ...) — mỗi phiên bản là 1 sản phẩm riêng, dùng chung thông tin nền,
            // mã VT/SKU tự thêm hậu tố theo giá trị thuộc tính (vd HI003_BK cho màu Đen).
            const attr1Name = (form.variantAttr1Name || "Thuộc tính 1").trim();
            const attr1Values = form.variantAttr1Values || [];
            const attr2Name = (form.variantAttr2Name || "Thuộc tính 2").trim();
            const attr2Values = form.variantAttr2Enabled ? (form.variantAttr2Values || []) : [];
            if (attr1Values.length === 0) { alert(`Vui lòng nhập ít nhất 1 giá trị cho "${attr1Name}".`); return; }
            const combos = cartesianProduct(attr2Values.length > 0 ? [attr1Values, attr2Values] : [attr1Values]);
            const groupId = uid();
            const newProducts = combos.map((combo) => {
                const suffix = combo.map((v) => variantValueCode(v)).join("_");
                const label = combo.join(", ");
                const variantAttrs = { [attr1Name]: combo[0] };
                if (combo[1] !== undefined) variantAttrs[attr2Name] = combo[1];
                return {
                    id: uid(), code: `${form.code}_${suffix}`, name: `${form.name} - ${label}`, unit: form.unit, category: form.category || "", brand: form.brand || "",
                    hasSeries: !!form.hasSeries, retailPrice: Number(form.retailPrice) || 0, wholesalePrice: Number(form.wholesalePrice) || 0, costPrice: Number(form.costPrice) || 0,
                    openingQty: Number(form.openingQty) || 0, minStockLevel: Number(form.minStockLevel) || 0,
                    sku: `${form.sku || nextSKU(products)}_${suffix}`, vat: form.vat || "VAT10", barcode: "", supplierId: form.supplierId || "", warrantyMonths: Number(form.warrantyMonths) || 0,
                    image: form.image || null, images: Array.isArray(form.images) ? form.images.filter(Boolean).slice(0, 3) : [],
                    variantGroupId: groupId, variantAttrs, movements: [],
                };
            });
            // Tránh trùng mã VT nếu vô tình bấm tạo 2 lần hoặc trùng với sản phẩm có sẵn.
            const dup = newProducts.find((np) => products.some((p) => p.code.toLowerCase() === np.code.toLowerCase()));
            if (dup) { alert(`Mã VT "${dup.code}" đã tồn tại — vui lòng đổi Mã VT gốc hoặc kiểm tra lại danh sách giá trị.`); return; }
            setProducts((prev) => [...prev, ...newProducts]);
            addLog("Tạo sản phẩm có phiên bản", `${form.name} · ${newProducts.length} phiên bản`);
        } else {
            setProducts((prev) => [...prev, {
                id: uid(), code: form.code, name: form.name, unit: form.unit, category: form.category || "", brand: form.brand || "",
                hasSeries: !!form.hasSeries, isService: !!form.isService, retailPrice: Number(form.retailPrice) || 0, wholesalePrice: Number(form.wholesalePrice) || 0, costPrice: Number(form.costPrice) || 0,
                openingQty: Number(form.openingQty) || 0, minStockLevel: Number(form.minStockLevel) || 0,
                sku: form.sku || nextSKU(products), vat: form.vat || "VAT10", barcode: form.barcode || "", supplierId: form.supplierId || "", warrantyMonths: Number(form.warrantyMonths) || 0, image: form.image || null, images: Array.isArray(form.images) ? form.images.filter(Boolean).slice(0, 3) : [],
                movements: [],
            }]);
            addLog("Thêm sản phẩm", `${form.code} · ${form.name}`);
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
    // Ảnh phụ (tối đa 3 ảnh) — idx từ 0 đến 2
    const onGalleryImagePick = async (idx, e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        if (file.size > 1.5 * 1024 * 1024) { alert("Ảnh quá lớn — vui lòng chọn ảnh dưới 1.5MB."); return; }
        const dataUrl = await fileToDataUrl(file);
        setForm((f) => {
            const imgs = Array.isArray(f.images) ? [...f.images] : [];
            imgs[idx] = dataUrl;
            return { ...f, images: imgs.slice(0, 3) };
        });
    };
    const removeGalleryImage = (idx) => setForm((f) => ({ ...f, images: (Array.isArray(f.images) ? f.images : []).filter((_, i) => i !== idx) }));
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
                const dups = findDuplicateSeries(product, series);
                if (dups.length > 0) {
                    alert(`Số series bị trùng với phiếu nhập trước đó (số series phải là duy nhất cho mỗi sản phẩm): ${dups.join(", ")}.`);
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
    const negativeStockProducts = products.filter((p) => productStats(p).closingQty < 0);
    const openProductDetail = (id) => { setViewingId(id); setHistoryPage(1); setHistorySearch(""); setHistoryFrom(""); setHistoryTo(""); setShowPriceHistory(false); };

    useEffect(() => {
        if (focusProductId) {
            openProductDetail(focusProductId);
            onFocusHandled && onFocusHandled();
        }
    }, [focusProductId]);

    return (
        <div>
            {negativeStockProducts.length > 0 && (
                <div className="mb-4 p-3 rounded-sm flex items-start gap-2.5" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44` }}>
                    <AlertTriangle size={16} style={{ color: RUST }} className="mt-0.5 shrink-0" />
                    <p className="text-sm" style={{ color: INK }}>
                        <span className="font-medium" style={{ color: RUST }}>Cảnh báo âm kho: </span>
                        {negativeStockProducts.map((p) => `${p.name} (${productStats(p).closingQty})`).join(", ")} — đã bán vượt tồn kho thực tế, cần nhập bù.
                    </p>
                </div>
            )}
            <div className="flex items-center justify-between mb-3 gap-3 flex-wrap">
                <div className="relative flex-1 max-w-xs">
                    <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
                    <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã VT hoặc tên…"
                           className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
                </div>
                <div className="flex gap-2 flex-wrap">
                    <button onClick={exportProducts} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border whitespace-nowrap" style={{ borderColor: FOREST, color: FOREST }}>
                        <FileSpreadsheet size={15} /> Xuất Excel{selectedIds.size > 0 ? ` (${selectedIds.size})` : ""}
                    </button>
                    {isAdmin && (
                        <button onClick={triggerImportFile} title="Dùng file đã tải từ nút Xuất Excel — thêm dòng sản phẩm mới rồi tải lên lại" className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border whitespace-nowrap" style={{ borderColor: BLUE, color: BLUE }}>
                            <ArrowUpFromLine size={15} /> Nhập từ Excel
                        </button>
                    )}
                    <input ref={importFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImportFile} />
                    {isAdmin && (
                        <button onClick={() => setManagingCategories(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border whitespace-nowrap" style={{ borderColor: LINE, color: INK }}>
                            <Filter size={15} /> Quản lý nhóm hàng
                        </button>
                    )}
                    {isAdmin && (
                        <button onClick={() => setManagingBrands(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border whitespace-nowrap" style={{ borderColor: LINE, color: INK }}>
                            <Filter size={15} /> Quản lý nhãn hiệu
                        </button>
                    )}
                    {!isCtv && (
                        <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white whitespace-nowrap" style={{ background: INK }}>
                            <Plus size={15} /> Thêm sản phẩm
                        </button>
                    )}
                </div>
            </div>

            <div className="flex flex-wrap items-end gap-2 mb-5">
                <label className="text-xs" style={{ width: 160 }}>
                    <span className="block opacity-60 mb-1">Nhóm hàng</span>
                    <select value={filterCategory} onChange={(e) => { setFilterCategory(e.target.value); if (filterBrand && !brandOptionsOf(e.target.value).includes(filterBrand)) setFilterBrand(""); }} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                        <option value="">Tất cả</option>
                        {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                </label>
                <label className="text-xs" style={{ width: 160 }}>
                    <span className="block opacity-60 mb-1">Nhãn hiệu</span>
                    <select value={filterBrand} onChange={(e) => setFilterBrand(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                        <option value="">Tất cả</option>
                        {(filterCategory ? brandOptionsOf(filterCategory) : brandOptions).map((b) => <option key={b} value={b}>{b}</option>)}
                    </select>
                </label>
                <label className="text-xs" style={{ width: 130 }}>
                    <span className="block opacity-60 mb-1">VAT</span>
                    <select value={filterVat} onChange={(e) => setFilterVat(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                        <option value="">Tất cả</option>
                        {VAT_OPTIONS.map((v) => <option key={v.id} value={v.id}>{v.label}</option>)}
                    </select>
                </label>
                {(filterCategory || filterBrand || filterVat) && (
                    <button onClick={() => { setFilterCategory(""); setFilterBrand(""); setFilterVat(""); }} className="text-xs opacity-50 hover:opacity-100 underline mb-1.5">Xoá lọc</button>
                )}
            </div>

            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 940 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        <th className="px-3 py-2.5 sticky top-0" style={{ background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}` }}><input type="checkbox" checked={selectedIds.size > 0 && selectedIds.size === filtered.length} onChange={toggleSelectAll} /></th>
                        {["", "", "Mã VT", "SKU", "Tên vật tư", "ĐVT", "Tồn ĐK", "Nhập", "Xuất", "Tồn CK", ...(isAdmin ? ["Giá nhập"] : []), ""].map((h, hi) => (
                            <th key={hi} className="text-left px-2 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap sticky top-0" style={h === "Tên vật tư" ? { color: INK, opacity: 0.6, background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}`, minWidth: 260 } : { color: INK, opacity: 0.6, background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}` }}>{h}</th>
                        ))}
                    </tr>
                    </thead>
                    <tbody>
                    {filtered.map((p) => {
                        const stats = productStats(p);
                        return (
                            <tr key={p.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                                <td className="px-3 py-3"><input type="checkbox" checked={selectedIds.has(p.id)} onChange={() => toggleSelect(p.id)} /></td>
                                <td className="px-2 py-3">
                                    <button onClick={() => openProductDetail(p.id)} className="opacity-50 hover:opacity-100" title="Xem chi tiết"><ChevronRight size={15} /></button>
                                </td>
                                <td className="px-2 py-3">
                                    {p.image ? (
                                        <img src={p.image} alt={p.name} onClick={() => setZoomImage({ images: [{ src: p.image, alt: p.name }, ...(p.images || []).map((im, i) => ({ src: im, alt: `${p.name} — ảnh phụ ${i + 1}` }))], index: 0 })} className="w-9 h-9 object-cover rounded-sm cursor-zoom-in" style={{ border: `1px solid ${LINE}` }} />
                                    ) : (
                                        <div className="w-9 h-9 rounded-sm flex items-center justify-center" style={{ background: PAPER, border: `1px dashed ${LINE}` }}>
                                            <ImageOff size={13} className="opacity-30" />
                                        </div>
                                    )}
                                </td>
                                <td className="px-3 py-3 font-medium">
                                    <button onClick={() => openProductDetail(p.id)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{p.code}</button>
                                </td>
                                <td className="px-3 py-3 opacity-70 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.sku || "—"}</td>
                                <td className="px-3 py-3" style={{ color: INK, minWidth: 260 }}>
                                    <button onClick={() => openProductDetail(p.id)} className="text-left hover:underline">{p.name}</button>
                                    <div className="flex gap-1.5 mt-1 flex-wrap">
                                        {p.isService && <span className="inline-block text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BLUE}1A`, color: BLUE }}>Dịch vụ</span>}
                                        {p.hasSeries && <span className="inline-flex items-center gap-1 text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BLUE}1A`, color: BLUE }}><Barcode size={10} /> Series</span>}
                                        {p.vat && <span className="inline-block text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BRASS}1A`, color: BRASS }}>{VAT_OPTIONS.find((v) => v.id === p.vat)?.label || p.vat}</span>}
                                        {p.variantAttrs && <span className="inline-block text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${PURPLE}1A`, color: PURPLE }}>{Object.values(p.variantAttrs).join(" / ")}</span>}
                                    </div>
                                </td>
                                <td className="px-2 py-3 opacity-70 whitespace-nowrap">{p.unit}</td>
                                {p.isService ? (
                                    <td colSpan={3} className="px-2 py-3 text-center opacity-40 text-xs">— không quản lý tồn kho —</td>
                                ) : (<>
                                    <td className="px-2 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.openingQty}</td>
                                    <td className="px-2 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>+{stats.importedQty}</td>
                                    <td className="px-2 py-3 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>-{stats.exportedQty}</td>
                                </>)}
                                <td className="px-2 py-3 text-right font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: stats.closingQty < 0 ? "#fff" : (stats.closingQty <= (p.minStockLevel ?? 5) ? RUST : INK) }}>
                                    {p.isService ? "—" : stats.closingQty < 0 ? (
                                        <span className="px-1.5 py-0.5 rounded-sm" style={{ background: RUST }}>{stats.closingQty}</span>
                                    ) : stats.closingQty}
                                </td>
                                {isAdmin && <td className="px-2 py-3 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.isService ? "—" : vnd(stats.avgCost)}</td>}
                                <td className="px-2 py-3">
                                    <div style={{ display: "grid", gridTemplateColumns: "auto auto", gap: 2, justifyContent: "end", marginLeft: "auto", width: "fit-content" }}>
                                        {isAdmin && !p.isService && <button onClick={() => openIO(p, "in")} title="Nhập kho" className="rounded-sm hover:bg-black/5" style={{ color: FOREST, padding: 4 }}><ArrowDownToLine size={13} /></button>}
                                        {!isCtv && !p.isService && <button onClick={() => openIO(p, "out")} title="Xuất kho" className="rounded-sm hover:bg-black/5" style={{ color: RUST, padding: 4 }}><ArrowUpFromLine size={13} /></button>}
                                        {isAdmin && (
                                            <>
                                                <button onClick={() => openEdit(p)} title="Sửa" className="rounded-sm hover:bg-black/5" style={{ opacity: 0.6, padding: 4 }}><Pencil size={13} /></button>
                                                <button onClick={() => removeProduct(p.id)} title="Xoá" className="rounded-sm hover:bg-black/5" style={{ color: RUST, opacity: 0.6, padding: 4 }}><Trash2 size={13} /></button>
                                            </>
                                        )}
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
                const galleryImgs = [
                    ...(viewingProduct.image ? [{ src: viewingProduct.image, alt: viewingProduct.name }] : []),
                    ...(viewingProduct.images || []).map((im, i) => ({ src: im, alt: `${viewingProduct.name} — ảnh phụ ${i + 1}` })),
                ];
                return (
                    <Modal title="Chi tiết sản phẩm" onClose={() => setViewingId(null)} size="2xl">
                        <div className="flex items-start gap-5 mb-3">
                            {viewingProduct.image ? (
                                <img src={viewingProduct.image} alt={viewingProduct.name} onClick={() => setZoomImage({ images: galleryImgs, index: 0 })}
                                     className="w-36 h-36 object-cover rounded-sm shrink-0 cursor-zoom-in" style={{ border: `1px solid ${LINE}` }} />
                            ) : (
                                <div className="w-36 h-36 rounded-sm flex items-center justify-center shrink-0" style={{ background: PAPER, border: `1px dashed ${LINE}` }}>
                                    <ImageOff size={30} className="opacity-30" />
                                </div>
                            )}
                            <div className="min-w-0">
                                <h4 style={{ fontFamily: "'Fraunces', serif", color: INK }} className="text-2xl leading-snug font-semibold">{viewingProduct.name}</h4>
                                {isAdmin && (() => {
                                    const sup = (suppliers || []).find((s) => s.id === viewingProduct.supplierId);
                                    return sup ? (
                                        <button onClick={() => goToSupplier && goToSupplier(sup.id)} className="inline-flex items-center gap-1 text-xs mt-1 hover:underline" style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>
                                            <Truck size={12} /> NCC: {sup.code}
                                        </button>
                                    ) : (
                                        <p className="text-xs mt-1 opacity-40">Chưa gán nhà cung cấp</p>
                                    );
                                })()}
                                <div className="flex flex-wrap gap-1.5 mt-2.5">
                                    <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: PAPER, color: INK, fontFamily: "'IBM Plex Mono', monospace", border: `1px solid ${LINE}` }}>{viewingProduct.code}</span>
                                    <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: PAPER, color: INK, fontFamily: "'IBM Plex Mono', monospace", border: `1px solid ${LINE}` }}>SKU {viewingProduct.sku || "—"}</span>
                                    {viewingProduct.hasSeries && <span className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${BLUE}1A`, color: BLUE }}><Barcode size={11} /> Series</span>}
                                    <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${BRASS}1A`, color: BRASS }}>{VAT_OPTIONS.find((v) => v.id === viewingProduct.vat)?.label || viewingProduct.vat}</span>
                                    {viewingProduct.warrantyMonths !== 0 && <span className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${FOREST}1A`, color: FOREST }}><ShieldCheck size={11} /> BH {warrantyLabel(viewingProduct.warrantyMonths)}</span>}
                                    {viewingProduct.brand && <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${PURPLE}1A`, color: PURPLE }}>{viewingProduct.brand}</span>}
                                    {viewingProduct.variantAttrs && <span className="text-[11px] px-2 py-0.5 rounded-sm" style={{ background: `${PURPLE}1A`, color: PURPLE }}>Phiên bản: {Object.entries(viewingProduct.variantAttrs).map(([k, v]) => `${k}: ${v}`).join(" · ")}</span>}
                                </div>
                                <p className="text-base font-medium mt-2.5" style={{ color: INK }}>{viewingProduct.category || "Chưa phân nhóm"}</p>
                                <p className="text-xs opacity-50 mt-1">{viewingProduct.unit}{viewingProduct.barcode ? ` · Barcode: ${viewingProduct.barcode}` : ""}</p>
                            </div>
                        </div>

                        {viewingProduct.images && viewingProduct.images.length > 0 && (
                            <div className="flex gap-2 mb-5">
                                {viewingProduct.images.map((img, i) => (
                                    <img key={i} src={img} alt={`${viewingProduct.name} — ảnh phụ ${i + 1}`} onClick={() => setZoomImage({ images: galleryImgs, index: i + 1 })}
                                         className="w-16 h-16 object-cover rounded-sm cursor-zoom-in" style={{ border: `1px solid ${LINE}` }} />
                                ))}
                            </div>
                        )}

                        {viewingProduct.variantGroupId && (() => {
                            const siblings = products.filter((p) => p.variantGroupId === viewingProduct.variantGroupId);
                            if (siblings.length <= 1) return null;
                            return (
                                <div className="mb-5 p-3 rounded-sm" style={{ background: `${PURPLE}08`, border: `1px solid ${PURPLE}33` }}>
                                    <p className="text-xs uppercase tracking-wider mb-2" style={{ color: PURPLE }}>Các phiên bản khác ({siblings.length})</p>
                                    <div className="flex flex-wrap gap-2">
                                        {siblings.map((sib) => {
                                            const sibStats = productStats(sib);
                                            const isCurrent = sib.id === viewingProduct.id;
                                            return (
                                                <button key={sib.id} onClick={() => !isCurrent && openProductDetail(sib.id)} disabled={isCurrent}
                                                        className="text-left px-3 py-2 rounded-sm text-xs"
                                                        style={{ border: `1px solid ${isCurrent ? PURPLE : LINE}`, background: isCurrent ? `${PURPLE}1A` : "#fff", cursor: isCurrent ? "default" : "pointer" }}>
                                                    <p className="font-medium" style={{ color: isCurrent ? PURPLE : INK }}>{sib.variantAttrs ? Object.values(sib.variantAttrs).join(" / ") : sib.name}</p>
                                                    <p className="opacity-50 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{sib.code} · Tồn {sibStats.closingQty}</p>
                                                </button>
                                            );
                                        })}
                                    </div>
                                </div>
                            );
                        })()}

                        <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-7 gap-2 mb-5">
                            {[
                                ["Tồn đầu kỳ", viewingProduct.openingQty, INK],
                                ["Nhập từ NCC", `+${stats.importedFromSupplierQty}`, FOREST],
                                ["Nhập lại (đổi trả)", `+${stats.importedFromReturnQty}`, BLUE],
                                ["Xuất trong kỳ", `-${stats.exportedQty}`, RUST],
                                ["Tồn cuối kỳ", stats.closingQty, stats.closingQty <= (viewingProduct.minStockLevel ?? 5) ? RUST : INK],
                                ...(isAdmin ? [["Giá nhập", vnd(viewingProduct.costPrice), INK]] : []),
                                ["Giá bán sỉ", vnd(viewingProduct.wholesalePrice), INK],
                                ["Giá bán lẻ", vnd(viewingProduct.retailPrice), INK],
                            ].map(([label, val, color], i) => (
                                <div key={i} className="p-2.5 rounded-sm text-center" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                    <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">{label}</p>
                                    <p className="text-sm font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color }}>{val}</p>
                                </div>
                            ))}
                        </div>

                        {isAdmin && viewingProduct.priceHistory.length > 0 && (
                            <div className="mb-5">
                                <button onClick={() => setShowPriceHistory((v) => !v)} className="text-xs underline opacity-60 hover:opacity-100 mb-2" style={{ color: INK }}>
                                    {showPriceHistory ? "Ẩn" : "Xem"} lịch sử thay đổi giá ({viewingProduct.priceHistory.length})
                                </button>
                                {showPriceHistory && (
                                    <div className="rounded-sm overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
                                        {viewingProduct.priceHistory.map((h) => (
                                            <div key={h.id} className="flex items-center justify-between gap-3 px-3 py-2 text-xs flex-wrap" style={{ borderBottom: `1px dashed ${LINE}` }}>
                        <span style={{ color: INK }}>
                          <b>{h.field}</b>: {vnd(h.oldValue)} → <b style={{ color: h.newValue > h.oldValue ? RUST : FOREST }}>{vnd(h.newValue)}</b>
                        </span>
                                                <span className="opacity-50 whitespace-nowrap">{formatDateTime(h.date)} · {h.changedBy}</span>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        )}

                        <div className="flex gap-2 mb-5 flex-wrap">
                            {isAdmin && <button onClick={() => openIOFromDetail(viewingProduct, "in")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm text-white" style={{ background: FOREST }}><ArrowDownToLine size={14} /> Nhập kho</button>}
                            {!isCtv && <button onClick={() => openIOFromDetail(viewingProduct, "out")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm text-white" style={{ background: RUST }}><ArrowUpFromLine size={14} /> Xuất kho</button>}
                            {isAdmin && (
                                <button onClick={() => openEditFromDetail(viewingProduct)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}><Pencil size={14} /> Sửa thông tin</button>
                            )}
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
                                                    <td className="py-1.5 px-2">{r.importDoc && goToDoc ? <button onClick={() => goToDoc(r.importDoc)} className="hover:underline" style={{ color: BLUE }}>{r.importDoc}</button> : (r.importDoc || "—")}</td>
                                                    <td className="py-1.5 px-2">{r.importDate}</td>
                                                    <td className="py-1.5 px-2">{r.exportDoc && goToDoc ? <button onClick={() => goToDoc(r.exportDoc)} className="hover:underline" style={{ color: BLUE }}>{r.exportDoc}</button> : (r.exportDoc || "—")}</td>
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

                        <div className="mt-5">
                            <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Lịch sử nhập / xuất / trả</p>
                            {viewingProduct.movements.length === 0 ? <p className="text-sm opacity-50">Chưa có bút toán nào.</p> : (() => {
                                const allSorted = [...viewingProduct.movements].sort((a, b) => (a.date < b.date ? 1 : -1));
                                const filteredHistory = allSorted.filter((m) => {
                                    if (historySearch.trim() && !(m.docNo || "").toLowerCase().includes(historySearch.trim().toLowerCase())) return false;
                                    if (historyFrom && m.date < historyFrom) return false;
                                    if (historyTo && m.date > historyTo) return false;
                                    return true;
                                });
                                const totalPages = Math.max(1, Math.ceil(filteredHistory.length / HISTORY_PAGE_SIZE));
                                const page = Math.min(historyPage, totalPages);
                                const pageRows = filteredHistory.slice((page - 1) * HISTORY_PAGE_SIZE, page * HISTORY_PAGE_SIZE);
                                return (
                                    <>
                                        <div className="flex flex-wrap items-end gap-2 mb-2">
                                            <label className="text-xs" style={{ width: 160 }}>
                                                <span className="block opacity-60 mb-1">Tìm số phiếu</span>
                                                <input value={historySearch} onChange={(e) => { setHistorySearch(e.target.value); setHistoryPage(1); }} placeholder="VD: POH001"
                                                       className="w-full border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }} />
                                            </label>
                                            <label className="text-xs" style={{ width: 140 }}>
                                                <span className="block opacity-60 mb-1">Từ ngày</span>
                                                <input type="date" value={historyFrom} onChange={(e) => { setHistoryFrom(e.target.value); setHistoryPage(1); }} className="w-full border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }} />
                                            </label>
                                            <label className="text-xs" style={{ width: 140 }}>
                                                <span className="block opacity-60 mb-1">Đến ngày</span>
                                                <input type="date" value={historyTo} onChange={(e) => { setHistoryTo(e.target.value); setHistoryPage(1); }} className="w-full border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }} />
                                            </label>
                                            {(historySearch || historyFrom || historyTo) && (
                                                <button onClick={() => { setHistorySearch(""); setHistoryFrom(""); setHistoryTo(""); setHistoryPage(1); }} className="text-xs opacity-50 hover:opacity-100 underline mb-1.5">Xoá lọc</button>
                                            )}
                                        </div>
                                        <div className="rounded-sm overflow-x-auto min-w-0" style={{ border: `1px solid ${LINE}` }}>
                                            <table className="w-full text-xs" style={{ minWidth: 460 }}>
                                                <thead style={{ background: PAPER }}><tr className="opacity-60">
                                                    <th className="text-left py-2 px-2">Loại</th><th className="text-left py-2 px-2">Số phiếu</th><th className="text-left py-2 px-2">Ngày</th><th className="text-right py-2 px-2">SL</th><th className="text-right py-2 px-2">Đơn giá</th>
                                                </tr></thead>
                                                <tbody>
                                                {pageRows.map((m) => {
                                                    const isReturn = (m.docNo || "").includes("-RT");
                                                    const isPO = (m.docNo || "").startsWith("POH");
                                                    const isDH = (m.docNo || "").startsWith("DH") && !isReturn;
                                                    let label = m.type === "in" ? "Nhập kho" : "Xuất kho";
                                                    if (isReturn) label = m.type === "in" ? "Nhập lại (đổi trả)" : "Xuất (đổi hàng)";
                                                    else if (isPO) label = "Nhập hàng";
                                                    else if (isDH) label = "Xuất bán";
                                                    return (
                                                        <tr key={m.id} style={{ borderTop: `1px dashed ${LINE}` }}>
                                                            <td className="py-1.5 px-2"><span style={{ color: m.type === "in" ? FOREST : RUST }}>{label}</span></td>
                                                            <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                                                                {m.docNo && goToDoc ? <button onClick={() => goToDoc(m.docNo)} className="hover:underline" style={{ color: BLUE }}>{m.docNo}</button> : <span style={{ color: BLUE }}>{m.docNo || "—"}</span>}
                                                            </td>
                                                            <td className="py-1.5 px-2 whitespace-nowrap">{m.date}</td>
                                                            <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: m.type === "in" ? FOREST : RUST }}>{m.type === "in" ? "+" : "-"}{m.qty}</td>
                                                            <td className="py-1.5 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(m.price)}</td>
                                                        </tr>
                                                    );
                                                })}
                                                {pageRows.length === 0 && <tr><td colSpan={5} className="text-center py-6 opacity-40">Không có bút toán phù hợp bộ lọc.</td></tr>}
                                                </tbody>
                                            </table>
                                        </div>
                                        {totalPages > 1 && (
                                            <div className="flex items-center justify-between mt-2">
                                                <span className="text-xs opacity-50">Trang {page}/{totalPages} · {filteredHistory.length} dòng</span>
                                                <div className="flex gap-1.5">
                                                    <button onClick={() => setHistoryPage((p2) => Math.max(1, p2 - 1))} disabled={page <= 1} className="text-xs px-2.5 py-1 rounded-sm border disabled:opacity-30" style={{ borderColor: LINE, color: INK }}>‹ Trước</button>
                                                    <button onClick={() => setHistoryPage((p2) => Math.min(totalPages, p2 + 1))} disabled={page >= totalPages} className="text-xs px-2.5 py-1 rounded-sm border disabled:opacity-30" style={{ borderColor: LINE, color: INK }}>Sau ›</button>
                                                </div>
                                            </div>
                                        )}
                                    </>
                                );
                            })()}
                        </div>
                    </Modal>
                );
            })()}
            {editing !== null && (
                <Modal title={editing.id ? "Sửa thông tin sản phẩm" : "Thêm sản phẩm"} onClose={() => setEditing(null)} size="xl">
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                        <Field label="Mã VT"><input className={inputCls} style={{ borderColor: LINE }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} disabled={!!editing.id} /></Field>
                        <Field label="Mã SKU" hint="Tự sinh — sửa được">
                            <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} />
                        </Field>
                        {!form.isService && (
                            <Field label="Đơn vị tính">
                                <select className={inputCls} style={{ borderColor: LINE }} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })}>
                                    {UNITS.map((u) => <option key={u} value={u}>{u}</option>)}
                                </select>
                            </Field>
                        )}
                    </div>

                    <Field label="Tên vật tư"><input className={inputCls} style={{ borderColor: LINE }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>

                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-4">
                        <Field label="Nhóm hàng" hint={isAdmin ? "Quản lý danh sách nhóm hàng ở nút bên trên" : "Chỉ chọn được nhóm hàng do quản trị viên tạo sẵn"}>
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.category || ""} onChange={(e) => setForm({ ...form, category: e.target.value, brand: e.target.value !== form.category ? "" : form.brand })}>
                                <option value="">— Chưa chọn —</option>
                                {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
                            </select>
                        </Field>
                        <Field label="Nhãn hiệu" hint={!form.category ? "Chọn nhóm hàng trước" : isAdmin ? "Quản lý danh sách nhãn hiệu ở nút bên trên" : "Chỉ chọn được nhãn hiệu do quản trị viên tạo sẵn"}>
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.brand || ""} onChange={(e) => setForm({ ...form, brand: e.target.value })} disabled={!form.category}>
                                <option value="">— Chưa chọn —</option>
                                {brandOptionsOf(form.category).map((b) => <option key={b} value={b}>{b}</option>)}
                            </select>
                        </Field>
                        <Field label="Barcode" hint="Bỏ trống nếu không có">
                            <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.barcode} onChange={(e) => setForm({ ...form, barcode: e.target.value })} placeholder="893xxxxxxxxxx" />
                        </Field>
                    </div>

                    <div className={`grid grid-cols-2 ${isAdmin ? "sm:grid-cols-3" : ""} gap-4`}>
                        {isAdmin && <Field label="Giá nhập (đ)"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.costPrice} onChange={(v) => setForm({ ...form, costPrice: v })} /></Field>}
                        <Field label="Giá bán lẻ (đ)"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.retailPrice} onChange={(v) => setForm({ ...form, retailPrice: v })} /></Field>
                        <Field label="Giá bán sỉ (đ)"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.wholesalePrice} onChange={(v) => setForm({ ...form, wholesalePrice: v })} /></Field>
                    </div>

                    <label className="flex items-center gap-2 mb-1 text-sm p-2.5 rounded-sm" style={{ color: INK, background: form.isService ? `${BLUE}0D` : "transparent", border: `1px solid ${form.isService ? BLUE : LINE}` }}>
                        <input type="checkbox" checked={!!form.isService} onChange={(e) => setForm({ ...form, isService: e.target.checked, hasSeries: e.target.checked ? false : form.hasSeries, openingQty: e.target.checked ? 0 : form.openingQty })} />
                        Sản phẩm dịch vụ (VD: phí sửa chữa, phí IT Helpdesk...) — không quản lý tồn kho, không trừ kho khi bán
                    </label>

                    {!form.isService && (
                        <div className="grid grid-cols-2 gap-4 items-end">
                            <Field label="Tồn đầu kỳ (SL)"><input type="number" className={inputCls} style={{ borderColor: LINE }} value={form.openingQty} onChange={(e) => setForm({ ...form, openingQty: e.target.value })} /></Field>
                            <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: INK }}>
                                <input type="checkbox" checked={!!form.hasSeries} onChange={(e) => setForm({ ...form, hasSeries: e.target.checked })} />
                                Quản lý theo số series
                            </label>
                        </div>
                    )}

                    {isAdmin && (
                        <Field label="Nhà cung cấp" hint="Không bắt buộc — nơi thường nhập sản phẩm này">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.supplierId || ""} onChange={(e) => setForm({ ...form, supplierId: e.target.value })}>
                                <option value="">— Chưa gán —</option>
                                {(suppliers || []).map((s) => <option key={s.id} value={s.id}>{s.code} - {s.name}</option>)}
                            </select>
                        </Field>
                    )}

                    {isAdmin && (
                        <Field label="Định mức tồn (cảnh báo khi tồn ≤ mức này)">
                            <div className="flex items-center gap-2">
                                {[0, 2, 5, 10].map((n) => (
                                    <button key={n} type="button" onClick={() => setForm({ ...form, minStockLevel: n })}
                                            className="px-3 py-1.5 rounded-sm text-sm border"
                                            style={{ borderColor: Number(form.minStockLevel) === n ? INK : LINE, background: Number(form.minStockLevel) === n ? INK : "transparent", color: Number(form.minStockLevel) === n ? "#fff" : INK }}>
                                        {n}
                                    </button>
                                ))}
                                <input type="number" min={0} value={form.minStockLevel} onChange={(e) => setForm({ ...form, minStockLevel: e.target.value })} className="w-20 border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} placeholder="Khác" />
                            </div>
                        </Field>
                    )}

                    <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
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

                        <Field label="Bảo hành" hint="Tuỳ sản phẩm — hiện trên đơn hàng và biên bản bàn giao khi in">
                            <div className="flex gap-2 flex-wrap">
                                <button type="button" onClick={() => setForm({ ...form, warrantyMonths: 0 })}
                                        className="px-3 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: !form.warrantyMonths ? INK : LINE, background: !form.warrantyMonths ? INK : "transparent", color: !form.warrantyMonths ? "#fff" : INK }}>
                                    Không BH
                                </button>
                                {WARRANTY_OPTIONS.map((m) => (
                                    <button key={m} type="button" onClick={() => setForm({ ...form, warrantyMonths: m })}
                                            className="px-3 py-1.5 rounded-sm text-sm border"
                                            style={{ borderColor: Number(form.warrantyMonths) === m ? INK : LINE, background: Number(form.warrantyMonths) === m ? INK : "transparent", color: Number(form.warrantyMonths) === m ? "#fff" : INK }}>
                                        {m} tháng
                                    </button>
                                ))}
                                <button type="button" onClick={() => setForm({ ...form, warrantyMonths: WARRANTY_LIFETIME })}
                                        className="px-3 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: Number(form.warrantyMonths) === WARRANTY_LIFETIME ? INK : LINE, background: Number(form.warrantyMonths) === WARRANTY_LIFETIME ? INK : "transparent", color: Number(form.warrantyMonths) === WARRANTY_LIFETIME ? "#fff" : INK }}>
                                    Vĩnh viễn
                                </button>
                            </div>
                        </Field>

                        <Field label="Hình ảnh sản phẩm" hint="Ảnh chính + tối đa 3 ảnh phụ — mỗi ảnh dưới 1.5MB">
                            <div className="flex flex-wrap items-start gap-3">
                                <div>
                                    <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Ảnh chính</p>
                                    {form.image ? (
                                        <div className="relative">
                                            <img src={form.image} alt="Ảnh chính" onClick={() => setZoomImage({ images: [{ src: form.image, alt: "Ảnh chính" }, ...(form.images || []).map((im, i) => ({ src: im, alt: `Ảnh phụ ${i + 1}` }))], index: 0 })}
                                                 className="w-16 h-16 object-cover rounded-sm cursor-zoom-in" style={{ border: `1px solid ${LINE}` }} />
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
                                </div>
                                {[0, 1, 2].map((i) => (
                                    <div key={i}>
                                        <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Ảnh phụ {i + 1}</p>
                                        {(form.images || [])[i] ? (
                                            <div className="relative">
                                                <img src={form.images[i]} alt={`Ảnh phụ ${i + 1}`} onClick={() => setZoomImage({ images: [{ src: form.image, alt: "Ảnh chính" }, ...(form.images || []).map((im, idx) => ({ src: im, alt: `Ảnh phụ ${idx + 1}` }))].filter((x) => x.src), index: (form.image ? 1 : 0) + i })}
                                                     className="w-16 h-16 object-cover rounded-sm cursor-zoom-in" style={{ border: `1px solid ${LINE}` }} />
                                                <button type="button" onClick={() => removeGalleryImage(i)}
                                                        className="absolute -top-2 -right-2 w-5 h-5 rounded-full flex items-center justify-center text-white" style={{ background: RUST }}>
                                                    <X size={11} />
                                                </button>
                                            </div>
                                        ) : (
                                            <label className="w-16 h-16 rounded-sm flex flex-col items-center justify-center cursor-pointer gap-0.5"
                                                   style={{ background: PAPER, border: `1px dashed ${LINE}`, color: INK }}>
                                                <ImagePlus size={16} className="opacity-50" />
                                                <input type="file" accept="image/*" className="hidden" onChange={(e) => onGalleryImagePick(i, e)} />
                                            </label>
                                        )}
                                    </div>
                                ))}
                            </div>
                            <span className="text-xs opacity-50 block mt-1.5">Bấm vào ảnh để phóng to, bấm ✕ để xoá / đổi ảnh khác.</span>
                        </Field>
                    </div>

                    {!editing.id && (
                        <>
                            <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                            <label className="flex items-center gap-2 mb-3 text-sm" style={{ color: INK }}>
                                <input type="checkbox" checked={!!form.hasVariants} onChange={(e) => setForm({ ...form, hasVariants: e.target.checked })} />
                                Sản phẩm có nhiều phiên bản (màu sắc, kích cỡ…)
                            </label>

                            {form.hasVariants && (
                                <div className="p-3 rounded-sm mb-3" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                    <p className="text-xs opacity-60 mb-3">Mỗi tổ hợp giá trị sẽ tạo thành 1 sản phẩm riêng (Mã VT/SKU tự thêm hậu tố, vd <code>{form.code || "HI003"}_BK</code> cho màu Đen). Thông tin nền (nhóm hàng, nhãn hiệu, giá, ảnh, VAT, bảo hành…) ở trên sẽ dùng chung cho mọi phiên bản, sau đó vẫn sửa lại được từng phiên bản riêng.</p>

                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-1">
                                        <Field label="Tên thuộc tính 1">
                                            <input className={inputCls} style={{ borderColor: LINE, background: "#fff" }} value={form.variantAttr1Name} onChange={(e) => setForm({ ...form, variantAttr1Name: e.target.value })} placeholder="VD: Màu sắc" />
                                        </Field>
                                        <Field label={`Giá trị "${form.variantAttr1Name || "Thuộc tính 1"}"`}>
                                            <VariantValueTagInput values={form.variantAttr1Values || []} setValues={(arr) => setForm({ ...form, variantAttr1Values: arr })} placeholder="VD: Đen, Trắng, Xanh dương…" />
                                        </Field>
                                    </div>

                                    <label className="flex items-center gap-2 my-3 text-sm" style={{ color: INK }}>
                                        <input type="checkbox" checked={!!form.variantAttr2Enabled} onChange={(e) => setForm({ ...form, variantAttr2Enabled: e.target.checked })} />
                                        Thêm thuộc tính thứ 2 (vd: kích cỡ)
                                    </label>

                                    {form.variantAttr2Enabled && (
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-1">
                                            <Field label="Tên thuộc tính 2">
                                                <input className={inputCls} style={{ borderColor: LINE, background: "#fff" }} value={form.variantAttr2Name} onChange={(e) => setForm({ ...form, variantAttr2Name: e.target.value })} placeholder="VD: Kích cỡ" />
                                            </Field>
                                            <Field label={`Giá trị "${form.variantAttr2Name || "Thuộc tính 2"}"`}>
                                                <VariantValueTagInput values={form.variantAttr2Values || []} setValues={(arr) => setForm({ ...form, variantAttr2Values: arr })} placeholder="VD: S, M, L…" />
                                            </Field>
                                        </div>
                                    )}

                                    {(form.variantAttr1Values || []).length > 0 && (
                                        <p className="text-xs mt-2" style={{ color: FOREST }}>
                                            Sẽ tạo <b>{(form.variantAttr1Values || []).length * (form.variantAttr2Enabled ? Math.max((form.variantAttr2Values || []).length, 1) : 1)}</b> phiên bản sản phẩm.
                                        </p>
                                    )}
                                </div>
                            )}
                        </>
                    )}

                    <button onClick={submitInfo} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>
                        {editing.id ? "Lưu thay đổi" : form.hasVariants ? "Tạo các phiên bản sản phẩm" : "Thêm sản phẩm"}
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
                            <MoneyInput className={inputCls} style={{ borderColor: LINE }} value={ioForm.price} onChange={(v) => setIoForm({ ...ioForm, price: v })} />
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

            {importResult && (
                <Modal title="Kết quả nhập sản phẩm từ Excel" onClose={() => setImportResult(null)}>
                    <div className="p-3 rounded-sm mb-3" style={{ background: `${FOREST}0D`, border: `1px solid ${FOREST}44` }}>
                        <p className="text-sm" style={{ color: FOREST }}>Đã thêm <b>{importResult.added}</b> sản phẩm mới.</p>
                    </div>
                    {importResult.skipped.length > 0 && (
                        <div className="p-3 rounded-sm" style={{ background: `${BRASS}0D`, border: `1px solid ${BRASS}44` }}>
                            <p className="text-sm mb-1.5" style={{ color: BRASS }}>Bỏ qua {importResult.skipped.length} mã VT đã tồn tại (không ghi đè):</p>
                            <p className="text-xs opacity-70" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{importResult.skipped.join(", ")}</p>
                        </div>
                    )}
                </Modal>
            )}

            {managingCategories && (
                <Modal title="Quản lý nhóm hàng" onClose={() => { setManagingCategories(false); setNewCategoryInput(""); }}>
                    <p className="text-xs opacity-60 mb-3">Đây là danh sách nhóm hàng dùng chung cho toàn bộ sản phẩm — chỉ quản trị viên tạo/xoá được, các sản phẩm chỉ chọn từ danh sách này.</p>
                    <div className="flex gap-2 mb-4">
                        <input value={newCategoryInput} onChange={(e) => setNewCategoryInput(e.target.value)} placeholder="Tên nhóm hàng mới…"
                               className="flex-1 border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }}
                               onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); const v = newCategoryInput.trim(); if (v && !(categories || []).includes(v)) { setCategories([...(categories || []), v]); addLog("Thêm nhóm hàng", v); } setNewCategoryInput(""); } }} />
                        <button onClick={() => { const v = newCategoryInput.trim(); if (v && !(categories || []).includes(v)) { setCategories([...(categories || []), v]); addLog("Thêm nhóm hàng", v); } setNewCategoryInput(""); }}
                                className="px-4 py-2 rounded-sm text-white text-sm shrink-0" style={{ background: INK }}>Thêm</button>
                    </div>
                    {categoryOptions.length === 0 ? (
                        <p className="text-sm opacity-50 text-center py-8">Chưa có nhóm hàng nào.</p>
                    ) : (
                        <div className="space-y-1.5 max-h-[50vh] overflow-y-auto">
                            {categoryOptions.map((c) => {
                                const usedCount = products.filter((p) => p.category === c).length;
                                const brandCount = (brands || []).filter((b) => b.category === c).length;
                                return (
                                    <div key={c} className="flex items-center justify-between p-2.5 rounded-sm" style={{ border: `1px solid ${LINE}` }}>
                                        <div>
                                            <span className="text-sm" style={{ color: INK }}>{c}</span>
                                            <span className="text-xs opacity-50 ml-2">{usedCount > 0 ? `${usedCount} sản phẩm` : "Chưa có sản phẩm"}{brandCount > 0 ? ` · ${brandCount} nhãn hiệu` : ""}</span>
                                        </div>
                                        <button
                                            onClick={() => {
                                                if (usedCount > 0) { alert(`Không thể xoá — đang có ${usedCount} sản phẩm thuộc nhóm hàng này. Hãy đổi nhóm hàng cho các sản phẩm đó trước.`); return; }
                                                if (brandCount > 0) { alert(`Không thể xoá — nhóm hàng này đang có ${brandCount} nhãn hiệu con. Hãy xoá hoặc chuyển các nhãn hiệu đó sang nhóm khác trước (mục Quản lý nhãn hiệu).`); return; }
                                                setCategories((categories || []).filter((x) => x !== c));
                                                addLog("Xoá nhóm hàng", c);
                                            }}
                                            className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}>
                                            <Trash2 size={14} />
                                        </button>
                                    </div>
                                );
                            })}
                        </div>
                    )}
                </Modal>
            )}

            {managingBrands && (
                <Modal title="Quản lý nhãn hiệu" onClose={() => { setManagingBrands(false); setNewBrandInputByCat({}); }} size="lg">
                    <p className="text-xs opacity-60 mb-3">Mỗi nhãn hiệu thuộc về 1 nhóm hàng cụ thể (vd nhóm <b>Phần mềm</b> có nhãn MS, Adobe…; nhóm <b>Gia dụng</b> có Ergotek, NB…). Khi tạo sản phẩm, chọn nhóm hàng trước rồi mới chọn được đúng nhãn hiệu trong nhóm đó.</p>

                    {categoryOptions.length === 0 ? (
                        <p className="text-sm opacity-50 text-center py-8">Chưa có nhóm hàng nào — vào "Quản lý nhóm hàng" để tạo trước, sau đó mới thêm được nhãn hiệu theo từng nhóm.</p>
                    ) : (
                        <div className="space-y-4 max-h-[60vh] overflow-y-auto">
                            {categoryOptions.map((cat) => {
                                const catBrands = (brands || []).filter((b) => b.category === cat).sort((a, b) => a.name.localeCompare(b.name));
                                const addBrand = () => {
                                    const v = (newBrandInputByCat[cat] || "").trim();
                                    if (!v) return;
                                    if (catBrands.some((b) => b.name.toLowerCase() === v.toLowerCase())) { setNewBrandInputByCat({ ...newBrandInputByCat, [cat]: "" }); return; }
                                    setBrands([...(brands || []), { id: uid(), name: v, category: cat }]);
                                    addLog("Thêm nhãn hiệu", `${v} (${cat})`);
                                    setNewBrandInputByCat({ ...newBrandInputByCat, [cat]: "" });
                                };
                                return (
                                    <div key={cat} className="p-3 rounded-sm" style={{ border: `1px solid ${LINE}` }}>
                                        <p className="text-sm font-medium mb-2" style={{ color: INK }}>{cat}</p>
                                        <div className="flex flex-wrap gap-1.5 mb-2">
                                            {catBrands.length === 0 && <span className="text-xs opacity-40">Chưa có nhãn hiệu nào trong nhóm này.</span>}
                                            {catBrands.map((b) => {
                                                const usedCount = products.filter((p) => p.brand === b.name).length;
                                                return (
                                                    <span key={b.id} className="inline-flex items-center gap-1.5 pl-2.5 pr-1.5 py-1 rounded-full text-xs" style={{ background: `${PURPLE}17`, color: PURPLE }} title={usedCount > 0 ? `${usedCount} sản phẩm đang dùng` : "Chưa có sản phẩm nào dùng"}>
                            {b.name}
                                                        <button type="button" onClick={() => {
                                                            if (usedCount > 0) { alert(`Không thể xoá — đang có ${usedCount} sản phẩm thuộc nhãn hiệu "${b.name}". Hãy đổi nhãn hiệu cho các sản phẩm đó trước.`); return; }
                                                            setBrands((brands || []).filter((x) => x.id !== b.id));
                                                            addLog("Xoá nhãn hiệu", `${b.name} (${cat})`);
                                                        }} className="hover:opacity-60 rounded-full" style={{ padding: 2 }}><X size={11} /></button>
                          </span>
                                                );
                                            })}
                                        </div>
                                        <div className="flex gap-2">
                                            <input value={newBrandInputByCat[cat] || ""} onChange={(e) => setNewBrandInputByCat({ ...newBrandInputByCat, [cat]: e.target.value })}
                                                   onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); addBrand(); } }}
                                                   placeholder={`Thêm nhãn hiệu vào "${cat}"…`} className="flex-1 border rounded-sm py-1.5 px-2.5 text-sm" style={{ borderColor: LINE }} />
                                            <button onClick={addBrand} className="px-3 py-1.5 rounded-sm text-white text-xs shrink-0" style={{ background: INK }}>Thêm</button>
                                        </div>
                                    </div>
                                );
                            })}

                            {(() => {
                                const unassigned = (brands || []).filter((b) => !b.category || !categoryOptions.includes(b.category));
                                if (unassigned.length === 0) return null;
                                return (
                                    <div className="p-3 rounded-sm" style={{ border: `1px dashed ${RUST}66` }}>
                                        <p className="text-sm font-medium mb-2" style={{ color: RUST }}>Chưa gán nhóm hàng ({unassigned.length})</p>
                                        <div className="space-y-1.5">
                                            {unassigned.map((b) => (
                                                <div key={b.id} className="flex items-center gap-2">
                                                    <span className="text-sm flex-1" style={{ color: INK }}>{b.name}</span>
                                                    <select onChange={(e) => { if (e.target.value) { setBrands((brands || []).map((x) => (x.id === b.id ? { ...x, category: e.target.value } : x))); addLog("Gán nhóm hàng cho nhãn hiệu", `${b.name} → ${e.target.value}`); } }}
                                                            defaultValue="" className="border rounded-sm py-1 px-2 text-xs" style={{ borderColor: LINE }}>
                                                        <option value="">— Gán vào nhóm hàng —</option>
                                                        {categoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
                                                    </select>
                                                    <button onClick={() => {
                                                        const usedCount = products.filter((p) => p.brand === b.name).length;
                                                        if (usedCount > 0) { alert(`Không thể xoá — đang có ${usedCount} sản phẩm thuộc nhãn hiệu "${b.name}".`); return; }
                                                        setBrands((brands || []).filter((x) => x.id !== b.id));
                                                    }} className="p-1 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={13} /></button>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                );
                            })()}
                        </div>
                    )}
                </Modal>
            )}

            {zoomImage && <ImageLightbox images={zoomImage.images} startIndex={zoomImage.index} onClose={() => setZoomImage(null)} />}
        </div>
    );
}

/* ---------------- Nhập hàng (Purchase Orders / Đơn nhập hàng) ---------------- */

// Nhập nhanh danh sách giá trị thuộc tính phiên bản (vd Đen, Trắng, Xanh dương) — gõ rồi Enter/dấu phẩy để thêm từng giá trị.
function VariantValueTagInput({ values, setValues, placeholder }) {
    const [text, setText] = useState("");
    const commit = () => {
        const v = text.trim().replace(/,$/, "");
        if (v && !values.some((x) => x.toLowerCase() === v.toLowerCase())) setValues([...values, v]);
        setText("");
    };
    return (
        <div className="w-full border rounded-sm p-2 flex flex-wrap gap-1.5 items-center" style={{ borderColor: LINE, background: "#fff" }}>
            {values.map((v, i) => (
                <span key={i} className="inline-flex items-center gap-1 pl-2.5 pr-1.5 py-1 rounded-full text-xs" style={{ background: `${BLUE}17`, color: BLUE }}>
          {v}
                    <button type="button" onClick={() => setValues(values.filter((_, idx) => idx !== i))} className="hover:opacity-60 rounded-full" style={{ padding: 2 }}><X size={11} /></button>
        </span>
            ))}
            <input value={text} onChange={(e) => { const val = e.target.value; if (/,$/.test(val)) { setText(val); commit(); } else setText(val); }}
                   onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); commit(); } else if (e.key === "Backspace" && text === "" && values.length > 0) setValues(values.slice(0, -1)); }}
                   onBlur={commit}
                   placeholder={placeholder || "Gõ giá trị rồi Enter…"}
                   className="flex-1 min-w-[120px] text-sm outline-none py-0.5" style={{ background: "transparent" }} />
        </div>
    );
}

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

function ProductPicker({ products, onPick, onQuickCreate, brands }) {
    const [query, setQuery] = useState("");
    const [open, setOpen] = useState(false);
    const [creatingNew, setCreatingNew] = useState(false);
    const [newForm, setNewForm] = useState({ code: "", name: "", unit: UNITS[0], brand: "", warrantyMonths: 0, vat: "VAT10" });
    const q = query.trim().toLowerCase();
    const matches = q
        ? products.filter((p) => p.name.toLowerCase().includes(q) || (p.code || "").toLowerCase().includes(q) || (p.sku || "").toLowerCase().includes(q)).slice(0, 60)
        : products.slice(0, 60);

    const startQuickCreate = () => { setNewForm({ code: "", name: query, unit: UNITS[0], brand: "", warrantyMonths: 0, vat: "VAT10" }); setCreatingNew(true); };
    const saveQuickCreate = () => {
        if (!newForm.code.trim() || !newForm.name.trim()) { alert("Vui lòng nhập Mã vật tư và Tên vật tư."); return; }
        const newProduct = onQuickCreate(newForm);
        if (newProduct) { onPick(newProduct.id, newProduct); setQuery(""); setOpen(false); setCreatingNew(false); }
    };

    return (
        <div className="relative">
            <div className="relative">
                <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 opacity-40" />
                <input
                    value={query}
                    onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
                    onFocus={() => setOpen(true)}
                    placeholder="Gõ tên hoặc mã sản phẩm để tìm…"
                    className="w-full pl-8 pr-8 py-2 text-sm rounded-sm border outline-none"
                    style={{ borderColor: LINE, background: "#fff" }}
                />
                {(open || creatingNew) && (
                    <button type="button" onClick={() => { setOpen(false); setCreatingNew(false); }} title="Đóng bảng gợi ý" className="absolute right-2 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-100" style={{ color: INK }}>
                        <X size={15} />
                    </button>
                )}
            </div>
            {(open || creatingNew) && (
                <div className="absolute z-20 mt-1 w-full max-h-[420px] overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    {creatingNew ? (
                        <div className="p-3">
                            <p className="text-xs uppercase tracking-wider opacity-50 mb-2">Tạo nhanh sản phẩm mới</p>
                            <div className="grid grid-cols-2 gap-2 mb-2">
                                <input value={newForm.code} onChange={(e) => setNewForm({ ...newForm, code: e.target.value })} placeholder="Mã vật tư *" className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                <select value={newForm.unit} onChange={(e) => setNewForm({ ...newForm, unit: e.target.value })} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                                    {UNITS.map((u) => <option key={u} value={u}>{u}</option>)}
                                </select>
                            </div>
                            <input value={newForm.name} onChange={(e) => setNewForm({ ...newForm, name: e.target.value })} placeholder="Tên vật tư *" className="w-full border rounded-sm py-1.5 px-2 text-sm mb-2" style={{ borderColor: LINE }} />
                            <select value={newForm.brand} onChange={(e) => setNewForm({ ...newForm, brand: e.target.value })} className="w-full border rounded-sm py-1.5 px-2 text-sm mb-2" style={{ borderColor: LINE }}>
                                <option value="">— Nhãn hiệu (không bắt buộc) —</option>
                                {[...new Set((brands || []).map((b) => (typeof b === "string" ? b : b.name)))].sort().map((b) => <option key={b} value={b}>{b}</option>)}
                            </select>
                            <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Bảo hành</p>
                            <div className="flex gap-1.5 flex-wrap mb-2">
                                <button type="button" onClick={() => setNewForm({ ...newForm, warrantyMonths: 0 })} className="px-2 py-1 rounded-sm text-xs border" style={{ borderColor: !newForm.warrantyMonths ? INK : LINE, background: !newForm.warrantyMonths ? INK : "transparent", color: !newForm.warrantyMonths ? "#fff" : INK }}>Không BH</button>
                                {WARRANTY_OPTIONS.map((m) => (
                                    <button key={m} type="button" onClick={() => setNewForm({ ...newForm, warrantyMonths: m })} className="px-2 py-1 rounded-sm text-xs border" style={{ borderColor: Number(newForm.warrantyMonths) === m ? INK : LINE, background: Number(newForm.warrantyMonths) === m ? INK : "transparent", color: Number(newForm.warrantyMonths) === m ? "#fff" : INK }}>{m}th</button>
                                ))}
                                <button type="button" onClick={() => setNewForm({ ...newForm, warrantyMonths: WARRANTY_LIFETIME })} className="px-2 py-1 rounded-sm text-xs border" style={{ borderColor: Number(newForm.warrantyMonths) === WARRANTY_LIFETIME ? INK : LINE, background: Number(newForm.warrantyMonths) === WARRANTY_LIFETIME ? INK : "transparent", color: Number(newForm.warrantyMonths) === WARRANTY_LIFETIME ? "#fff" : INK }}>Vĩnh viễn</button>
                            </div>
                            <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">VAT</p>
                            <div className="flex gap-1.5 flex-wrap mb-3">
                                {VAT_OPTIONS.map((v) => (
                                    <button key={v.id} type="button" onClick={() => setNewForm({ ...newForm, vat: v.id })} className="px-2 py-1 rounded-sm text-xs border" style={{ borderColor: newForm.vat === v.id ? INK : LINE, background: newForm.vat === v.id ? INK : "transparent", color: newForm.vat === v.id ? "#fff" : INK }}>{v.label}</button>
                                ))}
                            </div>
                            <div className="flex gap-2">
                                <button type="button" onClick={() => setCreatingNew(false)} className="flex-1 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}>Huỷ</button>
                                <button type="button" onClick={saveQuickCreate} className="flex-1 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>Tạo & thêm vào báo giá</button>
                            </div>
                        </div>
                    ) : (
                        <>
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
                            {onQuickCreate && (
                                <button type="button" onMouseDown={(e) => { e.preventDefault(); startQuickCreate(); }}
                                        className="w-full text-left px-3 py-2.5 text-sm flex items-center gap-1.5 font-medium" style={{ color: BLUE }}>
                                    <Plus size={14} /> Tạo sản phẩm mới nhanh{query ? `: "${query}"` : ""}
                                </button>
                            )}
                        </>
                    )}
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
                                        <input type="number" min={1} value={it.qty || ""} onChange={(e) => { const v = e.target.value; onUpdate(it.productId, { qty: v === "" ? "" : Math.max(0, Number(v)) }); }}
                                               onBlur={() => { if (!it.qty || it.qty < 1) onUpdate(it.productId, { qty: 1 }); }}
                                               className="w-full border rounded-sm py-2 px-1 text-center text-[15px] font-medium" style={{ borderColor: LINE }} />
                                    )}
                                </td>
                                <td className="px-1 py-3 align-top">
                                    {lockQtyPrice ? (
                                        <span className="block text-right text-[15px] font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(it.price)}</span>
                                    ) : (
                                        <MoneyInput value={it.price} onChange={(v) => onUpdate(it.productId, { price: v === "" ? 0 : v })}
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

function PurchaseOrders({ purchaseOrders, setPurchaseOrders, products, setProducts, suppliers, setSuppliers, employeeNames, addLog, focusPOId, onFocusHandled }) {
    const [creating, setCreating] = useState(false);
    const [form, setForm] = useState({});
    const [expanded, setExpanded] = useState(null);
    const [viewingId, setViewingId] = useState(null); // id của đơn đang xem/sửa
    const [editForm, setEditForm] = useState(null);
    const [returningPO, setReturningPO] = useState(null); // po đang lập phiếu trả hàng NCC
    const [returnForm, setReturnForm] = useState(null);
    const [quickAddingSupplier, setQuickAddingSupplier] = useState(false);
    const [quickSupplierForm, setQuickSupplierForm] = useState({});

    const openQuickAddSupplier = () => { setQuickSupplierForm({ code: nextSupplierCode(suppliers), name: "", taxCode: "", address: "", contactPerson: "", phone: "", email: "", paymentTerm: "cash", creditDays: 30 }); setQuickAddingSupplier(true); };
    const submitQuickAddSupplier = () => {
        if (!quickSupplierForm.name?.trim()) { alert("Vui lòng nhập tên nhà cung cấp."); return; }
        const newSupplier = { ...quickSupplierForm, id: uid(), code: quickSupplierForm.code || nextSupplierCode(suppliers), creditDays: Number(quickSupplierForm.creditDays) || 0 };
        setSuppliers((prev) => [...prev, newSupplier]);
        addLog("Thêm nhà cung cấp", newSupplier.name);
        // Nếu đang mở form tạo đơn nhập hàng, tự chọn luôn NCC vừa thêm cho tiện.
        setForm((f) => (f && f.supplierId !== undefined ? { ...f, supplierId: newSupplier.id, supplier: `${newSupplier.code} - ${newSupplier.name}` } : f));
        setQuickAddingSupplier(false);
    };

    const viewingPO = purchaseOrders.find((x) => x.id === viewingId) || null;

    // Số lượng của 1 sản phẩm trong đơn nhập này đã được trả lại NCC trước đó (cộng dồn các phiếu trả trước).
    const alreadyReturnedQty = (po, productId) => (po.returns || []).reduce((s, r) => s + r.items.filter((it) => it.productId === productId).reduce((s2, it) => s2 + it.qty, 0), 0);
    // Các số series thuộc đúng đơn nhập này, hiện vẫn còn tồn kho (chưa bán/chưa trả) — chỉ những series này mới được phép trả lại NCC.
    const availableSeriesFromPO = (po, productId) => {
        const p = products.find((x) => x.id === productId);
        if (!p) return [];
        return seriesList(p).filter((s) => s.status === "Còn tồn" && s.importDoc === po.code);
    };

    const openReturn = (po) => {
        setReturnForm({
            items: po.items.map((it) => ({ productId: it.productId, qty: 0, series: [], price: it.price })),
            note: "",
        });
        setReturningPO(po);
    };
    const updateReturnItem = (productId, patch) => setReturnForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
    const returnInvalid = () => {
        if (!returnForm || !returningPO) return true;
        const active = returnForm.items.filter((it) => it.qty > 0);
        if (active.length === 0) return true;
        for (const it of active) {
            const p = products.find((x) => x.id === it.productId);
            if (p?.hasSeries && it.series.length !== it.qty) return true;
            const stats = productStats(p);
            const already = alreadyReturnedQty(returningPO, it.productId);
            const orig = returningPO.items.find((x) => x.productId === it.productId)?.qty || 0;
            if (it.qty > orig - already) return true; // không trả quá số đã nhập (trừ đi phần đã trả trước đó)
            if (!p?.hasSeries && it.qty > stats.closingQty) return true; // không trả quá tồn kho hiện có (đã bán rồi thì không trả được nữa)
        }
        return false;
    };
    const submitReturn = () => {
        if (returnInvalid()) return;
        const code = nextSupplierReturnCode(returningPO);
        const now = new Date().toISOString();
        const items = returnForm.items.filter((it) => it.qty > 0);
        const rec = { id: uid(), code, createdAt: now, note: returnForm.note, items };
        setPurchaseOrders((prev) => prev.map((po) => (po.id === returningPO.id ? { ...po, returns: [...(po.returns || []), rec] } : po)));
        setProducts((prev) => prev.map((p) => {
            const it = items.find((i) => i.productId === p.id);
            if (!it) return p;
            return { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: code, date: todayISO(), qty: it.qty, price: it.price, series: it.series }] };
        }));
        const value = items.reduce((s, it) => s + it.qty * it.price, 0);
        addLog("Trả hàng NCC", `${code} · ${returningPO.supplier} · ${vnd(value)}${returnForm.note ? ` · Lý do: ${returnForm.note}` : ""}`);
        setReturningPO(null);
        setReturnForm(null);
    };

    const openNew = () => { setForm({ supplier: "", supplierId: "", branch: BRANCHES[0], createdBy: (employeeNames[0] || EMPLOYEES[0]), paymentMethod: "credit", creditDays: 30, invoiceNo: "", notes: "", tags: [], items: [] }); setCreating(true); };

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

    const validateItems = (itemsRaw, excludeDocNo) => {
        if (!itemsRaw || itemsRaw.length === 0) { alert("Chưa có sản phẩm nào trong đơn."); return null; }
        const built = [];
        for (const it of itemsRaw) {
            const p = products.find((x) => x.id === it.productId);
            if (p.hasSeries && it.series.length !== it.qty) {
                alert(`"${p.name}" quản lý theo series — cần đúng ${it.qty} số series (đang có ${it.series.length}). Vui lòng điền đủ số series trước khi nhập hàng.`);
                return null;
            }
            if (p.hasSeries && it.series.length > 0) {
                const dups = findDuplicateSeries(p, it.series, excludeDocNo);
                if (dups.length > 0) {
                    alert(`"${p.name}" có số series bị trùng với phiếu nhập trước đó (số series phải là duy nhất cho mỗi sản phẩm): ${dups.join(", ")}. Vui lòng kiểm tra lại.`);
                    return null;
                }
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
        if (!form.supplierId) { alert("Vui lòng chọn nhà cung cấp từ danh sách."); return; }
        const items = validateItems(form.items);
        if (!items) return;
        const now = new Date().toISOString();
        const paid = form.paymentMethod === "cash";
        const po = {
            id: uid(), code: nextPOCode(purchaseOrders), createdAt: now,
            status, receivedAt: status === "received" ? now : null,
            branch: form.branch, supplier: form.supplier, supplierId: form.supplierId, createdBy: form.createdBy,
            invoiceNo: form.invoiceNo || "", notes: form.notes || "", tags: form.tags || [],
            paymentMethod: form.paymentMethod, creditDays: form.paymentMethod === "credit" ? (Number(form.creditDays) || 0) : 0, paid, paidAt: paid ? now : null, items,
        };
        setPurchaseOrders((prev) => [po, ...prev]);
        if (status === "received") applyPOToStock(po, setProducts);
        addLog("Tạo đơn nhập hàng", `${po.code} · ${status === "received" ? "Đã nhập" : "Chờ giao"}`);
        setCreating(false);
    };

    const confirmReceive = (po) => {
        const receivedAt = new Date().toISOString();
        setPurchaseOrders((prev) => prev.map((x) => (x.id === po.id ? { ...x, status: "received", receivedAt } : x)));
        applyPOToStock(po, setProducts);
        addLog("Xác nhận nhập hàng", po.code);
    };
    const markPaid = (po) => {
        const paidAt = new Date().toISOString();
        setPurchaseOrders((prev) => prev.map((x) => (x.id === po.id ? { ...x, paid: true, paidAt } : x)));
    };

    // ----- Xem / sửa đơn đã tạo -----
    const openView = (po) => {
        setEditForm({
            supplier: po.supplier, supplierId: po.supplierId || "", branch: po.branch, createdBy: po.createdBy, paymentMethod: po.paymentMethod, creditDays: po.creditDays || 30,
            invoiceNo: po.invoiceNo || "", notes: po.notes || "", tags: [...(po.tags || [])],
            items: po.items.map((it) => ({ ...it, series: [...it.series] })),
        });
        setViewingId(po.id);
    };
    const closeView = () => { setViewingId(null); setEditForm(null); };

    useEffect(() => {
        if (focusPOId) {
            const po = purchaseOrders.find((x) => x.id === focusPOId);
            if (po) openView(po);
            onFocusHandled && onFocusHandled();
        }
    }, [focusPOId]);

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
            if (p?.hasSeries && it.series.length > 0) {
                const dups = findDuplicateSeries(p, it.series, viewingPO.code);
                if (dups.length > 0) {
                    alert(`"${p.name}" có số series bị trùng với phiếu nhập khác (số series phải là duy nhất cho mỗi sản phẩm): ${dups.join(", ")}.`);
                    return;
                }
            }
        }
        const items = editForm.items;
        setPurchaseOrders((prev) => prev.map((x) => {
            if (x.id !== viewingId) return x;
            return { ...x, supplier: editForm.supplier, supplierId: editForm.supplierId || "", branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, creditDays: editForm.paymentMethod === "credit" ? (Number(editForm.creditDays) || 0) : 0, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, items };
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
        setPurchaseOrders((prev) => prev.map((x) => (x.id === viewingId ? { ...x, supplier: editForm.supplier, supplierId: editForm.supplierId || "", branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, creditDays: editForm.paymentMethod === "credit" ? (Number(editForm.creditDays) || 0) : 0, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, paid, paidAt: paid ? (x.paidAt || new Date().toISOString()) : null, items } : x)));
        closeView();
    };
    const confirmReceiveFromEdit = () => {
        const items = validateItems(editForm.items);
        if (!items) return;
        const receivedAt = new Date().toISOString();
        const paid = editForm.paymentMethod === "cash";
        const updated = { ...viewingPO, supplier: editForm.supplier, supplierId: editForm.supplierId || "", branch: editForm.branch, createdBy: editForm.createdBy, paymentMethod: editForm.paymentMethod, creditDays: editForm.paymentMethod === "credit" ? (Number(editForm.creditDays) || 0) : 0, invoiceNo: editForm.invoiceNo, notes: editForm.notes, tags: editForm.tags, paid, paidAt: paid ? (viewingPO.paidAt || receivedAt) : null, items, status: "received", receivedAt };
        setPurchaseOrders((prev) => prev.map((x) => (x.id === viewingId ? updated : x)));
        applyPOToStock(updated, setProducts);
        closeView();
    };

    const dueSoonPOs = purchaseOrders.map((po) => ({ po, due: poDueInfo(po) })).filter((x) => x.due && x.due.nearDue);

    return (
        <div>
            {dueSoonPOs.length > 0 && (
                <div className="mb-4 p-3 rounded-sm flex items-start gap-2.5" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44` }}>
                    <AlertTriangle size={16} style={{ color: RUST }} className="mt-0.5 shrink-0" />
                    <div className="text-sm" style={{ color: INK }}>
                        <span className="font-medium" style={{ color: RUST }}>Công nợ NCC sắp/đã đến hạn: </span>
                        {dueSoonPOs.map(({ po, due }) => `${po.code} (${po.supplier || "NCC"}) — ${due.overdue ? `quá hạn ${-due.daysLeft} ngày` : due.daysLeft === 0 ? "đến hạn hôm nay" : `còn ${due.daysLeft} ngày`}`).join(", ")}
                    </div>
                </div>
            )}
            <div className="flex items-center justify-between mb-5">
                <p className="text-sm opacity-60">{purchaseOrders.length} đơn nhập hàng</p>
                <div className="flex gap-2">
                    <button onClick={openQuickAddSupplier} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}>
                        <Truck size={15} /> Thêm nhà cung cấp
                    </button>
                    <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
                        <Plus size={15} /> Tạo đơn nhập hàng
                    </button>
                </div>
            </div>

            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 720 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        {["", "Mã đơn nhập", "Ngày nhập", "Trạng thái nhập", "Chi nhánh nhập", "Nhà cung cấp", "Số hoá đơn", "Nhân viên tạo", "Thanh toán", "Giá trị đơn", ""].map((h, hi) => (
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
                                    <td className="px-3 py-3 opacity-70 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{po.invoiceNo || "—"}</td>
                                    <td className="px-3 py-3 opacity-70 whitespace-nowrap">{po.createdBy || "—"}</td>
                                    <td className="px-3 py-3">
                      <span className="inline-block rounded-full px-2.5 py-0.5 text-[11px] whitespace-nowrap"
                            style={{ background: po.paid ? `${FOREST}1A` : `${BRASS}1A`, color: po.paid ? FOREST : BRASS }}>
                        {po.paid ? "Hoàn thành" : "Đang giao dịch"}
                      </span>
                                    </td>
                                    <td className="px-3 py-3 text-right font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>
                                        {vnd(total)}
                                        {(po.returns || []).length > 0 && <span className="block text-[10px] font-normal" style={{ color: RUST }}>đã trả {vnd(poReturnedValue(po))}</span>}
                                    </td>
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
                                        <td colSpan={11} className="px-6 py-3" style={{ background: PAPER, borderBottom: `1px dashed ${LINE}` }}>
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
                    {purchaseOrders.length === 0 && <tr><td colSpan={11} className="text-center py-8 opacity-50">Chưa có đơn nhập hàng nào.</td></tr>}
                    </tbody>
                </table>
            </div>

            {/* Modal tạo đơn mới */}
            {creating && (
                <Modal title="Tạo đơn nhập hàng" onClose={() => setCreating(false)} size="2xl">
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Nhà cung cấp" hint={suppliers.length === 0 ? "Chưa có nhà cung cấp nào — vào mục Nhà cung cấp để thêm trước" : "Chỉ chọn được NCC đã có trong danh sách"}>
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.supplierId || ""}
                                    onChange={(e) => { const sup = suppliers.find((s) => s.id === e.target.value); setForm({ ...form, supplierId: e.target.value, supplier: sup ? `${sup.code} - ${sup.name}` : "" }); }}>
                                <option value="">— Chọn nhà cung cấp —</option>
                                {suppliers.map((s) => <option key={s.id} value={s.id}>{s.code} - {s.name}</option>)}
                            </select>
                        </Field>
                        <Field label="Chi nhánh nhập">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.branch} onChange={(e) => setForm({ ...form, branch: e.target.value })}>
                                {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                            </select>
                        </Field>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Nhân viên tạo">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.createdBy} onChange={(e) => setForm({ ...form, createdBy: e.target.value })}>
                                {(employeeNames.length ? employeeNames : EMPLOYEES).map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                            </select>
                        </Field>
                        <Field label="Số hóa đơn" hint="Có thể điền sau">
                            <input className={inputCls} style={{ borderColor: LINE }} value={form.invoiceNo} onChange={(e) => setForm({ ...form, invoiceNo: e.target.value })} placeholder="VD: HD0004521" />
                        </Field>
                    </div>

                    <Field label="Hình thức thanh toán">
                        <div className="flex gap-2 items-center flex-wrap">
                            {PAYMENT_METHODS.map((m) => (
                                <button key={m.id} type="button" onClick={() => setForm({ ...form, paymentMethod: m.id })}
                                        className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: form.paymentMethod === m.id ? INK : LINE, background: form.paymentMethod === m.id ? INK : "transparent", color: form.paymentMethod === m.id ? "#fff" : INK }}>
                                    {m.label}
                                </button>
                            ))}
                            {form.paymentMethod === "credit" && (
                                <label className="flex items-center gap-1.5 text-xs ml-1" style={{ color: INK }}>
                                    Số ngày công nợ:
                                    <input type="number" min={1} value={form.creditDays} onChange={(e) => setForm({ ...form, creditDays: e.target.value })} className="w-16 border rounded-sm py-1 px-2 text-sm" style={{ borderColor: LINE }} />
                                </label>
                            )}
                        </div>
                    </Field>

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
                        <button onClick={() => submit("pending")} disabled={!form.supplierId || itemsInvalid(form.items)} className="flex-1 py-2.5 rounded-sm text-sm border disabled:opacity-40" style={{ borderColor: BRASS, color: BRASS }}>Chờ giao</button>
                        <button onClick={() => submit("received")} disabled={!form.supplierId || itemsInvalid(form.items)} className="flex-1 py-2.5 rounded-sm text-white text-sm disabled:opacity-40" style={{ background: FOREST }}>Nhập hàng</button>
                    </div>
                </Modal>
            )}

            {/* Modal xem / sửa đơn đã tạo */}
            {viewingPO && editForm && (
                <Modal title={`Đơn nhập hàng ${viewingPO.code}`} onClose={closeView} size="2xl">
                    <POProgressStepper po={viewingPO} />

                    {(() => {
                        const due = poDueInfo(viewingPO);
                        return due ? (
                            <div className="mb-3 p-2.5 rounded-sm flex items-center gap-2 text-xs" style={{ background: due.nearDue ? `${RUST}10` : PAPER, border: `1px solid ${due.nearDue ? RUST + "44" : LINE}` }}>
                                <AlertTriangle size={13} style={{ color: due.nearDue ? RUST : BRASS }} />
                                <span style={{ color: due.nearDue ? RUST : INK }}>
                  Hạn công nợ: {due.dueDate.toLocaleDateString("vi-VN")} — {due.overdue ? `đã quá hạn ${-due.daysLeft} ngày` : due.daysLeft === 0 ? "đến hạn hôm nay" : `còn ${due.daysLeft} ngày`}
                </span>
                            </div>
                        ) : null;
                    })()}

                    <div className="flex gap-2 flex-wrap mb-5">
                        {!viewingPO.paid && (
                            <button onClick={() => markPaid(viewingPO)} className="px-3.5 py-1.5 rounded-sm text-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                                Đánh dấu đã thanh toán cho nhà cung cấp
                            </button>
                        )}
                        {viewingPO.status === "received" && (
                            <button onClick={() => openReturn(viewingPO)} className="px-3.5 py-1.5 rounded-sm text-sm border" style={{ borderColor: RUST, color: RUST }}>
                                Trả hàng NCC
                            </button>
                        )}
                    </div>

                    {(viewingPO.returns || []).length > 0 && (
                        <div className="mb-5 p-3 rounded-sm" style={{ background: `${RUST}08`, border: `1px solid ${RUST}33` }}>
                            <p className="text-xs uppercase tracking-wider mb-2" style={{ color: RUST }}>Đã trả hàng NCC ({viewingPO.returns.length} phiếu · tổng {vnd(poReturnedValue(viewingPO))})</p>
                            <div className="space-y-2">
                                {viewingPO.returns.map((r) => (
                                    <div key={r.id} className="text-xs">
                                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{r.code}</span>
                                        <span className="opacity-50"> · {formatDateTime(r.createdAt)} · </span>
                                        {r.items.map((it) => {
                                            const p = products.find((x) => x.id === it.productId);
                                            return `${p?.name || "?"} x${it.qty}`;
                                        }).join(", ")}
                                        {r.note && <span className="opacity-50"> · Lý do: {r.note}</span>}
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}

                    <div className="my-4" style={{ borderTop: `1px dashed ${LINE}` }} />

                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Nhà cung cấp">
                            <select className={inputCls} style={{ borderColor: LINE }} value={editForm.supplierId || ""}
                                    onChange={(e) => { const sup = suppliers.find((s) => s.id === e.target.value); setEditForm({ ...editForm, supplierId: e.target.value, supplier: sup ? `${sup.code} - ${sup.name}` : "" }); }}>
                                <option value="">— Chọn nhà cung cấp —</option>
                                {suppliers.map((s) => <option key={s.id} value={s.id}>{s.code} - {s.name}</option>)}
                            </select>
                        </Field>
                        <Field label="Chi nhánh nhập">
                            <select className={inputCls} style={{ borderColor: LINE }} value={editForm.branch} onChange={(e) => setEditForm({ ...editForm, branch: e.target.value })}>
                                {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                            </select>
                        </Field>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Nhân viên tạo">
                            <select className={inputCls} style={{ borderColor: LINE }} value={editForm.createdBy} onChange={(e) => setEditForm({ ...editForm, createdBy: e.target.value })}>
                                {(employeeNames.length ? employeeNames : EMPLOYEES).map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                            </select>
                        </Field>
                        <Field label="Số hóa đơn">
                            <input className={inputCls} style={{ borderColor: LINE }} value={editForm.invoiceNo} onChange={(e) => setEditForm({ ...editForm, invoiceNo: e.target.value })} placeholder="VD: HD0004521" />
                        </Field>
                    </div>

                    <Field label="Hình thức thanh toán">
                        <div className="flex gap-2 items-center flex-wrap">
                            {PAYMENT_METHODS.map((m) => (
                                <button key={m.id} type="button" onClick={() => setEditForm({ ...editForm, paymentMethod: m.id })}
                                        className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: editForm.paymentMethod === m.id ? INK : LINE, background: editForm.paymentMethod === m.id ? INK : "transparent", color: editForm.paymentMethod === m.id ? "#fff" : INK }}>
                                    {m.label}
                                </button>
                            ))}
                            {editForm.paymentMethod === "credit" && (
                                <label className="flex items-center gap-1.5 text-xs ml-1" style={{ color: INK }}>
                                    Số ngày công nợ:
                                    <input type="number" min={1} value={editForm.creditDays} onChange={(e) => setEditForm({ ...editForm, creditDays: e.target.value })} className="w-16 border rounded-sm py-1 px-2 text-sm" style={{ borderColor: LINE }} />
                                </label>
                            )}
                        </div>
                    </Field>

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

            {/* Modal lập phiếu trả hàng NCC (nhập sai/hàng lỗi...) */}
            {returningPO && returnForm && (
                <Modal title={`Trả hàng NCC — ${returningPO.code}`} onClose={() => { setReturningPO(null); setReturnForm(null); }} size="xl">
                    <p className="text-xs opacity-60 mb-4">Chọn sản phẩm và số lượng cần trả lại cho <b>{returningPO.supplier}</b>. Hệ thống sẽ tự trừ kho và trừ vào công nợ phải trả NCC.</p>
                    <div className="space-y-3 mb-4">
                        {returnForm.items.map((it) => {
                            const po_it = returningPO.items.find((x) => x.productId === it.productId);
                            const p = products.find((x) => x.id === it.productId);
                            const already = alreadyReturnedQty(returningPO, it.productId);
                            const maxByOrder = (po_it?.qty || 0) - already;
                            const stats = productStats(p);
                            const availSeries = p?.hasSeries ? availableSeriesFromPO(returningPO, it.productId) : [];
                            const maxQty = p?.hasSeries ? Math.min(maxByOrder, availSeries.length) : Math.min(maxByOrder, stats.closingQty);
                            if (maxByOrder <= 0) return null;
                            return (
                                <div key={it.productId} className="p-3 rounded-sm" style={{ border: `1px solid ${LINE}` }}>
                                    <div className="flex items-center justify-between gap-3 mb-2">
                                        <div>
                                            <p className="text-sm font-medium" style={{ color: INK }}>{p?.name}</p>
                                            <p className="text-[11px] opacity-50">Đã nhập {po_it?.qty} · {already > 0 ? `Đã trả ${already} · ` : ""}Tối đa trả được {maxQty}</p>
                                        </div>
                                        <input type="number" min={0} max={maxQty} value={it.qty || ""} onChange={(e) => {
                                            const v = Math.max(0, Math.min(maxQty, Number(e.target.value) || 0));
                                            updateReturnItem(it.productId, { qty: v, series: p?.hasSeries ? [] : it.series });
                                        }} className="w-20 border rounded-sm py-1.5 px-2 text-center text-sm" style={{ borderColor: LINE }} />
                                    </div>
                                    {p?.hasSeries && it.qty > 0 && (
                                        <SeriesPicker available={availSeries} selected={it.series} setSelected={(arr) => updateReturnItem(it.productId, { series: arr })} need={it.qty} />
                                    )}
                                </div>
                            );
                        })}
                    </div>
                    <Field label="Lý do trả hàng">
                        <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={returnForm.note} onChange={(e) => setReturnForm({ ...returnForm, note: e.target.value })} placeholder="VD: nhập sai model, hàng lỗi, NCC giao thừa…" />
                    </Field>
                    <button onClick={submitReturn} disabled={returnInvalid()} className="w-full py-2.5 rounded-sm text-white text-sm mt-2 disabled:opacity-40" style={{ background: RUST }}>
                        Xác nhận trả hàng NCC
                    </button>
                </Modal>
            )}

            {/* Modal thêm nhanh nhà cung cấp — thao tác tắt, không cần rời khỏi mục Nhập hàng */}
            {quickAddingSupplier && (
                <Modal title="Thêm nhà cung cấp" onClose={() => setQuickAddingSupplier(false)}>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Mã NCC" hint="Tự động sinh — có thể sửa lại">
                            <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={quickSupplierForm.code} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, code: e.target.value })} />
                        </Field>
                        <Field label="Mã số thuế"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={quickSupplierForm.taxCode} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, taxCode: e.target.value.replace(/\D/g, "") })} /></Field>
                    </div>
                    <Field label="Tên nhà cung cấp"><input className={inputCls} style={{ borderColor: LINE }} value={quickSupplierForm.name} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, name: e.target.value })} /></Field>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Người liên hệ"><input className={inputCls} style={{ borderColor: LINE }} value={quickSupplierForm.contactPerson} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, contactPerson: e.target.value })} /></Field>
                        <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={quickSupplierForm.phone} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, phone: e.target.value.replace(/\D/g, "") })} /></Field>
                    </div>
                    <Field label="Địa chỉ"><input className={inputCls} style={{ borderColor: LINE }} value={quickSupplierForm.address} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, address: e.target.value })} /></Field>
                    <Field label="Hình thức thanh toán">
                        <div className="flex gap-2 items-center flex-wrap">
                            {SUPPLIER_PAYMENT_TERMS.map((m) => (
                                <button key={m.id} type="button" onClick={() => setQuickSupplierForm({ ...quickSupplierForm, paymentTerm: m.id })}
                                        className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: quickSupplierForm.paymentTerm === m.id ? INK : LINE, background: quickSupplierForm.paymentTerm === m.id ? INK : "transparent", color: quickSupplierForm.paymentTerm === m.id ? "#fff" : INK }}>
                                    {m.label}
                                </button>
                            ))}
                            {quickSupplierForm.paymentTerm === "credit" && (
                                <label className="flex items-center gap-1.5 text-xs ml-1" style={{ color: INK }}>
                                    Số ngày công nợ:
                                    <input type="number" min={1} value={quickSupplierForm.creditDays} onChange={(e) => setQuickSupplierForm({ ...quickSupplierForm, creditDays: e.target.value })} className="w-16 border rounded-sm py-1 px-2 text-sm" style={{ borderColor: LINE }} />
                                </label>
                            )}
                        </div>
                    </Field>
                    <button onClick={submitQuickAddSupplier} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>Thêm nhà cung cấp</button>
                </Modal>
            )}
        </div>
    );
}



/* ---------------- Sản phẩm & Tồn kho — bao gồm 2 menu con ---------------- */

function Suppliers({ suppliers, setSuppliers, purchaseOrders, addLog, goToDoc, navTarget, onFocusHandled }) {
    const [view, setView] = useState("list"); // list | debt
    const [query, setQuery] = useState("");
    const [debtSort, setDebtSort] = useState("amount"); // amount | due
    const [editing, setEditing] = useState(null);
    const [form, setForm] = useState({});
    const [debtDetail, setDebtDetail] = useState(null); // { supplier, unpaidPOs } — xem chi tiết từng đơn nợ của 1 NCC

    useEffect(() => {
        if (navTarget?.type === "supplier") {
            const s = suppliers.find((x) => x.id === navTarget.id);
            if (s) { setView("list"); setForm({ ...s }); setEditing(s); }
            onFocusHandled && onFocusHandled();
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [navTarget]);

    const filtered = suppliers.filter((s) =>
        s.name.toLowerCase().includes(query.toLowerCase()) || s.code.toLowerCase().includes(query.toLowerCase()) || (s.contactPerson || "").toLowerCase().includes(query.toLowerCase())
    );

    const openNew = () => { setForm({ code: nextSupplierCode(suppliers), name: "", taxCode: "", address: "", contactPerson: "", phone: "", email: "", paymentTerm: "cash", creditDays: 30 }); setEditing({}); };
    const openEdit = (s) => { setForm({ ...s }); setEditing(s); };
    const submit = () => {
        if (!form.name) return;
        if (editing.id) {
            setSuppliers((prev) => prev.map((s) => (s.id === editing.id ? { ...s, ...form, creditDays: Number(form.creditDays) || 0 } : s)));
            addLog("Sửa nhà cung cấp", form.name);
        } else {
            setSuppliers((prev) => [...prev, { ...form, id: uid(), code: form.code || nextSupplierCode(suppliers), creditDays: Number(form.creditDays) || 0 }]);
            addLog("Thêm nhà cung cấp", form.name);
        }
        setEditing(null);
    };
    const remove = (id) => setSuppliers((prev) => prev.filter((s) => s.id !== id));

    // Công nợ phải trả từng NCC: cộng dồn các đơn nhập chưa thanh toán khớp với NCC đó.
    const debtList = suppliers.map((s) => {
        const unpaidPOs = purchaseOrders.filter((po) => po.status === "received" && !po.paid && poMatchesSupplier(po, s));
        const total = unpaidPOs.reduce((sum, po) => sum + poNetTotal(po), 0);
        const dues = unpaidPOs.map((po) => poDueInfo(po)).filter(Boolean);
        const nearestDue = dues.length ? dues.reduce((a, b) => (a.daysLeft < b.daysLeft ? a : b)) : null;
        return { supplier: s, total, poCount: unpaidPOs.length, nearestDue, unpaidPOs };
    }).filter((d) => d.total > 0);
    const grandTotal = debtList.reduce((s, d) => s + d.total, 0);
    const sortedDebt = [...debtList].sort((a, b) => {
        if (debtSort === "due") {
            if (!a.nearestDue) return 1;
            if (!b.nearestDue) return -1;
            return a.nearestDue.daysLeft - b.nearestDue.daysLeft;
        }
        return b.total - a.total;
    });

    // Xuất Excel toàn bộ công nợ NCC: 1 sheet tổng hợp theo từng NCC + 1 sheet chi tiết từng đơn nhập còn nợ (kèm số hoá đơn).
    const exportAllDebt = () => {
        if (sortedDebt.length === 0) { alert("Không có công nợ nào để xuất."); return; }
        const summaryRows = sortedDebt.map((d) => ({
            "Mã NCC": d.supplier.code, "Tên NCC": d.supplier.name, "Số đơn nợ": d.poCount, "Tổng nợ": d.total,
            "Hạn gần nhất": d.nearestDue ? (d.nearestDue.overdue ? `Quá hạn ${-d.nearestDue.daysLeft} ngày` : `Còn ${d.nearestDue.daysLeft} ngày`) : "",
        }));
        const detailRows = [];
        sortedDebt.forEach((d) => {
            d.unpaidPOs.forEach((po) => {
                const due = poDueInfo(po);
                detailRows.push({
                    "Mã NCC": d.supplier.code, "Tên NCC": d.supplier.name, "Mã đơn nhập": po.code, "Số hoá đơn": po.invoiceNo || "",
                    "Ngày nhập": formatDateTime(po.createdAt), "Giá trị đơn gốc": poTotal(po), "Đã trả NCC": poReturnedValue(po), "Còn nợ": poNetTotal(po),
                    "Hạn công nợ": due ? (due.overdue ? `Quá hạn ${-due.daysLeft} ngày` : `Còn ${due.daysLeft} ngày`) : "",
                });
            });
        });
        exportExcel(`CongNoNCC_${todayISO()}`, [{ name: "Tổng hợp theo NCC", rows: summaryRows }, { name: "Chi tiết đơn nợ", rows: detailRows }]);
        addLog("Xuất Excel công nợ NCC", `${sortedDebt.length} nhà cung cấp`);
    };
    const exportSupplierDebt = (d) => {
        const rows = d.unpaidPOs.map((po) => {
            const due = poDueInfo(po);
            return {
                "Mã đơn nhập": po.code, "Số hoá đơn": po.invoiceNo || "", "Ngày nhập": formatDateTime(po.createdAt), "Chi nhánh nhập": po.branch,
                "Nhân viên tạo": po.createdBy || "", "Giá trị đơn gốc": poTotal(po), "Đã trả NCC": poReturnedValue(po), "Còn nợ": poNetTotal(po),
                "Hạn công nợ": due ? (due.overdue ? `Quá hạn ${-due.daysLeft} ngày` : `Còn ${due.daysLeft} ngày`) : "",
            };
        });
        exportExcel(`CongNo_${d.supplier.code}_${todayISO()}`, [{ name: d.supplier.name.slice(0, 28), rows }]);
        addLog("Xuất Excel công nợ NCC", d.supplier.name);
    };

    return (
        <div>
            <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
                <div className="grid grid-cols-2 max-w-xs gap-2">
                    <button onClick={() => setView("list")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: view === "list" ? INK : LINE, background: view === "list" ? INK : "transparent", color: view === "list" ? "#fff" : INK }}>
                        Danh sách NCC
                    </button>
                    <button onClick={() => setView("debt")} className="px-3.5 py-2.5 rounded-full text-sm border text-center relative"
                            style={{ borderColor: view === "debt" ? INK : LINE, background: view === "debt" ? INK : "transparent", color: view === "debt" ? "#fff" : INK }}>
                        Công nợ NCC
                        {debtList.length > 0 && <span className="ml-1.5 text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: view === "debt" ? "rgba(255,255,255,0.25)" : `${RUST}1A`, color: view === "debt" ? "#fff" : RUST }}>{debtList.length}</span>}
                    </button>
                </div>
                {view === "list" && (
                    <div className="flex items-center gap-3 flex-wrap">
                        <div className="relative flex-1 max-w-xs">
                            <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
                            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã, tên hoặc người liên hệ…"
                                   className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
                        </div>
                        <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
                            <Plus size={15} /> Thêm nhà cung cấp
                        </button>
                    </div>
                )}
            </div>

            {view === "debt" ? (
                <div>
                    <div className="p-5 rounded-sm mb-5 flex items-center justify-between flex-wrap gap-3" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                        <div>
                            <p className="text-xs uppercase tracking-wider opacity-55 mb-1">Tổng phải trả nhà cung cấp</p>
                            <p className="text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(grandTotal)}</p>
                        </div>
                        <div className="flex items-center gap-2 flex-wrap">
                            <div className="flex gap-1.5">
                                <button onClick={() => setDebtSort("amount")} className="text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: debtSort === "amount" ? INK : LINE, background: debtSort === "amount" ? INK : "transparent", color: debtSort === "amount" ? "#fff" : INK }}>Số tiền</button>
                                <button onClick={() => setDebtSort("due")} className="text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: debtSort === "due" ? INK : LINE, background: debtSort === "due" ? INK : "transparent", color: debtSort === "due" ? "#fff" : INK }}>Hạn công nợ</button>
                            </div>
                            <button onClick={exportAllDebt} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                                <FileSpreadsheet size={13} /> Xuất Excel
                            </button>
                        </div>
                    </div>
                    {sortedDebt.length === 0 ? (
                        <p className="text-sm opacity-50 text-center py-16">Không có nhà cung cấp nào đang nợ.</p>
                    ) : (
                        <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                            <table className="w-full text-sm" style={{ minWidth: 700 }}>
                                <thead>
                                <tr style={{ borderBottom: `2px solid ${INK}` }}>
                                    {["Nhà cung cấp", "Số đơn nợ", "Tổng nợ", "Hạn gần nhất", ""].map((h, i) => (
                                        <th key={i} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
                                    ))}
                                </tr>
                                </thead>
                                <tbody>
                                {sortedDebt.map((d) => (
                                    <tr key={d.supplier.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                                        <td className="px-3 py-3">
                                            <p style={{ color: INK }}>{d.supplier.name}</p>
                                            <p className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.supplier.code}</p>
                                        </td>
                                        <td className="px-3 py-3" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.poCount}</td>
                                        <td className="px-3 py-3 font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(d.total)}</td>
                                        <td className="px-3 py-3 whitespace-nowrap">
                                            {d.nearestDue ? (
                                                <span className="text-xs px-2 py-0.5 rounded-full" style={{ background: d.nearestDue.nearDue ? `${RUST}1A` : PAPER, color: d.nearestDue.nearDue ? RUST : INK }}>
                            {d.nearestDue.overdue ? `Quá hạn ${-d.nearestDue.daysLeft} ngày` : d.nearestDue.daysLeft === 0 ? "Đến hạn hôm nay" : `Còn ${d.nearestDue.daysLeft} ngày`}
                          </span>
                                            ) : <span className="opacity-40">—</span>}
                                        </td>
                                        <td className="px-3 py-3 text-right">
                                            <button onClick={() => setDebtDetail(d)} className="text-xs px-2.5 py-1 rounded-sm border whitespace-nowrap" style={{ borderColor: LINE, color: INK }}>Chi tiết</button>
                                        </td>
                                    </tr>
                                ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            ) : (
                <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
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
            )}

            {editing !== null && (
                <Modal title={editing.id ? "Sửa nhà cung cấp" : "Thêm nhà cung cấp"} onClose={() => setEditing(null)}>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Field label="Mã NCC" hint="Tự động sinh — có thể sửa lại">
                            <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} />
                        </Field>
                        <Field label="Mã số thuế">
                            <input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={form.taxCode} onChange={(e) => setForm({ ...form, taxCode: e.target.value.replace(/\D/g, "") })} />
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
                            <input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value.replace(/\D/g, "") })} />
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

            {debtDetail && (
                <Modal title={`Công nợ — ${debtDetail.supplier.name}`} onClose={() => setDebtDetail(null)} size="2xl">
                    <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                        <p className="text-sm opacity-60">{debtDetail.poCount} đơn chưa thanh toán · Tổng nợ <b style={{ color: RUST }}>{vnd(debtDetail.total)}</b></p>
                        <button onClick={() => exportSupplierDebt(debtDetail)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                            <FileSpreadsheet size={12} /> Xuất Excel
                        </button>
                    </div>
                    <div className="rounded-sm overflow-auto" style={{ border: `1px solid ${LINE}`, maxHeight: "55vh" }}>
                        <table className="w-full text-sm">
                            <thead style={{ background: PAPER, position: "sticky", top: 0 }}><tr>
                                <th className="text-left px-3 py-2 text-xs uppercase tracking-wider opacity-60">Mã đơn nhập</th>
                                <th className="text-left px-3 py-2 text-xs uppercase tracking-wider opacity-60">Số hoá đơn</th>
                                <th className="text-left px-3 py-2 text-xs uppercase tracking-wider opacity-60">Ngày nhập</th>
                                <th className="text-left px-3 py-2 text-xs uppercase tracking-wider opacity-60">Hạn công nợ</th>
                                <th className="text-right px-3 py-2 text-xs uppercase tracking-wider opacity-60">Giá trị</th>
                            </tr></thead>
                            <tbody>
                            {debtDetail.unpaidPOs.map((po) => {
                                const due = poDueInfo(po);
                                return (
                                    <tr key={po.id} style={{ borderTop: `1px dashed ${LINE}` }}>
                                        <td className="px-3 py-2">
                                            {goToDoc ? (
                                                <button onClick={() => goToDoc(po.code)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{po.code}</button>
                                            ) : <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{po.code}</span>}
                                        </td>
                                        <td className="px-3 py-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{po.invoiceNo || "—"}</td>
                                        <td className="px-3 py-2 whitespace-nowrap opacity-80" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(po.createdAt)}</td>
                                        <td className="px-3 py-2 whitespace-nowrap">
                                            {due ? (
                                                <span className="text-xs px-2 py-0.5 rounded-full" style={{ background: due.nearDue ? `${RUST}1A` : PAPER, color: due.nearDue ? RUST : INK }}>
                            {due.overdue ? `Quá hạn ${-due.daysLeft} ngày` : due.daysLeft === 0 ? "Đến hạn hôm nay" : `Còn ${due.daysLeft} ngày`}
                          </span>
                                            ) : <span className="opacity-40">—</span>}
                                        </td>
                                        <td className="px-3 py-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>
                                            {vnd(poNetTotal(po))}
                                            {poReturnedValue(po) > 0 && <span className="block text-[10px] font-normal opacity-50">(đã trừ {vnd(poReturnedValue(po))} trả NCC)</span>}
                                        </td>
                                    </tr>
                                );
                            })}
                            {debtDetail.unpaidPOs.length === 0 && <tr><td colSpan={5} className="text-center py-8 opacity-40">Không có đơn nào.</td></tr>}
                            </tbody>
                        </table>
                    </div>
                </Modal>
            )}
        </div>
    );
}

// Kiểm kho — đối chiếu tồn kho thực tế (đếm/quét) với sổ sách, tự tính chênh lệch và tạo bút toán điều chỉnh.
// 3 chế độ đếm: quét Series (sản phẩm quản lý theo series), quét Barcode chung (mỗi lần quét +1), hoặc nhập tay số lượng.
// Phiếu bảo hành — tạo phiếu tiếp nhận bảo hành cho khách, có thể liên kết tới đúng đơn hàng/series đã bán, hoặc nhập tay.
function WarrantyTickets({ products, setProducts, orders, customers, warrantyTickets, setWarrantyTickets, currentUser, addLog, goToDoc }) {
    const [view, setView] = useState("history"); // history | new
    const [statusFilter, setStatusFilter] = useState("all");
    const [viewingTicket, setViewingTicket] = useState(null);
    const [printPaperSize, setPrintPaperSize] = useState("A5");
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);
    const [showWarrantyStock, setShowWarrantyStock] = useState(false);
    const [statusDraft, setStatusDraft] = useState(null); // "confirmed" | "rejected" — đang mở form soạn để xác nhận/từ chối, chưa lưu
    const [resolutionType, setResolutionType] = useState("exchange");
    const [exchangeSource, setExchangeSource] = useState("main"); // main = kho chính, warranty = hàng NCC trả lại (kho bảo hành)
    const [exchangeProductId, setExchangeProductId] = useState("");
    const [exchangeSeriesText, setExchangeSeriesText] = useState("");
    const [exchangeQty, setExchangeQty] = useState(1);
    const [selectedReplacementId, setSelectedReplacementId] = useState("");
    const [refundAmount, setRefundAmount] = useState(0);
    const [rejectReason, setRejectReason] = useState("");
    const [addingReplacement, setAddingReplacement] = useState(false);
    const [replProductId, setReplProductId] = useState("");
    const [replSeriesText, setReplSeriesText] = useState("");
    const [replQty, setReplQty] = useState(1);

    const openViewingTicket = (t) => {
        setViewingTicket(t);
        setPrintBlockedUrl(null);
        setStatusDraft(null);
        setAddingReplacement(false);
        setResolutionType(t.resolutionType || "exchange");
        setExchangeSource(t.exchangeSource || "main");
        setExchangeProductId(t.exchangeProductId || "");
        setExchangeSeriesText((t.exchangeSeries || []).join(", "));
        setExchangeQty(t.exchangeQty || 1);
        setSelectedReplacementId((t.replacementReceived || [])[0]?.id || "");
        setRefundAmount(t.refundAmount || 0);
        setRejectReason(t.rejectReason || "");
    };

    // Ghi nhận hàng NCC đã trả lại (thay thế hoặc đã sửa xong) — nhập thẳng vào Kho bảo hành của phiếu này, không đụng tới kho chính.
    const addReplacement = (ticket) => {
        if (!replProductId) { alert("Vui lòng chọn sản phẩm NCC đã trả lại."); return; }
        const p = products.find((x) => x.id === replProductId);
        const entry = {
            id: uid(), productId: replProductId, productName: p?.name || "", productCode: p?.code || "",
            series: replSeriesText.split(",").map((s) => s.trim()).filter(Boolean), qty: Number(replQty) || 1, receivedDate: todayISO(),
        };
        const newList = [...(ticket.replacementReceived || []), entry];
        setWarrantyTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, replacementReceived: newList } : t)));
        addLog("Nhập hàng thay thế NCC (Kho bảo hành)", `${ticket.code} · ${entry.productName} x${entry.qty}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, replacementReceived: newList } : v));
        setAddingReplacement(false); setReplProductId(""); setReplSeriesText(""); setReplQty(1);
    };

    // ----- Form tạo phiếu mới -----
    const emptyForm = () => ({ customerName: "", customerPhone: "", customerAddress: "", receivedDate: todayISO(), returnDate: "", note: "" });
    const [form, setForm] = useState(emptyForm());
    const [lines, setLines] = useState([]); // { key, orderId, orderCode, productId, productName, productCode, seriesText, qty, warrantyMonths, condition }
    const [orderQuery, setOrderQuery] = useState("");
    const [selectedOrder, setSelectedOrder] = useState(null);

    const orderMatches = orderQuery.trim()
        ? orders.filter((o) => o.status !== "cancelled" && (o.code.toLowerCase().includes(orderQuery.trim().toLowerCase()) || (customers.find((c) => c.id === o.customerId)?.name || "").toLowerCase().includes(orderQuery.trim().toLowerCase()))).slice(0, 20)
        : [];

    const pickOrder = (order) => {
        setSelectedOrder(order);
        const cust = customers.find((c) => c.id === order.customerId);
        const addr = order.shippingAddress;
        setForm((f) => ({
            ...f,
            customerName: addr?.recipientName || cust?.name || "",
            customerPhone: addr?.recipientPhone || cust?.phone || "",
            customerAddress: [addr?.addressDetail, addr?.ward, addr?.province].filter(Boolean).join(", ") || [cust?.addressDetail, cust?.ward, cust?.province].filter(Boolean).join(", "),
        }));
        setOrderQuery("");
    };

    const addLineFromOrder = (order, item) => {
        const p = products.find((x) => x.id === item.productId);
        setLines((prev) => [...prev, {
            key: uid(), orderId: order.id, orderCode: order.code, productId: item.productId, productName: p?.name || "?", productCode: p?.code || "",
            seriesText: "", qty: 1, warrantyMonths: p?.warrantyMonths || 0, condition: "",
        }]);
    };
    const addManualLine = (productId) => {
        const p = products.find((x) => x.id === productId);
        if (!p) return;
        setLines((prev) => [...prev, {
            key: uid(), orderId: "", orderCode: "", productId: p.id, productName: p.name, productCode: p.code,
            seriesText: "", qty: 1, warrantyMonths: p.warrantyMonths || 0, condition: "",
        }]);
    };
    const updateLine = (key, patch) => setLines((prev) => prev.map((l) => (l.key === key ? { ...l, ...patch } : l)));
    const removeLine = (key) => setLines((prev) => prev.filter((l) => l.key !== key));

    const resetForm = () => { setForm(emptyForm()); setLines([]); setOrderQuery(""); setSelectedOrder(null); };

    const submitTicket = () => {
        if (!form.customerName.trim()) { alert("Vui lòng nhập tên khách hàng."); return; }
        if (lines.length === 0) { alert("Vui lòng thêm ít nhất 1 sản phẩm vào phiếu bảo hành."); return; }
        const code = nextWarrantyCode(warrantyTickets);
        const ticket = normalizeWarrantyTicket({
            id: uid(), code, createdAt: new Date().toISOString(), createdBy: currentUser.fullName,
            customerName: form.customerName, customerPhone: form.customerPhone, customerAddress: form.customerAddress,
            receivedDate: form.receivedDate, returnDate: form.returnDate, status: "pending", note: form.note,
            items: lines.map((l) => ({
                orderId: l.orderId, orderCode: l.orderCode, productId: l.productId, productName: l.productName, productCode: l.productCode,
                series: l.seriesText.split(",").map((s) => s.trim()).filter(Boolean), qty: Number(l.qty) || 1, warrantyMonths: Number(l.warrantyMonths) || 0, condition: l.condition,
            })),
        });
        setWarrantyTickets((prev) => [ticket, ...prev]);
        addLog("Tạo phiếu bảo hành", `${code} · ${form.customerName} · ${lines.length} sản phẩm`);
        resetForm();
        setView("history");
    };

    // Đang xử lý / Đã trả khách: đổi trạng thái trực tiếp, không cần thêm dữ liệu.
    const setTicketStatusDirect = (ticket, status) => {
        const xtbhCode = (status === "completed" && !ticket.xtbhCode) ? `XTBH-${ticket.code}` : ticket.xtbhCode;
        // Nếu đổi SP lấy từ kho chính và chưa từng trừ kho cho phiếu này, trừ kho ngay lúc xác nhận đã trả cho khách.
        const shouldDeduct = status === "completed" && ticket.resolutionType === "exchange" && ticket.exchangeSource === "main" && ticket.exchangeProductId && !ticket.stockDeducted;
        if (shouldDeduct) {
            setProducts((prev) => prev.map((p) => (p.id === ticket.exchangeProductId
                ? { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: xtbhCode, date: todayISO(), qty: ticket.exchangeQty, price: p.costPrice || 0, series: ticket.exchangeSeries }] }
                : p)));
        }
        setWarrantyTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, status, xtbhCode, stockDeducted: t.stockDeducted || shouldDeduct } : t)));
        addLog("Cập nhật phiếu bảo hành", `${ticket.code} → ${WARRANTY_TICKET_STATUSES.find((s) => s.id === status)?.label}${status === "completed" ? ` · Phiếu xuất trả BH: ${xtbhCode}` : ""}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, status, xtbhCode, stockDeducted: v.stockDeducted || shouldDeduct } : v));
        setStatusDraft(null);
    };
    // Xác nhận: phải chọn đổi sản phẩm (từ kho chính hoặc kho bảo hành) hoặc hoàn tiền kèm thông tin cụ thể trước khi lưu.
    const saveConfirm = (ticket) => {
        if (resolutionType === "exchange" && exchangeSource === "main" && !exchangeProductId) { alert("Vui lòng chọn sản phẩm từ kho chính dùng để đổi cho khách."); return; }
        if (resolutionType === "exchange" && exchangeSource === "warranty" && !selectedReplacementId) { alert("Chưa có hàng NCC trả lại trong Kho bảo hành cho phiếu này — vui lòng dùng nút \"Nhập hàng thay thế NCC\" trước, hoặc chọn nguồn Kho chính."); return; }
        if (resolutionType === "refund" && (!refundAmount || refundAmount <= 0)) { alert("Vui lòng nhập số tiền hoàn lại cho khách."); return; }
        let patch;
        if (resolutionType === "exchange" && exchangeSource === "warranty") {
            const repl = (ticket.replacementReceived || []).find((r) => r.id === selectedReplacementId);
            patch = {
                status: "confirmed", resolutionType, exchangeSource: "warranty",
                exchangeProductId: repl.productId, exchangeProductName: repl.productName, exchangeProductCode: repl.productCode,
                exchangeSeries: repl.series, exchangeQty: repl.qty, refundAmount: 0,
            };
        } else {
            const p = products.find((x) => x.id === exchangeProductId);
            patch = {
                status: "confirmed", resolutionType, exchangeSource: "main",
                exchangeProductId: resolutionType === "exchange" ? exchangeProductId : "",
                exchangeProductName: resolutionType === "exchange" ? (p?.name || "") : "",
                exchangeProductCode: resolutionType === "exchange" ? (p?.code || "") : "",
                exchangeSeries: resolutionType === "exchange" ? exchangeSeriesText.split(",").map((s) => s.trim()).filter(Boolean) : [],
                exchangeQty: resolutionType === "exchange" ? (Number(exchangeQty) || 1) : 0,
                refundAmount: resolutionType === "refund" ? (Number(refundAmount) || 0) : 0,
            };
        }
        setWarrantyTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, ...patch } : t)));
        addLog("Xác nhận xử lý bảo hành", `${ticket.code} · ${resolutionType === "exchange" ? `Đổi SP (${exchangeSource === "warranty" ? "Kho bảo hành" : "Kho chính"}): ${patch.exchangeProductName}` : `Hoàn tiền: ${vnd(refundAmount)}`}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, ...patch } : v));
        setStatusDraft(null);
    };
    // Từ chối bảo hành: bắt buộc ghi lý do.
    const saveReject = (ticket) => {
        if (!rejectReason.trim()) { alert("Vui lòng nhập lý do từ chối bảo hành."); return; }
        const patch = { status: "rejected", rejectReason: rejectReason.trim() };
        setWarrantyTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, ...patch } : t)));
        addLog("Từ chối bảo hành", `${ticket.code} · Lý do: ${rejectReason.trim()}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, ...patch } : v));
        setStatusDraft(null);
    };
    // Mở lại phiếu đã hoàn tất (Đã trả khách / Từ chối BH) để sửa lại thông tin — vd lỡ tay bấm nhầm.
    // Nếu trước đó đã trừ kho chính (đổi SP từ kho chính) thì tự hoàn lại kho trước khi mở lại, tránh lệch tồn kho.
    const reopenTicket = (ticket) => {
        const wasCompleted = ticket.status === "completed";
        const newStatus = wasCompleted ? "confirmed" : "pending";
        if (wasCompleted && ticket.stockDeducted && ticket.resolutionType === "exchange" && ticket.exchangeSource === "main" && ticket.exchangeProductId) {
            setProducts((prev) => prev.map((p) => (p.id === ticket.exchangeProductId
                ? { ...p, movements: [...p.movements, { id: uid(), type: "in", docNo: `${ticket.xtbhCode || ticket.code}-MOLAI`, date: todayISO(), qty: ticket.exchangeQty, price: p.costPrice || 0, series: ticket.exchangeSeries }] }
                : p)));
        }
        setWarrantyTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, status: newStatus, stockDeducted: false } : t)));
        addLog("Mở lại phiếu bảo hành", `${ticket.code} → ${WARRANTY_TICKET_STATUSES.find((s) => s.id === newStatus)?.label}${wasCompleted && ticket.stockDeducted ? " · Đã hoàn lại kho chính" : ""}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, status: newStatus, stockDeducted: false } : v));
        setStatusDraft(null);
    };

    const doPrint = (ticket) => {
        const html = buildWarrantyTicketHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };
    const doPrintXTBH = (ticket) => {
        const html = buildWarrantyReturnSlipHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };
    const doPrintReject = (ticket) => {
        const html = buildWarrantyRejectSlipHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };

    // Kho bảo hành — liệt kê chi tiết từng sản phẩm/serial đang được giữ, kèm phiếu bảo hành gốc để bấm vào xem.
    // Gồm 2 loại: (1) hàng lỗi khách mang tới, chờ xử lý; (2) hàng NCC đã trả lại (thay thế/sửa xong), sẵn sàng giao khách.
    const warrantyStockList = useMemo(() => {
        const list = [];
        warrantyTickets.forEach((t) => {
            if (t.status !== "pending" && t.status !== "confirmed") return;
            t.items.forEach((it) => {
                if (!it.productId) return;
                if (it.series && it.series.length > 0) {
                    it.series.forEach((sn) => {
                        list.push({ key: `${t.id}-c-${it.productId}-${sn}`, ticket: t, productName: it.productName, productCode: it.productCode, serial: sn, qty: 1, source: "customer" });
                    });
                } else {
                    list.push({ key: `${t.id}-c-${it.productId}`, ticket: t, productName: it.productName, productCode: it.productCode, serial: "", qty: it.qty, source: "customer" });
                }
            });
            (t.replacementReceived || []).forEach((r) => {
                if (!r.productId) return;
                if (r.series && r.series.length > 0) {
                    r.series.forEach((sn) => {
                        list.push({ key: `${t.id}-s-${r.id}-${sn}`, ticket: t, productName: r.productName, productCode: r.productCode, serial: sn, qty: 1, source: "supplier" });
                    });
                } else {
                    list.push({ key: `${t.id}-s-${r.id}`, ticket: t, productName: r.productName, productCode: r.productCode, serial: "", qty: r.qty, source: "supplier" });
                }
            });
        });
        return list.sort((a, b) => a.productName.localeCompare(b.productName));
    }, [warrantyTickets]);

    const filteredTickets = warrantyTickets.filter((t) => statusFilter === "all" || t.status === statusFilter);

    if (view === "new") {
        return (
            <div>
                <button onClick={() => { resetForm(); setView("history"); }} className="text-sm mb-4 opacity-60 hover:opacity-100">← Quay lại danh sách phiếu bảo hành</button>

                <div className="p-5 rounded-sm mb-5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Liên kết đơn hàng (không bắt buộc)</p>
                    <div className="relative mb-2">
                        <input value={orderQuery} onChange={(e) => setOrderQuery(e.target.value)} placeholder="Tìm theo mã đơn hoặc tên khách hàng…"
                               className="w-full border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }} />
                        {orderMatches.length > 0 && (
                            <div className="absolute z-20 mt-1 w-full max-h-52 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                                {orderMatches.map((o) => (
                                    <button key={o.id} onClick={() => pickOrder(o)} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex justify-between gap-3" style={{ borderBottom: `1px dashed ${LINE}` }}>
                                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{o.code}</span>
                                        <span className="opacity-60 truncate">{customers.find((c) => c.id === o.customerId)?.name || "Khách lẻ"}</span>
                                    </button>
                                ))}
                            </div>
                        )}
                    </div>
                    {selectedOrder && (
                        <div className="p-3 rounded-sm" style={{ background: PAPER }}>
                            <div className="flex items-center justify-between mb-2">
                                <span className="text-sm font-medium" style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>{selectedOrder.code}</span>
                                <button onClick={() => setSelectedOrder(null)} className="text-xs underline opacity-60">Bỏ chọn</button>
                            </div>
                            <div className="space-y-1.5">
                                {selectedOrder.items.map((it) => {
                                    const p = products.find((x) => x.id === it.productId);
                                    return (
                                        <div key={it.productId} className="flex items-center justify-between gap-3 text-sm p-2 rounded-sm" style={{ background: "#fff", border: `1px dashed ${LINE}` }}>
                                            <span className="truncate">{p?.name} <span className="opacity-50 text-xs">(đã mua {it.qty}{it.series?.length ? ` · series: ${it.series.join(", ")}` : ""})</span></span>
                                            <button onClick={() => addLineFromOrder(selectedOrder, it)} className="text-xs px-2.5 py-1 rounded-sm border shrink-0" style={{ borderColor: LINE, color: INK }}>+ Thêm vào phiếu</button>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                    )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-5">
                    <Field label="Tên khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} /></Field>
                    <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerPhone} onChange={(e) => setForm({ ...form, customerPhone: e.target.value })} /></Field>
                </div>
                <Field label="Địa chỉ khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerAddress} onChange={(e) => setForm({ ...form, customerAddress: e.target.value })} /></Field>

                <div className="p-5 rounded-sm my-5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <div className="flex items-center justify-between mb-3">
                        <p className="text-xs uppercase tracking-wider opacity-60">Sản phẩm bảo hành ({lines.length})</p>
                    </div>
                    {lines.length > 0 && (
                        <div className="rounded-sm overflow-x-auto mb-3" style={{ border: `1px solid ${LINE}` }}>
                            <table className="w-full text-sm" style={{ minWidth: 640 }}>
                                <thead><tr style={{ background: PAPER }}>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Serial</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">SL</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">BH (tháng)</th>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Tình trạng</th>
                                    <th className="py-2 px-2"></th>
                                </tr></thead>
                                <tbody>
                                {lines.map((l) => (
                                    <tr key={l.key} style={{ borderTop: `1px dashed ${LINE}` }}>
                                        <td className="py-2 px-2">{l.productName}{l.orderCode ? <span className="block text-[10px] opacity-40" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>Đơn {l.orderCode}</span> : null}</td>
                                        <td className="py-2 px-2"><input value={l.seriesText} onChange={(e) => updateLine(l.key, { seriesText: e.target.value })} placeholder="Cách nhau bởi dấu phẩy" className="w-32 border rounded-sm py-1 px-1.5 text-xs" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} /></td>
                                        <td className="py-2 px-2 text-right"><input type="number" min={1} value={l.qty} onChange={(e) => updateLine(l.key, { qty: Number(e.target.value) || 1 })} className="w-14 border rounded-sm py-1 px-1.5 text-right text-sm" style={{ borderColor: LINE }} /></td>
                                        <td className="py-2 px-2 text-right"><input type="number" min={0} value={l.warrantyMonths} onChange={(e) => updateLine(l.key, { warrantyMonths: Number(e.target.value) || 0 })} className="w-16 border rounded-sm py-1 px-1.5 text-right text-sm" style={{ borderColor: LINE }} /></td>
                                        <td className="py-2 px-2"><input value={l.condition} onChange={(e) => updateLine(l.key, { condition: e.target.value })} placeholder="VD: không lên nguồn…" className="w-40 border rounded-sm py-1 px-1.5 text-xs" style={{ borderColor: LINE }} /></td>
                                        <td className="py-2 px-2 text-right"><button onClick={() => removeLine(l.key)} className="p-1 rounded-sm hover:bg-black/5 opacity-50"><X size={13} /></button></td>
                                    </tr>
                                ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Thêm sản phẩm thủ công (không qua đơn hàng)</p>
                    <ProductPicker products={products} onPick={addManualLine} />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-5">
                    <Field label="Ngày nhận"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.receivedDate} onChange={(e) => setForm({ ...form, receivedDate: e.target.value })} /></Field>
                    <Field label="Ngày hẹn trả"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.returnDate} onChange={(e) => setForm({ ...form, returnDate: e.target.value })} /></Field>
                </div>
                <Field label="Địa chỉ nhận bảo hành"><p className="text-sm py-1.5 px-1 opacity-70">{COMPANY_INFO.address}</p></Field>
                <Field label="Ghi chú chung (không bắt buộc)">
                    <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
                </Field>

                <button onClick={submitTicket} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>Tạo phiếu bảo hành</button>
            </div>
        );
    }

    return (
        <div>
            {warrantyStockList.length > 0 && (
                <div className="p-4 rounded-sm mb-5" style={{ background: `${BRASS}0A`, border: `1px solid ${BRASS}44` }}>
                    <button onClick={() => setShowWarrantyStock((v) => !v)} className="flex items-center justify-between w-full text-left">
                        <span className="text-xs uppercase tracking-wider" style={{ color: BRASS }}>Kho bảo hành — đang giữ, chờ xử lý với NCC ({warrantyStockList.reduce((s, w) => s + w.qty, 0)} sản phẩm)</span>
                        <span className="text-xs opacity-60">{showWarrantyStock ? "Thu gọn ▲" : "Xem chi tiết ▼"}</span>
                    </button>
                    {showWarrantyStock && (
                        <div className="mt-3 rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                            <table className="w-full text-sm" style={{ minWidth: 560 }}>
                                <thead><tr style={{ background: PAPER }}>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Serial</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">SL</th>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Nguồn</th>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Phiếu BH</th>
                                </tr></thead>
                                <tbody>
                                {warrantyStockList.map((w) => (
                                    <tr key={w.key} style={{ borderTop: `1px dashed ${LINE}` }}>
                                        <td className="py-2 px-2">{w.productName}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {w.productCode}</span></td>
                                        <td className="py-2 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{w.serial || "—"}</td>
                                        <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{w.qty}</td>
                                        <td className="py-2 px-2 whitespace-nowrap">
                        <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: w.source === "supplier" ? `${FOREST}1A` : `${BRASS}1A`, color: w.source === "supplier" ? FOREST : BRASS }}>
                          {w.source === "supplier" ? "NCC trả lại" : "Khách mang tới"}
                        </span>
                                        </td>
                                        <td className="py-2 px-2">
                                            <button onClick={() => openViewingTicket(w.ticket)} className="hover:underline" style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>{w.ticket.code}</button>
                                        </td>
                                    </tr>
                                ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            )}

            <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                <div className="flex gap-1.5 flex-wrap">
                    <FilterChip active={statusFilter === "all"} onClick={() => setStatusFilter("all")}>Tất cả</FilterChip>
                    {WARRANTY_TICKET_STATUSES.map((s) => <FilterChip key={s.id} active={statusFilter === s.id} onClick={() => setStatusFilter(s.id)} color={s.color}>{s.label}</FilterChip>)}
                </div>
                <button onClick={() => { resetForm(); setView("new"); }} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}>
                    <Plus size={15} /> Tạo phiếu bảo hành
                </button>
            </div>

            {filteredTickets.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-16">Chưa có phiếu bảo hành nào.</p>
            ) : (
                <div className="space-y-2.5">
                    {filteredTickets.map((t) => {
                        const st = WARRANTY_TICKET_STATUSES.find((s) => s.id === t.status);
                        const relatedOrders = [...new Set(t.items.map((it) => it.orderCode).filter(Boolean))];
                        return (
                            <button key={t.id} onClick={() => openViewingTicket(t)} className="w-full text-left p-4 rounded-sm hover:bg-black/[0.02]" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <div className="flex items-center justify-between gap-3 flex-wrap">
                                    <div>
                                        <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{t.code}</p>
                                        <p className="text-xs opacity-50">{t.customerName} · {t.items.length} sản phẩm · {formatDateTime(t.createdAt)}{relatedOrders.length > 0 ? ` · Đơn: ${relatedOrders.join(", ")}` : ""}</p>
                                    </div>
                                    <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${st.color}1A`, color: st.color }}>{st.label}</span>
                                </div>
                            </button>
                        );
                    })}
                </div>
            )}

            {viewingTicket && (
                <Modal title={`Phiếu bảo hành ${viewingTicket.code}`} onClose={() => { setViewingTicket(null); setPrintBlockedUrl(null); }} size="xl">
                    {(() => {
                        const STATUS_RANK = { pending: 0, confirmed: 1, completed: 2, rejected: 2 };
                        const isLocked = viewingTicket.status === "completed" || viewingTicket.status === "rejected";
                        return (
                            <div className="flex items-center gap-2 mb-4 flex-wrap">
                                {WARRANTY_TICKET_STATUSES.map((s) => {
                                    const isCurrent = viewingTicket.status === s.id;
                                    const blocked = isLocked ? !isCurrent : STATUS_RANK[s.id] < STATUS_RANK[viewingTicket.status];
                                    return (
                                        <button key={s.id} disabled={blocked} onClick={() => {
                                            if (s.id === "confirmed" || s.id === "rejected") setStatusDraft(statusDraft === s.id ? null : s.id);
                                            else setTicketStatusDirect(viewingTicket, s.id);
                                        }} className="text-xs px-3 py-1.5 rounded-full border disabled:opacity-35 disabled:cursor-not-allowed"
                                                style={{ borderColor: (isCurrent || statusDraft === s.id) ? s.color : LINE, background: (isCurrent || statusDraft === s.id) ? `${s.color}1A` : "transparent", color: (isCurrent || statusDraft === s.id) ? s.color : INK }}>
                                            {s.label}
                                        </button>
                                    );
                                })}
                                {isLocked && (
                                    <button onClick={() => reopenTicket(viewingTicket)} title="Lỡ tay bấm nhầm hoặc cần sửa lại thông tin xử lý — mở lại phiếu để chỉnh sửa"
                                            className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: BRASS, color: BRASS }}>
                                        <RotateCcw size={12} /> Mở lại phiếu để sửa
                                    </button>
                                )}
                                {viewingTicket.status === "pending" && (
                                    <>
                                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="ml-auto border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                                            <option value="A5">Khổ A5</option>
                                            <option value="A4">Khổ A4</option>
                                        </select>
                                        <button onClick={() => doPrint(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                                            <Printer size={13} /> In phiếu tiếp nhận BH
                                        </button>
                                    </>
                                )}
                                {viewingTicket.status === "completed" && (
                                    <>
                                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="ml-auto border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                                            <option value="A5">Khổ A5</option>
                                            <option value="A4">Khổ A4</option>
                                        </select>
                                        <button onClick={() => doPrintXTBH(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                                            <Printer size={13} /> In phiếu xuất trả BH
                                        </button>
                                    </>
                                )}
                                {viewingTicket.status === "rejected" && (
                                    <>
                                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="ml-auto border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                                            <option value="A5">Khổ A5</option>
                                            <option value="A4">Khổ A4</option>
                                        </select>
                                        <button onClick={() => doPrintReject(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: RUST, color: RUST }}>
                                            <Printer size={13} /> In phiếu từ chối BH
                                        </button>
                                    </>
                                )}
                            </div>
                        );
                    })()}

                    {viewingTicket.xtbhCode && (
                        <p className="text-xs mb-4"><span className="opacity-50">Phiếu xuất trả bảo hành: </span><b style={{ color: FOREST, fontFamily: "'IBM Plex Mono', monospace" }}>{viewingTicket.xtbhCode}</b></p>
                    )}

                    {(viewingTicket.status === "pending" || viewingTicket.status === "confirmed") && (
                        <div className="mb-4">
                            <button onClick={() => setAddingReplacement((v) => !v)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                                <ArrowDownToLine size={13} /> Nhập hàng thay thế NCC vào Kho bảo hành
                            </button>
                            {addingReplacement && (
                                <div className="p-3 rounded-sm mt-2" style={{ background: `${FOREST}0A`, border: `1px solid ${FOREST}33` }}>
                                    <p className="text-xs opacity-60 mb-2">Ghi nhận sản phẩm NCC đã trả lại (sửa xong hoặc hàng thay thế) — nhập thẳng vào Kho bảo hành, chưa động tới kho chính.</p>
                                    <ProductPicker products={products} onPick={(id) => setReplProductId(id)} />
                                    {replProductId && <p className="text-xs opacity-70 mt-1.5">Sản phẩm: <b>{products.find((p) => p.id === replProductId)?.name}</b></p>}
                                    <div className="grid grid-cols-2 gap-3 mt-2">
                                        <label className="text-xs">
                                            <span className="block opacity-60 mb-1">Serial (nếu có)</span>
                                            <input value={replSeriesText} onChange={(e) => setReplSeriesText(e.target.value)} placeholder="Cách nhau bởi dấu phẩy" className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                                        </label>
                                        <label className="text-xs">
                                            <span className="block opacity-60 mb-1">Số lượng</span>
                                            <input type="number" min={1} value={replQty} onChange={(e) => setReplQty(Number(e.target.value) || 1)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                                        </label>
                                    </div>
                                    <button onClick={() => addReplacement(viewingTicket)} className="w-full py-2 rounded-sm text-white text-sm mt-3" style={{ background: FOREST }}>Ghi nhận vào Kho bảo hành</button>
                                </div>
                            )}
                            {viewingTicket.replacementReceived && viewingTicket.replacementReceived.length > 0 && (
                                <div className="mt-2 space-y-1">
                                    {viewingTicket.replacementReceived.map((r) => (
                                        <p key={r.id} className="text-xs opacity-70">✓ Đã nhận từ NCC: <b>{r.productName}</b>{r.series.length ? ` · Serial: ${r.series.join(", ")}` : ""} · SL {r.qty}</p>
                                    ))}
                                </div>
                            )}
                        </div>
                    )}

                    {statusDraft === "confirmed" && (
                        <div className="p-4 rounded-sm mb-4" style={{ background: `${BLUE}0A`, border: `1px solid ${BLUE}44` }}>
                            <p className="text-xs uppercase tracking-wider mb-3" style={{ color: BLUE }}>Xác nhận xử lý bảo hành</p>
                            <div className="flex gap-2 mb-3">
                                <button onClick={() => setResolutionType("exchange")} className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: resolutionType === "exchange" ? INK : LINE, background: resolutionType === "exchange" ? INK : "transparent", color: resolutionType === "exchange" ? "#fff" : INK }}>
                                    Đổi sản phẩm
                                </button>
                                <button onClick={() => setResolutionType("refund")} className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: resolutionType === "refund" ? INK : LINE, background: resolutionType === "refund" ? INK : "transparent", color: resolutionType === "refund" ? "#fff" : INK }}>
                                    Hoàn tiền
                                </button>
                            </div>
                            {resolutionType === "exchange" ? (
                                <div className="space-y-2.5">
                                    <p className="text-[10px] uppercase tracking-wider opacity-50">Nguồn hàng đổi</p>
                                    <div className="flex gap-2 mb-1">
                                        <button onClick={() => setExchangeSource("main")} className="px-3 py-1.5 rounded-sm text-xs border"
                                                style={{ borderColor: exchangeSource === "main" ? INK : LINE, background: exchangeSource === "main" ? INK : "transparent", color: exchangeSource === "main" ? "#fff" : INK }}>
                                            Kho chính (trả trước cho khách)
                                        </button>
                                        <button onClick={() => setExchangeSource("warranty")} className="px-3 py-1.5 rounded-sm text-xs border"
                                                style={{ borderColor: exchangeSource === "warranty" ? INK : LINE, background: exchangeSource === "warranty" ? INK : "transparent", color: exchangeSource === "warranty" ? "#fff" : INK }}>
                                            Kho bảo hành (NCC đã trả lại)
                                        </button>
                                    </div>

                                    {exchangeSource === "main" ? (
                                        <>
                                            <ProductPicker products={products} onPick={(id) => setExchangeProductId(id)} />
                                            {exchangeProductId && <p className="text-xs opacity-70">Sản phẩm đổi: <b>{products.find((p) => p.id === exchangeProductId)?.name}</b></p>}
                                            <div className="grid grid-cols-2 gap-3">
                                                <label className="text-xs">
                                                    <span className="block opacity-60 mb-1">Serial đổi (nếu có)</span>
                                                    <input value={exchangeSeriesText} onChange={(e) => setExchangeSeriesText(e.target.value)} placeholder="Cách nhau bởi dấu phẩy" className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                                                </label>
                                                <label className="text-xs">
                                                    <span className="block opacity-60 mb-1">Số lượng</span>
                                                    <input type="number" min={1} value={exchangeQty} onChange={(e) => setExchangeQty(Number(e.target.value) || 1)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                                                </label>
                                            </div>
                                        </>
                                    ) : (
                                        (viewingTicket.replacementReceived || []).length === 0 ? (
                                            <p className="text-xs" style={{ color: RUST }}>Chưa có hàng NCC trả lại cho phiếu này — bấm nút "Nhập hàng thay thế NCC" ở trên trước.</p>
                                        ) : (
                                            <div className="space-y-1.5">
                                                {viewingTicket.replacementReceived.map((r) => (
                                                    <label key={r.id} className="flex items-center gap-2 p-2 rounded-sm text-sm cursor-pointer" style={{ border: `1px solid ${selectedReplacementId === r.id ? INK : LINE}` }}>
                                                        <input type="radio" name="replacement" checked={selectedReplacementId === r.id} onChange={() => setSelectedReplacementId(r.id)} />
                                                        <span>{r.productName}{r.series.length ? ` · Serial: ${r.series.join(", ")}` : ""} · SL {r.qty}</span>
                                                    </label>
                                                ))}
                                            </div>
                                        )
                                    )}
                                </div>
                            ) : (
                                <label className="text-xs block">
                                    <span className="block opacity-60 mb-1">Số tiền hoàn lại cho khách</span>
                                    <MoneyInput className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} value={refundAmount} onChange={(v) => setRefundAmount(v)} />
                                </label>
                            )}
                            <div className="flex gap-2 mt-3">
                                <button onClick={() => setStatusDraft(null)} className="flex-1 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}>Huỷ</button>
                                <button onClick={() => saveConfirm(viewingTicket)} className="flex-1 py-2 rounded-sm text-white text-sm" style={{ background: BLUE }}>Lưu xác nhận</button>
                            </div>
                        </div>
                    )}

                    {statusDraft === "rejected" && (
                        <div className="p-4 rounded-sm mb-4" style={{ background: `${RUST}0A`, border: `1px solid ${RUST}44` }}>
                            <p className="text-xs uppercase tracking-wider mb-3" style={{ color: RUST }}>Từ chối bảo hành</p>
                            <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={rejectReason} onChange={(e) => setRejectReason(e.target.value)} placeholder="VD: sản phẩm ngoài phạm vi bảo hành, hết hạn bảo hành, lỗi do va đập/vào nước…" />
                            <div className="flex gap-2 mt-3">
                                <button onClick={() => setStatusDraft(null)} className="flex-1 py-2 rounded-sm text-sm border" style={{ borderColor: LINE, color: INK }}>Huỷ</button>
                                <button onClick={() => saveReject(viewingTicket)} className="flex-1 py-2 rounded-sm text-white text-sm" style={{ background: RUST }}>Lưu từ chối</button>
                            </div>
                        </div>
                    )}

                    {!statusDraft && viewingTicket.status === "confirmed" && (
                        <div className="p-3 rounded-sm mb-4 text-sm" style={{ background: `${BLUE}0A`, border: `1px solid ${BLUE}33` }}>
                            {viewingTicket.resolutionType === "exchange"
                                ? <p><span className="opacity-60">Đã xác nhận đổi ({viewingTicket.exchangeSource === "warranty" ? "Kho bảo hành" : "Kho chính"}): </span><b>{viewingTicket.exchangeProductName}</b>{viewingTicket.exchangeSeries.length ? ` · Serial: ${viewingTicket.exchangeSeries.join(", ")}` : ""} · SL {viewingTicket.exchangeQty}</p>
                                : <p><span className="opacity-60">Đã xác nhận hoàn tiền: </span><b>{vnd(viewingTicket.refundAmount)}</b></p>}
                        </div>
                    )}
                    {!statusDraft && viewingTicket.status === "rejected" && (
                        <div className="p-3 rounded-sm mb-4 text-sm" style={{ background: `${RUST}0A`, border: `1px solid ${RUST}33` }}>
                            <span className="opacity-60">Lý do từ chối: </span>{viewingTicket.rejectReason}
                        </div>
                    )}

                    {printBlockedUrl && (
                        <div className="mb-4 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                            <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới để mở phiếu, sau đó tự bấm in (Ctrl+P / biểu tượng in) trong tab mới đó:</p>
                            <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở phiếu bảo hành để in</a>
                        </div>
                    )}

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4 text-sm">
                        <div>
                            <p><span className="opacity-50">Khách hàng: </span>{viewingTicket.customerName}</p>
                            <p><span className="opacity-50">SĐT: </span>{viewingTicket.customerPhone || "—"}</p>
                            <p><span className="opacity-50">Địa chỉ: </span>{viewingTicket.customerAddress || "—"}</p>
                        </div>
                        <div>
                            <p><span className="opacity-50">Người lập: </span>{viewingTicket.createdBy}</p>
                            <p><span className="opacity-50">Ngày nhận: </span>{viewingTicket.receivedDate ? new Date(viewingTicket.receivedDate).toLocaleDateString("vi-VN") : "—"}</p>
                            <p><span className="opacity-50">Hẹn trả: </span>{viewingTicket.returnDate ? new Date(viewingTicket.returnDate).toLocaleDateString("vi-VN") : "—"}</p>
                        </div>
                    </div>

                    {(() => {
                        const relatedOrders = [...new Set(viewingTicket.items.map((it) => it.orderCode).filter(Boolean))];
                        if (relatedOrders.length === 0) return null;
                        return (
                            <p className="text-sm mb-4">
                                <span className="opacity-50">Đơn hàng liên quan: </span>
                                {relatedOrders.map((code, i) => (
                                    <span key={code}>
                    {i > 0 && ", "}
                                        {goToDoc ? <button onClick={() => goToDoc(code)} className="hover:underline font-medium" style={{ color: BLUE }}>{code}</button> : <span style={{ color: BLUE }}>{code}</span>}
                  </span>
                                ))}
                            </p>
                        );
                    })()}

                    <div className="rounded-sm overflow-x-auto mb-4" style={{ border: `1px solid ${LINE}` }}>
                        <table className="w-full text-sm" style={{ minWidth: 620 }}>
                            <thead><tr style={{ background: PAPER }}>
                                <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Đơn hàng</th>
                                <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Serial</th>
                                <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">SL</th>
                                <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Bảo hành</th>
                                <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Tình trạng</th>
                            </tr></thead>
                            <tbody>
                            {viewingTicket.items.map((it, i) => (
                                <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                    <td className="py-2 px-2">{it.productName}</td>
                                    <td className="py-2 px-2 whitespace-nowrap">
                                        {it.orderCode ? (goToDoc ? <button onClick={() => goToDoc(it.orderCode)} className="hover:underline" style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>{it.orderCode}</button> : <span style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>{it.orderCode}</span>) : <span className="opacity-40">—</span>}
                                    </td>
                                    <td className="py-2 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series.length ? it.series.join(", ") : "—"}</td>
                                    <td className="py-2 px-2 text-right">{it.qty}</td>
                                    <td className="py-2 px-2 text-right">{warrantyLabel(it.warrantyMonths)}</td>
                                    <td className="py-2 px-2">{it.condition || "—"}</td>
                                </tr>
                            ))}
                            </tbody>
                        </table>
                    </div>
                    {viewingTicket.note && <p className="text-sm opacity-70"><b>Ghi chú:</b> {viewingTicket.note}</p>}
                </Modal>
            )}
        </div>
    );
}

// Vận chuyển — quản lý vận đơn thủ công, liên kết với đơn hàng. Chưa kết nối API thật với các hãng vận chuyển (cần backend riêng), nhập tay mã vận đơn.
function Shipping({ shippingTickets, setShippingTickets, orders, customers, currentUser, addLog }) {
    const [view, setView] = useState("history");
    const [statusFilter, setStatusFilter] = useState("all");
    const [carrierFilter, setCarrierFilter] = useState("all");
    const [viewingTicket, setViewingTicket] = useState(null);
    const [printPaperSize, setPrintPaperSize] = useState("A5");
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);
    const [editingId, setEditingId] = useState(null);

    const emptyForm = () => ({
        orderId: "", orderCode: "", carrier: SHIPPING_CARRIERS[0], trackingCode: "",
        recipientName: "", recipientPhone: "", recipientAddress: "",
        packDate: todayISO(), pickupDate: "", deliveredDate: "", shippingFee: 0, codAmount: 0, note: "",
    });
    const [form, setForm] = useState(emptyForm());
    const [orderQuery, setOrderQuery] = useState("");

    const orderMatches = orderQuery.trim()
        ? orders.filter((o) => o.status !== "cancelled" && (o.code.toLowerCase().includes(orderQuery.trim().toLowerCase()) || (customers.find((c) => c.id === o.customerId)?.name || "").toLowerCase().includes(orderQuery.trim().toLowerCase()))).slice(0, 20)
        : [];

    const pickOrder = (order) => {
        const cust = customers.find((c) => c.id === order.customerId);
        const addr = order.shippingAddress;
        setForm((f) => ({
            ...f, orderId: order.id, orderCode: order.code,
            recipientName: addr?.recipientName || cust?.name || "",
            recipientPhone: addr?.recipientPhone || cust?.phone || "",
            recipientAddress: [addr?.addressDetail, addr?.ward, addr?.province].filter(Boolean).join(", ") || [cust?.addressDetail, cust?.ward, cust?.province].filter(Boolean).join(", "),
            codAmount: orderCalc(order).remaining > 0 ? orderCalc(order).remaining : 0,
        }));
        setOrderQuery("");
    };

    const openEditInfo = (ticket) => { setForm({ ...ticket }); setEditingId(ticket.id); setView("new"); };
    const submitTicket = () => {
        if (!form.recipientName.trim()) { alert("Vui lòng nhập tên người nhận."); return; }
        if (editingId) {
            setShippingTickets((prev) => prev.map((t) => (t.id === editingId ? normalizeShippingTicket({ ...t, ...form }) : t)));
            addLog("Sửa thông tin phiếu vận chuyển", `${form.orderCode || form.recipientName}`);
            setViewingTicket((v) => (v && v.id === editingId ? normalizeShippingTicket({ ...v, ...form }) : v));
        } else {
            const code = nextShippingCode(shippingTickets);
            const ticket = normalizeShippingTicket({ id: uid(), code, createdAt: new Date().toISOString(), createdBy: currentUser.fullName, status: "packing", ...form });
            setShippingTickets((prev) => [ticket, ...prev]);
            addLog("Tạo phiếu vận chuyển", `${code} · ${form.orderCode || ""} · ${form.recipientName}`);
        }
        setForm(emptyForm()); setEditingId(null); setOrderQuery(""); setView("history");
    };
    const setStatus = (ticket, status) => {
        const patch = { status };
        if (status === "shipping" && !ticket.pickupDate) patch.pickupDate = todayISO();
        if (status === "delivered" && !ticket.deliveredDate) patch.deliveredDate = todayISO();
        setShippingTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, ...patch } : t)));
        addLog("Cập nhật phiếu vận chuyển", `${ticket.code} → ${SHIPPING_STATUSES.find((s) => s.id === status)?.label}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, ...patch } : v));
    };
    const doPrint = (ticket) => {
        const html = buildShippingLabelHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };

    const filtered = shippingTickets.filter((t) => (statusFilter === "all" || t.status === statusFilter) && (carrierFilter === "all" || t.carrier === carrierFilter));

    if (view === "new") {
        return (
            <div>
                <button onClick={() => { setForm(emptyForm()); setEditingId(null); setView("history"); }} className="text-sm mb-4 opacity-60 hover:opacity-100">← Quay lại danh sách vận đơn</button>

                <div className="p-5 rounded-sm mb-5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Liên kết đơn hàng (không bắt buộc)</p>
                    <div className="relative mb-2">
                        <input value={orderQuery} onChange={(e) => setOrderQuery(e.target.value)} placeholder="Tìm theo mã đơn hoặc tên khách hàng…"
                               className="w-full border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }} />
                        {orderMatches.length > 0 && (
                            <div className="absolute z-20 mt-1 w-full max-h-52 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                                {orderMatches.map((o) => (
                                    <button key={o.id} onClick={() => pickOrder(o)} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex justify-between gap-3" style={{ borderBottom: `1px dashed ${LINE}` }}>
                                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{o.code}</span>
                                        <span className="opacity-60 truncate">{customers.find((c) => c.id === o.customerId)?.name || "Khách lẻ"}</span>
                                    </button>
                                ))}
                            </div>
                        )}
                    </div>
                    {form.orderCode && (
                        <div className="flex items-center justify-between p-2.5 rounded-sm text-sm" style={{ background: PAPER }}>
                            <span style={{ color: BLUE, fontFamily: "'IBM Plex Mono', monospace" }}>{form.orderCode}</span>
                            <button onClick={() => setForm({ ...form, orderId: "", orderCode: "" })} className="text-xs underline opacity-60">Bỏ chọn</button>
                        </div>
                    )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Tên người nhận"><input className={inputCls} style={{ borderColor: LINE }} value={form.recipientName} onChange={(e) => setForm({ ...form, recipientName: e.target.value })} /></Field>
                    <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} value={form.recipientPhone} onChange={(e) => setForm({ ...form, recipientPhone: e.target.value })} /></Field>
                </div>
                <Field label="Địa chỉ nhận hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.recipientAddress} onChange={(e) => setForm({ ...form, recipientAddress: e.target.value })} /></Field>

                <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Đơn vị vận chuyển">
                        <select className={inputCls} style={{ borderColor: LINE }} value={form.carrier} onChange={(e) => setForm({ ...form, carrier: e.target.value })}>
                            {SHIPPING_CARRIERS.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                    </Field>
                    <Field label="Mã vận đơn" hint="Copy từ app/web của đơn vị vận chuyển">
                        <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.trackingCode} onChange={(e) => setForm({ ...form, trackingCode: e.target.value })} />
                    </Field>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Phí giao hàng"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.shippingFee} onChange={(v) => setForm({ ...form, shippingFee: v })} /></Field>
                    <Field label="Tiền thu hộ (COD)"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.codAmount} onChange={(v) => setForm({ ...form, codAmount: v })} /></Field>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
                    <Field label="Ngày đóng gói"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.packDate} onChange={(e) => setForm({ ...form, packDate: e.target.value })} /></Field>
                    <Field label="Ngày lấy hàng"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.pickupDate} onChange={(e) => setForm({ ...form, pickupDate: e.target.value })} /></Field>
                    <Field label="Ngày giao hàng"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.deliveredDate} onChange={(e) => setForm({ ...form, deliveredDate: e.target.value })} /></Field>
                </div>
                <Field label="Ghi chú giao hàng">
                    <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="VD: cho xem hàng, không cho thử…" />
                </Field>
                <button onClick={submitTicket} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>{editingId ? "Lưu thay đổi" : "Tạo phiếu vận chuyển"}</button>
            </div>
        );
    }

    return (
        <div>
            <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                <div className="flex gap-2 flex-wrap items-center">
                    <div className="flex gap-1.5 flex-wrap">
                        <FilterChip active={statusFilter === "all"} onClick={() => setStatusFilter("all")}>Tất cả</FilterChip>
                        {SHIPPING_STATUSES.map((s) => <FilterChip key={s.id} active={statusFilter === s.id} onClick={() => setStatusFilter(s.id)} color={s.color}>{s.label}</FilterChip>)}
                    </div>
                    <select value={carrierFilter} onChange={(e) => setCarrierFilter(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                        <option value="all">Tất cả đối tác</option>
                        {SHIPPING_CARRIERS.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                </div>
                <button onClick={() => { setForm(emptyForm()); setEditingId(null); setView("new"); }} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}>
                    <Plus size={15} /> Tạo phiếu vận chuyển
                </button>
            </div>

            {filtered.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-16">Chưa có phiếu vận chuyển nào.</p>
            ) : (
                <div className="space-y-2.5">
                    {filtered.map((t) => {
                        const st = SHIPPING_STATUSES.find((s) => s.id === t.status);
                        return (
                            <button key={t.id} onClick={() => { setViewingTicket(t); setPrintBlockedUrl(null); }} className="w-full text-left p-4 rounded-sm hover:bg-black/[0.02]" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <div className="flex items-center justify-between gap-3 flex-wrap">
                                    <div>
                                        <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{t.code}{t.orderCode ? <span className="opacity-50 font-normal"> · Đơn {t.orderCode}</span> : null}</p>
                                        <p className="text-xs opacity-50">{t.recipientName} · {t.carrier}{t.trackingCode ? ` · ${t.trackingCode}` : ""} · {formatDateTime(t.createdAt)}</p>
                                    </div>
                                    <div className="flex items-center gap-2">
                                        {t.codAmount > 0 && <span className="text-[11px] px-2 py-1 rounded-sm" style={{ background: `${BRASS}1A`, color: BRASS }}>COD {vnd(t.codAmount)}</span>}
                                        <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${st.color}1A`, color: st.color }}>{st.label}</span>
                                    </div>
                                </div>
                            </button>
                        );
                    })}
                </div>
            )}

            {viewingTicket && (
                <Modal title={`Phiếu vận chuyển ${viewingTicket.code}`} onClose={() => { setViewingTicket(null); setPrintBlockedUrl(null); }} size="lg">
                    <div className="flex items-center gap-2 mb-4 flex-wrap">
                        {SHIPPING_STATUSES.map((s) => (
                            <button key={s.id} onClick={() => setStatus(viewingTicket, s.id)} className="text-xs px-3 py-1.5 rounded-full border"
                                    style={{ borderColor: viewingTicket.status === s.id ? s.color : LINE, background: viewingTicket.status === s.id ? `${s.color}1A` : "transparent", color: viewingTicket.status === s.id ? s.color : INK }}>
                                {s.label}
                            </button>
                        ))}
                    </div>
                    <div className="flex items-center gap-2 mb-4 flex-wrap">
                        <button onClick={() => { setViewingTicket(null); openEditInfo(viewingTicket); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: BLUE, color: BLUE }}>
                            <Pencil size={13} /> Sửa thông tin
                        </button>
                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="ml-auto border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                            <option value="A5">Khổ A5</option>
                            <option value="A4">Khổ A4</option>
                        </select>
                        <button onClick={() => doPrint(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                            <Printer size={13} /> In phiếu vận chuyển
                        </button>
                    </div>

                    {printBlockedUrl && (
                        <div className="mb-4 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                            <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới để mở phiếu, sau đó tự bấm in (Ctrl+P / biểu tượng in) trong tab mới đó:</p>
                            <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở phiếu vận chuyển để in</a>
                        </div>
                    )}

                    {viewingTicket.orderCode && <p className="text-sm mb-3"><span className="opacity-50">Đơn hàng: </span><b style={{ color: BLUE }}>{viewingTicket.orderCode}</b></p>}

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4 text-sm">
                        <div>
                            <p><span className="opacity-50">Người nhận: </span>{viewingTicket.recipientName}</p>
                            <p><span className="opacity-50">SĐT: </span>{viewingTicket.recipientPhone || "—"}</p>
                            <p><span className="opacity-50">Địa chỉ: </span>{viewingTicket.recipientAddress || "—"}</p>
                        </div>
                        <div>
                            <p><span className="opacity-50">Đơn vị VC: </span>{viewingTicket.carrier}</p>
                            <p><span className="opacity-50">Mã vận đơn: </span>{viewingTicket.trackingCode || "—"}</p>
                            <p><span className="opacity-50">Người lập: </span>{viewingTicket.createdBy}</p>
                        </div>
                    </div>

                    <div className="grid grid-cols-3 gap-4 mb-4 text-sm">
                        <div><p className="opacity-50 text-xs mb-1">Đóng gói</p><p>{viewingTicket.packDate ? new Date(viewingTicket.packDate).toLocaleDateString("vi-VN") : "—"}</p></div>
                        <div><p className="opacity-50 text-xs mb-1">Lấy hàng</p><p>{viewingTicket.pickupDate ? new Date(viewingTicket.pickupDate).toLocaleDateString("vi-VN") : "—"}</p></div>
                        <div><p className="opacity-50 text-xs mb-1">Giao hàng</p><p>{viewingTicket.deliveredDate ? new Date(viewingTicket.deliveredDate).toLocaleDateString("vi-VN") : "—"}</p></div>
                    </div>

                    <div className="grid grid-cols-2 gap-4 mb-4">
                        <div>
                            <p className="text-xs opacity-50 mb-1">Phí giao hàng</p>
                            <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingTicket.shippingFee)}</p>
                        </div>
                        <div>
                            <p className="text-xs opacity-50 mb-1">Tiền thu hộ (COD)</p>
                            <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: viewingTicket.codAmount > 0 ? BRASS : INK }}>{vnd(viewingTicket.codAmount)}</p>
                        </div>
                    </div>
                    {viewingTicket.note && <p className="text-sm opacity-70"><b>Ghi chú:</b> {viewingTicket.note}</p>}
                </Modal>
            )}
        </div>
    );
}

function Stocktake({ products, setProducts, stocktakes, setStocktakes, currentUser, addLog }) {
    const [view, setView] = useState("history"); // history | new
    const [viewingStocktake, setViewingStocktake] = useState(null);
    const [mode, setMode] = useState("series"); // series | barcode | manual
    const [scanInput, setScanInput] = useState("");
    const [lines, setLines] = useState({}); // { [productId]: { productId, productName, productCode, systemQty, countedQty, scannedSeries? } }
    const [warnings, setWarnings] = useState([]); // { id, type, serial, message }
    const [note, setNote] = useState("");
    const scanRef = useRef(null);

    const resetSession = () => { setLines({}); setWarnings([]); setNote(""); setScanInput(""); setMode("series"); };

    const handleScan = () => {
        const raw = scanInput.trim();
        if (!raw) return;
        setScanInput("");
        if (mode === "series") {
            const found = findSeriesOwner(products, raw);
            if (!found) {
                setWarnings((w) => [{ id: uid(), type: "unknown_series", serial: raw, message: `Series "${raw}" chưa có trong hệ thống — cần kiểm tra thủ công, không tính vào số đếm.` }, ...w]);
                return;
            }
            if (found.row.status === "Đã xuất") {
                setWarnings((w) => [{ id: uid(), type: "sold_but_scanned", serial: raw, message: `Series "${raw}" (${found.product.name}) — hệ thống ghi ĐÃ BÁN nhưng vẫn quét thấy thực tế. Cần kiểm tra thủ công, không tự cộng vào tồn.` }, ...w]);
                return;
            }
            const p = found.product;
            const existing = lines[p.id];
            if (existing && (existing.scannedSeries || []).includes(raw)) {
                setWarnings((w) => [{ id: uid(), type: "duplicate_scan", serial: raw, message: `Series "${raw}" đã quét trước đó trong phiên này — bỏ qua, không cộng thêm.` }, ...w]);
                return;
            }
            setLines((prev) => ({ ...prev, [p.id]: {
                    productId: p.id, productName: p.name, productCode: p.code, systemQty: productStats(p).closingQty,
                    countedQty: (existing?.countedQty || 0) + 1, scannedSeries: [...(existing?.scannedSeries || []), raw],
                } }));
        } else if (mode === "barcode") {
            const p = products.find((x) => x.barcode && x.barcode.trim() === raw);
            if (!p) {
                setWarnings((w) => [{ id: uid(), type: "unknown_barcode", serial: raw, message: `Không tìm thấy sản phẩm nào có mã vạch "${raw}".` }, ...w]);
                return;
            }
            const existing = lines[p.id];
            setLines((prev) => ({ ...prev, [p.id]: {
                    productId: p.id, productName: p.name, productCode: p.code, systemQty: productStats(p).closingQty, countedQty: (existing?.countedQty || 0) + 1,
                } }));
        }
        scanRef.current && scanRef.current.focus();
    };

    const addManualLine = (productId) => {
        const p = products.find((x) => x.id === productId);
        if (!p || lines[productId]) return;
        setLines((prev) => ({ ...prev, [productId]: { productId, productName: p.name, productCode: p.code, systemQty: productStats(p).closingQty, countedQty: productStats(p).closingQty } }));
    };
    const updateManualQty = (productId, qty) => setLines((prev) => ({ ...prev, [productId]: { ...prev[productId], countedQty: Math.max(0, qty) } }));
    const removeLine = (productId) => setLines((prev) => { const next = { ...prev }; delete next[productId]; return next; });

    const lineList = Object.values(lines).sort((a, b) => a.productName.localeCompare(b.productName));
    const diffCount = lineList.filter((l) => l.countedQty !== l.systemQty).length;

    const finishStocktake = () => {
        if (lineList.length === 0) { alert("Chưa kiểm sản phẩm nào."); return; }
        const code = nextStocktakeCode(stocktakes);
        setProducts((prev) => prev.map((p) => {
            const line = lines[p.id];
            if (!line) return p;
            const diff = line.countedQty - line.systemQty;
            if (diff === 0) return p;
            if (diff > 0) {
                return { ...p, movements: [...p.movements, { id: uid(), type: "in", docNo: code, date: todayISO(), qty: diff, price: p.costPrice || 0, series: [] }] };
            }
            let seriesToRemove = [];
            if (p.hasSeries) {
                const stillInStock = seriesList(p).filter((s) => s.status === "Còn tồn").map((s) => s.serial);
                const scanned = line.scannedSeries || [];
                seriesToRemove = stillInStock.filter((s) => !scanned.includes(s)).slice(0, Math.abs(diff));
            }
            return { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: code, date: todayISO(), qty: Math.abs(diff), price: p.costPrice || 0, series: seriesToRemove }] };
        }));
        const record = normalizeStocktake({
            id: uid(), code, createdAt: new Date().toISOString(), createdBy: currentUser.fullName, note,
            lines: lineList.map((l) => ({ productId: l.productId, productName: l.productName, productCode: l.productCode, systemQty: l.systemQty, countedQty: l.countedQty, diff: l.countedQty - l.systemQty })),
            warnings: warnings.map((w) => w.message),
        });
        setStocktakes((prev) => [record, ...prev]);
        addLog("Kiểm kho", `${code} · ${lineList.length} sản phẩm · ${diffCount} sản phẩm lệch`);
        resetSession();
        setView("history");
    };

    if (view === "history") {
        return (
            <div>
                <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                    <p className="text-sm opacity-60">{stocktakes.length} phiếu kiểm kê</p>
                    <button onClick={() => { resetSession(); setView("new"); }} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
                        <Plus size={15} /> Bắt đầu kiểm kê mới
                    </button>
                </div>
                {stocktakes.length === 0 ? (
                    <p className="text-sm opacity-50 text-center py-16">Chưa có phiếu kiểm kê nào.</p>
                ) : (
                    <div className="space-y-2.5">
                        {stocktakes.map((s) => {
                            const diffLines = s.lines.filter((l) => l.diff !== 0);
                            return (
                                <button key={s.id} onClick={() => setViewingStocktake(s)} className="w-full text-left p-4 rounded-sm hover:bg-black/[0.02]" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                    <div className="flex items-center justify-between gap-3 flex-wrap">
                                        <div>
                                            <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{s.code}</p>
                                            <p className="text-xs opacity-50">{formatDateTime(s.createdAt)} · {s.createdBy}</p>
                                        </div>
                                        <div className="text-right">
                                            <p className="text-sm" style={{ color: INK }}>{s.lines.length} sản phẩm kiểm</p>
                                            <p className="text-xs" style={{ color: diffLines.length > 0 ? RUST : FOREST }}>{diffLines.length > 0 ? `${diffLines.length} sản phẩm lệch` : "Khớp sổ sách"}{s.warnings.length > 0 ? ` · ${s.warnings.length} cảnh báo` : ""}</p>
                                        </div>
                                    </div>
                                </button>
                            );
                        })}
                    </div>
                )}

                {viewingStocktake && (
                    <Modal title={`Phiếu kiểm kê ${viewingStocktake.code}`} onClose={() => setViewingStocktake(null)} size="xl">
                        <p className="text-xs opacity-50 mb-4">{formatDateTime(viewingStocktake.createdAt)} · Thực hiện bởi {viewingStocktake.createdBy}{viewingStocktake.note ? ` · Ghi chú: ${viewingStocktake.note}` : ""}</p>
                        <div className="rounded-sm overflow-x-auto mb-4" style={{ border: `1px solid ${LINE}` }}>
                            <table className="w-full text-sm" style={{ minWidth: 480 }}>
                                <thead><tr style={{ background: PAPER }}>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sổ sách</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Đã đếm</th>
                                    <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Chênh lệch</th>
                                </tr></thead>
                                <tbody>
                                {viewingStocktake.lines.map((l, i) => (
                                    <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                        <td className="py-2 px-2">{l.productName}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {l.productCode}</span></td>
                                        <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{l.systemQty}</td>
                                        <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{l.countedQty}</td>
                                        <td className="py-2 px-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: l.diff === 0 ? INK : l.diff > 0 ? FOREST : RUST }}>{l.diff > 0 ? `+${l.diff}` : l.diff}</td>
                                    </tr>
                                ))}
                                </tbody>
                            </table>
                        </div>
                        {viewingStocktake.warnings.length > 0 && (
                            <div className="p-3 rounded-sm" style={{ background: `${RUST}0D`, border: `1px solid ${RUST}44` }}>
                                <p className="text-xs uppercase tracking-wider mb-2" style={{ color: RUST }}>Cảnh báo trong phiên kiểm kê này</p>
                                <ul className="space-y-1">
                                    {viewingStocktake.warnings.map((w, i) => <li key={i} className="text-xs opacity-80">• {w}</li>)}
                                </ul>
                            </div>
                        )}
                    </Modal>
                )}
            </div>
        );
    }

    // view === "new"
    return (
        <div>
            <button onClick={() => { resetSession(); setView("history"); }} className="text-sm mb-4 opacity-60 hover:opacity-100">← Quay lại lịch sử kiểm kê</button>

            <div className="p-5 rounded-sm mb-5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Chế độ đếm</p>
                <div className="flex gap-2 mb-4 flex-wrap">
                    {[{ id: "series", label: "Quét Series" }, { id: "barcode", label: "Quét Barcode" }, { id: "manual", label: "Nhập tay số lượng" }].map((m) => (
                        <button key={m.id} onClick={() => { setMode(m.id); setScanInput(""); }} className="px-3.5 py-1.5 rounded-full text-sm border"
                                style={{ borderColor: mode === m.id ? INK : LINE, background: mode === m.id ? INK : "transparent", color: mode === m.id ? "#fff" : INK }}>
                            {m.label}
                        </button>
                    ))}
                </div>

                {(mode === "series" || mode === "barcode") && (
                    <div>
                        <input ref={scanRef} autoFocus value={scanInput} onChange={(e) => setScanInput(e.target.value)}
                               onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); handleScan(); } }}
                               placeholder={mode === "series" ? "Quét hoặc gõ số series rồi Enter…" : "Quét hoặc gõ mã vạch rồi Enter…"}
                               className="w-full border rounded-sm py-2.5 px-3 text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                        <p className="text-xs opacity-50 mt-1.5">{mode === "series" ? "Mỗi số series quét được sẽ cộng 1 vào tồn đếm của đúng sản phẩm sở hữu series đó." : "Mỗi lần quét trùng 1 mã vạch sẽ cộng thêm 1 vào số lượng đếm của sản phẩm có mã vạch đó."}</p>
                    </div>
                )}
                {mode === "manual" && (
                    <ProductPicker products={products} onPick={addManualLine} />
                )}
            </div>

            {warnings.length > 0 && (
                <div className="p-4 rounded-sm mb-5" style={{ background: `${RUST}0D`, border: `1px solid ${RUST}44` }}>
                    <p className="text-xs uppercase tracking-wider mb-2" style={{ color: RUST }}>Cảnh báo cần kiểm tra thủ công ({warnings.length})</p>
                    <ul className="space-y-1 max-h-40 overflow-y-auto">
                        {warnings.map((w) => <li key={w.id} className="text-xs opacity-80">• {w.message}</li>)}
                    </ul>
                </div>
            )}

            <div className="p-5 rounded-sm mb-5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <div className="flex items-center justify-between mb-3">
                    <p className="text-xs uppercase tracking-wider opacity-60">Kết quả kiểm kê ({lineList.length} sản phẩm{diffCount > 0 ? ` · ${diffCount} lệch` : ""})</p>
                </div>
                {lineList.length === 0 ? (
                    <p className="text-sm opacity-50 text-center py-10">Chưa có sản phẩm nào được đếm — bắt đầu quét hoặc chọn sản phẩm ở trên.</p>
                ) : (
                    <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
                        <table className="w-full text-sm" style={{ minWidth: 520 }}>
                            <thead><tr style={{ background: PAPER }}>
                                <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sổ sách</th>
                                <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Đã đếm</th>
                                <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Chênh lệch</th>
                                <th className="py-2 px-2"></th>
                            </tr></thead>
                            <tbody>
                            {lineList.map((l) => {
                                const diff = l.countedQty - l.systemQty;
                                return (
                                    <tr key={l.productId} style={{ borderTop: `1px dashed ${LINE}` }}>
                                        <td className="py-2 px-2">{l.productName}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {l.productCode}</span></td>
                                        <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{l.systemQty}</td>
                                        <td className="py-2 px-2 text-right">
                                            {mode === "manual" ? (
                                                <input type="number" min={0} value={l.countedQty} onChange={(e) => updateManualQty(l.productId, Number(e.target.value) || 0)}
                                                       className="w-20 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                            ) : (
                                                <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{l.countedQty}</span>
                                            )}
                                        </td>
                                        <td className="py-2 px-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: diff === 0 ? INK : diff > 0 ? FOREST : RUST }}>{diff > 0 ? `+${diff}` : diff}</td>
                                        <td className="py-2 px-2 text-right"><button onClick={() => removeLine(l.productId)} className="p-1 rounded-sm hover:bg-black/5 opacity-50"><X size={13} /></button></td>
                                    </tr>
                                );
                            })}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>

            <Field label="Ghi chú / lý do chênh lệch (không bắt buộc)">
                <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={note} onChange={(e) => setNote(e.target.value)} placeholder="VD: kiểm kê định kỳ cuối tháng, phát hiện thất thoát khu vực kho A…" />
            </Field>

            <button onClick={finishStocktake} disabled={lineList.length === 0} className="w-full py-2.5 rounded-sm text-white text-sm mt-2 disabled:opacity-40" style={{ background: INK }}>
                Hoàn tất kiểm kê & tạo phiếu điều chỉnh
            </button>
        </div>
    );
}

// Phiếu dịch vụ Sửa chữa — tiếp nhận thiết bị/sản phẩm khách mang tới sửa (không nhất thiết là bảo hành), theo dõi tiến độ và chi phí.
function RepairTickets({ repairTickets, setRepairTickets, currentUser, addLog }) {
    const [view, setView] = useState("history");
    const [statusFilter, setStatusFilter] = useState("all");
    const [viewingTicket, setViewingTicket] = useState(null);
    const [printPaperSize, setPrintPaperSize] = useState("A5");
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);
    const [editingId, setEditingId] = useState(null); // id phiếu đang sửa thông tin, null = đang tạo mới
    const emptyForm = () => ({
        customerName: "", customerPhone: "", customerAddress: "", deviceName: "", deviceBrand: "", serial: "", issueDescription: "",
        receivedDate: todayISO(), returnDate: "", estimatedCost: 0, actualCost: 0, vat: "VAT8", note: "",
    });
    const [form, setForm] = useState(emptyForm());

    const openEditInfo = (ticket) => {
        setForm({ ...ticket });
        setEditingId(ticket.id);
        setView("new");
    };
    const submitTicket = () => {
        if (!form.customerName.trim()) { alert("Vui lòng nhập tên khách hàng."); return; }
        if (!form.deviceName.trim()) { alert("Vui lòng nhập tên thiết bị/sản phẩm cần sửa."); return; }
        if (editingId) {
            setRepairTickets((prev) => prev.map((t) => (t.id === editingId ? normalizeRepairTicket({ ...t, ...form }) : t)));
            addLog("Sửa thông tin phiếu sửa chữa", `${form.customerName} · ${form.deviceName}`);
            setViewingTicket((v) => (v && v.id === editingId ? normalizeRepairTicket({ ...v, ...form }) : v));
        } else {
            const code = nextRepairCode(repairTickets);
            const ticket = normalizeRepairTicket({ id: uid(), code, createdAt: new Date().toISOString(), createdBy: currentUser.fullName, status: "received", ...form });
            setRepairTickets((prev) => [ticket, ...prev]);
            addLog("Tạo phiếu dịch vụ sửa chữa", `${code} · ${form.customerName} · ${form.deviceName}`);
        }
        setForm(emptyForm());
        setEditingId(null);
        setView("history");
    };
    const setStatus = (ticket, status) => {
        setRepairTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, status } : t)));
        addLog("Cập nhật phiếu sửa chữa", `${ticket.code} → ${REPAIR_STATUSES.find((s) => s.id === status)?.label}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, status } : v));
    };
    const updateTicketField = (ticket, field, value) => {
        setRepairTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, [field]: value } : t)));
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, [field]: value } : v));
    };
    const doPrintReceipt = (ticket) => {
        const html = buildRepairReceiptHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };

    const filtered = repairTickets.filter((t) => statusFilter === "all" || t.status === statusFilter);

    if (view === "new") {
        return (
            <div>
                <button onClick={() => { setForm(emptyForm()); setEditingId(null); setView("history"); }} className="text-sm mb-4 opacity-60 hover:opacity-100">← Quay lại danh sách phiếu sửa chữa</button>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Tên khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} /></Field>
                    <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerPhone} onChange={(e) => setForm({ ...form, customerPhone: e.target.value })} /></Field>
                </div>
                <Field label="Địa chỉ khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerAddress} onChange={(e) => setForm({ ...form, customerAddress: e.target.value })} /></Field>
                <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Tên thiết bị / sản phẩm"><input className={inputCls} style={{ borderColor: LINE }} value={form.deviceName} onChange={(e) => setForm({ ...form, deviceName: e.target.value })} placeholder="VD: PC Gaming, Laptop Dell..." /></Field>
                    <Field label="Hãng / thương hiệu"><input className={inputCls} style={{ borderColor: LINE }} value={form.deviceBrand} onChange={(e) => setForm({ ...form, deviceBrand: e.target.value })} /></Field>
                </div>
                <Field label="Serial / IMEI (nếu có)"><input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.serial} onChange={(e) => setForm({ ...form, serial: e.target.value })} /></Field>
                <Field label="Mô tả tình trạng / lỗi">
                    <textarea rows={3} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.issueDescription} onChange={(e) => setForm({ ...form, issueDescription: e.target.value })} placeholder="VD: không lên nguồn, màn hình bị sọc..." />
                </Field>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Ngày nhận"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.receivedDate} onChange={(e) => setForm({ ...form, receivedDate: e.target.value })} /></Field>
                    <Field label="Ngày hẹn trả"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.returnDate} onChange={(e) => setForm({ ...form, returnDate: e.target.value })} /></Field>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Chi phí dự kiến"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.estimatedCost} onChange={(v) => setForm({ ...form, estimatedCost: v })} /></Field>
                    <Field label="Chi phí thực tế (nếu đã có)"><MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.actualCost} onChange={(v) => setForm({ ...form, actualCost: v })} /></Field>
                </div>
                <Field label="VAT" hint="Chi phí trên đã bao gồm VAT theo mức chọn">
                    <div className="flex gap-2">
                        {VAT_OPTIONS.map((v) => (
                            <button key={v.id} type="button" onClick={() => setForm({ ...form, vat: v.id })} className="px-3.5 py-1.5 rounded-sm text-sm border"
                                    style={{ borderColor: form.vat === v.id ? INK : LINE, background: form.vat === v.id ? INK : "transparent", color: form.vat === v.id ? "#fff" : INK }}>
                                {v.label}
                            </button>
                        ))}
                    </div>
                </Field>
                <Field label="Ghi chú">
                    <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
                </Field>
                <button onClick={submitTicket} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>{editingId ? "Lưu thay đổi" : "Tạo phiếu dịch vụ sửa chữa"}</button>
            </div>
        );
    }

    return (
        <div>
            <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                <div className="flex gap-1.5 flex-wrap">
                    <FilterChip active={statusFilter === "all"} onClick={() => setStatusFilter("all")}>Tất cả</FilterChip>
                    {REPAIR_STATUSES.map((s) => <FilterChip key={s.id} active={statusFilter === s.id} onClick={() => setStatusFilter(s.id)} color={s.color}>{s.label}</FilterChip>)}
                </div>
                <button onClick={() => { setForm(emptyForm()); setView("new"); }} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}>
                    <Plus size={15} /> Tạo phiếu sửa chữa
                </button>
            </div>
            {filtered.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-16">Chưa có phiếu dịch vụ sửa chữa nào.</p>
            ) : (
                <div className="space-y-2.5">
                    {filtered.map((t) => {
                        const st = REPAIR_STATUSES.find((s) => s.id === t.status);
                        return (
                            <button key={t.id} onClick={() => setViewingTicket(t)} className="w-full text-left p-4 rounded-sm hover:bg-black/[0.02]" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <div className="flex items-center justify-between gap-3 flex-wrap">
                                    <div>
                                        <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{t.code}</p>
                                        <p className="text-xs opacity-50">{t.customerName} · {t.deviceName} · {formatDateTime(t.createdAt)}</p>
                                    </div>
                                    <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${st.color}1A`, color: st.color }}>{st.label}</span>
                                </div>
                            </button>
                        );
                    })}
                </div>
            )}

            {viewingTicket && (
                <Modal title={`Phiếu sửa chữa ${viewingTicket.code}`} onClose={() => { setViewingTicket(null); setPrintBlockedUrl(null); }} size="lg">
                    <div className="flex items-center gap-2 mb-4 flex-wrap">
                        {REPAIR_STATUSES.map((s) => (
                            <button key={s.id} onClick={() => setStatus(viewingTicket, s.id)} className="text-xs px-3 py-1.5 rounded-full border"
                                    style={{ borderColor: viewingTicket.status === s.id ? s.color : LINE, background: viewingTicket.status === s.id ? `${s.color}1A` : "transparent", color: viewingTicket.status === s.id ? s.color : INK }}>
                                {s.label}
                            </button>
                        ))}
                        <button onClick={() => { setViewingTicket(null); openEditInfo(viewingTicket); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border ml-auto" style={{ borderColor: BLUE, color: BLUE }}>
                            <Pencil size={13} /> Sửa thông tin
                        </button>
                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                            <option value="A5">Khổ A5</option>
                            <option value="A4">Khổ A4</option>
                        </select>
                        <button onClick={() => doPrintReceipt(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                            <Printer size={13} /> {viewingTicket.status === "received" ? "In phiếu tiếp nhận & báo giá" : "In phiếu báo giá thực tế"}
                        </button>
                    </div>

                    {printBlockedUrl && (
                        <div className="mb-4 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                            <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới để mở phiếu, sau đó tự bấm in (Ctrl+P / biểu tượng in) trong tab mới đó:</p>
                            <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở phiếu tiếp nhận để in</a>
                        </div>
                    )}

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4 text-sm">
                        <div>
                            <p><span className="opacity-50">Khách hàng: </span>{viewingTicket.customerName}</p>
                            <p><span className="opacity-50">SĐT: </span>{viewingTicket.customerPhone || "—"}</p>
                            <p><span className="opacity-50">Địa chỉ: </span>{viewingTicket.customerAddress || "—"}</p>
                        </div>
                        <div>
                            <p><span className="opacity-50">Người lập: </span>{viewingTicket.createdBy}</p>
                            <p><span className="opacity-50">Ngày nhận: </span>{viewingTicket.receivedDate ? new Date(viewingTicket.receivedDate).toLocaleDateString("vi-VN") : "—"}</p>
                            <p><span className="opacity-50">Hẹn trả: </span>{viewingTicket.returnDate ? new Date(viewingTicket.returnDate).toLocaleDateString("vi-VN") : "—"}</p>
                        </div>
                    </div>
                    <div className="p-3 rounded-sm mb-4 text-sm" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                        <p><span className="opacity-50">Thiết bị: </span><b>{viewingTicket.deviceName}</b>{viewingTicket.deviceBrand ? ` · ${viewingTicket.deviceBrand}` : ""}{viewingTicket.serial ? ` · SN: ${viewingTicket.serial}` : ""}</p>
                        <p className="mt-1"><span className="opacity-50">Tình trạng/lỗi: </span>{viewingTicket.issueDescription || "—"}</p>
                    </div>
                    <div className="grid grid-cols-2 gap-4 mb-2">
                        <div>
                            <p className="text-xs opacity-50 mb-1">Chi phí dự kiến</p>
                            <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingTicket.estimatedCost)}</p>
                        </div>
                        <div>
                            <p className="text-xs opacity-50 mb-1">Chi phí thực tế{viewingTicket.status !== "received" ? " (có thể sửa)" : ""}</p>
                            {viewingTicket.status !== "received" ? (
                                <MoneyInput className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={viewingTicket.actualCost} onChange={(v) => updateTicketField(viewingTicket, "actualCost", v)} />
                            ) : (
                                <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: viewingTicket.actualCost > 0 ? FOREST : INK }}>{vnd(viewingTicket.actualCost)}</p>
                            )}
                        </div>
                    </div>
                    <p className="text-xs opacity-50 mb-4">
                        VAT: <b>{VAT_OPTIONS.find((v) => v.id === viewingTicket.vat)?.label}</b>
                        {viewingTicket.actualCost > 0 && viewingTicket.vat !== "KCT" && ` · trong đó thuế GTGT: ${vnd(Math.round(viewingTicket.actualCost * vatPercent(viewingTicket.vat) / (100 + vatPercent(viewingTicket.vat))))}`}
                    </p>
                    {viewingTicket.note && <p className="text-sm opacity-70"><b>Ghi chú:</b> {viewingTicket.note}</p>}
                </Modal>
            )}
        </div>
    );
}

// Phiếu dịch vụ IT Helpdesk — tiếp nhận/xử lý các yêu cầu hỗ trợ kỹ thuật (cài đặt phần mềm, sự cố phần cứng, mạng...).
function HelpdeskTickets({ helpdeskTickets, setHelpdeskTickets, employeeNames, currentUser, addLog }) {
    const [view, setView] = useState("history");
    const [statusFilter, setStatusFilter] = useState("all");
    const [viewingTicket, setViewingTicket] = useState(null);
    const [printPaperSize, setPrintPaperSize] = useState("A5");
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);
    const [editingId, setEditingId] = useState(null); // id phiếu đang sửa thông tin, null = đang tạo mới
    const emptyForm = () => ({
        customerName: "", customerPhone: "", customerAddress: "", requestType: HELPDESK_TYPES[0], description: "",
        assignee: currentUser.fullName || "", receivedDate: todayISO(), completedDate: "", solution: "", note: "",
    });
    const [form, setForm] = useState(emptyForm());

    const openEditInfo = (ticket) => {
        setForm({ ...ticket });
        setEditingId(ticket.id);
        setView("new");
    };
    const submitTicket = () => {
        if (!form.customerName.trim()) { alert("Vui lòng nhập tên khách hàng / đơn vị yêu cầu."); return; }
        if (!form.description.trim()) { alert("Vui lòng mô tả yêu cầu hỗ trợ."); return; }
        if (editingId) {
            setHelpdeskTickets((prev) => prev.map((t) => (t.id === editingId ? normalizeHelpdeskTicket({ ...t, ...form }) : t)));
            addLog("Sửa thông tin phiếu IT Helpdesk", `${form.customerName} · ${form.requestType}`);
            setViewingTicket((v) => (v && v.id === editingId ? normalizeHelpdeskTicket({ ...v, ...form }) : v));
        } else {
            const code = nextHelpdeskCode(helpdeskTickets);
            const ticket = normalizeHelpdeskTicket({ id: uid(), code, createdAt: new Date().toISOString(), createdBy: currentUser.fullName, status: "new", ...form });
            setHelpdeskTickets((prev) => [ticket, ...prev]);
            addLog("Tạo phiếu IT Helpdesk", `${code} · ${form.customerName} · ${form.requestType}`);
        }
        setForm(emptyForm());
        setEditingId(null);
        setView("history");
    };
    const setStatus = (ticket, status) => {
        setHelpdeskTickets((prev) => prev.map((t) => (t.id === ticket.id ? { ...t, status } : t)));
        addLog("Cập nhật phiếu IT Helpdesk", `${ticket.code} → ${HELPDESK_STATUSES.find((s) => s.id === status)?.label}`);
        setViewingTicket((v) => (v && v.id === ticket.id ? { ...v, status } : v));
    };
    const doPrint = (ticket) => {
        const html = buildHelpdeskTicketHTML(ticket, printPaperSize);
        const result = printHTML(html);
        setPrintBlockedUrl(result.ok ? null : result.url);
    };

    const filtered = helpdeskTickets.filter((t) => statusFilter === "all" || t.status === statusFilter);

    if (view === "new") {
        return (
            <div>
                <button onClick={() => { setForm(emptyForm()); setEditingId(null); setView("history"); }} className="text-sm mb-4 opacity-60 hover:opacity-100">← Quay lại danh sách phiếu IT Helpdesk</button>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Tên khách hàng / đơn vị"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} /></Field>
                    <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerPhone} onChange={(e) => setForm({ ...form, customerPhone: e.target.value })} /></Field>
                </div>
                <Field label="Địa chỉ"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerAddress} onChange={(e) => setForm({ ...form, customerAddress: e.target.value })} /></Field>
                <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                <Field label="Loại yêu cầu">
                    <div className="flex gap-2 flex-wrap">
                        {HELPDESK_TYPES.map((tp) => (
                            <button key={tp} type="button" onClick={() => setForm({ ...form, requestType: tp })} className="px-3.5 py-1.5 rounded-sm text-sm border"
                                    style={{ borderColor: form.requestType === tp ? INK : LINE, background: form.requestType === tp ? INK : "transparent", color: form.requestType === tp ? "#fff" : INK }}>
                                {tp}
                            </button>
                        ))}
                    </div>
                </Field>
                <Field label="Mô tả yêu cầu / sự cố">
                    <textarea rows={3} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
                </Field>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                    <Field label="Người xử lý">
                        <select className={inputCls} style={{ borderColor: LINE }} value={form.assignee} onChange={(e) => setForm({ ...form, assignee: e.target.value })}>
                            <option value="">— Chưa gán —</option>
                            {employeeNames.map((n) => <option key={n} value={n}>{n}</option>)}
                        </select>
                    </Field>
                    <Field label="Ngày tiếp nhận"><input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.receivedDate} onChange={(e) => setForm({ ...form, receivedDate: e.target.value })} /></Field>
                </div>
                <Field label="Ghi chú">
                    <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
                </Field>
                <button onClick={submitTicket} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>{editingId ? "Lưu thay đổi" : "Tạo phiếu IT Helpdesk"}</button>
            </div>
        );
    }

    return (
        <div>
            <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                <div className="flex gap-1.5 flex-wrap">
                    <FilterChip active={statusFilter === "all"} onClick={() => setStatusFilter("all")}>Tất cả</FilterChip>
                    {HELPDESK_STATUSES.map((s) => <FilterChip key={s.id} active={statusFilter === s.id} onClick={() => setStatusFilter(s.id)} color={s.color}>{s.label}</FilterChip>)}
                </div>
                <button onClick={() => { setForm(emptyForm()); setView("new"); }} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}>
                    <Plus size={15} /> Tạo phiếu IT Helpdesk
                </button>
            </div>
            {filtered.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-16">Chưa có phiếu IT Helpdesk nào.</p>
            ) : (
                <div className="space-y-2.5">
                    {filtered.map((t) => {
                        const st = HELPDESK_STATUSES.find((s) => s.id === t.status);
                        return (
                            <button key={t.id} onClick={() => setViewingTicket(t)} className="w-full text-left p-4 rounded-sm hover:bg-black/[0.02]" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <div className="flex items-center justify-between gap-3 flex-wrap">
                                    <div>
                                        <p className="font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{t.code}</p>
                                        <p className="text-xs opacity-50">{t.customerName} · {t.requestType}{t.assignee ? ` · Người xử lý: ${t.assignee}` : ""} · {formatDateTime(t.createdAt)}</p>
                                    </div>
                                    <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${st.color}1A`, color: st.color }}>{st.label}</span>
                                </div>
                            </button>
                        );
                    })}
                </div>
            )}

            {viewingTicket && (
                <Modal title={`Phiếu IT Helpdesk ${viewingTicket.code}`} onClose={() => { setViewingTicket(null); setPrintBlockedUrl(null); }} size="lg">
                    <div className="flex items-center gap-2 mb-4 flex-wrap">
                        {HELPDESK_STATUSES.map((s) => (
                            <button key={s.id} onClick={() => setStatus(viewingTicket, s.id)} className="text-xs px-3 py-1.5 rounded-full border"
                                    style={{ borderColor: viewingTicket.status === s.id ? s.color : LINE, background: viewingTicket.status === s.id ? `${s.color}1A` : "transparent", color: viewingTicket.status === s.id ? s.color : INK }}>
                                {s.label}
                            </button>
                        ))}
                        <button onClick={() => { setViewingTicket(null); openEditInfo(viewingTicket); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border ml-auto" style={{ borderColor: BLUE, color: BLUE }}>
                            <Pencil size={13} /> Sửa thông tin
                        </button>
                        <select value={printPaperSize} onChange={(e) => setPrintPaperSize(e.target.value)} className="border rounded-sm py-1.5 px-2 text-xs" style={{ borderColor: LINE }}>
                            <option value="A5">Khổ A5</option>
                            <option value="A4">Khổ A4</option>
                        </select>
                        <button onClick={() => doPrint(viewingTicket)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                            <Printer size={13} /> In phiếu
                        </button>
                    </div>

                    {printBlockedUrl && (
                        <div className="mb-4 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                            <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới để mở phiếu, sau đó tự bấm in (Ctrl+P / biểu tượng in) trong tab mới đó:</p>
                            <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở phiếu IT Helpdesk để in</a>
                        </div>
                    )}

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4 text-sm">
                        <div>
                            <p><span className="opacity-50">Khách hàng/Đơn vị: </span>{viewingTicket.customerName}</p>
                            <p><span className="opacity-50">SĐT: </span>{viewingTicket.customerPhone || "—"}</p>
                            <p><span className="opacity-50">Địa chỉ: </span>{viewingTicket.customerAddress || "—"}</p>
                        </div>
                        <div>
                            <p><span className="opacity-50">Người lập: </span>{viewingTicket.createdBy}</p>
                            <p><span className="opacity-50">Người xử lý: </span>{viewingTicket.assignee || "—"}</p>
                            <p><span className="opacity-50">Ngày tiếp nhận: </span>{viewingTicket.receivedDate ? new Date(viewingTicket.receivedDate).toLocaleDateString("vi-VN") : "—"}</p>
                        </div>
                    </div>
                    <div className="p-3 rounded-sm mb-4 text-sm" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                        <p><span className="opacity-50">Loại yêu cầu: </span><b>{viewingTicket.requestType}</b></p>
                        <p className="mt-1"><span className="opacity-50">Mô tả: </span>{viewingTicket.description || "—"}</p>
                    </div>
                    {viewingTicket.status === "done" && (
                        <Field label="Giải pháp / kết quả xử lý">
              <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={viewingTicket.solution}
                        onChange={(e) => { const v = e.target.value; setHelpdeskTickets((prev) => prev.map((t) => (t.id === viewingTicket.id ? { ...t, solution: v } : t))); setViewingTicket({ ...viewingTicket, solution: v }); }} />
                        </Field>
                    )}
                    {viewingTicket.note && <p className="text-sm opacity-70"><b>Ghi chú:</b> {viewingTicket.note}</p>}
                </Modal>
            )}
        </div>
    );
}

// Phiếu dịch vụ — bao gồm 2 mảng tách riêng hoàn toàn: Sửa chữa và IT Helpdesk, chọn mục nào làm việc trên mục đó.
function ServiceTickets({ repairTickets, setRepairTickets, helpdeskTickets, setHelpdeskTickets, employeeNames, currentUser, addLog }) {
    const [serviceSub, setServiceSub] = useState("repair"); // repair | helpdesk
    return (
        <div>
            <div className="grid grid-cols-2 gap-2 mb-5" style={{ maxWidth: 420 }}>
                <button onClick={() => setServiceSub("repair")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: serviceSub === "repair" ? INK : LINE, background: serviceSub === "repair" ? INK : "transparent", color: serviceSub === "repair" ? "#fff" : INK }}>
                    Phiếu dịch vụ sửa chữa
                </button>
                <button onClick={() => setServiceSub("helpdesk")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: serviceSub === "helpdesk" ? INK : LINE, background: serviceSub === "helpdesk" ? INK : "transparent", color: serviceSub === "helpdesk" ? "#fff" : INK }}>
                    Phiếu dịch vụ IT Helpdesk
                </button>
            </div>
            {serviceSub === "repair" && <RepairTickets repairTickets={repairTickets} setRepairTickets={setRepairTickets} currentUser={currentUser} addLog={addLog} />}
            {serviceSub === "helpdesk" && <HelpdeskTickets helpdeskTickets={helpdeskTickets} setHelpdeskTickets={setHelpdeskTickets} employeeNames={employeeNames} currentUser={currentUser} addLog={addLog} />}
        </div>
    );
}

function ProductsSection({ products, setProducts, purchaseOrders, setPurchaseOrders, suppliers, setSuppliers, categories, setCategories, brands, setBrands, stocktakes, setStocktakes, warrantyTickets, setWarrantyTickets, repairTickets, setRepairTickets, helpdeskTickets, setHelpdeskTickets, orders, customers, employeeNames, currentUser, addLog, navTarget, onFocusHandled, goToDoc, goToSupplier }) {
    const [sub, setSub] = useState("list");
    const isAdmin = currentUser.role === "admin";
    const isCtv = currentUser.role === "ctv";

    useEffect(() => {
        if (navTarget?.type === "po" && isAdmin) setSub("purchase");
        else if (navTarget?.type === "product") setSub("list");
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [navTarget]);

    const tabCount = 1 + (isAdmin ? 2 : 0) + (!isCtv ? 2 : 0);

    return (
        <div>
            <div className={`grid gap-2 mb-5`} style={{ gridTemplateColumns: `repeat(${tabCount}, minmax(0, 1fr))`, maxWidth: tabCount === 1 ? 220 : tabCount * 170 }}>
                <button onClick={() => setSub("list")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: sub === "list" ? INK : LINE, background: sub === "list" ? INK : "transparent", color: sub === "list" ? "#fff" : INK }}>
                    Danh sách sản phẩm
                </button>
                {isAdmin && (
                    <button onClick={() => setSub("purchase")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: sub === "purchase" ? INK : LINE, background: sub === "purchase" ? INK : "transparent", color: sub === "purchase" ? "#fff" : INK }}>
                        Nhập hàng
                    </button>
                )}
                {isAdmin && (
                    <button onClick={() => setSub("stocktake")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: sub === "stocktake" ? INK : LINE, background: sub === "stocktake" ? INK : "transparent", color: sub === "stocktake" ? "#fff" : INK }}>
                        Kiểm kho
                    </button>
                )}
                {!isCtv && (
                    <button onClick={() => setSub("warranty")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: sub === "warranty" ? INK : LINE, background: sub === "warranty" ? INK : "transparent", color: sub === "warranty" ? "#fff" : INK }}>
                        Phiếu bảo hành
                    </button>
                )}
                {!isCtv && (
                    <button onClick={() => setSub("service")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: sub === "service" ? INK : LINE, background: sub === "service" ? INK : "transparent", color: sub === "service" ? "#fff" : INK }}>
                        Phiếu dịch vụ
                    </button>
                )}
            </div>
            {sub === "list" && <ProductsInventory products={products} setProducts={setProducts} addLog={addLog} currentUser={currentUser} focusProductId={navTarget?.type === "product" ? navTarget.id : null} onFocusHandled={onFocusHandled} goToDoc={goToDoc} suppliers={suppliers} goToSupplier={goToSupplier} categories={categories} setCategories={setCategories} brands={brands} setBrands={setBrands} />}
            {isAdmin && sub === "purchase" && <PurchaseOrders purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} products={products} setProducts={setProducts} suppliers={suppliers} setSuppliers={setSuppliers} employeeNames={employeeNames} addLog={addLog} focusPOId={navTarget?.type === "po" ? navTarget.id : null} onFocusHandled={onFocusHandled} />}
            {isAdmin && sub === "stocktake" && <Stocktake products={products} setProducts={setProducts} stocktakes={stocktakes} setStocktakes={setStocktakes} currentUser={currentUser} addLog={addLog} />}
            {!isCtv && sub === "warranty" && <WarrantyTickets products={products} setProducts={setProducts} orders={orders} customers={customers} warrantyTickets={warrantyTickets} setWarrantyTickets={setWarrantyTickets} currentUser={currentUser} addLog={addLog} goToDoc={goToDoc} />}
            {!isCtv && sub === "service" && <ServiceTickets repairTickets={repairTickets} setRepairTickets={setRepairTickets} helpdeskTickets={helpdeskTickets} setHelpdeskTickets={setHelpdeskTickets} employeeNames={employeeNames} currentUser={currentUser} addLog={addLog} />}
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
                    <div className="p-2 sticky top-0 flex items-center gap-2" style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
                        <div className="relative flex-1">
                            <Search size={13} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-40" />
                            <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm kiếm tỉnh/thành…"
                                   className="w-full pl-7 pr-2 py-1.5 text-sm rounded-sm border outline-none" style={{ borderColor: LINE }} />
                        </div>
                        <button type="button" onClick={() => { setOpen(false); setQuery(""); }} title="Đóng" className="shrink-0 opacity-60 hover:opacity-100 p-1" style={{ color: INK }}><X size={16} /></button>
                    </div>
                    {matches.length === 0 ? <p className="text-sm opacity-50 p-3">Không tìm thấy.</p> : matches.map((p) => (
                        <button key={p} onClick={() => { onChange(p); setQuery(""); setOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5" style={{ borderBottom: `1px dashed ${LINE}`, color: INK }}>{p}</button>
                    ))}
                </div>
            )}
        </div>
    );
}

// Chọn Phường/Xã — danh sách phụ thuộc vào Tỉnh/Thành đã chọn (theo địa giới mới sau sáp nhập 1/7/2025).
function WardSelect({ province, value, onChange }) {
    const [query, setQuery] = useState("");
    const [open, setOpen] = useState(false);
    const wards = WARDS_BY_PROVINCE[province] || [];
    const q = query.trim().toLowerCase();
    const matches = q ? wards.filter((w) => w.toLowerCase().includes(q)) : wards;

    if (!province) {
        return <p className="text-sm py-1.5 px-1 opacity-40 border-b-2" style={{ borderColor: LINE }}>Chọn Tỉnh/Thành phố trước</p>;
    }

    return (
        <div className="relative">
            <button type="button" onClick={() => setOpen((o) => !o)} className="w-full text-left border-b-2 outline-none py-1.5 px-1 text-[15px]" style={{ borderColor: LINE, color: value ? INK : "#999" }}>
                {value || "Chọn Phường/Xã"}
            </button>
            {open && (
                <div className="absolute z-20 mt-1 w-full max-h-64 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <div className="p-2 sticky top-0 flex items-center gap-2" style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
                        <div className="relative flex-1">
                            <Search size={13} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-40" />
                            <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm kiếm phường/xã…"
                                   className="w-full pl-7 pr-2 py-1.5 text-sm rounded-sm border outline-none" style={{ borderColor: LINE }} />
                        </div>
                        <button type="button" onClick={() => { setOpen(false); setQuery(""); }} title="Đóng" className="shrink-0 opacity-60 hover:opacity-100 p-1" style={{ color: INK }}><X size={16} /></button>
                    </div>
                    {matches.length === 0 ? <p className="text-sm opacity-50 p-3">Không tìm thấy.</p> : matches.map((w) => (
                        <button key={w} onClick={() => { onChange(w); setQuery(""); setOpen(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-black/5" style={{ borderBottom: `1px dashed ${LINE}`, color: INK }}>{w}</button>
                    ))}
                </div>
            )}
        </div>
    );
}

function Customers({ customers, setCustomers, orders, products, currentUser, addLog, goToDoc, employeeNames }) {
    const isAdmin = currentUser.role === "admin";
    const [view, setView] = useState("list"); // list | debt
    const [editing, setEditing] = useState(null);
    const [form, setForm] = useState({});
    const [viewingId, setViewingId] = useState(null);
    const [detailTab, setDetailTab] = useState("orders"); // orders | addresses
    const [addrForm, setAddrForm] = useState(null); // null = đóng; {} = thêm mới; {...addr} = sửa
    const [query, setQuery] = useState("");
    const [debtSort, setDebtSort] = useState("amount"); // amount | due

    const openNew = () => { setForm({ code: nextCustomerCode(customers), name: "", phone: "", contactPerson: "", email: "", taxCode: "", province: "", ward: "", addressDetail: "", note: "", group: "retail", representativeName: "", representativeTitle: "", assignedTo: isAdmin ? "" : currentUser.fullName }); setEditing({}); };
    const openEdit = (c) => { setForm({ ...c }); setEditing(c); };
    const submit = () => {
        if (!form.name) return;
        const assignedTo = isAdmin ? (form.assignedTo || "") : (editing.id ? (form.assignedTo || currentUser.fullName) : currentUser.fullName);
        if (editing.id) {
            setCustomers((prev) => prev.map((c) => (c.id === editing.id ? { ...c, ...form, assignedTo, code: form.code || c.code } : c)));
            addLog("Sửa khách hàng", form.name);
        } else {
            setCustomers((prev) => [...prev, { ...form, assignedTo, id: uid(), code: form.code || nextCustomerCode(customers) }]);
            addLog("Thêm khách hàng", `${form.name}${assignedTo ? ` · Phụ trách: ${assignedTo}` : ""}`);
        }
        setEditing(null);
    };
    const remove = (id) => setCustomers((prev) => prev.filter((c) => c.id !== id));
    const fullAddress = (c) => [c.addressDetail, c.ward, c.province].filter(Boolean).join(", ");
    const fullAddressOf = (a) => [a.addressDetail, a.ward, a.province].filter(Boolean).join(", ");

    // Sổ địa chỉ khách hàng — địa chỉ #1 luôn là địa chỉ gốc của khách hàng (customerAddressBook lo phần này tự động).
    // Các địa chỉ thêm ở đây (giao hàng, hoá đơn, chi nhánh…) xếp từ #2 trở đi, mỗi địa chỉ kèm người nhận/SĐT riêng.
    const openNewAddress = (customer) => setAddrForm({ customerId: customer.id, label: "Địa chỉ giao hàng", recipientName: customer.name, recipientPhone: customer.phone, province: "", ward: "", addressDetail: "", isDefault: false });
    const openEditAddress = (customer, addr) => setAddrForm({ customerId: customer.id, ...addr });
    const submitAddress = () => {
        if (!addrForm.addressDetail && !addrForm.province) { alert("Vui lòng nhập ít nhất Tỉnh/Thành hoặc địa chỉ cụ thể."); return; }
        setCustomers((prev) => prev.map((c) => {
            if (c.id !== addrForm.customerId) return c;
            const list = Array.isArray(c.addresses) ? c.addresses : [];
            let addresses;
            if (addrForm.id) {
                addresses = list.map((a) => (a.id === addrForm.id ? { ...a, ...addrForm } : a));
            } else {
                addresses = [...list, { ...addrForm, id: uid() }];
            }
            if (addrForm.isDefault) addresses = addresses.map((a) => ({ ...a, isDefault: a.id === (addrForm.id || addresses[addresses.length - 1].id) }));
            return { ...c, addresses };
        }));
        addLog(addrForm.id ? "Sửa địa chỉ khách hàng" : "Thêm địa chỉ khách hàng", customers.find((c) => c.id === addrForm.customerId)?.name || "");
        setAddrForm(null);
    };
    const removeAddress = (customer, addrId) => {
        setCustomers((prev) => prev.map((c) => (c.id !== customer.id ? c : { ...c, addresses: (c.addresses || []).filter((a) => a.id !== addrId) })));
    };
    // addrId === "primary" nghĩa là chọn lại địa chỉ gốc của khách hàng làm mặc định — chỉ cần bỏ cờ mặc định ở mọi địa chỉ thêm trong sổ.
    const setDefaultAddress = (customer, addrId) => {
        setCustomers((prev) => prev.map((c) => (c.id !== customer.id ? c : { ...c, addresses: (c.addresses || []).map((a) => ({ ...a, isDefault: addrId !== "primary" && a.id === addrId })) })));
    };

    const customerStats = (id) => {
        const custOrders = orders.filter((o) => o.customerId === id);
        const active = custOrders.filter((o) => o.status !== "cancelled");
        const totalSpent = active.reduce((s, o) => s + orderCalc(o).payable, 0);
        const debt = active.reduce((s, o) => s + Math.max(0, orderCalc(o).remaining), 0);
        const lastOrderAt = custOrders.reduce((max, o) => (!max || o.createdAt > max ? o.createdAt : max), null);
        return { orderCount: custOrders.length, totalSpent, debt, custOrders, lastOrderAt };
    };

    // Phạm vi khách hàng theo quyền: admin thấy tất cả; nhân viên/CTV chỉ thấy khách hàng do mình phụ trách.
    const myCustomers = isAdmin ? customers : customers.filter((c) => c.assignedTo === currentUser.fullName);

    const filtered = myCustomers.filter((c) =>
        c.name.toLowerCase().includes(query.toLowerCase()) || c.code.toLowerCase().includes(query.toLowerCase()) || (c.phone || "").includes(query)
    );

    const viewingCustomer = myCustomers.find((c) => c.id === viewingId) || null;

    // Xuất Excel danh sách khách hàng (theo đúng bộ lọc/tìm kiếm đang áp dụng) kèm số liệu tổng hợp.
    const exportCustomerList = () => {
        if (filtered.length === 0) { alert("Không có khách hàng nào để xuất."); return; }
        const rows = filtered.map((c) => {
            const s = customerStats(c.id);
            const g = CUSTOMER_GROUPS.find((x) => x.id === c.group) || CUSTOMER_GROUPS[0];
            return {
                "Mã KH": c.code, "Tên khách hàng": c.name, "Người liên hệ": c.contactPerson || "", "SĐT": c.phone || "", "Email": c.email || "", "MST": c.taxCode || "",
                "Nhóm khách hàng": g.label, "Phụ trách": c.assignedTo || "", "Địa chỉ": fullAddress(c),
                "Đại diện (Ông/Bà)": c.representativeName || "", "Chức vụ": c.representativeTitle || "",
                "Công nợ": s.debt, "Tổng chi tiêu": s.totalSpent, "Tổng SL đơn": s.orderCount, "Ghi chú": c.note || "",
            };
        });
        exportExcel(`KhachHang_${todayISO()}`, [{ name: "Khách hàng", rows }]);
        addLog("Xuất Excel khách hàng", `${filtered.length} khách hàng`);
    };

    // Nhập khách hàng hàng loạt từ file Excel — dùng đúng định dạng file "Xuất Excel": tải file xuất ra, thêm dòng mới rồi tải lên lại.
    // Mã KH đã có thì bỏ qua (không ghi đè), mã KH mới hoặc để trống thì thêm vào.
    const importFileRef = useRef(null);
    const [importResult, setImportResult] = useState(null); // { added, skipped: [] }

    const parseGroupCell = (v) => {
        const s = String(v ?? "").trim().toLowerCase();
        if (s.includes("doanh nghiệp") || s.includes("doanh nghiep")) return "enterprise";
        if (s.includes("b2b")) return "b2b";
        return "retail";
    };
    const triggerImportFile = () => importFileRef.current && importFileRef.current.click();
    const handleImportFile = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: "array" });
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                const existingCodes = new Set(customers.filter((c) => c.code).map((c) => c.code.trim().toLowerCase()));
                const newCustomers = [];
                const skipped = [];
                rows.forEach((row) => {
                    const name = String(row["Tên khách hàng"] ?? "").trim();
                    if (!name) return; // bỏ qua dòng thiếu tên khách hàng
                    let code = String(row["Mã KH"] ?? "").trim();
                    if (code) {
                        const codeKey = code.toLowerCase();
                        if (existingCodes.has(codeKey)) { skipped.push(code); return; }
                        existingCodes.add(codeKey);
                    } else {
                        code = nextCustomerCode([...customers, ...newCustomers]);
                        existingCodes.add(code.toLowerCase());
                    }
                    const rawAssignee = String(row["Phụ trách"] ?? "").trim();
                    const assignedTo = isAdmin ? (employeeNames.includes(rawAssignee) ? rawAssignee : "") : currentUser.fullName;
                    // Cột "Địa chỉ" trong file Xuất Excel là địa chỉ gộp đầy đủ — đưa nguyên vào Địa chỉ cụ thể, có thể tách lại Tỉnh/Phường sau khi sửa hồ sơ.
                    newCustomers.push(normalizeCustomer({
                        id: uid(), code, name, contactPerson: String(row["Người liên hệ"] ?? "").trim(), phone: String(row["SĐT"] ?? "").trim(),
                        email: String(row["Email"] ?? "").trim(), taxCode: String(row["MST"] ?? "").trim(), group: parseGroupCell(row["Nhóm khách hàng"]),
                        addressDetail: String(row["Địa chỉ"] ?? "").trim(),
                        representativeName: String(row["Đại diện (Ông/Bà)"] ?? "").trim(), representativeTitle: String(row["Chức vụ"] ?? "").trim(),
                        assignedTo, note: String(row["Ghi chú"] ?? "").trim(),
                    }));
                });
                if (newCustomers.length > 0) setCustomers((prev) => [...prev, ...newCustomers]);
                addLog("Nhập khách hàng từ Excel", `${newCustomers.length} khách hàng mới${skipped.length > 0 ? ` · Bỏ qua ${skipped.length} mã đã tồn tại` : ""}`);
                setImportResult({ added: newCustomers.length, skipped });
            } catch (err) {
                alert("Không đọc được file Excel — vui lòng dùng đúng file đã tải từ nút \"Xuất Excel\" (chỉ thêm dòng mới, không đổi tên cột) rồi thử lại.");
            }
            if (importFileRef.current) importFileRef.current.value = "";
        };
        reader.readAsArrayBuffer(file);
    };

    // Xuất Excel lịch sử mua hàng chi tiết của 1 khách cụ thể: từng đơn + từng sản phẩm trong đơn.
    const exportCustomerHistory = (cust) => {
        const s = customerStats(cust.id);
        if (s.custOrders.length === 0) { alert("Khách hàng này chưa có đơn hàng nào."); return; }
        const orderRows = s.custOrders.map((o) => {
            const c = orderCalc(o);
            const returns = o.returns || [];
            return {
                "Mã đơn": o.code, "Ngày tạo": formatDateTime(o.createdAt), "Trạng thái giao": STATUSES.find((st) => st.id === o.status)?.label || o.status,
                "Xuất hoá đơn": o.status === "cancelled" ? "Huỷ" : o.invoiceStatus === "issued" ? "Đã xuất" : "Chờ xuất", "Số hoá đơn": o.invoiceNo || "",
                "Tổng tiền": c.payable, "Đã trả": o.paidAmount, "Còn phải trả": c.remaining,
                "Số lần đổi trả": returns.length, "Đang chờ duyệt đổi trả": o.returnRequest ? "Có" : "",
                "Lý do huỷ": o.status === "cancelled" ? (o.cancelReason || "") : "", "Ghi chú": o.notes || "",
            };
        });
        const detailRows = [];
        s.custOrders.forEach((o) => {
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                detailRows.push({
                    "Mã đơn": o.code, "Ngày": formatDateTime(o.createdAt), "Sản phẩm": p?.name || "?", "Mã VT": p?.code || "",
                    "SL": it.qty, "Đơn giá": it.price, "Thành tiền": orderLineTotal(it), "Series": (it.series || []).join(", "),
                });
            });
        });
        // Đổi trả hàng — liệt kê đầy đủ từng phiếu đổi/trả của mọi đơn thuộc khách hàng này để dễ theo dõi.
        const returnRows = [];
        s.custOrders.forEach((o) => {
            (o.returns || []).forEach((r) => {
                const returnedNames = (r.returnedItems || []).map((it) => {
                    const p = products.find((x) => x.id === it.productId);
                    return `${p?.name || "?"} x${it.qty}`;
                }).join("; ");
                const exchangeNames = (r.exchangeItems || []).map((it) => {
                    const p = products.find((x) => x.id === it.productId);
                    return `${p?.name || "?"} x${it.qty}`;
                }).join("; ");
                const returnedValue = (r.returnedItems || []).reduce((s2, it) => s2 + it.qty * it.price, 0);
                const exchangeValue = (r.exchangeItems || []).reduce((s2, it) => s2 + it.qty * it.price, 0);
                returnRows.push({
                    "Mã đơn gốc": o.code, "Mã phiếu đổi trả": r.code, "Ngày": formatDateTime(r.createdAt),
                    "Loại": r.type === "exchange" ? "Đổi hàng" : "Hoàn tiền/Trả hàng",
                    "Sản phẩm trả": returnedNames, "Giá trị trả": returnedValue,
                    "Sản phẩm đổi": exchangeNames, "Giá trị đổi": r.type === "exchange" ? exchangeValue : "",
                    "Ghi chú": r.note || "",
                });
            });
        });
        exportExcel(`LichSuMuaHang_${cust.code}_${todayISO()}`, [
            { name: "Đơn hàng", rows: orderRows },
            { name: "Chi tiết sản phẩm", rows: detailRows },
            { name: "Đổi trả hàng", rows: returnRows },
        ]);
        addLog("Xuất Excel lịch sử mua hàng", `${cust.name} · ${s.custOrders.length} đơn`);
    };

    // Công nợ phải thu từng khách hàng: chỉ tính khách có công nợ > 0, kèm hạn công nợ gần nhất (nếu đơn có khai báo số ngày công nợ).
    const debtList = myCustomers.map((c) => {
        const s = customerStats(c.id);
        const unpaidOrders = s.custOrders.filter((o) => o.status !== "cancelled" && orderCalc(o).remaining > 0);
        const dues = unpaidOrders.map((o) => orderDueInfo(o)).filter(Boolean);
        const nearestDue = dues.length ? dues.reduce((a, b) => (a.daysLeft < b.daysLeft ? a : b)) : null;
        return { customer: c, debt: s.debt, orderCount: unpaidOrders.length, nearestDue };
    }).filter((d) => d.debt > 0);
    const grandDebtTotal = debtList.reduce((s, d) => s + d.debt, 0);
    const sortedDebt = [...debtList].sort((a, b) => {
        if (debtSort === "due") {
            if (!a.nearestDue) return 1;
            if (!b.nearestDue) return -1;
            return a.nearestDue.daysLeft - b.nearestDue.daysLeft;
        }
        return b.debt - a.debt;
    });

    return (
        <div>
            <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
                <div className="grid grid-cols-2 max-w-xs gap-2">
                    <button onClick={() => setView("list")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                            style={{ borderColor: view === "list" ? INK : LINE, background: view === "list" ? INK : "transparent", color: view === "list" ? "#fff" : INK }}>
                        Danh sách khách hàng
                    </button>
                    <button onClick={() => setView("debt")} className="px-3.5 py-2.5 rounded-full text-sm border text-center relative"
                            style={{ borderColor: view === "debt" ? INK : LINE, background: view === "debt" ? INK : "transparent", color: view === "debt" ? "#fff" : INK }}>
                        Công nợ khách hàng
                        {debtList.length > 0 && <span className="ml-1.5 text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: view === "debt" ? "rgba(255,255,255,0.25)" : `${RUST}1A`, color: view === "debt" ? "#fff" : RUST }}>{debtList.length}</span>}
                    </button>
                </div>
                {view === "list" && (
                    <div className="flex gap-2 flex-wrap">
                        <div className="relative flex-1 max-w-xs">
                            <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
                            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã, tên, SĐT…"
                                   className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
                        </div>
                        <button onClick={exportCustomerList} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                            <FileSpreadsheet size={15} /> Xuất Excel
                        </button>
                        <button onClick={triggerImportFile} title="Dùng file đã tải từ nút Xuất Excel — thêm dòng khách hàng mới rồi tải lên lại" className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border whitespace-nowrap" style={{ borderColor: BLUE, color: BLUE }}>
                            <ArrowUpFromLine size={15} /> Nhập từ Excel
                        </button>
                        <input ref={importFileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImportFile} />
                        <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}><Plus size={15} /> Thêm khách hàng</button>
                    </div>
                )}
            </div>

            {view === "debt" ? (
                <div>
                    <div className="p-5 rounded-sm mb-5 flex items-center justify-between flex-wrap gap-3" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                        <div>
                            <p className="text-xs uppercase tracking-wider opacity-55 mb-1">Tổng phải thu khách hàng</p>
                            <p className="text-2xl" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(grandDebtTotal)}</p>
                        </div>
                        <div className="flex gap-1.5">
                            <button onClick={() => setDebtSort("amount")} className="text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: debtSort === "amount" ? INK : LINE, background: debtSort === "amount" ? INK : "transparent", color: debtSort === "amount" ? "#fff" : INK }}>Số tiền</button>
                            <button onClick={() => setDebtSort("due")} className="text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: debtSort === "due" ? INK : LINE, background: debtSort === "due" ? INK : "transparent", color: debtSort === "due" ? "#fff" : INK }}>Hạn công nợ</button>
                        </div>
                    </div>
                    {sortedDebt.length === 0 ? (
                        <p className="text-sm opacity-50 text-center py-16">Không có khách hàng nào đang nợ.</p>
                    ) : (
                        <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                            <table className="w-full text-sm" style={{ minWidth: 700 }}>
                                <thead>
                                <tr style={{ borderBottom: `2px solid ${INK}` }}>
                                    {["Khách hàng", "Số đơn nợ", "Tổng nợ", "Hạn gần nhất"].map((h, i) => (
                                        <th key={i} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
                                    ))}
                                </tr>
                                </thead>
                                <tbody>
                                {sortedDebt.map((d) => (
                                    <tr key={d.customer.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02] cursor-pointer" onClick={() => setViewingId(d.customer.id)}>
                                        <td className="px-3 py-3">
                                            <p style={{ color: INK }}>{d.customer.name}</p>
                                            <p className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.customer.code}</p>
                                        </td>
                                        <td className="px-3 py-3" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.orderCount}</td>
                                        <td className="px-3 py-3 font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(d.debt)}</td>
                                        <td className="px-3 py-3 whitespace-nowrap">
                                            {d.nearestDue ? (
                                                <span className="text-xs px-2 py-0.5 rounded-full" style={{ background: d.nearestDue.overdue3 ? `${RUST}1A` : PAPER, color: d.nearestDue.overdue3 ? RUST : INK }}>
                            {d.nearestDue.overdue ? `Quá hạn ${-d.nearestDue.daysLeft} ngày` : d.nearestDue.daysLeft === 0 ? "Đến hạn hôm nay" : `Còn ${d.nearestDue.daysLeft} ngày`}
                          </span>
                                            ) : <span className="opacity-40">Không có hạn công nợ</span>}
                                        </td>
                                    </tr>
                                ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            ) : (
                <>

                    <div className="rounded-sm overflow-x-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                        <table className="w-full text-sm" style={{ minWidth: 760 }}>
                            <thead>
                            <tr style={{ borderBottom: `2px solid ${INK}` }}>
                                {(isAdmin ? ["Mã KH", "Tên khách hàng", "SĐT", "Nhóm khách hàng", "Phụ trách", "Công nợ", "Tổng chi tiêu", "Tổng SL đơn", ""] : ["Mã KH", "Tên khách hàng", "SĐT", "Nhóm khách hàng", "Công nợ", "Tổng chi tiêu", "Tổng SL đơn", ""]).map((h, hi) => (
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
                                        {isAdmin && (
                                            <td className="px-3 py-3 whitespace-nowrap text-sm opacity-70">{c.assignedTo || <span className="opacity-40">—</span>}</td>
                                        )}
                                        <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: s.debt > 0 ? RUST : "inherit", opacity: s.debt > 0 ? 1 : 0.4 }}>{vnd(s.debt)}</td>
                                        <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(s.totalSpent)}</td>
                                        <td className="px-3 py-3 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.orderCount}</td>
                                        <td className="px-3 py-3">
                                            {isAdmin && (
                                                <div className="flex gap-1.5 justify-end whitespace-nowrap">
                                                    <button onClick={() => openEdit(c)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                                                    <button onClick={() => remove(c.id)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                                                </div>
                                            )}
                                        </td>
                                    </tr>
                                );
                            })}
                            {filtered.length === 0 && <tr><td colSpan={isAdmin ? 9 : 8} className="text-center py-10 opacity-50">Không có khách hàng nào.</td></tr>}
                            </tbody>
                        </table>
                    </div>
                </>
            )}

            {/* Modal chi tiết khách hàng */}
            {importResult && (
                <Modal title="Kết quả nhập khách hàng từ Excel" onClose={() => setImportResult(null)}>
                    <div className="p-3 rounded-sm mb-3" style={{ background: `${FOREST}0D`, border: `1px solid ${FOREST}44` }}>
                        <p className="text-sm" style={{ color: FOREST }}>Đã thêm <b>{importResult.added}</b> khách hàng mới.</p>
                    </div>
                    {importResult.skipped.length > 0 && (
                        <div className="p-3 rounded-sm" style={{ background: `${BRASS}0D`, border: `1px solid ${BRASS}44` }}>
                            <p className="text-sm mb-1.5" style={{ color: BRASS }}>Bỏ qua {importResult.skipped.length} mã KH đã tồn tại (không ghi đè):</p>
                            <p className="text-xs opacity-70" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{importResult.skipped.join(", ")}</p>
                        </div>
                    )}
                </Modal>
            )}

            {viewingCustomer && (() => {
                const s = customerStats(viewingCustomer.id);
                const g = CUSTOMER_GROUPS.find((x) => x.id === viewingCustomer.group) || CUSTOMER_GROUPS[0];
                const addresses = customerAddressBook(viewingCustomer);
                return (
                    <Modal title={viewingCustomer.name} onClose={() => { setViewingId(null); setDetailTab("orders"); }} size="xl">
                        <div className="flex items-center gap-2 mb-4 flex-wrap">
                            <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${BLUE}15`, color: BLUE }}>{g.label}</span>
                            <span className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{viewingCustomer.code}</span>
                            <button onClick={() => exportCustomerHistory(viewingCustomer)} className="ml-auto text-xs px-3 py-1.5 rounded-sm border flex items-center gap-1" style={{ borderColor: FOREST, color: FOREST }}><FileSpreadsheet size={12} /> Xuất lịch sử mua hàng</button>
                            {isAdmin && (
                                <button onClick={() => { setViewingId(null); openEdit(viewingCustomer); }} className="text-xs px-3 py-1.5 rounded-sm border flex items-center gap-1" style={{ borderColor: LINE, color: INK }}><Pencil size={12} /> Sửa thông tin</button>
                            )}
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1.5 text-sm mb-4 p-3 rounded-sm" style={{ background: PAPER }}>
                            <p><span className="opacity-50">SĐT: </span>{viewingCustomer.phone || "—"}</p>
                            {viewingCustomer.contactPerson && <p><span className="opacity-50">Người liên hệ: </span>{viewingCustomer.contactPerson}</p>}
                            <p><span className="opacity-50">Email: </span>{viewingCustomer.email || "—"}</p>
                            <p><span className="opacity-50">MST: </span>{viewingCustomer.taxCode || "—"}</p>
                            <p><span className="opacity-50">Địa chỉ: </span>{fullAddress(viewingCustomer) || "—"}</p>
                            <p><span className="opacity-50">Phụ trách: </span>{viewingCustomer.assignedTo || <span className="opacity-40">Chưa gán</span>}</p>
                            {(viewingCustomer.representativeName || viewingCustomer.representativeTitle) && (
                                <p><span className="opacity-50">Đại diện: </span>{viewingCustomer.representativeName || "—"}{viewingCustomer.representativeTitle ? ` · ${viewingCustomer.representativeTitle}` : ""}</p>
                            )}
                            {viewingCustomer.note && <p className="sm:col-span-2"><span className="opacity-50">Ghi chú: </span>{viewingCustomer.note}</p>}
                        </div>

                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-5">
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
                            <div className="p-3 rounded-sm text-center" style={{ border: `1px solid ${LINE}` }}>
                                <p className="text-[10px] uppercase tracking-wider opacity-50 mb-1">Mua gần nhất</p>
                                <p className="text-sm font-semibold" style={{ color: INK }}>{s.lastOrderAt ? new Date(s.lastOrderAt).toLocaleDateString("vi-VN") : "—"}</p>
                            </div>
                        </div>

                        <div className="flex gap-2 mb-4" style={{ borderBottom: `1px solid ${LINE}` }}>
                            <button onClick={() => setDetailTab("orders")} className="px-3 py-2 text-sm -mb-px" style={{ color: detailTab === "orders" ? BLUE : INK, opacity: detailTab === "orders" ? 1 : 0.55, borderBottom: detailTab === "orders" ? `2px solid ${BLUE}` : "2px solid transparent" }}>Lịch sử mua hàng</button>
                            <button onClick={() => setDetailTab("addresses")} className="px-3 py-2 text-sm -mb-px" style={{ color: detailTab === "addresses" ? BLUE : INK, opacity: detailTab === "addresses" ? 1 : 0.55, borderBottom: detailTab === "addresses" ? `2px solid ${BLUE}` : "2px solid transparent" }}>Địa chỉ ({addresses.length})</button>
                        </div>

                        {detailTab === "orders" && (
                            <div className="rounded-sm overflow-x-auto min-w-0" style={{ border: `1px solid ${LINE}` }}>
                                <table className="w-full text-xs" style={{ minWidth: 560 }}>
                                    <thead style={{ background: PAPER }}><tr className="opacity-60">
                                        <th className="text-left py-2 px-2">Mã đơn</th><th className="text-left py-2 px-2">Trạng thái</th><th className="text-left py-2 px-2">Xuất hoá đơn</th><th className="text-right py-2 px-2">Giá trị</th><th className="text-right py-2 px-2">Còn phải trả</th><th className="text-left py-2 px-2">Ngày</th>
                                    </tr></thead>
                                    <tbody>
                                    {s.custOrders.map((o) => {
                                        const oc = orderCalc(o);
                                        return (
                                            <tr key={o.id} style={{ borderTop: `1px dashed ${LINE}` }}>
                                                <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                                                    {goToDoc ? <button onClick={() => goToDoc(o.code)} className="hover:underline" style={{ color: BLUE }}>{o.code}</button> : <span style={{ color: BLUE }}>{o.code}</span>}
                                                </td>
                                                <td className="py-1.5 px-2"><Stamp status={o.status} /></td>
                                                <td className="py-1.5 px-2 whitespace-nowrap">
                                                    {o.status === "cancelled" ? <span className="opacity-40">—</span> : o.invoiceStatus === "issued" ? (
                                                        <span style={{ color: FOREST, fontFamily: "'IBM Plex Mono', monospace" }}>{o.invoiceNo || "Đã xuất"}</span>
                                                    ) : <span style={{ color: RUST, opacity: 0.8 }}>Chờ</span>}
                                                </td>
                                                <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(oc.payable)}</td>
                                                <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: oc.remaining > 0 ? RUST : "inherit" }}>{vnd(oc.remaining)}</td>
                                                <td className="py-1.5 px-2 whitespace-nowrap">{formatDateTime(o.createdAt)}</td>
                                            </tr>
                                        );
                                    })}
                                    {s.custOrders.length === 0 && <tr><td colSpan={6} className="text-center py-6 opacity-40">Chưa có đơn hàng nào.</td></tr>}
                                    </tbody>
                                </table>
                            </div>
                        )}

                        {detailTab === "addresses" && (
                            <div>
                                <div className="flex justify-end mb-3">
                                    <button onClick={() => openNewAddress(viewingCustomer)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm text-white" style={{ background: INK }}>
                                        <Plus size={13} /> Thêm địa chỉ mới
                                    </button>
                                </div>
                                {addresses.length === 0 ? (
                                    <p className="text-sm opacity-50 text-center py-10">Chưa có địa chỉ nào — khách hàng có thể có nhiều địa chỉ (giao hàng, hoá đơn, chi nhánh…).</p>
                                ) : (
                                    <div className="space-y-2.5">
                                        {addresses.map((a) => (
                                            <div key={a.id} className="p-3 rounded-sm flex items-start justify-between gap-3 flex-wrap" style={{ border: `1px solid ${a.isDefault ? BLUE : LINE}`, background: a.isDefault ? `${BLUE}08` : "#fff" }}>
                                                <div className="min-w-0">
                                                    <div className="flex items-center gap-2 flex-wrap mb-1">
                                                        <span className="text-sm font-medium" style={{ color: INK }}>{a.label || "Địa chỉ"}</span>
                                                        {a.isDefault && <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: `${BLUE}1A`, color: BLUE }}>Mặc định</span>}
                                                    </div>
                                                    <p className="text-sm" style={{ color: INK }}>{a.recipientName || viewingCustomer.name}{a.recipientPhone ? ` · ${a.recipientPhone}` : ""}</p>
                                                    <p className="text-sm opacity-70">{fullAddressOf(a) || "—"}</p>
                                                </div>
                                                <div className="flex gap-1.5 shrink-0">
                                                    {!a.isDefault && <button onClick={() => setDefaultAddress(viewingCustomer, a.id)} className="text-xs px-2.5 py-1 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Đặt mặc định</button>}
                                                    <button onClick={() => (a.isPrimary ? (() => { setViewingId(null); openEdit(viewingCustomer); })() : openEditAddress(viewingCustomer, a))} title={a.isPrimary ? "Sửa trong thông tin khách hàng" : "Sửa địa chỉ"} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={13} /></button>
                                                    {!a.isPrimary && <button onClick={() => removeAddress(viewingCustomer, a.id)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={13} /></button>}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        )}
                    </Modal>
                );
            })()}

            {/* Modal thêm/sửa địa chỉ trong sổ địa chỉ khách hàng */}
            {addrForm && (
                <Modal title={addrForm.id ? "Sửa địa chỉ" : "Thêm mới địa chỉ"} onClose={() => setAddrForm(null)}>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Field label="Họ tên người nhận"><input className={inputCls} style={{ borderColor: LINE }} value={addrForm.recipientName} onChange={(e) => setAddrForm({ ...addrForm, recipientName: e.target.value })} /></Field>
                        <Field label="Số điện thoại nhận hàng"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={addrForm.recipientPhone} onChange={(e) => setAddrForm({ ...addrForm, recipientPhone: e.target.value.replace(/\D/g, "") })} /></Field>
                    </div>
                    <Field label="Nhãn địa chỉ" hint="VD: Địa chỉ giao hàng, Địa chỉ hoá đơn, Chi nhánh…">
                        <input className={inputCls} style={{ borderColor: LINE }} value={addrForm.label} onChange={(e) => setAddrForm({ ...addrForm, label: e.target.value })} />
                    </Field>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Field label="Tỉnh/Thành phố">
                            <ProvinceSelect value={addrForm.province} onChange={(v) => setAddrForm({ ...addrForm, province: v, ward: v !== addrForm.province ? "" : addrForm.ward })} />
                        </Field>
                        <Field label="Phường/Xã">
                            <WardSelect province={addrForm.province} value={addrForm.ward} onChange={(v) => setAddrForm({ ...addrForm, ward: v })} />
                        </Field>
                    </div>
                    <Field label="Địa chỉ cụ thể"><input className={inputCls} style={{ borderColor: LINE }} value={addrForm.addressDetail} onChange={(e) => setAddrForm({ ...addrForm, addressDetail: e.target.value })} placeholder="Số nhà, tên đường, thôn/khu vực…" /></Field>
                    <label className="flex items-center gap-2 text-sm mb-3 mt-1" style={{ color: INK }}>
                        <input type="checkbox" checked={!!addrForm.isDefault} onChange={(e) => setAddrForm({ ...addrForm, isDefault: e.target.checked })} />
                        Đặt làm địa chỉ mặc định
                    </label>
                    <button onClick={submitAddress} className="w-full py-2.5 rounded-sm text-white text-sm" style={{ background: INK }}>{addrForm.id ? "Lưu thay đổi" : "Thêm địa chỉ"}</button>
                </Modal>
            )}

            {editing !== null && (
                <Modal title={editing.id ? "Sửa khách hàng" : "Thêm khách hàng"} onClose={() => setEditing(null)} size="xl">
                    <Field label="Tên khách hàng"><input className={inputCls} style={{ borderColor: LINE }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Người liên hệ" hint="Không bắt buộc — dùng khi khách hàng là công ty">
                            <input className={inputCls} style={{ borderColor: LINE }} value={form.contactPerson || ""} onChange={(e) => setForm({ ...form, contactPerson: e.target.value })} />
                        </Field>
                        <Field label="Số điện thoại"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value.replace(/\D/g, "") })} /></Field>
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Mã khách hàng" hint="Tự động sinh — có thể sửa lại">
                            <input className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} />
                        </Field>
                        <Field label="Mã số thuế"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={form.taxCode} onChange={(e) => setForm({ ...form, taxCode: e.target.value.replace(/\D/g, "") })} /></Field>
                    </div>
                    <Field label="Email"><input className={inputCls} style={{ borderColor: LINE }} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>

                    <div className="grid grid-cols-2 gap-4">
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

                        <Field label="Nhân viên phụ trách" hint={isAdmin ? "Chỉ nhân viên/CTV được gán mới thấy công nợ, đơn hàng của khách này" : "Khách hàng do bạn tạo sẽ tự động do bạn phụ trách"}>
                            {isAdmin ? (
                                <select className={inputCls} style={{ borderColor: LINE }} value={form.assignedTo || ""} onChange={(e) => setForm({ ...form, assignedTo: e.target.value })}>
                                    <option value="">— Chưa gán (chỉ QTV thấy) —</option>
                                    {employeeNames.map((n) => <option key={n} value={n}>{n}</option>)}
                                </select>
                            ) : (
                                <p className="text-sm py-1.5 px-1 opacity-70">{form.assignedTo || currentUser.fullName}</p>
                            )}
                        </Field>
                    </div>

                    {(form.group === "b2b" || form.group === "enterprise") && (
                        <div className="grid grid-cols-2 gap-4">
                            <Field label="Đại diện (Ông/Bà)" hint="Dùng khi in biên bản bàn giao hàng hoá">
                                <input className={inputCls} style={{ borderColor: LINE }} value={form.representativeName || ""} onChange={(e) => setForm({ ...form, representativeName: e.target.value })} />
                            </Field>
                            <Field label="Chức vụ">
                                <input className={inputCls} style={{ borderColor: LINE }} value={form.representativeTitle || ""} onChange={(e) => setForm({ ...form, representativeTitle: e.target.value })} />
                            </Field>
                        </div>
                    )}

                    <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />
                    <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Địa chỉ</p>
                    <div className="grid grid-cols-2 gap-4">
                        <Field label="Tỉnh/Thành phố">
                            <ProvinceSelect value={form.province} onChange={(v) => setForm({ ...form, province: v, ward: v !== form.province ? "" : form.ward })} />
                        </Field>
                        <Field label="Phường/Xã" hint="Theo địa danh mới sau sáp nhập 1/7/2025">
                            <WardSelect province={form.province} value={form.ward} onChange={(v) => setForm({ ...form, ward: v })} />
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

function CustomerPicker({ customers, setCustomers, onPick, placeholder, currentUser }) {
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
            assignedTo: currentUser && currentUser.role !== "admin" ? currentUser.fullName : "",
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
                    onFocus={() => setOpen(true)}
                    placeholder={placeholder || "Tìm theo tên, SĐT khách hàng…"}
                    className="w-full pl-8 pr-8 py-2.5 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }}
                />
                {(open || creatingNew) && (
                    <button type="button" onClick={() => { setOpen(false); setCreatingNew(false); }} title="Đóng bảng" className="absolute right-2 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-100" style={{ color: INK }}>
                        <X size={15} />
                    </button>
                )}
            </div>
            {(open || creatingNew) && (
                <div className="absolute z-20 mt-1 w-full max-h-72 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    {creatingNew ? (
                        <div className="p-3">
                            <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Tạo khách hàng mới</p>
                            <input autoFocus value={newForm.name} onChange={(e) => setNewForm({ ...newForm, name: e.target.value })} placeholder="Tên khách hàng"
                                   className="w-full border rounded-sm py-2 px-2.5 text-sm mb-2" style={{ borderColor: LINE }} />
                            <input value={newForm.phone} onChange={(e) => setNewForm({ ...newForm, phone: e.target.value.replace(/\D/g, "") })} placeholder="Số điện thoại" inputMode="numeric"
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
    const pick = (serial) => { if (selected.length >= need) return; setSelected([...selected, serial]); setQuery(""); };
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
                        onFocus={() => setOpen(true)}
                        placeholder={selected.length >= need ? "Đã đủ series" : "Gõ để tìm số series còn tồn…"}
                        className="w-full pl-8 pr-8 py-2 text-xs rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }}
                    />
                    {open && (
                        <button type="button" onClick={() => setOpen(false)} title="Đóng" className="absolute right-2 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-100" style={{ color: INK }}>
                            <X size={13} />
                        </button>
                    )}
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

function SalesItemsTable({ items, products, onUpdate, onRemove, role }) {
    return (
        <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
            <table className="w-full text-sm table-fixed" style={{ minWidth: 560 }}>
                <colgroup>
                    <col style={{ width: 28 }} />
                    <col style={{ width: 46 }} />
                    <col />
                    <col style={{ width: 76 }} />
                    <col style={{ width: 150 }} />
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
                    const floor = p ? minSellPrice(p, role) : 0;
                    const belowFloor = p && floor > 0 && it.price < floor;
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
                                    <input type="number" min={1} value={it.qty || ""} onChange={(e) => { const v = e.target.value; onUpdate(it.productId, { qty: v === "" ? "" : Math.max(0, Number(v)) }); }}
                                           onBlur={() => { if (!it.qty || it.qty < 1) onUpdate(it.productId, { qty: 1 }); }}
                                           className="w-full border rounded-sm py-2 px-1 text-center text-[15px] font-medium" style={{ borderColor: LINE }} />
                                </td>
                                <td className="px-1 py-3 align-top">
                                    <MoneyInput value={it.price} onChange={(v) => onUpdate(it.productId, { price: v === "" ? 0 : v })}
                                                className="w-full border rounded-sm py-2 px-2 text-right text-[15px] font-medium" style={{ borderColor: belowFloor ? RUST : LINE, fontFamily: "'IBM Plex Mono', monospace", color: belowFloor ? RUST : INK }} />
                                    {p && (
                                        <div className="flex gap-1 mt-1">
                                            <button type="button" onClick={() => onUpdate(it.productId, { price: p.wholesalePrice })} className="text-[10px] px-1.5 py-0.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Giá sỉ</button>
                                            <button type="button" onClick={() => onUpdate(it.productId, { price: p.retailPrice })} className="text-[10px] px-1.5 py-0.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Giá lẻ</button>
                                        </div>
                                    )}
                                    {belowFloor && <span className="block text-[10px] mt-0.5" style={{ color: RUST }}>Dưới giá niêm yết tối thiểu ({vnd(floor)})</span>}
                                </td>
                                <td className="px-2 py-3 text-right font-bold text-base whitespace-nowrap align-top" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(orderLineTotal(it))}</td>
                                <td className="px-1 py-3 align-top"><button onClick={() => onRemove(it.productId)} style={{ color: RUST }}><X size={14} /></button></td>
                            </tr>
                            {p?.hasSeries && (
                                <tr style={{ borderBottom: `1px dashed ${LINE}` }}>
                                    <td></td>
                                    <td colSpan={6} className="px-2 pb-3">
                                        <div className="flex items-center gap-2 mb-1">
                                            <span className="text-xs opacity-60">Chọn series xuất bán — cần {it.qty} (còn tồn {available.length})</span>
                                            {it.series.length < it.qty && (
                                                <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BRASS}1A`, color: BRASS }}>Sẽ ở trạng thái chờ hàng</span>
                                            )}
                                        </div>
                                        <SeriesPicker available={available} selected={it.series} setSelected={(arr) => onUpdate(it.productId, { series: arr })} need={it.qty} />
                                        {it.series.length < it.qty && <p className="text-[11px] opacity-50 mt-1">Chưa đủ series (thường vì chưa đủ hàng trong kho) — vẫn tạo được đơn, bổ sung series sau khi nhập hàng về.</p>}
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
                    {order.cancelReason && <p className="text-[11px] opacity-70 mt-0.5">Lý do: {order.cancelReason}</p>}
                </div>
            </div>
        );
    }
    // Dùng trạng thái GIAO HÀNG hiện tại để xác định bước nào đã hoàn tất — tránh trường hợp
    // đơn từng được chuyển tới "Đã giao" rồi bị chỉnh lùi về "Chờ xử lý" nhưng bước cũ vẫn hiện xanh.
    const DELIVERY_RANK = { pending: 0, shipping: 1, delivered: 2, done: 3 };
    const rank = DELIVERY_RANK[order.status] ?? 0;
    const steps = [
        { label: "Đặt hàng", done: true, at: order.createdAt },
        { label: "Đang giao", done: rank >= 1, at: order.shippingAt },
        { label: "Đã giao", done: rank >= 2, at: order.deliveredAt },
        { label: "Đã thanh toán", done: !!order.paidCompleteAt, at: order.paidCompleteAt },
        { label: "Hoàn thành", done: rank >= 3, at: order.deliveredAt && order.paidCompleteAt ? (order.deliveredAt > order.paidCompleteAt ? order.deliveredAt : order.paidCompleteAt) : (order.deliveredAt || order.paidCompleteAt) },
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

function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Mở hoá đơn ra 1 TAB TRÌNH DUYỆT MỚI (thoát khỏi khung sandbox của artifact) và tự gọi lệnh in
// ngay trong chính tab đó — cách này không bị giới hạn bởi việc app đang chạy trong iframe cách ly.
// Trả về true nếu mở tab thành công, false nếu bị trình duyệt chặn popup (cần bấm link thủ công).
function printHTML(html) {
    const autoPrintHtml = html.replace(
        "</body>",
        `<script>
      window.addEventListener("load", function () { setTimeout(function () { window.print(); }, 150); });
    </script></body>`
    );
    const blob = new Blob([autoPrintHtml], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const win = window.open(url, "_blank");
    if (!win) return { ok: false, url };
    return { ok: true, url };
}

const PAPER_SIZES = [
    { id: "A4", label: "Khổ A4" },
    { id: "A5", label: "Khổ A5" },
    { id: "thermal80", label: "In nhiệt 80mm" },
    { id: "thermal58", label: "In nhiệt 58mm" },
];
const DEFAULT_PRINT_SETTINGS = {
    defaultPaperSize: "A4",
    showCompanyInfo: true, showBankAccount: true, showCustomerAddress: true, showCustomerPhone: true,
    showSeriesCol: true, showNotes: true, showAmountInWords: true, showSignatures: true,
};
function normalizePrintSettings(s) {
    return { ...DEFAULT_PRINT_SETTINGS, ...(s || {}) };
}

// Tạo HTML hoá đơn hoàn chỉnh (tự chứa CSS) theo khổ giấy và các mục thông tin được bật/tắt.
function buildInvoiceHTML(order, products, customer, paperSize, settings) {
    const c = orderCalc(order);
    const vp = vatPercent(order.vat);
    const isThermal = paperSize === "thermal80" || paperSize === "thermal58";
    const pageCss = {
        A4: `@page { size: A4; margin: 12mm; } body{ width:186mm; }`,
        A5: `@page { size: A5; margin: 8mm; } body{ width:132mm; }`,
        thermal80: `@page { size: 80mm auto; margin: 3mm; } body{ width:74mm; }`,
        thermal58: `@page { size: 58mm auto; margin: 2mm; } body{ width:54mm; }`,
    }[paperSize];
    const baseFont = isThermal ? 11 : (paperSize === "A5" ? 11.5 : 13);
    const custAddr = customer ? [customer.addressDetail, customer.ward, customer.province].filter(Boolean).join(", ") : "";
    const custPhone = customer?.phone || "";
    // Địa chỉ/điện thoại chính trên hoá đơn luôn là thông tin gốc (mặc định) của khách hàng.
    // Chỉ khi đơn này được chọn giao tới 1 địa chỉ khác với địa chỉ mặc định thì mới thêm riêng dòng "Giao hàng tại".
    const shipAddr = order.shippingAddress;
    const hasShipAddr = shipAddr && (shipAddr.addressDetail || shipAddr.province || shipAddr.ward);
    const shipAddrStr = hasShipAddr ? [shipAddr.addressDetail, shipAddr.ward, shipAddr.province].filter(Boolean).join(", ") : "";
    const isDifferentShip = hasShipAddr && (
        shipAddrStr !== custAddr ||
        (shipAddr.recipientPhone && shipAddr.recipientPhone !== custPhone) ||
        (shipAddr.recipientName && shipAddr.recipientName !== customer?.name)
    );
    const shipLine = isDifferentShip
        ? [shipAddrStr, shipAddr.recipientName, shipAddr.recipientPhone].filter(Boolean).join(" · ")
        : "";

    let body;
    if (isThermal) {
        const itemsHtml = order.items.map((it, i) => {
            const p = products.find((x) => x.id === it.productId);
            return `
        <div style="margin-bottom:5px;">
          <div>${i + 1}. ${escapeHtml(p?.name)}</div>
          <div style="display:flex;justify-content:space-between;">
            <span>${it.qty} x ${it.price.toLocaleString("vi-VN")}</span>
            <span><b>${orderLineTotal(it).toLocaleString("vi-VN")}</b></span>
          </div>
          ${settings.showSeriesCol && it.series?.length ? `<div style="opacity:0.7;">SN: ${escapeHtml(it.series.join(", "))}</div>` : ""}
        </div>`;
        }).join("");
        body = `
      <div style="text-align:center;margin-bottom:6px;">
        ${settings.showCompanyInfo ? `<div style="font-weight:bold;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 2}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 2}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>` : ""}
      </div>
      <div style="border-top:1px dashed #000;border-bottom:1px dashed #000;padding:5px 0;margin:6px 0;text-align:center;font-weight:bold;">HOÁ ĐƠN BÁN HÀNG</div>
      <div>Số: ${escapeHtml(order.code)}</div>
      <div>Ngày: ${new Date(order.createdAt).toLocaleString("vi-VN")}</div>
      <div>Khách: ${escapeHtml(customer?.name || "Khách lẻ")}</div>
      ${settings.showCustomerPhone && custPhone ? `<div>SĐT: ${escapeHtml(custPhone)}</div>` : ""}
      ${isDifferentShip ? `<div>Giao hàng tại: ${escapeHtml(shipLine)}</div>` : ""}
      <div style="border-top:1px dashed #000;margin:6px 0;"></div>
      ${itemsHtml}
      <div style="border-top:1px dashed #000;margin:6px 0;"></div>
      <div style="display:flex;justify-content:space-between;"><span>Tổng tiền</span><span>${c.subtotal.toLocaleString("vi-VN")}</span></div>
      ${order.orderDiscount > 0 ? `<div style="display:flex;justify-content:space-between;"><span>Chiết khấu</span><span>-${c.discountAmount.toLocaleString("vi-VN")}</span></div>` : ""}
      ${order.shippingFee > 0 ? `<div style="display:flex;justify-content:space-between;"><span>Phí giao</span><span>${order.shippingFee.toLocaleString("vi-VN")}</span></div>` : ""}
      <div style="display:flex;justify-content:space-between;font-weight:bold;border-top:1px dashed #000;padding-top:4px;margin-top:4px;"><span>TỔNG CỘNG</span><span>${c.payable.toLocaleString("vi-VN")}</span></div>
      <div style="border-top:1px dashed #000;margin:8px 0;"></div>
      <div style="text-align:center;">Cảm ơn quý khách!</div>
    `;
    } else {
        const itemsHtml = order.items.map((it, i) => {
            const p = products.find((x) => x.id === it.productId);
            return `
        <tr>
          <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
          <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(p?.name)}</td>
          <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${escapeHtml(p?.unit)}</td>
          <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${warrantyLabel(p?.warrantyMonths || 0)}</td>
          <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
          <td style="border:1px solid #111;padding:6px 8px;text-align:right;">${it.price.toLocaleString("vi-VN")}</td>
          <td style="border:1px solid #111;padding:6px 8px;text-align:right;">${orderLineTotal(it).toLocaleString("vi-VN")}</td>
        </tr>`;
        }).join("");
        // Danh sách số series (nếu có) — trình bày thành bảng riêng, đặt phía dưới khối chữ ký. Mỗi số series 1 dòng.
        const seriesRows = [];
        order.items.forEach((it) => {
            const p = products.find((x) => x.id === it.productId);
            (it.series || []).forEach((sn) => { seriesRows.push({ product: p?.name || "", sn }); });
        });
        const seriesTableHtml = (settings.showSeriesCol && seriesRows.length > 0) ? `
      <div style="page-break-before:always; break-before:page; padding-top:6mm;">
        <p style="font-weight:bold;margin-bottom:8px;">DANH SÁCH SỐ SERIES — Đơn ${escapeHtml(order.code)}</p>
        <table style="width:100%;border-collapse:collapse;">
          <thead><tr>
            <th style="border:1px solid #111;padding:5px 8px;font-size:12px;width:40px;">STT</th>
            <th style="border:1px solid #111;padding:5px 8px;font-size:12px;">Sản phẩm</th>
            <th style="border:1px solid #111;padding:5px 8px;font-size:12px;">Số Series</th>
          </tr></thead>
          <tbody>
            ${seriesRows.map((r, i) => `
              <tr>
                <td style="border:1px solid #111;padding:5px 8px;text-align:center;">${i + 1}</td>
                <td style="border:1px solid #111;padding:5px 8px;">${escapeHtml(r.product)}</td>
                <td style="border:1px solid #111;padding:5px 8px;font-family:monospace;">${escapeHtml(r.sn)}</td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>` : "";
        body = `
      ${settings.showCompanyInfo ? `
      <div style="border-bottom:2px solid #111;padding-bottom:10px;margin-bottom:14px;line-height:1.6;display:flex;align-items:flex-start;gap:14px;">
        ${HILI_LOGO_SRC ? `<img src="${HILI_LOGO_SRC}" alt="Hili" style="width:78px;height:78px;object-fit:cover;border-radius:4px;flex-shrink:0;" />` : ""}
        <div>
          <div><b>Tên đơn vị: </b>${escapeHtml(COMPANY_INFO.name)}</div>
          <div><b>Địa chỉ: </b>${escapeHtml(COMPANY_INFO.address)}</div>
          <div><b>MST: </b>${escapeHtml(COMPANY_INFO.taxCode)}</div>
          ${settings.showBankAccount ? `<div><b>Tài khoản số: </b>${escapeHtml(COMPANY_INFO.bankAccount)}</div>` : ""}
          <div><b>Liên hệ: </b>${escapeHtml(COMPANY_INFO.phone)}</div>
        </div>
      </div>` : ""}
      <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:4px 0 10px;">HÓA ĐƠN BÁN HÀNG</h1>
      <div style="display:flex;justify-content:center;gap:24px;margin-bottom:14px;">
        <span>Ngày ${new Date(order.createdAt).getDate()} tháng ${new Date(order.createdAt).getMonth() + 1} năm ${new Date(order.createdAt).getFullYear()}</span>
        <span><b>Số:</b> ${escapeHtml(order.code)}</span>
      </div>
      <div style="margin-bottom:10px;"><b>Khách hàng: </b>${escapeHtml(customer?.name || "Khách lẻ")}</div>
      ${settings.showCustomerAddress && custAddr ? `<div style="margin-bottom:10px;"><b>Địa chỉ: </b>${escapeHtml(custAddr)}</div>` : ""}
      ${settings.showCustomerPhone && custPhone ? `<div style="margin-bottom:10px;"><b>Điện thoại: </b>${escapeHtml(custPhone)}</div>` : ""}
      ${isDifferentShip ? `<div style="margin-bottom:14px;"><b>Giao hàng tại: </b>${escapeHtml(shipLine)}</div>` : ""}
      <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
        <thead><tr>
          ${["STT", "Tên SP", "ĐVT", "Bảo hành", "SL", "Đơn giá", "Thành tiền"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
        </tr></thead>
        <tbody>
          ${itemsHtml}
          <tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Cộng:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${c.subtotal.toLocaleString("vi-VN")}</td></tr>
          <tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">${order.vat === "KCT" ? "Thuế KCT" : `Đã bao gồm Thuế (${vp}%)`}:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${order.vat === "KCT" ? "—" : c.vatTotal.toLocaleString("vi-VN")}</td></tr>
          ${order.orderDiscount > 0 ? `<tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Chiết khấu${order.discountType === "percent" ? ` (${order.orderDiscount}%)` : ""}:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${c.discountAmount.toLocaleString("vi-VN")}</td></tr>` : ""}
          ${order.shippingFee > 0 ? `<tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Phí giao hàng:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${order.shippingFee.toLocaleString("vi-VN")}</td></tr>` : ""}
          <tr style="font-weight:bold;"><td colspan="6" style="border:1px solid #111;padding:6px 8px;text-align:right;">Tổng cộng:</td><td style="border:1px solid #111;padding:6px 8px;text-align:right;">${c.payable.toLocaleString("vi-VN")}</td></tr>
        </tbody>
      </table>
      ${settings.showAmountInWords ? `<p style="margin-bottom:18px;"><i>Tổng số tiền bằng chữ: ${escapeHtml(soTienBangChu(c.payable))}.</i></p>` : ""}
      <p style="font-weight:bold;margin-bottom:${settings.showSignatures ? 28 : 6}px;">VUI LÒNG KIỂM TRA HÀNG NGAY KHI NHẬN</p>
      ${settings.showSignatures ? `
      <div style="display:flex;justify-content:space-between;text-align:center;">
        <div style="width:30%;"><b>Người lập phiếu</b><br/><span style="font-size:11px;">(Ký, họ tên)</span><br/><br/><br/>${escapeHtml(order.seller)}</div>
        <div style="width:30%;"><b>Người nhận hàng</b><br/><span style="font-size:11px;">(Ký, họ tên)</span></div>
        <div style="width:30%;"><b>Giám đốc</b><br/><span style="font-size:11px;">(Ký, họ tên)</span></div>
      </div>` : ""}
      ${seriesTableHtml}
    `;
    }

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(order.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
      table { page-break-inside: auto; }
      tr { page-break-inside: avoid; }
    </style>
  </head><body>${body}</body></html>`;
}

// Biên bản bàn giao hàng hóa (BBBG) — in từ đơn hàng. Bên B (Hilitek) cố định theo COMPANY_INFO,
// Bên A lấy thông tin khách hàng của đơn (địa chỉ, điện thoại, đại diện/chức vụ đã lưu trong hồ sơ khách B2B/Doanh nghiệp).
function buildHandoverHTML(order, products, customer, paperSize) {
    const isA5 = paperSize === "A5";
    const pageCss = isA5
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = isA5 ? 10.5 : 12.5;
    const custAddr = customer ? [customer.addressDetail, customer.ward, customer.province].filter(Boolean).join(", ") : "";
    const d = new Date(order.deliveredAt || order.date || order.createdAt);
    const custRepName = customer?.representativeName || "";
    const custRepTitle = customer?.representativeTitle || "";
    const DOTS = "……………………………………";
    // Địa điểm giao hàng — ưu tiên địa chỉ giao hàng đã chọn riêng cho đơn này, nếu không có thì dùng địa chỉ khách hàng.
    const shipAddr = order.shippingAddress;
    const deliveryLocation = (shipAddr && (shipAddr.addressDetail || shipAddr.province || shipAddr.ward))
        ? [shipAddr.addressDetail, shipAddr.ward, shipAddr.province].filter(Boolean).join(", ")
        : custAddr;

    const itemsRows = order.items.map((it, i) => {
        const p = products.find((x) => x.id === it.productId);
        const hasSeries = !!p?.hasSeries || (it.series && it.series.length > 0);
        return `<tr>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(p?.name)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${escapeHtml(p?.unit)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${warrantyLabel(p?.warrantyMonths || 0)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">Hộp sản phẩm</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${hasSeries ? "Kèm theo phiếu series" : ""}</td>
    </tr>`;
    }).join("") + `
    <tr>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
      <td style="border:1px solid #111;padding:16px 8px;">&nbsp;</td>
    </tr>`;

    // Danh sách số series — đặt ở trang 2, giống cách trình bày của hoá đơn.
    const seriesRows = [];
    order.items.forEach((it) => {
        const p = products.find((x) => x.id === it.productId);
        (it.series || []).forEach((sn) => { seriesRows.push({ product: p?.name || "", sn }); });
    });
    const seriesTableHtml = seriesRows.length > 0 ? `
    <div style="page-break-before:always; break-before:page; padding-top:6mm;">
      <p style="font-weight:bold;margin-bottom:8px;">DANH SÁCH SỐ SERIES — ${escapeHtml(order.code)}</p>
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr>
          <th style="border:1px solid #111;padding:5px 8px;font-size:12px;width:40px;">STT</th>
          <th style="border:1px solid #111;padding:5px 8px;font-size:12px;">Sản phẩm</th>
          <th style="border:1px solid #111;padding:5px 8px;font-size:12px;">Số Series</th>
        </tr></thead>
        <tbody>
          ${seriesRows.map((r, i) => `
            <tr>
              <td style="border:1px solid #111;padding:5px 8px;text-align:center;">${i + 1}</td>
              <td style="border:1px solid #111;padding:5px 8px;">${escapeHtml(r.product)}</td>
              <td style="border:1px solid #111;padding:5px 8px;font-family:monospace;">${escapeHtml(r.sn)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>` : "";

    const body = `
    <table style="width:100%;border-collapse:collapse;margin-bottom:2px;">
      <tr>
        <td style="width:50%;text-align:center;vertical-align:top;padding:0;">
          <div style="font-weight:bold;">${escapeHtml(COMPANY_INFO.name)}</div>
          <div>Số: <b style="color:#b0462f;">${escapeHtml(order.code)}</b></div>
        </td>
        <td style="width:50%;text-align:center;vertical-align:top;padding:0;">
          <div style="font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
          <div style="font-weight:bold;display:inline-block;border-bottom:1px solid #111;padding-bottom:2px;">Độc lập – Tự do – Hạnh phúc</div>
          <div style="margin-top:2px;"><i>Ngày ${d.getDate()} tháng ${d.getMonth() + 1} năm ${d.getFullYear()}</i></div>
        </td>
      </tr>
    </table>

    <h1 style="text-align:center;font-size:${isA5 ? 17 : 21}px;letter-spacing:1px;margin:16px 0 10px;">BIÊN BẢN BÀN GIAO HÀNG HÓA</h1>

    <div style="display:flex;justify-content:space-between;margin-bottom:12px;">
      <span><i>Hôm nay, ngày ${d.getDate()} tháng ${d.getMonth() + 1} năm ${d.getFullYear()}</i></span>
      <span><i>Tại ${escapeHtml(COMPANY_INFO.city)}</i></span>
    </div>

    <p style="margin-bottom:4px;"><b>BÊN A (Bên nhận hàng): </b>${escapeHtml(customer?.name || "Khách lẻ")}</p>
    <p style="margin-bottom:4px;"><b>Địa chỉ: </b>${escapeHtml(custAddr || DOTS)}</p>
    <p style="margin-bottom:4px;"><b>Điện thoại: </b>${escapeHtml(customer?.phone || DOTS)}</p>
    <div style="display:flex;justify-content:space-between;margin-bottom:14px;">
      <span><b>Đại diện Ông/bà: </b>${escapeHtml(custRepName || DOTS)}</span>
      <span><b>Chức vụ: </b>${escapeHtml(custRepTitle || DOTS)}</span>
    </div>

    <p style="margin-bottom:4px;"><b>BÊN B (Bên giao hàng): </b>${escapeHtml(COMPANY_INFO.name)}</p>
    <p style="margin-bottom:4px;"><b>Địa chỉ: </b>${escapeHtml(COMPANY_INFO.address)}</p>
    <p style="margin-bottom:4px;"><b>Điện thoại: </b>${escapeHtml(COMPANY_INFO.phone)}</p>
    <div style="display:flex;justify-content:space-between;margin-bottom:16px;">
      <span><b>Đại diện Ông/bà: </b>${escapeHtml(COMPANY_INFO.representativeName)}</span>
      <span><b>Chức vụ: </b>${escapeHtml(COMPANY_INFO.representativeTitle)}</span>
    </div>

    <p style="margin-bottom:8px;">Hai bên cùng nhau thống nhất số lượng giao hàng như sau:</p>

    <table style="width:100%;border-collapse:collapse;margin-bottom:16px;">
      <thead><tr>
        ${["STT", "Tên SP", "Đơn vị tính", "Bảo hành", "Số lượng", "Quy cách", "Ghi chú"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
      </tr></thead>
      <tbody>${itemsRows}</tbody>
    </table>

    <p style="margin-bottom:14px;">Bên B đã giao hàng đầy đủ cho Bên A trong tình trạng hàng mới 100%.</p>
    <p style="margin-bottom:14px;"><b>Địa điểm giao hàng: </b>${escapeHtml(deliveryLocation || DOTS)}</p>
    <p style="margin-bottom:6px;">Bên B đã giao các giấy tờ kèm theo cho Bên A bao gồm:</p>
    <div style="margin-bottom:14px;padding-left:4px;">
      <div style="margin-bottom:4px;">☐ Hóa đơn tài chính</div>
      <div style="margin-bottom:4px;">☐ Phiếu và tem bảo hành theo máy</div>
      <div style="margin-bottom:4px;">☐ Biên bản bàn giao</div>
    </div>
    <p style="margin-bottom:8px;">Bên A xác nhận Bên B đã giao cho Bên A đúng chủng loại và đủ số lượng hàng như trên.</p>
    <p style="margin-bottom:28px;">Hai bên đồng ý, thống nhất ký tên. Biên bản được lập thành 02 bản, mỗi bên giữ 01 bản có giá trị pháp lý như nhau.</p>

    <div style="display:flex;justify-content:space-between;text-align:center;">
      <div style="width:45%;">
        <b>ĐẠI DIỆN BÊN A</b>
        <div style="height:70px;"></div>
        <b>${escapeHtml(custRepName)}</b>
      </div>
      <div style="width:45%;">
        <b>ĐẠI DIỆN BÊN B</b>
        <div style="height:70px;"></div>
        <b>${escapeHtml(COMPANY_INFO.representativeName)}</b>
      </div>
    </div>
    ${seriesTableHtml}
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>BBBG ${escapeHtml(order.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
      table { page-break-inside: auto; }
      tr { page-break-inside: avoid; }
    </style>
  </head><body>${body}</body></html>`;
}
function buildQuoteHTML(q, products, paperSize) {
    const c = quoteCalc(q);
    const vp = vatPercent(q.vat);
    const isThermal = paperSize === "thermal80" || paperSize === "thermal58";
    const pageCss = {
        A4: `@page { size: A4; margin: 12mm; } body{ width:186mm; }`,
        A5: `@page { size: A5; margin: 8mm; } body{ width:132mm; }`,
        thermal80: `@page { size: 80mm auto; margin: 3mm; } body{ width:74mm; }`,
        thermal58: `@page { size: 58mm auto; margin: 2mm; } body{ width:54mm; }`,
    }[paperSize];
    const baseFont = isThermal ? 11 : (paperSize === "A5" ? 11.5 : 13);
    const exp = quoteExpiryInfo(q);
    const validDays = Math.max(1, Math.round((new Date(q.expiryDate).getTime() - new Date(q.createdAt).getTime()) / 86400000));
    const itemsHtml = q.items.map((it, i) => {
        const p = products.find((x) => x.id === it.productId);
        if (isThermal) {
            return `<div style="margin-bottom:5px;"><div>${i + 1}. ${escapeHtml(p?.name)}</div>
        <div style="display:flex;justify-content:space-between;"><span>${it.qty} x ${it.price.toLocaleString("vi-VN")}</span><span><b>${(it.qty * it.price).toLocaleString("vi-VN")}</b></span></div></div>`;
        }
        return `<tr>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(p?.name)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${escapeHtml(p?.unit)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${warrantyLabel(p?.warrantyMonths || 0)}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:right;">${it.price.toLocaleString("vi-VN")}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:right;">${(it.qty * it.price).toLocaleString("vi-VN")}</td>
    </tr>`;
    }).join("");

    let body;
    if (isThermal) {
        body = `
      <div style="text-align:center;margin-bottom:6px;">
        <div style="font-weight:bold;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 2}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 2}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="border-top:1px dashed #000;border-bottom:1px dashed #000;padding:5px 0;margin:6px 0;text-align:center;font-weight:bold;">BẢNG BÁO GIÁ</div>
      <div>Số: ${escapeHtml(q.code)}</div>
      <div>Ngày: ${new Date(q.createdAt).toLocaleString("vi-VN")}</div>
      <div>Hiệu lực đến: ${new Date(q.expiryDate).toLocaleDateString("vi-VN")}</div>
      <div>Khách: ${escapeHtml(q.customerName || "—")}</div>
      ${q.customerPhone ? `<div>SĐT: ${escapeHtml(q.customerPhone)}</div>` : ""}
      <div style="border-top:1px dashed #000;margin:6px 0;"></div>
      ${itemsHtml}
      <div style="border-top:1px dashed #000;margin:6px 0;"></div>
      <div style="display:flex;justify-content:space-between;"><span>Tổng tiền</span><span>${c.subtotal.toLocaleString("vi-VN")}</span></div>
      ${q.orderDiscount > 0 ? `<div style="display:flex;justify-content:space-between;"><span>Chiết khấu</span><span>-${c.discountAmount.toLocaleString("vi-VN")}</span></div>` : ""}
      <div style="display:flex;justify-content:space-between;font-weight:bold;border-top:1px dashed #000;padding-top:4px;margin-top:4px;"><span>TỔNG CỘNG</span><span>${c.total.toLocaleString("vi-VN")}</span></div>
      <div style="border-top:1px dashed #000;margin:8px 0;"></div>
      <div style="text-align:center;">Cảm ơn quý khách!</div>
    `;
    } else {
        body = `
      <div style="border-bottom:2px solid #111;padding-bottom:10px;margin-bottom:14px;line-height:1.6;display:flex;align-items:flex-start;gap:14px;">
        ${HILI_LOGO_SRC ? `<img src="${HILI_LOGO_SRC}" alt="Hili" style="width:78px;height:78px;object-fit:cover;border-radius:4px;flex-shrink:0;" />` : ""}
        <div>
          <div><b>Tên đơn vị: </b>${escapeHtml(COMPANY_INFO.name)}</div>
          <div><b>Địa chỉ: </b>${escapeHtml(COMPANY_INFO.address)}</div>
          <div><b>MST: </b>${escapeHtml(COMPANY_INFO.taxCode)}</div>
          <div><b>Tài khoản số: </b>${escapeHtml(COMPANY_INFO.bankAccount)}</div>
          <div><b>Liên hệ: </b>${escapeHtml(COMPANY_INFO.phone)}</div>
        </div>
      </div>
      <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:4px 0 10px;">BẢNG BÁO GIÁ</h1>
      <div style="display:flex;justify-content:center;gap:24px;margin-bottom:6px;">
        <span>Ngày ${new Date(q.createdAt).getDate()} tháng ${new Date(q.createdAt).getMonth() + 1} năm ${new Date(q.createdAt).getFullYear()}</span>
        <span><b>Số:</b> ${escapeHtml(q.code)}</span>
      </div>
      <div style="margin-bottom:6px;"><i>Kính gửi,</i></div>
      <div style="margin-bottom:4px;"><b>Khách hàng: </b>${escapeHtml(q.customerName || "—")}</div>
      ${q.customerAddress ? `<div style="margin-bottom:4px;"><b>Địa chỉ: </b>${escapeHtml(q.customerAddress)}</div>` : ""}
      ${q.customerTaxCode ? `<div style="margin-bottom:4px;"><b>Mã số thuế: </b>${escapeHtml(q.customerTaxCode)}</div>` : ""}
      ${q.customerPhone ? `<div style="margin-bottom:14px;"><b>Điện thoại: </b>${escapeHtml(q.customerPhone)}</div>` : ""}
      <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
        <thead><tr>
          ${["STT", "Tên SP", "ĐVT", "Bảo hành", "SL", "Đơn giá", "Thành tiền"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
        </tr></thead>
        <tbody>
          ${itemsHtml}
          <tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Cộng:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${c.subtotal.toLocaleString("vi-VN")}</td></tr>
          <tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">${q.vat === "KCT" ? "Thuế KCT" : `Đã bao gồm Thuế (${vp}%)`}:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${q.vat === "KCT" ? "—" : c.vatTotal.toLocaleString("vi-VN")}</td></tr>
          ${q.orderDiscount > 0 ? `<tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Chiết khấu${q.discountType === "percent" ? ` (${q.orderDiscount}%)` : ""}:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${c.discountAmount.toLocaleString("vi-VN")}</td></tr>` : ""}
          ${q.shippingFee > 0 ? `<tr><td colspan="6" style="border:1px solid #111;padding:5px 8px;text-align:right;">Phí giao hàng:</td><td style="border:1px solid #111;padding:5px 8px;text-align:right;">${q.shippingFee.toLocaleString("vi-VN")}</td></tr>` : ""}
          <tr style="font-weight:bold;"><td colspan="6" style="border:1px solid #111;padding:6px 8px;text-align:right;">Tổng cộng:</td><td style="border:1px solid #111;padding:6px 8px;text-align:right;">${c.total.toLocaleString("vi-VN")}</td></tr>
        </tbody>
      </table>
      <p style="margin-bottom:18px;"><i>Tổng số tiền bằng chữ: ${escapeHtml(soTienBangChu(c.total))}.</i></p>
      <div style="font-size:12px;line-height:1.9;">
        <div>1. Báo giá đã bao gồm VAT.</div>
        <div>2. Cam kết hàng chính hãng, bảo hành chính hãng.</div>
        <div>3. Báo giá có hiệu lực trong vòng ${validDays} ngày kể từ ngày lập.</div>
        <div>4. Báo giá chưa bao gồm chi phí phát sinh (nếu có).</div>
        <div>5. Vui lòng liên hệ ${escapeHtml(COMPANY_INFO.phone)} hoặc email ${escapeHtml(COMPANY_INFO.email)} để được tư vấn thêm.</div>
      </div>
      <div style="display:flex;justify-content:flex-end;margin-top:26px;">
        <div style="text-align:center;">
          <div>${thuNgayThangNam(q.createdAt)}</div>
          <div style="font-weight:bold;margin-top:6px;">Nhân viên lập phiếu</div>
          <div style="font-size:11px;">(Ký và ghi rõ họ tên)</div>
          <div style="height:55px;"></div>
        </div>
      </div>
    `;
    }
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(q.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

/* ---------------- Phiếu bảo hành ---------------- */
function buildWarrantyTicketHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const statusLabel = WARRANTY_TICKET_STATUSES.find((s) => s.id === ticket.status)?.label || "";
    // Chỉ dùng nội bộ (mã đơn hàng gốc) — không thể hiện ra phiếu in cho khách.
    const itemsHtml = ticket.items.map((it, i) => `
    <tr>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productCode)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productName)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.series.join(", ") || "—")}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${warrantyLabel(it.warrantyMonths)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.condition || "—")}</td>
    </tr>`).join("");
    const totalQty = ticket.items.reduce((s, it) => s + it.qty, 0);

    const isPending = ticket.status === "pending";
    const title = isPending ? "PHIẾU TIẾP NHẬN BẢO HÀNH" : "PHIẾU BẢO HÀNH";
    const returnLine = isPending
        ? `<div><b>Ngày hẹn trả: </b>Chờ kiểm tra và báo lại${ticket.returnDate ? ` trước ${new Date(ticket.returnDate).toLocaleDateString("vi-VN")}` : ""}</div>`
        : `<div><b>Ngày hẹn trả: </b>${ticket.returnDate ? new Date(ticket.returnDate).toLocaleDateString("vi-VN") : "—"}</div>`;

    // Bảng kết quả xử lý (đổi sản phẩm / hoàn tiền) — chỉ hiện khi đã Xác nhận hoặc Đã trả khách.
    const resolutionHtml = (ticket.status === "confirmed" || ticket.status === "completed") && ticket.resolutionType ? `
    <table style="width:100%;border-collapse:collapse;margin:14px 0;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">KẾT QUẢ XỬ LÝ BẢO HÀNH</th></tr></thead>
      <tbody>
        ${ticket.resolutionType === "exchange" ? `
          <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Hình thức</td><td style="border:1px solid #111;padding:6px 8px;">Đổi sản phẩm tương tự</td></tr>
          <tr><td style="border:1px solid #111;padding:6px 8px;">Sản phẩm đổi</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.exchangeProductName)} (${escapeHtml(ticket.exchangeProductCode)})</td></tr>
          ${ticket.exchangeSeries.length > 0 ? `<tr><td style="border:1px solid #111;padding:6px 8px;">Serial</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.exchangeSeries.join(", "))}</td></tr>` : ""}
          <tr><td style="border:1px solid #111;padding:6px 8px;">Số lượng</td><td style="border:1px solid #111;padding:6px 8px;">${ticket.exchangeQty}</td></tr>
        ` : `
          <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Hình thức</td><td style="border:1px solid #111;padding:6px 8px;">Hoàn tiền</td></tr>
          <tr><td style="border:1px solid #111;padding:6px 8px;">Số tiền hoàn lại</td><td style="border:1px solid #111;padding:6px 8px;font-weight:bold;">${ticket.refundAmount.toLocaleString("vi-VN")}đ</td></tr>
        `}
      </tbody>
    </table>` : "";

    const rejectHtml = ticket.status === "rejected" ? `
    <div style="margin:14px 0;padding:10px;border:1px solid #111;">
      <b>TỪ CHỐI BẢO HÀNH</b>
      <div style="margin-top:6px;"><b>Sản phẩm bị từ chối bảo hành:</b></div>
      <ul style="margin:4px 0 8px 20px;padding:0;">
        ${ticket.items.map((it) => `<li>${escapeHtml(it.productName)} (${escapeHtml(it.productCode)})${it.series.length ? ` — Serial: ${escapeHtml(it.series.join(", "))}` : ""} — SL: ${it.qty}</li>`).join("")}
      </ul>
      <div><b>Lý do từ chối: </b>${escapeHtml(ticket.rejectReason || "—")}</div>
    </div>` : "";

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.code)}</div>
        <div>Ngày tạo: ${new Date(ticket.createdAt).toLocaleDateString("vi-VN")}</div>
        <div>Trạng thái: ${escapeHtml(statusLabel)}</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:8px 0 16px;">${title}</h1>

    <div style="display:flex;gap:24px;margin-bottom:12px;">
      <div style="flex:1;">
        <div><b>Khách hàng: </b>${escapeHtml(ticket.customerName || "—")}</div>
        <div><b>Điện thoại: </b>${escapeHtml(ticket.customerPhone || "—")}</div>
        <div><b>Địa chỉ: </b>${escapeHtml(ticket.customerAddress || "—")}</div>
      </div>
      <div style="flex:1;">
        <div><b>Người lập phiếu: </b>${escapeHtml(ticket.createdBy)}</div>
        <div><b>Địa chỉ nhận bảo hành: </b>${escapeHtml(COMPANY_INFO.address)}</div>
        <div><b>Ngày nhận: </b>${ticket.receivedDate ? new Date(ticket.receivedDate).toLocaleDateString("vi-VN") : "—"}</div>
        ${returnLine}
      </div>
    </div>

    <table style="width:100%;border-collapse:collapse;margin-bottom:12px;">
      <thead><tr>
        ${["STT", "Mã SKU", "Tên sản phẩm", "Serial", "SL", "Bảo hành", "Tình trạng sản phẩm"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
      </tr></thead>
      <tbody>${itemsHtml}</tbody>
    </table>
    <div style="text-align:right;margin-bottom:8px;"><b>Tổng số lượng: ${totalQty}</b></div>

    ${resolutionHtml}
    ${rejectHtml}

    ${ticket.note ? `<div style="margin:14px 0;"><b>Ghi chú: </b>${escapeHtml(ticket.note)}</div>` : ""}

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Khách hàng</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
      <div>
        <b>Người lập phiếu</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(ticket.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

// Phiếu Xuất trả bảo hành (XTBH-{mã phiếu bảo hành gốc}) — chứng từ xác nhận đã xuất trả/đổi sản phẩm bảo hành cho khách,
// lập khi phiếu bảo hành chuyển sang trạng thái "Đã trả khách".
function buildWarrantyReturnSlipHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const resolutionRows = ticket.resolutionType === "exchange" ? `
    <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Hình thức</td><td style="border:1px solid #111;padding:6px 8px;">Đổi sản phẩm tương tự</td></tr>
    <tr><td style="border:1px solid #111;padding:6px 8px;">Sản phẩm xuất trả</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.exchangeProductName)} (${escapeHtml(ticket.exchangeProductCode)})</td></tr>
    ${ticket.exchangeSeries.length > 0 ? `<tr><td style="border:1px solid #111;padding:6px 8px;">Serial</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.exchangeSeries.join(", "))}</td></tr>` : ""}
    <tr><td style="border:1px solid #111;padding:6px 8px;">Số lượng</td><td style="border:1px solid #111;padding:6px 8px;">${ticket.exchangeQty}</td></tr>
  ` : `
    <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Hình thức</td><td style="border:1px solid #111;padding:6px 8px;">Hoàn tiền</td></tr>
    <tr><td style="border:1px solid #111;padding:6px 8px;">Số tiền đã hoàn</td><td style="border:1px solid #111;padding:6px 8px;font-weight:bold;">${ticket.refundAmount.toLocaleString("vi-VN")}đ</td></tr>
  `;
    const originalItemsHtml = ticket.items.map((it, i) => `
    <tr>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productCode)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productName)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.series.join(", ") || "—")}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
    </tr>`).join("");

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.xtbhCode || `XTBH-${ticket.code}`)}</div>
        <div>Phiếu bảo hành gốc: ${escapeHtml(ticket.code)}</div>
        <div>Ngày lập: ${new Date().toLocaleDateString("vi-VN")}</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:8px 0 16px;">PHIẾU XUẤT TRẢ BẢO HÀNH</h1>

    <div style="margin-bottom:12px;">
      <div><b>Khách hàng: </b>${escapeHtml(ticket.customerName || "—")}</div>
      <div><b>Điện thoại: </b>${escapeHtml(ticket.customerPhone || "—")}</div>
      <div><b>Địa chỉ: </b>${escapeHtml(ticket.customerAddress || "—")}</div>
    </div>

    <p style="margin-bottom:6px;"><b>Sản phẩm bảo hành đã nhận trước đó (theo phiếu ${escapeHtml(ticket.code)}):</b></p>
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr>
        ${["STT", "Mã SKU", "Tên sản phẩm", "Serial", "SL"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
      </tr></thead>
      <tbody>${originalItemsHtml}</tbody>
    </table>

    <p style="margin-bottom:6px;"><b>Kết quả xuất trả cho khách:</b></p>
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
      <tbody>${resolutionRows}</tbody>
    </table>

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Khách hàng nhận hàng</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
      <div>
        <b>Người lập phiếu</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(ticket.xtbhCode || `XTBH-${ticket.code}`)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

// Phiếu Từ chối bảo hành — xác nhận sản phẩm không đủ điều kiện bảo hành, xuất trả lại nguyên trạng cho khách.
function buildWarrantyRejectSlipHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const itemsHtml = ticket.items.map((it, i) => `
    <tr>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productCode)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.productName)}</td>
      <td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(it.series.join(", ") || "—")}</td>
      <td style="border:1px solid #111;padding:6px 8px;text-align:center;">${it.qty}</td>
    </tr>`).join("");

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.code)}</div>
        <div>Ngày lập: ${new Date().toLocaleDateString("vi-VN")}</div>
        <div>Trạng thái: Từ chối bảo hành</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:8px 0 16px;">PHIẾU TỪ CHỐI BẢO HÀNH</h1>

    <div style="margin-bottom:12px;">
      <div><b>Khách hàng: </b>${escapeHtml(ticket.customerName || "—")}</div>
      <div><b>Điện thoại: </b>${escapeHtml(ticket.customerPhone || "—")}</div>
      <div><b>Địa chỉ: </b>${escapeHtml(ticket.customerAddress || "—")}</div>
    </div>

    <p style="margin-bottom:6px;">Sau khi kiểm tra, các sản phẩm dưới đây <b>không đủ điều kiện bảo hành</b> và được xuất trả lại nguyên trạng cho khách hàng:</p>
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr>
        ${["STT", "Mã SKU", "Tên sản phẩm", "Serial", "SL"].map((h) => `<th style="border:1px solid #111;padding:6px 8px;font-size:12px;">${h}</th>`).join("")}
      </tr></thead>
      <tbody>${itemsHtml}</tbody>
    </table>

    <div style="margin-bottom:20px;padding:10px;border:1px solid #111;">
      <b>Lý do từ chối bảo hành: </b>${escapeHtml(ticket.rejectReason || "—")}
    </div>

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Khách hàng nhận lại hàng</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
      <div>
        <b>Người lập phiếu</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Tu-choi-BH-${escapeHtml(ticket.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

/* ---------------- Phiếu dịch vụ sửa chữa — Phiếu tiếp nhận & báo chi phí dự kiến ---------------- */
function buildRepairReceiptHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const isReceived = ticket.status === "received";
    const vp = vatPercent(ticket.vat);
    const vatLabel = VAT_OPTIONS.find((v) => v.id === ticket.vat)?.label || "";

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.code)}</div>
        <div>Ngày lập: ${new Date(ticket.createdAt).toLocaleDateString("vi-VN")}</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 17 : 21}px;letter-spacing:1px;margin:8px 0 4px;">${isReceived ? "PHIẾU TIẾP NHẬN" : "PHIẾU BÁO GIÁ SỬA CHỮA"}</h1>
    <p style="text-align:center;font-size:${baseFont}px;margin:0 0 16px;">${isReceived ? "VÀ BÁO GIÁ CHI PHÍ SỬA CHỮA DỰ KIẾN" : "CHI PHÍ THỰC TẾ"}</p>

    <div style="display:flex;gap:24px;margin-bottom:12px;">
      <div style="flex:1;">
        <div><b>Khách hàng: </b>${escapeHtml(ticket.customerName || "—")}</div>
        <div><b>Điện thoại: </b>${escapeHtml(ticket.customerPhone || "—")}</div>
        <div><b>Địa chỉ: </b>${escapeHtml(ticket.customerAddress || "—")}</div>
      </div>
      <div style="flex:1;">
        <div><b>Người lập phiếu: </b>${escapeHtml(ticket.createdBy)}</div>
        <div><b>Ngày nhận: </b>${ticket.receivedDate ? new Date(ticket.receivedDate).toLocaleDateString("vi-VN") : "—"}</div>
        <div><b>Ngày hẹn trả: </b>${ticket.returnDate ? new Date(ticket.returnDate).toLocaleDateString("vi-VN") : "Chưa xác định — sẽ báo sau khi kiểm tra"}</div>
      </div>
    </div>

    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">THÔNG TIN THIẾT BỊ TIẾP NHẬN</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Tên thiết bị / sản phẩm</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.deviceName || "—")}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Hãng / thương hiệu</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.deviceBrand || "—")}</td></tr>
        ${ticket.serial ? `<tr><td style="border:1px solid #111;padding:6px 8px;">Serial / IMEI</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.serial)}</td></tr>` : ""}
        <tr><td style="border:1px solid #111;padding:6px 8px;">Tình trạng / mô tả lỗi</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.issueDescription || "—")}</td></tr>
      </tbody>
    </table>

    ${isReceived ? `
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">BÁO GIÁ CHI PHÍ SỬA CHỮA DỰ KIẾN</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:8px;width:35%;font-weight:bold;">Chi phí dự kiến</td><td style="border:1px solid #111;padding:8px;font-weight:bold;font-size:${baseFont + 2}px;">${ticket.estimatedCost > 0 ? `${ticket.estimatedCost.toLocaleString("vi-VN")}đ` : "Chưa xác định — sẽ báo sau khi kiểm tra chi tiết"}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">VAT</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(vatLabel)}</td></tr>
      </tbody>
    </table>
    <p style="font-size:${baseFont - 1}px;font-style:italic;margin-bottom:16px;">* Chi phí trên chỉ là báo giá dự kiến ban đầu, có thể thay đổi sau khi kiểm tra kỹ thuật chi tiết. Chúng tôi sẽ liên hệ khách hàng trước khi tiến hành sửa chữa nếu chi phí thực tế phát sinh khác biệt.</p>
    ` : `
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">CHI PHÍ SỬA CHỮA THỰC TẾ</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:8px;width:35%;font-weight:bold;">Chi phí thực tế</td><td style="border:1px solid #111;padding:8px;font-weight:bold;font-size:${baseFont + 2}px;">${ticket.actualCost > 0 ? `${ticket.actualCost.toLocaleString("vi-VN")}đ` : "Chưa xác định"}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">VAT (${escapeHtml(vatLabel)})</td><td style="border:1px solid #111;padding:6px 8px;">${ticket.actualCost > 0 && ticket.vat !== "KCT" ? `${Math.round(ticket.actualCost * vp / (100 + vp)).toLocaleString("vi-VN")}đ` : "—"}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;opacity:0.6;">Chi phí dự kiến ban đầu (tham khảo)</td><td style="border:1px solid #111;padding:6px 8px;opacity:0.6;">${ticket.estimatedCost > 0 ? `${ticket.estimatedCost.toLocaleString("vi-VN")}đ` : "—"}</td></tr>
      </tbody>
    </table>
    `}

    ${ticket.note ? `<div style="margin-bottom:20px;"><b>Ghi chú: </b>${escapeHtml(ticket.note)}</div>` : ""}

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Khách hàng</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên — đồng ý mức chi phí ${isReceived ? "dự kiến" : "thực tế"} trên)</div>
      </div>
      <div>
        <b>Người lập phiếu</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${isReceived ? "Tiep-nhan" : "Bao-gia-thuc-te"}-${escapeHtml(ticket.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

/* ---------------- Phiếu dịch vụ IT Helpdesk ---------------- */
function buildHelpdeskTicketHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const statusLabel = HELPDESK_STATUSES.find((s) => s.id === ticket.status)?.label || "";

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.code)}</div>
        <div>Ngày lập: ${new Date(ticket.createdAt).toLocaleDateString("vi-VN")}</div>
        <div>Trạng thái: ${escapeHtml(statusLabel)}</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:8px 0 16px;">PHIẾU DỊCH VỤ IT HELPDESK</h1>

    <div style="display:flex;gap:24px;margin-bottom:12px;">
      <div style="flex:1;">
        <div><b>Khách hàng / Đơn vị: </b>${escapeHtml(ticket.customerName || "—")}</div>
        <div><b>Điện thoại: </b>${escapeHtml(ticket.customerPhone || "—")}</div>
        <div><b>Địa chỉ: </b>${escapeHtml(ticket.customerAddress || "—")}</div>
      </div>
      <div style="flex:1;">
        <div><b>Người lập phiếu: </b>${escapeHtml(ticket.createdBy)}</div>
        <div><b>Người xử lý: </b>${escapeHtml(ticket.assignee || "—")}</div>
        <div><b>Ngày tiếp nhận: </b>${ticket.receivedDate ? new Date(ticket.receivedDate).toLocaleDateString("vi-VN") : "—"}</div>
        ${ticket.completedDate ? `<div><b>Ngày hoàn thành: </b>${new Date(ticket.completedDate).toLocaleDateString("vi-VN")}</div>` : ""}
      </div>
    </div>

    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">NỘI DUNG YÊU CẦU</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Loại yêu cầu</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.requestType)}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Mô tả yêu cầu / sự cố</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.description || "—")}</td></tr>
      </tbody>
    </table>

    ${ticket.solution ? `
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">GIẢI PHÁP / KẾT QUẢ XỬ LÝ</th></tr></thead>
      <tbody><tr><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.solution)}</td></tr></tbody>
    </table>` : ""}

    ${ticket.note ? `<div style="margin-bottom:20px;"><b>Ghi chú: </b>${escapeHtml(ticket.note)}</div>` : ""}

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Khách hàng / Đơn vị</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
      <div>
        <b>Người xử lý</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(ticket.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}

/* ---------------- Vận chuyển — Phiếu/nhãn vận đơn ---------------- */
function buildShippingLabelHTML(ticket, paperSize) {
    const pageCss = paperSize === "A5"
        ? `@page { size: A5; margin: 8mm; } body{ width:132mm; }`
        : `@page { size: A4; margin: 12mm; } body{ width:186mm; }`;
    const baseFont = paperSize === "A5" ? 11.5 : 13;
    const statusLabel = SHIPPING_STATUSES.find((s) => s.id === ticket.status)?.label || "";

    const body = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
      <div>
        <div style="font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(COMPANY_INFO.name)}</div>
        <div style="font-size:${baseFont - 1}px;">${escapeHtml(COMPANY_INFO.address)}</div>
        <div style="font-size:${baseFont - 1}px;">ĐT: ${escapeHtml(COMPANY_INFO.phone)}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-weight:bold;">Mã phiếu: ${escapeHtml(ticket.code)}</div>
        <div>Đơn hàng: ${escapeHtml(ticket.orderCode || "—")}</div>
        <div>Trạng thái: ${escapeHtml(statusLabel)}</div>
      </div>
    </div>
    <h1 style="text-align:center;font-size:${paperSize === "A5" ? 18 : 22}px;letter-spacing:1px;margin:8px 0 16px;">PHIẾU VẬN CHUYỂN</h1>

    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">NGƯỜI NHẬN</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Họ tên</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.recipientName || "—")}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Điện thoại</td><td style="border:1px solid #111;padding:6px 8px;font-weight:bold;font-size:${baseFont + 2}px;">${escapeHtml(ticket.recipientPhone || "—")}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Địa chỉ</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.recipientAddress || "—")}</td></tr>
      </tbody>
    </table>

    <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
      <thead><tr><th colspan="2" style="border:1px solid #111;padding:6px 8px;font-size:12px;background:#f2f2f2;">THÔNG TIN VẬN CHUYỂN</th></tr></thead>
      <tbody>
        <tr><td style="border:1px solid #111;padding:6px 8px;width:35%;">Đơn vị vận chuyển</td><td style="border:1px solid #111;padding:6px 8px;">${escapeHtml(ticket.carrier || "—")}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Mã vận đơn</td><td style="border:1px solid #111;padding:6px 8px;font-weight:bold;">${escapeHtml(ticket.trackingCode || "—")}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Tiền thu hộ (COD)</td><td style="border:1px solid #111;padding:6px 8px;font-weight:bold;font-size:${baseFont + 2}px;">${ticket.codAmount > 0 ? `${ticket.codAmount.toLocaleString("vi-VN")}đ` : "Không thu hộ"}</td></tr>
        <tr><td style="border:1px solid #111;padding:6px 8px;">Phí giao hàng</td><td style="border:1px solid #111;padding:6px 8px;">${ticket.shippingFee > 0 ? `${ticket.shippingFee.toLocaleString("vi-VN")}đ` : "—"}</td></tr>
      </tbody>
    </table>

    ${ticket.note ? `<div style="margin-bottom:20px;padding:10px;border:1px dashed #111;"><b>Ghi chú giao hàng: </b>${escapeHtml(ticket.note)}</div>` : ""}

    <div style="display:flex;justify-content:space-between;margin-top:40px;text-align:center;">
      <div>
        <b>Người giao</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
      <div>
        <b>Người nhận</b>
        <div style="font-size:11px;opacity:0.7;">(Ký, ghi rõ họ tên)</div>
      </div>
    </div>
  `;

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(ticket.code)}</title>
    <style>
      * { box-sizing: border-box; }
      body { font-family: 'Times New Roman', Times, serif; color:#111; font-size:${baseFont}px; margin:0 auto; }
      ${pageCss}
    </style>
  </head><body>${body}</body></html>`;
}




const QUOTE_STATUS_LABEL = { active: "Còn hiệu lực", converted: "Đã chuyển đơn", cancelled: "Đã huỷ", expired: "Hết hiệu lực" };

function QuoteItemsTable({ items, products, onUpdate, onRemove }) {
    return (
        <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
            <table className="w-full text-sm table-fixed" style={{ minWidth: 480 }}>
                <colgroup><col style={{ width: 28 }} /><col /><col style={{ width: 70 }} /><col style={{ width: 130 }} /><col style={{ width: 120 }} /><col style={{ width: 32 }} /></colgroup>
                <thead>
                <tr style={{ borderBottom: `2px solid ${INK}`, background: PAPER }}>
                    <th className="text-left px-2 py-2.5 text-[11px] uppercase tracking-wider opacity-50">STT</th>
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
                    return (
                        <tr key={it.productId} style={{ borderBottom: `1px dashed ${LINE}` }}>
                            <td className="px-2 py-3 text-xs opacity-40 align-top">{idx + 1}</td>
                            <td className="px-2 py-3 align-top">
                                <div className="font-semibold text-base leading-snug" style={{ color: INK }}>{p?.name}</div>
                                <div className="text-xs opacity-50 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p?.sku || p?.code}</div>
                            </td>
                            <td className="px-1 py-3 align-top">
                                <input type="number" min={1} value={it.qty || ""} onChange={(e) => { const v = e.target.value; onUpdate(it.productId, { qty: v === "" ? "" : Math.max(0, Number(v)) }); }}
                                       onBlur={() => { if (!it.qty || it.qty < 1) onUpdate(it.productId, { qty: 1 }); }}
                                       className="w-full border rounded-sm py-2 px-1 text-center text-[15px] font-medium" style={{ borderColor: LINE }} />
                            </td>
                            <td className="px-1 py-3 align-top">
                                <MoneyInput value={it.price} onChange={(v) => onUpdate(it.productId, { price: v === "" ? 0 : v })}
                                            className="w-full border rounded-sm py-2 px-2 text-right text-[15px] font-medium" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                            </td>
                            <td className="px-2 py-3 text-right font-bold text-base whitespace-nowrap align-top" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(it.qty * it.price)}</td>
                            <td className="px-1 py-3 align-top"><button onClick={() => onRemove(it.productId)} style={{ color: RUST }}><X size={14} /></button></td>
                        </tr>
                    );
                })}
                {items.length === 0 && <tr><td colSpan={6} className="text-center py-8 opacity-40">Chưa có thông tin sản phẩm.</td></tr>}
                </tbody>
            </table>
        </div>
    );
}

function Quotations({ quotations, setQuotations, orders, setOrders, products, setProducts, customers, setCustomers, employeeNames, currentUser, addLog, goToDoc, brands }) {
    const isAdmin = currentUser.role === "admin";
    const [creating, setCreating] = useState(false);
    const [editingQuoteId, setEditingQuoteId] = useState(null);
    const [form, setForm] = useState({});
    const [query, setQuery] = useState("");
    const [viewingId, setViewingId] = useState(null);
    const [printPaperSize, setPrintPaperSize] = useState("A4");
    const [printModalOpen, setPrintModalOpen] = useState(false);
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);
    const [reactivateDays, setReactivateDays] = useState(7);

    // Báo giá hết hiệu lực sẽ tự động chuyển sang trạng thái "Hết hiệu lực" (coi như bị huỷ).
    // Chỉ quản trị viên mới có thể kích hoạt lại hoặc sửa đổi báo giá đã hết hiệu lực.
    useEffect(() => {
        const toExpire = quotations.filter((q) => q.status === "active" && quoteExpiryInfo(q).expired);
        if (toExpire.length === 0) return;
        setQuotations((prev) => prev.map((q) => (toExpire.some((e) => e.id === q.id) ? { ...q, status: "expired" } : q)));
        toExpire.forEach((q) => addLog("Báo giá tự động hết hiệu lực", q.code));
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [quotations]);

    const expiryFromDays = (days) => { const d = new Date(); d.setDate(d.getDate() + (Number(days) || 7)); return d.toISOString(); };

    const openNew = () => {
        setForm({
            customerId: "", customerName: "", customerPhone: "", branch: BRANCHES[0], seller: currentUser.fullName,
            items: [], vat: "VAT10", orderDiscount: 0, discountType: "amount", shippingFee: 0, notes: "", tags: [], validDays: 7,
        });
        setEditingQuoteId(null);
        setCreating(true);
    };
    // Sửa đổi báo giá — báo giá còn hiệu lực thì mọi vai trò đều sửa được; báo giá đã hết hiệu lực thì chỉ quản trị viên.
    // Mở lại form với dữ liệu hiện có, khi lưu sẽ tính lại hạn hiệu lực và kích hoạt lại nếu đang hết hạn.
    const openEdit = (q) => {
        const days = Math.max(1, Math.round((new Date(q.expiryDate).getTime() - new Date(q.createdAt).getTime()) / 86400000)) || 7;
        setForm({
            customerId: q.customerId, customerName: q.customerName, customerPhone: q.customerPhone, branch: q.branch, seller: q.seller,
            items: q.items.map((it) => ({ ...it })), vat: q.vat, orderDiscount: q.orderDiscount, discountType: q.discountType,
            shippingFee: q.shippingFee, notes: q.notes, tags: q.tags || [], validDays: days,
        });
        setEditingQuoteId(q.id);
        setViewingId(null);
        setCreating(true);
    };
    const addItem = (productId, productOverride) => {
        if (!productId) return;
        setForm((f) => {
            if (f.items.some((it) => it.productId === productId)) return f;
            const p = productOverride || products.find((x) => x.id === productId);
            if (!p) return f;
            // Nếu đây là sản phẩm đầu tiên được thêm vào báo giá, tự đặt VAT của báo giá theo VAT sản phẩm (vẫn cho đổi lại sau).
            const vat = f.items.length === 0 && p?.vat ? p.vat : f.vat;
            return { ...f, items: [...f.items, { productId, qty: 1, price: p.retailPrice }], vat };
        });
    };
    // Tạo nhanh sản phẩm mới ngay trong lúc lập báo giá (khi chưa kịp tạo/nhập kho sản phẩm) — trả về product vừa tạo để thêm luôn vào báo giá.
    const quickCreateProduct = (data) => {
        const code = data.code.trim();
        if (products.some((p) => p.code.toLowerCase() === code.toLowerCase())) {
            alert(`Mã vật tư "${code}" đã tồn tại — vui lòng dùng mã khác.`);
            return null;
        }
        const newProduct = normalizeProduct({
            id: uid(), code, name: data.name.trim(), unit: data.unit || UNITS[0], brand: data.brand || "",
            warrantyMonths: Number(data.warrantyMonths) || 0, vat: data.vat || "VAT10", sku: nextSKU(products),
        });
        setProducts((prev) => [...prev, newProduct]);
        addLog("Tạo nhanh sản phẩm (từ báo giá)", `${newProduct.code} · ${newProduct.name}`);
        return newProduct;
    };
    const updateItem = (productId, patch) => setForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
    const removeItem = (productId) => setForm((f) => ({ ...f, items: f.items.filter((it) => it.productId !== productId) }));
    const calc = quoteCalc({ items: form.items || [], vat: form.vat, orderDiscount: form.orderDiscount, discountType: form.discountType, shippingFee: form.shippingFee });

    const submit = () => {
        if (!form.items || form.items.length === 0) { alert("Vui lòng thêm ít nhất một sản phẩm."); return; }
        const cust = customers.find((c) => c.id === form.customerId);
        const custAddress = cust ? [cust.addressDetail, cust.ward, cust.province].filter(Boolean).join(", ") : "";
        if (editingQuoteId) {
            const old = quotations.find((q) => q.id === editingQuoteId);
            setQuotations((prev) => prev.map((q) => {
                if (q.id !== editingQuoteId) return q;
                return normalizeQuote({
                    ...q, customerId: form.customerId, customerName: cust ? cust.name : form.customerName, customerPhone: cust ? cust.phone : form.customerPhone,
                    customerAddress: custAddress, customerTaxCode: cust ? (cust.taxCode || "") : "",
                    branch: form.branch, seller: form.seller, items: form.items, vat: form.vat, orderDiscount: form.orderDiscount, discountType: form.discountType,
                    shippingFee: form.shippingFee, notes: form.notes, tags: form.tags, expiryDate: expiryFromDays(form.validDays), status: "active",
                });
            }));
            addLog("Sửa báo giá", old ? old.code : editingQuoteId);
        } else {
            const newQuote = normalizeQuote({
                code: nextQuoteCode(quotations), customerId: form.customerId, customerName: cust ? cust.name : form.customerName, customerPhone: cust ? cust.phone : form.customerPhone,
                customerAddress: custAddress, customerTaxCode: cust ? (cust.taxCode || "") : "",
                branch: form.branch, seller: form.seller, items: form.items, vat: form.vat, orderDiscount: form.orderDiscount, discountType: form.discountType,
                shippingFee: form.shippingFee, notes: form.notes, tags: form.tags, expiryDate: expiryFromDays(form.validDays),
            });
            setQuotations((prev) => [newQuote, ...prev]);
            addLog("Tạo báo giá", `${newQuote.code} · ${vnd(quoteCalc(newQuote).total)}`);
        }
        setCreating(false);
        setEditingQuoteId(null);
    };

    const filtered = quotations.filter((q) => {
        if (!query.trim()) return true;
        const s = query.trim().toLowerCase();
        return q.code.toLowerCase().includes(s) || (q.customerName || "").toLowerCase().includes(s);
    });
    const viewingQuote = quotations.find((q) => q.id === viewingId) || null;

    // Chuyển báo giá thành đơn hàng thật: tự tạo khách hàng mới nếu báo giá chưa gắn khách có sẵn.
    const convertToOrder = (q) => {
        let customerId = q.customerId;
        if (!customerId && q.customerName) {
            const newCust = { id: uid(), code: nextCustomerCode(customers), name: q.customerName, phone: q.customerPhone || "", email: "", taxCode: "", province: "", ward: "", addressDetail: "", note: `Tạo tự động khi chuyển báo giá ${q.code}`, group: "retail" };
            setCustomers((prev) => [...prev, newCust]);
            customerId = newCust.id;
        }
        const code = nextOrderCode(orders);
        const now = new Date().toISOString();
        const itemsWithFulfilled = q.items.map((it) => {
            const p = products.find((x) => x.id === it.productId);
            const fulfilled = !p?.hasSeries; // sản phẩm series cần chọn số series thủ công sau khi chuyển — sẽ ở trạng thái chờ hàng
            return { ...it, series: [], fulfilled };
        });
        const newOrder = {
            id: uid(), code, createdAt: now, date: todayISO(), customerId, channel: "store", branch: q.branch, seller: q.seller, deliveryDate: "",
            tags: q.tags, notes: `Chuyển từ báo giá ${q.code}.${q.notes ? " " + q.notes : ""}`, status: "pending", items: itemsWithFulfilled, vat: q.vat,
            shippingAt: null, deliveredAt: null, paidCompleteAt: null, cancelledAt: null,
            orderDiscount: q.orderDiscount, discountType: q.discountType, creditDays: 0, shippingFee: q.shippingFee, paidAmount: 0,
            approvalStatus: "approved", approvalReason: "", createdByRole: currentUser.role,
        };
        setOrders((prev) => [newOrder, ...prev]);
        setProducts((prev) => prev.map((p) => {
            const it = itemsWithFulfilled.find((i) => i.productId === p.id);
            if (!it || !it.fulfilled) return p;
            return { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: code, date: todayISO(), qty: it.qty, price: it.price, series: [] }] };
        }));
        setQuotations((prev) => prev.map((x) => (x.id === q.id ? { ...x, status: "converted", convertedOrderId: newOrder.id } : x)));
        addLog("Chuyển báo giá thành đơn hàng", `${q.code} → ${code}`);
        setViewingId(null);
    };

    const cancelQuote = (q) => {
        setQuotations((prev) => prev.map((x) => (x.id === q.id ? { ...x, status: "cancelled" } : x)));
        addLog("Huỷ báo giá", q.code);
    };
    const removeQuote = (id) => setQuotations((prev) => prev.filter((q) => q.id !== id));
    // Kích hoạt lại báo giá đã hết hiệu lực — chỉ quản trị viên. Cấp thêm số ngày hiệu lực tính từ thời điểm kích hoạt.
    const reactivateQuote = (q, days) => {
        setQuotations((prev) => prev.map((x) => (x.id === q.id ? { ...x, status: "active", expiryDate: expiryFromDays(days) } : x)));
        addLog("Kích hoạt lại báo giá", `${q.code} · thêm ${days} ngày`);
    };

    const executePrint = () => {
        const html = buildQuoteHTML(viewingQuote, products, printPaperSize);
        const result = printHTML(html);
        addLog("In báo giá", `${viewingQuote.code} · ${PAPER_SIZES.find((p) => p.id === printPaperSize)?.label}`);
        if (result.ok) { setPrintModalOpen(false); setPrintBlockedUrl(null); } else { setPrintBlockedUrl(result.url); }
    };

    return (
        <div>
            <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
                <div className="relative flex-1 max-w-xs">
                    <Search size={15} className="absolute left-2 top-1/2 -translate-y-1/2 opacity-50" />
                    <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm theo mã hoặc tên khách…"
                           className="w-full pl-7 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
                </div>
                <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}>
                    <Plus size={15} /> Tạo báo giá
                </button>
            </div>

            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 920 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        {["Mã BG", "Ngày tạo", "Khách hàng", "Hạn hiệu lực", "Tổng tiền", "Trạng thái", "Đơn hàng", ""].map((h, hi) => (
                            <th key={hi} className="text-left px-4 py-3 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
                        ))}
                    </tr>
                    </thead>
                    <tbody>
                    {filtered.map((q) => {
                        const c = quoteCalc(q);
                        const exp = quoteExpiryInfo(q);
                        const effectiveStatus = q.status === "active" && exp.expired ? "expired" : q.status;
                        const convertedOrder = q.convertedOrderId ? orders.find((o) => o.id === q.convertedOrderId) : null;
                        return (
                            <tr key={q.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                                <td className="px-4 py-3.5 font-medium whitespace-nowrap"><button onClick={() => setViewingId(q.id)} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{q.code}</button></td>
                                <td className="px-4 py-3.5 whitespace-nowrap opacity-80" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(q.createdAt)}</td>
                                <td className="px-4 py-3.5 whitespace-nowrap">{q.customerName || "—"}</td>
                                <td className="px-4 py-3.5 whitespace-nowrap" style={{ color: exp.expired ? RUST : INK }}>{new Date(q.expiryDate).toLocaleDateString("vi-VN")}</td>
                                <td className="px-4 py-3.5 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(c.total)}</td>
                                <td className="px-4 py-3.5 whitespace-nowrap">
                    <span className="text-[11px] px-2 py-0.5 rounded-full" style={{
                        background: effectiveStatus === "converted" ? `${FOREST}1A` : effectiveStatus === "cancelled" ? "rgba(0,0,0,0.08)" : effectiveStatus === "expired" ? `${RUST}1A` : `${BLUE}1A`,
                        color: effectiveStatus === "converted" ? FOREST : effectiveStatus === "cancelled" ? INK : effectiveStatus === "expired" ? RUST : BLUE,
                    }}>
                      {QUOTE_STATUS_LABEL[effectiveStatus] || QUOTE_STATUS_LABEL[q.status]}
                    </span>
                                </td>
                                <td className="px-4 py-3.5 whitespace-nowrap">
                                    {convertedOrder ? (
                                        <button onClick={() => goToDoc && goToDoc(convertedOrder.code)} className="hover:underline font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>{convertedOrder.code}</button>
                                    ) : <span className="opacity-30">—</span>}
                                </td>
                                <td className="px-4 py-3.5 text-right">
                                    {q.status === "active" && !exp.expired && (
                                        <button onClick={() => convertToOrder(q)} className="text-xs px-2.5 py-1 rounded-sm text-white whitespace-nowrap" style={{ background: FOREST }}>Chuyển đơn hàng</button>
                                    )}
                                </td>
                            </tr>
                        );
                    })}
                    {filtered.length === 0 && <tr><td colSpan={8} className="text-center py-10 opacity-50">Chưa có báo giá nào.</td></tr>}
                    </tbody>
                </table>
            </div>

            {/* Modal tạo / sửa báo giá */}
            {creating && (
                <Modal title={editingQuoteId ? "Sửa báo giá" : "Tạo phiếu báo giá"} onClose={() => { setCreating(false); setEditingQuoteId(null); }} size="2xl">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Field label="Khách hàng có sẵn (không bắt buộc)">
                            <FilterSearchSelect options={customers.map((c) => ({ id: c.id, label: c.name }))} value={form.customerId} onChange={(v) => setForm({ ...form, customerId: v })} placeholder="Gõ tên khách hàng…" />
                        </Field>
                        <Field label="Bán bởi">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.seller} onChange={(e) => setForm({ ...form, seller: e.target.value })}>
                                {(employeeNames.length ? employeeNames : EMPLOYEES).map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                            </select>
                        </Field>
                    </div>
                    {!form.customerId && (
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <Field label="Tên khách hàng (nếu chưa có trong danh sách)"><input className={inputCls} style={{ borderColor: LINE }} value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} /></Field>
                            <Field label="SĐT"><input className={inputCls} style={{ borderColor: LINE }} inputMode="numeric" value={form.customerPhone} onChange={(e) => setForm({ ...form, customerPhone: e.target.value.replace(/\D/g, "") })} /></Field>
                        </div>
                    )}
                    <p className="text-[11px] opacity-50 -mt-2 mb-3">Khách hàng chỉ được thêm chính thức vào danh sách khi báo giá được chuyển thành đơn hàng.</p>

                    <Field label="Thêm sản phẩm vào báo giá">
                        <ProductPicker products={products} onPick={addItem} onQuickCreate={quickCreateProduct} brands={brands} />
                    </Field>
                    <QuoteItemsTable items={form.items || []} products={products} onUpdate={updateItem} onRemove={removeItem} />

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4">
                        <Field label="Hiệu lực báo giá (số ngày)" hint={editingQuoteId ? "Tính từ thời điểm lưu chỉnh sửa." : "Tính từ thời điểm tạo báo giá."}>
                            <input type="number" min={1} value={form.validDays ?? 7}
                                   onChange={(e) => setForm({ ...form, validDays: e.target.value === "" ? "" : Math.max(1, Number(e.target.value)) })}
                                   onBlur={() => { if (!form.validDays) setForm({ ...form, validDays: 7 }); }}
                                   className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                        </Field>
                        <Field label="Chiết khấu">
                            <div className="flex items-center gap-1.5">
                                <div className="flex rounded-sm border overflow-hidden" style={{ borderColor: LINE }}>
                                    <button type="button" onClick={() => setForm({ ...form, discountType: "amount" })} className="px-2 py-1.5 text-xs" style={{ background: (form.discountType || "amount") === "amount" ? INK : "transparent", color: (form.discountType || "amount") === "amount" ? "#fff" : INK }}>đ</button>
                                    <button type="button" onClick={() => setForm({ ...form, discountType: "percent" })} className="px-2 py-1.5 text-xs" style={{ background: form.discountType === "percent" ? INK : "transparent", color: form.discountType === "percent" ? "#fff" : INK }}>%</button>
                                </div>
                                <MoneyInput value={form.orderDiscount} onChange={(v) => setForm({ ...form, orderDiscount: v })} className="flex-1 border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                            </div>
                        </Field>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <Field label="Phí giao hàng"><MoneyInput value={form.shippingFee} onChange={(v) => setForm({ ...form, shippingFee: v })} className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} /></Field>
                    </div>

                    <div className="p-3 rounded-sm space-y-1.5 mb-3" style={{ background: PAPER }}>
                        <div className="flex justify-between text-sm"><span className="opacity-60">Tạm tính</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.subtotal)}</span></div>
                        <div className="flex justify-between text-sm font-semibold" style={{ color: INK }}><span>Tổng cộng</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.total)}</span></div>
                    </div>

                    <TagsNotesCompact tags={form.tags} setTags={(arr) => setForm({ ...form, tags: arr })} notes={form.notes} setNotes={(v) => setForm({ ...form, notes: v })} />
                    <button onClick={submit} className="w-full py-2.5 rounded-sm text-white text-sm mt-4" style={{ background: INK }}>
                        {editingQuoteId ? "Lưu thay đổi (kích hoạt lại)" : `Tạo báo giá (hiệu lực ${form.validDays || 7} ngày)`}
                    </button>
                </Modal>
            )}

            {/* Modal chi tiết báo giá */}
            {viewingQuote && (() => {
                const c = quoteCalc(viewingQuote);
                const exp = quoteExpiryInfo(viewingQuote);
                const isExpired = viewingQuote.status === "expired" || (viewingQuote.status === "active" && exp.expired);
                return (
                    <Modal title={`Báo giá ${viewingQuote.code}`} onClose={() => setViewingId(null)} size="xl">
                        <div className="flex items-center gap-2 mb-4 flex-wrap">
              <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: isExpired ? `${RUST}1A` : `${BLUE}1A`, color: isExpired ? RUST : BLUE }}>
                Hiệu lực đến {new Date(viewingQuote.expiryDate).toLocaleDateString("vi-VN")}{isExpired ? " — đã hết hiệu lực" : ` (còn ${exp.daysLeft} ngày)`}
              </span>
                            {viewingQuote.status === "converted" && <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: `${FOREST}1A`, color: FOREST }}>Đã chuyển đơn hàng</span>}
                            {viewingQuote.status === "cancelled" && <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: "rgba(0,0,0,0.08)", color: INK }}>Đã huỷ</span>}
                        </div>
                        <p className="text-sm mb-1"><b>Khách hàng:</b> {viewingQuote.customerName || "—"}{viewingQuote.customerPhone ? ` · ${viewingQuote.customerPhone}` : ""}</p>
                        {viewingQuote.customerAddress && <p className="text-sm mb-1"><b>Địa chỉ:</b> {viewingQuote.customerAddress}</p>}
                        {viewingQuote.customerTaxCode && <p className="text-sm mb-1"><b>MST:</b> {viewingQuote.customerTaxCode}</p>}
                        <p className="text-sm mb-4"><b>Bán bởi:</b> {viewingQuote.seller}</p>
                        <QuoteItemsTable items={viewingQuote.items} products={products} onUpdate={() => {}} onRemove={() => {}} />
                        <div className="p-3 rounded-sm space-y-1.5 my-4" style={{ background: PAPER }}>
                            <div className="flex justify-between text-sm"><span className="opacity-60">Tạm tính</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.subtotal)}</span></div>
                            {viewingQuote.orderDiscount > 0 && <div className="flex justify-between text-sm"><span className="opacity-60">Chiết khấu</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.discountAmount)}</span></div>}
                            <div className="flex justify-between text-sm font-semibold" style={{ color: INK }}><span>Tổng cộng</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.total)}</span></div>
                        </div>

                        {/* Báo giá hết hiệu lực: chỉ quản trị viên được kích hoạt lại hoặc sửa đổi */}
                        {isExpired && viewingQuote.status !== "converted" && (
                            <div className="p-3 rounded-sm mb-4" style={{ background: `${RUST}0D`, border: `1px solid ${RUST}33` }}>
                                <p className="text-sm mb-2" style={{ color: RUST }}>Báo giá này đã hết hiệu lực và tự động bị huỷ.</p>
                                {isAdmin ? (
                                    <div className="flex items-center gap-2 flex-wrap">
                                        <span className="text-xs opacity-70">Kích hoạt lại thêm</span>
                                        <input type="number" min={1} value={reactivateDays} onChange={(e) => setReactivateDays(Math.max(1, Number(e.target.value) || 1))}
                                               className="w-16 border rounded-sm py-1 px-2 text-sm text-center" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                        <span className="text-xs opacity-70">ngày</span>
                                        <button onClick={() => reactivateQuote(viewingQuote, reactivateDays)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm text-white" style={{ background: FOREST }}>
                                            <RotateCcw size={13} /> Kích hoạt lại
                                        </button>
                                        <button onClick={() => openEdit(viewingQuote)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                                            <Pencil size={13} /> Sửa báo giá
                                        </button>
                                    </div>
                                ) : (
                                    <p className="text-xs opacity-60">Chỉ quản trị viên mới có thể kích hoạt lại hoặc sửa đổi báo giá đã hết hiệu lực.</p>
                                )}
                            </div>
                        )}

                        <div className="flex gap-2 flex-wrap">
                            <button onClick={() => { setPrintPaperSize("A4"); setPrintBlockedUrl(null); setPrintModalOpen(true); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                                <Printer size={13} /> In báo giá
                            </button>
                            {viewingQuote.status === "active" && !exp.expired && (
                                <button onClick={() => convertToOrder(viewingQuote)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm text-white" style={{ background: FOREST }}>
                                    <Check size={13} /> Chuyển thành đơn hàng
                                </button>
                            )}
                            {viewingQuote.status === "active" && (
                                <button onClick={() => openEdit(viewingQuote)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                                    <Pencil size={13} /> Sửa báo giá
                                </button>
                            )}
                            {viewingQuote.status === "active" && (
                                <button onClick={() => cancelQuote(viewingQuote)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: RUST, color: RUST }}>
                                    <XCircle size={13} /> Huỷ báo giá
                                </button>
                            )}
                            {isAdmin && (
                                <button onClick={() => { removeQuote(viewingQuote.id); setViewingId(null); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border ml-auto" style={{ borderColor: RUST, color: RUST }}>
                                    <Trash2 size={13} /> Xoá
                                </button>
                            )}
                        </div>

                        {printModalOpen && (
                            <div className="mt-4 p-4 rounded-sm" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                <Field label="Khổ giấy">
                                    <div className="grid grid-cols-2 gap-2">
                                        {PAPER_SIZES.map((p) => (
                                            <button key={p.id} type="button" onClick={() => setPrintPaperSize(p.id)}
                                                    className="px-3 py-2 rounded-sm text-sm border text-center"
                                                    style={{ borderColor: printPaperSize === p.id ? INK : LINE, background: printPaperSize === p.id ? INK : "transparent", color: printPaperSize === p.id ? "#fff" : INK }}>
                                                {p.label}
                                            </button>
                                        ))}
                                    </div>
                                </Field>
                                <button onClick={executePrint} className="w-full py-2.5 rounded-sm text-white text-sm flex items-center justify-center gap-1.5" style={{ background: INK }}>
                                    <Printer size={14} /> In ngay
                                </button>
                                {printBlockedUrl && (
                                    <div className="mt-3 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                                        <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới:</p>
                                        <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở phiếu báo giá để in</a>
                                    </div>
                                )}
                            </div>
                        )}
                    </Modal>
                );
            })()}
        </div>
    );
}

function Orders({ orders, setOrders, products, setProducts, customers, setCustomers, employeeNames, currentUser, addLog, focusOrderId, initialFilterStatus, onFocusHandled, printSettings, setPrintSettings }) {
    const isCtv = currentUser.role === "ctv";
    const isAdmin = currentUser.role === "admin";
    const [creating, setCreating] = useState(false);
    const [editingOrderId, setEditingOrderId] = useState(null);
    const [form, setForm] = useState({});
    const [pickingAddress, setPickingAddress] = useState(false);
    const [filterStatus, setFilterStatus] = useState("all");
    const [searchQuery, setSearchQuery] = useState("");
    const [showFilters, setShowFilters] = useState(false);
    const [filterSeller, setFilterSeller] = useState("");
    const [filterFrom, setFilterFrom] = useState("");
    const [filterTo, setFilterTo] = useState("");
    const [filterInvoice, setFilterInvoice] = useState("");
    const [filterText, setFilterText] = useState("");
    const [filterCustomer, setFilterCustomer] = useState("");
    const [filterProduct, setFilterProduct] = useState("");
    const [filterApprovalOnly, setFilterApprovalOnly] = useState(false);
    const [filterCancelledOnly, setFilterCancelledOnly] = useState(false);
    const [selectedIds, setSelectedIds] = useState(new Set());
    const [viewingId, setViewingId] = useState(null);
    const [payInput, setPayInput] = useState("");
    const [refundInput, setRefundInput] = useState("");
    const [pendingFulfillDraft, setPendingFulfillDraft] = useState({});
    const [returning, setReturning] = useState(false);
    const [returnForm, setReturnForm] = useState(null);
    const [viewingReturnId, setViewingReturnId] = useState(null);
    const [printModalOpen, setPrintModalOpen] = useState(false);
    const [printDocType, setPrintDocType] = useState("invoice"); // invoice | handover (BBBG)
    const [printPaperSize, setPrintPaperSize] = useState(printSettings?.defaultPaperSize || "A4");
    const [showPrintOptions, setShowPrintOptions] = useState(false);
    const [printBlockedUrl, setPrintBlockedUrl] = useState(null);

    const viewingOrder = orders.find((o) => o.id === viewingId) || null;
    const executePrint = () => {
        const cust = customers.find((c) => c.id === viewingOrder.customerId);
        if (printDocType === "handover") {
            const html = buildHandoverHTML(viewingOrder, products, cust, printPaperSize);
            const result = printHTML(html);
            addLog("In biên bản bàn giao (BBBG)", `${viewingOrder.code} · ${PAPER_SIZES.find((p) => p.id === printPaperSize)?.label}`);
            if (result.ok) { setPrintModalOpen(false); setPrintBlockedUrl(null); } else { setPrintBlockedUrl(result.url); }
            return;
        }
        const html = buildInvoiceHTML(viewingOrder, products, cust, printPaperSize, printSettings);
        const result = printHTML(html);
        setPrintSettings((s) => ({ ...s, defaultPaperSize: printPaperSize }));
        addLog("In đơn hàng", `${viewingOrder.code} · ${PAPER_SIZES.find((p) => p.id === printPaperSize)?.label}`);
        if (result.ok) {
            setPrintModalOpen(false);
            setPrintBlockedUrl(null);
        } else {
            setPrintBlockedUrl(result.url);
        }
    };

    useEffect(() => {
        if (focusOrderId) {
            setViewingId(focusOrderId);
            setPayInput("");
            onFocusHandled && onFocusHandled();
        }
    }, [focusOrderId]);

    useEffect(() => {
        if (initialFilterStatus) {
            setFilterStatus(initialFilterStatus);
            setSelectedIds(new Set());
            onFocusHandled && onFocusHandled();
        }
    }, [initialFilterStatus]);

    const openNew = () => {
        setForm({
            customerId: "", channel: "store", branch: BRANCHES[0], seller: (currentUser.fullName || employeeNames[0] || EMPLOYEES[0]), deliveryDate: "",
            shippingAddress: null,
            tags: [], notes: "", items: [], vat: "VAT10", orderDiscount: 0, discountType: "amount", creditDays: 0, shippingFee: 0, paidAmount: 0,
        });
        setEditingOrderId(null);
        setCreating(true);
    };

    const addItem = (productId) => {
        if (!productId) return;
        setForm((f) => {
            if (f.items.some((it) => it.productId === productId)) return f;
            const p = products.find((x) => x.id === productId);
            // Nếu đây là sản phẩm đầu tiên được thêm vào đơn, tự đặt VAT của đơn theo VAT sản phẩm (vẫn cho đổi lại sau).
            const vat = f.items.length === 0 && p?.vat ? p.vat : f.vat;
            return { ...f, items: [...f.items, { productId, qty: 1, price: p.retailPrice, series: [] }], vat };
        });
    };
    const updateItem = (productId, patch) => setForm((f) => ({ ...f, items: f.items.map((it) => (it.productId === productId ? { ...it, ...patch } : it)) }));
    const removeItem = (productId) => setForm((f) => ({ ...f, items: f.items.filter((it) => it.productId !== productId) }));

    const calc = orderCalc({ items: form.items || [], vat: form.vat, orderDiscount: form.orderDiscount, discountType: form.discountType, shippingFee: form.shippingFee, paidAmount: form.paidAmount });

    // Không cho xuất dưới giá sàn theo vai trò (kể cả sau chiết khấu) — trừ Admin. Nhân viên/CTV vẫn tạo được nhưng đơn cần Admin duyệt.
    const discountRate = (() => {
        const subtotal = (form.items || []).reduce((s, it) => s + it.qty * it.price, 0);
        if (form.discountType === "percent") return (Number(form.orderDiscount) || 0) / 100;
        return subtotal > 0 ? (Number(form.orderDiscount) || 0) / subtotal : 0;
    })();
    const belowListPriceItems = (form.items || []).filter((it) => {
        const p = products.find((x) => x.id === it.productId);
        if (!p) return false;
        const floor = minSellPrice(p, currentUser.role);
        return floor > 0 && it.price * (1 - discountRate) < floor;
    });

    const itemsInvalid = () => {
        if (!form.items || form.items.length === 0) return true;
        return form.items.some((it) => {
            const p = products.find((x) => x.id === it.productId);
            if (!p) return true;
            // Sản phẩm có series: cho phép series chưa đủ (đơn sẽ ở trạng thái "chờ hàng", bổ sung series sau).
            // Sản phẩm không series: vẫn cho tạo đơn kể cả khi vượt tồn kho (ghi nhận âm kho).
            return false;
        });
    };

    const submit = () => {
        if (itemsInvalid()) { alert("Vui lòng thêm ít nhất một sản phẩm vào đơn."); return; }
        const now = new Date().toISOString();
        const itemsWithFulfilled = form.items.map((it) => {
            const p = products.find((x) => x.id === it.productId);
            const fulfilled = !p?.hasSeries || it.series.length === it.qty;
            return { ...it, fulfilled };
        });

        if (editingOrderId) {
            // Sửa đơn hàng đang "Chờ xử lý": hoàn trả kho cho các sản phẩm cũ đã xuất, rồi xuất lại kho theo danh sách sản phẩm mới.
            const order = orders.find((o) => o.id === editingOrderId);
            if (!order) { setCreating(false); setEditingOrderId(null); return; }
            setProducts((prev) => prev.map((p) => {
                if (p.isService) return p; // sản phẩm dịch vụ: không quản lý tồn kho
                let movements = p.movements;
                const oldIt = order.items.find((i) => i.productId === p.id && i.fulfilled);
                if (oldIt) movements = [...movements, { id: uid(), type: "in", docNo: `${order.code}-SUA`, date: todayISO(), qty: oldIt.qty, price: oldIt.price, series: oldIt.series }];
                const newIt = itemsWithFulfilled.find((i) => i.productId === p.id && i.fulfilled);
                if (newIt) movements = [...movements, { id: uid(), type: "out", docNo: order.code, date: todayISO(), qty: newIt.qty, price: newIt.price, series: newIt.series }];
                return movements === p.movements ? p : { ...p, movements };
            }));
            setOrders((prev) => prev.map((o) => (o.id === editingOrderId ? {
                ...o, customerId: form.customerId, channel: form.channel, branch: form.branch, seller: form.seller, deliveryDate: form.deliveryDate,
                shippingAddress: form.shippingAddress || null, tags: form.tags, notes: form.notes, items: itemsWithFulfilled, vat: form.vat,
                orderDiscount: Number(form.orderDiscount) || 0, discountType: form.discountType || "amount", creditDays: Number(form.creditDays) || 0, shippingFee: Number(form.shippingFee) || 0,
            } : o)));
            addLog("Sửa đơn hàng", `${order.code} · ${vnd(calc.payable)}`);
            setCreating(false);
            setEditingOrderId(null);
            return;
        }

        const code = nextOrderCode(orders);
        const paidAmount = Number(form.paidAmount) || 0;
        const payable = calc.payable;
        const needsApproval = !isAdmin && !isCtv && belowListPriceItems.length > 0;
        const approvalReason = belowListPriceItems.length > 0 ? "Bán dưới giá niêm yết" : "";
        const newOrder = {
            id: uid(), code, createdAt: now, date: todayISO(),
            customerId: form.customerId, channel: form.channel, branch: form.branch, seller: form.seller, deliveryDate: form.deliveryDate,
            shippingAddress: form.shippingAddress || null,
            tags: form.tags, notes: form.notes, status: "pending", items: itemsWithFulfilled, vat: form.vat,
            shippingAt: null, deliveredAt: null, paidCompleteAt: paidAmount >= payable && payable > 0 ? now : null, cancelledAt: null,
            orderDiscount: Number(form.orderDiscount) || 0, discountType: form.discountType || "amount", creditDays: Number(form.creditDays) || 0, shippingFee: Number(form.shippingFee) || 0, paidAmount,
            approvalStatus: needsApproval ? "pending" : "approved", approvalReason, createdByRole: currentUser.role,
        };
        setOrders((prev) => [newOrder, ...prev]);
        setProducts((prev) => prev.map((p) => {
            if (p.isService) return p; // sản phẩm dịch vụ: không quản lý tồn kho, không trừ kho
            const it = itemsWithFulfilled.find((i) => i.productId === p.id);
            if (!it || !it.fulfilled) return p; // sản phẩm chờ hàng: chưa trừ kho, sẽ trừ khi bổ sung đủ series
            return { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: code, date: todayISO(), qty: it.qty, price: it.price, series: it.series }] };
        }));
        addLog(needsApproval ? "Tạo đơn (chờ duyệt)" : "Tạo đơn hàng", `${code} · ${vnd(payable)}${approvalReason ? " · " + approvalReason : ""}`);
        setCreating(false);
    };
    // Sửa đơn hàng — mọi vai trò đều được sửa khi đơn đang "Chờ xử lý" (mở lại đúng form tạo đơn với dữ liệu hiện có).
    const openEditOrder = (order) => {
        setForm({
            customerId: order.customerId, channel: order.channel, branch: order.branch, seller: order.seller, deliveryDate: order.deliveryDate,
            shippingAddress: order.shippingAddress || null, tags: order.tags, notes: order.notes,
            items: order.items.map((it) => ({ ...it, series: [...(it.series || [])] })),
            vat: order.vat, orderDiscount: order.orderDiscount, discountType: order.discountType, creditDays: order.creditDays, shippingFee: order.shippingFee, paidAmount: order.paidAmount,
        });
        setEditingOrderId(order.id);
        setViewingId(null);
        setCreating(true);
    };

    const setStatus = (id, status) => {
        const order = orders.find((o) => o.id === id);
        if (order && order.status === "cancelled") return;
        if (status === "shipping" || status === "delivered" || status === "done") {
            if (order && order.items.some((it) => !it.fulfilled)) {
                alert("Đơn còn sản phẩm chờ hàng (chưa đủ số series) — vui lòng bổ sung series trước khi chuyển sang giao hàng.");
                return;
            }
            if (order && order.approvalStatus === "pending") {
                alert("Đơn này do CTV tạo và đang chờ quản trị viên duyệt — chưa thể chuyển sang giao hàng.");
                return;
            }
        }
        if (status === "done" && order && orderCalc(order).remaining > 0) {
            alert("Chỉ có thể chuyển sang \"Hoàn thành\" khi khách đã thanh toán đủ.");
            return;
        }
        const now = new Date().toISOString();
        setOrders((prev) => prev.map((o) => {
            if (o.id !== id) return o;
            const next = { ...o, status };
            if (status === "shipping" && !o.shippingAt) next.shippingAt = now;
            if (status === "delivered") { if (!o.shippingAt) next.shippingAt = now; if (!o.deliveredAt) next.deliveredAt = now; }
            if (status === "done") { if (!o.shippingAt) next.shippingAt = now; if (!o.deliveredAt) next.deliveredAt = now; }
            if (status === "cancelled" && !o.cancelledAt) next.cancelledAt = now;
            return next;
        }));
        addLog("Đổi trạng thái đơn", `${order?.code} → ${STATUSES.find((s) => s.id === status)?.label}`);
    };
    // Hàm huỷ đơn dùng chung cho MỌI đường huỷ (Admin/Nhân viên/CTV huỷ trực tiếp khi đơn còn "Chờ xử lý", hoặc không duyệt đơn chờ duyệt):
    // tự động nhập lại toàn bộ sản phẩm đã xuất kho của đơn về kho, và hoàn tiền cho khách nếu đã thanh toán.
    const executeCancelOrder = (order, reason, logAction) => {
        const now = new Date().toISOString();
        const code = `${order.code}-HUY`;
        setProducts((prev) => prev.map((p) => {
            if (p.isService) return p; // sản phẩm dịch vụ: không quản lý tồn kho, không cần hoàn kho
            const it = order.items.find((i) => i.productId === p.id && i.fulfilled);
            if (!it) return p;
            return { ...p, movements: [...p.movements, { id: uid(), type: "in", docNo: code, date: todayISO(), qty: it.qty, price: it.price, series: it.series }] };
        }));
        setOrders((prev) => prev.map((o) => (o.id === order.id ? {
            ...o, status: "cancelled", cancelledAt: now, cancelReason: reason || "", cancelRequest: null, approvalStatus: "approved", paidAmount: 0,
            cancelledByRole: currentUser.role, cancelledByName: currentUser.fullName,
        } : o)));
        const refundNote = order.paidAmount > 0 ? ` · đã hoàn ${vnd(order.paidAmount)} cho khách` : "";
        addLog(logAction, `${order.code}${reason ? ` · Lý do: ${reason}` : ""} · đã nhập lại kho${refundNote}`);
    };

    const approveOrder = (order) => {
        setOrders((prev) => prev.map((o) => (o.id === order.id ? { ...o, approvalStatus: "approved" } : o)));
        addLog("Duyệt đơn hàng", `${order.code} (tạo bởi ${ACCOUNT_ROLES.find((r) => r.id === order.createdByRole)?.label || order.createdByRole})`);
    };
    const rejectOrder = (order) => {
        executeCancelOrder(order, order.approvalReason ? `Không duyệt — ${order.approvalReason}` : "Không duyệt đơn", "Không duyệt đơn (tự động huỷ + hoàn kho)");
    };
    const [cancelling, setCancelling] = useState(false);
    const [cancelReasonInput, setCancelReasonInput] = useState("");
    // Huỷ đơn trực tiếp — mọi vai trò đều huỷ được khi đơn đang "Chờ xử lý" (không cần chờ duyệt); admin sẽ được báo qua thông báo/nhật ký.
    const cancelOrder = () => {
        if (!viewingOrder) return;
        const reason = cancelReasonInput || viewingOrder.cancelRequest?.reason || "";
        executeCancelOrder(viewingOrder, reason, isAdmin ? "Huỷ đơn hàng" : "Huỷ đơn hàng (tự huỷ khi Chờ xử lý)");
        setCancelling(false);
        setCancelReasonInput("");
    };
    const approveCancelRequest = (order) => {
        executeCancelOrder(order, order.cancelRequest?.reason || "", "Duyệt yêu cầu huỷ đơn");
    };
    const rejectCancelRequest = (order) => {
        setOrders((prev) => prev.map((o) => (o.id === order.id ? { ...o, cancelRequest: null } : o)));
        addLog("Từ chối yêu cầu huỷ đơn", order.code);
    };
    // Bổ sung series cho sản phẩm "chờ hàng" sau khi đã nhập hàng về, rồi trừ kho ngay lúc này.
    const fulfillPendingItem = (order, productId, series) => {
        setOrders((prev) => prev.map((o) => (o.id === order.id ? { ...o, items: o.items.map((it) => (it.productId === productId ? { ...it, series, fulfilled: true } : it)) } : o)));
        const item = order.items.find((it) => it.productId === productId);
        setProducts((prev) => prev.map((p) => (p.id === productId ? { ...p, movements: [...p.movements, { id: uid(), type: "out", docNo: order.code, date: todayISO(), qty: item.qty, price: item.price, series }] } : p)));
    };
    const remove = (id) => setOrders((prev) => prev.filter((o) => o.id !== id));
    const [deletingOrder, setDeletingOrder] = useState(null);
    const [deletePasswordInput, setDeletePasswordInput] = useState("");
    const [deletePasswordError, setDeletePasswordError] = useState("");
    const confirmDeleteOrder = async () => {
        const ok = await verifyPassword(deletePasswordInput, currentUser.passwordSalt, currentUser.passwordHash);
        if (!ok) { setDeletePasswordError("Sai mật khẩu."); return; }
        remove(deletingOrder.id);
        addLog("Xoá đơn hàng", deletingOrder.code);
        setDeletingOrder(null); setDeletePasswordInput(""); setDeletePasswordError("");
    };
    const addPayment = () => {
        const amt = Number(payInput);
        if (!amt || amt <= 0) return;
        const now = new Date().toISOString();
        const order = orders.find((o) => o.id === viewingId);
        setOrders((prev) => prev.map((o) => {
            if (o.id !== viewingId) return o;
            const newPaid = (o.paidAmount || 0) + amt;
            const c = orderCalc({ ...o, paidAmount: newPaid });
            return { ...o, paidAmount: newPaid, paidCompleteAt: o.paidCompleteAt || (c.remaining <= 0 ? now : null), payments: [...(o.payments || []), { id: uid(), date: now, amount: amt, type: "thu" }] };
        }));
        addLog("Ghi nhận thanh toán", `${order?.code} · ${vnd(amt)}`);
        setPayInput("");
    };
    // Ghi nhận số tiền đã hoàn lại cho khách (trường hợp "Cần hoàn lại cho khách") — trừ vào paidAmount,
    // khi khớp đủ (còn lại = 0) thì tự đánh dấu đã hoàn tất thanh toán.
    const addRefund = () => {
        const amt = Number(refundInput);
        if (!amt || amt <= 0) return;
        const now = new Date().toISOString();
        const order = orders.find((o) => o.id === viewingId);
        setOrders((prev) => prev.map((o) => {
            if (o.id !== viewingId) return o;
            const newPaid = (o.paidAmount || 0) - amt;
            const c = orderCalc({ ...o, paidAmount: newPaid });
            return { ...o, paidAmount: newPaid, paidCompleteAt: Math.abs(c.remaining) < 1 ? now : o.paidCompleteAt, payments: [...(o.payments || []), { id: uid(), date: now, amount: amt, type: "hoan" }] };
        }));
        addLog("Ghi nhận hoàn tiền", `${order?.code} · ${vnd(amt)}`);
        setRefundInput("");
    };
    const [invoiceNoInput, setInvoiceNoInput] = useState("");
    const setInvoiceStatus = (id, invoiceStatus, invoiceNo) => {
        setOrders((prev) => prev.map((o) => (o.id === id ? { ...o, invoiceStatus, invoiceNo: invoiceNo !== undefined ? invoiceNo : o.invoiceNo } : o)));
        const order = orders.find((o) => o.id === id);
        addLog("Cập nhật xuất hoá đơn", `${order?.code} → ${invoiceStatus === "issued" ? "Đã xuất" : "Chờ xuất"}`);
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

    // Thực sự tạo phiếu đổi trả + điều chỉnh tồn kho (chỉ gọi khi admin huỷ trực tiếp, hoặc khi admin duyệt yêu cầu).
    const executeReturn = (order, type, returnedItemsIn, exchangeItemsIn, note) => {
        const code = nextReturnCode(order);
        const now = new Date().toISOString();
        const returnedItems = returnedItemsIn.filter((it) => it.qty > 0);
        const exchangeItems = type === "exchange" ? exchangeItemsIn : [];
        const rec = { id: uid(), code, createdAt: now, type, note, returnedItems, exchangeItems };

        setOrders((prev) => prev.map((o) => (o.id === order.id ? { ...o, returns: [...(o.returns || []), rec], returnRequest: null } : o)));
        setProducts((prev) => prev.map((p) => {
            let movs = [];
            const ret = returnedItems.find((i) => i.productId === p.id);
            if (ret) movs.push({ id: uid(), type: "in", docNo: code, date: todayISO(), qty: ret.qty, price: ret.price, series: ret.series });
            const exc = exchangeItems.find((i) => i.productId === p.id);
            if (exc) movs.push({ id: uid(), type: "out", docNo: code, date: todayISO(), qty: exc.qty, price: exc.price, series: exc.series });
            if (movs.length === 0) return p;
            return { ...p, movements: [...p.movements, ...movs] };
        }));
        addLog(type === "exchange" ? "Tạo phiếu đổi hàng" : "Tạo phiếu hoàn tiền", `${code} (đơn ${order.code})`);
    };

    const submitReturn = () => {
        if (returnInvalid() || !viewingOrder) return;
        const returnedItems = returnForm.returnedItems.filter((it) => it.qty > 0);
        const exchangeItems = returnForm.type === "exchange" ? returnForm.exchangeItems : [];
        if (isAdmin) {
            executeReturn(viewingOrder, returnForm.type, returnedItems, exchangeItems, returnForm.note);
        } else {
            setOrders((prev) => prev.map((o) => (o.id === viewingId ? {
                ...o, returnRequest: {
                    type: returnForm.type, returnedItems, exchangeItems, note: returnForm.note,
                    requestedAt: new Date().toISOString(), requestedByRole: currentUser.role, requestedByName: currentUser.fullName,
                },
            } : o)));
            addLog("Yêu cầu đổi trả hàng", `Đơn ${viewingOrder.code}`);
        }
        setReturning(false);
        setReturnForm(null);
    };
    const approveReturnRequest = (order) => {
        const r = order.returnRequest;
        if (!r) return;
        executeReturn(order, r.type, r.returnedItems, r.exchangeItems, r.note);
        addLog("Duyệt yêu cầu đổi trả hàng", order.code);
    };
    const rejectReturnRequest = (order) => {
        setOrders((prev) => prev.map((o) => (o.id === order.id ? { ...o, returnRequest: null } : o)));
        addLog("Từ chối yêu cầu đổi trả hàng", order.code);
    };

    // Đơn hàng chứa phiếu đổi trả đang xem chi tiết (nếu có)
    const viewingReturnOrder = orders.find((o) => (o.returns || []).some((r) => r.id === viewingReturnId)) || null;
    const viewingReturn = viewingReturnOrder ? viewingReturnOrder.returns.find((r) => r.id === viewingReturnId) : null;
    const saveReturnNote = (note) => {
        setOrders((prev) => prev.map((o) => (o.id === viewingReturnOrder.id ? { ...o, returns: o.returns.map((r) => (r.id === viewingReturnId ? { ...r, note } : r)) } : o)));
    };

    const visible = orders
        .filter((o) => {
            if (filterStatus === "all") return true;
            if (filterStatus === "approval_pending") return o.status !== "cancelled" && o.approvalStatus === "pending";
            if (filterStatus === "return_request") return o.status !== "cancelled" && (o.cancelRequest || o.returnRequest);
            return o.status === filterStatus;
        })
        .filter((o) => {
            if (!searchQuery.trim()) return true;
            const q = searchQuery.trim().toLowerCase();
            const cust = customers.find((c) => c.id === o.customerId);
            return o.code.toLowerCase().includes(q) || (cust?.name || "khách lẻ").toLowerCase().includes(q);
        })
        .filter((o) => !filterSeller || o.seller === filterSeller)
        .filter((o) => !filterInvoice || o.invoiceStatus === filterInvoice)
        .filter((o) => !filterCustomer || o.customerId === filterCustomer)
        .filter((o) => !filterProduct || o.items.some((it) => it.productId === filterProduct))
        .filter((o) => !filterApprovalOnly || o.approvalStatus === "pending")
        .filter((o) => !filterCancelledOnly || o.status === "cancelled")
        .filter((o) => {
            if (!filterFrom && !filterTo) return true;
            const t = new Date(o.createdAt).getTime();
            if (filterFrom && t < new Date(filterFrom).getTime()) return false;
            if (filterTo && t > new Date(filterTo).getTime() + 24 * 3600 * 1000 - 1) return false;
            return true;
        })
        .filter((o) => {
            if (!filterText.trim()) return true;
            const q = filterText.trim().toLowerCase();
            return (o.notes || "").toLowerCase().includes(q) || (o.tags || []).some((t) => t.toLowerCase().includes(q));
        });
    const activeFilterCount = [filterSeller, filterFrom, filterTo, filterInvoice, filterText, filterCustomer, filterProduct].filter(Boolean).length + (filterApprovalOnly ? 1 : 0) + (filterCancelledOnly ? 1 : 0);
    const clearFilters = () => {
        setFilterSeller(""); setFilterFrom(""); setFilterTo(""); setFilterInvoice(""); setFilterText("");
        setFilterCustomer(""); setFilterProduct(""); setFilterApprovalOnly(false); setFilterCancelledOnly(false);
    };

    const toggleSelect = (id) => setSelectedIds((prev) => { const next = new Set(prev); next.has(id) ? next.delete(id) : next.add(id); return next; });
    const toggleSelectAll = () => setSelectedIds((prev) => (prev.size === visible.length ? new Set() : new Set(visible.map((o) => o.id))));

    // Chuyển trạng thái hàng loạt cho các đơn đã chọn — tự động bỏ qua đơn không đủ điều kiện (còn chờ hàng/chưa duyệt/chưa thanh toán đủ/đã huỷ/đang ở bước sau).
    const bulkSetStatus = (newStatus) => {
        const targets = orders.filter((o) => selectedIds.has(o.id));
        if (targets.length === 0) return;
        const rank = { pending: 0, shipping: 1, delivered: 2, done: 3 };
        const eligible = [];
        let skipped = 0;
        targets.forEach((o) => {
            if (o.status === "cancelled") { skipped++; return; }
            if (rank[newStatus] < rank[o.status]) { skipped++; return; }
            if (newStatus === "shipping" || newStatus === "delivered" || newStatus === "done") {
                if (o.items.some((it) => !it.fulfilled) || o.approvalStatus === "pending") { skipped++; return; }
            }
            if (newStatus === "done" && orderCalc(o).remaining > 0) { skipped++; return; }
            eligible.push(o.id);
        });
        if (eligible.length === 0) { alert("Không có đơn nào đủ điều kiện để chuyển trạng thái."); return; }
        const now = new Date().toISOString();
        setOrders((prev) => prev.map((o) => {
            if (!eligible.includes(o.id)) return o;
            const next = { ...o, status: newStatus };
            if (newStatus === "shipping" && !o.shippingAt) next.shippingAt = now;
            if (newStatus === "delivered") { if (!o.shippingAt) next.shippingAt = now; if (!o.deliveredAt) next.deliveredAt = now; }
            if (newStatus === "done") { if (!o.shippingAt) next.shippingAt = now; if (!o.deliveredAt) next.deliveredAt = now; }
            return next;
        }));
        addLog("Đổi trạng thái đơn (hàng loạt)", `${eligible.length} đơn → ${STATUSES.find((s) => s.id === newStatus)?.label}${skipped > 0 ? ` · Bỏ qua ${skipped} đơn không đủ điều kiện` : ""}`);
        setSelectedIds(new Set());
        if (skipped > 0) alert(`Đã chuyển ${eligible.length} đơn. Bỏ qua ${skipped} đơn không đủ điều kiện (còn sản phẩm chờ hàng / chưa duyệt / chưa thanh toán đủ / đã huỷ / trạng thái không hợp lệ).`);
    };

    // Xuất Excel: nếu có chọn dòng thì xuất đúng các dòng đã chọn, không thì xuất theo danh sách đang lọc/hiển thị hiện tại
    // (đã áp dụng sẵn: khoảng ngày, khách hàng, sản phẩm, nhân viên, trạng thái... qua bộ lọc phía trên).
    const exportOrders = () => {
        const list = selectedIds.size > 0 ? orders.filter((o) => selectedIds.has(o.id)) : visible;
        if (list.length === 0) { alert("Không có đơn hàng nào để xuất."); return; }
        const summaryRows = list.map((o) => {
            const cust = customers.find((c) => c.id === o.customerId);
            const c = orderCalc(o);
            return {
                "Mã đơn": o.code, "Ngày tạo": formatDateTime(o.createdAt), "Khách hàng": cust?.name || "Khách lẻ", "SĐT khách": cust?.phone || "",
                "Bán tại": o.branch, "Bán bởi": o.seller, "Nguồn": o.channel === "online" ? "Online" : "Tại cửa hàng",
                "Trạng thái giao": STATUSES.find((s) => s.id === o.status)?.label || o.status,
                "Trạng thái duyệt": o.approvalStatus === "pending" ? "Chờ duyệt" : "Đã duyệt",
                "Xuất hoá đơn": o.status === "cancelled" ? "Huỷ" : o.invoiceStatus === "issued" ? "Đã xuất" : "Chờ xuất",
                "Số hoá đơn": o.invoiceNo || "", "Tổng tiền": c.payable, "Đã trả": o.paidAmount, "Còn phải trả": c.remaining,
                "Ghi chú": o.notes || "", "Tags": (o.tags || []).join(", "),
            };
        });
        const detailRows = [];
        list.forEach((o) => {
            const cust = customers.find((c) => c.id === o.customerId);
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                detailRows.push({
                    "Mã đơn": o.code, "Khách hàng": cust?.name || "Khách lẻ", "Sản phẩm": p?.name || "?", "Mã VT": p?.code || "",
                    "SL": it.qty, "Đơn giá": it.price, "Thành tiền": orderLineTotal(it), "Series": (it.series || []).join(", "),
                });
            });
        });
        exportExcel(`DonHang_${todayISO()}`, [{ name: "Đơn hàng", rows: summaryRows }, { name: "Chi tiết sản phẩm", rows: detailRows }]);
        addLog("Xuất Excel đơn hàng", `${list.length} đơn`);
    };

    return (
        <div>
            <div className="flex items-center justify-between mb-4 gap-3 flex-wrap">
                <div className="grid grid-cols-3 sm:grid-cols-4 lg:grid-cols-8 gap-2 flex-1 max-w-3xl">
                    <FilterChip active={filterStatus === "all"} onClick={() => setFilterStatus("all")}>Tất cả</FilterChip>
                    <FilterChip active={filterStatus === "approval_pending"} onClick={() => setFilterStatus("approval_pending")} color={RUST}>Chờ duyệt</FilterChip>
                    {STATUSES.map((s) => <FilterChip key={s.id} active={filterStatus === s.id} onClick={() => setFilterStatus(s.id)} color={s.color}>{s.label}</FilterChip>)}
                    <FilterChip active={filterStatus === "return_request"} onClick={() => setFilterStatus("return_request")} color={RUST}>Yêu cầu huỷ/đổi trả</FilterChip>
                </div>
                <div className="flex gap-2 flex-wrap">
                    {filterStatus === "pending" && selectedIds.size > 0 && (
                        <button onClick={() => bulkSetStatus("shipping")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: STATUSES.find((s) => s.id === "shipping")?.color, color: STATUSES.find((s) => s.id === "shipping")?.color }}>
                            <Truck size={15} /> Chuyển sang Đang giao ({selectedIds.size})
                        </button>
                    )}
                    {filterStatus === "shipping" && selectedIds.size > 0 && (
                        <button onClick={() => bulkSetStatus("delivered")} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: STATUSES.find((s) => s.id === "delivered")?.color, color: STATUSES.find((s) => s.id === "delivered")?.color }}>
                            <PackageCheck size={15} /> Chuyển sang Đã giao ({selectedIds.size})
                        </button>
                    )}
                    <button onClick={exportOrders} className="flex items-center gap-1.5 px-3.5 py-2 rounded-sm text-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                        <FileSpreadsheet size={15} /> Xuất Excel{selectedIds.size > 0 ? ` (${selectedIds.size})` : ""}
                    </button>
                    <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white" style={{ background: INK }}><Plus size={15} /> Tạo đơn và giao hàng</button>
                </div>
            </div>

            <div className="flex items-center gap-2 mb-5 flex-wrap">
                <div className="relative flex-1 min-w-[220px] max-w-sm">
                    <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 opacity-40" />
                    <input value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="Tìm theo mã đơn hoặc tên khách hàng…"
                           className="w-full pl-8 pr-2 py-2 text-sm rounded-sm border outline-none" style={{ borderColor: LINE, background: "#fff" }} />
                </div>
                <button onClick={() => setShowFilters((s) => !s)} className="flex items-center gap-1.5 text-sm px-3.5 py-2 rounded-sm border" style={{ borderColor: activeFilterCount ? INK : LINE, color: INK, background: showFilters ? PAPER : "#fff" }}>
                    <Filter size={14} /> Bộ lọc {activeFilterCount > 0 && <span className="text-[10px] px-1.5 rounded-full text-white" style={{ background: INK }}>{activeFilterCount}</span>}
                </button>
                {activeFilterCount > 0 && <button onClick={clearFilters} className="text-xs opacity-50 hover:opacity-100 underline">Xoá bộ lọc</button>}
            </div>

            {showFilters && (
                <div className="flex flex-wrap items-end gap-3 mb-5 p-4 rounded-sm overflow-x-auto" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <label className="text-xs shrink-0" style={{ width: 150 }}>
                        <span className="block opacity-60 mb-1">Nhân viên tạo</span>
                        <select value={filterSeller} onChange={(e) => setFilterSeller(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                            <option value="">Tất cả</option>
                            {employeeNames.map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                        </select>
                    </label>
                    <label className="text-xs shrink-0" style={{ width: 190 }}>
                        <span className="block opacity-60 mb-1">Khách hàng</span>
                        <FilterSearchSelect options={customers.map((c) => ({ id: c.id, label: c.name }))} value={filterCustomer} onChange={setFilterCustomer} placeholder="Gõ tên khách hàng…" />
                    </label>
                    <label className="text-xs shrink-0" style={{ width: 190 }}>
                        <span className="block opacity-60 mb-1">Sản phẩm</span>
                        <FilterSearchSelect options={products.map((p) => ({ id: p.id, label: p.name }))} value={filterProduct} onChange={setFilterProduct} placeholder="Gõ tên sản phẩm…" />
                    </label>
                    <label className="text-xs shrink-0" style={{ width: 140 }}>
                        <span className="block opacity-60 mb-1">Từ ngày</span>
                        <input type="date" value={filterFrom} onChange={(e) => setFilterFrom(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="text-xs shrink-0" style={{ width: 140 }}>
                        <span className="block opacity-60 mb-1">Đến ngày</span>
                        <input type="date" value={filterTo} onChange={(e) => setFilterTo(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="text-xs shrink-0" style={{ width: 140 }}>
                        <span className="block opacity-60 mb-1">Xuất hoá đơn</span>
                        <select value={filterInvoice} onChange={(e) => setFilterInvoice(e.target.value)} className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }}>
                            <option value="">Tất cả</option>
                            <option value="issued">Đã xuất</option>
                            <option value="pending">Chờ xuất</option>
                        </select>
                    </label>
                    <label className="text-xs flex-1" style={{ minWidth: 160 }}>
                        <span className="block opacity-60 mb-1">Ghi chú / Tag</span>
                        <input value={filterText} onChange={(e) => setFilterText(e.target.value)} placeholder="Tìm trong ghi chú, tag…" className="w-full border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="flex items-center gap-1.5 text-xs shrink-0 pb-1.5">
                        <input type="checkbox" checked={filterApprovalOnly} onChange={(e) => setFilterApprovalOnly(e.target.checked)} />
                        Chỉ đơn chờ duyệt
                    </label>
                    <label className="flex items-center gap-1.5 text-xs shrink-0 pb-1.5">
                        <input type="checkbox" checked={filterCancelledOnly} onChange={(e) => setFilterCancelledOnly(e.target.checked)} />
                        Chỉ đơn đã huỷ
                    </label>
                </div>
            )}

            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 900 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        <th className="px-3 py-2.5 sticky top-0" style={{ background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}` }}><input type="checkbox" checked={selectedIds.size > 0 && selectedIds.size === visible.length} onChange={toggleSelectAll} /></th>
                        {["Mã đơn", "Ngày tạo", "Khách hàng", "Bán tại", "Trạng thái giao", "Xuất hoá đơn", "Tổng tiền", "Còn phải trả", ""].map((h, hi) => (
                            <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap sticky top-0" style={{ color: INK, opacity: 0.6, background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}` }}>{h}</th>
                        ))}
                    </tr>
                    </thead>
                    <tbody>
                    {visible.map((o) => {
                        const cust = customers.find((c) => c.id === o.customerId);
                        const c = orderCalc(o);
                        return (
                            <tr key={o.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                                <td className="px-3 py-3"><input type="checkbox" checked={selectedIds.has(o.id)} onChange={() => toggleSelect(o.id)} /></td>
                                <td className="px-3 py-3 font-medium whitespace-nowrap">
                                    <button onClick={() => { setViewingId(o.id); setPayInput(""); }} className="hover:underline" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{o.code}</button>
                                </td>
                                <td className="px-3 py-3 whitespace-nowrap opacity-80" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(o.createdAt)}</td>
                                <td className="px-3 py-3 whitespace-nowrap">{cust?.name || "Khách lẻ"}</td>
                                <td className="px-3 py-3 whitespace-nowrap opacity-70">{o.branch}</td>
                                <td className="px-3 py-3">
                                    <div className="flex items-center gap-1.5">
                                        <Stamp status={o.status} />
                                        {o.status !== "cancelled" && o.items.some((it) => !it.fulfilled) && <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BRASS}1A`, color: BRASS }}>Chờ hàng</span>}
                                        {o.status !== "cancelled" && o.approvalStatus === "pending" && <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${RUST}1A`, color: RUST }}>Chờ duyệt</span>}
                                        {o.status !== "cancelled" && o.cancelRequest && <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${RUST}1A`, color: RUST }}>Yêu cầu huỷ</span>}
                                        {o.status !== "cancelled" && o.returnRequest && <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: `${BRASS}1A`, color: BRASS }}>Yêu cầu đổi trả</span>}
                                    </div>
                                </td>
                                <td className="px-3 py-3 whitespace-nowrap">
                    <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: o.status === "cancelled" ? "rgba(0,0,0,0.08)" : o.invoiceStatus === "issued" ? `${FOREST}1A` : `${RUST}1A`, color: o.status === "cancelled" ? INK : o.invoiceStatus === "issued" ? FOREST : RUST, opacity: o.status === "cancelled" ? 0.6 : 1 }}>
                      {o.status === "cancelled" ? "Huỷ" : o.invoiceStatus === "issued" ? `Đã xuất${o.invoiceNo ? ` · ${o.invoiceNo}` : ""}` : "Chờ xuất"}
                    </span>
                                </td>
                                <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(c.payable)}</td>
                                <td className="px-3 py-3 text-right whitespace-nowrap font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: c.remaining > 0 ? RUST : FOREST }}>{vnd(c.remaining)}</td>
                                <td className="px-3 py-3">
                                    {isAdmin && (
                                        <button onClick={() => setDeletingOrder(o)} className="opacity-50 hover:opacity-100" style={{ color: RUST }}><Trash2 size={14} /></button>
                                    )}
                                </td>
                            </tr>
                        );
                    })}
                    {visible.length === 0 && <tr><td colSpan={10} className="text-center py-10 opacity-50">Không có đơn hàng nào.</td></tr>}
                    </tbody>
                </table>
            </div>

            {/* Modal tạo đơn và giao hàng */}
            {creating && (
                <Modal title={editingOrderId ? "Sửa đơn hàng" : "Tạo đơn và giao hàng"} onClose={() => { setCreating(false); setEditingOrderId(null); }} size="3xl">
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
                        <div className="lg:col-span-2 space-y-5 min-w-0">
                            <div className="p-4 rounded-sm" style={{ border: `2px solid ${INK}`, background: "#fff" }}>
                                <p className="text-xs uppercase tracking-wider mb-2 font-medium" style={{ color: INK }}>Thông tin khách hàng</p>
                                <CustomerPicker customers={customers} setCustomers={setCustomers} currentUser={currentUser} onPick={(id) => {
                                    const cust = customers.find((c) => c.id === id);
                                    const def = customerDefaultAddress(cust);
                                    const shippingAddress = def
                                        ? { recipientName: def.recipientName || cust?.name || "", recipientPhone: def.recipientPhone || cust?.phone || "", province: def.province, ward: def.ward, addressDetail: def.addressDetail }
                                        : (cust ? { recipientName: cust.name, recipientPhone: cust.phone, province: cust.province, ward: cust.ward, addressDetail: cust.addressDetail } : null);
                                    setForm({ ...form, customerId: id, shippingAddress });
                                }} />
                                {form.customerId && (() => {
                                    const cust = customers.find((c) => c.id === form.customerId);
                                    return (
                                        <div className="mt-2.5 p-3 rounded-sm flex items-center justify-between" style={{ background: `${FOREST}0F`, border: `1px solid ${FOREST}44` }}>
                                            <div>
                                                <p className="text-[10px] uppercase tracking-wider opacity-60 mb-0.5">Khách hàng</p>
                                                <p className="text-lg font-bold uppercase" style={{ color: INK, letterSpacing: "0.02em" }}>{cust?.name}</p>
                                                {cust?.phone && <p className="text-xs opacity-60 mt-0.5" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{cust.phone}</p>}
                                            </div>
                                            <button onClick={() => setForm({ ...form, customerId: "", shippingAddress: null })} className="text-xs opacity-50 hover:opacity-100 underline shrink-0">Bỏ chọn</button>
                                        </div>
                                    );
                                })()}
                                {!form.customerId && <p className="text-sm mt-2 opacity-50">Khách lẻ (chưa chọn khách hàng)</p>}

                                {form.customerId && (() => {
                                    const cust = customers.find((c) => c.id === form.customerId);
                                    const addr = form.shippingAddress;
                                    const hasAddr = addr && (addr.addressDetail || addr.province || addr.ward);
                                    return (
                                        <div className="mt-2.5 p-3 rounded-sm" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                            <div className="flex items-center justify-between mb-1">
                                                <p className="text-[10px] uppercase tracking-wider opacity-60">Địa chỉ giao hàng</p>
                                                {customerAddressBook(cust).length > 1 && (
                                                    <button type="button" onClick={() => setPickingAddress(true)} className="text-xs underline" style={{ color: BLUE }}>Thay đổi</button>
                                                )}
                                            </div>
                                            {hasAddr ? (
                                                <>
                                                    <p className="text-sm" style={{ color: INK }}>{addr.recipientName || cust?.name || "—"}{addr.recipientPhone ? ` · ${addr.recipientPhone}` : ""}</p>
                                                    <p className="text-sm opacity-70">{[addr.addressDetail, addr.ward, addr.province].filter(Boolean).join(", ")}</p>
                                                </>
                                            ) : (
                                                <p className="text-sm opacity-50">Khách hàng chưa có địa chỉ nào — có thể thêm trong hồ sơ khách hàng (mục Khách hàng &gt; Địa chỉ).</p>
                                            )}
                                        </div>
                                    );
                                })()}
                            </div>

                            <div className="p-4 rounded-sm" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Thông tin xuất hàng</p>
                                <div className="grid grid-cols-2 gap-3">
                                    <Field label="Bán tại">
                                        <select className={inputCls} style={{ borderColor: LINE }} value={form.branch} onChange={(e) => setForm({ ...form, branch: e.target.value })}>
                                            {BRANCHES.map((b) => <option key={b} value={b}>{b}</option>)}
                                        </select>
                                    </Field>
                                    <Field label="Bán bởi">
                                        <select className={inputCls} style={{ borderColor: LINE }} value={form.seller} onChange={(e) => setForm({ ...form, seller: e.target.value })}>
                                            {(employeeNames.length ? employeeNames : EMPLOYEES).map((e2) => <option key={e2} value={e2}>{e2}</option>)}
                                        </select>
                                    </Field>
                                </div>
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
                                <div className="grid grid-cols-2 gap-3">
                                    <Field label="Hẹn giao">
                                        <input type="date" className={inputCls} style={{ borderColor: LINE }} value={form.deliveryDate} onChange={(e) => setForm({ ...form, deliveryDate: e.target.value })} />
                                    </Field>
                                    <Field label="Mã đơn"><input disabled className={inputCls} style={{ borderColor: LINE, opacity: 0.5 }} value={nextOrderCode(orders)} /></Field>
                                </div>
                                <Field label="Công nợ khách hàng (số ngày)" hint="Bỏ trống/0 nếu khách trả ngay — dùng cho khách B2B mua trả chậm">
                                    <input type="number" min={0} className={inputCls} style={{ borderColor: LINE }} value={form.creditDays} onChange={(e) => setForm({ ...form, creditDays: e.target.value })} placeholder="0" />
                                </Field>
                            </div>

                            {belowListPriceItems.length > 0 && (
                                <div className="p-3 rounded-sm flex items-start gap-2.5" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44` }}>
                                    <AlertTriangle size={15} style={{ color: RUST }} className="mt-0.5 shrink-0" />
                                    <p className="text-xs" style={{ color: INK }}>
                                        <span className="font-medium" style={{ color: RUST }}>Có sản phẩm bán dưới giá niêm yết: </span>
                                        {belowListPriceItems.map((it) => products.find((p) => p.id === it.productId)?.name).join(", ")}.
                                        {isAdmin ? " (Admin — vẫn tạo đơn bình thường)" : " Đơn sẽ cần Quản trị viên duyệt trước khi giao."}
                                    </p>
                                </div>
                            )}

                            <div className="p-4 rounded-sm" style={{ border: `2px solid ${INK}`, background: "#fff" }}>
                                <p className="text-xs uppercase tracking-wider mb-2 font-medium" style={{ color: INK }}>Thông tin sản phẩm</p>
                                <ProductPicker products={products} onPick={addItem} />
                                <div className="mt-3">
                                    <SalesItemsTable items={form.items || []} products={products} onUpdate={updateItem} onRemove={removeItem} role={currentUser.role} />
                                </div>
                            </div>
                        </div>

                        <div className="space-y-4 min-w-0">
                            <div className="p-4 rounded-sm space-y-2.5" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                                <div className="text-sm">
                                    <div className="flex justify-between mb-1.5"><span className="opacity-60">Tổng tiền ({form.items?.length || 0} sản phẩm)</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(calc.subtotal)}</span></div>
                                    <span className="opacity-60 block mb-1.5">VAT</span>
                                    <div className="flex gap-1.5 flex-wrap">
                                        {VAT_OPTIONS.map((v) => (
                                            <button key={v.id} type="button" onClick={() => setForm({ ...form, vat: v.id })}
                                                    className="px-2.5 py-1.5 rounded-sm text-xs border"
                                                    style={{ borderColor: form.vat === v.id ? INK : LINE, background: form.vat === v.id ? INK : "transparent", color: form.vat === v.id ? "#fff" : INK }}>
                                                {v.label}
                                            </button>
                                        ))}
                                    </div>
                                </div>
                                <div className="flex justify-between text-sm opacity-60"><span>Trong đó VAT</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.vatTotal)}</span></div>
                                <div className="flex justify-between items-center text-sm">
                                    <span className="opacity-60">Chiết khấu</span>
                                    <div className="flex items-center gap-1.5">
                                        <div className="flex rounded-sm border overflow-hidden" style={{ borderColor: LINE }}>
                                            <button type="button" onClick={() => setForm({ ...form, discountType: "amount" })} className="px-2 py-1 text-[11px]" style={{ background: (form.discountType || "amount") === "amount" ? INK : "transparent", color: (form.discountType || "amount") === "amount" ? "#fff" : INK }}>đ</button>
                                            <button type="button" onClick={() => setForm({ ...form, discountType: "percent" })} className="px-2 py-1 text-[11px]" style={{ background: form.discountType === "percent" ? INK : "transparent", color: form.discountType === "percent" ? "#fff" : INK }}>%</button>
                                        </div>
                                        <MoneyInput value={form.orderDiscount} onChange={(v) => setForm({ ...form, orderDiscount: v })} className="w-20 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                    </div>
                                </div>
                                {calc.discountAmount > 0 && form.discountType === "percent" && (
                                    <div className="flex justify-between text-xs opacity-50"><span>Số tiền chiết khấu</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.discountAmount)}</span></div>
                                )}
                                <div className="flex justify-between items-center text-sm">
                                    <span className="opacity-60">Phí giao hàng</span>
                                    <MoneyInput value={form.shippingFee} onChange={(v) => setForm({ ...form, shippingFee: v })} className="w-28 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                </div>
                                <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: INK }}><span>Khách phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.payable)}</span></div>
                                <div className="flex justify-between items-center text-sm">
                                    <span className="opacity-60">Khách đã trả</span>
                                    <MoneyInput value={form.paidAmount} onChange={(v) => setForm({ ...form, paidAmount: v })} className="w-28 border rounded-sm py-1 px-2 text-right text-sm" style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} />
                                </div>
                                <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: calc.remaining > 0 ? RUST : FOREST }}><span>Còn phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(calc.remaining)}</span></div>
                            </div>

                            <button onClick={submit} disabled={itemsInvalid()} className="w-full py-3 rounded-sm text-white text-sm font-medium disabled:opacity-40" style={{ background: INK }}>{editingOrderId ? "Lưu thay đổi" : "Tạo đơn hàng"}</button>
                        </div>
                    </div>

                    <TagsNotesCompact tags={form.tags} setTags={(arr) => setForm({ ...form, tags: arr })} notes={form.notes} setNotes={(v) => setForm({ ...form, notes: v })} />
                </Modal>
            )}

            {/* Modal chọn địa chỉ giao hàng trong sổ địa chỉ của khách đang chọn */}
            {pickingAddress && (() => {
                const cust = customers.find((c) => c.id === form.customerId);
                const addrs = customerAddressBook(cust);
                return (
                    <Modal title="Chọn địa chỉ giao hàng" onClose={() => setPickingAddress(false)}>
                        {addrs.length === 0 ? (
                            <p className="text-sm opacity-50 text-center py-8">Khách hàng này chưa có địa chỉ nào trong sổ địa chỉ.</p>
                        ) : (
                            <div className="space-y-2">
                                {addrs.map((a) => (
                                    <button key={a.id} onClick={() => {
                                        setForm({ ...form, shippingAddress: { recipientName: a.recipientName || cust.name, recipientPhone: a.recipientPhone || cust.phone, province: a.province, ward: a.ward, addressDetail: a.addressDetail } });
                                        setPickingAddress(false);
                                    }} className="w-full text-left p-3 rounded-sm border hover:bg-black/5" style={{ borderColor: LINE }}>
                                        <div className="flex items-center gap-2 mb-1">
                                            <span className="text-sm font-medium" style={{ color: INK }}>{a.label || "Địa chỉ"}</span>
                                            {a.isDefault && <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: `${BLUE}1A`, color: BLUE }}>Mặc định</span>}
                                        </div>
                                        <p className="text-sm" style={{ color: INK }}>{a.recipientName || cust.name}{a.recipientPhone ? ` · ${a.recipientPhone}` : ""}</p>
                                        <p className="text-sm opacity-70">{[a.addressDetail, a.ward, a.province].filter(Boolean).join(", ")}</p>
                                    </button>
                                ))}
                            </div>
                        )}
                    </Modal>
                );
            })()}

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

                        <div className="flex flex-wrap gap-1.5 mb-1.5">
                            {viewingOrder.status === "cancelled" ? (
                                <span className="text-[11px] px-2.5 py-1.5 rounded-sm opacity-50" style={{ background: PAPER, color: RUST }}>Đơn đã huỷ — không thể đổi trạng thái giao hàng</span>
                            ) : STATUSES.filter((s) => s.id !== "cancelled").map((s) => {
                                const rank = { pending: 0, shipping: 1, delivered: 2, done: 3 };
                                const isPast = rank[s.id] < rank[viewingOrder.status];
                                if (s.id === "done") {
                                    const blocked = orderCalc(viewingOrder).remaining > 0 || isPast;
                                    return (
                                        <button key={s.id} onClick={() => setStatus(viewingOrder.id, s.id)} disabled={blocked} className="text-[11px] px-2.5 py-1.5 rounded-sm disabled:opacity-35"
                                                style={{ background: s.id === viewingOrder.status ? `${s.color}22` : PAPER, color: s.color }}
                                                title={isPast ? "Đã qua bước này — không thể quay lại" : blocked ? "Chỉ hoàn thành được khi khách đã thanh toán đủ" : undefined}>{s.label}</button>
                                    );
                                }
                                const blockedByFlow = (s.id === "shipping" || s.id === "delivered") && (viewingOrder.items.some((it) => !it.fulfilled) || viewingOrder.approvalStatus === "pending");
                                const blockedByPaid = viewingOrder.paidAmount > 0;
                                const blocked = blockedByFlow || blockedByPaid || isPast;
                                return (
                                    <button key={s.id} onClick={() => setStatus(viewingOrder.id, s.id)} disabled={blocked} className="text-[11px] px-2.5 py-1.5 rounded-sm disabled:opacity-35"
                                            style={{ background: s.id === viewingOrder.status ? `${s.color}22` : PAPER, color: s.color }}
                                            title={isPast ? "Đã qua bước này — không thể quay lại" : blockedByPaid ? "Đơn đã ghi nhận thanh toán — chỉ có thể chuyển sang Hoàn thành" : blockedByFlow ? "Còn sản phẩm chờ hàng hoặc đơn chưa được duyệt" : undefined}>{s.label}</button>
                                );
                            })}
                        </div>
                        {viewingOrder.items.some((it) => !it.fulfilled) && (
                            <p className="text-[11px] mb-1" style={{ color: BRASS }}>⚠ Còn sản phẩm chờ hàng — bổ sung series bên dưới trước khi chuyển "Đang giao"/"Đã giao".</p>
                        )}
                        {viewingOrder.paidAmount > 0 && viewingOrder.status !== "cancelled" && viewingOrder.status !== "done" && (
                            <p className="text-[11px] mb-1" style={{ color: FOREST }}>✓ Đơn đã có thanh toán — Chờ xử lý/Đang giao/Đã giao đã khoá, dùng nút "Hoàn thành" khi đã thanh toán đủ.</p>
                        )}
                        {viewingOrder.approvalStatus === "pending" && (
                            <div className="flex items-center gap-2 mb-3 p-3 rounded-sm flex-wrap" style={{ background: `${RUST}0D`, border: `1px solid ${RUST}44` }}>
                                <span className="text-[11px]" style={{ color: RUST }}>⚠ Đơn cần duyệt{viewingOrder.approvalReason ? ` — Lý do: ${viewingOrder.approvalReason}` : ""} — đang chờ quản trị viên duyệt.</span>
                                {isAdmin && (
                                    <div className="flex gap-2 ml-auto">
                                        <button onClick={() => rejectOrder(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm border flex items-center gap-1" style={{ borderColor: RUST, color: RUST }}>
                                            <XCircle size={12} /> Không duyệt
                                        </button>
                                        <button onClick={() => approveOrder(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm text-white flex items-center gap-1" style={{ background: FOREST }}>
                                            <ShieldCheck size={12} /> Duyệt đơn
                                        </button>
                                    </div>
                                )}
                            </div>
                        )}

                        {viewingOrder.cancelRequest && (
                            <div className="flex items-center gap-2 mb-3 p-3 rounded-sm flex-wrap" style={{ background: `${RUST}0D`, border: `1px solid ${RUST}44` }}>
                <span className="text-[11px]" style={{ color: RUST }}>
                  ⚠ {viewingOrder.cancelRequest.requestedByName} ({ACCOUNT_ROLES.find((r) => r.id === viewingOrder.cancelRequest.requestedByRole)?.label}) yêu cầu huỷ đơn
                    {viewingOrder.cancelRequest.reason ? ` — Lý do: ${viewingOrder.cancelRequest.reason}` : ""}. Đang chờ quản trị viên duyệt.
                </span>
                                {isAdmin && (
                                    <div className="flex gap-2 ml-auto">
                                        <button onClick={() => rejectCancelRequest(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Từ chối</button>
                                        <button onClick={() => approveCancelRequest(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm text-white flex items-center gap-1" style={{ background: RUST }}>
                                            <XCircle size={12} /> Duyệt huỷ
                                        </button>
                                    </div>
                                )}
                            </div>
                        )}

                        {viewingOrder.returnRequest && (
                            <div className="flex items-center gap-2 mb-3 p-3 rounded-sm flex-wrap" style={{ background: `${BRASS}0D`, border: `1px solid ${BRASS}44` }}>
                <span className="text-[11px]" style={{ color: BRASS }}>
                  ⚠ {viewingOrder.returnRequest.requestedByName} ({ACCOUNT_ROLES.find((r) => r.id === viewingOrder.returnRequest.requestedByRole)?.label}) yêu cầu {viewingOrder.returnRequest.type === "exchange" ? "đổi hàng" : "hoàn tiền"}
                    {viewingOrder.returnRequest.note ? ` — Lý do: ${viewingOrder.returnRequest.note}` : ""}. Đang chờ quản trị viên duyệt.
                </span>
                                {isAdmin && (
                                    <div className="flex gap-2 ml-auto">
                                        <button onClick={() => rejectReturnRequest(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Từ chối</button>
                                        <button onClick={() => approveReturnRequest(viewingOrder)} className="text-[11px] px-3 py-1.5 rounded-sm text-white flex items-center gap-1" style={{ background: BRASS }}>
                                            <RotateCcw size={12} /> Duyệt đổi trả
                                        </button>
                                    </div>
                                )}
                            </div>
                        )}

                        <div className="flex flex-wrap gap-2 mb-4">
                            {!isCtv && (
                                <button onClick={() => { setPrintDocType("invoice"); setPrintPaperSize(printSettings?.defaultPaperSize || "A4"); setPrintBlockedUrl(null); setPrintModalOpen(true); }} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: LINE, color: INK }}>
                                    <Printer size={13} /> In đơn hàng
                                </button>
                            )}
                            {viewingOrder.status === "pending" && (
                                <button onClick={() => openEditOrder(viewingOrder)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: BLUE, color: BLUE }}>
                                    <Pencil size={13} /> Sửa đơn hàng
                                </button>
                            )}
                            {!viewingOrder.returnRequest && viewingOrder.status !== "cancelled" && viewingOrder.approvalStatus !== "pending" && (
                                <button onClick={() => openReturn(viewingOrder)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: BRASS, color: BRASS }}>
                                    <RotateCcw size={13} /> {isAdmin ? "Đổi trả hàng" : "Yêu cầu đổi trả hàng"}
                                </button>
                            )}
                        </div>
                        {viewingOrder.status === "pending" && viewingOrder.paidAmount === 0 && (
                            <div className="-mt-2 mb-4">
                                <button onClick={() => setCancelling(true)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: RUST, color: RUST }}>
                                    <XCircle size={13} /> Huỷ đơn
                                </button>
                            </div>
                        )}
                        {(viewingOrder.status === "shipping" || viewingOrder.status === "delivered" || viewingOrder.status === "done") && (
                            <p className="text-[11px] mb-3 opacity-50">Đơn đã qua bước "Đang giao" — không thể huỷ, chỉ có thể đổi trả hàng.</p>
                        )}


                        {!isCtv && viewingOrder.status !== "cancelled" && viewingOrder.approvalStatus !== "pending" && (
                            <div className="flex items-center gap-2 mb-4 p-3 rounded-sm flex-wrap" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                <span className="text-xs opacity-60 shrink-0">Xuất hoá đơn:</span>
                                <button onClick={() => setInvoiceStatus(viewingOrder.id, "pending")} className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: viewingOrder.invoiceStatus === "pending" ? `${RUST}1A` : "transparent", color: RUST, border: `1px solid ${viewingOrder.invoiceStatus === "pending" ? RUST : LINE}` }}>Chờ xuất</button>
                                <button onClick={() => setInvoiceStatus(viewingOrder.id, "issued", viewingOrder.invoiceNo)} className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: viewingOrder.invoiceStatus === "issued" ? `${FOREST}1A` : "transparent", color: FOREST, border: `1px solid ${viewingOrder.invoiceStatus === "issued" ? FOREST : LINE}` }}>Đã xuất</button>
                                {viewingOrder.invoiceStatus === "issued" && (
                                    <input defaultValue={viewingOrder.invoiceNo} onBlur={(e) => setInvoiceStatus(viewingOrder.id, "issued", e.target.value)} placeholder="Số hoá đơn"
                                           className="text-xs border rounded-sm py-1 px-2 flex-1 min-w-[120px]" style={{ borderColor: LINE }} />
                                )}
                            </div>
                        )}
                        {viewingOrder.status === "cancelled" && (
                            <p className="text-xs mb-4 opacity-50">Đơn đã huỷ — không xuất hoá đơn.</p>
                        )}
                        {viewingOrder.approvalStatus === "pending" && viewingOrder.status !== "cancelled" && (
                            <p className="text-xs mb-4" style={{ color: BRASS }}>Đơn đang chờ duyệt — tạm khoá xuất hoá đơn và ghi nhận thanh toán.</p>
                        )}

                        <p className="text-[10px] uppercase tracking-wider opacity-50 mb-0.5">Khách hàng</p>
                        <p className="text-base font-bold uppercase mb-1" style={{ color: INK, letterSpacing: "0.02em" }}>{cust?.name || "Khách lẻ"}</p>
                        {viewingOrder.shippingAddress && (viewingOrder.shippingAddress.addressDetail || viewingOrder.shippingAddress.province || viewingOrder.shippingAddress.ward) && (
                            <p className="text-sm mb-1 flex items-start gap-1.5" style={{ color: INK, opacity: 0.8 }}>
                                <MapPin size={13} className="shrink-0 mt-0.5 opacity-50" />
                                <span>
                  {viewingOrder.shippingAddress.recipientName || cust?.name || "—"}{viewingOrder.shippingAddress.recipientPhone ? ` · ${viewingOrder.shippingAddress.recipientPhone}` : ""}
                                    {" — "}{[viewingOrder.shippingAddress.addressDetail, viewingOrder.shippingAddress.ward, viewingOrder.shippingAddress.province].filter(Boolean).join(", ")}
                </span>
                            </p>
                        )}
                        <p className="text-sm mb-1 opacity-70">Bán tại {viewingOrder.branch} · Bán bởi {viewingOrder.seller}{viewingOrder.deliveryDate ? ` · Hẹn giao ${viewingOrder.deliveryDate}` : ""}</p>
                        {viewingOrder.notes && <p className="text-sm mb-3 opacity-70">Ghi chú: {viewingOrder.notes}</p>}

                        <div className="my-3" style={{ borderTop: `1px dashed ${LINE}` }} />

                        <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Thông tin sản phẩm</p>
                        <div className="rounded-sm overflow-x-auto mb-4" style={{ border: `1px solid ${LINE}` }}>
                            <table className="w-full text-sm" style={{ minWidth: isAdmin && viewingOrder.approvalStatus === "pending" ? 720 : 520 }}>
                                <thead><tr className="opacity-60" style={{ background: PAPER }}>
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap">Sản phẩm</th><th className="text-right py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap">SL</th><th className="text-right py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap">Đơn giá</th><th className="text-right py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap">Thành tiền</th>
                                    {isAdmin && viewingOrder.approvalStatus === "pending" && (<>
                                        <th className="text-right py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: RUST }}>Giá nhập</th>
                                        <th className="text-right py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: RUST }}>Lời/Lỗ</th>
                                    </>)}
                                    <th className="text-left py-2 px-2 text-xs uppercase tracking-wider whitespace-nowrap">Series</th>
                                </tr></thead>
                                <tbody>
                                {viewingOrder.items.map((it, i) => {
                                    const p = products.find((x) => x.id === it.productId);
                                    const returnedQty = returnedQtyOf(viewingOrder, it.productId);
                                    const cost = p?.costPrice || 0;
                                    const profitPct = cost > 0 ? ((it.price - cost) / cost) * 100 : null;
                                    return (
                                        <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                            <td className="py-2.5 px-2">
                                                <span className="font-medium" style={{ color: INK }}>{p?.name || "?"}</span>
                                                {returnedQty > 0 && <span className="ml-1.5 text-[10px] whitespace-nowrap" style={{ color: RUST }}>(đã trả {returnedQty})</span>}
                                                {!it.fulfilled && <span className="ml-1.5 text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider whitespace-nowrap" style={{ background: `${BRASS}1A`, color: BRASS }}>Chờ hàng</span>}
                                            </td>
                                            <td className="py-2.5 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{it.qty}</td>
                                            <td className="py-2.5 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(it.price)}</td>
                                            <td className="py-2.5 px-2 text-right font-semibold whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(orderLineTotal(it))}</td>
                                            {isAdmin && viewingOrder.approvalStatus === "pending" && (<>
                                                <td className="py-2.5 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK, opacity: 0.75 }}>{cost > 0 ? vnd(cost) : "—"}</td>
                                                <td className="py-2.5 px-2 text-right font-semibold whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: profitPct === null ? INK : profitPct >= 0 ? FOREST : RUST }}>
                                                    {profitPct === null ? "—" : `${profitPct >= 0 ? "+" : ""}${profitPct.toFixed(1)}%`}
                                                </td>
                                            </>)}
                                            <td className="py-2.5 px-2 whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series?.length ? it.series.join(", ") : "—"}</td>
                                        </tr>
                                    );
                                })}
                                </tbody>
                            </table>
                        </div>

                        {viewingOrder.items.some((it) => !it.fulfilled) && (
                            <div className="mb-4 p-4 rounded-sm" style={{ background: `${BRASS}0D`, border: `1px solid ${BRASS}44` }}>
                                <p className="text-sm font-medium mb-1" style={{ color: BRASS }}>Sản phẩm chờ hàng — bổ sung series khi đã nhập hàng về</p>
                                <p className="text-xs opacity-60 mb-3">Đơn chưa trừ kho cho các sản phẩm này. Chọn đủ số series rồi xác nhận để xuất kho — sau đó mới chuyển được sang "Đang giao".</p>
                                <div className="space-y-3">
                                    {viewingOrder.items.filter((it) => !it.fulfilled).map((it) => {
                                        const p = products.find((x) => x.id === it.productId);
                                        const available = p ? seriesList(p).filter((s) => s.status === "Còn tồn") : [];
                                        const draft = pendingFulfillDraft[it.productId] || it.series || [];
                                        return (
                                            <div key={it.productId} className="p-3 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                                                <div className="flex items-center justify-between text-sm mb-2">
                                                    <span style={{ color: INK }} className="font-medium">{p?.name}</span>
                                                    <span className="text-xs opacity-50">cần {it.qty} · còn tồn hiện tại {available.length}</span>
                                                </div>
                                                <SeriesPicker available={available} selected={draft} setSelected={(arr) => setPendingFulfillDraft((d) => ({ ...d, [it.productId]: arr }))} need={it.qty} />
                                                <button
                                                    onClick={() => { fulfillPendingItem(viewingOrder, it.productId, draft); setPendingFulfillDraft((d) => { const n = { ...d }; delete n[it.productId]; return n; }); }}
                                                    disabled={draft.length !== it.qty}
                                                    className="mt-2 text-xs px-3 py-1.5 rounded-sm text-white disabled:opacity-40" style={{ background: FOREST }}>
                                                    Xác nhận đã có hàng — xuất kho
                                                </button>
                                            </div>
                                        );
                                    })}
                                </div>
                            </div>
                        )}

                        {viewingOrder.returns?.length > 0 && (
                            <div className="mb-4">
                                <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Lịch sử đổi trả</p>
                                <div className="space-y-2">
                                    {viewingOrder.returns.map((r) => {
                                        const rTotal = r.returnedItems.reduce((s, it) => s + returnLineTotal(it), 0);
                                        return (
                                            <button key={r.id} onClick={() => setViewingReturnId(r.id)} className="w-full flex items-center justify-between px-4 py-3 rounded-sm text-left hover:brightness-95" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
                                                <div>
                                                    <span className="text-sm font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{r.code}</span>
                                                    <span className="ml-2 text-[11px] px-2 py-0.5 rounded-full" style={{ background: r.type === "exchange" ? `${FOREST}1A` : `${BRASS}1A`, color: r.type === "exchange" ? FOREST : BRASS }}>{r.type === "exchange" ? "Đổi hàng" : "Hoàn tiền"}</span>
                                                    <p className="text-xs opacity-50 mt-1">{formatDateTime(r.createdAt)}</p>
                                                </div>
                                                <span className="text-sm font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(rTotal)}</span>
                                            </button>
                                        );
                                    })}
                                </div>
                            </div>
                        )}

                        <div className="p-4 rounded-sm space-y-2.5 mb-4" style={{ border: `1px solid ${LINE}`, background: PAPER }}>
                            <div className="flex justify-between text-sm"><span className="opacity-60">Tổng tiền</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.subtotal)}</span></div>
                            <div className="flex justify-between text-sm opacity-60"><span>Trong đó VAT ({VAT_OPTIONS.find((v) => v.id === viewingOrder.vat)?.label})</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.vatTotal)}</span></div>
                            <div className="flex justify-between text-sm"><span className="opacity-60">Chiết khấu{viewingOrder.discountType === "percent" ? ` (${viewingOrder.orderDiscount}%)` : ""}</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.discountAmount)}</span></div>
                            <div className="flex justify-between text-sm"><span className="opacity-60">Phí giao hàng</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingOrder.shippingFee)}</span></div>
                            {c.returnedValue > 0 && <div className="flex justify-between text-sm"><span className="opacity-60">Giá trị hàng trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>-{vnd(c.returnedValue)}</span></div>}
                            {c.exchangeValue > 0 && <div className="flex justify-between text-sm"><span className="opacity-60">Giá trị hàng đổi</span><span style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>+{vnd(c.exchangeValue)}</span></div>}
                            <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: INK }}><span>Khách phải trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(c.payable)}</span></div>
                            <div className="flex justify-between text-sm"><span className="opacity-60">Khách đã trả</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(viewingOrder.paidAmount)}</span></div>
                            <div className="flex justify-between text-sm font-semibold pt-2" style={{ borderTop: `1px dashed ${LINE}`, color: c.remaining > 0 ? RUST : FOREST }}>
                                <span>{c.remaining >= 0 ? "Còn phải trả" : "Cần hoàn lại cho khách"}</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(Math.abs(c.remaining))}</span>
                            </div>
                        </div>

                        {viewingOrder.status !== "cancelled" && viewingOrder.approvalStatus !== "pending" && c.remaining > 0 && (
                            <div className="flex gap-2">
                                <MoneyInput value={payInput === "" ? "" : Number(payInput)} onChange={(v) => setPayInput(v)} placeholder="Số tiền khách trả thêm"
                                            className="flex-1 border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }} />
                                <button onClick={addPayment} className="px-4 py-2 rounded-sm text-white text-sm" style={{ background: FOREST }}>Ghi nhận thanh toán</button>
                            </div>
                        )}
                        {viewingOrder.status !== "cancelled" && viewingOrder.approvalStatus !== "pending" && c.remaining < 0 && (
                            <div className="flex gap-2">
                                <MoneyInput value={refundInput === "" ? "" : Number(refundInput)} onChange={(v) => setRefundInput(v)} placeholder="Số tiền đã hoàn cho khách"
                                            className="flex-1 border rounded-sm py-2 px-3 text-sm" style={{ borderColor: LINE }} />
                                <button onClick={addRefund} className="px-4 py-2 rounded-sm text-white text-sm" style={{ background: FOREST }}>Ghi nhận hoàn tiền</button>
                            </div>
                        )}
                    </Modal>
                );
            })()}

            {/* Modal tạo phiếu đổi trả hàng */}
            {returning && returnForm && viewingOrder && (
                <Modal title={isAdmin ? `Đổi trả hàng — Đơn ${viewingOrder.code}` : `Yêu cầu đổi trả hàng — Đơn ${viewingOrder.code}`} onClose={() => { setReturning(false); setReturnForm(null); }} size="2xl">
                    {!isAdmin && <p className="text-xs opacity-60 mb-3">Yêu cầu của bạn sẽ được gửi tới quản trị viên để duyệt trước khi tồn kho được điều chỉnh.</p>}
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
                                <SalesItemsTable items={returnForm.exchangeItems} products={products} onUpdate={updateExchangeItem} onRemove={removeExchangeItem} role={currentUser.role} />
                            </div>
                        </>
                    )}

                    <Field label="Ghi chú" hint="VD: lý do đổi trả">
                        <textarea rows={2} className="w-full border rounded-sm p-2 text-sm mt-1" style={{ borderColor: LINE }} value={returnForm.note} onChange={(e) => setReturnForm({ ...returnForm, note: e.target.value })} />
                    </Field>

                    <button onClick={submitReturn} disabled={returnInvalid()} className="w-full py-2.5 rounded-sm text-white text-sm disabled:opacity-40 mt-2" style={{ background: INK }}>
                        {isAdmin ? (returnForm.type === "exchange" ? "Tạo phiếu đổi hàng" : "Tạo phiếu hoàn tiền") : "Gửi yêu cầu đổi trả"}
                    </button>
                </Modal>
            )}

            {/* Modal chi tiết phiếu đổi trả */}
            {viewingReturn && (() => {
                const rTotal = viewingReturn.returnedItems.reduce((s, it) => s + returnLineTotal(it), 0);
                const eTotal = viewingReturn.exchangeItems.reduce((s, it) => s + returnLineTotal(it), 0);
                return (
                    <Modal title={`Phiếu đổi trả ${viewingReturn.code}`} onClose={() => setViewingReturnId(null)} size="lg">
                        <div className="flex items-center gap-2 mb-4">
              <span className="text-[11px] px-2.5 py-1 rounded-full" style={{ background: viewingReturn.type === "exchange" ? `${FOREST}1A` : `${BRASS}1A`, color: viewingReturn.type === "exchange" ? FOREST : BRASS }}>
                {viewingReturn.type === "exchange" ? "Đổi hàng" : "Hoàn tiền"}
              </span>
                            <span className="text-xs opacity-50" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(viewingReturn.createdAt)}</span>
                            <span className="text-xs opacity-40">· thuộc đơn {viewingReturnOrder.code}</span>
                        </div>

                        <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Sản phẩm trả lại</p>
                        <div className="rounded-sm overflow-x-auto mb-4" style={{ border: `1px solid ${LINE}` }}>
                            <table className="w-full text-xs" style={{ minWidth: 380 }}>
                                <thead style={{ background: PAPER }}><tr className="opacity-60">
                                    <th className="text-left py-2 px-2">Sản phẩm</th><th className="text-left py-2 px-2">SL</th><th className="text-right py-2 px-2">Giá</th><th className="text-left py-2 px-2">Series</th>
                                </tr></thead>
                                <tbody>
                                {viewingReturn.returnedItems.map((it, i) => {
                                    const p = products.find((x) => x.id === it.productId);
                                    return (
                                        <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                            <td className="py-1.5 px-2">{p?.name || "?"}</td>
                                            <td className="py-1.5 px-2">{it.qty}</td>
                                            <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td>
                                            <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series?.length ? it.series.join(", ") : "—"}</td>
                                        </tr>
                                    );
                                })}
                                </tbody>
                            </table>
                        </div>
                        <p className="text-sm text-right mb-4" style={{ color: RUST }}>Giá trị hàng trả: -{vnd(rTotal)}</p>

                        {viewingReturn.type === "exchange" && viewingReturn.exchangeItems.length > 0 && (
                            <>
                                <p className="text-xs uppercase tracking-wider mb-2 opacity-60">Sản phẩm đổi lấy</p>
                                <div className="rounded-sm overflow-x-auto mb-2" style={{ border: `1px solid ${LINE}` }}>
                                    <table className="w-full text-xs" style={{ minWidth: 380 }}>
                                        <thead style={{ background: PAPER }}><tr className="opacity-60">
                                            <th className="text-left py-2 px-2">Sản phẩm</th><th className="text-left py-2 px-2">SL</th><th className="text-right py-2 px-2">Giá</th><th className="text-left py-2 px-2">Series</th>
                                        </tr></thead>
                                        <tbody>
                                        {viewingReturn.exchangeItems.map((it, i) => {
                                            const p = products.find((x) => x.id === it.productId);
                                            return (
                                                <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                                    <td className="py-1.5 px-2">{p?.name || "?"}</td>
                                                    <td className="py-1.5 px-2">{it.qty}</td>
                                                    <td className="py-1.5 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(it.price)}</td>
                                                    <td className="py-1.5 px-2" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{it.series?.length ? it.series.join(", ") : "—"}</td>
                                                </tr>
                                            );
                                        })}
                                        </tbody>
                                    </table>
                                </div>
                                <p className="text-sm text-right mb-4" style={{ color: FOREST }}>Giá trị hàng đổi: +{vnd(eTotal)}</p>
                            </>
                        )}

                        <Field label="Ghi chú">
                            <textarea rows={2} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} defaultValue={viewingReturn.note} onBlur={(e) => saveReturnNote(e.target.value)} placeholder="Lý do đổi trả…" />
                        </Field>
                        <p className="text-xs opacity-40 mt-1">Số lượng/sản phẩm trong phiếu đã khoá vì đã cập nhật tồn kho — chỉ có thể sửa ghi chú.</p>
                    </Modal>
                );
            })()}

            {/* Modal huỷ đơn hàng (kèm lý do) — mọi vai trò huỷ trực tiếp được khi đơn còn "Chờ xử lý" và chưa thanh toán. */}
            {cancelling && viewingOrder && (
                <Modal title={`Huỷ đơn hàng ${viewingOrder.code}`} onClose={() => setCancelling(false)}>
                    {!isAdmin && <p className="text-xs opacity-60 mb-3">Đơn sẽ bị huỷ ngay và tự động hoàn lại kho. Quản trị viên sẽ nhận được thông báo về việc này.</p>}
                    <Field label="Lý do huỷ đơn">
                        <textarea rows={3} className="w-full border rounded-sm p-2 text-sm" style={{ borderColor: LINE }} value={cancelReasonInput} onChange={(e) => setCancelReasonInput(e.target.value)} placeholder="VD: khách đổi ý, hết hàng, đặt nhầm…" />
                    </Field>
                    <button onClick={cancelOrder} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: RUST }}>Xác nhận huỷ đơn</button>
                </Modal>
            )}

            {/* Modal xác nhận xoá đơn hàng — chỉ admin, cần nhập mật khẩu */}
            {deletingOrder && (
                <Modal title={`Xoá đơn hàng ${deletingOrder.code}`} onClose={() => { setDeletingOrder(null); setDeletePasswordInput(""); setDeletePasswordError(""); }}>
                    <p className="text-sm mb-3" style={{ color: INK }}>Hành động này không thể hoàn tác. Nhập mật khẩu quản trị viên để xác nhận xoá vĩnh viễn đơn <b>{deletingOrder.code}</b>.</p>
                    <Field label="Mật khẩu">
                        <input type="password" className={inputCls} style={{ borderColor: LINE }} value={deletePasswordInput}
                               onChange={(e) => { setDeletePasswordInput(e.target.value); setDeletePasswordError(""); }}
                               onKeyDown={(e) => { if (e.key === "Enter") confirmDeleteOrder(); }} autoFocus />
                    </Field>
                    {deletePasswordError && <p className="text-sm mb-2" style={{ color: RUST }}>{deletePasswordError}</p>}
                    <button onClick={confirmDeleteOrder} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: RUST }}>Xác nhận xoá</button>
                </Modal>
            )}

            {/* Modal chọn loại chứng từ + khổ giấy + tuỳ chỉnh thông tin in */}
            {printModalOpen && viewingOrder && (
                <Modal title={printDocType === "handover" ? `In biên bản bàn giao (BBBG) ${viewingOrder.code}` : `In đơn hàng ${viewingOrder.code}`} onClose={() => setPrintModalOpen(false)}>
                    <Field label="Loại chứng từ">
                        <div className="grid grid-cols-2 gap-2">
                            <button type="button" onClick={() => setPrintDocType("invoice")}
                                    className="px-3 py-2 rounded-sm text-sm border text-center"
                                    style={{ borderColor: printDocType === "invoice" ? INK : LINE, background: printDocType === "invoice" ? INK : "transparent", color: printDocType === "invoice" ? "#fff" : INK }}>
                                Đơn hàng
                            </button>
                            <button type="button" onClick={() => { setPrintDocType("handover"); if (printPaperSize !== "A4" && printPaperSize !== "A5") setPrintPaperSize("A4"); }}
                                    className="px-3 py-2 rounded-sm text-sm border text-center"
                                    style={{ borderColor: printDocType === "handover" ? INK : LINE, background: printDocType === "handover" ? INK : "transparent", color: printDocType === "handover" ? "#fff" : INK }}>
                                Biên bản bàn giao (BBBG)
                            </button>
                        </div>
                    </Field>

                    <Field label="Khổ giấy">
                        <div className="grid grid-cols-2 gap-2">
                            {(printDocType === "handover" ? PAPER_SIZES.filter((p) => p.id === "A4" || p.id === "A5") : PAPER_SIZES).map((p) => (
                                <button key={p.id} type="button" onClick={() => setPrintPaperSize(p.id)}
                                        className="px-3 py-2 rounded-sm text-sm border text-center"
                                        style={{ borderColor: printPaperSize === p.id ? INK : LINE, background: printPaperSize === p.id ? INK : "transparent", color: printPaperSize === p.id ? "#fff" : INK }}>
                                    {p.label}
                                </button>
                            ))}
                        </div>
                    </Field>

                    {printDocType === "handover" && (() => {
                        const cust = customers.find((c) => c.id === viewingOrder.customerId);
                        const missingRep = cust && (cust.group === "b2b" || cust.group === "enterprise") && !cust.representativeName;
                        return (
                            <p className="text-[11px] opacity-60 -mt-2 mb-3">
                                Bên B (Hilitek) và Bên A lấy tự động từ hồ sơ khách hàng.
                                {missingRep ? " Khách hàng này chưa có thông tin Đại diện/Chức vụ — vào Khách hàng để bổ sung, nếu không phần này sẽ để trống trên biên bản." : ""}
                            </p>
                        );
                    })()}

                    {printDocType === "invoice" && (
                        <button type="button" onClick={() => setShowPrintOptions((s) => !s)} className="text-xs underline opacity-60 hover:opacity-100 mb-2">
                            {showPrintOptions ? "Ẩn tuỳ chỉnh thông tin in" : "Tuỳ chỉnh thông tin in"}
                        </button>
                    )}

                    {printDocType === "invoice" && showPrintOptions && (
                        <div className="p-3 rounded-sm mb-3 space-y-1.5" style={{ background: PAPER }}>
                            {[
                                ["showCompanyInfo", "Thông tin công ty (tên, địa chỉ, MST)"],
                                ["showBankAccount", "Số tài khoản ngân hàng"],
                                ["showCustomerAddress", "Địa chỉ khách hàng"],
                                ["showCustomerPhone", "Số điện thoại khách hàng"],
                                ["showSeriesCol", "Cột số series"],
                                ["showAmountInWords", "Số tiền bằng chữ"],
                                ["showSignatures", "Khối chữ ký"],
                            ].map(([key, label]) => (
                                <label key={key} className="flex items-center gap-2 text-xs" style={{ color: INK }}>
                                    <input type="checkbox" checked={!!printSettings[key]} onChange={(e) => setPrintSettings((s) => ({ ...s, [key]: e.target.checked }))} />
                                    {label}
                                </label>
                            ))}
                            <p className="text-[11px] opacity-50 pt-1">Các tuỳ chỉnh này áp dụng chung cho mọi lần in sau, không phải riêng đơn này.</p>
                        </div>
                    )}

                    <button onClick={executePrint} className="w-full py-2.5 rounded-sm text-white text-sm flex items-center justify-center gap-1.5" style={{ background: INK }}>
                        <Printer size={14} /> In ngay
                    </button>

                    {printBlockedUrl && (
                        <div className="mt-3 p-3 rounded-sm text-xs" style={{ background: `${RUST}10`, border: `1px solid ${RUST}44`, color: INK }}>
                            <p className="mb-2">Trình duyệt đã chặn cửa sổ in tự động. Bấm vào liên kết bên dưới để mở hoá đơn, sau đó tự bấm in (Ctrl+P / biểu tượng in) trong tab mới đó:</p>
                            <a href={printBlockedUrl} target="_blank" rel="noreferrer" className="underline font-medium" style={{ color: BLUE }}>Mở hoá đơn để in</a>
                        </div>
                    )}
                </Modal>
            )}
        </div>
    );
}
function FilterChip({ active, onClick, children, color }) {
    return (
        <button onClick={onClick} className="flex-1 text-center text-sm px-4 py-2 rounded-full border whitespace-nowrap"
                style={{ borderColor: active ? (color || INK) : LINE, background: active ? (color || INK) : "transparent", color: active ? "#fff" : INK }}>
            {children}
        </button>
    );
}

// Ô lọc dạng tìm kiếm gõ-để-gợi-ý — dùng cho các bộ lọc có thể có rất nhiều lựa chọn (khách hàng, sản phẩm...)
function FilterSearchSelect({ options, value, onChange, placeholder }) {
    const [query, setQuery] = useState("");
    const [open, setOpen] = useState(false);
    const selected = options.find((o) => o.id === value);
    const q = query.trim().toLowerCase();
    const filtered = (q ? options.filter((o) => o.label.toLowerCase().includes(q)) : options).slice(0, 8);

    return (
        <div className="relative">
            <input
                value={open ? query : (selected ? selected.label : "")}
                onChange={(e) => { setQuery(e.target.value); if (!e.target.value) onChange(""); }}
                onFocus={() => { setQuery(""); setOpen(true); }}
                placeholder={placeholder}
                className="w-full border rounded-sm py-1.5 pl-2 pr-6 text-sm" style={{ borderColor: LINE }}
            />
            {selected && !open && (
                <button type="button" onClick={() => { onChange(""); setQuery(""); }} title="Bỏ chọn" className="absolute right-1.5 top-1/2 -translate-y-1/2 opacity-40 hover:opacity-100"><X size={12} /></button>
            )}
            {open && (
                <button type="button" onClick={() => setOpen(false)} title="Đóng" className="absolute right-1.5 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-100" style={{ color: INK }}><X size={13} /></button>
            )}
            {open && (
                <div className="absolute z-50 mt-1 w-full max-h-52 overflow-y-auto rounded-sm shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    {value && (
                        <button type="button" onClick={() => { onChange(""); setQuery(""); setOpen(false); }} className="w-full text-left px-2.5 py-1.5 text-xs opacity-50 hover:bg-black/5" style={{ borderBottom: `1px dashed ${LINE}` }}>✕ Bỏ lọc</button>
                    )}
                    {filtered.length === 0 && <div className="px-2.5 py-2 text-xs opacity-50">Không tìm thấy</div>}
                    {filtered.map((o) => (
                        <button key={o.id} type="button" onClick={() => { onChange(o.id); setQuery(""); setOpen(false); }}
                                className="w-full text-left px-2.5 py-1.5 text-sm hover:bg-black/5 truncate" style={{ color: o.id === value ? BLUE : INK }}>
                            {o.label}
                        </button>
                    ))}
                </div>
            )}
        </div>
    );
}

/* ---------------- Reports ---------------- */

function ActivityLog({ log, accounts }) {
    const [filterUser, setFilterUser] = useState("");
    const filtered = filterUser ? log.filter((l) => l.userId === filterUser) : log;

    return (
        <div>
            <div className="flex items-center justify-between mb-5 gap-3 flex-wrap">
                <p className="text-sm opacity-60">{filtered.length} hoạt động — ghi lại mọi thao tác quan trọng để admin theo dõi (đóng vai trò như thông báo).</p>
                <select value={filterUser} onChange={(e) => setFilterUser(e.target.value)} className="text-sm border rounded-sm py-1.5 px-2" style={{ borderColor: LINE }}>
                    <option value="">Tất cả nhân sự</option>
                    {accounts.map((a) => <option key={a.id} value={a.id}>{a.fullName}</option>)}
                </select>
            </div>
            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "70vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 640 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        {["Thời gian", "Nhân sự", "Vai trò", "Hoạt động", "Chi tiết"].map((h, hi) => (
                            <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap sticky top-0" style={{ color: INK, opacity: 0.6, background: "#fff", zIndex: 2, boxShadow: `0 1px 0 0 ${INK}` }}>{h}</th>
                        ))}
                    </tr>
                    </thead>
                    <tbody>
                    {filtered.map((l) => (
                        <tr key={l.id} style={{ borderBottom: `1px dashed ${LINE}` }}>
                            <td className="px-3 py-2.5 whitespace-nowrap opacity-70" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{formatDateTime(l.at)}</td>
                            <td className="px-3 py-2.5 whitespace-nowrap" style={{ color: INK }}>{l.userName}</td>
                            <td className="px-3 py-2.5 whitespace-nowrap">
                  <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: l.role === "admin" ? `${BRASS}1A` : l.role === "ctv" ? `${RUST}1A` : `${BLUE}15`, color: l.role === "admin" ? BRASS : l.role === "ctv" ? RUST : BLUE }}>
                    {ACCOUNT_ROLES.find((r) => r.id === l.role)?.label || l.role}
                  </span>
                            </td>
                            <td className="px-3 py-2.5 whitespace-nowrap font-medium" style={{ color: INK }}>{l.action}</td>
                            <td className="px-3 py-2.5 opacity-70">{l.detail}</td>
                        </tr>
                    ))}
                    {filtered.length === 0 && <tr><td colSpan={5} className="text-center py-10 opacity-50">Chưa có hoạt động nào.</td></tr>}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

function monthLabel(month) {
    const [y, m] = (month || "").split("-");
    return m ? `Tháng ${parseInt(m, 10)}/${y}` : month;
}

function Plans({ plans, setPlans, orders, purchaseOrders, products, employeeNames }) {
    const [sub, setSub] = useState("sales");
    const [editing, setEditing] = useState(null);
    const [form, setForm] = useState({});

    const categories = [...new Set(products.map((p) => p.category).filter(Boolean))].sort();

    const openNew = () => { setForm({ month: todayISO().slice(0, 7), targetValue: "", note: "", sellerName: "", scope: "company", targetCategory: "", targetProductId: "", metric: "revenue" }); setEditing({}); };
    const openEdit = (p) => { setForm({ ...p }); setEditing(p); };
    const submit = () => {
        if (!form.month || !form.targetValue) return;
        if (form.scope === "category" && !form.targetCategory) { alert("Vui lòng chọn nhóm hàng."); return; }
        if (form.scope === "product" && !form.targetProductId) { alert("Vui lòng chọn sản phẩm."); return; }
        const payload = {
            month: form.month, targetValue: Number(form.targetValue), note: form.note, sellerName: sub === "sales" ? (form.sellerName || "") : "",
            scope: form.scope || "company", targetCategory: form.scope === "category" ? form.targetCategory : "", targetProductId: form.scope === "product" ? form.targetProductId : "",
            metric: form.metric === "qty" ? "qty" : "revenue",
        };
        if (editing.id) setPlans((prev) => prev.map((p) => (p.id === editing.id ? { ...p, ...payload } : p)));
        else setPlans((prev) => [...prev, { id: uid(), type: sub, ...payload }]);
        setEditing(null);
    };
    const remove = (id) => setPlans((prev) => prev.filter((p) => p.id !== id));

    const list = plans.filter((p) => p.type === sub).sort((a, b) => (a.month < b.month ? 1 : -1));

    // Xu hướng 6 tháng gần nhất — chỉ tính các kế hoạch chung toàn công ty, theo giá trị (không giao riêng, không theo SP/nhóm hàng) để biểu đồ có ý nghĩa tổng thể.
    const trend = useMemo(() => {
        const now = new Date();
        const months = [];
        for (let i = 5; i >= 0; i--) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
            const companyPlans = plans.filter((p) => p.type === sub && p.month === key && !p.sellerName && p.scope === "company" && p.metric !== "qty");
            const target = companyPlans.reduce((s, p) => s + p.targetValue, 0);
            const actual = companyPlans.length > 0 ? planActual(companyPlans[0], orders, purchaseOrders, products) : 0;
            months.push({ label: `T${d.getMonth() + 1}/${String(d.getFullYear()).slice(2)}`, target, actual });
        }
        return months;
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [plans, sub, orders, purchaseOrders, products]);

    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };

    const scopeLabel = (p) => p.scope === "product" ? (products.find((x) => x.id === p.targetProductId)?.name || "Sản phẩm đã xoá")
        : p.scope === "category" ? `Nhóm: ${p.targetCategory}` : "Toàn công ty";
    const fmtValue = (p, n) => p.metric === "qty" ? `${(n || 0).toLocaleString("vi-VN")} sản phẩm` : vnd(n);

    return (
        <div>
            <div className="grid grid-cols-2 gap-2 mb-5 max-w-md">
                <button onClick={() => setSub("sales")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: sub === "sales" ? INK : LINE, background: sub === "sales" ? INK : "transparent", color: sub === "sales" ? "#fff" : INK }}>
                    Kế hoạch bán hàng
                </button>
                <button onClick={() => setSub("purchase")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: sub === "purchase" ? INK : LINE, background: sub === "purchase" ? INK : "transparent", color: sub === "purchase" ? "#fff" : INK }}>
                    Kế hoạch nhập hàng
                </button>
            </div>

            <div className="p-5 rounded-sm mb-6" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <h4 className="text-sm uppercase tracking-wider mb-4" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>
                    Xu hướng 6 tháng — mục tiêu so với thực tế {sub === "sales" ? "(toàn công ty, theo giá trị)" : "(theo giá trị)"}
                </h4>
                <ResponsiveContainer width="100%" height={220}>
                    <BarChart data={trend} barCategoryGap="28%">
                        <CartesianGrid vertical={false} strokeDasharray="3 3" stroke={LINE} />
                        <XAxis dataKey="label" tick={axisTick} axisLine={{ stroke: LINE }} tickLine={false} />
                        <YAxis tick={axisTick} axisLine={false} tickLine={false} width={52} tickFormatter={(v) => (v >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                        <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                        <Bar dataKey="target" name="Mục tiêu" fill={LINE} radius={[6, 6, 0, 0]} maxBarSize={26} />
                        <Bar dataKey="actual" name="Thực tế" fill={sub === "sales" ? FOREST : BLUE} radius={[6, 6, 0, 0]} maxBarSize={26} />
                    </BarChart>
                </ResponsiveContainer>
            </div>

            <div className="flex items-center justify-between mb-5">
                <p className="text-sm opacity-60">{list.length} kế hoạch — mục tiêu {sub === "sales" ? "bán hàng" : "nhập hàng"} theo tháng, có thể theo giá trị hoặc số lượng, toàn công ty hoặc theo nhóm hàng/sản phẩm</p>
                <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}><Plus size={15} /> Thêm kế hoạch</button>
            </div>

            {list.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-12">Chưa có kế hoạch nào.</p>
            ) : (
                <div className="space-y-4">
                    {list.map((p) => {
                        const actual = planActual(p, orders, purchaseOrders, products);
                        const pct = p.targetValue > 0 ? Math.round((actual / p.targetValue) * 100) : 0;
                        const over = actual >= p.targetValue;
                        const isCurrentMonth = p.month === todayISO().slice(0, 7);
                        const statusLabel = over ? "Đạt mục tiêu" : pct >= 80 ? "Sắp đạt" : isCurrentMonth ? "Đang trong tháng" : "Chưa đạt";
                        const statusColor = over ? FOREST : pct >= 80 ? BRASS : isCurrentMonth ? BLUE : RUST;
                        return (
                            <div key={p.id} className="p-4 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                                <div className="flex items-start justify-between gap-3 mb-3">
                                    <div>
                                        <div className="flex items-center gap-2 flex-wrap">
                                            <p className="font-medium" style={{ color: INK, fontFamily: "'Fraunces', serif" }}>{monthLabel(p.month)}</p>
                                            {p.sellerName ? (
                                                <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: `${PURPLE}1A`, color: PURPLE }}>{p.sellerName}</span>
                                            ) : (
                                                <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: PAPER, color: INK, opacity: 0.6 }}>Toàn công ty (chưa giao riêng)</span>
                                            )}
                                            <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: `${BLUE}1A`, color: BLUE }}>{scopeLabel(p)}</span>
                                            <span className="text-[11px] px-2 py-0.5 rounded-full font-medium" style={{ background: `${statusColor}1A`, color: statusColor }}>{statusLabel}</span>
                                        </div>
                                        {p.note && <p className="text-xs opacity-50 mt-1">{p.note}</p>}
                                    </div>
                                    <div className="flex gap-1.5 shrink-0">
                                        <button onClick={() => openEdit(p)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                                        <button onClick={() => remove(p.id)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                                    </div>
                                </div>
                                <div className="flex justify-between items-baseline text-sm mb-1.5">
                                    <span className="opacity-60">Thực tế: <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK, fontWeight: 600 }}>{fmtValue(p, actual)}</span></span>
                                    <span className="opacity-60">Mục tiêu: <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{fmtValue(p, p.targetValue)}</span></span>
                                </div>
                                <div className="h-2.5 rounded-full overflow-hidden" style={{ background: PAPER }}>
                                    <div className="h-2.5 rounded-full transition-all" style={{ width: `${Math.min(100, pct)}%`, background: over ? FOREST : BRASS }} />
                                </div>
                                <p className="text-right text-xs mt-1 font-medium" style={{ color: over ? FOREST : BRASS }}>{pct}% đạt mục tiêu</p>
                            </div>
                        );
                    })}
                </div>
            )}

            {editing !== null && (
                <Modal title={editing.id ? "Sửa kế hoạch" : `Thêm ${sub === "sales" ? "kế hoạch bán hàng" : "kế hoạch nhập hàng"}`} onClose={() => setEditing(null)}>
                    <Field label="Tháng">
                        <input type="month" className={inputCls} style={{ borderColor: LINE }} value={form.month} onChange={(e) => setForm({ ...form, month: e.target.value })} />
                    </Field>

                    <Field label="Phạm vi mục tiêu">
                        <div className="grid grid-cols-3 gap-2">
                            {[["company", "Toàn công ty"], ["category", "Theo nhóm hàng"], ["product", "Theo sản phẩm"]].map(([id, label]) => (
                                <button key={id} type="button" onClick={() => setForm({ ...form, scope: id })}
                                        className="px-2 py-2 rounded-sm text-xs border text-center"
                                        style={{ borderColor: form.scope === id ? INK : LINE, background: form.scope === id ? INK : "transparent", color: form.scope === id ? "#fff" : INK }}>
                                    {label}
                                </button>
                            ))}
                        </div>
                    </Field>

                    {form.scope === "category" && (
                        <Field label="Nhóm hàng">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.targetCategory || ""} onChange={(e) => setForm({ ...form, targetCategory: e.target.value })}>
                                <option value="">— Chọn nhóm hàng —</option>
                                {categories.map((c) => <option key={c} value={c}>{c}</option>)}
                            </select>
                        </Field>
                    )}
                    {form.scope === "product" && (
                        <Field label="Sản phẩm">
                            <FilterSearchSelect options={products.map((pr) => ({ id: pr.id, label: `${pr.name} (${pr.code})` }))} value={form.targetProductId} onChange={(v) => setForm({ ...form, targetProductId: v })} placeholder="Gõ tên hoặc mã sản phẩm…" />
                        </Field>
                    )}

                    <Field label="Đơn vị mục tiêu">
                        <div className="grid grid-cols-2 gap-2">
                            <button type="button" onClick={() => setForm({ ...form, metric: "revenue" })} className="px-3 py-2 rounded-sm text-sm border text-center"
                                    style={{ borderColor: (form.metric || "revenue") === "revenue" ? INK : LINE, background: (form.metric || "revenue") === "revenue" ? INK : "transparent", color: (form.metric || "revenue") === "revenue" ? "#fff" : INK }}>
                                Giá trị (đ)
                            </button>
                            <button type="button" onClick={() => setForm({ ...form, metric: "qty" })} className="px-3 py-2 rounded-sm text-sm border text-center"
                                    style={{ borderColor: form.metric === "qty" ? INK : LINE, background: form.metric === "qty" ? INK : "transparent", color: form.metric === "qty" ? "#fff" : INK }}>
                                Số lượng (SL)
                            </button>
                        </div>
                    </Field>

                    <Field label={form.metric === "qty" ? "Mục tiêu số lượng (SL)" : sub === "sales" ? "Mục tiêu doanh thu (đ)" : "Mục tiêu giá trị nhập (đ)"}>
                        {form.metric === "qty" ? (
                            <input type="number" min={0} className={inputCls} style={{ borderColor: LINE, fontFamily: "'IBM Plex Mono', monospace" }} value={form.targetValue} onChange={(e) => setForm({ ...form, targetValue: e.target.value })} />
                        ) : (
                            <MoneyInput className={inputCls} style={{ borderColor: LINE }} value={form.targetValue} onChange={(v) => setForm({ ...form, targetValue: v })} />
                        )}
                    </Field>
                    {sub === "sales" && (
                        <Field label="Giao cho nhân viên/CTV (không bắt buộc)" hint="Để trống nếu đây là mục tiêu chung toàn công ty">
                            <select className={inputCls} style={{ borderColor: LINE }} value={form.sellerName || ""} onChange={(e) => setForm({ ...form, sellerName: e.target.value })}>
                                <option value="">— Toàn công ty —</option>
                                {employeeNames.map((n) => <option key={n} value={n}>{n}</option>)}
                            </select>
                        </Field>
                    )}
                    <Field label="Ghi chú"><input className={inputCls} style={{ borderColor: LINE }} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
                    <button onClick={submit} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>{editing.id ? "Lưu thay đổi" : "Thêm kế hoạch"}</button>
                </Modal>
            )}
        </div>
    );
}

// Xếp hạng doanh số nhân viên/CTV — phục vụ tính hoa hồng, thưởng KPI. Lọc theo khoảng ngày, xem/xuất chi tiết từng sản phẩm đã bán của mỗi người.
function SalesRanking({ orders, products, accounts }) {
    const [dateFrom, setDateFrom] = useState("");
    const [dateTo, setDateTo] = useState("");
    const [detailSeller, setDetailSeller] = useState(null);

    const inRange = (o) => {
        const d = o.createdAt.slice(0, 10);
        if (dateFrom && d < dateFrom) return false;
        if (dateTo && d > dateTo) return false;
        return true;
    };
    const allOrdersInRange = orders.filter((o) => inRange(o));
    const relevantOrders = allOrdersInRange.filter((o) => o.status !== "cancelled");

    const sellerNames = [...new Set([
        ...accounts.filter((a) => a.active).map((a) => a.fullName).filter(Boolean),
        ...allOrdersInRange.map((o) => o.seller).filter(Boolean),
    ])];

    const productBreakdown = (sellerOrders) => {
        const map = {};
        sellerOrders.forEach((o) => {
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                if (!map[it.productId]) map[it.productId] = { name: p?.name || "?", code: p?.code || "", qty: 0, revenue: 0 };
                map[it.productId].qty += it.qty;
                map[it.productId].revenue += orderLineTotal(it);
            });
        });
        return Object.values(map).sort((a, b) => b.revenue - a.revenue);
    };

    const ranking = sellerNames.map((name) => {
        const acc = accounts.find((a) => a.fullName === name);
        const sellerOrders = relevantOrders.filter((o) => o.seller === name);
        const allSellerOrders = allOrdersInRange.filter((o) => o.seller === name);
        const totalOrderCount = allSellerOrders.length;
        const cancelledCount = allSellerOrders.filter((o) => o.status === "cancelled").length;
        const returnedCount = allSellerOrders.filter((o) => (o.returns || []).length > 0).length;
        const revenue = sellerOrders.reduce((s, o) => s + orderCalc(o).payable, 0);
        const qty = sellerOrders.reduce((s, o) => s + o.items.reduce((s2, it) => s2 + it.qty, 0), 0);
        return {
            name, role: acc?.role || "", orderCount: sellerOrders.length, totalOrderCount, cancelledCount, returnedCount,
            cancelRate: totalOrderCount > 0 ? (cancelledCount / totalOrderCount) * 100 : 0,
            returnRate: totalOrderCount > 0 ? (returnedCount / totalOrderCount) * 100 : 0,
            qty, revenue, orders: sellerOrders,
        };
    }).filter((r) => r.totalOrderCount > 0).sort((a, b) => b.revenue - a.revenue);

    const maxRevenue = Math.max(1, ...ranking.map((r) => r.revenue));

    const exportRanking = () => {
        if (ranking.length === 0) { alert("Không có dữ liệu để xuất."); return; }
        const rows = ranking.map((r, i) => ({
            "Hạng": i + 1, "Nhân viên / CTV": r.name, "Vai trò": ACCOUNT_ROLES.find((x) => x.id === r.role)?.label || "—",
            "Số đơn": r.orderCount, "Tổng SL bán": r.qty, "Doanh số": r.revenue,
            "Tổng đơn tạo (kể cả huỷ)": r.totalOrderCount, "Số đơn huỷ": r.cancelledCount, "Tỷ lệ huỷ": `${r.cancelRate.toFixed(1)}%`,
            "Số đơn có đổi trả": r.returnedCount, "Tỷ lệ đổi trả": `${r.returnRate.toFixed(1)}%`,
        }));
        exportExcel(`XepHangBanHang_${todayISO()}`, [{ name: "Xếp hạng", rows }]);
    };
    const exportSellerDetail = (r) => {
        const rows = productBreakdown(r.orders).map((p) => ({ "Mã VT": p.code, "Tên sản phẩm": p.name, "SL bán": p.qty, "Doanh số": p.revenue }));
        exportExcel(`ChiTietBanHang_${r.name}_${todayISO()}`, [{ name: r.name.slice(0, 28), rows }]);
    };

    return (
        <div>
            <div className="flex items-end justify-between gap-3 mb-5 flex-wrap">
                <div className="flex items-end gap-3 flex-wrap">
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Từ ngày</span>
                        <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Đến ngày</span>
                        <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    {(dateFrom || dateTo) && (
                        <button onClick={() => { setDateFrom(""); setDateTo(""); }} className="text-xs underline opacity-60 hover:opacity-100 mb-1.5">Xoá lọc</button>
                    )}
                </div>
                <button onClick={exportRanking} className="flex items-center gap-1.5 text-xs px-3 py-2 rounded-sm border shrink-0" style={{ borderColor: FOREST, color: FOREST }}>
                    <FileSpreadsheet size={13} /> Xuất Excel xếp hạng
                </button>
            </div>

            {ranking.length === 0 ? (
                <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu bán hàng trong khoảng thời gian này.</p>
            ) : (
                <div className="p-5 sm:p-6 rounded-sm mb-6" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <h4 className="text-sm uppercase tracking-wider mb-5" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Biểu đồ xếp hạng doanh số</h4>
                    <ResponsiveContainer width="100%" height={Math.max(220, ranking.length * 42)}>
                        <BarChart data={ranking} layout="vertical" margin={{ left: 8, right: 24 }}>
                            <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                            <XAxis type="number" tick={{ fontSize: 12, fill: INK }} axisLine={false} tickLine={false} tickFormatter={(v) => (v >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                            <YAxis type="category" dataKey="name" tick={{ fontSize: 12, fill: INK }} axisLine={false} tickLine={false} width={140} />
                            <Tooltip formatter={(v, n) => (n === "revenue" ? [vnd(v), "Doanh số"] : [v, n])} labelFormatter={(label) => label}
                                     contentStyle={{ fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" }}
                                     cursor={{ fill: PAPER }} />
                            <Bar dataKey="revenue" radius={[0, 8, 8, 0]} maxBarSize={26}>
                                {ranking.map((_, i) => <Cell key={i} fill={i === 0 ? BRASS : BLUE} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </div>
            )}

            {ranking.length === 0 ? null : (
                <div className="space-y-3">
                    {ranking.map((r, i) => (
                        <div key={r.name} className="p-4 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                            <div className="flex items-center justify-between gap-3 mb-2 flex-wrap">
                                <div className="flex items-center gap-2.5 min-w-0">
                  <span className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-semibold shrink-0"
                        style={{ background: i === 0 ? `${BRASS}25` : PAPER, color: i === 0 ? BRASS : INK }}>{i + 1}</span>
                                    <div className="min-w-0">
                                        <p className="font-medium truncate" style={{ color: INK }}>{r.name}</p>
                                        <p className="text-[11px] opacity-50">{ACCOUNT_ROLES.find((x) => x.id === r.role)?.label || "—"} · {r.orderCount} đơn · {r.qty} sản phẩm</p>
                                        <div className="flex gap-1.5 mt-1 flex-wrap">
                      <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: r.cancelRate >= 15 ? `${RUST}1A` : PAPER, color: r.cancelRate >= 15 ? RUST : INK, opacity: r.cancelRate >= 15 ? 1 : 0.6 }}>
                        Huỷ {r.cancelRate.toFixed(1)}% ({r.cancelledCount}/{r.totalOrderCount})
                      </span>
                                            <span className="text-[10px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider" style={{ background: r.returnRate >= 15 ? `${RUST}1A` : PAPER, color: r.returnRate >= 15 ? RUST : INK, opacity: r.returnRate >= 15 ? 1 : 0.6 }}>
                        Đổi trả {r.returnRate.toFixed(1)}% ({r.returnedCount}/{r.totalOrderCount})
                      </span>
                                        </div>
                                    </div>
                                </div>
                                <div className="flex items-center gap-2 shrink-0">
                                    <span className="text-sm font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(r.revenue)}</span>
                                    <button onClick={() => setDetailSeller(r)} className="text-xs px-2.5 py-1 rounded-sm border" style={{ borderColor: LINE, color: INK }}>Chi tiết</button>
                                </div>
                            </div>
                            <div className="h-2 rounded-full" style={{ background: PAPER }}>
                                <div className="h-2 rounded-full" style={{ width: `${(r.revenue / maxRevenue) * 100}%`, background: i === 0 ? BRASS : FOREST }} />
                            </div>
                        </div>
                    ))}
                </div>
            )}

            {detailSeller && (
                <Modal title={`Chi tiết bán hàng — ${detailSeller.name}`} onClose={() => setDetailSeller(null)} size="lg">
                    <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
                        <p className="text-sm opacity-60">{detailSeller.orderCount} đơn · {detailSeller.qty} sản phẩm · Doanh số {vnd(detailSeller.revenue)}</p>
                        <button onClick={() => exportSellerDetail(detailSeller)} className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-sm border" style={{ borderColor: FOREST, color: FOREST }}>
                            <FileSpreadsheet size={12} /> Xuất Excel
                        </button>
                    </div>
                    <div className="rounded-sm overflow-auto" style={{ border: `1px solid ${LINE}`, maxHeight: "50vh" }}>
                        <table className="w-full text-sm">
                            <thead style={{ background: PAPER, position: "sticky", top: 0 }}><tr>
                                <th className="text-left px-3 py-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                                <th className="text-right px-3 py-2 text-xs uppercase tracking-wider opacity-60">SL bán</th>
                                <th className="text-right px-3 py-2 text-xs uppercase tracking-wider opacity-60">Doanh số</th>
                            </tr></thead>
                            <tbody>
                            {productBreakdown(detailSeller.orders).map((p, i) => (
                                <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                    <td className="px-3 py-2" style={{ color: INK }}>{p.name}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {p.code}</span></td>
                                    <td className="px-3 py-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.qty}</td>
                                    <td className="px-3 py-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(p.revenue)}</td>
                                </tr>
                            ))}
                            {productBreakdown(detailSeller.orders).length === 0 && <tr><td colSpan={3} className="text-center py-8 opacity-40">Chưa có sản phẩm nào.</td></tr>}
                            </tbody>
                        </table>
                    </div>
                </Modal>
            )}
        </div>
    );
}

// Biểu đồ "Hoạt động kinh doanh" — doanh thu (cột) + lợi nhuận gộp (đường), ghi nhận theo ngày giao hàng thành công.
// Lợi nhuận gộp = (giá bán - giá nhập) x số lượng, cộng dồn theo từng đơn đã giao trong khoảng thời gian chọn.
// Công nợ phải thu tổng hợp — tổng công nợ toàn hệ thống, top khách nợ nhiều nhất, và công nợ theo ngày phát sinh đơn (30 ngày gần nhất).
function DebtOverviewReport({ orders, customers }) {
    const chartColors = [INK, BRASS, FOREST, BLUE, RUST, "#8B6FB5"];
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };
    const moneyTick = (v) => (Math.abs(v) >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`);

    const debtByCustomer = useMemo(() => {
        return customers.map((c) => {
            const debt = orders.filter((o) => o.customerId === c.id && o.status !== "cancelled").reduce((s, o) => s + Math.max(0, orderCalc(o).remaining), 0);
            return { name: c.name, debt };
        }).filter((c) => c.debt > 0).sort((a, b) => b.debt - a.debt);
    }, [orders, customers]);
    const totalDebt = debtByCustomer.reduce((s, c) => s + c.debt, 0);
    const top10Debt = debtByCustomer.slice(0, 10);

    // Công nợ theo ngày phát sinh đơn (30 ngày gần nhất) — số dư còn lại tính tại thời điểm hiện tại, gộp theo ngày đơn được tạo.
    const debtByOriginDate = useMemo(() => {
        const days = [];
        for (let i = 29; i >= 0; i--) {
            const d = new Date(); d.setDate(d.getDate() - i); d.setHours(0, 0, 0, 0);
            days.push({ key: d.toISOString().slice(0, 10), label: `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`, debt: 0 });
        }
        const map = {}; days.forEach((d) => { map[d.key] = d; });
        orders.filter((o) => o.status !== "cancelled").forEach((o) => {
            const b = map[o.createdAt.slice(0, 10)];
            if (!b) return;
            b.debt += Math.max(0, orderCalc(o).remaining);
        });
        return days;
    }, [orders]);

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-5" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Công nợ phải thu</h4>

            <div className="grid grid-cols-2 gap-4 mb-6 max-w-md">
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng công nợ phải thu</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(totalDebt)}</p>
                </div>
                <div>
                    <p className="text-xs opacity-50 mb-1">Số khách đang nợ</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{debtByCustomer.length}</p>
                </div>
            </div>

            <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Top 10 khách nợ nhiều nhất</p>
            {top10Debt.length === 0 ? <p className="text-sm opacity-50 text-center py-10">Không có khách hàng nào đang nợ.</p> : (
                <ResponsiveContainer width="100%" height={Math.max(200, top10Debt.length * 34)}>
                    <BarChart data={top10Debt} layout="vertical" margin={{ left: 8, right: 24 }}>
                        <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                        <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={moneyTick} />
                        <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={140} />
                        <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                        <Bar dataKey="debt" radius={[0, 8, 8, 0]} maxBarSize={22}>
                            {top10Debt.map((_, i) => <Cell key={i} fill={i === 0 ? RUST : BRASS} />)}
                        </Bar>
                    </BarChart>
                </ResponsiveContainer>
            )}

            <div className="mt-7">
                <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Công nợ theo ngày phát sinh đơn (30 ngày gần nhất)</p>
                <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={debtByOriginDate} barCategoryGap="30%">
                        <CartesianGrid vertical={false} strokeDasharray="3 3" stroke={LINE} />
                        <XAxis dataKey="label" tick={axisTick} axisLine={{ stroke: LINE }} tickLine={false} interval={2} />
                        <YAxis tick={axisTick} axisLine={false} tickLine={false} width={52} tickFormatter={moneyTick} />
                        <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                        <Bar dataKey="debt" fill={RUST} radius={[4, 4, 0, 0]} maxBarSize={18} />
                    </BarChart>
                </ResponsiveContainer>
            </div>
        </div>
    );
}

// Dòng tiền thu/chi theo ngày — Thu = các lần ghi nhận thanh toán từ khách (trừ hoàn tiền), Chi = các đơn nhập hàng đã đánh dấu thanh toán cho NCC.
function CashFlowReport({ orders, purchaseOrders }) {
    const [preset, setPreset] = useState("30d");
    const [customFrom, setCustomFrom] = useState("");
    const [customTo, setCustomTo] = useState("");

    const { from, to } = useMemo(() => {
        const now = new Date(); now.setHours(23, 59, 59, 999);
        const startOf = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
        switch (preset) {
            case "7d": { const f = new Date(now); f.setDate(f.getDate() - 6); return { from: startOf(f), to: now }; }
            case "14d": { const f = new Date(now); f.setDate(f.getDate() - 13); return { from: startOf(f), to: now }; }
            case "30d": { const f = new Date(now); f.setDate(f.getDate() - 29); return { from: startOf(f), to: now }; }
            case "3m": { const f = new Date(now); f.setMonth(f.getMonth() - 3); return { from: startOf(f), to: now }; }
            case "6m": { const f = new Date(now); f.setMonth(f.getMonth() - 6); return { from: startOf(f), to: now }; }
            case "custom": {
                const f = customFrom ? startOf(new Date(customFrom)) : startOf(now);
                const t = customTo ? (() => { const x = new Date(customTo); x.setHours(23, 59, 59, 999); return x; })() : now;
                return { from: f, to: t };
            }
            default: return { from: startOf(now), to: now };
        }
    }, [preset, customFrom, customTo]);

    const { buckets, totalIn, totalOut } = useMemo(() => {
        const days = Math.max(1, Math.round((to - from) / 86400000) + 1);
        const byMonth = days > 31;
        const list = [];
        if (!byMonth) {
            for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
                list.push({ key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`, label: `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`, "Thu": 0, "Chi": 0 });
            }
        } else {
            const cursor = new Date(from.getFullYear(), from.getMonth(), 1);
            const end = new Date(to.getFullYear(), to.getMonth(), 1);
            while (cursor <= end) {
                list.push({ key: `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`, label: `T${cursor.getMonth() + 1}/${String(cursor.getFullYear()).slice(2)}`, "Thu": 0, "Chi": 0 });
                cursor.setMonth(cursor.getMonth() + 1);
            }
        }
        const map = {}; list.forEach((b) => { map[b.key] = b; });
        const keyOf = (d) => byMonth ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}` : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
        orders.forEach((o) => {
            (o.payments || []).forEach((p) => {
                const d = new Date(p.date);
                if (d < from || d > to) return;
                const b = map[keyOf(d)];
                if (!b) return;
                b["Thu"] += p.type === "hoan" ? -p.amount : p.amount;
            });
        });
        purchaseOrders.forEach((po) => {
            if (!po.paid || !po.paidAt) return;
            const d = new Date(po.paidAt);
            if (d < from || d > to) return;
            const b = map[keyOf(d)];
            if (!b) return;
            b["Chi"] += poNetTotal(po);
        });
        return { buckets: list, totalIn: list.reduce((s, b) => s + b["Thu"], 0), totalOut: list.reduce((s, b) => s + b["Chi"], 0) };
    }, [orders, purchaseOrders, from, to]);

    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };
    const PRESETS = [
        { id: "7d", label: "7 ngày qua" }, { id: "14d", label: "14 ngày qua" }, { id: "30d", label: "30 ngày qua" },
        { id: "3m", label: "3 tháng qua" }, { id: "6m", label: "6 tháng qua" }, { id: "custom", label: "Tuỳ chọn ngày" },
    ];

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Dòng tiền thu / chi</h4>
            <p className="text-xs opacity-50 mb-4">Thu = tiền khách thanh toán (trừ hoàn tiền) · Chi = đơn nhập hàng đã thanh toán cho NCC</p>

            <div className="flex flex-wrap gap-1.5 mb-4">
                {PRESETS.map((p) => (
                    <button key={p.id} onClick={() => setPreset(p.id)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                            style={{ borderColor: preset === p.id ? INK : LINE, background: preset === p.id ? INK : "transparent", color: preset === p.id ? "#fff" : INK }}>
                        {p.label}
                    </button>
                ))}
            </div>

            {preset === "custom" && (
                <div className="flex items-end gap-3 flex-wrap mb-4">
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Từ ngày</span>
                        <input type="date" value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Đến ngày</span>
                        <input type="date" value={customTo} onChange={(e) => setCustomTo(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                </div>
            )}

            <div className="grid grid-cols-3 gap-4 mb-5 max-w-lg">
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng thu</p>
                    <p className="text-lg font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: FOREST }}>{vnd(totalIn)}</p>
                </div>
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng chi</p>
                    <p className="text-lg font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(totalOut)}</p>
                </div>
                <div>
                    <p className="text-xs opacity-50 mb-1">Chênh lệch</p>
                    <p className="text-lg font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: totalIn - totalOut >= 0 ? FOREST : RUST }}>{vnd(totalIn - totalOut)}</p>
                </div>
            </div>

            <ResponsiveContainer width="100%" height={300}>
                <BarChart data={buckets} barCategoryGap="30%">
                    <CartesianGrid vertical={false} strokeDasharray="3 3" stroke={LINE} />
                    <XAxis dataKey="label" tick={axisTick} axisLine={{ stroke: LINE }} tickLine={false} interval={buckets.length > 20 ? Math.ceil(buckets.length / 15) : 0} />
                    <YAxis tick={axisTick} axisLine={false} tickLine={false} width={56} tickFormatter={(v) => (Math.abs(v) >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                    <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                    <Legend wrapperStyle={{ fontSize: 13, fontFamily: "'Inter', sans-serif" }} />
                    <Bar dataKey="Thu" fill={FOREST} radius={[6, 6, 0, 0]} maxBarSize={22} />
                    <Bar dataKey="Chi" fill={RUST} radius={[6, 6, 0, 0]} maxBarSize={22} />
                </BarChart>
            </ResponsiveContainer>
        </div>
    );
}

// Hàng tồn kho lâu ngày không bán được ("hàng chết") — cảnh báo vốn bị đọng. Ngày bán gần nhất chỉ tính các lượt xuất do bán hàng (mã DH...).
function DeadStockReport({ products }) {
    const [threshold, setThreshold] = useState(60);
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };

    const lastSoldDate = (p) => {
        const soldMoves = (p.movements || []).filter((m) => m.type === "out" && /^DH/.test(m.docNo || ""));
        if (soldMoves.length === 0) return null;
        return soldMoves.reduce((max, m) => (!max || m.date > max ? m.date : max), null);
    };

    const deadStock = useMemo(() => {
        const now = new Date();
        return products.map((p) => {
            const stats = productStats(p);
            if (stats.closingQty <= 0) return null;
            const lastSold = lastSoldDate(p);
            const daysSince = lastSold ? Math.floor((now - new Date(lastSold)) / 86400000) : null;
            return { name: p.name, code: p.code, qty: stats.closingQty, value: stats.closingQty * (p.costPrice || 0), lastSold, daysSince };
        }).filter((p) => p && (p.daysSince === null || p.daysSince >= threshold)).sort((a, b) => b.value - a.value);
    }, [products, threshold]);

    const totalDeadValue = deadStock.reduce((s, p) => s + p.value, 0);

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div className="flex items-center justify-between mb-1 flex-wrap gap-3">
                <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Hàng tồn kho lâu ngày ("hàng chết")</h4>
                <div className="flex gap-1.5">
                    {[30, 60, 90, 180].map((d) => (
                        <button key={d} onClick={() => setThreshold(d)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                style={{ borderColor: threshold === d ? INK : LINE, background: threshold === d ? INK : "transparent", color: threshold === d ? "#fff" : INK }}>
                            &gt; {d} ngày
                        </button>
                    ))}
                </div>
            </div>
            <p className="text-xs opacity-50 mb-4">Sản phẩm còn tồn kho nhưng không bán được (hoặc chưa từng bán) trong khoảng thời gian đã chọn</p>

            <div className="mb-5">
                <p className="text-xs opacity-50 mb-1">Tổng vốn bị đọng</p>
                <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(totalDeadValue)}</p>
            </div>

            {deadStock.length === 0 ? <p className="text-sm opacity-50 text-center py-10">Không có sản phẩm nào tồn đọng quá lâu — tốt!</p> : (
                <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}`, maxHeight: 360, overflowY: "auto" }}>
                    <table className="w-full text-sm" style={{ minWidth: 560 }}>
                        <thead><tr style={{ background: PAPER, position: "sticky", top: 0 }}>
                            <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Tồn kho</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Vốn tồn đọng</th>
                            <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Bán gần nhất</th>
                            <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Không bán</th>
                        </tr></thead>
                        <tbody>
                        {deadStock.map((p, i) => (
                            <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                <td className="py-2 px-2">{p.name}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {p.code}</span></td>
                                <td className="py-2 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.qty}</td>
                                <td className="py-2 px-2 text-right font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{vnd(p.value)}</td>
                                <td className="py-2 px-2 whitespace-nowrap">{p.lastSold ? new Date(p.lastSold).toLocaleDateString("vi-VN") : "Chưa từng bán"}</td>
                                <td className="py-2 px-2 whitespace-nowrap">{p.daysSince !== null ? `${p.daysSince} ngày` : "—"}</td>
                            </tr>
                        ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}

// Vòng quay tồn kho — ước tính số ngày để bán hết tồn kho hiện tại, theo tốc độ bán trung bình trong khoảng thời gian đã chọn.
function InventoryTurnoverReport({ orders, products }) {
    const [periodDays, setPeriodDays] = useState(90);
    const chartColors = [INK, BRASS, FOREST, BLUE, RUST, "#8B6FB5"];
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };

    const turnoverData = useMemo(() => {
        const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - periodDays); cutoff.setHours(0, 0, 0, 0);
        const soldQtyMap = {};
        orders.filter((o) => o.status !== "cancelled" && new Date(o.createdAt) >= cutoff).forEach((o) => {
            o.items.forEach((it) => { soldQtyMap[it.productId] = (soldQtyMap[it.productId] || 0) + it.qty; });
        });
        return products.map((p) => {
            const stats = productStats(p);
            const sold = soldQtyMap[p.id] || 0;
            const dailyRate = sold / periodDays;
            const turnoverDays = (stats.closingQty > 0 && dailyRate > 0) ? stats.closingQty / dailyRate : null;
            return { name: p.name, code: p.code, category: p.category || "Khác", closingQty: stats.closingQty, sold, turnoverDays };
        }).filter((r) => r.closingQty > 0);
    }, [orders, products, periodDays]);

    const byCategory = useMemo(() => {
        const catMap = {};
        turnoverData.forEach((r) => {
            catMap[r.category] = catMap[r.category] || { name: r.category, closingQty: 0, sold: 0 };
            catMap[r.category].closingQty += r.closingQty;
            catMap[r.category].sold += r.sold;
        });
        return Object.values(catMap).map((c) => {
            const dailyRate = c.sold / periodDays;
            const turnoverDays = (c.closingQty > 0 && dailyRate > 0) ? c.closingQty / dailyRate : null;
            return { name: c.name, turnoverDays };
        }).filter((c) => c.turnoverDays !== null).sort((a, b) => b.turnoverDays - a.turnoverDays);
    }, [turnoverData, periodDays]);

    const slowestProducts = [...turnoverData].filter((r) => r.turnoverDays !== null).sort((a, b) => b.turnoverDays - a.turnoverDays).slice(0, 10);
    const neverSoldCount = turnoverData.filter((r) => r.sold === 0).length;
    const PERIODS = [{ d: 30, label: "30 ngày" }, { d: 90, label: "90 ngày" }, { d: 180, label: "180 ngày" }, { d: 365, label: "1 năm" }];

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div className="flex items-center justify-between mb-1 flex-wrap gap-3">
                <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Vòng quay tồn kho</h4>
                <div className="flex gap-1.5">
                    {PERIODS.map((p) => (
                        <button key={p.d} onClick={() => setPeriodDays(p.d)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                style={{ borderColor: periodDays === p.d ? INK : LINE, background: periodDays === p.d ? INK : "transparent", color: periodDays === p.d ? "#fff" : INK }}>
                            {p.label}
                        </button>
                    ))}
                </div>
            </div>
            <p className="text-xs opacity-50 mb-5">Số ngày ước tính để bán hết tồn kho hiện tại, theo tốc độ bán trung bình trong khoảng thời gian đã chọn</p>

            {byCategory.length > 0 && (
                <>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Theo nhóm hàng</p>
                    <ResponsiveContainer width="100%" height={Math.max(160, byCategory.length * 34)}>
                        <BarChart data={byCategory} layout="vertical" margin={{ left: 8, right: 24 }}>
                            <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                            <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={(v) => `${Math.round(v)}d`} />
                            <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={100} />
                            <Tooltip formatter={(v) => [`${Math.round(v)} ngày`, "Vòng quay"]} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                            <Bar dataKey="turnoverDays" radius={[0, 8, 8, 0]} maxBarSize={22}>
                                {byCategory.map((_, i) => <Cell key={i} fill={chartColors[i % chartColors.length]} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </>
            )}

            <p className="text-xs uppercase tracking-wider mb-3 mt-7 opacity-60">10 sản phẩm quay vòng chậm nhất</p>
            {slowestProducts.length === 0 ? <p className="text-sm opacity-50 text-center py-8">Chưa đủ dữ liệu bán hàng trong khoảng thời gian này.</p> : (
                <div className="space-y-1.5">
                    {slowestProducts.map((p, i) => (
                        <div key={i} className="flex items-center justify-between text-sm p-2 rounded-sm" style={{ border: `1px dashed ${LINE}` }}>
                            <span className="truncate" style={{ color: INK }}>{i + 1}. {p.name}</span>
                            <span className="whitespace-nowrap font-medium ml-2" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>~{Math.round(p.turnoverDays)} ngày</span>
                        </div>
                    ))}
                </div>
            )}
            {neverSoldCount > 0 && <p className="text-xs opacity-50 mt-4">Còn {neverSoldCount} sản phẩm còn tồn kho nhưng chưa bán được lần nào trong khoảng thời gian này — xem chi tiết ở mục "Hàng tồn kho lâu ngày" bên trên.</p>}
        </div>
    );
}

// Sản phẩm bị trả lại nhiều nhất — giúp phát hiện sản phẩm lỗi/chất lượng kém.
// Sản phẩm có tỷ lệ bảo hành cao — dựa trên các phiếu bảo hành đã lập, đối chiếu với số lượng đã bán để tính tỷ lệ %.
function WarrantyRateReport({ orders, products, warrantyTickets }) {
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };

    const data = useMemo(() => {
        const soldMap = {};
        orders.filter((o) => o.status !== "cancelled").forEach((o) => {
            o.items.forEach((it) => { soldMap[it.productId] = (soldMap[it.productId] || 0) + it.qty; });
        });
        const map = {};
        warrantyTickets.forEach((t) => {
            t.items.forEach((it) => {
                if (!it.productId) return;
                const p = products.find((x) => x.id === it.productId);
                map[it.productId] = map[it.productId] || { name: p?.name || it.productName || "?", code: p?.code || it.productCode || "", total: 0, pending: 0, confirmed: 0, rejected: 0, refundTotal: 0 };
                map[it.productId].total += it.qty;
                if (t.status === "pending") map[it.productId].pending += it.qty;
                if (t.status === "confirmed" || t.status === "completed") map[it.productId].confirmed += it.qty;
                if (t.status === "rejected") map[it.productId].rejected += it.qty;
                if ((t.status === "confirmed" || t.status === "completed") && t.resolutionType === "refund") map[it.productId].refundTotal += t.refundAmount;
            });
        });
        return Object.entries(map).map(([productId, v]) => {
            const sold = soldMap[productId] || 0;
            return { productId, ...v, sold, rate: sold > 0 ? (v.total / sold) * 100 : null };
        }).sort((a, b) => {
            if (a.rate === null && b.rate === null) return b.total - a.total;
            if (a.rate === null) return 1;
            if (b.rate === null) return -1;
            return b.rate - a.rate;
        });
    }, [orders, products, warrantyTickets]);

    const chartData = data.filter((d) => d.rate !== null).slice(0, 10);
    const totalWarranty = data.reduce((s, d) => s + d.total, 0);
    const totalConfirmedDefect = data.reduce((s, d) => s + d.confirmed, 0);

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Sản phẩm có tỷ lệ bảo hành cao</h4>
            <p className="text-xs opacity-50 mb-5">Tỷ lệ = tổng số lượng yêu cầu bảo hành ÷ số lượng đã bán × 100% · giúp phát hiện sản phẩm lỗi/chất lượng kém qua tỷ lệ khách mang bảo hành</p>

            <div className="grid grid-cols-2 gap-4 mb-6 max-w-md">
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng lượt bảo hành</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{totalWarranty}</p>
                </div>
                <div>
                    <p className="text-xs opacity-50 mb-1">Đã xác nhận lỗi thực sự</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{totalConfirmedDefect}</p>
                </div>
            </div>

            {chartData.length > 0 && (
                <>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Top 10 theo tỷ lệ bảo hành</p>
                    <ResponsiveContainer width="100%" height={Math.max(200, chartData.length * 34)}>
                        <BarChart data={chartData} layout="vertical" margin={{ left: 8, right: 24 }}>
                            <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                            <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={(v) => `${v.toFixed(0)}%`} />
                            <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={140} />
                            <Tooltip formatter={(v) => [`${v.toFixed(1)}%`, "Tỷ lệ BH"]} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                            <Bar dataKey="rate" radius={[0, 8, 8, 0]} maxBarSize={22}>
                                {chartData.map((_, i) => <Cell key={i} fill={i === 0 ? RUST : BRASS} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </>
            )}

            <p className="text-xs uppercase tracking-wider mb-3 mt-7 opacity-60">Chi tiết đầy đủ theo sản phẩm</p>
            {data.length === 0 ? <p className="text-sm opacity-50 text-center py-10">Chưa có phiếu bảo hành nào.</p> : (
                <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
                    <table className="w-full text-sm" style={{ minWidth: 720 }}>
                        <thead><tr style={{ background: PAPER }}>
                            <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">SL đã bán</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">SL yêu cầu BH</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Tỷ lệ</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Đang xử lý</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Đã xác nhận lỗi</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Từ chối BH</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Tổng tiền hoàn</th>
                        </tr></thead>
                        <tbody>
                        {data.map((d) => (
                            <tr key={d.productId} style={{ borderTop: `1px dashed ${LINE}` }}>
                                <td className="py-2 px-2">{d.name}<span className="opacity-40 text-xs" style={{ fontFamily: "'IBM Plex Mono', monospace" }}> · {d.code}</span></td>
                                <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.sold}</td>
                                <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.total}</td>
                                <td className="py-2 px-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: d.rate === null ? INK : d.rate >= 10 ? RUST : d.rate >= 3 ? BRASS : FOREST }}>{d.rate === null ? "—" : `${d.rate.toFixed(1)}%`}</td>
                                <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BRASS }}>{d.pending || "—"}</td>
                                <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", color: RUST }}>{d.confirmed || "—"}</td>
                                <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono', monospace", opacity: 0.6 }}>{d.rejected || "—"}</td>
                                <td className="py-2 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{d.refundTotal > 0 ? vnd(d.refundTotal) : "—"}</td>
                            </tr>
                        ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}

function TopReturnedProductsReport({ orders, products }) {
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };

    const data = useMemo(() => {
        const soldMap = {};
        orders.filter((o) => o.status !== "cancelled").forEach((o) => {
            o.items.forEach((it) => { soldMap[it.productId] = (soldMap[it.productId] || 0) + it.qty; });
        });
        const map = {};
        orders.forEach((o) => {
            (o.returns || []).forEach((r) => {
                (r.returnedItems || []).forEach((it) => { map[it.productId] = (map[it.productId] || 0) + it.qty; });
            });
        });
        return Object.entries(map).map(([productId, qty]) => {
            const p = products.find((x) => x.id === productId);
            const sold = soldMap[productId] || 0;
            return { name: p?.name || "?", code: p?.code || "", qty, sold, rate: sold > 0 ? (qty / sold) * 100 : null };
        }).sort((a, b) => b.qty - a.qty).slice(0, 10);
    }, [orders, products]);

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Sản phẩm bị trả lại nhiều nhất</h4>
            <p className="text-xs opacity-50 mb-5">Giúp phát hiện sản phẩm lỗi/chất lượng kém · Tỷ lệ trả = số lượng trả ÷ số lượng đã bán</p>
            {data.length === 0 ? <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu đổi/trả hàng.</p> : (
                <div className="grid lg:grid-cols-2 gap-6">
                    <ResponsiveContainer width="100%" height={Math.max(200, data.length * 34)}>
                        <BarChart data={data} layout="vertical" margin={{ left: 8, right: 24 }}>
                            <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                            <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} allowDecimals={false} />
                            <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={140} />
                            <Tooltip formatter={(v) => [`${v} SP`, "Đã trả"]} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                            <Bar dataKey="qty" radius={[0, 8, 8, 0]} maxBarSize={22}>
                                {data.map((_, i) => <Cell key={i} fill={i === 0 ? RUST : BRASS} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                    <div className="space-y-1.5 pt-1">
                        {data.map((p, i) => (
                            <div key={i} className="flex items-center justify-between text-sm p-2.5 rounded-sm" style={{ border: `1px dashed ${LINE}` }}>
                                <span className="truncate" style={{ color: INK }}>{i + 1}. {p.name}</span>
                                <span className="opacity-70 whitespace-nowrap ml-2 text-xs">{p.qty} trả / {p.sold} bán{p.rate !== null ? ` · ${p.rate.toFixed(1)}%` : ""}</span>
                            </div>
                        ))}
                    </div>
                </div>
            )}
        </div>
    );
}

// Biên lợi nhuận gộp theo nhóm hàng/sản phẩm — phát hiện sản phẩm bán nhiều nhưng lời ít, và sản phẩm lời nhiều dù bán ít.
function ProfitMarginReport({ orders, products }) {
    const [periodDays, setPeriodDays] = useState(30);
    const chartColors = [INK, BRASS, FOREST, BLUE, RUST, "#8B6FB5"];
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };
    const PERIODS = [{ d: 30, label: "30 ngày" }, { d: 90, label: "90 ngày" }, { d: 180, label: "180 ngày" }, { d: 365, label: "1 năm" }];

    const productData = useMemo(() => {
        const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - periodDays); cutoff.setHours(0, 0, 0, 0);
        const map = {};
        orders.filter((o) => o.status !== "cancelled" && new Date(o.createdAt) >= cutoff).forEach((o) => {
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                if (!p) return;
                map[p.id] = map[p.id] || { name: p.name, code: p.code, category: p.category || "Khác", revenue: 0, profit: 0, qty: 0 };
                const rev = orderLineTotal(it);
                map[p.id].revenue += rev;
                map[p.id].profit += (it.price - (p.costPrice || 0)) * it.qty;
                map[p.id].qty += it.qty;
            });
        });
        return Object.values(map).map((r) => ({ ...r, margin: r.revenue > 0 ? (r.profit / r.revenue) * 100 : 0 }));
    }, [orders, products, periodDays]);

    const byCategory = useMemo(() => {
        const catMap = {};
        productData.forEach((r) => {
            catMap[r.category] = catMap[r.category] || { name: r.category, revenue: 0, profit: 0 };
            catMap[r.category].revenue += r.revenue;
            catMap[r.category].profit += r.profit;
        });
        return Object.values(catMap).map((c) => ({ ...c, margin: c.revenue > 0 ? (c.profit / c.revenue) * 100 : 0 })).sort((a, b) => b.revenue - a.revenue);
    }, [productData]);

    const topByRevenue = [...productData].sort((a, b) => b.revenue - a.revenue).slice(0, 12);

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div className="flex items-center justify-between mb-1 flex-wrap gap-3">
                <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Biên lợi nhuận gộp theo nhóm hàng / sản phẩm</h4>
                <div className="flex gap-1.5">
                    {PERIODS.map((p) => (
                        <button key={p.d} onClick={() => setPeriodDays(p.d)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                style={{ borderColor: periodDays === p.d ? INK : LINE, background: periodDays === p.d ? INK : "transparent", color: periodDays === p.d ? "#fff" : INK }}>
                            {p.label}
                        </button>
                    ))}
                </div>
            </div>
            <p className="text-xs opacity-50 mb-5">Biên lợi nhuận gộp = (giá bán − giá nhập) ÷ doanh thu × 100% · bảng bên dưới sắp theo doanh thu để dễ so sánh "bán nhiều — lời ít"</p>

            {byCategory.length > 0 && (
                <>
                    <p className="text-xs uppercase tracking-wider mb-3 opacity-60">Biên lợi nhuận theo nhóm hàng</p>
                    <ResponsiveContainer width="100%" height={Math.max(160, byCategory.length * 34)}>
                        <BarChart data={byCategory} layout="vertical" margin={{ left: 8, right: 24 }}>
                            <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                            <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={(v) => `${v.toFixed(0)}%`} />
                            <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={100} />
                            <Tooltip formatter={(v) => [`${v.toFixed(1)}%`, "Biên LN gộp"]} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                            <Bar dataKey="margin" radius={[0, 8, 8, 0]} maxBarSize={22}>
                                {byCategory.map((_, i) => <Cell key={i} fill={chartColors[i % chartColors.length]} />)}
                            </Bar>
                        </BarChart>
                    </ResponsiveContainer>
                </>
            )}

            <p className="text-xs uppercase tracking-wider mb-3 mt-7 opacity-60">Sản phẩm bán chạy nhất — kèm biên lợi nhuận</p>
            {topByRevenue.length === 0 ? <p className="text-sm opacity-50 text-center py-8">Chưa có dữ liệu bán hàng trong khoảng thời gian này.</p> : (
                <div className="rounded-sm overflow-x-auto" style={{ border: `1px solid ${LINE}` }}>
                    <table className="w-full text-sm" style={{ minWidth: 520 }}>
                        <thead><tr style={{ background: PAPER }}>
                            <th className="text-left py-2 px-2 text-xs uppercase tracking-wider opacity-60">Sản phẩm</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Doanh thu</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Lợi nhuận gộp</th>
                            <th className="text-right py-2 px-2 text-xs uppercase tracking-wider opacity-60">Biên LN</th>
                        </tr></thead>
                        <tbody>
                        {topByRevenue.map((p, i) => (
                            <tr key={i} style={{ borderTop: `1px dashed ${LINE}` }}>
                                <td className="py-2 px-2">{p.name}</td>
                                <td className="py-2 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{vnd(p.revenue)}</td>
                                <td className="py-2 px-2 text-right whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: p.profit >= 0 ? FOREST : RUST }}>{vnd(p.profit)}</td>
                                <td className="py-2 px-2 text-right font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: p.margin >= 20 ? FOREST : p.margin >= 8 ? BRASS : RUST }}>{p.margin.toFixed(1)}%</td>
                            </tr>
                        ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}

// So sánh kỳ này với kỳ trước — đo tăng trưởng % theo doanh thu, lợi nhuận gộp và số đơn hàng.
function PeriodComparisonReport({ orders, products }) {
    const [mode, setMode] = useState("month"); // month | quarter | year

    const ranges = useMemo(() => {
        const now = new Date();
        if (mode === "quarter") {
            const q = Math.floor(now.getMonth() / 3);
            return {
                curStart: new Date(now.getFullYear(), q * 3, 1), curEnd: now,
                prevStart: new Date(now.getFullYear(), (q - 1) * 3, 1), prevEnd: new Date(now.getFullYear(), q * 3, 0, 23, 59, 59, 999),
                curLabel: "Quý này", prevLabel: "Quý trước",
            };
        }
        if (mode === "year") {
            return {
                curStart: new Date(now.getFullYear(), 0, 1), curEnd: now,
                prevStart: new Date(now.getFullYear() - 1, 0, 1), prevEnd: new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59, 999),
                curLabel: "Năm nay", prevLabel: "Năm trước",
            };
        }
        return {
            curStart: new Date(now.getFullYear(), now.getMonth(), 1), curEnd: now,
            prevStart: new Date(now.getFullYear(), now.getMonth() - 1, 1), prevEnd: new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999),
            curLabel: "Tháng này", prevLabel: "Tháng trước",
        };
    }, [mode]);

    const metricsFor = (start, end) => {
        const list = orders.filter((o) => o.status !== "cancelled" && new Date(o.createdAt) >= start && new Date(o.createdAt) <= end);
        const revenue = list.reduce((s, o) => s + orderCalc(o).payable, 0);
        const profit = list.reduce((s, o) => s + o.items.reduce((s2, it) => { const p = products.find((x) => x.id === it.productId); return s2 + (it.price - (p?.costPrice || 0)) * it.qty; }, 0), 0);
        return { revenue, profit, orderCount: list.length };
    };

    const cur = metricsFor(ranges.curStart, ranges.curEnd);
    const prev = metricsFor(ranges.prevStart, ranges.prevEnd);
    const pct = (c, p) => (p === 0 ? (c > 0 ? 100 : 0) : ((c - p) / p) * 100);

    const MODES = [{ id: "month", label: "Tháng này / Tháng trước" }, { id: "quarter", label: "Quý này / Quý trước" }, { id: "year", label: "Năm này / Năm trước" }];
    const metrics = [
        { key: "revenue", label: "Doanh thu", cur: cur.revenue, prev: prev.revenue, format: vnd },
        { key: "profit", label: "Lợi nhuận gộp", cur: cur.profit, prev: prev.profit, format: vnd },
        { key: "orderCount", label: "Số đơn hàng", cur: cur.orderCount, prev: prev.orderCount, format: (v) => `${v} đơn` },
    ];

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div className="flex items-center justify-between mb-1 flex-wrap gap-3">
                <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>So sánh kỳ này với kỳ trước</h4>
                <div className="flex gap-1.5">
                    {MODES.map((m) => (
                        <button key={m.id} onClick={() => setMode(m.id)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                style={{ borderColor: mode === m.id ? INK : LINE, background: mode === m.id ? INK : "transparent", color: mode === m.id ? "#fff" : INK }}>
                            {m.label}
                        </button>
                    ))}
                </div>
            </div>
            <p className="text-xs opacity-50 mb-5">{ranges.curLabel} (tính đến hôm nay) so với toàn bộ {ranges.prevLabel.toLowerCase()}</p>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {metrics.map((m) => {
                    const change = pct(m.cur, m.prev);
                    const up = change >= 0;
                    return (
                        <div key={m.key} className="p-4 rounded-sm" style={{ border: `1px solid ${LINE}` }}>
                            <p className="text-xs opacity-50 mb-1.5">{m.label}</p>
                            <p className="text-xl font-semibold mb-1" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{m.format(m.cur)}</p>
                            <div className="flex items-center gap-1.5 text-xs">
                <span className="px-1.5 py-0.5 rounded-sm font-medium" style={{ background: up ? `${FOREST}1A` : `${RUST}1A`, color: up ? FOREST : RUST }}>
                  {up ? "▲" : "▼"} {Math.abs(change).toFixed(1)}%
                </span>
                                <span className="opacity-50">so với {m.format(m.prev)}</span>
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
    );
}

function BusinessActivityChart({ orders, products }) {
    const [preset, setPreset] = useState("30d");
    const [customFrom, setCustomFrom] = useState("");
    const [customTo, setCustomTo] = useState("");

    const { from, to } = useMemo(() => {
        const now = new Date(); now.setHours(23, 59, 59, 999);
        const startOf = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
        switch (preset) {
            case "7d": { const f = new Date(now); f.setDate(f.getDate() - 6); return { from: startOf(f), to: now }; }
            case "14d": { const f = new Date(now); f.setDate(f.getDate() - 13); return { from: startOf(f), to: now }; }
            case "30d": { const f = new Date(now); f.setDate(f.getDate() - 29); return { from: startOf(f), to: now }; }
            case "this_month": return { from: new Date(now.getFullYear(), now.getMonth(), 1), to: now };
            case "3m": { const f = new Date(now); f.setMonth(f.getMonth() - 3); return { from: startOf(f), to: now }; }
            case "6m": { const f = new Date(now); f.setMonth(f.getMonth() - 6); return { from: startOf(f), to: now }; }
            case "this_year": return { from: new Date(now.getFullYear(), 0, 1), to: now };
            case "last_year": return { from: new Date(now.getFullYear() - 1, 0, 1), to: new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59, 999) };
            case "custom": {
                const f = customFrom ? startOf(new Date(customFrom)) : startOf(now);
                const t = customTo ? (() => { const x = new Date(customTo); x.setHours(23, 59, 59, 999); return x; })() : now;
                return { from: f, to: t };
            }
            default: return { from: startOf(now), to: now };
        }
    }, [preset, customFrom, customTo]);

    const { buckets, totalRevenue, totalProfit } = useMemo(() => {
        const days = Math.max(1, Math.round((to - from) / 86400000) + 1);
        const byMonth = days > 31;
        const list = [];
        if (!byMonth) {
            for (let d = new Date(from); d <= to; d.setDate(d.getDate() + 1)) {
                list.push({ key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`, label: `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`, "Doanh thu": 0, "Lợi nhuận gộp": 0 });
            }
        } else {
            const cursor = new Date(from.getFullYear(), from.getMonth(), 1);
            const end = new Date(to.getFullYear(), to.getMonth(), 1);
            while (cursor <= end) {
                list.push({ key: `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}`, label: `T${cursor.getMonth() + 1}/${String(cursor.getFullYear()).slice(2)}`, "Doanh thu": 0, "Lợi nhuận gộp": 0 });
                cursor.setMonth(cursor.getMonth() + 1);
            }
        }
        const map = {}; list.forEach((b) => { map[b.key] = b; });
        orders.filter((o) => o.deliveredAt && o.status !== "cancelled").forEach((o) => {
            const d = new Date(o.deliveredAt);
            if (d < from || d > to) return;
            const key = byMonth ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}` : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
            const b = map[key];
            if (!b) return;
            const oc = orderCalc(o);
            b["Doanh thu"] += oc.payable;
            const profit = o.items.reduce((s, it) => { const p = products.find((x) => x.id === it.productId); return s + (it.price - (p?.costPrice || 0)) * it.qty; }, 0);
            b["Lợi nhuận gộp"] += Math.round(profit);
        });
        return { buckets: list, totalRevenue: list.reduce((s, b) => s + b["Doanh thu"], 0), totalProfit: list.reduce((s, b) => s + b["Lợi nhuận gộp"], 0) };
    }, [orders, products, from, to]);

    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };
    const PRESETS = [
        { id: "7d", label: "7 ngày qua" }, { id: "14d", label: "14 ngày qua" }, { id: "30d", label: "30 ngày qua" },
        { id: "this_month", label: "Tháng này" }, { id: "3m", label: "3 tháng qua" }, { id: "6m", label: "6 tháng qua" },
        { id: "this_year", label: "Năm nay" }, { id: "last_year", label: "Năm trước" }, { id: "custom", label: "Tuỳ chọn ngày" },
    ];

    return (
        <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <h4 className="text-sm uppercase tracking-wider mb-1" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Hoạt động kinh doanh</h4>
            <p className="text-xs opacity-50 mb-4">Ghi nhận theo ngày giao hàng thành công · Lợi nhuận gộp = (giá bán − giá nhập) × số lượng</p>

            <div className="flex flex-wrap gap-1.5 mb-4">
                {PRESETS.map((p) => (
                    <button key={p.id} onClick={() => setPreset(p.id)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                            style={{ borderColor: preset === p.id ? INK : LINE, background: preset === p.id ? INK : "transparent", color: preset === p.id ? "#fff" : INK }}>
                        {p.label}
                    </button>
                ))}
            </div>

            {preset === "custom" && (
                <div className="flex items-end gap-3 flex-wrap mb-4">
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Từ ngày</span>
                        <input type="date" value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                    <label className="text-xs">
                        <span className="block opacity-60 mb-1">Đến ngày</span>
                        <input type="date" value={customTo} onChange={(e) => setCustomTo(e.target.value)} className="border rounded-sm py-1.5 px-2 text-sm" style={{ borderColor: LINE }} />
                    </label>
                </div>
            )}

            <div className="grid grid-cols-2 gap-4 mb-5 max-w-md">
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng doanh thu</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{vnd(totalRevenue)}</p>
                </div>
                <div>
                    <p className="text-xs opacity-50 mb-1">Tổng lợi nhuận gộp</p>
                    <p className="text-xl font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#C97F0F" }}>{vnd(totalProfit)}</p>
                </div>
            </div>

            <ResponsiveContainer width="100%" height={320}>
                <ComposedChart data={buckets} barCategoryGap="30%">
                    <CartesianGrid vertical={false} strokeDasharray="3 3" stroke={LINE} />
                    <XAxis dataKey="label" tick={axisTick} axisLine={{ stroke: LINE }} tickLine={false} interval={buckets.length > 20 ? Math.ceil(buckets.length / 15) : 0} />
                    <YAxis tick={axisTick} axisLine={false} tickLine={false} width={56} tickFormatter={(v) => (Math.abs(v) >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                    <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                    <Legend wrapperStyle={{ fontSize: 13, fontFamily: "'Inter', sans-serif" }} />
                    <Bar dataKey="Doanh thu" fill={BLUE} radius={[6, 6, 0, 0]} maxBarSize={48} />
                    <Line type="monotone" dataKey="Lợi nhuận gộp" stroke="#F5A623" strokeWidth={3} dot={{ r: 3, fill: "#F5A623" }} activeDot={{ r: 5 }} />
                </ComposedChart>
            </ResponsiveContainer>
        </div>
    );
}

function Reports({ orders, products, customers, accounts, purchaseOrders, warrantyTickets }) {
    const [sub, setSub] = useState("overview"); // overview | ranking
    const byCategory = useMemo(() => {
        const map = {};
        products.forEach((p) => { const cat = p.category || "Khác"; map[cat] = (map[cat] || 0) + productStats(p).exportedValue; });
        return Object.entries(map).map(([name, total]) => ({ name, total })).sort((a, b) => b.total - a.total);
    }, [products]);

    const topCustomers = useMemo(() => {
        return customers
            .map((c) => ({ name: c.name, total: orders.filter((o) => o.customerId === c.id && o.status !== "cancelled").reduce((s, o) => s + orderCalc(o).payable, 0) }))
            .filter((c) => c.total > 0)
            .sort((a, b) => b.total - a.total)
            .slice(0, 5);
    }, [customers, orders]);

    // Sản phẩm bán chạy nhất theo N tháng gần nhất (tính đến hôm nay), xếp theo số lượng bán — có thể lọc theo 1 nhóm hàng cụ thể.
    const PRODUCT_PERIODS = [{ months: 1, label: "1 tháng" }, { months: 3, label: "3 tháng" }, { months: 6, label: "6 tháng" }, { months: 12, label: "1 năm" }];
    const computeTopProducts = (months, categoryFilter) => {
        const now = new Date();
        const cutoff = new Date(now.getFullYear(), now.getMonth() - months, now.getDate());
        const map = {};
        orders.filter((o) => o.status !== "cancelled" && new Date(o.createdAt) >= cutoff).forEach((o) => {
            o.items.forEach((it) => {
                const p = products.find((x) => x.id === it.productId);
                if (!p) return;
                if (categoryFilter && p.category !== categoryFilter) return;
                map[p.id] = map[p.id] || { name: p.name, code: p.code, qty: 0 };
                map[p.id].qty += it.qty;
            });
        });
        return Object.values(map).sort((a, b) => b.qty - a.qty).slice(0, 10);
    };
    const productCategoryOptions = [...new Set(products.map((p) => p.category).filter(Boolean))].sort();
    const [top10Period, setTop10Period] = useState(1);
    const [catPeriod, setCatPeriod] = useState(1);
    const [topCategory, setTopCategory] = useState(productCategoryOptions[0] || "");
    const top10Products = useMemo(() => computeTopProducts(top10Period, ""), [orders, products, top10Period]);
    const topCategoryProducts = useMemo(() => computeTopProducts(catPeriod, topCategory), [orders, products, catPeriod, topCategory]);

    const stockAging = products
        .map((p) => ({ code: p.code, name: p.name, ton: productStats(p).closingQty, giaTri: Math.round(productStats(p).closingQty * productStats(p).avgCost) }))
        .sort((a, b) => Math.abs(b.giaTri) - Math.abs(a.giaTri));
    const maxStockValue = Math.max(1, ...stockAging.map((s) => Math.abs(s.giaTri)));

    const chartColors = [INK, BRASS, FOREST, BLUE, RUST, "#8B6FB5"];
    const tooltipStyle = { fontFamily: "'Inter', sans-serif", fontSize: 13, border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(31,42,68,0.12)", padding: "8px 12px" };
    const axisTick = { fontSize: 12, fill: INK, fontFamily: "'Inter', sans-serif" };

    return (
        <div className="space-y-6">
            <div className="grid grid-cols-2 gap-2 mb-1 max-w-md">
                <button onClick={() => setSub("overview")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: sub === "overview" ? INK : LINE, background: sub === "overview" ? INK : "transparent", color: sub === "overview" ? "#fff" : INK }}>
                    Tổng quan
                </button>
                <button onClick={() => setSub("ranking")} className="px-3.5 py-2.5 rounded-full text-sm border text-center"
                        style={{ borderColor: sub === "ranking" ? INK : LINE, background: sub === "ranking" ? INK : "transparent", color: sub === "ranking" ? "#fff" : INK }}>
                    Xếp hạng bán hàng
                </button>
            </div>

            {sub === "ranking" ? <SalesRanking orders={orders} products={products} accounts={accounts} /> : (<>
                <BusinessActivityChart orders={orders} products={products} />

                <PeriodComparisonReport orders={orders} products={products} />
                <ProfitMarginReport orders={orders} products={products} />

                <div className="grid lg:grid-cols-2 gap-6">
                    <DebtOverviewReport orders={orders} customers={customers} />
                    <CashFlowReport orders={orders} purchaseOrders={purchaseOrders} />
                </div>

                <div className="grid lg:grid-cols-2 gap-6">
                    <DeadStockReport products={products} />
                    <InventoryTurnoverReport orders={orders} products={products} />
                </div>

                <TopReturnedProductsReport orders={orders} products={products} />
                <WarrantyRateReport orders={orders} products={products} warrantyTickets={warrantyTickets} />

                <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
                        <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Top 10 sản phẩm bán chạy nhất</h4>
                        <div className="flex gap-1.5">
                            {PRODUCT_PERIODS.map((p) => (
                                <button key={p.months} onClick={() => setTop10Period(p.months)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                        style={{ borderColor: top10Period === p.months ? INK : LINE, background: top10Period === p.months ? INK : "transparent", color: top10Period === p.months ? "#fff" : INK }}>
                                    {p.label}
                                </button>
                            ))}
                        </div>
                    </div>
                    {top10Products.length === 0 ? <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu bán hàng trong khoảng thời gian này.</p> : (
                        <div className="grid grid-cols-2 gap-6">
                            <div className="space-y-3.5 pt-1">
                                {top10Products.map((p, i) => (
                                    <div key={i}>
                                        <div className="flex justify-between items-baseline gap-3 mb-1">
                                            <span className="text-sm truncate" style={{ color: INK }}>{i + 1}. {p.name}</span>
                                            <span className="text-sm font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{p.qty} SP</span>
                                        </div>
                                        <div className="h-1.5 rounded-full" style={{ background: PAPER }}>
                                            <div className="h-1.5 rounded-full" style={{ width: `${(p.qty / top10Products[0].qty) * 100}%`, background: chartColors[i % chartColors.length] }} />
                                        </div>
                                    </div>
                                ))}
                            </div>
                            <ResponsiveContainer width="100%" height={280}>
                                <PieChart>
                                    <Pie data={top10Products} dataKey="qty" nameKey="name" cx="50%" cy="50%" outerRadius={100} innerRadius={48} paddingAngle={1}>
                                        {top10Products.map((_, i) => <Cell key={i} fill={chartColors[i % chartColors.length]} />)}
                                    </Pie>
                                    <Tooltip formatter={(v, n) => [`${v} SP`, n]} contentStyle={tooltipStyle} />
                                </PieChart>
                            </ResponsiveContainer>
                        </div>
                    )}
                </div>

                <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <div className="flex items-center justify-between mb-3 flex-wrap gap-3">
                        <h4 className="text-sm uppercase tracking-wider" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Top sản phẩm bán chạy theo nhóm hàng</h4>
                        <div className="flex gap-1.5">
                            {PRODUCT_PERIODS.map((p) => (
                                <button key={p.months} onClick={() => setCatPeriod(p.months)} className="text-xs px-3 py-1.5 rounded-full border whitespace-nowrap"
                                        style={{ borderColor: catPeriod === p.months ? INK : LINE, background: catPeriod === p.months ? INK : "transparent", color: catPeriod === p.months ? "#fff" : INK }}>
                                    {p.label}
                                </button>
                            ))}
                        </div>
                    </div>
                    {productCategoryOptions.length === 0 ? (
                        <p className="text-sm opacity-50 text-center py-14">Chưa có nhóm hàng nào — vào mục Sản phẩm để tạo nhóm hàng trước.</p>
                    ) : (
                        <>
                            <div className="mb-5">
                                <select value={topCategory} onChange={(e) => setTopCategory(e.target.value)} className="border rounded-sm py-1.5 px-2.5 text-sm" style={{ borderColor: LINE }}>
                                    {productCategoryOptions.map((c) => <option key={c} value={c}>{c}</option>)}
                                </select>
                            </div>
                            {topCategoryProducts.length === 0 ? <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu bán hàng cho nhóm hàng này trong khoảng thời gian đã chọn.</p> : (
                                <div className="grid grid-cols-2 gap-6">
                                    <div className="space-y-3.5 pt-1">
                                        {topCategoryProducts.map((p, i) => (
                                            <div key={i}>
                                                <div className="flex justify-between items-baseline gap-3 mb-1">
                                                    <span className="text-sm truncate" style={{ color: INK }}>{i + 1}. {p.name}</span>
                                                    <span className="text-sm font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{p.qty} SP</span>
                                                </div>
                                                <div className="h-1.5 rounded-full" style={{ background: PAPER }}>
                                                    <div className="h-1.5 rounded-full" style={{ width: `${(p.qty / topCategoryProducts[0].qty) * 100}%`, background: chartColors[i % chartColors.length] }} />
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                    <ResponsiveContainer width="100%" height={280}>
                                        <PieChart>
                                            <Pie data={topCategoryProducts} dataKey="qty" nameKey="name" cx="50%" cy="50%" outerRadius={100} innerRadius={48} paddingAngle={1}>
                                                {topCategoryProducts.map((_, i) => <Cell key={i} fill={chartColors[i % chartColors.length]} />)}
                                            </Pie>
                                            <Tooltip formatter={(v, n) => [`${v} SP`, n]} contentStyle={tooltipStyle} />
                                        </PieChart>
                                    </ResponsiveContainer>
                                </div>
                            )}
                        </>
                    )}
                </div>

                <div className="grid lg:grid-cols-2 gap-6">
                    <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                        <h4 className="text-sm uppercase tracking-wider mb-5" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Doanh số xuất theo nhóm hàng</h4>
                        {byCategory.length === 0 ? <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu.</p> : (
                            <ResponsiveContainer width="100%" height={240}>
                                <BarChart data={byCategory} layout="vertical" margin={{ left: 8 }}>
                                    <CartesianGrid horizontal={false} strokeDasharray="3 3" stroke={LINE} />
                                    <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={(v) => (v >= 1000000 ? `${(v / 1000000).toFixed(0)}tr` : `${v / 1000}k`)} />
                                    <YAxis type="category" dataKey="name" tick={axisTick} axisLine={false} tickLine={false} width={90} />
                                    <Tooltip formatter={(v) => vnd(v)} contentStyle={tooltipStyle} cursor={{ fill: PAPER }} />
                                    <Bar dataKey="total" radius={[0, 8, 8, 0]} maxBarSize={28}>
                                        {byCategory.map((_, i) => <Cell key={i} fill={chartColors[i % chartColors.length]} />)}
                                    </Bar>
                                </BarChart>
                            </ResponsiveContainer>
                        )}
                    </div>

                    <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                        <h4 className="text-sm uppercase tracking-wider mb-5" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Khách hàng chi tiêu nhiều nhất</h4>
                        {topCustomers.length === 0 ? <p className="text-sm opacity-50 text-center py-14">Chưa có dữ liệu bán hàng.</p> : (
                            <div className="space-y-3.5 pt-1">
                                {topCustomers.map((c, i) => (
                                    <div key={i}>
                                        <div className="flex justify-between items-baseline mb-1">
                                            <span className="text-sm" style={{ color: INK }}>{i + 1}. {c.name}</span>
                                            <span className="text-sm font-medium" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{vnd(c.total)}</span>
                                        </div>
                                        <div className="h-1.5 rounded-full" style={{ background: PAPER }}>
                                            <div className="h-1.5 rounded-full" style={{ width: `${(c.total / topCustomers[0].total) * 100}%`, background: chartColors[i % chartColors.length] }} />
                                        </div>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>

                <div className="p-5 sm:p-6 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <h4 className="text-sm uppercase tracking-wider mb-5" style={{ color: INK, opacity: 0.55, letterSpacing: "0.06em" }}>Giá trị tồn kho theo mã VT</h4>
                    {stockAging.length === 0 ? <p className="text-sm opacity-50 text-center py-8">Chưa có sản phẩm.</p> : (
                        <div className="space-y-3 max-h-72 overflow-y-auto pr-1">
                            {stockAging.map((s) => {
                                const negative = s.giaTri < 0;
                                return (
                                    <div key={s.code}>
                                        <div className="flex justify-between items-baseline mb-1 gap-3">
                    <span className="text-sm truncate" style={{ color: INK }}>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: BLUE }}>{s.code}</span>
                      <span className="opacity-50"> · {s.name}</span>
                      <span className="opacity-40"> ({s.ton} đvt)</span>
                    </span>
                                            <span className="text-sm font-medium shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace", color: negative ? RUST : INK }}>{vnd(s.giaTri)}</span>
                                        </div>
                                        <div className="h-1.5 rounded-full" style={{ background: PAPER }}>
                                            <div className="h-1.5 rounded-full" style={{ width: `${Math.min(100, (Math.abs(s.giaTri) / maxStockValue) * 100)}%`, background: negative ? RUST : FOREST }} />
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                    )}
                </div>
            </>)}
        </div>
    );
}

/* ---------------- App shell ---------------- */

const HILI_LOGO_SRC = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPAAAAD2CAIAAABnbp3YAABdVUlEQVR42u29a6xtS3YeNL4x5z7n3Hd33364X+62TbdtnMQ4iiCKwUjIEYqcRLEtIlmJrYAsQJGMhAJIIKT8iWTEDyyEUQADgQQUJYFYUQwiMomDjIMJNoaOiduPNnbTtvt9349z9prj40c9ZlXNqjlrzrXWPqftbN907t17rfmoGjUe3xjjG3jz/3peACEFAhGhUJIfiiD8rwgAkuE/KaDQ/RJCgUIolMl9UihMvkz4S8LdAPMdwu81ubElH6KIAOIfJT4hQH9V91gCASEiDLc0/yQijLcX0F3MXQsUuneQxo+JlGvj1yFc3D+ivy3mr8L/EslrSv76xVJDxF05PGN6MfdHusWdN4pD/XLuCu5q8cOImxqezv05PDiKz4Z1ca8bHmz+j/nTDHvjxQnJi/m1Cp+mFzdo+KJ5AYEXKbdfcBLmnoJhgebVSPZahEIKR+Ujob8D02cA0geF/xvESamIxiWlEBDE1WH892yZw9X9fSz9nbugzKcHnDcDfmEZhMvfJPnyvGksD4sXgLCF6UNBzN+CAhGEbeB8kJlvsSS76y7uLz0/VHwHEmGxLRNMzNKeC6wkkjLLWilmYfPIcFYnAIUiAulfh0LMuy6J1kh1htDCJTAvJOZNhYQFZyajFoSFkggcRAtxFrr1kPyc5UeQIvCL6Bc1Kqi4QJy1YrIq7psqHJms8qweBOn6IBMCFBuzUDdIjjhzdRElm/XvRqmblUFYj6Aaw4KiquZKpTYfsfJpmZ6f9HkxPy4EIlq7MnMpLdeneDhUn6q+rY0vFTenYj69CC84rzsAIeYP1CwDsl1IHzn8O9PX53KrmPwb2g+MzNgm2n1ptmZxnG0dG8sS78i4jMBo6VsxtZiVtWayAUFlOxWU2d94PZNhKQGAmmh27LzCMFR2OhimQk3BZgcgebZM5QuiGsuU16wPROYnnD0bEZP5i1WZY+pyGJgcm9lGZzo2M8HZSh75SY56ph0Akkh8PKcCGqefcR2YnE/UDxFZlarUUYQV6j8IabKbTblnJlfseqr0LyYCkbF9sLbUX+abYSE3waYXZ5pMn1LYNqvzSk6zQGQKxjuynA2MpK4wve2NC+X1DVJLWNyS9dNb+9Glg7ryhcXVcJZAz8oRC18Ixa0Kd7/m5VT/ulyR4uwjCjGj5zZ7qokzAN1xVNMTwrVVYn6eIYBglLv6OVMnrX9917VJLgzob9ufuG6Jz81VtV/iAptXTg/Dyjal56p2xrw5bR+/rp/xggu3+RBbC3rcACen1DrMNfq37belcNf+pusWRgrzt3XZ1F3LPNUgJ947ChIsmTnNTku/wJAcj8tQyzBdS0mj/ksIqAEf5IaRpkXvCOn+sBGBXOUHq3dZ+StqC065vK3R2ray+3GTCJOVTfGOIntiCTTe+uIaunBuwE4lfYYOrl0Z6nDlADCdNvzODDFCtmreD7miQPvFoTrYde01a+vJLMDnrDKJVPml/3sJ3cF1mxywf0GqVKEiKmIgXV7i6HNoqaa2NujufOhrqjznPlgB+PZI2BX8nyu+ZWdIefZL2V7be2kLdnybxivpoeIJyATSB6+1TNhngu9Smg/oy71fib5pv5WdPdrstFjn7VBDn+myTYfPA2cPO1jUHRc5LNDrckNVmM2fgYQ80XERYs1xN2HIjqLz3ENKh5tbfm2nNBUnue47YuNGrPkhXLweDp95pgEHquuL7sdD/oUY07ksATFnXY/FrzFfyc6XPirQsYyiOFyIyWMnYioiYi7JwDM2IjiQaXYey01p34ALgcZ2xNN/1BQQ0YC0MzEXC/hszS9agIkMKfniD0lKuFMjc04aU0JNzlyLkiYTV9Df0oWFPx5ZQprMIGKcsbDW/4Jnuxwd8RNjTdBlYvHLB20XCWRjrcglPalUYS6eMIK1JLPk7myjUZwB7woGLwGAGCUCZ4r9y4IEg5OLuW+hXGxOnHDHRoyHt3Ar08H1NM/jja/OcW0b7yuXA3iQKcx97rUTdG67Nhd6d5JnuUBNM3Uk87WCQyOzR9FKBod9LmxbGCCzWMXFBCJ7IsAEwEH31qFLkdY8cgFyZ9tIRr8q/e2h3cQu4WgI5VzRUTl5IjXlXRSj9chT+pknYX+57nIgFugu1BtWFNV6edMT8tq9QbxPMVAaYGpa13i5Q3dFMBGCrmVB5xs9aVvc1NDBW0W+WUi8fB7ZLMmKVO9aPbecpdkWxUKCq/lLmxba9xuct0TBiv6O+oFsFifVlpUHqyBQiXFa29adQ7hKCEh6iFYkKTbAuULcLdb7hLl19NDxmheMIp4QmR4vsz/dgVioOdxdpeR7uojDWEereACiiVBMc9gAKTojeAEdgtBW1LhYntRg4uBllbGoAD/MEC62a4kChIA9sN8Z6mbX51c/DGkLAFddjt9+Pzyk8a9htnHBy8YWyz2K1ieS7lI331mVwfiYZIt9/uTlhKuu3x6Hm8cLCwq8Uu/1BK8px82qpruR5uM49GEXNhXoPCnQDuPu8tlwxbW+dhiw6w5XEeu0bTn3Cn4bamgGsFZgIhOpEN9/v3xbpOedsdnJOutmNhQ9CkTZOkX5jHJMgoOrqBRB8NSH7ZJ6h+W7fJ4QiPm+xxp5pWB5tiNaKxLx25/TVARqgsX5xJ69K79I6mNYD+71ILhZvP9EOOltDYTwI1/hjV9bep3rYhdKfc7AdJ5EH5pzI7eswsO5+8vDEnY30Y8T2epbkK7mfXfSxBfRJ3Wh1+yi6HfMWvqF27XB+Qcu/i7j3a8I4GglWPSKF9mNYhn39pbtsxjIvngNiWFCNRR61blyn7S59ZjYHXiLA1/MWl8jFUitGI08SzH1PJULCgsf6EoQDLPe9zngX7s7aQls3bgvdakVunuQkjIg8Bwka/2LbqcpoELMFZra9YK2q4uy86FhxVcSPrUap4XyDl7nCcehIcQxl+OJSoBBVAQiCj6hPV/7litqEDqx5i6F+MShHEVXens5mH+JiwO9sogmosIhJLWMR3sf9rwUd6YnRISAOiIVT/dRP3iTUGPz4/pb5LXOpVW80kHtuDgSPp1JRCCDZElWuQQGdQGBvvgCsVpwV7CKxDhq6SgiEiC6QMrpBFpDoPfTw1V5JHYLCkMu3VRRQOxLWaCchJawIWKv1ryqzekqcQkQRUr8l6ylXT5vtAc1Oy7Q4TSnnXz0FsfZoODXFjomkTME05WQmmY/Ovfyi0JVqPXFQjMjhYYFLPp3Ax/cdgVw/jrsj0EtFmF7Gtk1Jd3mFuoIy2oL3uTv3aAyKsOAYDhcRY1TxkrHroaJMqFr3XBpaea5LkfCYsjST3CcDHCZIuQMhashfRnDQahQMPAJUgZ0ERVg04zWf9NRAbwwMiXLVuNwJjanO7bu3/jci9uyTs0vbvYYwxf+e64616E8za9GbbkcdxnPnBMU5qR6gViZFJFJBMDAQDm+3LmoNlyZ5uKdTeCkXEWm0JTTkSEjakBhE/Q4B9ZNhTj8O2SR4b8MtnhEXe1Lha+7ZHFl1WX9AJe4pdx6SXgCmpLOoQJzGemQivUq2GJ3k3ezVBOtxVQNamDtXmxA3BUTocD6Nn6DeeUq0EWi3gu+Nudbk7K3AOs6zuU25yDXXjMjZmBKXe+B1+kxec8X1dCZTppXzZBxk6duHMLoiNgVYuUJARP+fF8BzS29TKnTWrvSWZ3bG7nCk9mv2bKBDTm3S+4saghkm+55Z3EwmqT3XaESZNtlbxx+JCMMZlpietdCF4zlHVHbHvOx19yc1fUdJw0kcV7hJub87w30YwH6xDNAXFqx3W2PBv30mcfGpsV1bd2xDiykOQ+IKJiEjx97vpQPHTjWI+E4jDLNY2HEErYkhmE9YQpMpX/ESdsgMmTTMbCD8t43aXcQAFywiX82ybQEzN5VioR6F4pjDSKlsxwQ6v09MpYMHFKOc2juBTpLoIxhrsi5wHPFZmJI/mqZQewQg/FYVJQRXgEef2CqhhlHNUSRdr31zuuo8YLFGT9DGKAVHY7+gkmSIAgxX69IcJ7PYN6Dp5pOjtwysGieof1yM83ViVfYs8171DuRFrMlLSkrRTLlc5Xc/0nfO9O8ifPjbL0Ge5kbitOWkHtr9LWxGskXgtobATqEAEQP9+5FUt8xqLcWOofSzWUz/oCHPGMRxz5aeYcuQeLICKr5EDzMk1MOxEQI7OL+wDnqsP6N7sqZ/WlUpjWPtaZJRK+PQXW5h9nOodSZIwGpg1pezwBKBFanMhztWrtRLlmRWEF4WyepJtwIVtsgCjlIXYnxgQ7PwO6lNOaQSYTkQD7k9LKIgArYTu18MTGvncRKCNQP7R1yLSAz51OrysS8Mp6FO6VmZJX9tm7TGiMsXPVluOAEDn6qJAVi3AETUcTG1nZ2ZHc2yvQa0POKcENkEmdiXG9LoBas4gTVcziM93/yZz/8Ez/9vI5PDQOG4Z5wMCOHW5mmRw/ffv97XvkT38mn9SVwHkP4WGPH+sbUXJvN8IvJ1FMwDPiiTSJW0/RTRDKEq0+Cue4AM1R1VnVNLKUkHRblXRqIZaWzjEnlrrM9pqUwK+nT/e7lge4MLk1fNl+REuJKyd3F+Hkbh/t//xNP/Sd/9X33n3oPdBRRV+lGnMTs0aO3ftdH/+Ef/8Mv6L0vTQREE73IXQ0BuLAEt0LN3oEMPhJg1CVwlKixcXO1HLkGevqhomwUnzB38uF5IGdURMPkTGMuGMFpkWS8IDMVWJnJwrabHlFZish4aT10pleKqlUsZF64zEv53TIbhvHp+0+9eO/+e2XGM0AYJ6M+fXPvafDVVVf+cWnlzTWx7uugNWgvv5euoddrvHVcuLcapy27aNTcqF+ktEQsQMMql/YC0emZfuZD4DoO3T2rSi+xf835AxW7D2wi05ABcgMZFYP5XlSQQqoAOtyM9+77aFoOzoW5WqkNGhs244CArPjWxbt4ql0uWet68jnOu9B8MiSygaj5GlAs7I4HKzhXSroLDhJI+TuHDTQl4QAOfQ0PcjmbdR0w3l/tANKgMgzlPrpsomK4Ge9DNEJIdzApohugQBOMi4ZmVaABRdhQilEeBWyoOIGbCivJ5C8wF5KxXT8GswCFE2EBUHfOh0bvIgLtM8N6rjfJ7dRPj0jMPvTufWrMxF3V3BQfwaY0VCjqGNu+bJhrXttXhwxSVMQoE3MpATzWqTog5GyRlFP1C+uez8OVsDINc7OvWHOVKpV8Q/o6pC33Pp0439j7lOVH21Y3K7cK6JO7i2ZJgTDKOEgsRSgcYCAoMAQnmGJJC5L2+HuBuV39iASyPaxMPWx3hr7hLj264tpfUAsCSLjqClcMZoyDr5EEk7tud/TZdnY5BQSgSI4ULkpZ0era6RuE+kW96DJNG0W27p84XezFMR1DkSIeofPXXJHpwvLsaahzzO2Yx+LIJrPhuSMp2Me0jlj7L1iqpVZrxi7nif7/QA5mqgUWhQTKcFVPoVakcv3tWsrOA09uBMpYnyucWk5Atvxmz2PN2rjL2jKmy6uJoC/Ur2uAIP0OOm0dlz0Q0sIU6szLFMokqXmChIy9EV0DlcNRMW/QKmAAy8TKHYbySGCg7QxCz6zlsmjOg5hailGoMJmmiWwyAntujUwFMvWCdipkthIWQUiz4q318vy8bKQuoORt0NDby5g1JXAoyJDC4F1LNIYnLc0r0JibiyB/8PStiepH4mX1rGjEIM2XlKDJkRAfcbxMmtB3K6Rtf+n0NCXnfqj1QvCNevxVDBjKQf3cMpWBmJZ+NkmRU0D+KLXJOq6LDBhocH0Z4dW0bhdn9IgNXMkF/7Pf7G5kEWITBY0k13r9fdFmrkQzO4ZkmsBiXpWHhblUMiFxtSiMhq8+yKG02PMdW/IJh4JbSBwwITCEn/niGNEAwSRIKmzqr+mWdgha3TpzAJfR0FCLANljB3fjcLxUU4ajzaIdt+pohPIY80Xb4B7Ho6jgmV2OwrhLIIkLHU1yDut+b1O9BZq81CKg4VkFyQwv4FNbDCmuVd3k1LlG8C1mWMLwt9Khr8VjLYcbVRdYzwsK04ULyqPUDrEFzatDoODyYwFxQlVI97+yOd+zxEPmNOqixC0y6jEd3MR0+aPrycw5S1bZkpXWLBMAhFpNzcC4MKMnCYlRIDNgpW4lchFJSPypYEU/rfW3FqWNoTwodWyDSiqZ3V1QavD+K5EAFUlpV9K9H3Q3Y9JAw9mJDphPK3Lbh+YiouElNXQA3FhYvRQ9gSfVdGQa0+pEllyTqmZGvKNNo0i75Ca3kViiVBNXyXrF36fjvDpof4HsrHoqJo1yWVuJMD8TrIslEuyWXeJ7VCXNBz5/PE1cpzlR48C5gL2CRpSjeWc0GjGqm89uNle4wKRrIAEjvLtcntkEoFugnQfmdS0Z0Y3Atq2AhbiFCxhoIlRkEDG6087B9ztIrDrInMe5TBEuyTf5SXO1HETmWIMaWWjrOloARagdZplCs1qmAzVUzX3Y1QQPoea1Cq6FHkOYOKoHaGC0MOEYwg/nZyvjGOm8trqnfsTpTphDhTEjtKyobbpSegopmpWuZ2ac8FbNSSMS19fdjoHENiFOgHk/iuodnNhWSpnn5EHFICCi+S48qNBeXjZJWyULQ9FdGrro9UhLCVv6Yh6THM8oNbweArqO9dAVkXXKQUDSiUltGJqGIldP2rRR6REV+eCjO+8coiQDQYyT4GwXEpwnw8RZvPLSsZ3rhLZrHRhVKxvOJ2pOoyw65Qw+28cS2Kg2aMy+FFO33XOTBCeaYgzUK1gTtqTySaqFH2lAzB0uxwYZ4ex9pooZDLAUaJx77AKmRoXXX8bQMBtsGxIvN8Zkg3jGFmvwy5DGyaa1bAVgfspVzV0TFRnpoFbYVg27+ggUJrPLWFSfuQyZCjW8u4VMPF2FiW+Z8mAZIabi4lFO7gPLJB/ZKB6aKzxt3osUWNYCqwk7OwSf3iIDPCjqEAY4BY1YOoQkpWnzmUAFXRcJKfEY4lA0qm1n4rlww5ikIzQxslup7z6NhqbiTMGpOBY+am+qT2JAgEnm1UAGJ2Q+ARN2hAS0dpFM4B6qYbp++m1ZUpPtec7JUjol3m6E8L1CPpZE4kYGfezZIJBys8eiiaS+WcXDswZfZ6yLXZxDsRz3QB7yGy2uthahHEVEk2oWpiPIwIz9FhGHDgI0BVuKmFuhL73A3AaetFQNaf4FmQknKRCoEwIVxN6taaHTkwSZRHJ4ZhmeLeoP9vUUomqCkeAZEKXQAlNUuLHOlpMM3VVpHsGfal9lmLIaZBN8fak/aJCBBLOFLSFfiCLx9OLwFH8sQs1Hxf3MwAiruug5qi01cDrBW2Y+vtggDYbXDbsrzPRDaGzgahKBlNCgCSkPMHPiggULd9QAJoL5EJUuR8grxeZBRj0QEHVS514EQe5/zugKkxVmwgM3b3Qh0GkeQ3pEWQKSMm57msl6hWrG6odr5RNIqmV8zoAyk6ChjTW2XmDqRGSreBA3MvnMFbkmo5GXiAcFDNQimdNZYxXLI05CMLgWo44MO5fJc2fyXPdStXcrCynBarUuffkbQ7wVBihSw9F23h0K2Di5KBL4UkJJtHt/gxoJITzUBpu5hmaNi8UMSJ2db1J2DkNq4dAOlKgkxdoTI72ZKALBxGn2Fs1RfjuHhLAisbxa7eQibd0o6g3Z6krsGHiOtpMjUGHkyXUyasV9g0kJ0GSkAWeKk6QKPqAZHCAiYohzPsv6QQQ9rXQANJW4jWrcolWHgst+E3oaRTinxRIMHMFZoNgovqXFRE5J4fINhJQJHngd5nGIMXiYDdGUhIEqopRJ9TYgOyNtcBX/oRw1WuI51MsP6xBK8+Y60P7SgxUNrYlesYUTAxZxCWPuC2FpnA7zm835A5F3Z+Ji0HJbvzJUcG3oZ1jg94QU4KVLb5WsKRV0M3hE1JDyWnZJMYEmI8ysDYOTMnxK0jDF2fYhIj/mqk3C/qtEtteMLrGwjfMZC1nwQbjIXyL1RzTp/bEk36RCP69M/ReHxDNFzmWFYIRVeCPiMtvBTJnX9+SQmCLHWCBRcFFNUc55HtKh2TxLoItQqQgzCyMYQgdfRyK+QCAAGkyHIs/cJVK7VCNfg9nUbrRn0rAoWWYiZ0gqH1f5PkK/hVgtkWerZNUr6aRZbRNZtSLEl+DkvUmWucJId1tXwCiE/GfuSFtuPRBgwMRHnn8Vx3oir/5Nu+Aw5//CFfL+cOZxLVB4p8X+h8IRZiOfe5X0uIXTTzFyd2clBrBA3elOoEHkaTlXvwXBQJ9KAG0QOP4Xax3ApNLcFV6hNmopL9tFM9LlAsdt1FWSMglcW9FU3IKV5NGGW58UgjJpv2s7+HRUIgyjCr2hoKQ0xWWMkCTbjHLr4u1sHqb5/A8pSEBAdSQlPoE/JVRfgxdPashoWkkRFiJwVzsqIuQwI18wlMc7HmOynNCZbYubqiAIy9Ch/samt8rTDMSmOq7ekGKJKtWwhUlENU/lGIwjXHmhQ7vg9spq6biVahU2sgXOUNtWWmTTLoXaYqkbhDMbI7yb3raKFBObS4GRtC7luGWj7NanezyUWcIgflcnRNGkAytBMcoplE+Coi7W8WUcsS45/HvIGMThS+bUltGB3pXy97IOaZaTGtIfUjIhqYb1HWwLNCaZs7kRRdMtgWAspgop/SmtonCUEZBke+hTSW31nPg8npCKNRwgRSstRmdVA7Lwzqt1P2y5EDFVkZzbdQljDWKy+MiUwpMkC5Z/CBNMrBppRDgyVMU5dbsUgZkYgJE+Zs7+YqZ0CxqSvsIOc3UHXdYsFi8gDKkJxVreN2oW/Cbp5hWAKzJbFEniSqrUrd+4lR+EVApr2i7sXDUXlToXGY1IhGchJ4yqKfEtZWmqDNIoio6F52StsaGVK3YIRptwFrKpC0vYrt8szGgugKTkN2xb0nfHBOgVEzaSwJoUZyEJGqRZGQVWEuJI1jM/z7ScITmiM5o3znRCq9V+hWVjGICqfpFlQr4nU1iURLYOHBzpDTJY1XxRLNWBm/S1AQyopFPY5g3LRr69RX4wY/eEEMp8IAsqSsAnlHK4YBXnWaC8q+7KCkViBs9JlmMrX5zYAGQWse/sejGbgmNFDJD6ISQrAD/TCi3zqfDUb56TV/AJGtQpODo8yTnXXWHRgFEsybufgXJQCGrBFL6morLupnSL1PdaY5rzCzDElDU7CUZ9XUSOIaaONUUGijLQY9bTbBtyGesxhhnu80VL69tTDKbocdZlLiJduii9jCtLfVVgU7JVJt8AZLwzvVBkmJP5uwjB1llfFipMRawcTLWChXFTQzPm8TfyELFGMoWZotb0S6ukKWNiJXEkzFftVKEWSzZ6qCVCkfvBk6ohkprM1RVJ9OIQ3cYohqwaz29nzohFjU+1SIZLTT23yvesZgatZad3BJ6MwRaQaF+kyTx4jJUpvtcIecMMBvqWoFBSwzmvPXJunl30jEbA0bexzPVGuWmJDLY6iy/o0DFAAnkDWqhcV3FSVXtmQ+c35vFwIVe7fg5w5HVdkduqpZi7FaEorjwljhGnyoo7MYPWnToPWCEMiHOrkg+zuncOJ2U5II65wYRJzJtIa/BkERMjYSlhI9svaV3aann6/LWxb6FDZoQVf45z7WY1qnOZgsgobDnNfRJMylBXRXtkAwg5mPprZONi201xc/8zUPgDASRJghiUyc5WTmpjkZPwJgtazFg2IzRn1SVVyJXQMTaJOa9Xgs7jEo7I10pj2AMMSR6A8KT0llhEmWn4gTTpwvk0MKmny0NbBMPItFFFO8cvndskS27Pi2BE0dcIMBPmpKMPkxLfH2ISW2hw55GziHUKYtlr/syoXihpgHSPnqkEM5zNvSKCxGtctZxzjr7XBrN8S+oAcsVNYpjfVw3lG4gZE5iml6Vor0Cz0c1vK6EV6DKOQ3CvtBAaR9oObsMZrVBARE1BCDFwTrGW7USGuZKnaNlHI8GRRCeFDxBZzbrsBom8VXH9duHNOIPuNueNJEyLQF50K0VL0hKbSn5p5inRkPCNS1bKPIPiFgJECe1wjoJ6Sh7IREaPVi3EL6FEllTtLGU6EwMiyWl0ScK4X3Wh22PtAbkuocVi9OrgOF8tbF1h9AV+Wtgi9y9sIP5zhUg0aNo6uncZW2ova/AOl9XIQJ7mIBA86NDN7SMEzP9uSUmo+bZR7tm6rAOmDmt2/uxokg33GOj53CeNAR5ZZHEr7BAhyl3YaJ9YoaXFa7Mp7wGEQjWA+/+NdK1DzFqCYy7xDPLcDtFhB3Ce5iDysaKupSNlwrAUEC1T3j4p1GjFpO93nPNZdA1TMySYdqYA0CKcjSStdDjwHEvPLTmMXUmg8CQiQmMCtmUnJItfLdSqBGYSlLBTp0OXNqkc0tB09USGtmuQDgaIyiARcbZ8/ANeaVLZ6uouKW1KDz4OGpzemyIkLyyCSyj+HKR5ai1aPqS5nAReI2h0bfKA1hBrBsvgZ8Rlw9uz+9I/9zZDSOmga4j5pnNYinpS380QgaY5OXuGDlbaujwLlduwKdG/TgtMKN3TCudsJP6doRw3lYNzoKb0Wi0pUkDU3hGoYe0hc5xxheZ+03PImhHD6MR4njUV9GTkAsXNQhFteQxzOo0zvNKg903xb9/02si0UVzLgeM+KGM3hPYYF0fmDjJVxPUkL7sfdsz3CLhHEZHvcdWC/RoPqZtYnxknomgMOpoqFpMQNAEGbHGq1tJvB5zIDQ0RKK0uMYsteWZnQHtY0SpvlNQvS5V/lewzaESE3maum4ynMpN+uiq5rCaKxT0xBdxNqySi9UXelOyQuiqceNkxTGI+7a0m2aIQJOukT94E6ko4bICIqJFTwtFXuR9kFHGRjpUJlyZd+5zbcxx/czn2kn00QuOB4jUe/RTyShMFyw1Y1lvWwBytpZli3ZzGPhfmDGC5oDeJ+1mB/NNvchNFBSahhjLUGQBrEHBNIXkjZrbQKfMYdnIQDiSJKR3MXqF63BD06HxaQtMmVWHoV1gHXI5487S3TwVc4IVYqnbMuKZyz2zdmLNdujJL5QqE/vPl5mc8VJeFPFxaOCK1doZHWHQlr9uQOvHK0nXbdCbbnPQ5ZcXcv13MgebOF4y+otUE5sju7CNrZEZdkEYbftSGrM9M8cfRIGFQAZhAQivUZnEtsQUZAjChcW6GRXTR6HW3+qoXgmArXu6PrBM7s50X7LxF0uSyo641o3wW35qNDSgm/R2kOu4jQHvBCllQZykmU5+r2XJ8c49ZG33y/ceD+1CO3KVzohVGzSFW86cTI0tuuTiIRsO8RxERz46V1aY169QqkxZY5k6vy+bbOT/Kih06NsnzbOpu7gQZtwOQpIssThjjipOQzjS7m0GmXfXQ8BxcrtuMwXfVNFYPJMqVPGLAJfy0UBUCnFzXO+NQJusgR650Xy8RQKyIDmcjwwr5t3Rw7uRoTCayDM6oztwGiyt37uv+7WdLWMmW5lvo1HysUT6iqpjeaTPk154rsB/rOEvdADr2qSIuanZ3T4KJ/LiuYwWiVuKsa8BqO9SNdVszFEhPZ2r1YJesTVLjhXRhJi68C9x73pp2S0RzkRuQQCQIzoFrpp3hpUXsIEBMi6jWPOZGBeemrmEPc1LKZILVPUTEQBaQZ8o3xYT9aubb6dFPi2VAGQM5WsGFlXBiDkfcCl9bm+loHJZprqAWjcLi2R5fSqAbK8ZVI7P8ItMiQcxhafE6WXnnuhwnaObSejQrQvdrdIdSWpfLETqaiD1z2ZLJzZzruEGL0z7SckEVhJraKpDZ1dzjAQ5D0kibcNsRJaH3RdxWVKKxDZD4WAzPs75dQR5rrxzZLukqqqwGFNaQ5mJI07aN5fmYxlLqOgTa5+R6w62qTMQmZ8+ww9BmH+er5S1uZ7gCkcagJUW4grDgwJdI4jEI9OabpJTH06YDU/Sq7wwG1konDkQR7Kq28wOLfEhXmRRSrrKRgpxsITRcMgx8zfuJWTdS3VotiYcc/z09NUvmwnsmWUWOOS7FiqudhpnLVkBQq+4+D0luccuNGWLlKulSm0bHeEkbx7nu3w2CbYXqC0bqilBwofuz1M5mRV5aoNZ3gntT33GimXV3fVboxBk6nmAqzKCgiNf3WOpVvEwaDBuzl19WrDXydJshSdUFOl+UL/LDgtFjFXOo9T8HZhUEwqa2Z9yDjmehBncYl73GWfr4oZ0Fn8KJ3/QxtOarzczV5l3xpfg2yZ7y3MruZViOS/P6O5AdQ7NZ9nvlb9HPh77t74Lt5i92eCjL6TvL6fBVd39lQlPx8ADSdlL2IY/xIbCRIL8aDp1Ephp6RbWZA+NisEheccZwPWGgf8WOvPeuBF6EhJO9TKd5VDj5Und2v33gZm3GkZpYdg3nWzHTqEdmXqgWXyyTz5FRryamnJe6g9cvnMfWFKeLCjRWG7YT3HFYm5BHzISirPKhx/mEFiedke2pvVU3YMNlpCcmcLNNlvSo7ldqQUt4KgIk1DipPutEo1YaczLEao9MH474irrE6vDwNW1qDtrwHOgNac7eDmxtSKn+1XuuiDMhpSgJLrymnYsm3UOD2FGiOtN/hcadxe7ORXIM5UMXbDINrgJdFxDXZzmGr9hxD7fNVlM2N8idtxVsla5Wa8GDs5hV7qNpN1pFmhWBTijOa3QLtSF9h0OpsfVuRy1vHEMxLJoj0t44ERGzC0pzaIyF+YFTFF4524rdPsNxV+oKMp9VDSDr9DLX3MgEtk+rMpIroO9dVBpg9rrPgv2LNl5oby0PBJFPFWmpSetHcLfqe3KGcBLyuGeOfyX9ePdvHmx8vvfTJaKX36KxJyTvViQqKU1jRYnOtFoBvd8d9deWZy6KCjXWashHNYZSUhoQxhcty0ev5Nrueqk7gfXq/JErNq1tcDQxvLZpt3KXT6tzcNatWXvRINLOFHqX/6IrTt8WVB8AcJ6Gnj11NIJSkqyBX8dqO3+n6/MzVqz/6wfc6HFFc1x0jyMtmqf4F3JX9XLXw/iKM03mP8xhUkgQKn0i0SUskbC4b9U6EjPMd0359+kMFBME70JS83cvu+9yBLBT37HBWNtY6jyZvlcI9Y6PddK2jv4v9p/U9ofjjBzNPBXuugtm0qCjT7X5RYY+5Kuy8xxb3mNfKV6/L8Q/+PrjXcnzPIWJsl08nSqM2hldOugaRslruJVl5IGhB8H83L5Bcp64vLM6rTex5GGcG9457LBbJ1VCpF7FdLQ8cFnErFt+4Myt2C5Z0a3Xx/ZgvuyWQ2UGH2Ql2TiuLpZcxLDWoMe7/wlUlgJuFZG2rpCQ5ux7iwZ1RuuTaWE9D5jdA77s+ryBBdbWKiXdToj1v0q78uTIFCxfDusZV+fJXBeI5fuLIA+1IRH5ImYhCGOnJ1balHO1p8VOLFgrkrHBjSJbiCaX5RqKAviOtv0sJHu1dVYkGBUn10vQsDhgG/vZ7PMltirbd9aU8iDRzBMfg9fWJSlL0GM9ollouxvKQ9XcPzE/ic8KEU5P4rb2kQyOls68bexhVT2fWW3uZjm2ltaSCv3ajWoTHkLNYyBRLoNRzgPO5pGmCSt40mGAJa+ckFTM5ThplxfWX5RxDEZozpFEoedKPUzidkQMad24IF0u2BJwKMCHVUpirXlTkbV/ZWcXIw+TncK+KTC5X8LtVq6NK4MrPYU8qMIOqof4QDtwvJ5xNcdqRdJRZ8SZXtCmk3DWol1km0gSt2EPuBP70guuxsoKb6LXiD506prEKZ/VQ36pxugUZizbpNedTKajhurVi8uKsPBd8/xgXwk58Y2Tk6vuxK3C0b6YuN/o/KJfWx55tQOKpvNb44GVRV8B3nkuSc+Bbv7JxGylqgt7dBvLO9bynNzz4glrLQ6u3mLYe653FnP4VsiXZodsp846XCm5VQiJzq2vHvkxGQxPAaokI4WTnRaq9hUWdQXsS93Cai16fCTOZKcInPfiZ/gu+n/M3NsxUDztdVeQq7Fk2p0vCl7IOKQ+ojYOpuB6RrI8fGg8G5HgWawqoP7jgp3K8pK6Cqz6C3V761INC897PHDaaoXtvdNwa0woVfZVNny9JqDjB1I4NibqddNsvtfuKggx1/+Q652lqQqtQdvtZEugN6Lgl9DNu+bkXEb3V1yOxrVqyix3444aGZsZHhvqONk15BYnTvtkGNSeUI4uATLWuzAOOrjIHm4zXYJ81l8c/leZhScV9heJDFHSzHhF0CZ6t0D24RXl5xEn3xNulwh/txIuq8JeJCOrXZLVZxxzMqRa0zpAS4H+jDSt+62DJPnC+/i4WXPmsiF0MZIlNiQvmxKUgSoXi1mSgVo66dY9g003ec4Q2YYUVC5eIiQUObOzScAEmVaQzwexPIr9gZ4vt/JyHScEggkgOD9eOleJ7F+WTrdkXigm/57x6DFz5ImWPV+2F6BBOTlK0dYvNZYKTwZcGwJ/pCkdhUzUcO6qIomd8HVWkPBBQ42/lL7XMKyRKngsBz6/DVa323u1lh0kpGQBaMgJDx6z+HpxtEIyNq2X9+oifhRbcDZXgDN0XCpzMl0L6T6UowmawlBV2AtXhAni1uN2ZmEodr1kPzxyDeguwGZLXwKsu2rcA+LW5irGovzKglAlGS2w92CcHSFoj0JfB+PWnyE4fjwo0KtroDVzVWxAb8q52uPYT2OQTPktjv91087kymSVlkOpreC4kC2RbYh5MZgmZV/nRd/0Yi0Rm/HfjhrdvQKdemB9rgaWs0GAutdfcbgZe9sij8d2IZh3Kjwx4yzBjEy6kR7S/8fByYWB6iw3eS2ZY08rZPNsLwv9UYfJW0VwLWqxQbLhxW2HIYc5+8r6S/UZtf7RBr/aaU+ifBxyOZAM2+ygQF2QNpCdJzsmMbCiwFoOLMm2g1rX35dzotmyvwnJyTFneR9lesfFU/5Bbt86R7+lRqqUf6XGIttHFdIj5ee6HD1V6kllDCEQxZwEye3UOolUDOH2DtpwYDRQM/RwLKexXwYpcFYZmILMzT3fznZ/dzl2rBzxtqtZpv0Mk8zTONGjzhLUYq5h3HC4F90tPUnsw3Rqe0cjb2+MlXHsWbB5zLP0DnLoOAYIaRe2ahGOND9vy8GhGSvc9QAtwSpkLiBolimQNgFNdf7dYUN3EcaS1l/HXeK/a1fyML/S7FTJKcTX8HoxkWNqOzrxoCsgk4ByY8CYe4EJr6CGrKmm7MBStmDp1gqwR38c0uhxDI8wY1RNZshm96pGljNLF4jo6eVHS32kU6iq1Cghrka7d4EhailouYmdY2G9QnQpMy5CkoJ7srq24zHFc0BLXRkBDbeAGwDfrCwFRHWYufYudJLrF2GQwstlyGs3avd3OT2yOAOLcSfn/KAHm+uXsQyx3TRH0sCh+/e1DxqsjZXgpnxodlB3qTTHw6gAVDEER3nZYEkRVR2HQU2MND9sfFYnGwVfCxuyGpHNOSNeSpRRVWpLVykpIyimIyzduTUcULiC2aXAxd5Du7eItB5NLoPCvdft9Ag3py923LSY18Z2IE8R0UEVUFC1VZnkF8BzEYDCKXkXTzK5Xp4TySa30fQsAW8x0coEhsRxrdzACpnOzUxaVw5BY2063YMRxR4BOGjTxiPUjPAFHvX4g7WtlSgtKwbIJAwmUojqc6LPJeXHSAJmd6gMbi6tmlAEN1DF/fdwfM4HhwmLJuFamgS44fg+6vuJk+itw62hCtE4ebFt65gHQyh/GZ1RzbkiVRQT7K3p9PbJbpkAYVtIfo9oaCHPoVQrvYpLKWazcrIgw7OjMb9g9OKtBfUwzPnePDZzPQbWXjMOkveVvgtksGkNuEoFtn1qsc8/7qcagVD1+Z/4P77qpz/xgg5PTzyZTSKiGFQHASAKmHGappPZpDqaTWa3w6DDzTt/7pPPjsP9pSCCEJXPv/T8v/9fyGAfNXlrMBPoMAzDOEJv3CuZmcX2gOxcUGAhWkttMaP6D2wdNJpDD1V1GMdhIPTRiy/Y7/mG22/86Jdv5AsyPRQ/7fyS0UoV8e0n3boqwcQ+pTkTSO9bomOp75C4Qg2BXh5i7g2/BOMzP/3z7/yRv/6B4eZFO00W6sIGdTKtEiYDmZEGwUlwcqVGNzfPDsMDy5/PnUIM+MLLL/zFH70n8rTII5lU9GYYBlEoPBRC0uwkwgYAJxCNfCk0BruhgVbPSJKTEFAdhsEV+p3skZxee+GZL377H7j/A9/34kfe86nb0+tqlxGUYgRgaWHWp8u1S6sPB/eP5VQ4v/HcWo48Uk4zYUwzW3vJoyA6js/d3Huv3rzTxkfGSURpUAyKIWlgVRJGCh4JJuEgHACYKJFURsXnhAnu3b8/Cu9PBG8wKKCa7ygHmQJvGMO8xYQuGYpgu+fZ53OhGwORpyPZc+7OdKOUmxdft4/893/35V/+1V/5wX/zG77pw588yWvV1Oku0al9/SrC1G+Ktxoid9uCbm1YD9V1ZXSSB1Os6FozgVGyf6JntJUU1PBPcKFsLpUING9B62MyORFGNQMNNKWokEqOIoMLzF1vfeZ7UZQCDpRpgk2AQaADMVDESLOJdkve0mHZEgbQUTyBOkVsFI5C0ISTiimo4CAcIYPQNcs4QsoROgJKihkoKgpAx+He/fvv+38+87v/7H/I33zpo8BY5cY6qggB6vIfcf8sW8HcXw3zP/6Te27J5J8kw7D8SUU++1ZSA88zknFO6nINDVvpuq76ZPMU0YMhaoO6yHB7mh6d3h7xEGbGIVCYqwiht/OkeooQ5CRCyCgYBKJDTICHz5mbijqIiE0PTcydSxfN0FzvltLV18KEk9APBCWFogIFA7+1y8jwRkQEtz6aDJOEQxgb+z5OpCiGm8GUpIqOT/3cJ7/mL/0Pv/yvf9879fYLlEGu0ox+BqPOMUgRO9Tt+rSXw/jJWDlxtOVE6BXNUc7pou6ggU1nYc0gA4RvvP+5z/7jH3gd9543wuxG/K6DMgE2QIZIgqCYJjtNJ7FR5ekvvv6el954UXNYzCAiCnn04Oa19773yzd4/WY46SCT0TiJkKbkwBnANoiJTKpG3hhHSWlqZAImcpAkiw6Pi+bFEqBQofLGW8999kvPc3xepwEy8d67//b/9pnv+2Pvfu8zX6LZFWQMdyDQhTO5Wdm8Tlqeis2KWK8Dx+MZb4u0HaKAajNnqYkj16uEKbDb177nj9x813dA8LZQA7RCmmsc8oOagMAICiExnWy4997/+C/f/Gd/5d79p17wGJFrkAIhmKbTV7/vlR/+cw/fee8zCogYp1s3hI40EVBGckAcqCUGUUeSRNCSCEFJwSQipIojyUTO1wYKTiIUDqoPXrn9wH/0X771N/7ONI3vEOqgT3/xy+/5jd986/0fH2/xtlDRYDqM7dzdhYr9mvvc4K9bj/o6zc2BmIdumvkUYw3q3MHXBucE2Xzs0gn3gQmmhrmm9c2Bvnb2/THdv/f5+/wcIJ6i3EGT5ggJnFmwZC4sAMFTMoyvPHNzI/KicACm2UtyeVLeG4eH7376M8/f+weTAUlTXaBInN8/OithEUGkLF5W5GvSZclTEhiAdz7zq3/yO//Jv/33Xnr10TtFoYK3b+998eUbYMBGHAYRS5jM+n25GqFe9/TOS4KJHuTUDbEN3RFzAnJzGjRzdgkeQjlSptcs+bWyNOzVGzPZBWP1/JQOE3NEQY4kjm5QnlsHutBSgUHANLGRk0g4bTeZZawZTBILSJtb5t9nkQ0WmaOUboqx/xV+R/H2W/fw5fv3buQRICBlsvHt23uCASLgYLCGY3okw4dixMy2+4udF+9xJotbe2qnTVNzSRqDnpeJjZjJbm43CkNWnH0UxIT5eRhC4bIB5suePYkkU/4hCmnm+RBQh3sggIkYwGZ6CCXmI5UC62xXTHJSmdhk7t56CgEqXDQJigzCUaAUZUKIunAWsUnC246r7BqwcXedWWmV9+R5DoeJGHe1+qUpKESfgVt2j/udOabKIxh90BcoQmTJuECQYi5LJ5F3cJZj+i4Dm8zMKFSBlZNMKgVo1umMUbLy5dQ2QVR1ulEHotEjMwJ4wFuW6lkq5FXoimh2/JjLB4Uy2mHug0b0JXmBDAuPhKfm4HzAmllM1jR0N1/M4ojMlRVoXMPlF+qdo1zUA6SLteCVmSfQNp9WAy1rAD7L6TPmfWVz8N+MmupK9UYys2vZALwoKkw0Sgn46wAKlPAV2YQrlZgCuIR25s/TBNe7laItyinQtvlnHcGJQjiKC0hUHZw7yUkBbReFs4Nbp8BbggbsZ4+L6bCa9KIuruNFXKjsl6EEPfD20DfXcs8t+hZuBoMyt4NS5eVID2LHFNQWs8QmSlXWYCFVABZI8E0wic4pyRZMG/+zY6RdM13QWDeFAWaihnEcVFXx6PYRzTAphnKMI6ucY2VjRLliefcp5EIp9FYGZndQWDxKLSeZz/32G3YwydlTpJr7DJFNBhv+3V39IEwcjSBjAPUsTCKV9jCHtUVYZig6OTjc8VCIUW7lnW+eXnjj1Xe8+egp1dOzD157+qlX7g8vQR4Kdb0wffFgFYaGMJDzgBeOXGEV+oXVg5FOkt3phYGeT2kjz47lpUupZerIE/RzIyq8T1UeWZad4Vw4JGEqoU+vQmCpYzZTYicXzCPAQtFzS21kijwXTcC9ns1JR8Ex4Gx2MfuYoSFmogqDjHrvuS+/9uEf+zv3/tb/Onz28089Oj0jcnr2qZtv+cbn/tg//8Hf97u+oLefEUPoPradj3eUbFdyKD9uXhYaMh8wOYvEOcVJPMwFsFYzNTcNLSUFSYkzq/a2iVWFYwU5CW5FJsgN5nKTrFEfyyF5aT9fxzun2ayArJNU2k24ky4eeLezFxhCkPqqW9wlqgRs0Kff/Quf/oZ/788Pf+/nXpjk3bh5VhWkfe6V06d+49Wf+Nnf+t4//Mz3f9fTD/QXSaModpe6Xiw92aYYkGYLVvd1fci1JLHccIxiuwx55pt0NsEvPuZEx8hJZHKzprK7QObGxAVpL/vWGA2PyHWpurx6CBmdgin4jQ7iZd40dy2qQia5eebnf/3r/60ffPBLn/6Y3H8GosAt7VYEw80D3H/6y2+/78//5d949fWHf+Zf/MB9foZiIipid1A+iiTirregL2RKlo7OzrtVht2e+apX7dKdhc3bsKGVtWqVuXrcxPsrh2h+hMAkekqOiG5CgVdZCMXLb3/kz/3w9Iuf+drhwfMDRPgyb3/xvv6fIz9hjz4vdhqhdu/D/82PvffH/pf3DeMLZmrcV3Di10orE0p79vqwRI1SQUBr2AqMSTC+zTTsfZJkHrCEfD43IAIApGE5QGIGgFAz/W791Dfm13hSSaXcCAfGpyiZeTVcN3tBxcx+G943IUOiq7SWVlGXm3BOCmUyoVIdgAfeo6ir9Dt+oC3yvtSUOt2gKs9iDJiOL/yPf+veT//8e+7ff4E0nF795o/90vf80Ycfff/rb75176d+5jd+9MffePn26yAw+fBf+O8+923f8r7nHrw8qiKZl7AuyoUwFpmRLkqkGpyyYBeRkrCFOu7TMlf4ucs+CMTCa/YajS0a/F0OFBZbFUsHusKRTXer+EClU1rw1umFH/9JGt9lnKbT7e/+6k//0J/Fh9/xCw9ffVn05lu/5R/7yIee/sH/9DcfynuB+7/06ef+9//70bf//mfAN6p37vEAL7vFa4cKe6q5l+dmB4YNKcq+l5eNv6fnked6n0s9y5Cke5pgXS492VNV+PStKET39zWKMa1Px1b9ZO0VPMRBNpyc/ZPltzaRX37t5tc/+65xfMowcfrCd3/H6x969h88fOMLGDmMj/TRJ//QH/jc13/150+3b02Gh/biL316Gu/dYM97LX9pqyWyDC5K1rS+mHpfnFWgMuFEr6QLKxUU+U+/Tltu8jrouzr1eEekdanwYHUfLbY79oDQB4TbkfzGxVDom2/i4e29QW9E+GC4/cgHb99++yVguBl1hNotx+HzH37/7e3tGwJRvffWm8M4PuBAnCMAh9Zw5+K7wZudMu0r2c1XISRloi097ftJFxAAE2KrBDmPk0RkZtIthhggdstbBKdd0UEYDUEBbTFJhVnvo/uwulqKOStZjp5m7P4vhoAVkEI+Ekpj7OSfKzYvixoFRvX9hoRM8A6nrU7J9s0vbpF8jQV1GclII4sZTowPPMYB4+iJtk6GV18bMIgCgoFihtPJ7r3xhqgMInqj+tQz94GYJe/JT6SHqcv3cHhFJDkIXXeu3gXdRxrbKMfynG2q2M4WhoaHiZ4nqaFy4TN0xZ/dITmj0KBBRritAunbwAox2kMQc1GAtmK1YiLa8M7n7PkX3jA7QfTEZ3/8p4bh5mPjqJBHgNw8ePHXP/ehn/uV4d7NU6o28PaDXyWQt4ALcONvcs+ebwx135L5iWkqooG1bTuGrUreqvNQd9bbszaKnJ75lqsSsFx4Yo6GMHSGJk+uIhp9uqo7G1OOAqgO/jqCErhyvS9+QFcpZ1VqotpayaKcYxO0VmBw/6RPNcnpuWdf+paPvzE9ekXI8d6z//NPve+//tGveX36/XbzsdPNN33y//unfug/f+rLr3+djveFbz937zO/52O3Mr15QQqR2uAvbhIetUKvBWx3XoC5Pq/k/ANXDkXtYSETqliTq9GTKs2so5mAZ+Wj7daxUNtESdyK+sk8OEDkUhJT3MJExN76zm+//Zt/91ffmN6l+uAkH/oP/uIXfvzv82Mffcdrb+jPfGL8/Etfde/mXSZvP3rrt771n33laz/40nR7K9CCyOZusKnrEs3Ewkg6bFXWx1HF9iGjAEyA24WvXaCYro86eMdbHMmuLpE+XaJQiDKvqwBgyBTConxC6qXheYzhqweLjMmK9grkIKCrWPUNZpAb5yZuu0SzAx96YNiJ41ZqBECV01vf/PVf+Je+6/4P/7e/xgdfTwwyvP/nfvHFn/2FW1IHfWoYQbHT7Zsffd+n/tU/KQ/0sycbQO6Az3Z6SmASXHV8K6hzWRKH6/XOSiWiB0NR7PE07/H7p8Kbt7Jj36gA2YXVyMysjtKA8vKkMB0o0C2mz33/v/D6v/zHvzSefvn20S1Jxf1xeHYcngYgPE2Pvvh1X/WJH/wzj77hQ79CewSZfN32tbfp7J8jLVjr5rLlTHMe+Z1ojmwI5xzS5uXhO0nzbaO/vBWE7+p2XqYw1nkKsQgM0J7ALqvUamf5bCD5QOT0AP/vD/yp4eMfu/9f/ZVPffLXP3jSZ1wulzY9//Rn/9A/88Xv/+5HX/OeT58evWGqM/7XMXFvfcUuMTAumVW3SEiNZyq9kNLGmaFrTn/Y+141aZ2JPtpJrKL15MgOFYNf2xNGkjeUpBpWAWr00rmwt8fkRmoF6/mbO4YdodzenH75u77t5W/6+Df/qX/jC7/50rODqlBOt69/73ef/rXv+RIe/eJ0+yYVIqachy6f42bs+DpFsNsOoDrre2UqUaozfIEsBJBq4UoczLvtYKRTchfxL5qxX0Qu0kLmE3Dre0NYF+GI70KM+ymLYn3KZkc7SUf05Roi6ftX/BRlQiMkH3lIW8SWm0N0lph0qyKQIE2FQpxu3/qtZ/Qdo36QYqSjJZuev/el4fRrp+ktRvT7+txLUmve2+WsuyhlPHbur4FBVj5caVWofphFCLWksOB1+bBYBRxTLZshbo7lwCfwiEMLviLxvY6T0GjTNAE3gKYz8QKmfrRI/2JOxY6SY0JsKdAXJEwIT7MmAtheEdau2eKsMS8fToAy9KoyyYnCS1RvpnyDsjWw2o+fcj7HBRXHDle7pBYAOaVYEzQ0g15IeV0qANh8Le6uhy49Fiz7BUtthApAUNRWlsqNKy5v8Wsu+cRYaTiJV+ln3eMetQ6RDpwEmSJvFVFdG9iNKgBOeFWNDx1vZfriTxx+kZqP9t/HM8RZPUFWLSQKjHKZ6M0Ob+QqSBDImLH2qXyvMxBGEyxoOApoNAImyMpFciXm5yN3SF4/d1ZXMbOrVvCzIgC6rqZZyPRMn2iHFkTiW4BQ6gAMg0JpEyJv4CUc54vgetW+7IrIiYjoeIk7HZ81n3+xSemCKg7ImgcC2pb3txNQ57qt2Bo+luLQ8ymLaRye4fccclWRnlLXL+zHDFyFBPUyWrkrSCBEVA+L8gpIlyQUycVo3GbWfkmB3R0GZdLnWMWw7Ris0B9mNCN7nmp7TZL1CK7+waLnI/pvrkFRn2ZRz3UWJhng0jXYx9Vlw/InIpTrMoGNZ65jsxUZkY8BCV4USz0lc7KxpFlhHjnWy/fyuwf6uNC2DcmIBOJUVpcoJPuIfA44k/PDz+9O0uhS+gqC0EGGgOHhEAx/vjMduDsBeHprR7+kcanOeapM6y3VwdaUooJMupiPmFjpdBNN7+aw5eQCWChLBobFrCN/PTJb8gQjMLJBuDAImcewOXNywWSFdJZlh2pctDR6fmQNYbLGYnFscwReR5ohiswDKSpaLhoZcvHPuU5IVdLGc86fNKZd1W/JGkyGdERx8lnTFOSSrt41z60bnByWnm4LRNm/jrtNsutYSHUSZ7/ayHyu4BmQ3H6fe66MRVw380DR0YrtA2OiDr145S/jxdDNC5xh5EpR0kKOzanlK+rxCfjBzFvACA5F3+iOXI72uqUF3xSZDgydX6YqLxsU1nRKJQgamyJRS7m13OX+IKnA2jjT/aAhk64Th8DaMyRfAejHorZc/MzRmT+QOGqeyNRanHIbx4cL8AeezwPxoIaURuw529Qg/ZR/RX1YdXLkzKDt1pdGM0e9rmFQyS5qrJXTcg7UXjsk8YKsYah3rOZQutBJ51zV5RKR/kJNy9pTznu49mViy5a08ymLU4dYNu0vbmQQZFxjtTvZSgNfdaQtxvmgeEPXXNBdbjri4/m36ZshxxnMX3XEq4YsY1Jce6qsEgGLiUlBfpJ8T6s3UdbueMSkLtKCND+fM429Vqz2XsdvpVYzHxXimoGNfT2wVdNXNSwrVW6bnv3h1T43KOydduyGNAg68TFwKfF9JU1h9kPBYZ6mNlbQpc6XytvAe7X/MIyLpsDQnLskpk1kcW8CZe+oP1yCreHw1y8boS0Gb85BjJ4d7Kfm2ELmedGSSNY902SCVfJUaZVINkQrKj0Hp6S+sjsh3oeASRwSR+RmkTnDctnslI7I6LSkcTCpioL0rWViQvWEGROMHBqdxXtNQTWHJSu+LQIGI8oAOcWZdv1S6I+fi11KyZnrwN3Ujr3nc6H+ESvQlo7HuYmVLWWQQXIVgoU2ul6OBOg0BPs1zTlTtbvvEepgESn7hWIGIyaETE9Vy+7Vzf2xY33dcI6G3mjjR+OhNvt9jmlo7j0rpYdairUlQi1yuEw2OfHVG/k69OAlO0DHA2F5tWjx751895L3XB2w/iKDiAkf0W6ZTFLC8EB1VH1bGpqVGbNea07XESQkmEfhcnVYr/q9Xpv6pjvef+ux/7hLo29iXTMll7obvDcom4siB4dbj8IyDIKJ8pA2xQkvw3gz3jwPvVF1vlJRUxopl2qHObN4qFaN74lVfVVHctCv7uxeyYqOuTCkUJ9bud7c+FL0mVTVFrNzzgRxallD+KSAn3BsoFEKUx6RQg1IMK59uLz9wGQCVRM8NNogo4hO0/3Pv/Q0bt6l06sGm04mrrie8QWpMcwNtU3qmt+oMFJP4GBCkUFpENiurkjlKNCshRcCnWSK7PZFecIG4rFNK4dkfgh6ZHcF9nGE3QcY/JnNlJTLToE+ErmvXRyyUqXhA0argL+Xsqd5rw1jUEgOzz5zeuGFt37zJSpAcDq946/9zVd/7zf+3o++793T6dXpNM2AkAgwjMMN1Zeb042jI2kn4yQioxogA6dx/MIN3jZAxEDsecII2WyiZttrs96NuvjrhYaWL8iOO4PCwJYiXWzsVQ2azBhWOTSRd0UThv/nCvzJUGwXveR5dPoZeZcDxjHItgoePf/MG//Ex1/45Ke+gHtfTZFhfPoXf/1D/8q//bkPvPcDtOenZO64KMfh5ubmxgxGUVUA4PTo0dun00M38GgcpgH3nxm++O/86fsf/8ivTdNDcAy0juxB98NJy8BNW8jlXn7yTW9tU3ek8rP6YQrPmLFyWPmG59vIonfi8Js385m4MEK5cwLp5R3EpO0FcqLIwFf/yD/3gf/pJz//+ukjhCon6IPPvfbRz798EsJSdBKqCh1gJ5JQVQUok9k0TY8Ek1MKIx48O/7yq29/Fvwt5W2j+37fTq2r65Wjchdg0e6gcMsHitQFOVyN5TS7vazGADwKHp6hPW3SIdB1pJyBiLVwKWaqDt8dbEKNLsF5S1xDxEu4B8phmt78fd/00vd8x9M/8td+Qx98GAJ1KfFxFA4DPEGOBlomigw30YcwUiGqGAQTYCLjgAGDCnw5UTkGaUuePI6oRZg5zyzAolH2LvnsilHc1QVfeg16vQe6s8C21DHQnZ+/gwzZQDUKlL/2p//Em9/9Bz8tDz95y4emA2RwwbNRRQZgNIGRk5mbUzQRJ8rJYPR1tAIVuQe5ceMclRRMhunYoj+ZzbCHpWi83tMclshzjkSYK8EF6hLiAIEeajxbcYq2271xoopM93E6Pbj5h//uD3zsaz80/aW/8eXf+tKLt3wHZYSIYITPpUlo5hVS55GsYoITQNWBcBM39eHtK9PpLW+VYIdGMiBP/fIJrL7t1yZjzxGu1RpsuKRnWXCJLVIVJcKNXkBfmt6c2ZxhrbxYuB26XdGaMUsR3EIEp9N9/ML3/tEX/+C3vucTv/Dypz6jr70xjMO94f4oPGGep0gzo2eqdrsAIYZxUFUMRj4SDnJ65T3vesXklqvVCs2V1ox4Kq0wuRstfNROYFtDNxnLCawN6+aFtSxYqj2ihGq4BtqFqs7S/SNCCx3z09EX+qz0Tcw1fM1GWj9q0yCQEyfAPveBd3z+w//0feK+DA+G8SnhBHlbxcgYTkhS+AshaPcgI3SiugmiQoOd3pxOpiKEYldYSFCmMMcuAPmBmPjqxBwZV8QumfZQfRUd73U5jroQlfaqqzoeDPjvCrcAL6mbD+gTCkwHpZHTQ+jbsFd4AinUySONoejfU+P5AI7JnHJ4vhLRoFFVMB1Wde0sE6+rnQ8GMNzW0J2h8R6PnunOdHZcnpng0KQ4n4WO7Kj57B8OeczBd+LsodlBJABtoE+hCFQWT+qydkk3/ejaTKKeBY+FK+ksxnhoJBmcu10/0JjQJ12Ms2vGdgNWCk1MJd/W2OrP2YmwaBmYLT/G3rORThqM3At97MIe5oHv+TRJksYSRwIyotSJHchzTt4M17xJsqYk0A3feseHJBUamrHEkjY+Ij5tVAgI7pSru3L76cYHhDUhOg4kk1jQN5uFDhxJprlsv055IyRWcaYW0Mg1XKM+axUyVFCO+pi1np7CQ8oJyyWLRjG98jmIXt+DeYZIFuvoihMiWREFIpYNiGGNia/ayr+7gbQSQTJrJVwXHD9wyMcyeV8TdsZVviTRHyrIkDILhOFGvgGVO7be0tULVdYzm31Bn5LF6BeqknDXHS+EwloFjsi69uu6P1ddmi26j6WQAuxbTxsqPKVsDk3XzMxChQJT+35gYfvHUq0cznUaquJF0KCE2gEySsFGGByMUNTPCw/KQG60ex39clnQ1aF4WRx6Zc5fD7l/BFxBmVKnrp//3XVWs6bRKdUyX8o1s1/L5kts2dbzTOUeWL3iTzOD7a8RCUZwg3vBQRZdENX1GbOjnHGoX3I1+/ZmWY9yJEB0kVX58v4PLnNhQdntwzrKEGo2PbX4Bgu/eQlXXWr1EKGvDXTIq73I/keBjPDlYhbGiDnm2EJRFx7FojmSS/lZ0KrsLAwrWJcZzSiqpUtuQjAK97fODXwn9H1MH+AKY7CcFDOQ1et6vdTKVKQEsmUZtleIzoo/8Wqrt0tTurpEI40ymWvHBCHm5rCsblAHOJkC25dyXVIPse6nNfihebb5WwDJzI/0vnkL+zY1wDksPXovf8MwKoaJc+Nsjm0E0GaJ1aSfNRVELhZWoA6i7sf6UCFr8GQFq8b8DEuadlR3q4N6j4AMwZAJRsHgEEQ3wcgotyIPPRhY3HLORLJpHNtxwmVR7eUgMqcvrljLsQIYXV3Vhzanpf8GqsoNANXJZFpYI+dCWJTuludFKILlnmshEggiDBp125l60mUroyfGXtCzL/4dbZXD9P5ldz1TSj3M2LafpDCI3Ac0BM8GnBJ+pyU/24pM7+3nPcuJrWq9mkCza6L9OVP0GjG45qbtyGt691BA15eaGkARpchwIw/eM9681+w08RYmoQYDmfPrQc4kEMEEEcroCzXhaE5PAhMZhCPlJHIK1MuR1l8DsRijXAcL4sqJAmFzaAqbQRnM8kSKinppw8RQ4CEEOUYKJ/r7gIWNCikdZXxfkYHDMAz330vcF1EZ3APYPD1Ilu7oulIuA7F6hBNAfsoxOkiHuM/jBK+HcpwZ5fDwuSwu4euJ/dwLBvfLoMMXX3nXX/jrz97HN9tpmmyi3FImVVUNg+xFhaCYwJVNC6lC37BH73wrMAnMvBJnDKXm8oQQCPhniVuYlmjQEMFD+A4xKHzvHVXUBCQHGgJjlwkmlVtxbFIE5R4FopF/WCFDmSdydeJGmSdciShU8XD64KsP70GdvAOQ84uTNgDW2NwLyFGm92RoYFNDn8VpsiV/PanvciTFMiO1yajr5oNERt0kq+d0Ej730vM/9COvUT7qci/ESXCLGREcSCUhYk5kSQhH1w0AMXqtqYMYQeMAjpBJcBIOLiktMomYidKz0phmAh18FELknohpkpVwZO+u71UIwAQiHIQDZaK3nwBOjsjPl+YFP0JFSQiHVFUW5JSZSQQh4/jATUR2SzAIZWv44hlNWZn/w8OzOC6PQ/e4Gftiu4uAhgzFxFlI4zwEEJPIMN57p+tagecnM4iG5BiYesLedKuIiXO4MRCEmxlFBOqnojIQIi7ymoskvMMdUDsGJY75S4xeq3rLHAJ7V6JkM5M7EibLCNoFT8aS+yYBsaTuTurSu+PWOx7sCZzvnePQMqWd4j1A4WVfaaUJ/shEnAhgL85K8BbcfRjg4DAGM42pyolFlmAYk4MD0onikGIyeZnTircqYqflkDCZ56FbYqkmyQb9SJmxT9zeBEhiI1SbQ1hAEQ2jLyK7Jxg5jzzo8pt3qEEmz4XAPreK0+6K08ZClfLSA+p6WnyvAXPINu9ei6Cxs2KkGiSxB6Ht+8hmddElkN3FjQAdhqGOJdbchrNzAzhrXq1coOv7bM+klSqjXnCXeh+sCvBdzo86VtiN/YvcyqtvbkqaaYvpvEX9y1VE40wn505RjsNUsBfVOuxIzuEwSrgcfevhinmYaKH4eb3VPu+7/lXU1Qx4CPW6A6UC9MOeLFs/seBYggmyWnDsIxoNcVLTy6kNjFLxWFjP65qbo5j5fBvFvhQRNahpQMpIc03g5iUsgg3R03ZOZHXGkZQONZKgLMt8pDEh8+xfQJRD5Dd5GBXJ7zQtOMg3OXUukZoSSooehcKJ2SHOQmzM3NyzKmeajwTVlEbD9MhopKrDLFsnuTZodxeozPNIROJ5QI/LcYxh+3H/QMQIC+2ikY/eT7QPxZFzIRFBilGgUBjnIwqN40XAmnplFPAcfC78haSphGIQhQzxa8Fk+kcWX8HvCyqWV1MHJkMZCnXC1W0mt6bGSDNTfbO2HQCbx7wncSOEkGk63drtJGIUWyrpimDwrFaqy6McFdRLcHEBDcaFZQFdO7xlbNAIFPdJFyobwDZMxewRb9/i+OikJM0mgCPmCXFz+xIcumc0mP8lZq0c/BXHDTCKKL15ce0/VsMckJThBU44n6WbJDY1+UnoTkFqKJMK9a1IxypaoundXYwxE4OEya8k4bY49jRddIgb2qaFqYRwspNMnB69brdvisADj8JWSICswuTqVbh1qq1FTDn2BRK8mk7dRiHaz1BxkTndfvSrXvu2b/n0cP9LtzyZmZiK3FAGNzrXZ/U8enwCTgGmdeBveX+KuMy2yGB+gF90l6Pdj+5CTOtGtgT3fyZyEpkwtxQ6JEZFBvoua0sMNyuD3nyzySQyaTKg1OQmg859rm8CTvMpmPttScdQo7fqZ/u6d51up1ubJnv40nvf/Tb5FnErEnI0l5bLXb1LHe1kyYdf/5kHW1UZizHXCx9601GJbYGzYiA6+1VD4N0CRpS0oIPcDr5o8uIE0JLJ2l7D0ytd/zQnwa1wDNuGujaKwk31GhecXVFfPDK7pYEWQOLEl9lAgR4GLylAZzpowJSTgELLS1BH4eiymHn1tYahSoDNx3BGqh13pSshdLXgYZxcVO0+5CDNbNCXb+TlQe5RTodnby5xa3De08j1synWTbFE6T27o9mHcjCtYEx7rp4gTzo2XwFfGvGlQRRqkuuo2QvwSpnBuWa+PovKT1drhJBO9Ccs1Adhns/CdMWzDIrvzvUVGOlsuDwbPbsxzqRkvP3+ZOc51cSfHmbrYpn3D4EKnVMMcRl1l+/0uEwwKAoSJkoxlcmugHV05pj3d6CS7CofTdsxUtO6N2rch1stOg/6sF4QHIVQmZj1os8ur8Vubl/pZq4XvJCSqE1mSbXIVIOsIrQ61wELpib4rlxWMuWJffBd3A5mANPSJQlTeVISAzEUar4I0xBCvjlSdGdaE3qJWEwiIs7DvqWNhavdn7QrxTHx7zcGznJGwuPO1oz//O202GmsOSXXULRVOJZ78gmoXaq8CAX0UXzhESMx7sHQYKaeYDIJVDgxacmkmK/E0GzHcgOWKP6IoeRTcr3vApWIiDn6xdrQUaawWuaqswiHi4lVZYVHBrOR6cEI+jrDguGy7qEL6xjaDSlbtuZ4A9LMySfxJTvFqfzPpOs7PgQ7aOx4kEUpiBlFY1AfjKKtxojuJeeWsnS4QZnxWqzWHN/Nw8wksFGUaSumg0aqfAOuKjVfulCjhLQkBNmxyzppGF2ZQLlRsOqbRG7bGEvJgqUcpZlqNgZhiS1yAbUnS0E9PLoFqKUtF/OMsYpyJ70IRWxTHTnJgHLAGFqRGKhGLuUNtxK/hShs2gQAF8RZEvnox/C5YishFbFjrmZSxpxMK9fY4KMpK3opVlLcK63jKcvyOTvbCvoL9p9jExZrDmqrABNtLRaYk/hYGVRTMo27SMDszxmt9892lggzr7jfYoG626RUB0H943jUSsZqnb+4FhRi5mC7iKSunfVkR7Mq1rw6hohGc6PuItND6HXEDpRZsNZgV75vHvyh6CYMHTXtGL/iD/T7ldhyfNmdIlgDZCFJd86lkxb1xDprD+ZzAWP7Ta9T9LyiIbi2qgFq3dVJwS655LEs1+YGLrULlk7kRj8ZumGEx2fTrvkQ20Plik82qMA4t+v06ela33LUTE1OKgcFZPzTKevXwnMVEK2I4shsTKAIGHYVYeZhTG+Y3LOqczoNKUco1iVso6oQvdQROENtXUoRuoRYPt+7HuAWimy86EGqw3AbI7q4iyn48m5ce5jn+rf2qXXWTs4xp63Gw8knvDPq0tp6lw9dW7sdS8YESQJTBxNxhGtsCI0cXRGoDMgr54TzLptmsjUGvVWrjT1SAsS6jp6kT8qSkfj3Cx+xOSljT9dP3lPXBm3yy2YOfXfTA+Z8DHHRZppekCBAYK6DbLwIFFA/ScyR+ZiDlYTrJEuSINN/FSeb2w0gDT+7+TrMOYq6V6C33aK5Pqz7iPFhCrak/Y4+w7jRYrny8vezSd7Ia7vSXSCSePR7bDuX14Ns0lQXt4/fPPY0H49SnQ7YoVkWWdm+MYqHmq/2xlH5QKqK6PQcaPFdGnv0OpOwFcbHPe7N9/2EQXYh0bO+yLxYG1/nBicTH0NdznUAzg0ymuuAA8cue9GHaVFF9n/3WjtyZz8jLSU720WjYeu/B5D6jk5nAJMnIlol8Ktg0iQioWXxnB0EtccHH2b/oSWoEDoGEiekklNkUpzouRIlq9dZDUjQv/hISza2yIhX14R5fYFUlr2yJpfTDmJQ0NKx8CuLNCfDr9v1fZjy66qP1Cncm9MCKi8SM55ErXbijtxK36uVY6BnQkYA5O6A7y5PsE54nsf+sjfI3XWzsLLzNBNIxS5U4GfMJRDsdnmby0SthGULNV/UJ6yg6WVZ38xbWu1nC0V43fnJZXfwOgKAvDGp3fOGfi1DNsnMr+ND67HDo3JHZOZLc9Zgs6xERXMc3dngcMBrXLngEiUg2eaTnV3YfFIOrj6hdfc+bjwVVrXDpUexXMZR9BXcLICzZkGq5fUJNbYEd0Fzs8EscStTfwgRuKw00swDDkUCradHR1KcjxXluuWS6sohy3mcXXu1oOlhMHOIC/QtEbC69FPX4pCGsUo10Wys6hfZwY4Zq9tKU8D0GXRzztqVXNbFjeAYihe9yftRDu45NCRpsW7liv5XWjBYvUvLW9iZGjy3KvJJAFX2XQr+5ysJ5eh4XOZV8JkHth05RT5xVGKy5dcTQp1e9AGikXkAKZ6Qe8AOcqmy1SSDJllUlXR651dKPs+VwX2fvpRby7kSEJ3n57Gn351B0tJaXnhA3QzX37FuWwVYWByInfjGb/OfVCv327FLLxrbz7buQy+9jgOCt3TjEoA2zPatvLklWCmSMdNIbAFTHbAKmsVjGmcylPorIwZIVbUjI/C5HorVMdf0NfN1u9YZwAwnc9miViz7GdhUKbW+bwkryM+dn7O0orNmPLljaNB2pnod30U7vkngQlRRpztyRpdVJ49PQT5eZz22hLV653qw/MNyXxsJvgvl6JHmnAr/WIxSfUP1bBUoQadaexIig9bWjdIGpwPB3F2JEjtXsnePz/ahZ0ldKPtl1+CVFhFAIG9oitAKbed48f2orjsXVJ1Ykm2jYRQWQ7cWD9AeiLbD3b9rPXiJ51na38NeBztOTi8nPHnOUJGmXLGjH23ce3wSUoWu6YtLNT/XyAVCLTdMvezsLjqcuaBiqDXsI9b19ivmI3NrcbgDKg4nXjWJe0rPsZhuekCYsv5qre88uFzwPTJwULmk6Viuugmym9T6bMduOcw9/c80TXi8nOhOpjg/QY/HruG8Vww/Lv1267vPS2roa6Fpq3XASeS9wcc+oxnnrqg8qcBdLbZOmJ7xmLHRi2Ag66QROF+gi4E65z4yWJy0rbg+Y9ar9YcXE7arLUySD2sr/obUicTlMhRtk4qyLbzGKJTPUd7srg1DPVFL0vd4LxX348k52RulVDyooXkRtvY9mIMhhaU3e1yCis2DdOPq/HS0Z8NfY2P2Dsrun/wOqTjo2CkudmAdrg9R9zXJ9g/uvh4RRK3dCMXfuxer3jrQ8XVsVsTvdwb6orq26GCfRYSfqsjzfI96hQ+SuS4VXpn2i8fRAnIeImrrd8QBDX2HWaIdpbqtL2YV3k+YH4w+Ur/jS3eWTGsRb2bqP0Ae+3cH+3cz/XbXkKEx4eXGeiLwKiqqA+sBdM8qpMVJCaUB0XB8sNy5Y9bTJzhr2Fa+MoHy/NK+WY9TvvM6rU74ueyyRy6AXa/Tq8uYOJZxwvSYTCrB2oTDu9bKlzjZmY27wMW3Bb3rjtwpZ0dk+jyPlns+hktftlugK6NwMV6pbmJFW5Os1Q4dtt158CdWdz2rBquBZhw/2I0czWpD2tq6HTZ6d1Z9f4UYceoMZ1HpDeF4wUC187sXXOsaDnpMsT0OFKrNtZxOi7qSWrlecH+EbauUyw0vdB3lEIA0260yW046dX+wr73hLXjE1q1YLbBP0VoeLT02zXeRS+0aNnzsdj242WEZW7ng+JWiz7qc7Wvf8StqkXok79rK+0rXXzknKo+nantuNo49pE9Ou97vkJ+v6Mac1sNrVfbPPz0bkVP8x0Mr+Ee7+49+LvLz/wPajs+khkbXwAAAAABJRU5ErkJggg==";
function HiliLogo({ size = 40, radius = 8 }) {
    return (
        <img src={HILI_LOGO_SRC} alt="Hilitek" width={size} height={size} style={{ borderRadius: radius, display: "block", flexShrink: 0, objectFit: "cover" }} />
    );
}

function LoginScreen({ accounts, onLogin }) {
    const [username, setUsername] = useState("");
    const [password, setPassword] = useState("");
    const [showPassword, setShowPassword] = useState(false);
    const [error, setError] = useState("");

    const submit = async () => {
        const acc = accounts.find((a) => a.username === username.trim().toLowerCase());
        if (!acc || !acc.active) { setError("Tài khoản không tồn tại hoặc đã bị khoá."); return; }
        const ok = await verifyPassword(password, acc.passwordSalt, acc.passwordHash);
        if (!ok) { setError("Sai mật khẩu."); return; }
        setError("");
        onLogin(acc.id);
    };

    return (
        <div className="min-h-screen w-full flex items-center justify-center p-5" style={{ background: PAPER, fontFamily: "'Inter', sans-serif" }}>
            <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');`}</style>
            <div className="w-full max-w-md p-9 rounded-sm" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <div className="flex items-center gap-3.5 mb-2">
                    <HiliLogo size={52} />
                    <h1 style={{ fontFamily: "'Fraunces', serif", color: INK }} className="text-2xl leading-tight">Quản lý bán hàng Hilitek</h1>
                </div>
                <p className="text-xs uppercase tracking-widest mb-8" style={{ color: BRASS }}>đăng nhập để tiếp tục</p>

                <Field label="Tên đăng nhập">
                    <input autoFocus className={inputCls} style={{ borderColor: LINE }} value={username} onChange={(e) => setUsername(e.target.value)}
                           onKeyDown={(e) => { if (e.key === "Enter") submit(); }} placeholder="admin" />
                </Field>
                <Field label="Mật khẩu">
                    <div className="relative">
                        <input type={showPassword ? "text" : "password"} className={inputCls} style={{ borderColor: LINE, paddingRight: 28 }} value={password} onChange={(e) => setPassword(e.target.value)}
                               onKeyDown={(e) => { if (e.key === "Enter") submit(); }} placeholder="••••••••" />
                        <button type="button" onClick={() => setShowPassword((s) => !s)} className="absolute right-0 top-1/2 -translate-y-1/2 opacity-50 hover:opacity-100" style={{ color: INK }}>
                            {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                        </button>
                    </div>
                </Field>
                {error && <p className="text-sm mb-3" style={{ color: RUST }}>{error}</p>}
                <button type="button" onClick={submit} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>Đăng nhập</button>
            </div>
        </div>
    );
}

function Accounts({ accounts, setAccounts, currentUser, onResetTestData }) {
    const [editing, setEditing] = useState(null);
    const [form, setForm] = useState({});
    const [confirmReset, setConfirmReset] = useState(false);
    const [confirmResetText, setConfirmResetText] = useState("");
    const openNew = () => { setForm({ username: "", password: "", fullName: "", role: "staff", active: true }); setEditing({}); };
    const openEdit = (a) => { setForm({ ...a, password: "" }); setEditing(a); };
    const submit = async () => {
        const username = form.username.trim().toLowerCase();
        if (!username || !form.fullName) return;
        if (!editing.id && !form.password) { alert("Vui lòng đặt mật khẩu cho tài khoản mới."); return; }
        if (accounts.some((a) => a.username === username && a.id !== editing.id)) { alert("Tên đăng nhập đã tồn tại."); return; }
        let passFields = {};
        if (form.password) {
            const salt = randomSalt();
            const hash = await hashPassword(form.password, salt);
            passFields = { passwordHash: hash, passwordSalt: salt };
        }
        if (editing.id) {
            setAccounts((prev) => prev.map((a) => (a.id === editing.id ? { ...a, username, fullName: form.fullName, role: form.role, active: form.active, ...passFields } : a)));
        } else {
            setAccounts((prev) => [...prev, { id: uid(), username, fullName: form.fullName, role: form.role, active: true, ...passFields }]);
        }
        setEditing(null);
    };
    const toggleActive = (a) => {
        if (a.id === currentUser.id) { alert("Không thể tự khoá tài khoản đang đăng nhập."); return; }
        setAccounts((prev) => prev.map((x) => (x.id === a.id ? { ...x, active: !x.active } : x)));
    };
    const remove = (a) => {
        if (a.id === currentUser.id) { alert("Không thể xoá tài khoản đang đăng nhập."); return; }
        setAccounts((prev) => prev.filter((x) => x.id !== a.id));
    };

    return (
        <div>
            <div className="flex items-center justify-between mb-5">
                <p className="text-sm opacity-60">Chỉ Quản trị viên mới thấy và chỉnh sửa được mục này.</p>
                <button onClick={openNew} className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-sm text-white shrink-0" style={{ background: INK }}><Plus size={15} /> Thêm tài khoản</button>
            </div>

            <div className="rounded-sm overflow-auto min-w-0" style={{ border: `1px solid ${LINE}`, background: "#fff", maxHeight: "65vh" }}>
                <table className="w-full text-sm" style={{ minWidth: 620 }}>
                    <thead>
                    <tr style={{ borderBottom: `2px solid ${INK}` }}>
                        {["Tên đăng nhập", "Họ tên", "Vai trò", "Trạng thái", ""].map((h, hi) => (
                            <th key={hi} className="text-left px-3 py-2.5 text-xs uppercase tracking-wider whitespace-nowrap" style={{ color: INK, opacity: 0.6 }}>{h}</th>
                        ))}
                    </tr>
                    </thead>
                    <tbody>
                    {accounts.map((a) => {
                        const role = ACCOUNT_ROLES.find((r) => r.id === a.role);
                        return (
                            <tr key={a.id} style={{ borderBottom: `1px dashed ${LINE}` }} className="hover:bg-black/[0.02]">
                                <td className="px-3 py-3 font-medium whitespace-nowrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: INK }}>{a.username}{a.id === currentUser.id && <span className="ml-1.5 text-[10px] opacity-40">(bạn)</span>}</td>
                                <td className="px-3 py-3 whitespace-nowrap" style={{ color: INK }}>{a.fullName}</td>
                                <td className="px-3 py-3 whitespace-nowrap">
                                    <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: a.role === "admin" ? `${BRASS}1A` : `${BLUE}15`, color: a.role === "admin" ? BRASS : BLUE }}>{role.label}</span>
                                </td>
                                <td className="px-3 py-3 whitespace-nowrap">
                                    <button onClick={() => toggleActive(a)} className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: a.active ? `${FOREST}1A` : `${RUST}1A`, color: a.active ? FOREST : RUST }}>
                                        {a.active ? "Đang hoạt động" : "Đã khoá"}
                                    </button>
                                </td>
                                <td className="px-3 py-3">
                                    <div className="flex gap-1.5 justify-end whitespace-nowrap">
                                        <button onClick={() => openEdit(a)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60"><Pencil size={14} /></button>
                                        <button onClick={() => remove(a)} className="p-1.5 rounded-sm hover:bg-black/5 opacity-60" style={{ color: RUST }}><Trash2 size={14} /></button>
                                    </div>
                                </td>
                            </tr>
                        );
                    })}
                    </tbody>
                </table>
            </div>

            {onResetTestData && (
                <div className="mt-8 p-4 rounded-sm" style={{ border: `1px solid ${RUST}44`, background: `${RUST}08` }}>
                    <p className="text-sm font-medium mb-1" style={{ color: RUST }}>Vùng nguy hiểm</p>
                    <p className="text-xs opacity-70 mb-3">Xoá toàn bộ đơn bán, đơn nhập hàng và báo giá đang có (thường dùng để dọn sạch dữ liệu test trước khi đưa vào dùng thật). Sản phẩm, khách hàng, nhà cung cấp, kế hoạch và tài khoản sẽ được giữ nguyên.</p>
                    <button onClick={() => { setConfirmResetText(""); setConfirmReset(true); }} className="text-xs px-3.5 py-2 rounded-sm border" style={{ borderColor: RUST, color: RUST }}>
                        Đặt lại dữ liệu giao dịch (đơn bán / đơn nhập / báo giá)
                    </button>
                </div>
            )}

            {confirmReset && (
                <Modal title="Xác nhận đặt lại dữ liệu" onClose={() => setConfirmReset(false)}>
                    <p className="text-sm mb-3">Thao tác này sẽ <b style={{ color: RUST }}>xoá vĩnh viễn</b> toàn bộ đơn bán, đơn nhập hàng, báo giá và các bút toán kho liên quan. Không thể hoàn tác. Sản phẩm, khách hàng, nhà cung cấp, kế hoạch, tài khoản sẽ không bị ảnh hưởng.</p>
                    <Field label={<>Gõ <b>XOA</b> để xác nhận</>}>
                        <input className={inputCls} style={{ borderColor: LINE }} value={confirmResetText} onChange={(e) => setConfirmResetText(e.target.value)} placeholder="XOA" />
                    </Field>
                    <button onClick={() => { onResetTestData(); setConfirmReset(false); }} disabled={confirmResetText.trim().toUpperCase() !== "XOA"}
                            className="w-full py-2.5 rounded-sm text-white text-sm mt-2 disabled:opacity-40" style={{ background: RUST }}>
                        Xác nhận xoá toàn bộ dữ liệu giao dịch
                    </button>
                </Modal>
            )}

            {editing !== null && (
                <Modal title={editing.id ? "Sửa tài khoản" : "Thêm tài khoản"} onClose={() => setEditing(null)}>
                    <Field label="Tên đăng nhập"><input className={inputCls} style={{ borderColor: LINE }} value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} disabled={!!editing.id} /></Field>
                    <Field label="Họ tên"><input className={inputCls} style={{ borderColor: LINE }} value={form.fullName} onChange={(e) => setForm({ ...form, fullName: e.target.value })} /></Field>
                    <Field label="Mật khẩu" hint={editing.id ? "Để trống nếu không đổi mật khẩu" : ""}>
                        <input type="password" className={inputCls} style={{ borderColor: LINE }} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} />
                    </Field>
                    <Field label="Vai trò">
                        <div className="flex gap-2">
                            {ACCOUNT_ROLES.map((r) => (
                                <button key={r.id} type="button" onClick={() => setForm({ ...form, role: r.id })} className="px-3.5 py-1.5 rounded-sm text-sm border"
                                        style={{ borderColor: form.role === r.id ? INK : LINE, background: form.role === r.id ? INK : "transparent", color: form.role === r.id ? "#fff" : INK }}>
                                    {r.label}
                                </button>
                            ))}
                        </div>
                    </Field>
                    <button onClick={submit} className="w-full py-2.5 rounded-sm text-white text-sm mt-2" style={{ background: INK }}>{editing.id ? "Lưu thay đổi" : "Thêm tài khoản"}</button>
                </Modal>
            )}
        </div>
    );
}

const NOTIF_CATEGORIES = [
    { id: "approval", label: "Đơn hàng cần duyệt", color: RUST },
    { id: "pending_stock", label: "Đơn chờ hàng (thiếu series/tồn kho)", color: RUST },
    { id: "po_due", label: "Công nợ NCC sắp/đã đến hạn", color: RUST },
    { id: "b2b_due", label: "Khách B2B quá hạn công nợ", color: RUST },
    { id: "neg_stock", label: "Sản phẩm âm tồn", color: RUST },
    { id: "low_stock", label: "Sản phẩm dưới định mức tồn", color: BRASS },
    { id: "plan_kpi", label: "Kế hoạch có nguy cơ không đạt", color: BRASS },
    { id: "order_cancelled", label: "Đơn hàng bị huỷ bởi nhân viên/CTV", color: RUST },
];

function NotificationBell({ notifications, markRead, markAllRead, onGoto }) {
    const [open, setOpen] = useState(false);
    const unread = notifications.filter((n) => !n.read);
    const sorted = [...notifications].sort((a, b) => (a.read === b.read ? (a.createdAt < b.createdAt ? 1 : -1) : a.read ? 1 : -1));
    const isNavigable = (n) => { const prefix = n.key.split(":")[0]; return ["appr", "cxreq", "rtreq", "b2bdue", "pend", "podue", "low", "neg"].includes(prefix); };

    return (
        <div>
            <button onClick={() => setOpen((o) => !o)} className="w-full flex items-center gap-3 px-4 py-3 rounded-sm text-sm relative" style={{ color: "rgba(255,255,255,0.75)", background: open ? "rgba(255,255,255,0.08)" : "transparent" }}>
                <Bell size={18} />
                Thông báo
                {unread.length > 0 && <span className="ml-auto text-[10px] w-5 h-5 rounded-full flex items-center justify-center text-white font-semibold" style={{ background: RUST }}>{unread.length > 99 ? "99+" : unread.length}</span>}
            </button>
            {open && (
                <div className="mt-1 rounded-sm p-2.5 max-h-[45vh] overflow-y-auto" style={{ background: "rgba(0,0,0,0.18)", border: "1px solid rgba(255,255,255,0.1)" }}>
                    <div className="flex items-center justify-between mb-2 px-0.5">
                        <span className="text-[11px]" style={{ color: "rgba(255,255,255,0.45)" }}>{unread.length} chưa đọc · tự xoá sau 3 ngày</span>
                        {unread.length > 0 && <button onClick={markAllRead} className="text-[11px] underline" style={{ color: "rgba(255,255,255,0.6)" }}>Đánh dấu tất cả</button>}
                    </div>
                    {sorted.length === 0 ? <p className="text-xs text-center py-6" style={{ color: "rgba(255,255,255,0.4)" }}>Không có thông báo nào.</p> : (
                        <div className="space-y-1.5">
                            {sorted.map((n) => {
                                const cat = NOTIF_CATEGORIES.find((c) => c.id === n.category);
                                const nav = isNavigable(n);
                                return (
                                    <div key={n.id} className="flex items-start gap-2 p-2 rounded-sm" style={{ background: n.read ? "transparent" : "rgba(255,255,255,0.06)", opacity: n.read ? 0.5 : 1 }}>
                                        <button onClick={() => markRead(n.id)} title={n.read ? "Đã đọc" : "Đánh dấu đã đọc"} className="mt-0.5 w-4 h-4 rounded-sm border flex items-center justify-center shrink-0"
                                                style={{ borderColor: n.read ? FOREST : "rgba(255,255,255,0.3)", background: n.read ? FOREST : "transparent" }}>
                                            {n.read && <Check size={11} color="#fff" />}
                                        </button>
                                        <button onClick={() => { if (nav) { onGoto(n); setOpen(false); } }} disabled={!nav} className="min-w-0 text-left" style={{ cursor: nav ? "pointer" : "default" }}>
                                            <p className="text-[10px] uppercase tracking-wider font-medium" style={{ color: cat?.color || "rgba(255,255,255,0.7)" }}>{cat?.label}</p>
                                            <p className="text-xs mt-0.5" style={{ color: "rgba(255,255,255,0.85)", textDecoration: nav ? "underline" : "none", textDecorationColor: "rgba(255,255,255,0.25)" }}>{n.detail}</p>
                                        </button>
                                    </div>
                                );
                            })}
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}

const TABS = [
    { id: "dashboard", label: "Tổng quan", icon: LayoutDashboard },
    { id: "products", label: "Sản phẩm & Tồn kho", icon: Package },
    { id: "quotes", label: "Báo giá", icon: FileText },
    { id: "orders", label: "Bán hàng", icon: ShoppingCart },
    { id: "shipping", label: "Vận chuyển", icon: PackageCheck },
    { id: "customers", label: "Khách hàng", icon: Users },
    { id: "suppliers", label: "Nhà cung cấp", icon: Truck },
    { id: "plans", label: "Kế hoạch", icon: Target },
    { id: "reports", label: "Báo cáo", icon: BarChart3 },
];

export default function SalesManager() {
    const [tab, setTab] = useState("dashboard");
    const [products, setProducts] = useState([]);
    const [orders, setOrders] = useState([]);
    const [customers, setCustomers] = useState([]);
    const [purchaseOrders, setPurchaseOrders] = useState([]);
    const [suppliers, setSuppliers] = useState([]);
    const [categories, setCategories] = useState([]);
    const [brands, setBrands] = useState([]);
    const [stocktakes, setStocktakes] = useState([]);
    const [warrantyTickets, setWarrantyTickets] = useState([]);
    const [repairTickets, setRepairTickets] = useState([]);
    const [helpdeskTickets, setHelpdeskTickets] = useState([]);
    const [shippingTickets, setShippingTickets] = useState([]);
    const [plans, setPlans] = useState([]);
    const [accounts, setAccounts] = useState([]);
    const [activityLog, setActivityLog] = useState([]);
    const [notifications, setNotifications] = useState([]);
    const [quotations, setQuotations] = useState([]);
    const [navTarget, setNavTarget] = useState(null); // { type: 'order'|'po'|'product', id }
    // Điều hướng nhanh tới đúng đơn bán/đơn nhập từ mã số phiếu (dùng ở mọi nơi hiển thị mã đơn: lịch sử tồn kho, lịch sử khách hàng...).
    const goToDoc = (docNo) => {
        if (!docNo) return;
        if (docNo.startsWith("POH")) {
            const po = purchaseOrders.find((p) => p.code === docNo);
            if (po) { setTab("products"); setNavTarget({ type: "po", id: po.id }); }
        } else if (docNo.startsWith("DH")) {
            const orderCode = docNo.split("-")[0];
            const o = orders.find((x) => x.code === orderCode);
            if (o) { setTab("orders"); setNavTarget({ type: "order", id: o.id }); }
        }
    };
    const goToSupplier = (supplierId) => {
        if (!supplierId || currentUser.role !== "admin") return;
        setTab("suppliers"); setNavTarget({ type: "supplier", id: supplierId });
    };
    // status: "approval_pending" | "return_request" | id trong STATUSES (pending/shipping/delivered/done/cancelled)
    const goToOrdersFilter = (status) => {
        setTab("orders"); setNavTarget({ type: "orders-filter", status });
    };
    const [printSettings, setPrintSettings] = useState(DEFAULT_PRINT_SETTINGS);
    const [currentUserId, setCurrentUserId] = useState(null);
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
                    // Danh mục nhóm hàng: nếu dữ liệu cũ chưa có danh sách quản lý riêng, tự sinh từ các category đã dùng trên sản phẩm (không mất dữ liệu).
                    const existingCats = [...new Set((data.products || []).map((p) => p.category).filter(Boolean))];
                    setCategories(Array.isArray(data.categories) && data.categories.length > 0 ? data.categories : existingCats);
                    // Danh mục nhãn hiệu: tương tự — tự sinh từ các brand đã dùng trên sản phẩm nếu chưa có danh sách quản lý riêng.
                    // Mỗi nhãn hiệu giờ thuộc về 1 nhóm hàng cụ thể; dữ liệu cũ (chuỗi đơn, chưa có nhóm hàng) được tự chuyển sang dạng mới.
                    const existingBrands = [...new Set((data.products || []).map((p) => p.brand).filter(Boolean))];
                    const rawBrands = Array.isArray(data.brands) && data.brands.length > 0 ? data.brands : existingBrands;
                    setBrands(rawBrands.map(normalizeBrandEntry));
                    setStocktakes((data.stocktakes || []).map(normalizeStocktake));
                    setWarrantyTickets((data.warrantyTickets || []).map(normalizeWarrantyTicket));
                    setRepairTickets((data.repairTickets || []).map(normalizeRepairTicket));
                    setHelpdeskTickets((data.helpdeskTickets || []).map(normalizeHelpdeskTicket));
                    setShippingTickets((data.shippingTickets || []).map(normalizeShippingTicket));
                    setPlans((data.plans || []).map(normalizePlan));
                    setActivityLog((data.activityLog || []).map(normalizeLog));
                    setNotifications((data.notifications || []).map(normalizeNotif));
                    setQuotations((data.quotations || []).map(normalizeQuote));
                    setPrintSettings(normalizePrintSettings(data.printSettings));
                    const rawAccs = (data.accounts && data.accounts.length > 0) ? data.accounts.map(normalizeAccount) : seedAccounts();
                    const accs = await migrateAccountPasswords(rawAccs);
                    setAccounts(accs);
                    setCurrentUserId((data.session && data.session.userId && accs.some((a) => a.id === data.session.userId)) ? data.session.userId : null);
                } else {
                    const seed = seedData();
                    setProducts(seed.products.map(normalizeProduct));
                    setOrders(seed.orders.map(normalizeOrder));
                    setCustomers(seed.customers.map(normalizeCustomer));
                    setPurchaseOrders([]);
                    setSuppliers([]);
                    setCategories([...new Set(seed.products.map((p) => p.category).filter(Boolean))]);
                    setBrands([...new Set(seed.products.map((p) => p.brand).filter(Boolean))].map(normalizeBrandEntry));
                    setStocktakes([]);
                    setWarrantyTickets([]);
                    setRepairTickets([]);
                    setHelpdeskTickets([]);
                    setShippingTickets([]);
                    setPlans([]);
                    setActivityLog([]);
                    setNotifications([]);
                    setAccounts(await migrateAccountPasswords(seedAccounts()));
                    setCurrentUserId(null);
                }
            } catch (e) {
                console.error("Lỗi chuẩn hoá dữ liệu, dùng dữ liệu mẫu:", e);
                const seed = seedData();
                setProducts(seed.products);
                setOrders(seed.orders);
                setCustomers(seed.customers);
                setPurchaseOrders([]);
                setSuppliers([]);
                setPlans([]);
                setActivityLog([]);
                setNotifications([]);
                setAccounts(await migrateAccountPasswords(seedAccounts()));
                setCurrentUserId(null);
            }
            setLoaded(true);
        })();
    }, []);

    useEffect(() => {
        if (!loaded) return;
        const t = setTimeout(() => { saveData({ products, orders, customers, purchaseOrders, suppliers, categories, brands, stocktakes, warrantyTickets, repairTickets, helpdeskTickets, shippingTickets, plans, accounts, activityLog, notifications, printSettings, quotations, session: { userId: currentUserId } }); }, 400);
        return () => clearTimeout(t);
    }, [products, orders, customers, purchaseOrders, suppliers, categories, brands, stocktakes, warrantyTickets, repairTickets, helpdeskTickets, shippingTickets, plans, accounts, activityLog, notifications, printSettings, quotations, currentUserId, loaded]);

    // Đồng bộ danh sách thông báo cho Admin từ dữ liệu thật: sản phẩm dưới định mức/âm tồn, công nợ NCC,
    // công nợ khách B2B quá hạn, đơn hàng cần duyệt (kể cả yêu cầu huỷ/đổi trả). Thông báo đã đọc tự xoá sau 3 ngày.
    useEffect(() => {
        if (!loaded) return;
        const activeKeys = new Set();
        const fresh = [];
        const push = (key, category, detail) => {
            activeKeys.add(key);
            fresh.push({ key, category, detail });
        };
        products.forEach((p) => {
            const s = productStats(p);
            if (s.closingQty < 0) push(`neg:${p.id}`, "neg_stock", `${p.name} (${s.closingQty})`);
            else if (s.closingQty <= (p.minStockLevel ?? 5)) push(`low:${p.id}`, "low_stock", `${p.name} (còn ${s.closingQty}, định mức ${p.minStockLevel ?? 5})`);
        });
        purchaseOrders.forEach((po) => {
            const due = poDueInfo(po);
            if (due && due.nearDue) push(`podue:${po.id}`, "po_due", `${po.code} (${po.supplier || "NCC"}) — ${due.overdue ? `quá hạn ${-due.daysLeft} ngày` : due.daysLeft === 0 ? "đến hạn hôm nay" : `còn ${due.daysLeft} ngày`}`);
        });
        orders.forEach((o) => {
            const cust = customers.find((c) => c.id === o.customerId);
            if (cust && cust.group === "b2b") {
                const due = orderDueInfo(o);
                if (due && due.overdue3) push(`b2bdue:${o.id}`, "b2b_due", `${o.code} (${cust.name}) — quá hạn ${-due.daysLeft} ngày`);
            }
            if (o.status !== "cancelled") {
                if (o.approvalStatus === "pending") push(`appr:${o.id}`, "approval", `${o.code} — ${o.approvalReason || "CTV tạo đơn"}`);
                if (o.cancelRequest) push(`cxreq:${o.id}`, "approval", `${o.code} — Yêu cầu huỷ bởi ${o.cancelRequest.requestedByName}`);
                if (o.returnRequest) push(`rtreq:${o.id}`, "approval", `${o.code} — Yêu cầu ${o.returnRequest.type === "exchange" ? "đổi hàng" : "hoàn tiền"} bởi ${o.returnRequest.requestedByName}`);
                const pendingItems = o.items.filter((it) => !it.fulfilled);
                if (pendingItems.length > 0) {
                    const names = pendingItems.map((it) => products.find((p) => p.id === it.productId)?.name || "?").join(", ");
                    push(`pend:${o.id}`, "pending_stock", `${o.code} — thiếu series/tồn kho: ${names}. Cần nhập hàng về để bổ sung.`);
                }
            }
            if (o.status === "cancelled" && o.cancelledByRole && o.cancelledByRole !== "admin") {
                const cancelledDaysAgo = o.cancelledAt ? Math.floor((Date.now() - new Date(o.cancelledAt).getTime()) / 86400000) : 99;
                if (cancelledDaysAgo <= 3) {
                    push(`cancelled:${o.id}`, "order_cancelled", `${o.code} — Huỷ bởi ${o.cancelledByName} (${ACCOUNT_ROLES.find((r) => r.id === o.cancelledByRole)?.label || o.cancelledByRole})${o.cancelReason ? ` · Lý do: ${o.cancelReason}` : ""}`);
                }
            }
        });

        // Cảnh báo kế hoạch có nguy cơ không đạt: đã qua kỳ mà tiến độ chậm hơn hẳn thời gian đã trôi qua, hoặc gần hết tháng mà còn thấp.
        {
            const now = new Date();
            const curMonth = todayISO().slice(0, 7);
            const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
            const daysLeftInMonth = daysInMonth - now.getDate();
            const expectedPct = (now.getDate() / daysInMonth) * 100;
            plans.forEach((pl) => {
                if (pl.month !== curMonth || !pl.targetValue) return;
                const actual = planActual(pl, orders, purchaseOrders, products);
                const pct = pl.targetValue > 0 ? (actual / pl.targetValue) * 100 : 0;
                const behindSchedule = pct < expectedPct - 15;
                const nearEndLow = daysLeftInMonth <= 5 && pct < 90;
                if (!behindSchedule && !nearEndLow) return;
                const scope = pl.scope === "product" ? (products.find((x) => x.id === pl.targetProductId)?.name || "Sản phẩm đã xoá")
                    : pl.scope === "category" ? `Nhóm ${pl.targetCategory}` : (pl.type === "sales" ? "Bán hàng" : "Nhập hàng");
                const who = pl.sellerName ? ` · ${pl.sellerName}` : "";
                push(`kpi:${pl.id}`, "plan_kpi", `${monthLabel(pl.month)} — ${scope}${who}: đạt ${Math.round(pct)}% mục tiêu${daysLeftInMonth <= 5 ? ` (còn ${daysLeftInMonth} ngày)` : ""}`);
            });
        }

        setNotifications((prev) => {
            const now = new Date().toISOString();
            const threeDaysAgo = Date.now() - 3 * 24 * 3600 * 1000;
            // Xoá thông báo đã đọc quá 3 ngày, hoặc điều kiện gốc không còn đúng nữa (đã xử lý xong).
            let next = prev.filter((n) => {
                if (n.read && n.readAt && new Date(n.readAt).getTime() < threeDaysAgo) return false;
                if (!activeKeys.has(n.key)) return false;
                return true;
            });
            fresh.forEach((f) => {
                const existing = next.find((n) => n.key === f.key);
                if (existing) { if (existing.detail !== f.detail) existing.detail = f.detail; }
                else next.push({ id: uid(), key: f.key, category: f.category, title: "", detail: f.detail, createdAt: now, read: false, readAt: null });
            });
            return next;
        });
    }, [loaded, products, purchaseOrders, orders, customers, plans]);

    // Nhắc công nợ NCC: khi còn ≤3 ngày tới hạn (hoặc đã quá hạn) mà chưa thanh toán, tự ghi 1 dòng
    // vào Nhật ký hoạt động mỗi ngày (không lặp lại trong cùng 1 ngày) để admin theo dõi như thông báo.
    useEffect(() => {
        if (!loaded) return;
        const today = todayISO();
        setActivityLog((prev) => {
            const toLog = [];
            purchaseOrders.forEach((po) => {
                const due = poDueInfo(po);
                if (!due || !due.nearDue) return;
                const already = prev.some((l) => l.action === "Nhắc công nợ NCC" && l.detail.startsWith(po.code + " ") && l.at.slice(0, 10) === today);
                if (already) return;
                toLog.push({
                    id: uid(), at: new Date().toISOString(), userId: "system", userName: "Hệ thống", role: "admin",
                    action: "Nhắc công nợ NCC",
                    detail: `${po.code} ${po.supplier ? "· " + po.supplier + " " : ""}· ${due.overdue ? `quá hạn ${-due.daysLeft} ngày` : due.daysLeft === 0 ? "đến hạn hôm nay" : `còn ${due.daysLeft} ngày đến hạn`}`,
                });
            });
            return toLog.length > 0 ? [...toLog, ...prev].slice(0, 500) : prev;
        });
    }, [loaded, purchaseOrders]);

    if (!loaded) {
        return <div className="flex items-center justify-center h-96" style={{ color: INK }}><Loader2 className="animate-spin mr-2" size={18} /> Đang tải dữ liệu…</div>;
    }

    const currentUser = accounts.find((a) => a.id === currentUserId) || null;
    if (!currentUser) {
        return <LoginScreen accounts={accounts} onLogin={(id) => setCurrentUserId(id)} />;
    }

    const addLog = (action, detail) => {
        setActivityLog((prev) => [{ id: uid(), at: new Date().toISOString(), userId: currentUser.id, userName: currentUser.fullName, role: currentUser.role, action, detail }, ...prev].slice(0, 500));
    };

    // Đặt lại dữ liệu giao dịch test (đơn bán, đơn nhập hàng, báo giá) — dùng khi cần dọn sạch dữ liệu thử nghiệm trước khi đưa vào dùng thật.
    // Giữ nguyên: sản phẩm, khách hàng, nhà cung cấp, kế hoạch, tài khoản. Tự động dọn luôn các bút toán kho phát sinh từ các đơn bị xoá
    // (mã chứng từ bắt đầu bằng DH hoặc POH, kể cả các bút toán hoàn/sửa kho -HUY/-SUA), không đụng tới các lần nhập/xuất kho thủ công khác.
    const resetTestData = () => {
        setOrders([]);
        setPurchaseOrders([]);
        setQuotations([]);
        setProducts((prev) => prev.map((p) => ({
            ...p,
            movements: (p.movements || []).filter((m) => !/^(DH|POH)/i.test(m.docNo || "")),
        })));
        addLog("Đặt lại dữ liệu test", "Đã xoá toàn bộ đơn bán, đơn nhập hàng, báo giá và các bút toán kho liên quan");
    };

    const roleTabIds = currentUser.role === "admin" ? ["dashboard", "products", "quotes", "orders", "shipping", "customers", "suppliers", "plans", "reports"]
        : currentUser.role === "staff" ? ["dashboard", "products", "quotes", "orders", "shipping", "customers"]
            : ["dashboard", "products", "quotes", "orders", "shipping", "customers"]; // ctv — chỉ xem sản phẩm, không tạo/sửa/xoá
    const visibleTabs = [
        ...TABS.filter((t) => roleTabIds.includes(t.id)),
        ...(currentUser.role === "admin" ? [{ id: "activity", label: "Nhật ký", icon: History }, { id: "accounts", label: "Tài khoản", icon: KeyRound }] : []),
    ];
    const employeeNames = accounts.filter((a) => a.active).map((a) => a.fullName);

    return (
        <div style={{ background: PAPER, minHeight: "100%", fontFamily: "'Inter', sans-serif" }} className="w-full overflow-x-hidden">
            <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        select { appearance: none; }
      `}</style>
            <div id="app-shell" className="flex flex-col md:flex-row">
                <div className="md:w-72 shrink-0 p-5 md:min-h-screen flex flex-col" style={{ background: INK }}>
                    <div className="mb-8 flex items-center gap-2.5">
                        <HiliLogo size={34} />
                        <div>
                            <h1 style={{ fontFamily: "'Fraunces', serif", color: "#fff" }} className="text-base leading-tight">Quản lý bán hàng<br />Hilitek</h1>
                        </div>
                    </div>
                    <nav className="flex flex-col gap-2 w-full">
                        {visibleTabs.map((t) => (
                            <button key={t.id} onClick={() => setTab(t.id)} className="flex items-center gap-3 px-4 py-3.5 rounded-sm text-base font-medium transition-colors w-full text-left"
                                    style={{ background: tab === t.id ? "rgba(255,255,255,0.1)" : "transparent", color: tab === t.id ? "#fff" : "rgba(255,255,255,0.6)", borderLeft: tab === t.id ? `3px solid ${BRASS}` : "3px solid transparent" }}>
                                <t.icon size={20} className="shrink-0" /> {t.label}
                            </button>
                        ))}
                    </nav>
                    {currentUser.role === "admin" && (
                        <div className="mt-3 pt-3" style={{ borderTop: "1px solid rgba(255,255,255,0.12)" }}>
                            <NotificationBell
                                notifications={notifications}
                                markRead={(id) => setNotifications((prev) => prev.map((n) => (n.id === id ? { ...n, read: !n.read, readAt: !n.read ? new Date().toISOString() : null } : n)))}
                                markAllRead={() => { const now = new Date().toISOString(); setNotifications((prev) => prev.map((n) => (n.read ? n : { ...n, read: true, readAt: now }))); }}
                                onGoto={(n) => {
                                    const [prefix, id] = n.key.split(":");
                                    if (["appr", "cxreq", "rtreq", "b2bdue", "pend", "cancelled"].includes(prefix)) { setTab("orders"); setNavTarget({ type: "order", id }); }
                                    else if (prefix === "podue") { setTab("products"); setNavTarget({ type: "po", id }); }
                                    else if (prefix === "low" || prefix === "neg") { setTab("products"); setNavTarget({ type: "product", id }); }
                                }}
                            />
                        </div>
                    )}
                    <div className="mt-auto pt-6 flex items-center justify-between gap-2" style={{ borderTop: "1px solid rgba(255,255,255,0.12)" }}>
                        <div className="min-w-0 pt-4">
                            <p className="text-sm font-medium truncate" style={{ color: "#fff" }}>{currentUser.fullName}</p>
                            <p className="text-[11px] opacity-50" style={{ color: "#fff" }}>{ACCOUNT_ROLES.find((r) => r.id === currentUser.role)?.label}</p>
                        </div>
                        <button onClick={() => setCurrentUserId(null)} title="Đăng xuất" className="pt-4 opacity-60 hover:opacity-100 shrink-0" style={{ color: "#fff" }}><LogOut size={16} /></button>
                    </div>
                </div>
                <div className="flex-1 p-5 md:p-8 min-w-0">
                    <div className="flex items-center justify-between mb-6">
                        <h2 style={{ fontFamily: "'Fraunces', serif", color: INK }} className="text-2xl">{visibleTabs.find((t) => t.id === tab)?.label || "Bán hàng"}</h2>
                        <span className="text-xs opacity-40" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{todayISO()}</span>
                    </div>
                    <AppErrorBoundary key={tab}>
                        {tab === "dashboard" && <Dashboard products={products} orders={orders} goToOrdersFilter={goToOrdersFilter} />}
                        {tab === "products" && roleTabIds.includes("products") && <ProductsSection products={products} setProducts={setProducts} purchaseOrders={purchaseOrders} setPurchaseOrders={setPurchaseOrders} suppliers={suppliers} setSuppliers={setSuppliers} categories={categories} setCategories={setCategories} brands={brands} setBrands={setBrands} stocktakes={stocktakes} setStocktakes={setStocktakes} warrantyTickets={warrantyTickets} setWarrantyTickets={setWarrantyTickets} repairTickets={repairTickets} setRepairTickets={setRepairTickets} helpdeskTickets={helpdeskTickets} setHelpdeskTickets={setHelpdeskTickets} orders={orders} customers={customers} employeeNames={employeeNames} currentUser={currentUser} addLog={addLog} navTarget={tab === "products" ? navTarget : null} onFocusHandled={() => setNavTarget(null)} goToDoc={goToDoc} goToSupplier={goToSupplier} />}
                        {tab === "quotes" && <Quotations quotations={quotations} setQuotations={setQuotations} orders={orders} setOrders={setOrders} products={products} setProducts={setProducts} customers={customers} setCustomers={setCustomers} employeeNames={employeeNames} currentUser={currentUser} addLog={addLog} goToDoc={goToDoc} brands={brands} />}
                        {tab === "orders" && <Orders orders={orders} setOrders={setOrders} products={products} setProducts={setProducts} customers={customers} setCustomers={setCustomers} employeeNames={employeeNames} currentUser={currentUser} addLog={addLog} focusOrderId={tab === "orders" ? navTarget?.type === "order" ? navTarget.id : null : null} initialFilterStatus={tab === "orders" && navTarget?.type === "orders-filter" ? navTarget.status : null} onFocusHandled={() => setNavTarget(null)} printSettings={printSettings} setPrintSettings={setPrintSettings} />}
                        {tab === "shipping" && roleTabIds.includes("shipping") && <Shipping shippingTickets={shippingTickets} setShippingTickets={setShippingTickets} orders={orders} customers={customers} currentUser={currentUser} addLog={addLog} />}
                        {tab === "customers" && <Customers customers={customers} setCustomers={setCustomers} orders={orders} products={products} currentUser={currentUser} addLog={addLog} goToDoc={goToDoc} employeeNames={employeeNames} />}
                        {tab === "suppliers" && <Suppliers suppliers={suppliers} setSuppliers={setSuppliers} purchaseOrders={purchaseOrders} addLog={addLog} goToDoc={goToDoc} navTarget={tab === "suppliers" ? navTarget : null} onFocusHandled={() => setNavTarget(null)} />}
                        {tab === "plans" && roleTabIds.includes("plans") && <Plans plans={plans} setPlans={setPlans} orders={orders} purchaseOrders={purchaseOrders} products={products} employeeNames={employeeNames} />}
                        {tab === "reports" && roleTabIds.includes("reports") && <Reports orders={orders} products={products} customers={customers} accounts={accounts} purchaseOrders={purchaseOrders} warrantyTickets={warrantyTickets} />}
                        {tab === "activity" && currentUser.role === "admin" && <ActivityLog log={activityLog} accounts={accounts} />}
                        {tab === "accounts" && currentUser.role === "admin" && <Accounts accounts={accounts} setAccounts={setAccounts} currentUser={currentUser} onResetTestData={resetTestData} />}
                    </AppErrorBoundary>
                </div>
            </div>
        </div>
    );
}
