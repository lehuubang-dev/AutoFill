import { useState, useMemo, type ReactNode, type ChangeEvent } from "react";
import {
  Plus,
  Trash2,
  Download,
  AlertTriangle,
  FileText,
  Pencil,
  X as XIcon,
  Upload,
  Loader2,
  Info,
} from "lucide-react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import {
  createWorker,
  PSM,
  type Worker as TesseractWorker,
} from "tesseract.js";

// Worker cho pdfjs-dist (Vite). Cần: npm install pdfjs-dist xlsx lucide-react
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

/**
 * BẢNG DỮ LIỆU LỆNH BẮT BỊ CAN — 11 CỘT CỐ ĐỊNH
 * ------------------------------------------------------------------
 *  1. STT              7. Tôn giáo
 *  2. Họ tên           8. Thẻ CCCD/Thẻ CC (12 số, lưu dạng text)
 *  3. Giới tính        9. Nơi thường trú (chuẩn hoá theo NĐ 30/2020)
 *  4. Sinh ngày       10. Lệnh
 *  5. Quốc tịch       11. Bị bắt ngày
 *  6. Dân tộc
 *
 * Thêm cột X — mặc định "GP" cho mọi bản ghi, không nhập tay.
 * Mọi trường khác của mẫu 74 đều bỏ, không đọc, không lưu.
 *
 * Nguyên tắc đọc PDF: chỉ điền phần in sẵn/đánh máy đọc chắc chắn.
 * Trường viết tay (ngày bắt thực tế) được thử đọc ở mục "Ghi chú:
 * Bắt bị can ngày..." và đối chiếu với mục giao lệnh; nếu không đọc
 * được thì để trống cho người dùng tự nhập, không suy diễn.
 */

// ---------------------------------------------------------------------------
// SCHEMA — đúng 10 trường nhập (STT tự sinh, X cố định "GP")
// ---------------------------------------------------------------------------

type FieldKey =
  | "hoTen"
  | "gioiTinh"
  | "sinhNgay"
  | "quocTich"
  | "danToc"
  | "tonGiao"
  | "cccd"
  | "noiThuongTru"
  | "lenh"
  | "biBatNgay";

type FieldType = "text" | "textarea" | "date" | "datestr" | "select";

interface FieldConfig {
  key: FieldKey;
  label: string;
  type: FieldType;
  selectOptions?: string[];
  datalistOptions?: string[];
  placeholder?: string;
  full?: boolean;
}

const DAN_TOC_OPTIONS = ["Kinh", "Tày", "Thái", "Mường", "Khmer", "Nùng", "Hoa", "Dao"];
const TON_GIAO_OPTIONS = ["Không", "Phật giáo", "Công giáo", "Tin lành", "Hòa Hảo", "Cao Đài"];

const FIELDS: FieldConfig[] = [
  { key: "hoTen", label: "Họ tên", type: "text", placeholder: "Nguyễn Văn A" },
  { key: "gioiTinh", label: "Giới tính", type: "select", selectOptions: ["Nam", "Nữ"] },
  { key: "sinhNgay", label: "Sinh ngày", type: "datestr", placeholder: "dd/mm/yyyy" },
  { key: "quocTich", label: "Quốc tịch", type: "text", datalistOptions: ["Việt Nam"] },
  { key: "danToc", label: "Dân tộc", type: "text", datalistOptions: DAN_TOC_OPTIONS },
  { key: "tonGiao", label: "Tôn giáo", type: "text", datalistOptions: TON_GIAO_OPTIONS },
  { key: "cccd", label: "Thẻ CCCD/Thẻ CC", type: "text", placeholder: "12 số" },
  { key: "biBatNgay", label: "Bị bắt ngày", type: "datestr", placeholder: "dd/mm/yyyy" },
  { key: "noiThuongTru", label: "Nơi thường trú", type: "textarea", full: true },
  { key: "lenh", label: "Lệnh", type: "textarea", full: true },
];

const FIELD_MAP: Record<FieldKey, FieldConfig> = FIELDS.reduce((acc, f) => {
  acc[f.key] = f;
  return acc;
}, {} as Record<FieldKey, FieldConfig>);

// Nhóm trường dùng chung cho cả form nhập liệu lẫn hộp thoại xem chi
// tiết, để hai nơi này luôn hiển thị đồng nhất.
interface FieldSection {
  title: string;
  keys: FieldKey[];
}

const FIELD_SECTIONS: FieldSection[] = [
  {
    title: "Nhân thân",
    keys: ["hoTen", "gioiTinh", "sinhNgay", "quocTich", "danToc", "tonGiao", "cccd"],
  },
  {
    title: "Bắt giữ & nơi cư trú",
    keys: ["biBatNgay", "noiThuongTru", "lenh"],
  },
];

const ALL_FIELD_KEYS: FieldKey[] = FIELDS.map((f) => f.key);
const DATE_KEYS: FieldKey[] = FIELDS.filter((f) => f.type === "date").map((f) => f.key);
const LABEL_MAP = FIELDS.reduce((acc, f) => {
  acc[f.key] = f.label;
  return acc;
}, {} as Record<FieldKey, string>);

const DEFAULTS: Partial<Record<FieldKey, string>> = {
  danToc: "Kinh",
  quocTich: "Việt Nam",
  tonGiao: "Không",
};

// Cột hiển thị ở bảng chính — rút gọn để dễ quét mắt; bấm vào một dòng
// bất kỳ sẽ mở hộp thoại xem đầy đủ cả 11 trường.
const TABLE_COLUMNS: { key: "stt" | "x" | FieldKey; label: string }[] = [
  { key: "stt", label: "STT" },
  { key: "x", label: "X" },
  { key: "hoTen", label: "Họ tên" },
  { key: "gioiTinh", label: "GT" },
  { key: "sinhNgay", label: "Sinh ngày" },
  { key: "cccd", label: "Thẻ CCCD/Thẻ CC" },
  { key: "biBatNgay", label: "Bị bắt ngày" },
];

// Ánh xạ khi xuất Excel: đặt đúng vào vị trí cột của file mẫu có sẵn
// (không phải cột tự sinh liền nhau như bảng hiển thị). Cột nào không
// nằm trong danh sách này (ví dụ "Quê quán" ở cột D — hệ thống hiện
// chưa thu thập dữ liệu này) vẫn được chừa chỗ đúng vị trí nhưng để
// trống, không suy diễn dữ liệu.
interface ExportField {
  col: string; // ký hiệu cột trong file mẫu, vd "A", "AA", "AS"
  label: string;
  key?: FieldKey | "x";
  getValue?: (r: BiCanRecord) => string;
  isText?: boolean; // ép định dạng text (không mất số 0 đầu) khi xuất
}

const EXPORT_FIELDS: ExportField[] = [
  { col: "A", label: "X", key: "x" },
  { col: "B", label: "Họ tên", key: "hoTen" },
  { col: "C", label: "Năm sinh", key: "sinhNgay", getValue: (r) => r.sinhNgay.split("/").pop() ?? "" },
  { col: "E", label: "ĐKTT", key: "noiThuongTru" },
  { col: "L", label: "GT", key: "gioiTinh" },
  { col: "M", label: "Dân tộc", key: "danToc" },
  { col: "N", label: "Quốc tịch", key: "quocTich" },
  {
    col: "AA",
    label: "CMND/CCCD",
    key: "cccd",
    getValue: (r) => r.cccd + (r.cccdNote ? ` ${r.cccdNote}` : ""),
    isText: true,
  },
  { col: "AG", label: "Ngày bắt", key: "biBatNgay" },
  { col: "AL", label: "Lệnh tạm giam", key: "lenh" },
  { col: "AS", label: "Tôn giáo", key: "tonGiao" },
];

type Draft = Record<FieldKey, string>;

interface BiCanRecord extends Draft {
  _id: number;
  x: string;
  cccdNote: string;
  batNote: string;
}

const emptyDraft: Draft = ALL_FIELD_KEYS.reduce((acc, k) => {
  acc[k] = DEFAULTS[k] ?? "";
  return acc;
}, {} as Draft);

// ---------------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------------

// Chuyển ký hiệu cột Excel kiểu "A", "L", "AA", "AS"... sang chỉ số cột
// 0-based (A -> 0, Z -> 25, AA -> 26...) để đặt đúng dữ liệu vào đúng
// cột của file mẫu có sẵn (không phải cột đầu tiên tự sinh).
function colLetterToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function pad2(n: string): string {
  return n.padStart(2, "0");
}

function toDisplayDate(isoDate: string): string {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-");
  if (!y || !m || !d) return "";
  return `${d}/${m}/${y}`;
}

function toIsoDate(displayDate: string): string {
  if (!displayDate) return "";
  const [d, m, y] = displayDate.split("/");
  if (!d || !m || !y) return "";
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

// Tự chèn dấu "/" khi người dùng gõ số liền tay, giới hạn đúng 8 chữ số
// (dd mm yyyy). Không tự đoán/sửa nếu người dùng đã gõ dấu "/" sẵn.
function maskDateTyping(raw: string): string {
  const digits = raw.replace(/\D/g, "").slice(0, 8);
  const parts = [digits.slice(0, 2), digits.slice(2, 4), digits.slice(4, 8)].filter(Boolean);
  return parts.join("/");
}

// Chuẩn hoá khi lưu: đệm số 0 cho ngày/tháng 1 chữ số (vd 3/9/1995 ->
// 03/09/1995). Không đúng dạng thì giữ nguyên để người dùng tự sửa,
// không suy đoán bừa.
function normalizeDateStr(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return "";
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(trimmed);
  if (m) return `${pad2(m[1])}/${pad2(m[2])}/${m[3]}`;
  return trimmed;
}

// Cột 8: đúng 12 số, lưu dạng chuỗi. Khác 12 số thì ghi chú tình trạng.
function parseCccd(raw: string): { value: string; note: string } {
  const digits = raw.replace(/\D/g, "");
  if (!digits) return { value: "", note: "" };
  if (digits.length !== 12) {
    return {
      value: digits,
      note: `(đọc được ${digits.length} số — cần đủ 12 số, kiểm tra lại)`,
    };
  }
  return { value: digits, note: "" };
}

// Giá trị hiển thị/xuất ra cho một trường, kèm ghi chú (nếu có) ngay
// trong cùng ô — dùng chung cho việc xuất Excel.
function cellValue(r: BiCanRecord, key: "x" | FieldKey): string {
  if (key === "x") return r.x;
  if (key === "cccd") return r.cccd + (r.cccdNote ? ` ${r.cccdNote}` : "");
  if (key === "biBatNgay") return r.biBatNgay;
  return r[key] ?? "";
}

// Cột 9: chuẩn hoá viết hoa/thường theo NĐ 30/2020/NĐ-CP.
// Danh từ chung chỉ đơn vị hành chính/định danh viết thường; tên riêng
// địa lý viết hoa chữ đầu mỗi âm tiết.
const COMMON_NOUNS = [
  "tổ dân phố",
  "khu phố",
  "thị trấn",
  "thị xã",
  "thành phố",
  "quốc lộ",
  "tỉnh lộ",
  "hẻm",
  "ngõ",
  "ngách",
  "số nhà",
  "số",
  "tổ",
  "ấp",
  "thôn",
  "khóm",
  "bản",
  "buôn",
  "làng",
  "xóm",
  "đường",
  "phố",
  "phường",
  "xã",
  "quận",
  "huyện",
  "tỉnh",
  "khu",
  "lô",
  "tầng",
  "căn hộ",
  "chung cư",
];

function titleCaseWord(w: string): string {
  if (!w) return w;
  if (/^\d/.test(w)) return w; // số giữ nguyên
  return w.charAt(0).toUpperCase() + w.slice(1);
}

function normalizeAddress(raw: string): string {
  if (!raw) return "";
  const segments = raw
    .replace(/\s+/g, " ")
    .replace(/\s*,\s*/g, ",")
    .trim()
    .replace(/[.,;]+$/, "")
    .split(",");

  const normalized = segments.map((seg) => {
    let s = seg.trim().toLowerCase();
    if (!s) return "";
    // Viết hoa toàn bộ, sau đó hạ các danh từ chung về chữ thường.
    const words = s.split(" ").map(titleCaseWord);
    s = words.join(" ");
    COMMON_NOUNS.forEach((noun) => {
      const titled = noun.split(" ").map(titleCaseWord).join(" ");
      s = s.replace(new RegExp(`(^|\\s)${titled}(?=\\s|$)`, "g"), (_m, p1) => `${p1}${noun}`);
    });
    return s;
  });

  return normalized.filter(Boolean).join(", ");
}

// Chữ tiếng Việt trong PDF thường ở dạng tổ hợp (NFD) hoặc dùng bộ mã
// riêng của font, nên so khớp thẳng "Họ tên:" rất hay trượt. Cách làm ở
// đây: tạo một bản "không dấu" có ĐỘ DÀI BẰNG ĐÚNG bản gốc (mỗi ký tự
// đổi thành đúng một ký tự), dò nhãn trên bản không dấu rồi cắt giá trị
// ra từ bản gốc theo đúng vị trí đó.
function deaccent(s: string): string {
  return Array.from(s)
    .map((ch) => {
      if (ch === "đ") return "d";
      if (ch === "Đ") return "D";
      const base = ch.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      return base.length > 0 ? base[0] : ch;
    })
    .join("");
}

// Cắt phần giá trị đứng sau một nhãn, dừng ở nhãn kế tiếp.
function sliceAfterLabel(
  block: string,
  flat: string,
  label: RegExp,
  stops: RegExp[],
  maxLen = 300
): string | null {
  const lm = new RegExp(label.source, "i").exec(flat);
  if (!lm) return null;
  const start = lm.index + lm[0].length;
  let end = Math.min(flat.length, start + maxLen);
  stops.forEach((st) => {
    const re = new RegExp(st.source, "gi");
    re.lastIndex = start;
    const r = re.exec(flat);
    if (r && r.index < end) end = r.index;
  });
  const value = block.slice(start, end).replace(/\s+/g, " ").replace(/^[:\-–\s]+/, "").replace(/[;,.\s]+$/, "").trim();
  return value || null;
}

// Tách file nhiều trang thành từng khối cho từng bị can. Thử lần lượt
// nhiều mốc khác nhau vì số văn bản không phải lúc nào cũng đọc được.
// Hỗ trợ cả mẫu 74 (Lệnh bắt bị can để tạm giam) và mẫu 77 (Quyết định
// tạm giữ) — chỉ khác nhau ở nhãn văn bản, cấu trúc phần thân giống nhau.
function splitLenhBlocks(fullText: string): string[] {
  const flat = deaccent(fullText);
  const markers = [
    /\d{1,6}\s*\/\s*(?:LB|QD)[-\s]*[A-Z]{2,10}/g,
    /Lenh\s*bat\s*bi\s*can\s*de\s*tam\s*giam/gi,
    /Quyet\s*dinh\s*tam\s*giu/gi,
    /CONG\s*HOA\s*XA\s*HOI\s*CHU\s*NGHIA/gi,
    /Ho\s*(?:va\s*)?ten\s*:/gi,
  ];
  for (const marker of markers) {
    const indices: number[] = [];
    let m: RegExpExecArray | null;
    marker.lastIndex = 0;
    while ((m = marker.exec(flat)) !== null) indices.push(m.index);
    if (indices.length >= 1) {
      return indices.map((start, i) =>
        fullText.slice(start, i + 1 < indices.length ? indices[i + 1] : fullText.length)
      );
    }
  }
  return [fullText];
}

interface Extracted {
  fields: Partial<Record<FieldKey, string>>;
  batNote: string;
}

function extractFromBlock(block: string): Extracted {
  const f: Partial<Record<FieldKey, string>> = {};
  const flat = deaccent(block);
  let batNote = "";

  const fmtDate = (d: string, m: string, y: string) => `${pad2(d)}/${pad2(m)}/${y}`;

  // Họ tên
  const hoTen = sliceAfterLabel(
    block,
    flat,
    /Ho\s*(?:va\s*)?ten\s*:?/,
    [/Gioi\s*tinh/, /Ten\s*goi\s*khac/, /Sinh\s*ngay/, /Sinh\s*nam/, /;/],
    120
  );
  if (hoTen) f.hoTen = hoTen;

  // Giới tính
  const gt = /Gioi\s*tinh\s*:?\s*(Nam|Nu)\b/i.exec(flat);
  if (gt) f.gioiTinh = gt[1].toLowerCase() === "nam" ? "Nam" : "Nữ";

  // Sinh ngày (đủ ngày/tháng/năm, hoặc chỉ có năm)
  const sinh =
    /Sinh\s*(?:ngay)?\s*:?\s*(\d{1,2})\s*(?:thang|\/|-|\.)\s*(\d{1,2})\s*(?:nam|\/|-|\.)\s*(\d{4})/i.exec(
      flat
    );
  if (sinh) {
    f.sinhNgay = fmtDate(sinh[1], sinh[2], sinh[3]);
  } else {
    const sinhNam = /Sinh\s*(?:nam)?\s*:?\s*(\d{4})\b/i.exec(flat);
    if (sinhNam) f.sinhNgay = `01/01/${sinhNam[1]}`;
  }

  // Quốc tịch / Dân tộc / Tôn giáo
  const qt = sliceAfterLabel(block, flat, /Quoc\s*tich\s*:?/, [/;/, /Dan\s*toc/, /Ton\s*giao/, /Nghe\s*nghiep/], 60);
  if (qt) f.quocTich = qt;

  const dt = sliceAfterLabel(block, flat, /Dan\s*toc\s*:?/, [/;/, /Ton\s*giao/, /Quoc\s*tich/, /Nghe\s*nghiep/], 60);
  if (dt) f.danToc = dt;

  const tg = sliceAfterLabel(
    block,
    flat,
    /Ton\s*giao\s*:?/,
    [/;/, /Nghe\s*nghiep/, /Quoc\s*tich/, /Dan\s*toc/, /Trinh\s*do/, /So\s*CMND/, /The\s*CCCD/],
    60
  );
  if (tg) f.tonGiao = tg;

  // CCCD: bắt mọi biến thể nhãn rồi lọc chữ số
  const cc =
    /(?:The\s*CCCD|CCCD|CMND|CMTND|Can\s*cuoc\s*cong\s*dan|Ho\s*chieu)[^0-9]{0,40}(\d[\d .]{7,22})/i.exec(
      flat
    );
  if (cc) f.cccd = cc[1].replace(/\D/g, "");

  // Nơi thường trú
  const tt = sliceAfterLabel(
    block,
    flat,
    /Noi\s*(?:dang\s*ky\s*)?thuong\s*tru\s*:?|DKTT\s*:?/,
    [/Noi\s*o\s*hien\s*tai/, /Cho\s*o\s*hien/, /Nghe\s*nghiep/, /Quoc\s*tich/, /Ho\s*ten\s*cha/, /Tien\s*an/],
    250
  );
  if (tt) f.noiThuongTru = normalizeAddress(tt);

  // Cột 10 — Lệnh: nhận diện văn bản là Lệnh bắt bị can (mẫu 74, /LB-...)
  // hay Quyết định tạm giữ (mẫu 77, /QĐ-...) để dựng đúng câu chữ.
  const isTamGiu = /Quyet\s*dinh\s*tam\s*giu/i.test(flat);
  const soLenh =
    /So\s*:?\s*(\d{1,6})\s*\/\s*((?:LB|QD)[-\s]*[A-Z0-9]{2,10})/i.exec(flat) ||
    /(\d{1,6})\s*\/\s*((?:LB|QD)[-\s]*[A-Z0-9]{2,10})/i.exec(flat);
  const ngayBanHanh =
    /Ha\s*Noi\s*,?\s*ngay\s*(\d{1,2})\s*thang\s*(\d{1,2})\s*nam\s*(\d{4})/i.exec(flat) ||
    /ngay\s*(\d{1,2})\s*thang\s*(\d{1,2})\s*nam\s*(\d{4})/i.exec(flat);
  if (soLenh || ngayBanHanh) {
    const soPart = soLenh
      ? `${soLenh[1]}/${soLenh[2].replace(/\s+/g, "").toUpperCase()}`
      : "[chưa đọc được số]";
    const ngayPart = ngayBanHanh
      ? fmtDate(ngayBanHanh[1], ngayBanHanh[2], ngayBanHanh[3])
      : "[chưa đọc được ngày]";
    const tenVanBan = isTamGiu ? "Quyết định tạm giữ" : "Lệnh bắt bị can để tạm giam";
    f.lenh = `${tenVanBan} số: ${soPart}, ngày ${ngayPart} của Cơ quan CSĐT Bộ Công an`;
  }

  // Cột 11 — Bị bắt ngày: ưu tiên theo thứ tự
  // - mục ghi chú viết tay "Bắt bị can ngày..." (mẫu 74)
  // - đối chiếu với dòng giao văn bản cho bị can/người bị tạm giữ
  // - nếu cả hai đều không đọc được (viết tay bị bỏ trống hoặc chưa
  //   giao), tạm lấy theo ngày ban hành của chính Lệnh/Quyết định —
  //   đây chỉ là suy đoán gần đúng nên luôn ghi chú rõ để người dùng
  //   tự xác nhận lại, không âm thầm coi là chắc chắn.
  const ghiChu =
    /Bat\s*bi\s*can\s*(?:vao\s*)?(?:hoi\s*[\d\sgiophut:.]{0,12})?ngay\s*(\d{1,2})\s*(?:thang|\/|-|\.)\s*(\d{1,2})\s*(?:nam|\/|-|\.)\s*(\d{4})/i.exec(
      flat
    );
  const giaoLenh =
    /giao\s*cho\s*(?:bi\s*can|nguoi\s*bi\s*tam\s*giu)[\s\S]{0,250}?ngay\s*(\d{1,2})\s*(?:thang|\/|-|\.)\s*(\d{1,2})\s*(?:nam|\/|-|\.)\s*(\d{4})/i.exec(
      flat
    );
  const g1 = ghiChu ? fmtDate(ghiChu[1], ghiChu[2], ghiChu[3]) : null;
  const g2 = giaoLenh ? fmtDate(giaoLenh[1], giaoLenh[2], giaoLenh[3]) : null;
  const g3 = ngayBanHanh ? fmtDate(ngayBanHanh[1], ngayBanHanh[2], ngayBanHanh[3]) : null;

  if (g1) {
    f.biBatNgay = g1;
    if (g2 && g2 !== g1) batNote = `(lệch với ngày giao lệnh ${g2} — cần đối chiếu)`;
  } else if (g2) {
    f.biBatNgay = g2;
    batNote = "(lấy theo ngày giao lệnh, mục ghi chú viết tay chưa đọc được)";
  } else if (g3) {
    f.biBatNgay = g3;
    const tenVB = isTamGiu ? "quyết định" : "lệnh";
    batNote = `(chưa đọc được ngày viết tay — tạm lấy theo ngày ban hành ${tenVB}, cần xác nhận lại)`;
  }

  return { fields: f, batNote };
}

// Một khối được coi là hợp lệ khi đọc được ít nhất một dấu hiệu nhận dạng.
function hasUsableData(ex: Extracted): boolean {
  const f = ex.fields;
  return Boolean(f.hoTen || f.cccd || f.lenh || f.noiThuongTru);
}

interface PendingItem {
  draft: Draft;
  batNote: string;
}

// Dữ liệu đọc từ PDF -> giá trị điền thẳng vào form.
// Dữ liệu đọc từ PDF/ảnh -> giá trị điền thẳng vào form. Các trường
// "sinhNgay"/"biBatNgay" đã ở dạng dd/mm/yyyy sẵn (kiểu "datestr") nên
// giữ nguyên, không cần đổi sang yyyy-MM-dd.
function toDraft(ex: Extracted): PendingItem {
  const draft = ALL_FIELD_KEYS.reduce((acc, k) => {
    const raw = ex.fields[k] ?? DEFAULTS[k] ?? "";
    acc[k] = DATE_KEYS.includes(k) ? toIsoDate(raw) : raw;
    return acc;
  }, {} as Draft);
  draft.cccd = parseCccd(draft.cccd).value;
  return { draft, batNote: ex.batNote };
}

// Đọc lớp văn bản có sẵn (nếu là PDF được soạn thảo, không phải scan).
async function extractTextLayer(pdf: pdfjsLib.PDFDocumentProxy): Promise<string> {
  let fullText = "";
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const pageText = content.items
      .map((it) => {
        if (!("str" in it)) return "";
        const item = it as { str: string; hasEOL?: boolean };
        return item.str + (item.hasEOL ? " " : "");
      })
      .join("");
    fullText += pageText + " ";
  }
  return fullText;
}

// File scan ảnh (mộc dấu, chữ ký, chụp lại...) không có lớp văn bản
// nhúng sẵn, nên getTextContent() trả về gần như rỗng. Trường hợp này
// phải render từng trang thành ảnh rồi nhận dạng chữ (OCR) bằng
// tesseract.js, ngôn ngữ tiếng Việt ("vie"). Chỉ chạy khi thật sự cần,
// vì OCR chậm hơn nhiều so với đọc lớp văn bản có sẵn.
let ocrWorkerPromise: Promise<TesseractWorker> | null = null;
function getOcrWorker(): Promise<TesseractWorker> {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = createWorker("vie");
  }
  return ocrWorkerPromise;
}

async function extractTextByOcr(
  pdf: pdfjsLib.PDFDocumentProxy,
  onProgress?: (page: number, total: number) => void
): Promise<string> {
  const worker = await getOcrWorker();
  // Đặt cùng chế độ phân đoạn với nhánh OCR ảnh (xem giải thích ở
  // extractTextFromImage) vì worker dùng chung, tránh kết quả phụ
  // thuộc vào việc hàm nào chạy trước.
  await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_BLOCK });
  let fullText = "";
  for (let p = 1; p <= pdf.numPages; p++) {
    onProgress?.(p, pdf.numPages);
    const page = await pdf.getPage(p);
    // Scale 2.5x để chữ in trên form rõ hơn cho OCR, ảnh hưởng tốc độ
    // không đáng kể vì mỗi trang chỉ mất vài giây.
    const viewport = page.getViewport({ scale: 2.5 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext("2d");
    if (!ctx) continue;
    await page.render({ canvasContext: ctx, viewport, canvas }).promise;
    const { data } = await worker.recognize(canvas);
    fullText += data.text + " ";
  }
  return fullText;
}

function cleanupText(raw: string): string {
  // Phải xử lý ký tự ẩn TRƯỚC khi normalize NFC: nếu một ký tự ẩn nằm
  // xen giữa chữ cái gốc và dấu tổ hợp (dạng NFD mà PDF hay xuất ra),
  // normalize chạy trước sẽ không ghép được dấu đúng.
  // NBSP (\u00A0) thực sự đóng vai trò dấu cách nên đổi thành " ".
  // Ký tự rộng-0 (zero-width) và BOM không có độ rộng hiển thị — nếu
  // đổi thành " " sẽ vô tình chèn dấu cách giữa hai ký tự vốn liền
  // nhau (ví dụ tách đôi một dãy số CCCD), nên phải xoá hẳn, không
  // thay bằng khoảng trắng.
  return raw
    .replace(/(?:\u200B|\u200C|\u200D|\uFEFF)/g, "")
    .replace(/\u00A0/g, " ")
    .normalize("NFC")
    .replace(/\s+/g, " ")
    .trim();
}

interface PdfExtraction {
  text: string;
  usedOcr: boolean;
}

async function extractTextFromPdf(
  file: File,
  onOcrProgress?: (page: number, total: number) => void
): Promise<PdfExtraction> {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

  const layerText = cleanupText(await extractTextLayer(pdf));
  // Ngưỡng thực nghiệm: mỗi trang mẫu 74/77 gõ máy có vài trăm ký tự.
  // Dưới ~40 ký tự/trang gần như chắc chắn là ảnh scan không có chữ.
  const perPageThreshold = 40 * pdf.numPages;
  if (layerText.length >= perPageThreshold) {
    return { text: layerText, usedOcr: false };
  }

  const ocrText = cleanupText(await extractTextByOcr(pdf, onOcrProgress));
  return { text: ocrText, usedOcr: true };
}

// Ảnh chụp bằng điện thoại thường có độ phân giải hiệu dụng thấp cho
// phần chữ nhỏ trên form (Họ tên, Giới tính, Sinh ngày, Nơi thường trú
// chỉ chiếm vài chục pixel chiều cao), lại lẫn ánh sáng không đều,
// mộc dấu, QR code quanh đó. Phóng to lên một kích thước tối thiểu và
// chuyển sang thang xám kéo giãn tương phản giúp Tesseract đọc chữ nhỏ
// tốt hơn nhiều so với đưa thẳng ảnh gốc vào — đúng như cách PDF scan
// đã được phóng to 2.5x trước khi OCR ở trên.
async function preprocessImageForOcr(file: File): Promise<HTMLCanvasElement> {
  const bitmap = await createImageBitmap(file);
  const MIN_WIDTH = 1800;
  const scale = bitmap.width < MIN_WIDTH ? MIN_WIDTH / bitmap.width : 1;
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  const ctx = canvas.getContext("2d");
  if (!ctx) return canvas;
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);

  // Thang xám + kéo giãn tương phản đơn giản (min-max stretch) để chữ
  // in mờ/ám vàng của ảnh scan/chụp nổi rõ hơn so với nền giấy.
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = imageData.data;
  let min = 255;
  let max = 0;
  for (let i = 0; i < d.length; i += 4) {
    const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    d[i] = d[i + 1] = d[i + 2] = gray;
    if (gray < min) min = gray;
    if (gray > max) max = gray;
  }
  const range = Math.max(1, max - min);
  for (let i = 0; i < d.length; i += 4) {
    const stretched = ((d[i] - min) / range) * 255;
    d[i] = d[i + 1] = d[i + 2] = stretched;
  }
  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

// Ảnh chụp/scan (jpg, png...) không có khái niệm "lớp văn bản" — luôn
// phải OCR trực tiếp trên ảnh.
async function extractTextFromImage(
  file: File,
  onOcrProgress?: (page: number, total: number) => void
): Promise<PdfExtraction> {
  onOcrProgress?.(1, 1);
  const worker = await getOcrWorker();
  const canvas = await preprocessImageForOcr(file);
  // PSM 6: coi cả ảnh là một khối văn bản đồng nhất. Form có bố cục
  // nhãn/giá trị nằm cạnh nhau (vd "Họ tên: ... Giới tính: ...") khiến
  // chế độ tự động phân đoạn mặc định của Tesseract dễ nhảy cột hoặc
  // đọc sai thứ tự dòng, làm rớt đúng các trường này.
  await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_BLOCK });
  const { data } = await worker.recognize(canvas);
  return { text: cleanupText(data.text), usedOcr: true };
}

const IMAGE_TYPES = ["image/jpeg", "image/png", "image/webp", "image/bmp"];

// Điểm vào chung: chọn cách đọc theo loại file (ảnh -> OCR thẳng,
// PDF -> đọc lớp văn bản trước rồi mới OCR nếu cần).
async function extractTextFromFile(
  file: File,
  onOcrProgress?: (page: number, total: number) => void
): Promise<PdfExtraction> {
  const isImage = IMAGE_TYPES.includes(file.type) || /\.(jpe?g|png|webp|bmp)$/i.test(file.name);
  if (isImage) return extractTextFromImage(file, onOcrProgress);
  return extractTextFromPdf(file, onOcrProgress);
}

// ---------------------------------------------------------------------------
// COMPONENT
// ---------------------------------------------------------------------------

export default function LenhBatBiCanForm() {
  const [records, setRecords] = useState<BiCanRecord[]>([]);
  const [draft, setDraft] = useState<Draft>(emptyDraft);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [isParsingPdf, setIsParsingPdf] = useState(false);
  const [pdfImportSummary, setPdfImportSummary] = useState<string | null>(null);
  const [draftBatNote, setDraftBatNote] = useState("");
  const [pendingQueue, setPendingQueue] = useState<PendingItem[]>([]);
  const [rawPdfText, setRawPdfText] = useState("");
  const [selectedRecord, setSelectedRecord] = useState<BiCanRecord | null>(null);

  const filteredRecords = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    if (!q) return records;
    return records.filter((r) =>
      [r.hoTen, r.cccd, r.noiThuongTru, r.lenh].some((v) => (v ?? "").toLowerCase().includes(q))
    );
  }, [records, searchTerm]);

  const missingFieldLabels = useMemo<Record<string, string[]>>(() => {
    const report: Record<string, string[]> = {};
    records.forEach((r) => {
      const missing = ALL_FIELD_KEYS.filter((k) => !r[k]?.trim()).map((k) => LABEL_MAP[k]);
      if (missing.length) report[r.hoTen || `(Bản ghi #${r._id})`] = missing;
    });
    return report;
  }, [records]);

  function updateDraft(key: FieldKey, value: string) {
    setDraft((d) => ({ ...d, [key]: value }));
  }

  function handleFieldChange(key: FieldKey) {
    return (e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) =>
      updateDraft(key, e.target.value);
  }

  function resetDraft() {
    setDraft(emptyDraft);
    setDraftBatNote("");
    setEditingId(null);
  }

  // Nạp hồ sơ kế tiếp đã đọc từ PDF vào form; hết hàng đợi thì để form trống.
  function loadNextPending() {
    setEditingId(null);
    setPendingQueue((q) => {
      const [next, ...rest] = q;
      if (next) {
        setDraft(next.draft);
        setDraftBatNote(next.batNote);
      } else {
        setDraft(emptyDraft);
        setDraftBatNote("");
      }
      return rest;
    });
  }

  function saveDraft() {
    if (!draft.hoTen.trim()) return;
    const { value: cccdValue, note } = parseCccd(draft.cccd);
    const row = ALL_FIELD_KEYS.reduce((acc, k) => {
      if (k === "cccd") acc[k] = cccdValue;
      else if (k === "sinhNgay" || k === "biBatNgay") acc[k] = normalizeDateStr(draft[k]);
      else if (DATE_KEYS.includes(k)) acc[k] = toDisplayDate(draft[k]);
      else if (k === "noiThuongTru") acc[k] = normalizeAddress(draft[k]);
      else acc[k] = draft[k].trim();
      return acc;
    }, {} as Draft);

    if (editingId !== null) {
      setRecords((rs) =>
        rs.map((r) => (r._id === editingId ? { ...r, ...row, cccdNote: note, batNote: draftBatNote } : r))
      );
    } else {
      setRecords((rs) => [
        ...rs,
        { _id: Date.now(), x: "GP", cccdNote: note, batNote: draftBatNote, ...row },
      ]);
    }
    loadNextPending();
  }

  function editRow(r: BiCanRecord) {
    setSelectedRecord(null);
    setEditingId(r._id);
    setDraftBatNote(r.batNote);
    setDraft(
      ALL_FIELD_KEYS.reduce((acc, k) => {
        acc[k] = DATE_KEYS.includes(k) ? toIsoDate(r[k]) : r[k];
        return acc;
      }, {} as Draft)
    );
  }

  function removeRow(id: number) {
    setRecords((rs) => rs.filter((r) => r._id !== id));
    if (editingId === id) resetDraft();
    if (selectedRecord?._id === id) setSelectedRecord(null);
  }

  async function handleDocFiles(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return;
    setIsParsingPdf(true);
    setPdfImportSummary(null);
    setRawPdfText("");
    try {
      const items: PendingItem[] = [];
      let rawAll = "";
      let anyOcr = false;

      for (const file of Array.from(fileList)) {
        const { text, usedOcr } = await extractTextFromFile(file, (page, total) => {
          setPdfImportSummary(
            total > 1
              ? `Không thấy chữ gõ máy trong "${file.name}" — đang nhận dạng chữ (OCR) trang ${page}/${total}...`
              : `Đang nhận dạng chữ (OCR) trong "${file.name}"...`
          );
        });
        anyOcr = anyOcr || usedOcr;
        rawAll += `\n===== ${file.name}${usedOcr ? " (đã OCR)" : ""} (${text.length} ký tự) =====\n${text}\n`;
        splitLenhBlocks(text).forEach((block) => {
          const ex = extractFromBlock(block);
          if (!hasUsableData(ex)) return;
          items.push(toDraft(ex));
        });
      }
      setRawPdfText(rawAll.trim());

      if (items.length > 0) {
        const [first, ...rest] = items;
        setEditingId(null);
        setDraft(first.draft);
        setDraftBatNote(first.batNote);
        setPendingQueue(rest);
        setPdfImportSummary(
          (anyOcr ? "Đã nhận dạng chữ từ ảnh scan (OCR). " : "") +
            (rest.length === 0
              ? 'Đã điền 1 hồ sơ vào form bên dưới. Kiểm tra lại (nhất là các mục viết tay) rồi bấm "Thêm vào bảng".'
              : `Đọc được ${items.length} hồ sơ. Hồ sơ đầu tiên đã điền sẵn vào form; mỗi lần bấm "Thêm vào bảng" sẽ tự nạp hồ sơ kế tiếp.`)
        );
      } else if (rawAll.replace(/=+[^\n]*\n/g, "").trim().length < 200) {
        setPdfImportSummary(
          "Không đọc được chữ nào kể cả sau khi OCR. Ảnh scan có thể quá mờ hoặc nghiêng — thử chụp/scan lại rõ hơn."
        );
      } else {
        setPdfImportSummary(
          "Đã đọc được văn bản (kể cả bằng OCR) nhưng không khớp nhãn nào của mẫu 74. Mở mục \"Xem văn bản đọc được\" bên dưới để kiểm tra nhãn thực tế trong file."
        );
      }
    } catch (err) {
      console.error(err);
      setPdfImportSummary("Không đọc được file PDF. Thử lại hoặc nhập tay.");
    } finally {
      setIsParsingPdf(false);
    }
  }


  function exportExcel() {
    if (!records.length) return;

    const colIndexes = EXPORT_FIELDS.map((f) => colLetterToIndex(f.col));
    const totalCols = Math.max(...colIndexes) + 1;

    const emptyRow = () => Array(totalCols).fill("");

    const header = emptyRow();
    EXPORT_FIELDS.forEach((f, idx) => {
      header[colIndexes[idx]] = f.label;
    });

    const body = records.map((r) => {
      const row = emptyRow();
      EXPORT_FIELDS.forEach((f, idx) => {
        if (f.getValue) row[colIndexes[idx]] = f.getValue(r);
        else if (f.key) row[colIndexes[idx]] = cellValue(r, f.key);
        // Không có key/getValue (vd "Quê quán") -> để trống, không suy diễn.
      });
      return row;
    });

    const ws = XLSX.utils.aoa_to_sheet([header, ...body]);

    // Ép các cột cần giữ nguyên dạng text (không mất số 0 đầu, vd CCCD).
    EXPORT_FIELDS.forEach((f, idx) => {
      if (!f.isText) return;
      records.forEach((_, i) => {
        const addr = XLSX.utils.encode_cell({ r: i + 1, c: colIndexes[idx] });
        if (ws[addr]) {
          ws[addr].t = "s";
          ws[addr].z = "@";
        }
      });
    });

    const cols = Array(totalCols).fill({ wch: 4 });
    EXPORT_FIELDS.forEach((f, idx) => {
      cols[colIndexes[idx]] = { wch: Math.max(14, f.label.length + 4) };
    });
    ws["!cols"] = cols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Danh sách");
    XLSX.writeFile(wb, "danh_sach_lenh_bat.xlsx");
  }

  function renderInput(cfg: FieldConfig) {
    const value = draft[cfg.key];
    if (cfg.type === "textarea")
      return (
        <textarea
          className="input"
          rows={2}
          value={value}
          placeholder={cfg.placeholder}
          onChange={handleFieldChange(cfg.key)}
        />
      );
    if (cfg.type === "date")
      return <input type="date" className="input" value={value} onChange={handleFieldChange(cfg.key)} />;
    if (cfg.type === "datestr")
      return (
        <input
          type="text"
          inputMode="numeric"
          className="input"
          value={value}
          placeholder={cfg.placeholder}
          maxLength={10}
          onChange={(e) => updateDraft(cfg.key, maskDateTyping(e.target.value))}
        />
      );
    if (cfg.type === "select")
      return (
        <select className="input" value={value} onChange={handleFieldChange(cfg.key)}>
          <option value="">— Chọn —</option>
          {cfg.selectOptions?.map((o) => (
            <option key={o} value={o}>
              {o}
            </option>
          ))}
        </select>
      );
    const listId = cfg.datalistOptions ? `dl-${cfg.key}` : undefined;
    return (
      <>
        <input
          className="input"
          value={value}
          placeholder={cfg.placeholder}
          list={listId}
          onChange={handleFieldChange(cfg.key)}
        />
        {cfg.datalistOptions && (
          <datalist id={listId}>
            {cfg.datalistOptions.map((o) => (
              <option key={o} value={o} />
            ))}
          </datalist>
        )}
      </>
    );
  }

  return (
    <div className="min-h-screen bg-[#F4F3EF] text-[#1B2A4A]">
      <div className="mx-auto max-w-6xl px-6 py-10">
        <header className="mb-8 border-b border-[#D8D6CC] pb-6">
          <div className="flex flex-wrap items-end justify-between gap-4">
            <div>
              <p className="text-xs tracking-wide text-[#8A6A3B]">Cơ quan Cảnh sát điều tra</p>
              <h1 className="mt-1 font-serif text-3xl">BẢNG DỮ LIỆU LỆNH TẠM GIAM BỊ CAN</h1>
              <p className="mt-2 max-w-xl text-sm text-[#5B5B54]">
                11 cột cố định, kèm cột X mặc định "GP". Trường không có
                trên văn bản thì để trống — hệ thống tô đỏ, không suy diễn.
              </p>
            </div>
            <div className="flex gap-2">
              <StatChip label="Hồ sơ" value={records.length} />
              <StatChip
                label="Thiếu dữ liệu"
                value={Object.keys(missingFieldLabels).length}
                tone={Object.keys(missingFieldLabels).length > 0 ? "warn" : "ok"}
              />
            </div>
          </div>
        </header>

        <section className="mb-8 rounded-sm border border-dashed border-[#8A6A3B]/40 bg-[#FBF9F2] p-6 sm:p-7">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-start">
            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-full bg-[#8A6A3B]/10 text-[#8A6A3B]">
              <Upload size={18} />
            </div>
            <div className="flex-1">
              <h2 className="font-serif text-lg">Đọc từ file PDF / ảnh</h2>
              <p className="mb-4 mt-1 text-sm text-[#5B5B54]">
                Chọn một hoặc nhiều file "Lệnh bắt bị can để tạm giam" (mẫu số
                74) hoặc "Quyết định tạm giữ" (mẫu số 77) — nhận cả PDF (gõ
                máy hoặc scan) lẫn ảnh chụp/scan jpg, png. Hệ thống tự nhận
                dạng chữ bằng OCR khi cần, có thể mất vài giây mỗi trang/ảnh.
                Nội dung đọc được sẽ tự điền vào form bên dưới.
              </p>
              <label className="btn-secondary inline-flex cursor-pointer">
                {isParsingPdf ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
                {isParsingPdf ? "Đang đọc..." : "Chọn file PDF hoặc ảnh"}
                <input
                  type="file"
                  accept="application/pdf,image/jpeg,image/png,image/webp,image/bmp"
                  multiple
                  disabled={isParsingPdf}
                  className="hidden"
                  onChange={(e) => {
                    handleDocFiles(e.target.files);
                    e.target.value = "";
                  }}
                />
              </label>
              {pdfImportSummary && (
                <p className="mt-3 flex items-start gap-2 rounded-sm bg-white px-3 py-2 text-sm">
                  <Info size={15} className="mt-0.5 shrink-0 text-[#8A6A3B]" />
                  {pdfImportSummary}
                </p>
              )}
              {rawPdfText && (
                <details className="mt-3 rounded-sm bg-white px-3 py-2 text-sm">
                  <summary className="cursor-pointer text-[#8A6A3B]">
                    Xem văn bản đọc được từ file
                  </summary>
                  <pre className="mt-2 max-h-64 overflow-auto whitespace-pre-wrap wrap-break-word text-xs text-[#5B5B54]">
                    {rawPdfText}
                  </pre>
                </details>
              )}
            </div>
          </div>
        </section>

        <section className="mb-10 rounded-sm border border-[#D8D6CC] bg-white p-6 sm:p-7">
          <div className="mb-6 flex items-center justify-between border-b border-[#EDEBE3] pb-4">
            <h2 className="flex items-center gap-2 font-serif text-lg">
              <FileText size={18} className="text-[#8A6A3B]" />
              {editingId !== null ? "Sửa dòng dữ liệu" : "Thêm dòng mới"}
            </h2>
            {editingId !== null && (
              <span className="rounded-sm bg-[#FBF3E4] px-2.5 py-1 text-xs font-medium text-[#8A6A3B]">
                Đang sửa
              </span>
            )}
          </div>

          {FIELD_SECTIONS.map((section, idx) => (
            <div key={section.title} className={idx > 0 ? "mt-6" : undefined}>
              <h3 className="mb-3 text-xs font-semibold uppercase tracking-wide text-[#8A6A3B]">
                {section.title}
              </h3>
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
                {section.keys.map((k) => {
                  const cfg = FIELD_MAP[k];
                  return (
                    <Field key={k} label={cfg.label} full={cfg.full}>
                      {renderInput(cfg)}
                    </Field>
                  );
                })}
              </div>
            </div>
          ))}

          <div className="mt-6 flex flex-wrap items-center gap-3 border-t border-[#EDEBE3] pt-5">
            <button onClick={saveDraft} className="btn-primary">
              {editingId !== null ? <Pencil size={16} /> : <Plus size={16} />}
              {editingId !== null ? "Lưu thay đổi" : "Thêm vào bảng"}
            </button>
            {editingId !== null && (
              <button onClick={resetDraft} className="btn-secondary">
                <XIcon size={16} /> Huỷ sửa
              </button>
            )}
            {editingId === null && pendingQueue.length > 0 && (
              <>
                <button onClick={loadNextPending} className="btn-secondary">
                  <XIcon size={16} /> Bỏ qua hồ sơ này
                </button>
                <span className="text-sm text-[#5B5B54]">
                  Còn {pendingQueue.length} hồ sơ chờ kiểm tra.
                </span>
              </>
            )}
          </div>
        </section>

        <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
          <div className="relative w-full max-w-sm">
      
            <input
              className="input pl-9"
              placeholder="Tìm theo họ tên, CCCD, địa chỉ, số lệnh..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <p className="text-sm text-[#8A8A80]">
            {filteredRecords.length}/{records.length} dòng · bấm vào một dòng để xem đầy đủ
          </p>
        </div>

        <section className="mb-8 overflow-hidden rounded-sm border border-[#D8D6CC] bg-white">
          <div className="overflow-x-auto">
            <table className="w-full min-w-205 border-collapse text-sm">
              <thead>
                <tr className="bg-[#1B2A4A] text-left text-white">
                  {TABLE_COLUMNS.map((c) => (
                    <th key={c.key} className="px-3 py-2.5 font-semibold">
                      {c.label}
                    </th>
                  ))}
                  <th className="px-3 py-2.5 font-semibold">Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {filteredRecords.length === 0 && (
                  <tr>
                    <td
                      colSpan={TABLE_COLUMNS.length + 1}
                      className="px-3 py-10 text-center text-[#8A8A80]"
                    >
                      {records.length === 0
                        ? "Chưa có dữ liệu. Đọc file PDF/ảnh hoặc nhập tay ở form phía trên."
                        : "Không có dòng nào khớp từ khoá."}
                    </td>
                  </tr>
                )}
                {filteredRecords.map((r, i) => (
                  <tr
                    key={r._id}
                    onClick={() => setSelectedRecord(r)}
                    className="cursor-pointer border-t border-[#EDEBE3] transition-colors hover:bg-[#F4F3EF]"
                  >
                    {TABLE_COLUMNS.map((c) => {
                      if (c.key === "stt")
                        return (
                          <td key="stt" className="px-3 py-2.5 text-[#8A8A80]">
                            {i + 1}
                          </td>
                        );
                      if (c.key === "x")
                        return (
                          <td key="x" className="px-3 py-2.5">
                            <span className="badge-x">{r.x}</span>
                          </td>
                        );
                      const val = r[c.key];
                      const isEmpty = !String(val ?? "").trim();
                      const note = c.key === "cccd" ? r.cccdNote : "";
                      return (
                        <td
                          key={c.key}
                          className={`max-w-55 truncate px-3 py-2.5 ${
                            isEmpty ? "text-[#B3261E]" : ""
                          } ${c.key === "hoTen" ? "font-medium" : ""}`}
                        >
                          {isEmpty ? "— thiếu" : val}
                          {!isEmpty && note && (
                            <AlertTriangle
                              size={12}
                              className="ml-1 inline-block shrink-0 text-[#B3261E]"
                            />
                          )}
                        </td>
                      );
                    })}
                    <td className="px-3 py-2.5" onClick={(e) => e.stopPropagation()}>
                      <div className="flex gap-2">
                        <button onClick={() => editRow(r)} className="icon-btn" aria-label="Sửa">
                          <Pencil size={14} />
                        </button>
                        <button onClick={() => removeRow(r._id)} className="icon-btn" aria-label="Xoá">
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <div className="mb-10 flex items-center justify-between">
          <p className="text-sm text-[#5B5B54]">{records.length} dòng trong bảng.</p>
          <button
            onClick={exportExcel}
            disabled={!records.length}
            className="btn-primary disabled:cursor-not-allowed disabled:opacity-40"
          >
            <Download size={16} /> Xuất file Excel (.xlsx)
          </button>
        </div>

        <section className="rounded-sm border border-[#D8D6CC] bg-white p-6">
          <h2 className="mb-4 flex items-center gap-2 font-serif text-lg">
            <AlertTriangle size={18} className="text-[#B3261E]" />
            Báo cáo dữ liệu
          </h2>
          {records.length === 0 ? (
            <p className="text-sm text-[#8A8A80]">Chưa có dữ liệu để kiểm tra.</p>
          ) : Object.keys(missingFieldLabels).length === 0 ? (
            <p className="text-sm text-[#5B5B54]">Không thiếu dữ liệu.</p>
          ) : (
            <ul className="space-y-3 text-sm">
              {records
                .filter((r) => missingFieldLabels[r.hoTen || `(Bản ghi #${r._id})`])
                .map((r) => {
                  const name = r.hoTen || `(Bản ghi #${r._id})`;
                  const fields = missingFieldLabels[name];
                  return (
                    <li key={r._id} className="border-b border-[#EDEBE3] pb-3 last:border-0 last:pb-0">
                      <button
                        onClick={() => setSelectedRecord(r)}
                        className="font-medium text-[#1B2A4A] underline decoration-[#D8D6CC] underline-offset-2 hover:decoration-[#1B2A4A]"
                      >
                        {name}
                      </button>
                      <div className="mt-1.5 flex flex-wrap gap-1.5">
                        {fields.map((f) => (
                          <span key={f} className="rounded-sm bg-[#FBE1DF] px-2 py-0.5 text-xs text-[#B3261E]">
                            {f}
                          </span>
                        ))}
                      </div>
                    </li>
                  );
                })}
            </ul>
          )}
        </section>
      </div>

      {selectedRecord && (
        <DetailModal
          record={selectedRecord}
          onClose={() => setSelectedRecord(null)}
          onEdit={() => editRow(selectedRecord)}
          onDelete={() => removeRow(selectedRecord._id)}
        />
      )}

      <style>{`
        .input { width:100%; border:1px solid #D8D6CC; background:#FDFCFA; border-radius:2px;
          padding:0.5rem 0.65rem; font-size:0.875rem; color:#1B2A4A; outline:none; }
        .input:focus { border-color:#1B2A4A; }
        .btn-primary { display:inline-flex; align-items:center; gap:0.4rem; background:#1B2A4A;
          color:#fff; padding:0.55rem 1.1rem; border-radius:2px; font-size:0.875rem; font-weight:500; }
        .btn-primary:hover { background:#22355E; }
        .btn-secondary { display:inline-flex; align-items:center; gap:0.4rem; background:transparent;
          color:#1B2A4A; border:1px solid #D8D6CC; padding:0.55rem 1.1rem; border-radius:2px; font-size:0.875rem; }
        .btn-secondary:hover { border-color:#1B2A4A; }
        .icon-btn { display:inline-flex; align-items:center; justify-content:center; width:26px; height:26px;
          border:1px solid #D8D6CC; border-radius:2px; color:#5B5B54; }
        .icon-btn:hover { border-color:#1B2A4A; color:#1B2A4A; }
        .badge-x { display:inline-flex; align-items:center; justify-content:center; min-width:26px;
          padding:0.1rem 0.4rem; border-radius:2px; background:#1B2A4A; color:#fff; font-size:0.75rem;
          font-weight:600; letter-spacing:0.02em; }
      `}</style>
    </div>
  );
}

function StatChip({
  label,
  value,
  tone = "neutral",
}: {
  label: string;
  value: number;
  tone?: "neutral" | "warn" | "ok";
}) {
  const toneClass =
    tone === "warn"
      ? "border-[#B3261E]/30 bg-[#FBE1DF] text-[#B3261E]"
      : tone === "ok"
      ? "border-[#1B2A4A]/15 bg-[#F4F3EF] text-[#5B5B54]"
      : "border-[#D8D6CC] bg-white text-[#1B2A4A]";
  return (
    <div className={`rounded-sm border px-4 py-2 text-right ${toneClass}`}>
      <p className="font-serif text-xl leading-none">{value}</p>
      <p className="mt-1 text-xs text-[#5B5B54]">{label}</p>
    </div>
  );
}

function DetailModal({
  record,
  onClose,
  onEdit,
  onDelete,
}: {
  record: BiCanRecord;
  onClose: () => void;
  onEdit: () => void;
  onDelete: () => void;
}) {
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[88vh] w-full max-w-2xl overflow-y-auto rounded-sm bg-white p-6 shadow-xl sm:p-7"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-5 flex items-start justify-between gap-4 border-b border-[#D8D6CC] pb-4">
          <div>
            <p className="text-xs tracking-wide text-[#8A6A3B]">X: {record.x}</p>
            <h2 className="font-serif text-xl">{record.hoTen || "(Chưa có họ tên)"}</h2>
          </div>
          <button onClick={onClose} className="icon-btn shrink-0" aria-label="Đóng">
            <XIcon size={16} />
          </button>
        </div>

        {FIELD_SECTIONS.map((section) => (
          <div key={section.title} className="mb-5 last:mb-0">
            <h3 className="mb-2 text-xs font-semibold uppercase tracking-wide text-[#8A6A3B]">
              {section.title}
            </h3>
            <div className="grid grid-cols-1 gap-x-6 gap-y-3 sm:grid-cols-2">
              {section.keys.map((k) => {
                const cfg = FIELD_MAP[k];
                const val = record[k];
                const isEmpty = !val?.trim();
                const note = k === "cccd" ? record.cccdNote : "";
                return (
                  <div key={k} className={cfg.full ? "sm:col-span-2" : ""}>
                    <p className="text-xs text-[#8A8A80]">{cfg.label}</p>
                    <p
                      className={`whitespace-pre-wrap text-sm ${
                        isEmpty ? "text-[#B3261E]" : "text-[#1B2A4A]"
                      }`}
                    >
                      {isEmpty ? "— (trống)" : val}
                    </p>
                    {note && <p className="mt-0.5 text-xs text-[#B3261E]">{note}</p>}
                  </div>
                );
              })}
            </div>
          </div>
        ))}

        <div className="mt-2 flex gap-3 border-t border-[#D8D6CC] pt-5">
          <button onClick={onEdit} className="btn-primary">
            <Pencil size={16} /> Sửa hồ sơ này
          </button>
          <button onClick={onDelete} className="btn-secondary border-[#B3261E] text-[#B3261E]">
            <Trash2 size={16} /> Xoá
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children, full }: { label: string; children: ReactNode; full?: boolean }) {
  return (
    <label className={`flex flex-col gap-1 ${full ? "sm:col-span-2 lg:col-span-4" : ""}`}>
      <span className="text-xs font-medium text-[#5B5B54]">{label}</span>
      {children}
    </label>
  );
}