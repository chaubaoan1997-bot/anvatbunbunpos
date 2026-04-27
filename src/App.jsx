import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import {
  addDoc,
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  onSnapshot,
  orderBy,
  query,
  serverTimestamp,
  setDoc,
  updateDoc,
  writeBatch,
} from "firebase/firestore";
import {
  Home,
  ShoppingBag,
  History,
  Wallet,
  BarChart3,
  Search,
  Package,
  Printer,
  Trash2,
  Plus,
  Pencil,
  X,
  Calendar,
  Receipt,
  CreditCard,
  BadgeDollarSign,
  CheckCircle2,
  Minus,
  Download,
  Boxes,
  ChevronLeft,
  ChevronRight,
} from "lucide-react";

const COLORS = {
  bg: "#f3f6fb",
  sidebar: "#071633",
  sidebarBorder: "rgba(255,255,255,0.06)",
  primary: "#2f66e9",
  primarySoft: "#e9f0ff",
  text: "#0f172a",
  textSoft: "#64748b",
  border: "#dfe6ef",
  white: "#ffffff",
  success: "#16a34a",
  successSoft: "#eaf8ee",
  danger: "#ef4444",
  dangerSoft: "#fdecec",
  warn: "#f59e0b",
  warnSoft: "#fff7e8",
  cardShadow: "0 8px 24px rgba(15,23,42,0.06)",
};

const defaultCategories = [
  "Tất cả",
  "Bánh tráng trộn",
  "Bánh tráng bịch",
  "Trà",
  "Bánh tráng cuộn",
  "Matcha",
  "Ăn vặt",
];

const defaultProducts = [
  { id: "seed-1", name: "Bánh tráng fulltopping", price: 40000, stock: 20, category: "Bánh tráng trộn", cost: 0, status: "Còn hàng" },
  { id: "seed-2", name: "Bánh tráng muối béo", price: 15000, stock: 20, category: "Bánh tráng bịch", cost: 0, status: "Còn hàng" },
  { id: "seed-3", name: "Bánh tráng muối gánh", price: 12000, stock: 20, category: "Bánh tráng bịch", cost: 0, status: "Còn hàng" },
  { id: "seed-4", name: "Bánh tráng muối sặc", price: 10000, stock: 20, category: "Bánh tráng bịch", cost: 0, status: "Còn hàng" },
  { id: "seed-5", name: "Bánh tráng muối sặc đặc biệt", price: 45000, stock: 15, category: "Bánh tráng bịch", cost: 0, status: "Còn hàng" },
  { id: "seed-6", name: "Bánh tráng vô tri", price: 10000, stock: 18, category: "Bánh tráng bịch", cost: 0, status: "Còn hàng" },
  { id: "seed-7", name: "Chanh dây nhiệt đới", price: 35000, stock: 10, category: "Trà", cost: 7240, status: "Còn hàng" },
  { id: "seed-8", name: "Cuộn bơ chấm sốt bò", price: 30000, stock: 16, category: "Bánh tráng cuộn", cost: 0, status: "Còn hàng" },
  { id: "seed-9", name: "Cuộn chấm sốt bơ me", price: 30000, stock: 16, category: "Bánh tráng cuộn", cost: 0, status: "Còn hàng" },
  { id: "seed-10", name: "Cuộn lắc muối bò", price: 40000, stock: 12, category: "Bánh tráng cuộn", cost: 0, status: "Còn hàng" },
  { id: "seed-11", name: "Cuộn trộn full topping", price: 40000, stock: 10, category: "Bánh tráng cuộn", cost: 0, status: "Còn hàng" },
  { id: "seed-12", name: "Hồng trà kem machiato", price: 35000, stock: 8, category: "Trà", cost: 0, status: "Còn hàng" },
  { id: "seed-13", name: "Hồng trà tắc mật ong", price: 18000, stock: 12, category: "Trà", cost: 0, status: "Còn hàng" },
  { id: "seed-14", name: "Matcha kem machiato", price: 30000, stock: 12, category: "Matcha", cost: 0, status: "Còn hàng" },
];

const pageTitle = { fontSize: 42, fontWeight: 800, lineHeight: 1.1 };
const pageSub = { fontSize: 16, color: COLORS.textSoft, marginTop: 8 };
const qtyBtn = {
  width: 30,
  height: 30,
  borderRadius: 8,
  border: `1px solid ${COLORS.border}`,
  background: COLORS.white,
  cursor: "pointer",
  fontWeight: 700,
  display: "grid",
  placeItems: "center",
};
const primaryBtn = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 10,
  border: "none",
  background: COLORS.primary,
  color: "#fff",
  padding: "12px 18px",
  borderRadius: 14,
  fontWeight: 700,
  cursor: "pointer",
};
const ghostBtn = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 10,
  border: `1px solid ${COLORS.border}`,
  background: COLORS.white,
  color: COLORS.text,
  padding: "12px 18px",
  borderRadius: 14,
  fontWeight: 700,
  cursor: "pointer",
};
const cellInput = {
  width: "100%",
  padding: "8px 10px",
  borderRadius: 10,
  border: `1px solid ${COLORS.border}`,
  outline: "none",
  background: "#fff",
  boxSizing: "border-box",
};
const iconBtn = {
  width: 36,
  height: 36,
  borderRadius: 10,
  border: `1px solid ${COLORS.border}`,
  background: COLORS.white,
  cursor: "pointer",
};

const money = (v) => `${Number(v || 0).toLocaleString("vi-VN")} đ`;

const dateInputValue = () => {
  const d = new Date();
  const y = d.getFullYear();
  const m = `${d.getMonth() + 1}`.padStart(2, "0");
  const day = `${d.getDate()}`.padStart(2, "0");
  return `${y}-${m}-${day}`;
};

function pageBtn(active) {
  return {
    minWidth: 36,
    height: 36,
    padding: "0 10px",
    borderRadius: 10,
    border: `1px solid ${COLORS.border}`,
    background: active ? COLORS.primary : COLORS.white,
    color: active ? "#fff" : COLORS.text,
    fontWeight: 700,
    cursor: "pointer",
  };
}

function pageBtnDisabled(disabled) {
  return {
    width: 36,
    height: 36,
    borderRadius: 10,
    border: `1px solid ${COLORS.border}`,
    background: disabled ? "#f8fafc" : COLORS.white,
    color: COLORS.text,
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.5 : 1,
    display: "grid",
    placeItems: "center",
  };
}

function exportExcel(filename, rows) {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
  XLSX.writeFile(workbook, filename);
}

function useWindowWidth() {
  const [width, setWidth] = useState(
    typeof window !== "undefined" ? window.innerWidth : 1280
  );

  useEffect(() => {
    const onResize = () => setWidth(window.innerWidth);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  return width;
}

function SidebarItem({ icon: Icon, label, active, onClick, mobile = false }) {
  return (
    <button
      onClick={onClick}
      style={{
        width: mobile ? "auto" : "100%",
        display: "flex",
        alignItems: "center",
        gap: 12,
        padding: mobile ? "10px 14px" : "15px 16px",
        borderRadius: 14,
        border: "none",
        background: active ? COLORS.primary : "transparent",
        color: active ? "#fff" : mobile ? COLORS.text : "#fff",
        cursor: "pointer",
        fontSize: mobile ? 14 : 18,
        fontWeight: 600,
        textAlign: "left",
        whiteSpace: "nowrap",
      }}
    >
      <Icon size={mobile ? 18 : 22} />
      <span>{label}</span>
    </button>
  );
}

function SectionCard({ children, style, ...props }) {
  return (
    <div
      {...props}
      style={{
        background: COLORS.white,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 22,
        boxShadow: COLORS.cardShadow,
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function Chip({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        borderRadius: 999,
        padding: "10px 18px",
        border: active
          ? `1px solid ${COLORS.primary}`
          : `1px solid ${COLORS.border}`,
        background: active ? COLORS.primary : COLORS.white,
        color: active ? "#fff" : COLORS.text,
        fontWeight: 600,
        fontSize: 14,
        cursor: "pointer",
        whiteSpace: "nowrap",
        flex: "0 0 auto",
      }}
    >
      {children}
    </button>
  );
}

function EmptyState({ icon: Icon, title, subtitle }) {
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        height: "100%",
        color: COLORS.textSoft,
        gap: 14,
      }}
    >
      <div
        style={{
          width: 88,
          height: 88,
          borderRadius: 24,
          background: "#f8fafc",
          display: "grid",
          placeItems: "center",
          border: `1px solid ${COLORS.border}`,
        }}
      >
        <Icon size={42} color="#cbd5e1" />
      </div>
      <div style={{ fontSize: 16 }}>{title}</div>
      {subtitle ? <div style={{ fontSize: 14 }}>{subtitle}</div> : null}
    </div>
  );
}

function ProductModal({ open, onClose, onSave, categories, product }) {
  const [form, setForm] = useState({
    name: "",
    price: "",
    stock: "",
    category: categories[1] || "",
    cost: "0",
  });

  useEffect(() => {
    if (product) {
      setForm({
        name: product.name || "",
        price: product.price || "",
        stock: product.stock || "",
        category: product.category || categories[1] || "",
        cost: product.cost || 0,
      });
    } else {
      setForm({
        name: "",
        price: "",
        stock: "",
        category: categories[1] || "",
        cost: "0",
      });
    }
  }, [product, categories, open]);

  if (!open) return null;

  const fieldStyle = {
    width: "100%",
    padding: "12px 14px",
    borderRadius: 12,
    border: `1px solid ${COLORS.border}`,
    outline: "none",
    fontSize: 15,
    boxSizing: "border-box",
  };

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(2,6,23,0.4)",
        display: "grid",
        placeItems: "center",
        zIndex: 999,
      }}
    >
      <div
        style={{
          width: "min(520px, calc(100vw - 24px))",
          background: COLORS.white,
          borderRadius: 22,
          boxShadow: "0 24px 60px rgba(2,6,23,0.25)",
          border: `1px solid ${COLORS.border}`,
        }}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            padding: 22,
            borderBottom: `1px solid ${COLORS.border}`,
          }}
        >
          <div style={{ fontSize: 24, fontWeight: 700 }}>
            {product ? "Sửa sản phẩm" : "Thêm sản phẩm"}
          </div>
          <button
            onClick={onClose}
            style={{ border: "none", background: "transparent", cursor: "pointer" }}
          >
            <X />
          </button>
        </div>
        <div style={{ padding: 22, display: "grid", gap: 14 }}>
          <input
            style={fieldStyle}
            placeholder="Tên sản phẩm"
            value={form.name}
            onChange={(e) => setForm({ ...form, name: e.target.value })}
          />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <input
              style={fieldStyle}
              placeholder="Giá bán"
              type="number"
              value={form.price}
              onChange={(e) => setForm({ ...form, price: e.target.value })}
            />
            <input
              style={fieldStyle}
              placeholder="Tồn kho"
              type="number"
              value={form.stock}
              onChange={(e) => setForm({ ...form, stock: e.target.value })}
            />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <select
              style={fieldStyle}
              value={form.category}
              onChange={(e) => setForm({ ...form, category: e.target.value })}
            >
              {categories
                .filter((c) => c !== "Tất cả")
                .map((c) => (
                  <option key={c}>{c}</option>
                ))}
            </select>
            <input
              style={fieldStyle}
              placeholder="Giá vốn"
              type="number"
              value={form.cost}
              onChange={(e) => setForm({ ...form, cost: e.target.value })}
            />
          </div>
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            gap: 10,
            padding: 22,
            borderTop: `1px solid ${COLORS.border}`,
          }}
        >
          <button
            onClick={onClose}
            style={{
              padding: "12px 18px",
              borderRadius: 12,
              border: `1px solid ${COLORS.border}`,
              background: COLORS.white,
              cursor: "pointer",
              fontWeight: 600,
            }}
          >
            Hủy
          </button>
          <button
            onClick={() =>
              onSave({
                ...form,
                price: Number(form.price || 0),
                stock: Number(form.stock || 0),
                cost: Number(form.cost || 0),
              })
            }
            style={{
              padding: "12px 18px",
              borderRadius: 12,
              border: "none",
              background: COLORS.primary,
              color: "#fff",
              cursor: "pointer",
              fontWeight: 700,
            }}
          >
            Lưu sản phẩm
          </button>
        </div>
      </div>
    </div>
  );
}

function PaymentButton({ active, onClick, icon: Icon, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 10,
        padding: "15px 18px",
        borderRadius: 14,
        fontWeight: 700,
        cursor: "pointer",
        border: active
          ? `2px solid ${COLORS.primary}`
          : `1px solid ${COLORS.border}`,
        background: active ? COLORS.primary : COLORS.white,
        color: active ? "#fff" : COLORS.text,
      }}
    >
      {Icon ? <Icon size={18} /> : null}
      {children}
    </button>
  );
}

function DiscountButton({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        minWidth: 38,
        height: 30,
        borderRadius: 10,
        border: active
          ? `1px solid ${COLORS.primary}`
          : `1px solid ${COLORS.border}`,
        background: active ? COLORS.primarySoft : COLORS.white,
        color: active ? COLORS.primary : COLORS.textSoft,
        fontWeight: 700,
        cursor: "pointer",
      }}
    >
      {children}
    </button>
  );
}

function Pagination({ page, totalPage, onPageChange }) {
  if (totalPage <= 1) return null;

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 8,
        paddingTop: 14,
        flexWrap: "wrap",
      }}
    >
      <button
        onClick={() => onPageChange(Math.max(1, page - 1))}
        disabled={page === 1}
        style={pageBtnDisabled(page === 1)}
      >
        <ChevronLeft size={16} />
      </button>
      {Array.from({ length: totalPage }).map((_, i) => {
        const value = i + 1;
        return (
          <button
            key={value}
            onClick={() => onPageChange(value)}
            style={pageBtn(page === value)}
          >
            {value}
          </button>
        );
      })}
      <button
        onClick={() => onPageChange(Math.min(totalPage, page + 1))}
        disabled={page === totalPage}
        style={pageBtnDisabled(page === totalPage)}
      >
        <ChevronRight size={16} />
      </button>
    </div>
  );
}

function MetricCard({ title, value, sub, color, soft, icon }) {
  return (
    <SectionCard style={{ padding: 28, background: soft }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ color: COLORS.textSoft, fontSize: 16 }}>{title}</div>
        {icon || <div style={{ width: 18, height: 18 }} />}
      </div>
      <div style={{ marginTop: 18, fontSize: 42, fontWeight: 800, color }}>{value}</div>
      {sub ? <div style={{ marginTop: 8, color: COLORS.textSoft }}>{sub}</div> : null}
    </SectionCard>
  );
}

export default function App() {
  const width = useWindowWidth();
  const isMobile = width < 768;

  const [page, setPage] = useState("sales");
  const [productPage, setProductPage] = useState(1);
  const pageSize = isMobile ? 6 : 8;

  const [products, setProducts] = useState([]);
  const [orders, setOrders] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [categories] = useState(defaultCategories);
  const [cart, setCart] = useState([]);
  const [search, setSearch] = useState("");
  const [productSearch, setProductSearch] = useState("");
  const [historySearch, setHistorySearch] = useState("");
  const [activeCategory, setActiveCategory] = useState("Tất cả");
  const [selectedOrder, setSelectedOrder] = useState(null);
  const [showProductModal, setShowProductModal] = useState(false);
  const [editingProduct, setEditingProduct] = useState(null);
  const [reportDate, setReportDate] = useState(dateInputValue());
  const [reportType, setReportType] = useState("day"); // day | week | month
  const [expenseDate, setExpenseDate] = useState(dateInputValue());
  const [historyDate, setHistoryDate] = useState(dateInputValue());
  const [historyType, setHistoryType] = useState("day"); // day | week | month
  const [paymentMethod, setPaymentMethod] = useState("Tiền mặt");
  const [paidMessage, setPaidMessage] = useState("");
  // ===== CHẤM CÔNG =====
  const [employees, setEmployees] = useState([]);
  const [attendance, setAttendance] = useState([]);

  const [employeeName, setEmployeeName] = useState("");
  const [selectedEmp, setSelectedEmp] = useState(null);

  const [attendanceDate, setAttendanceDate] = useState(dateInputValue());
  const addEmployee = async () => {
    if (!employeeName) return;

    await addDoc(collection(db, "employees"), {
      name: employeeName,
      salary: 20000,
      createdAt: serverTimestamp(),
    });

    setEmployeeName("");
  };
  const handleCheckIn = async () => {
    if (!selectedEmp) return alert("Bạn chưa chọn nhân viên");

    const now = new Date();
    const time = now.toTimeString().slice(0, 5);

    await addDoc(collection(db, "attendance"), {
      empId: selectedEmp.id,
      start: time,
      end: "",
      hours: 0,
      dateKey: attendanceDate,
      createdAt: serverTimestamp(),
    });
  };

  const handleCheckOut = async (record) => {
    if (!record) return;

    const now = new Date();
    const end = now.toTimeString().slice(0, 5);

    const startTime = new Date(`1970-01-01T${record.start}`);
    const endTime = new Date(`1970-01-01T${end}`);

    let hours = (endTime - startTime) / 3600000;
    if (hours < 0) hours += 24;

    await updateDoc(doc(db, "attendance", record.id), {
      end,
      hours,
    });
  };

  const [wholesalePayment, setWholesalePayment] = useState("Tiền mặt");
  const [tempOrders, setTempOrders] = useState([]);
  const [wholesaleCart, setWholesaleCart] = useState([]);
  const [sellerInfo, setSellerInfo] = useState({ name: "", });
  const [customerInfo, setCustomerInfo] = useState({
    name: "",
    phone: ""
  });
  const [editingWholesaleOrder, setEditingWholesaleOrder] = useState(null);

  const [reportProductSearch, setReportProductSearch] = useState("");
  const [reportCategory, setReportCategory] = useState("Tất cả");

  const [showPaymentModal, setShowPaymentModal] = useState(false);
  const [tempOrder, setTempOrder] = useState(null);


  const [filterType, setFilterType] = useState("all");

  const [showDeliveryModal, setShowDeliveryModal] = useState(false);
  const [deliveryForm, setDeliveryForm] = useState({
    name: "",
    phone: "",
    address: "",
    payment: "Tiền mặt",
  });

  const isSameWeek = (dateStr, selectedDate) => {
    const d1 = new Date(dateStr + "T00:00:00");
    const d2 = new Date(selectedDate + "T00:00:00");

    // đưa về thứ 2 (chuẩn)
    const day = d2.getDay() || 7;

    const first = new Date(d2);
    first.setDate(d2.getDate() - day + 1);

    const last = new Date(first);
    last.setDate(first.getDate() + 6);

    return d1 >= first && d1 <= last;
  };

  const isSameMonth = (dateStr, selectedDate) => {
    const d1 = new Date(dateStr);
    const d2 = new Date(selectedDate);
    return (
      d1.getMonth() === d2.getMonth() &&
      d1.getFullYear() === d2.getFullYear()
    );
  };

  useEffect(() => {
    const ensureSeed = async () => {
      try {
        const snap = await getDocs(collection(db, "products"));

        if (!snap.empty) return; // nếu đã có data thì không seed nữa

        for (const item of defaultProducts) {
          await setDoc(doc(db, "products", item.id), item);
        }
      } catch (err) {
        console.error("Seed products error:", err);
      }
    };

    ensureSeed();
  }, []);

  useEffect(() => {
    const unsubProducts = onSnapshot(
      query(collection(db, "products")),
      (snap) => {
        setProducts(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      },
      (err) => console.error("Products snapshot error:", err)
    );

    const unsubOrders = onSnapshot(
      query(collection(db, "orders"), orderBy("createdAt", "desc")),
      (snap) => {
        const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        setOrders(data);
        setSelectedOrder((prev) => data.find((x) => x.id === prev?.id) || data[0] || null);
      },
      (err) => console.error("Orders snapshot error:", err)
    );

    const unsubExpenses = onSnapshot(
      query(collection(db, "expenses"), orderBy("createdAt", "desc")),
      (snap) => {
        setExpenses(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      },
      (err) => console.error("Expenses snapshot error:", err)
    );

    // ===== CHẤM CÔNG =====
    const unsubEmployees = onSnapshot(
      collection(db, "employees"),
      (snap) => {
        setEmployees(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      },
      (err) => console.error("Employees snapshot error:", err)
    );

    const unsubAttendance = onSnapshot(
      collection(db, "attendance"),
      (snap) => {
        setAttendance(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
      },
      (err) => console.error("Attendance snapshot error:", err)
    );

    return () => {
      unsubProducts();
      unsubOrders();
      unsubExpenses();
      unsubEmployees();
      unsubAttendance();
    };
  }, []);

  useEffect(() => {
    setProductPage(1);
  }, [search, activeCategory]);

  useEffect(() => {
    if (!paidMessage) return;
    const t = setTimeout(() => setPaidMessage(""), 2200);
    return () => clearTimeout(t);
  }, [paidMessage]);

  const filteredProducts = useMemo(() => {
    return products.filter(
      (p) =>
        (p.name || "").toLowerCase().includes(search.toLowerCase()) &&
        (activeCategory === "Tất cả" ? true : p.category === activeCategory)
    );
  }, [products, search, activeCategory]);

  const totalProductPage = Math.max(1, Math.ceil(filteredProducts.length / pageSize));
  const paginatedProducts = filteredProducts.slice(
    (productPage - 1) * pageSize,
    productPage * pageSize
  );

  const filteredProductTable = useMemo(() => {
    return products.filter((p) =>
      `${p.name || ""} ${p.category || ""}`.toLowerCase().includes(productSearch.toLowerCase())
    );
  }, [products, productSearch]);

  const historyFiltered = useMemo(() => {
    return orders.filter((o) => {
      const codeOk = (o.code || "")
        .toLowerCase()
        .includes(historySearch.toLowerCase());

      let dateOk = true;

      if (historyType === "day") {
        dateOk = o.dateKey === historyDate;
      } else if (historyType === "week") {
        dateOk = isSameWeek(o.dateKey, historyDate);
      } else if (historyType === "month") {
        dateOk = isSameMonth(o.dateKey, historyDate);
      }

      return codeOk && dateOk;
    });
  }, [orders, historySearch, historyDate, historyType]);

  const expenseFiltered = useMemo(
    () => expenses.filter((e) => e.dateKey === expenseDate),
    [expenses, expenseDate]
  );

  const reportOrders = useMemo(() => {
    return orders.filter((o) => {
      if (reportType === "day") return o.dateKey === reportDate;
      if (reportType === "week") return isSameWeek(o.dateKey, reportDate);
      if (reportType === "month") return isSameMonth(o.dateKey, reportDate);
      return true;
    });
  }, [orders, reportDate, reportType]);

  const wholesaleStats = useMemo(() => {
    const map = {};

    reportOrders
      .filter(o => o.type === "wholesale" && o.status === "Đã thanh toán")
      .forEach(order => {
        const seller = order.sellerName || "Không tên";

        if (!map[seller]) {
          map[seller] = {
            seller,
            orders: 0,
            revenue: 0,
            discount: 0,
          };
        }

        map[seller].orders += 1;
        map[seller].revenue += Number(order.total || 0);
        map[seller].discount += Number(order.discount || 0);
      });

    return Object.values(map);
  }, [reportOrders]);

  const reportProductStats = useMemo(() => {
    const map = {};

    reportOrders
      .filter(o => o.status === "Đã thanh toán")
      .forEach(order => {
        (order.items || []).forEach(item => {
          if (!map[item.name]) {
            map[item.name] = {
              name: item.name,
              qty: 0,
              revenue: 0,
            };
          }

          const raw = item.price * item.qty;
          const final = raw - (raw * (item.discount || 0)) / 100;

          map[item.name].qty += item.qty;
          map[item.name].revenue += final;
        });
      });

    return Object.values(map);
  }, [reportOrders]);

  const filteredReportProducts = useMemo(() => {
    return reportProductStats.filter(p => {
      const matchSearch = p.name
        .toLowerCase()
        .includes(reportProductSearch.toLowerCase());

      const product = products.find(x => x.name === p.name);

      const matchCategory =
        reportCategory === "Tất cả" ||
        (product && product.category === reportCategory);

      return matchSearch && matchCategory;
    });
  }, [reportProductStats, reportProductSearch, reportCategory, products]);

  const reportExpenses = useMemo(() => {
    return expenses.filter((e) => {
      if (reportType === "day") return e.dateKey === reportDate;
      if (reportType === "week") return isSameWeek(e.dateKey, reportDate);
      if (reportType === "month") return isSameMonth(e.dateKey, reportDate);
      return true;
    });
  }, [expenses, reportDate, reportType]);

  const lowStockProducts = useMemo(
    () => products.filter((p) => Number(p.stock || 0) <= 5),
    [products]
  );

  const getCartItemTotal = (item) => {
    const base = Number(item.price || 0) * Number(item.qty || 0);
    const discountValue = (base * Number(item.discount || 0)) / 100;
    return base - discountValue;
  };

  const total = cart.reduce((sum, item) => sum + getCartItemTotal(item), 0);
  const orderCount = reportOrders.length;
  const revenue = reportOrders
    .filter(o => o.status === "Đã thanh toán")
    .reduce((sum, o) => sum + Number(o.total || 0), 0);
  const cost = reportExpenses.reduce((sum, e) => sum + Number(e.amount || 0), 0);
  const cashRevenue = reportOrders
    .filter(o => o.method === "Tiền mặt" && o.status === "Đã thanh toán")
    .reduce((s, o) => s + Number(o.total || 0), 0);
  const bankRevenue = reportOrders
    .filter(o => o.method === "Chuyển khoản" && o.status === "Đã thanh toán")
    .reduce((s, o) => s + Number(o.total || 0), 0);
  const profit = revenue - cost;
  const totalExpenseToday = expenseFiltered.reduce((s, e) => s + Number(e.amount || 0), 0);

  const addToCart = (product) => {
    if (Number(product.stock || 0) <= 0) return;

    const found = cart.find((i) => i.id === product.id);
    if (found) {
      if (found.qty >= Number(product.stock || 0)) return;
      setCart(cart.map((i) => (i.id === product.id ? { ...i, qty: i.qty + 1 } : i)));
    } else {
      setCart([...cart, { ...product, qty: 1, discount: 0 }]);
    }
  };

  const addToWholesaleCart = (product) => {
    const found = wholesaleCart.find(i => i.id === product.id);

    if (found) {
      setWholesaleCart(
        wholesaleCart.map(i =>
          i.id === product.id ? { ...i, qty: i.qty + 1 } : i
        )
      );
    } else {
      setWholesaleCart([
        ...wholesaleCart,
        {
          ...product,
          qty: 1,
          customPrice: product.price,
          discountCash: 0
        }
      ]);
    }
  };

  const getWholesaleTotal = () => {
    return wholesaleCart.reduce((sum, i) => {
      const price = Number(i.customPrice || 0);
      const qty = Number(i.qty || 0);

      return sum + price * qty;
    }, 0);
  };

  const getWholesaleDiscount = () => {
    return wholesaleCart.reduce((sum, i) => {
      return sum + (Number(i.discountCash || 0) * i.qty);
    }, 0);
  };

  const loadTempOrderToCart = (order) => {
    if (!order) return;

    if (order.type === "wholesale") {
      setWholesaleCart(order.items || []);
      setSellerInfo({
        name: order.sellerName || "",
        phone: order.sellerPhone || ""
      });
      setEditingWholesaleOrder(order); // thêm dòng này
      setPage("wholesale");
      return;
    }

    setCart(
      (order.items || []).map(i => ({
        ...i,
        discount: i.discount || 0
      }))
    );
    setEditingOrder(order);
    setPage("sales");
  };

  const updateCartQty = (id, next) => {
    const product = products.find((p) => p.id === id);
    if (next <= 0) {
      setCart(cart.filter((i) => i.id !== id));
      return;
    }
    if (product && next > Number(product.stock || 0)) return;
    setCart(cart.map((i) => (i.id === id ? { ...i, qty: next } : i)));
  };

  const updateDiscount = (id, discount) => {
    let value = Number(discount || 0);
    if (value < 0) value = 0;
    if (value > 100) value = 100;
    setCart(cart.map((i) => (i.id === id ? { ...i, discount: value } : i)));
  };

  const printReceipt = (orderData) => {
    const data = orderData || {
      items: cart,
      total,
      code: "Đơn tạm",
      method: paymentMethod,
      timeText: new Date().toLocaleString("vi-VN"),
      status: "Đã thanh toán",
    };

    const html = `
      <html>
      <head>
        <title>In hóa đơn</title>
        <style>
          body{font-family:Arial,sans-serif;width:280px;padding:10px;color:#111}
          .center{text-align:center}
          .line{border-top:1px dashed #000;margin:8px 0}
          .row{display:flex;justify-content:space-between;font-size:12px;margin:3px 0}
          .b{font-weight:700}
        </style>
      </head>
      <body>
        <div class="center b">AN VAT BUNBUN</div>
        <div class="center">${data.code || "Hóa đơn"}</div>
        <div class="line"></div>
        ${(data.items || [])
        .map((i) => {
          const raw = Number(i.price || 0) * Number(i.qty || 0);
          const after = raw - (raw * Number(i.discount || 0)) / 100;
          return `<div class="row"><span>${i.name} x${i.qty}${i.discount ? ` (-${i.discount}%)` : ""}</span><span>${Number(after).toLocaleString("vi-VN")} đ</span></div>`;
        })
        .join("")}
        <div class="line"></div>
        <div class="row b"><span>Tổng cộng</span><span>${Number(data.total || 0).toLocaleString("vi-VN")} đ</span></div>
        <div class="row"><span>Thanh toán</span><span>${data.method || "Tiền mặt"}</span></div>
        <div class="row"><span>Trạng thái</span><span>${data.status || "Đã thanh toán"}</span></div>
        <div class="row"><span>Thời gian</span><span>${data.timeText || new Date().toLocaleString("vi-VN")}</span></div>
        <div class="center b">Cám ơn quý khách ❤️</div>
        <script>window.onload=()=>window.print()</script>
      </body>
      </html>`;

    const w = window.open("", "_blank", "width=320,height=700");
    if (w) {
      w.document.write(html);
      w.document.close();
    }
  };

  const checkout = async () => {
    if (!cart.length) return;
    setPaidMessage("Đang xử lý...");

    try {
      const now = new Date();
      const code = `DH${Date.now().toString().slice(-6)}`;

      const orderPayload = {
        code,
        items: cart,
        total,
        method: paymentMethod,
        status: "Đã thanh toán",
        dateKey: dateInputValue(),
        createdAt: serverTimestamp(),
        timeText: now.toLocaleString("vi-VN"),
      };

      await addDoc(collection(db, "orders"), orderPayload);

      const batch = writeBatch(db);
      for (const item of cart) {
        const ref = doc(db, "products", item.id);
        const snap = await getDoc(ref);
        if (!snap.exists()) continue;
        const currentStock = Number(snap.data().stock || 0);
        const nextStock = Math.max(0, currentStock - Number(item.qty || 0));
        batch.update(ref, {
          stock: nextStock,
          status: nextStock > 0 ? "Còn hàng" : "Hết hàng",
        });
      }
      await batch.commit();

      setCart([]);
      setPaidMessage(`Đã thanh toán • ${paymentMethod}`);
    } catch (err) {
      console.error(err);
      setPaidMessage("Lỗi thanh toán!");
    }
  };

  const createTempOrder = async () => {
    try {
      if (!cart.length) return;

      const now = new Date();
      const code = `TMP${Date.now().toString().slice(-6)}`;

      const orderPayload = {
        code,
        items: cart,
        total,
        method: "Chưa thanh toán",
        status: "Đơn tạm",
        isTemp: true,
        dateKey: dateInputValue(),
        createdAt: serverTimestamp(),
        timeText: now.toLocaleString("vi-VN"),
      };

      await addDoc(collection(db, "orders"), orderPayload);

      setCart([]);
      setPaidMessage("Đã lưu đơn tạm");
    } catch (err) {
      console.error("Lỗi tạo đơn tạm:", err);
      alert("Không lưu được đơn tạm");
    }
  };

  const createWholesaleOrder = async () => {
    if (!wholesaleCart.length) return;

    const now = new Date();
    const code = `SỈ${Date.now().toString().slice(-6)}`;

    const orderPayload = {
      code,
      items: wholesaleCart,
      total: getWholesaleTotal(),
      discount: getWholesaleDiscount(),

      method: wholesalePayment, // 🔥 lấy từ nút chọn
      status: "Đã thanh toán",

      type: "wholesale", // 🔥 QUAN TRỌNG

      sellerName: sellerInfo.name,
      customerName: customerInfo.name,
      customerPhone: customerInfo.phone,

      dateKey: dateInputValue(),
      createdAt: serverTimestamp(),
      timeText: now.toLocaleString("vi-VN"),
    };

    await addDoc(collection(db, "orders"), orderPayload);

    setWholesaleCart([]);
    setSellerInfo({ name: "", phone: "" });

    setPaidMessage("Đã thanh toán đơn sỉ");
  };

  const saveWholesaleTemp = async () => {
    if (!wholesaleCart.length) return;

    const now = new Date();
    const code = `SỈ_TMP${Date.now().toString().slice(-6)}`;

    const orderPayload = {
      code,
      items: wholesaleCart,
      total: getWholesaleTotal(),

      method: wholesalePayment,
      status: "Đơn tạm",
      isTemp: true,
      type: "wholesale",   // 🔥 phân biệt đơn sỉ

      sellerName: sellerInfo.name,
      sellerPhone: sellerInfo.phone,

      dateKey: dateInputValue(),
      createdAt: serverTimestamp(),
      timeText: now.toLocaleString("vi-VN"),
    };

    await addDoc(collection(db, "orders"), orderPayload);

    setWholesaleCart([]);
    setSellerInfo({ name: "", phone: "" });

    setPaidMessage("Đã lưu đơn sỉ tạm");
  };

  const updateWholesaleTempOrder = async () => {
    if (!editingWholesaleOrder) return;

    await updateDoc(doc(db, "orders", editingWholesaleOrder.id), {
      items: wholesaleCart,
      total: getWholesaleTotal(),
      sellerName: sellerInfo.name,
      sellerPhone: sellerInfo.phone,
      updatedAt: serverTimestamp(),
    });

    setWholesaleCart([]);
    setSellerInfo({ name: "", phone: "" });
    setEditingWholesaleOrder(null);
    setSelectedOrder(null);
    setPage("history");
    setPaidMessage("Đã cập nhật đơn tạm sỉ");
  };

  const updateTempOrder = async () => {
    if (!editingOrder) return;

    await updateDoc(doc(db, "orders", editingOrder.id), {
      items: cart,
      total,
    });

    setCart([]);
    setEditingOrder(null);
    setSelectedOrder(null);
    setPage("history");

    setPaidMessage("Đã cập nhật đơn tạm");
  };

  const confirmTempOrder = async (order, method) => {
    try {
      if (!order) return;

      await updateDoc(doc(db, "orders", order.id), {
        status: "Đã thanh toán",
        method: method || "Tiền mặt", // 🔥 FIX
        isTemp: false,
      });

      const batch = writeBatch(db);

      for (const item of order.items || []) {
        const ref = doc(db, "products", item.id);
        const snap = await getDoc(ref);
        if (!snap.exists()) continue;

        const currentStock = Number(snap.data().stock || 0);
        const nextStock = Math.max(0, currentStock - Number(item.qty || 0));

        batch.update(ref, {
          stock: nextStock,
          status: nextStock > 0 ? "Còn hàng" : "Hết hàng",
        });
      }

      await batch.commit();

      setPaidMessage(`Đã thanh toán đơn tạm • ${method || "Tiền mặt"}`);
    } catch (err) {
      console.error(err);
    }
  };

  const createDeliveryOrder = async () => {
    if (!cart.length) return;

    const now = new Date();
    const code = `GH${Date.now().toString().slice(-6)}`;

    const orderPayload = {
      code,
      items: cart,
      total,
      method: deliveryForm.payment,
      status: "Chờ giao",
      isDelivery: true,

      customer: {
        name: deliveryForm.name || "",
        phone: deliveryForm.phone || "",
        address: deliveryForm.address || "",
      },

      dateKey: dateInputValue(),
      createdAt: serverTimestamp(),
      timeText: now.toLocaleString("vi-VN"),

      isPaid: deliveryForm.payment !== "Nợ",
    };

    await addDoc(collection(db, "orders"), orderPayload);

    setCart([]);
    setDeliveryForm({
      name: "",
      phone: "",
      address: "",
      payment: "Tiền mặt",
    });
    setShowDeliveryModal(false);
  };

  const updateDeliveryOrder = async (orderId) => {
    if (!orderId) return;

    await updateDoc(doc(db, "orders", orderId), {
      method: deliveryForm.payment,

      customer: {
        name: deliveryForm.name,
        phone: deliveryForm.phone,
        address: deliveryForm.address,
      }
    });

    setShowDeliveryModal(false);
  };

  const [editingOrder, setEditingOrder] = useState(null);

  const completeDelivery = async (order) => {
    try {
      if (!order) return;

      await updateDoc(doc(db, "orders", order.id), {
        status: "Đã thanh toán",
        isDelivery: false,
      });

      // trừ kho
      const batch = writeBatch(db);

      for (const item of order.items || []) {
        const ref = doc(db, "products", item.id);
        const snap = await getDoc(ref);
        if (!snap.exists()) continue;

        const currentStock = Number(snap.data().stock || 0);
        const nextStock = Math.max(0, currentStock - Number(item.qty || 0));

        batch.update(ref, {
          stock: nextStock,
          status: nextStock > 0 ? "Còn hàng" : "Hết hàng",
        });
      }

      await batch.commit();

      setPaidMessage("Đã giao hàng & thanh toán");
    } catch (err) {
      console.error(err);
    }
  };

  const saveProduct = async (payload) => {
    if (!payload.name) return;
    const status = Number(payload.stock) > 0 ? "Còn hàng" : "Hết hàng";
    if (editingProduct) {
      await updateDoc(doc(db, "products", editingProduct.id), { ...payload, status });
    } else {
      await addDoc(collection(db, "products"), { ...payload, status });
    }
    setShowProductModal(false);
    setEditingProduct(null);
  };

  const quickEditProduct = async (id, field, value) => {
    const payload = { [field]: value };
    if (field === "stock") payload.status = Number(value) > 0 ? "Còn hàng" : "Hết hàng";
    await updateDoc(doc(db, "products", id), payload);
  };

  const deleteOrder = async (id) => {
    const target = orders.find((o) => o.id === id);

    if (target) {
      const batch = writeBatch(db);
      for (const item of target.items || []) {
        const ref = doc(db, "products", item.id);
        const snap = await getDoc(ref);
        if (!snap.exists()) continue;
        const currentStock = Number(snap.data().stock || 0);
        batch.update(ref, {
          stock: currentStock + Number(item.qty || 0),
          status: "Còn hàng",
        });
      }
      await batch.commit();
    }

    await deleteDoc(doc(db, "orders", id));
    if (selectedOrder?.id === id) setSelectedOrder(null);
  };

  const addExpense = async () => {
    const name = prompt("Tên khoản chi:");
    if (!name) return;
    const amount = Number(prompt("Số tiền:", "0") || 0);
    await addDoc(collection(db, "expenses"), {
      name,
      amount,
      dateKey: expenseDate,
      createdAt: serverTimestamp(),
      timeText: new Date().toLocaleString("vi-VN"),
    });
  };

  const updateExpense = async (id, oldExpense) => {
    const name = prompt("Sửa tên khoản chi:", oldExpense.name);
    if (!name) return;
    const amount = Number(prompt("Sửa số tiền:", oldExpense.amount) || 0);
    await updateDoc(doc(db, "expenses", id), {
      name,
      amount,
    });
  };

  const deleteExpenseItem = async (id) => {
    if (!confirm("Xóa khoản chi này?")) return;
    await deleteDoc(doc(db, "expenses", id));
  };

  const exportOrdersExcel = () =>
    exportExcel(
      `lich-su-don-hang-${historyDate}.xlsx`,
      historyFiltered.map((o) => ({
        MaDon: o.code,
        ThoiGian: o.timeText,
        ThanhToan: o.method,
        TrangThai: o.status,
        TongTien: o.total,
      }))
    );

  const exportProductsExcel = () =>
    exportExcel(
      "san-pham.xlsx",
      filteredProductTable.map((p) => ({
        TenSanPham: p.name,
        DanhMuc: p.category,
        GiaVon: p.cost || 0,
        GiaBan: p.price,
        TonKho: p.stock,
        TrangThai: p.status,
      }))
    );

  const exportReportExcel = () =>
    exportExcel(
      `bao-cao-${reportDate}.xlsx`,
      reportOrders.map((o) => ({
        MaDon: o.code,
        ThoiGian: o.timeText,
        ThanhToan: o.method,
        TongTien: o.total,
        TrangThai: o.status,
      }))
    );

  const pages = [
    { key: "sales", label: "Bán hàng", icon: Home },
    { key: "wholesale", label: "Bán sỉ", icon: ShoppingBag },
    { key: "products", label: "Sản phẩm", icon: ShoppingBag },
    { key: "history", label: "Lịch sử", icon: History },
    { key: "expense", label: "Chi phí", icon: Wallet },
    { key: "report", label: "Báo cáo", icon: BarChart3 },
    { key: "attendance", label: "Chấm công", icon: Calendar },
  ];

  const salesPage = (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: isMobile ? "1fr" : "minmax(0, 1fr) 400px",
        gap: 18,
        height: "100%",
        minHeight: 0,
      }}
    >
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 16,
          minWidth: 0,
          minHeight: 0,
          position: "relative",
          zIndex: 1,
        }}
      >
        {!!lowStockProducts.length && (
          <SectionCard
            style={{
              padding: 14,
              background: COLORS.warnSoft,
              display: "flex",
              alignItems: "center",
              gap: 10,
              flexShrink: 0,
            }}
          >
            <Boxes size={18} color={COLORS.warn} />
            <strong>Cảnh báo tồn kho thấp:</strong>
            <span>
              {lowStockProducts
                .slice(0, 4)
                .map((p) => `${p.name} (${p.stock})`)
                .join(", ")}
            </span>
          </SectionCard>
        )}

        <div
          style={{
            position: "relative",
            width: "100%",
            maxWidth: 520,
            flexShrink: 0,
            zIndex: 0,
          }}
        >
          <Search size={20} color="#64748b" style={{ position: "absolute", left: 18, top: 17 }} />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Tìm kiếm sản phẩm..."
            style={{
              width: "100%",
              height: 56,
              paddingLeft: 52,
              borderRadius: 16,
              border: `1px solid ${COLORS.border}`,
              outline: "none",
              background: COLORS.white,
              fontSize: 16,
              boxShadow: COLORS.cardShadow,
              boxSizing: "border-box",
            }}
          />
        </div>

        <div style={{ display: "flex", gap: 10, overflowX: "auto", paddingBottom: 4, flexShrink: 0 }}>
          {categories.map((c) => (
            <Chip key={c} active={activeCategory === c} onClick={() => setActiveCategory(c)}>
              {c}
            </Chip>
          ))}
        </div>

        <div style={{ overflowY: "auto", paddingRight: 4, position: "relative", zIndex: 2 }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: isMobile
                ? "repeat(2, minmax(0, 1fr))"
                : "repeat(4, minmax(0, 1fr))",
              gap: 18,
            }}
          >
            {paginatedProducts.map((p) => (
              <SectionCard
                key={p.id}
                onClick={() => Number(p.stock || 0) > 0 && addToCart(p)}
                style={{
                  padding: 18,
                  minHeight: 170,
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "space-between",
                  cursor: Number(p.stock || 0) > 0 ? "pointer" : "not-allowed",
                  borderColor: "#d9e2ee",
                  overflow: "hidden",
                  opacity: Number(p.stock || 0) > 0 ? 1 : 0.55,
                }}
              >
                <div>
                  <div
                    style={{
                      fontSize: 17,
                      lineHeight: 1.3,
                      fontWeight: 600,
                      color: COLORS.text,
                      minHeight: 48,
                      display: "-webkit-box",
                      WebkitLineClamp: 2,
                      WebkitBoxOrient: "vertical",
                      overflow: "hidden",
                    }}
                  >
                    {p.name}
                  </div>
                  <div
                    style={{
                      marginTop: 10,
                      color: Number(p.stock || 0) > 0 ? COLORS.success : COLORS.danger,
                      fontWeight: 700,
                      fontSize: 13,
                    }}
                  >
                    {Number(p.stock || 0) > 0 ? `Còn hàng • Tồn ${p.stock}` : "Hết hàng"}
                  </div>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", gap: 10 }}>
                  <div style={{ fontSize: 18, color: COLORS.primary, fontWeight: 700 }}>
                    {money(p.price)}
                  </div>
                  <button
                    disabled={Number(p.stock || 0) <= 0}
                    onClick={(e) => {
                      e.stopPropagation();
                      addToCart(p);
                    }}
                    style={{
                      border: "none",
                      background: COLORS.primary,
                      color: "#fff",
                      borderRadius: 12,
                      padding: "10px 12px",
                      cursor: Number(p.stock || 0) > 0 ? "pointer" : "not-allowed",
                      fontWeight: 700,
                      flexShrink: 0,
                      opacity: Number(p.stock || 0) > 0 ? 1 : 0.5,
                    }}
                  >
                    <Plus size={18} />
                  </button>
                </div>
              </SectionCard>
            ))}
          </div>
          <Pagination page={productPage} totalPage={totalProductPage} onPageChange={setProductPage} />
        </div>
      </div>

      <SectionCard style={{ display: "flex", flexDirection: "column", overflow: "hidden", minHeight: 0 }}>
        <div
          style={{
            padding: 18,
            borderBottom: `1px solid ${COLORS.border}`,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
            flexWrap: "wrap",
          }}
        >
          <div style={{ fontSize: 22, fontWeight: 800, color: COLORS.text }}>Đơn hàng hiện tại</div>
          {paidMessage ? (
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 8,
                padding: "8px 12px",
                borderRadius: 999,
                background: COLORS.successSoft,
                color: COLORS.success,
                fontWeight: 700,
                fontSize: 13,
              }}
            >
              <CheckCircle2 size={16} /> {paidMessage}
            </div>
          ) : null}
        </div>

        <div style={{ flex: 1, overflowY: "auto", padding: 18, minHeight: 0 }}>
          {!cart.length ? (
            <EmptyState icon={Package} title="Chưa có sản phẩm nào trong giỏ hàng" />
          ) : (
            <div style={{ display: "grid", gap: 16 }}>
              {cart.map((i) => (
                <div key={i.id} style={{ borderBottom: `1px solid ${COLORS.border}`, paddingBottom: 14 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 16, lineHeight: 1.35 }}>{i.name}</div>
                      <div style={{ color: COLORS.textSoft, marginTop: 6 }}>{money(i.price)}</div>
                    </div>
                    <button
                      onClick={() => updateCartQty(i.id, 0)}
                      style={{ border: "none", background: "transparent", cursor: "pointer" }}
                    >
                      <Trash2 size={18} color="#94a3b8" />
                    </button>
                  </div>

                  <div style={{ marginTop: 12, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        border: `1px solid ${COLORS.border}`,
                        borderRadius: 12,
                        padding: 4,
                      }}
                    >
                      <button onClick={() => updateCartQty(i.id, i.qty - 1)} style={qtyBtn}>
                        <Minus size={14} />
                      </button>
                      <div style={{ minWidth: 30, textAlign: "center", fontWeight: 700 }}>{i.qty}</div>
                      <button onClick={() => updateCartQty(i.id, i.qty + 1)} style={qtyBtn}>
                        <Plus size={14} />
                      </button>
                    </div>
                    <div style={{ color: COLORS.primary, fontWeight: 800, fontSize: 16 }}>
                      {money(getCartItemTotal(i))}
                    </div>
                  </div>

                  <div style={{ marginTop: 12, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <div style={{ color: COLORS.textSoft, fontSize: 13, fontWeight: 600 }}>Giảm giá:</div>
                    {[0, 5, 10].map((d) => (
                      <DiscountButton
                        key={d}
                        active={Number(i.discount || 0) === d}
                        onClick={() => updateDiscount(i.id, d)}
                      >
                        {d}%
                      </DiscountButton>
                    ))}
                    <input
                      type="number"
                      placeholder="Nhập %"
                      value={i.discount ?? ""}
                      onChange={(e) => updateDiscount(i.id, Number(e.target.value || 0))}
                      style={{
                        width: 80,
                        padding: "6px 8px",
                        borderRadius: 8,
                        border: `1px solid ${COLORS.border}`,
                        outline: "none",
                      }}
                    />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div style={{ borderTop: `1px solid ${COLORS.border}`, padding: 18, display: "grid", gap: 16, flexShrink: 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", color: COLORS.textSoft, fontSize: 15 }}>
            <div>Tạm tính ({cart.reduce((s, i) => s + i.qty, 0)} món)</div>
            <div>{money(total)}</div>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 24, fontWeight: 800, color: COLORS.primary }}>
            <div>Tổng cộng</div>
            <div>{money(total)}</div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <PaymentButton
              active={paymentMethod === "Tiền mặt"}
              onClick={() => setPaymentMethod("Tiền mặt")}
              icon={BadgeDollarSign}
            >
              Tiền mặt
            </PaymentButton>
            <PaymentButton
              active={paymentMethod === "Chuyển khoản"}
              onClick={() => setPaymentMethod("Chuyển khoản")}
              icon={CreditCard}
            >
              Chuyển khoản
            </PaymentButton>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 14 }}>
            <button style={ghostBtn} onClick={() => printReceipt()}>
              <Printer size={18} /> In hóa đơn
            </button>

            <button
              style={ghostBtn}
              onClick={() => {
                setDeliveryForm({
                  name: "",
                  phone: "",
                  address: "",
                  payment: "Tiền mặt",
                });
                setShowDeliveryModal(true);
              }}
            >
              🚚 Giao hàng
            </button>

            <button
              style={ghostBtn}
              onClick={createTempOrder}
            >
              🧾 Đơn tạm
            </button>

            <button
              style={{ ...primaryBtn, opacity: total ? 1 : 0.55 }}
              onClick={checkout}
            >
              Thanh toán {money(total)}
            </button>

            {editingOrder && (
              <button style={ghostBtn} onClick={updateTempOrder}>
                💾 Cập nhật đơn tạm
              </button>
            )}
          </div>
        </div>
      </SectionCard>
    </div>
  );
  const attendancePage = (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: isMobile ? "1fr" : "1fr 400px",
        gap: 18,
        height: "100%",
      }}
    >

      {/* LEFT - NHÂN VIÊN */}
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>

        {/* ADD NV */}
        <SectionCard style={{ padding: 12 }}>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              placeholder="Tên nhân viên"
              value={employeeName}
              onChange={(e) => setEmployeeName(e.target.value)}
              style={cellInput}
            />
            <button style={primaryBtn} onClick={addEmployee}>
              +
            </button>
          </div>
        </SectionCard>

        {/* LIST NV */}
        <div style={{ display: "grid", gap: 10 }}>
          {employees.map(emp => (
            <SectionCard
              key={emp.id}
              onClick={() => setSelectedEmp(emp)}
              style={{
                padding: 14,
                cursor: "pointer",
                background:
                  selectedEmp?.id === emp.id
                    ? COLORS.primarySoft
                    : COLORS.white,
              }}
            >
              {emp.name}
            </SectionCard>
          ))}
        </div>
      </div>

      {/* RIGHT - CHẤM CÔNG */}
      <SectionCard style={{ display: "flex", flexDirection: "column" }}>

        <div style={{
          padding: 16,
          borderBottom: `1px solid ${COLORS.border}`
        }}>
          <strong>{selectedEmp?.name || "Chọn nhân viên"}</strong>
        </div>

        <div style={{ padding: 16 }}>

          {/* DATE */}
          <input
            type="date"
            value={attendanceDate}
            onChange={(e) => setAttendanceDate(e.target.value)}
            style={cellInput}
          />

          {/* BUTTON */}
          <div style={{ marginTop: 12, display: "flex", gap: 10 }}>
            <button style={primaryBtn} onClick={handleCheckIn}>
              Vào ca
            </button>
          </div>

          {/* DATA */}
          <div style={{ marginTop: 16 }}>
            {attendance
              .filter(a =>
                a.empId === selectedEmp?.id &&
                a.dateKey === attendanceDate
              )
              .map(a => (
                <div key={a.id} style={{ marginBottom: 10 }}>

                  {a.start} - {a.end || "..."} = {a.hours?.toFixed(2) || 0}h

                  {!a.end && (
                    <button
                      style={ghostBtn}
                      onClick={() => handleCheckOut(a)}
                    >
                      Ra ca
                    </button>
                  )}

                </div>
              ))}
          </div>

        </div>
      </SectionCard>

    </div>
  );

  const wholesalePage = (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: isMobile ? "1fr" : "minmax(0, 1fr) 400px",
        gap: 18,
        height: "100%",
        minHeight: 0,
      }}
    >

      {/* ===== LEFT: COPY Y CHANG SALES ===== */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 16,
          minWidth: 0,
          minHeight: 0,
          position: "relative",
          zIndex: 1,
        }}>

        {/* SEARCH */}
        <div style={{ position: "relative", maxWidth: 520 }}>
          <input
            placeholder="Tìm kiếm sản phẩm..."
            value={search || ""}
            onChange={(e) => setSearch(e.target.value)}
            style={{
              width: "100%",
              height: 56,
              paddingLeft: 16,
              borderRadius: 16,
              border: `1px solid ${COLORS.border}`,
              background: COLORS.white,
              fontSize: 16,
            }}
          />
        </div>

        <div
          style={{
            display: "flex",
            gap: 8,
            overflowX: "auto"
          }}
        >

          {/* TẤT CẢ */}
          <button
            style={{
              ...(activeCategory === "all" ? primaryBtn : ghostBtn),
              width: "auto",
              flexShrink: 0,
              whiteSpace: "nowrap",
            }}
            onClick={() => setActiveCategory("all")}
          >
            Tất cả
          </button>

          {/* CATEGORY */}
          {categories
            .filter(c => c !== "Tất cả")
            .map((c) => (
              <button
                key={c}
                style={{
                  ...(activeCategory === c ? primaryBtn : ghostBtn),
                  width: "auto",
                  flexShrink: 0,
                  whiteSpace: "nowrap",
                }}
                onClick={() => setActiveCategory(c)}
              >
                {c}
              </button>
            ))}

        </div>

        {/* NGƯỜI BÁN */}
        <div style={{ display: "grid", gap: 8 }}>
          <input
            placeholder="Tên người bán"
            value={sellerInfo.name}
            onChange={(e) =>
              setSellerInfo({ ...sellerInfo, name: e.target.value })
            }
            style={cellInput}
          />
          {/* KHÁCH HÀNG */}
          <div style={{ display: "grid", gap: 8 }}>
            <input
              placeholder="Tên khách hàng"
              value={customerInfo.name}
              onChange={(e) =>
                setCustomerInfo({ ...customerInfo, name: e.target.value })
              }
              style={cellInput}
            />
            <input
              placeholder="SĐT khách hàng"
              value={customerInfo.phone}
              onChange={(e) =>
                setCustomerInfo({ ...customerInfo, phone: e.target.value })
              }
              style={cellInput}
            />
          </div>

          {/* NHÂN VIÊN */}
        </div>

        {/* PRODUCTS */}
        <div style={{ flex: 1, overflowY: "auto", minHeight: 0, height: "100%" }}>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: isMobile
                ? "repeat(2,1fr)"
                : "repeat(4,1fr)",
              gap: 16,
            }}
          >
            {products
              .filter(p => {
                const s = (search || "").toLowerCase();

                const matchSearch = p.name
                  .toLowerCase()
                  .includes(s);

                const matchCate =
                  activeCategory === "all" || p.category === activeCategory;

                return matchSearch && matchCate;
              })
              .map((p) => (
                <SectionCard
                  key={p.id}
                  onClick={() => addToWholesaleCart(p)}
                  style={{ padding: 16, cursor: "pointer" }}
                >
                  <div>{p.name}</div>
                  <div style={{ color: COLORS.primary }}>
                    {money(p.price)}
                  </div>
                </SectionCard>
              ))}
          </div>
        </div>
      </div>

      {/* ===== RIGHT: UI SALES 100% - CHỈ ĐỔI LOGIC ===== */}
      <SectionCard style={{ display: "flex", flexDirection: "column" }}>

        {/* HEADER */}
        <div style={{
          padding: 18,
          borderBottom: `1px solid ${COLORS.border}`,
          fontWeight: 800
        }}>
          Đơn hàng hiện tại
        </div>

        {/* CART */}
        <div style={{ flex: 1, overflowY: "auto", padding: 16 }}>
          {!wholesaleCart.length ? (
            <div>Chưa có sản phẩm nào trong giỏ hàng</div>
          ) : (
            wholesaleCart.map(i => {
              const totalItem =
                (Number(i.customPrice) || 0) * (Number(i.qty) || 0);

              return (
                <div key={i.id} style={{
                  borderBottom: `1px solid ${COLORS.border}`,
                  marginBottom: 12,
                  paddingBottom: 12
                }}>

                  <div style={{
                    display: "flex",
                    justifyContent: "space-between"
                  }}>
                    <div>
                      <div style={{ fontWeight: 700 }}>{i.name}</div>
                      <div style={{ color: COLORS.textSoft }}>
                        {money(i.customPrice)}
                      </div>
                    </div>

                    <button onClick={() =>
                      setWholesaleCart(wholesaleCart.filter(x => x.id !== i.id))
                    }>✕</button>
                  </div>

                  {/* QTY */}
                  <div style={{
                    display: "flex",
                    justifyContent: "space-between",
                    marginTop: 8
                  }}>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        border: `1px solid ${COLORS.border}`,
                        borderRadius: 8,
                        padding: "2px 4px",
                        width: "fit-content"
                      }}
                    >
                      {/* Nút trừ */}
                      <button
                        style={qtyBtn}
                        onClick={() =>
                          setWholesaleCart(wholesaleCart.map(x =>
                            x.id === i.id
                              ? { ...x, qty: Math.max(1, i.qty - 1) }
                              : x
                          ))
                        }
                      >
                        -
                      </button>

                      {/* 🔥 INPUT NHẬP TAY */}
                      <input
                        type="number"
                        value={i.qty}
                        onChange={(e) => {
                          const value = Math.max(1, Number(e.target.value) || 1);
                          setWholesaleCart(wholesaleCart.map(x =>
                            x.id === i.id
                              ? { ...x, qty: value }
                              : x
                          ));
                        }}
                        style={{
                          width: 45,
                          textAlign: "center",
                          border: "none",
                          outline: "none",
                          fontWeight: 700,
                          fontSize: 14
                        }}
                      />

                      {/* Nút cộng */}
                      <button
                        style={qtyBtn}
                        onClick={() =>
                          setWholesaleCart(wholesaleCart.map(x =>
                            x.id === i.id
                              ? { ...x, qty: i.qty + 1 }
                              : x
                          ))
                        }
                      >
                        +
                      </button>
                    </div>

                    <div style={{ fontWeight: 800, color: COLORS.primary }}>
                      {money(totalItem)}
                    </div>
                  </div>

                  {/* GIÁ */}
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "1fr 1fr",
                      gap: 6,
                      marginTop: 8
                    }}
                  >

                    {/* Giá sỉ */}
                    <div
                      style={{
                        background: "#f1f5f9",
                        border: `1px solid ${COLORS.border}`,
                        borderRadius: 10,
                        padding: "2px 6px",
                        display: "flex",
                        flexDirection: "column"
                      }}
                    >
                      <span
                        style={{
                          fontSize: 10,
                          color: "#64748b",
                          fontWeight: 600,
                          marginBottom: 1
                        }}
                      >
                        Giá sỉ
                      </span>

                      <input
                        value={i.customPrice}
                        onChange={(e) =>
                          setWholesaleCart(wholesaleCart.map(x =>
                            x.id === i.id
                              ? { ...x, customPrice: Number(e.target.value) }
                              : x
                          ))
                        }
                        style={{
                          border: "none",
                          outline: "none",
                          background: "transparent",
                          fontWeight: 700,
                          fontSize: 13
                        }}
                      />
                    </div>

                    {/* Giá chiết khấu */}
                    <div
                      style={{
                        background: "#fef2f2",
                        border: "1px solid #fecaca",
                        borderRadius: 10,
                        padding: "2px 6px",
                        display: "flex",
                        flexDirection: "column"
                      }}
                    >
                      <span
                        style={{
                          fontSize: 10,
                          color: "#ef4444",
                          fontWeight: 600,
                          marginBottom: 1
                        }}
                      >
                        Giá chiết khấu
                      </span>

                      <input
                        value={i.discountCash}
                        onChange={(e) =>
                          setWholesaleCart(wholesaleCart.map(x =>
                            x.id === i.id
                              ? { ...x, discountCash: Number(e.target.value) }
                              : x
                          ))
                        }
                        style={{
                          border: "none",
                          outline: "none",
                          background: "transparent",
                          fontWeight: 700,
                          fontSize: 13,
                          color: "#ef4444"
                        }}
                      />
                    </div>
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* FOOTER (GIỐNG SALES) */}
        <div style={{
          borderTop: `1px solid ${COLORS.border}`,
          padding: 18
        }}>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <span>Tạm tính</span>
            <span>{money(getWholesaleTotal())}</span>
          </div>

          <div style={{
            display: "flex",
            justifyContent: "space-between",
            fontSize: 24,
            fontWeight: 800,
            marginTop: 6
          }}>
            <span>Tổng cộng</span>
            <span>{money(getWholesaleTotal())}</span>
          </div>

          <div style={{ color: "#888" }}>
            Chiết khấu: {money(getWholesaleDiscount())}
          </div>

          {/* PAYMENT */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 12 }}>
            <PaymentButton
              active={wholesalePayment === "Tiền mặt"}
              onClick={() => setWholesalePayment("Tiền mặt")}
              icon={BadgeDollarSign}
            >
              Tiền mặt
            </PaymentButton>

            <PaymentButton
              active={wholesalePayment === "Chuyển khoản"}
              onClick={() => setWholesalePayment("Chuyển khoản")}
              icon={CreditCard}
            >
              Chuyển khoản
            </PaymentButton>
          </div>

          {/* ACTION */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr 1fr 1fr",
              gap: 14,
              marginTop: 10
            }}
          >
            <button style={ghostBtn}>
              <Printer size={18} /> In hóa đơn
            </button>

            <button style={ghostBtn}>
              🚚 Giao hàng
            </button>

            <button
              style={ghostBtn}
              onClick={saveWholesaleTemp}   // 🔥 thêm dòng này
            >
              🧾 Đơn tạm
            </button>

            {editingWholesaleOrder && (
              <button
                style={ghostBtn}
                onClick={updateWholesaleTempOrder}
              >
                💾 Cập nhật đơn sỉ
              </button>
            )}

            <button
              style={{ ...primaryBtn, opacity: getWholesaleTotal() ? 1 : 0.55 }}
              onClick={createWholesaleOrder}
            >
              Thanh toán {money(getWholesaleTotal())}
            </button>
          </div>

        </div>

      </SectionCard>
    </div>
  );

  const historyPage = (
    <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "440px 1fr", gap: 18, height: "100%", minHeight: 0 }}>
      <SectionCard style={{ overflow: "hidden", display: "flex", flexDirection: "column", minHeight: 0 }}>
        <div style={{ padding: 18, borderBottom: `1px solid ${COLORS.border}` }}>
          <div style={pageTitle}>Lịch sử đơn hàng</div>
          <div style={{ marginTop: 16, display: "grid", gap: 12 }}>
            <div style={{ position: "relative" }}>
              <Calendar size={18} color="#111" style={{ position: "absolute", right: 16, top: 13 }} />
              <input
                type="date"
                value={historyDate}
                onChange={(e) => setHistoryDate(e.target.value)}
                style={{
                  width: "100%",
                  padding: "12px 14px",
                  borderRadius: 14,
                  border: `1px solid ${COLORS.border}`,
                  fontSize: 16,
                  background: "#f8fafc",
                  boxSizing: "border-box",
                }}
              />
            </div>
            <div style={{ position: "relative" }}>
              <Search size={18} color="#64748b" style={{ position: "absolute", left: 14, top: 12 }} />
              <input
                value={historySearch}
                onChange={(e) => setHistorySearch(e.target.value)}
                placeholder="Tìm mã đơn hàng..."
                style={{
                  width: "100%",
                  padding: "12px 14px 12px 40px",
                  borderRadius: 14,
                  border: `1px solid ${COLORS.border}`,
                  boxSizing: "border-box",
                }}
              />
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
              <button
                onClick={() => setHistoryType("day")}
                style={{
                  ...ghostBtn,
                  background: historyType === "day" ? COLORS.primary : COLORS.white,
                  color: historyType === "day" ? "#fff" : COLORS.text,
                }}
              >
                Ngày
              </button>

              <button
                onClick={() => setHistoryType("week")}
                style={{
                  ...ghostBtn,
                  background: historyType === "week" ? COLORS.primary : COLORS.white,
                  color: historyType === "week" ? "#fff" : COLORS.text,
                }}
              >
                Tuần
              </button>

              <button
                onClick={() => setHistoryType("month")}
                style={{
                  ...ghostBtn,
                  background: historyType === "month" ? COLORS.primary : COLORS.white,
                  color: historyType === "month" ? "#fff" : COLORS.text,
                }}
              >
                Tháng
              </button>
            </div>
            <button style={ghostBtn} onClick={exportOrdersExcel}>
              <Download size={18} /> Xuất Excel lịch sử
            </button>
          </div>
          <div style={{
            display: "flex",
            gap: 8,
            marginTop: 10,
            background: "#f1f5f9",
            padding: 6,
            borderRadius: 12
          }}>

            <button
              onClick={() => setFilterType("all")}
              style={{
                flex: 1,
                padding: "8px 0",
                borderRadius: 10,
                border: "none",
                cursor: "pointer",
                background: filterType === "all" ? "#2f66e9" : "transparent",
                color: filterType === "all" ? "#fff" : "#334155",
                fontWeight: 600
              }}
            >
              Tất cả
            </button>

            <button
              onClick={() => setFilterType("temp")}
              style={{
                flex: 1,
                padding: "8px 0",
                borderRadius: 10,
                border: "none",
                cursor: "pointer",
                background: filterType === "temp" ? "#2f66e9" : "transparent",
                color: filterType === "temp" ? "#fff" : "#334155",
                fontWeight: 600
              }}
            >
              Đơn tạm
            </button>

            <button
              style={filterType === "wholesale" ? primaryBtn : ghostBtn}
              onClick={() => setFilterType("wholesale")}
            >
              Sỉ
            </button>

            <button
              onClick={() => setFilterType("delivery")}
              style={{
                flex: 1,
                padding: "8px 0",
                borderRadius: 10,
                border: "none",
                cursor: "pointer",
                background: filterType === "delivery" ? "#2f66e9" : "transparent",
                color: filterType === "delivery" ? "#fff" : "#334155",
                fontWeight: 600
              }}
            >
              Giao hàng
            </button>

          </div>
        </div>

        <div style={{ flex: 1, overflow: "auto" }}>
          {!historyFiltered.length ? (
            <div style={{ height: "100%" }}>
              <EmptyState icon={Receipt} title="Không có đơn hàng nào" />
            </div>
          ) : (
            historyFiltered
              .filter(o => {
                if (filterType === "all") return true;
                if (filterType === "temp") return o.status === "Đơn tạm";
                if (filterType === "delivery") return o.isDelivery;
                if (filterType === "wholesale") return o.type === "wholesale";
                return true;
              }).map((o) => (
                <button
                  key={o.id}
                  onClick={() => {
                    setSelectedOrder(o);

                    if (o.isDelivery) {
                      setEditingOrder(o);
                      setDeliveryForm({
                        name: o.customer?.name || "",
                        phone: o.customer?.phone || "",
                        address: o.customer?.address || "",
                        payment: o.method || "Tiền mặt"
                      });

                      setShowDeliveryModal(true);
                    }
                  }}
                  style={{
                    width: "100%",
                    textAlign: "left",
                    border: "none",
                    background: selectedOrder?.id === o.id ? "#eef4ff" : COLORS.white,
                    padding: 16,
                    borderBottom: `1px solid ${COLORS.border}`,
                    cursor: "pointer",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
                    <div style={{ fontWeight: 700, color: COLORS.text }}>
                      {o.code}

                      {o.type === "wholesale" && (
                        <span
                          style={{
                            marginLeft: 8,
                            padding: "2px 6px",
                            background: "#e9f0ff",
                            color: "#2f66e9",
                            borderRadius: 6,
                            fontSize: 12,
                            fontWeight: 700
                          }}
                        >
                          SỈ
                        </span>
                      )}
                    </div>

                    {/* 🔥 THÊM NGAY DƯỚI */}
                    {o.type === "wholesale" && (
                      <div style={{ fontSize: 13, color: "#555", marginTop: 4 }}>
                        👤 {o.customerName || "Chưa có tên"} - {o.customerPhone || ""}
                      </div>
                    )}
                    <div style={{ color: COLORS.success, fontSize: 13, fontWeight: 700 }}>
                      {o.status || "Đã thanh toán"}
                    </div>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, color: COLORS.textSoft }}>
                    <span>{o.timeText}</span>
                    <span style={{ color: COLORS.primary, fontWeight: 700 }}>{money(o.total)}</span>
                  </div>
                </button>
              ))
          )}
        </div>
      </SectionCard>

      <SectionCard style={{ overflow: "hidden", minHeight: 0 }}>
        {!selectedOrder ? (
          <div style={{ height: "100%" }}>
            <EmptyState icon={Receipt} title="Chọn một đơn hàng để xem chi tiết" />
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
            <div
              style={{
                padding: 18,
                borderBottom: `1px solid ${COLORS.border}`,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 12,
                flexWrap: "wrap",
              }}
            >
              <div>
                <div style={{ fontSize: 24, fontWeight: 800 }}>{selectedOrder.code}</div>
                <div style={{ color: COLORS.textSoft, marginTop: 6 }}>{selectedOrder.timeText}</div>
                <div style={{ color: COLORS.success, marginTop: 6, fontWeight: 700 }}>
                  {selectedOrder.status || "Đã thanh toán"} • {selectedOrder.method}
                </div>
                {selectedOrder.type === "wholesale" && (
                  <div style={{ marginTop: 6, fontSize: 14, color: "#555" }}>
                    {selectedOrder.customerName || "Khách lẻ"}
                    - {selectedOrder.customerPhone || "---"}
                    - {selectedOrder.sellerName || "---"}
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={ghostBtn} onClick={() => printReceipt(selectedOrder)}>
                  <Printer size={18} /> In lại
                </button>

                {selectedOrder?.status === "Đơn tạm" && (
                  <>
                    <button
                      style={ghostBtn}
                      onClick={() => loadTempOrderToCart(selectedOrder)}
                    >
                      ✏️ Sửa đơn
                    </button>

                    <button
                      style={primaryBtn}
                      onClick={() => {
                        setTempOrder(selectedOrder);
                        setShowPaymentModal(true);
                      }}
                    >
                      Xác nhận thanh toán
                    </button>
                  </>
                )}

                {selectedOrder?.status === "Chờ giao" && (
                  <button
                    style={primaryBtn}
                    onClick={() => completeDelivery(selectedOrder)}
                  >
                    Đã giao xong
                  </button>
                )}

                <button
                  style={{ ...ghostBtn, color: COLORS.danger, borderColor: "#fecaca" }}
                  onClick={() => deleteOrder(selectedOrder.id)}
                >
                  <Trash2 size={18} /> Xóa
                </button>
              </div>
            </div>

            <div style={{ padding: 18, flex: 1, overflow: "auto" }}>
              {(selectedOrder.items.map((i) => {
                const price = Number(i.customPrice || i.price || 0);
                const qty = Number(i.qty || 0);

                return (
                  <div
                    key={i.id}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      marginBottom: 8
                    }}
                  >
                    <div>
                      <div style={{ fontWeight: 600 }}>
                        {i.name} x{qty}
                      </div>

                      {/* 🔥 HIỆN GIÁ SỈ */}
                      {selectedOrder.type === "wholesale" && (
                        <div style={{ fontSize: 12, color: "#64748b" }}>
                          Giá sỉ: {money(price)}
                        </div>
                      )}
                    </div>

                    <div style={{ fontWeight: 700 }}>
                      {money(price * qty)}
                    </div>
                  </div>
                );
              }))}
            </div>

            <div
              style={{
                padding: 18,
                borderTop: `1px solid ${COLORS.border}`,
                display: "flex",
                justifyContent: "space-between",
                fontSize: 24,
                fontWeight: 800,
                color: COLORS.primary,
              }}
            >
              <span>Tổng cộng</span>
              <span>{money(selectedOrder.total)}</span>
            </div>
          </div>
        )}
      </SectionCard>
    </div>
  );

  const productsPage = (
    <div style={{ display: "grid", gap: 18, minHeight: 0 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div style={pageTitle}>Sản phẩm</div>
          <div style={pageSub}>Quản lý danh sách sản phẩm và giá bán</div>
        </div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <button style={ghostBtn} onClick={exportProductsExcel}>
            <Download size={18} /> Xuất Excel
          </button>
          <button
            style={primaryBtn}
            onClick={() => {
              setEditingProduct(null);
              setShowProductModal(true);
            }}
          >
            <Plus size={18} /> Thêm sản phẩm
          </button>
        </div>
      </div>

      <SectionCard style={{ overflow: "hidden", minHeight: 0 }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 14,
            padding: 16,
            borderBottom: `1px solid ${COLORS.border}`,
            flexWrap: "wrap",
          }}
        >
          <div style={{ position: "relative", width: 320, maxWidth: "100%" }}>
            <Search size={18} color="#64748b" style={{ position: "absolute", left: 14, top: 12 }} />
            <input
              value={productSearch}
              onChange={(e) => setProductSearch(e.target.value)}
              placeholder="Tìm tên, danh mục..."
              style={{
                width: "100%",
                padding: "11px 12px 11px 40px",
                borderRadius: 12,
                border: `1px solid ${COLORS.border}`,
                outline: "none",
                background: "#f8fafc",
                boxSizing: "border-box",
              }}
            />
          </div>
          <div style={{ color: COLORS.textSoft }}>{filteredProductTable.length} sản phẩm</div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: isMobile
              ? "1.8fr 1.2fr 1fr 1fr"
              : "2.2fr 1.3fr 1fr 1fr 1fr 0.8fr",
            padding: "14px 16px",
            background: "#f8fafc",
            color: COLORS.textSoft,
            fontWeight: 700,
          }}
        >
          <div>Tên sản phẩm</div>
          <div>Danh mục</div>
          <div>Giá bán</div>
          <div>Tồn kho</div>
          {!isMobile && (
            <>
              <div>Giá vốn</div>
              <div>Thao tác</div>
            </>
          )}
        </div>

        <div style={{ maxHeight: "62vh", overflow: "auto" }}>
          {filteredProductTable.map((p) => (
            <div
              key={p.id}
              style={{
                display: "grid",
                gridTemplateColumns: isMobile
                  ? "1.8fr 1.2fr 1fr 1fr"
                  : "2.2fr 1.3fr 1fr 1fr 1fr 0.8fr",
                padding: "14px 16px",
                borderTop: `1px solid ${COLORS.border}`,
                alignItems: "center",
                background: COLORS.white,
                gap: 10,
              }}
            >
              <input
                defaultValue={p.name}
                onBlur={(e) => quickEditProduct(p.id, "name", e.target.value)}
                style={cellInput}
              />
              <select
                defaultValue={p.category}
                onChange={(e) => quickEditProduct(p.id, "category", e.target.value)}
                style={cellInput}
              >
                {categories
                  .filter((c) => c !== "Tất cả")
                  .map((c) => (
                    <option key={c}>{c}</option>
                  ))}
              </select>
              <input
                type="number"
                defaultValue={p.price}
                onBlur={(e) => quickEditProduct(p.id, "price", Number(e.target.value || 0))}
                style={{ ...cellInput, color: COLORS.primary, fontWeight: 700 }}
              />
              <input
                type="number"
                defaultValue={p.stock || 0}
                onBlur={(e) => quickEditProduct(p.id, "stock", Number(e.target.value || 0))}
                style={cellInput}
              />
              {!isMobile && (
                <>
                  <input
                    type="number"
                    defaultValue={p.cost || 0}
                    onBlur={(e) => quickEditProduct(p.id, "cost", Number(e.target.value || 0))}
                    style={cellInput}
                  />
                  <div style={{ display: "flex", gap: 10 }}>
                    <button
                      style={iconBtn}
                      onClick={() => {
                        setEditingProduct(p);
                        setShowProductModal(true);
                      }}
                    >
                      <Pencil size={16} />
                    </button>
                  </div>
                </>
              )}
            </div>
          ))}
        </div>
      </SectionCard>
    </div>
  );

  const expensePage = (
    <div style={{ display: "grid", gap: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div style={pageTitle}>Danh mục chi phí</div>
          <div style={pageSub}>Theo dõi các khoản chi tiêu trong ngày</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div style={{ color: COLORS.text, fontWeight: 600 }}>Ngày:</div>
          <div style={{ position: "relative" }}>
            <Calendar size={18} style={{ position: "absolute", right: 14, top: 12 }} />
            <input
              type="date"
              value={expenseDate}
              onChange={(e) => setExpenseDate(e.target.value)}
              style={{ padding: "11px 40px 11px 14px", borderRadius: 14, border: `1px solid ${COLORS.border}` }}
            />
          </div>
          <button style={primaryBtn} onClick={addExpense}>
            <Plus size={18} /> Thêm chi phí
          </button>
        </div>
      </div>

      <SectionCard style={{ width: isMobile ? "100%" : 360, padding: 26, background: "#fff3f3" }}>
        <div style={{ color: COLORS.textSoft, fontSize: 16 }}>Tổng chi phí</div>
        <div style={{ fontSize: 40, fontWeight: 800, color: COLORS.danger, marginTop: 14 }}>
          {money(totalExpenseToday)}
        </div>
        <div style={{ color: COLORS.textSoft, marginTop: 6 }}>{expenseFiltered.length} khoản chi</div>
      </SectionCard>

      <SectionCard style={{ overflow: "hidden" }}>
        <div style={{ padding: 18, borderBottom: `1px solid ${COLORS.border}`, fontSize: 20, fontWeight: 700 }}>
          Danh sách chi phí
        </div>
        <div style={{ minHeight: 360, maxHeight: "55vh", overflow: "auto" }}>
          {!expenseFiltered.length ? (
            <div style={{ height: 320 }}>
              <EmptyState icon={Wallet} title="Chưa có chi phí nào trong ngày" />
            </div>
          ) : (
            expenseFiltered.map((e) => (
              <div
                key={e.id}
                style={{
                  padding: 18,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  gap: 12,
                  borderTop: `1px solid ${COLORS.border}`,
                }}
              >
                <div>
                  <div style={{ fontWeight: 700 }}>{e.name}</div>
                  <div style={{ color: COLORS.textSoft, marginTop: 4 }}>{e.timeText}</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <div style={{ color: COLORS.danger, fontWeight: 800 }}>{money(e.amount)}</div>
                  <button onClick={() => updateExpense(e.id, e)} style={iconBtn}>
                    <Pencil size={16} />
                  </button>
                  <button onClick={() => deleteExpenseItem(e.id)} style={{ ...iconBtn, borderColor: "#fecaca" }}>
                    <Trash2 size={16} color="#ef4444" />
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
      </SectionCard>
    </div>
  );

  const productStats = {};

  orders
    .filter(o => o.status === "Đã thanh toán")
    .forEach(o => {
      (o.items || []).forEach(item => {
        const product = products.find(p => p.name === item.name);

        if (
          reportCategory === "Tất cả" ||
          (product && product.category === reportCategory)
        ) {
          if (!productStats[item.name]) {
            productStats[item.name] = {
              qty: 0,
              total: 0
            };
          }

          productStats[item.name].qty += item.qty;
          productStats[item.name].total += item.qty * item.price;
        }
      });
    });

  const productList = Object.entries(productStats)
    .map(([name, data]) => ({
      name,
      qty: data.qty,
      total: data.total
    }))
    .sort((a, b) => b.qty - a.qty);



  const reportPage = (
    <div style={{ display: "grid", gap: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div style={pageTitle}>Báo cáo doanh thu</div>
          <div style={pageSub}>Tổng quan tình hình kinh doanh trong ngày</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div style={{ color: COLORS.text, fontWeight: 600 }}>Ngày báo cáo:</div>
          <div style={{ position: "relative" }}>
            <Calendar size={18} style={{ position: "absolute", right: 14, top: 12 }} />
            <input
              type="date"
              value={reportDate}
              onChange={(e) => setReportDate(e.target.value)}
              style={{ padding: "11px 40px 11px 14px", borderRadius: 14, border: `1px solid ${COLORS.border}` }}
            />
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              onClick={() => setReportType("day")}
              style={{
                ...ghostBtn,
                background: reportType === "day" ? COLORS.primary : COLORS.white,
                color: reportType === "day" ? "#fff" : COLORS.text,
              }}
            >
              Ngày
            </button>

            <button
              onClick={() => setReportType("week")}
              style={{
                ...ghostBtn,
                background: reportType === "week" ? COLORS.primary : COLORS.white,
                color: reportType === "week" ? "#fff" : COLORS.text,
              }}
            >
              Tuần
            </button>

            <button
              onClick={() => setReportType("month")}
              style={{
                ...ghostBtn,
                background: reportType === "month" ? COLORS.primary : COLORS.white,
                color: reportType === "month" ? "#fff" : COLORS.text,
              }}
            >
              Tháng
            </button>
          </div>
          <button style={ghostBtn} onClick={exportReportExcel}>
            <Download size={18} /> Xuất Excel
          </button>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.2fr 1.2fr 1.1fr", gap: 18 }}>
        <MetricCard title="Tổng doanh thu" value={money(revenue)} sub={`Lợi nhuận gộp: ${money(revenue)}`} color="#2f66e9" soft="#eef4ff" />
        <MetricCard title="Tổng chi phí" value={money(cost)} sub="Chi phí vận hành trong ngày" color="#ef4444" soft="#fff1f2" />
        <MetricCard title="Lợi nhuận ròng" value={money(profit)} sub="Sau khi trừ chi phí" color="#0f9d58" soft="#effaf4" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr 1fr", gap: 18 }}>
        <MetricCard title="Tiền mặt" value={money(cashRevenue)} color="#111827" soft="#ffffff" icon={<BadgeDollarSign size={18} color="#22c55e" />} />
        <MetricCard title="Chuyển khoản" value={money(bankRevenue)} color="#111827" soft="#ffffff" icon={<CreditCard size={18} color="#2f66e9" />} />
        <MetricCard title="Số đơn hàng" value={`${orderCount} đơn`} color="#111827" soft="#ffffff" icon={<ShoppingBag size={18} color="#f97316" />} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.35fr 1fr", gap: 18 }}>
        <SectionCard style={{ minHeight: 340, padding: 18 }}>
          <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 12 }}>
            SẢN PHẨM ĐÃ BÁN
          </div>

          <div style={{ maxHeight: 280, overflow: "auto" }}>
            {!productList.length ? (
              <div style={{ height: 220 }}>
                <EmptyState icon={Package} title="Chưa có dữ liệu bán hàng" />
              </div>
            ) : (
              productList.map((p, i) => (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    padding: "10px 0",
                    borderBottom: `1px solid ${COLORS.border}`,
                    alignItems: "center"
                  }}
                >
                  <div style={{ fontWeight: 500 }}>
                    {i === 0 ? "🔥 " : ""}{p.name}
                  </div>

                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontWeight: 600 }}>
                      {p.qty} món
                    </div>
                    <div style={{ fontSize: 12, color: "#64748b" }}>
                      {money(p.total)}
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        </SectionCard>

        {/* ===== THỐNG KÊ SẢN PHẨM ===== */}
        <SectionCard style={{ marginTop: 18 }}>
          <div style={{
            padding: 18,
            borderBottom: `1px solid ${COLORS.border}`,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center"
          }}>
            <div style={{ fontSize: 20, fontWeight: 700 }}>
              SẢN PHẨM ĐÃ BÁN TRONG NGÀY
            </div>

            <input
              placeholder="Tìm sản phẩm..."
              value={reportProductSearch}
              onChange={(e) => setReportProductSearch(e.target.value)}
              style={{
                padding: "8px 12px",
                borderRadius: 10,
                border: `1px solid ${COLORS.border}`
              }}
            />
          </div>

          <div style={{
            display: "grid",
            gridTemplateColumns: "2fr 1fr 1fr",
            padding: "12px 18px",
            background: "#f8fafc",
            fontWeight: 700
          }}>
            <div>Sản phẩm</div>
            <div>Số lượng</div>
            <div>Doanh thu</div>
          </div>
          <div style={{
            display: "flex",
            gap: 8,
            padding: "10px 18px",
            borderBottom: `1px solid ${COLORS.border}`,
            flexWrap: "wrap"
          }}>
            {categories.map((c) => (
              <button
                key={c}
                onClick={() => setReportCategory(c)}
                style={{
                  padding: "5px 10px",
                  borderRadius: 8,
                  border: "none",
                  cursor: "pointer",
                  background:
                    reportCategory === c ? COLORS.primary : "#f1f5f9",
                  color: reportCategory === c ? "#fff" : "#334155",
                  fontWeight: 600,
                  fontSize: 13
                }}
              >
                {c}
              </button>
            ))}
          </div>

          <div style={{ maxHeight: 300, overflow: "auto" }}>
            {!filteredReportProducts.length ? (
              <EmptyState icon={Package} title="Chưa có dữ liệu" />
            ) : (
              filteredReportProducts.map((p, i) => (
                <div key={i} style={{
                  display: "grid",
                  gridTemplateColumns: "2fr 1fr 1fr",
                  padding: "12px 18px",
                  borderTop: `1px solid ${COLORS.border}`
                }}>
                  <div>{p.name}</div>
                  <div>{p.qty}</div>
                  <div>{money(p.revenue)}</div>
                </div>
              ))
            )}
          </div>
        </SectionCard>

        <SectionCard style={{ padding: 20, marginTop: 20 }}>
          <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 12 }}>
            Thống kê đơn sỉ
          </div>

          {!wholesaleStats.length ? (
            <div>Không có dữ liệu</div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
                  <th>Người bán</th>
                  <th>Số đơn</th>
                  <th>Doanh thu</th>
                  <th>Chiết khấu</th>
                </tr>
              </thead>
              <tbody>
                {wholesaleStats.map((s, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #eee" }}>
                    <td>{s.seller}</td>
                    <td>{s.orders}</td>
                    <td>{money(s.revenue)}</td>
                    <td>{money(s.discount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </SectionCard>
      </div>
    </div>
  );

  return (
    <div
      style={{
        height: "100vh",
        background: COLORS.bg,
        color: COLORS.text,
        fontFamily: "Inter, system-ui, sans-serif",
        display: isMobile ? "block" : "grid",
        gridTemplateColumns: isMobile ? undefined : "280px 1fr",
      }}
    >
      {!isMobile ? (
        <aside
          style={{
            background: COLORS.sidebar,
            color: "#fff",
            padding: 14,
            display: "flex",
            flexDirection: "column",
            borderRight: `1px solid ${COLORS.sidebarBorder}`,
            height: "100vh",
          }}
        >
          <div
            style={{
              height: 70,
              display: "flex",
              alignItems: "center",
              gap: 14,
              padding: "0 12px",
              borderBottom: `1px solid ${COLORS.sidebarBorder}`,
              marginBottom: 14,
            }}
          >
            <div
              style={{
                width: 36,
                height: 36,
                borderRadius: 10,
                border: "2px solid #3b82f6",
                display: "grid",
                placeItems: "center",
              }}
            >
              <Home size={22} color="#3b82f6" />
            </div>
            <div style={{ fontSize: 22, fontWeight: 800 }}>AN VAT BUNBUN</div>
          </div>
          <div style={{ display: "grid", gap: 10 }}>
            {pages.map((item) => (
              <SidebarItem
                key={item.key}
                icon={item.icon}
                label={item.label}
                active={page === item.key}
                onClick={() => setPage(item.key)}
              />
            ))}
          </div>
          <div style={{ marginTop: "auto", color: "#94a3b8", fontSize: 14, padding: 12 }}>
            © 2026 AN VAT BUNBUN
          </div>
        </aside>
      ) : (
        <div
          style={{
            background: COLORS.white,
            padding: 12,
            borderBottom: `1px solid ${COLORS.border}`,
            position: "sticky",
            top: 0,
            zIndex: 20,
          }}
        >
          <div style={{ fontSize: 22, fontWeight: 800, marginBottom: 10 }}>AN VAT BUNBUN</div>
          <div style={{ display: "flex", gap: 8, overflowX: "auto" }}>
            {pages.map((item) => (
              <SidebarItem
                key={item.key}
                icon={item.icon}
                label={item.label}
                active={page === item.key}
                onClick={() => setPage(item.key)}
                mobile
              />
            ))}
          </div>
        </div>
      )}

      <main
        style={{
          padding: 18,
          overflow: isMobile ? "visible" : "hidden",
          height: isMobile ? "auto" : "100vh",
        }}
      >
        {page === "sales" && salesPage}
        {page === "products" && productsPage}
        {page === "history" && historyPage}
        {page === "expense" && expensePage}
        {page === "report" && reportPage}
        {page === "wholesale" && wholesalePage}
        {page === "attendance" && attendancePage}
      </main>

      <ProductModal
        open={showProductModal}
        onClose={() => {
          setShowProductModal(false);
          setEditingProduct(null);
        }}
        onSave={saveProduct}
        categories={categories}
        product={editingProduct}
      />
      <PaymentModal
        open={showPaymentModal}
        onClose={() => setShowPaymentModal(false)}
        onConfirm={(method) => {
          confirmTempOrder(tempOrder, method);
          setShowPaymentModal(false);
          setSelectedOrder(null);
        }}
      />
      <DeliveryModal
        open={showDeliveryModal}
        onClose={() => {
          setShowDeliveryModal(false);
          setEditingOrder(null);
        }}
        onConfirm={createDeliveryOrder}
        editingOrder={editingOrder}   // 👈 THÊM
        form={deliveryForm}
        setForm={setDeliveryForm}
      />
    </div>
  );
}
function DeliveryModal({ open, onClose, onConfirm, editingOrder, form, setForm }) {
  if (!open) return null;

  return (
    <div style={{
      position: "fixed",
      top: 0, left: 0,
      width: "100%", height: "100%",
      background: "rgba(0,0,0,0.3)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 999
    }}>
      <div style={{
        background: "#fff",
        padding: 20,
        borderRadius: 10,
        width: 350
      }}>
        <h3>Giao hàng</h3>

        <input
          placeholder="Tên khách"
          value={form.name}
          onChange={(e) => setForm({ ...form, name: e.target.value })}
          style={{ width: "100%", marginBottom: 12 }}
        />

        <input
          placeholder="SĐT"
          value={form.phone}
          onChange={(e) => setForm({ ...form, phone: e.target.value })}
          style={{ width: "100%", marginBottom: 12 }}
        />

        <input
          placeholder="Địa chỉ"
          value={form.address}
          onChange={(e) => setForm({ ...form, address: e.target.value })}
          style={{ width: "100%", marginBottom: 12 }}
        />

        <select
          value={form.payment}
          onChange={(e) => setForm({ ...form, payment: e.target.value })}
          style={{ width: "100%", marginBottom: 12 }}
        >
          <option value="Tiền mặt">Tiền mặt</option>
          <option value="Chuyển khoản">Chuyển khoản</option>
          <option value="Nợ">Đơn nợ</option>
        </select>

        <div style={{ display: "flex", gap: 12 }}>
          <button onClick={onConfirm}>Xác nhận</button>
          <button onClick={onClose}>Đóng</button>
        </div>
      </div>
    </div>
  );
}
function PaymentModal({ open, onClose, onConfirm }) {
  if (!open) return null;

  return (
    <div style={{
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,0.4)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 999
    }}>
      <div style={{
        background: "#fff",
        padding: 24,
        borderRadius: 16,
        width: 320,
        boxShadow: "0 10px 30px rgba(0,0,0,0.2)",
        textAlign: "center",
        animation: "fadeIn 0.2s ease"
      }}>

        <h3 style={{
          margin: 0,
          marginBottom: 20,
          fontSize: 20,
          fontWeight: 700
        }}>
          Chọn phương thức thanh toán
        </h3>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>

          <button
            onClick={() => onConfirm("Tiền mặt")}
            style={{
              padding: "14px",
              borderRadius: 12,
              border: "none",
              background: "#16a34a",
              color: "#fff",
              fontWeight: 700,
              cursor: "pointer",
              fontSize: 15
            }}
          >
            💵 Tiền mặt
          </button>

          <button
            onClick={() => onConfirm("Chuyển khoản")}
            style={{
              padding: "14px",
              borderRadius: 12,
              border: "none",
              background: "#2f66e9",
              color: "#fff",
              fontWeight: 700,
              cursor: "pointer",
              fontSize: 15
            }}
          >
            💳 Chuyển khoản
          </button>
        </div>

        <button
          onClick={onClose}
          style={{
            marginTop: 16,
            background: "transparent",
            border: "none",
            color: "#64748b",
            cursor: "pointer",
            fontSize: 14
          }}
        >
          Hủy
        </button>
      </div>
    </div>
  );
}