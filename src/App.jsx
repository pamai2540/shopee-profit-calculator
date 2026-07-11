import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ================= DESIGN TOKENS =================
   สมุดบัญชีแม่ค้า digital ledger:
   ink green #123B2F / paper #F5F8F4 / line #D8E4DA
   profit green #1E7A4C / loss red #C23B3B / amber #C98A12
   Numbers = tabular, receipt-style waterfall as signature
=================================================== */

const T = {
  ink: "#123B2F",
  paper: "#F5F8F4",
  card: "#FFFFFF",
  line: "#D8E4DA",
  dim: "#5C7266",
  profit: "#1E7A4C",
  loss: "#C23B3B",
  amber: "#C98A12",
};

const fmt = (n, d = 2) =>
  isNaN(n) || n === null
    ? "–"
    : n.toLocaleString("th-TH", { minimumFractionDigits: d, maximumFractionDigits: d });

const num = (v) => {
  if (typeof v === "number") return v;
  if (v === undefined || v === null) return 0;
  const c = parseFloat(String(v).replace(/[,฿\s]/g, ""));
  return isNaN(c) ? 0 : c;
};

/* ---------- shared small components ---------- */

const Field = ({ label, value, onChange, suffix, hint, step = "any" }) => (
  <label style={{ display: "block", marginBottom: 14 }}>
    <div style={{ fontSize: 13, color: T.dim, marginBottom: 4 }}>{label}</div>
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <input
        type="number"
        step={step}
        value={value}
        onChange={(e) => onChange(e.target.value === "" ? "" : Number(e.target.value))}
        style={{
          flex: 1,
          padding: "10px 12px",
          border: `1px solid ${T.line}`,
          borderRadius: 8,
          fontSize: 16,
          fontVariantNumeric: "tabular-nums",
          background: "#fff",
          color: T.ink,
          outline: "none",
          width: "100%",
        }}
      />
      {suffix && <span style={{ fontSize: 13, color: T.dim, whiteSpace: "nowrap" }}>{suffix}</span>}
    </div>
    {hint && <div style={{ fontSize: 11.5, color: T.dim, marginTop: 3 }}>{hint}</div>}
  </label>
);

const Card = ({ title, children, accent }) => (
  <div
    style={{
      background: T.card,
      border: `1px solid ${T.line}`,
      borderTop: accent ? `3px solid ${accent}` : `1px solid ${T.line}`,
      borderRadius: 12,
      padding: 18,
      marginBottom: 16,
    }}
  >
    {title && (
      <div
        style={{
          fontSize: 12,
          letterSpacing: 1,
          textTransform: "uppercase",
          color: T.dim,
          marginBottom: 12,
          fontWeight: 600,
        }}
      >
        {title}
      </div>
    )}
    {children}
  </div>
);

/* receipt-style line */
const RLine = ({ label, value, bold, color, indent, sign }) => (
  <div
    style={{
      display: "flex",
      justifyContent: "space-between",
      padding: "6px 0",
      paddingLeft: indent ? 16 : 0,
      borderBottom: bold ? "none" : `1px dashed ${T.line}`,
      fontWeight: bold ? 700 : 400,
      fontSize: bold ? 17 : 14.5,
      color: color || T.ink,
      fontVariantNumeric: "tabular-nums",
    }}
  >
    <span>{label}</span>
    <span>
      {sign}
      ฿{fmt(Math.abs(value))}
    </span>
  </div>
);

/* ตารางค่าคอมมิชชั่นโดยประมาณ (% รวม VAT, อ้างอิงประกาศ มิ.ย. 2026)
   Mall สูงกว่า Non-Mall ราว 1.5–3pp — เรทจริงต่างกันตาม leaf category
   ของแต่ละร้าน ควรเทียบกับ Seller Centre เสมอ */
const CATS = [
  { key: "fashion", label: "แฟชั่น / เสื้อผ้า / รองเท้า / กระเป๋า", nonmall: 9.63, mall: 12.84 },
  { key: "beauty", label: "ความงาม / สกินแคร์ / ของใช้ส่วนตัว", nonmall: 10.7, mall: 13.91 },
  { key: "fmcg", label: "FMCG / ของใช้ในบ้าน / อุปโภคบริโภค", nonmall: 10.7, mall: 13.91 },
  { key: "food", label: "อาหารและเครื่องดื่ม", nonmall: 9.63, mall: 12.84 },
  { key: "pet", label: "สัตว์เลี้ยงและอุปกรณ์", nonmall: 9.63, mall: 12.84 },
  { key: "momkid", label: "แม่และเด็ก / ของเล่น", nonmall: 9.63, mall: 12.84 },
  { key: "home", label: "บ้านและสวน / เฟอร์นิเจอร์", nonmall: 9.63, mall: 12.84 },
  { key: "electronics", label: "เครื่องใช้ไฟฟ้า / แกดเจ็ต / อุปกรณ์ IT", nonmall: 7.49, mall: 9.63 },
  { key: "mobile", label: "มือถือ / แท็บเล็ต", nonmall: 7.49, mall: 8.56 },
  { key: "auto", label: "ยานยนต์ / อะไหล่", nonmall: 8.56, mall: 10.7 },
  { key: "sport", label: "กีฬาและกิจกรรมกลางแจ้ง", nonmall: 9.63, mall: 12.84 },
  { key: "custom", label: "กำหนดเอง (ใส่เลขเอง)", nonmall: null, mall: null },
];

/* ================= MAIN APP ================= */

export default function App() {
  const [tab, setTab] = useState("calc");

  /* --- fee settings (editable, defaults = มิ.ย. 2026) --- */
  const [fees, setFees] = useState({
    shopType: "nonmall", // nonmall | mall
    category: "fashion",
    commission: 9.63, // % รวม VAT — auto จากหมวด + ประเภทร้าน, แก้เองได้
    serviceProgram: 6.42, // % โปรแกรมส่งฟรีพิเศษ/โค้ดคุ้ม (รวม VAT)
    joinProgram: true,
    payment: 3.21, // % ของ (ราคาสุทธิ + ค่าส่งที่ลูกค้าจ่าย)
    infra: 1, // บาท/ออเดอร์
  });

  const pickCategory = (catKey, shopType) => {
    const cat = CATS.find((c) => c.key === catKey);
    const rate = cat && cat[shopType === "mall" ? "mall" : "nonmall"];
    setFees((f) => ({
      ...f,
      category: catKey,
      shopType,
      commission: rate !== null && rate !== undefined ? rate : f.commission,
    }));
  };

  /* --- staff groups: มีได้ทั้งรายวันและรายเดือนพร้อมกัน --- */
  const [staff, setStaff] = useState([
    { id: 1, count: 1, payType: "monthly", rate: 12000, workDays: 26 },
  ]);
  const updStaff = (id, patch) => setStaff((s) => s.map((g) => (g.id === id ? { ...g, ...patch } : g)));
  const addStaff = () =>
    setStaff((s) => [...s, { id: Date.now(), count: 1, payType: "daily", rate: 400, workDays: 26 }]);
  const delStaff = (id) => setStaff((s) => s.filter((g) => g.id !== id));

  /* --- hidden costs --- */
  const [hc, setHc] = useState({
    boxCost: 4,
    bubbleCost: 2,
    tapeCost: 0.5,
    otherPack: 0,
    rentMonthly: 0,
    utilMonthly: 0,
    otherMonthly: 0,
    ordersPerMonth: 300,
  });

  /* derived hidden cost */
  const hidden = useMemo(() => {
    const staffMonthly = staff.reduce(
      (s, g) =>
        s +
        (g.payType === "monthly"
          ? num(g.count) * num(g.rate)
          : num(g.count) * num(g.rate) * num(g.workDays)),
      0
    );
    const fixedMonthly = staffMonthly + num(hc.rentMonthly) + num(hc.utilMonthly) + num(hc.otherMonthly);
    const packPerOrder = num(hc.boxCost) + num(hc.bubbleCost) + num(hc.tapeCost) + num(hc.otherPack);
    const orders = Math.max(num(hc.ordersPerMonth), 1);
    const fixedPerOrder = fixedMonthly / orders;
    return { staffMonthly, fixedMonthly, packPerOrder, fixedPerOrder, perOrder: packPerOrder + fixedPerOrder };
  }, [hc, staff]);

  /* --- manual calculator inputs --- */
  const [calc, setCalc] = useState({
    price: 500,
    sellerDiscount: 0,
    shippingBuyerPays: 0,
    cogs: 250,
  });

  const result = useMemo(() => {
    const net = num(calc.price) - num(calc.sellerDiscount);
    const commission = (net * num(fees.commission)) / 100;
    const service = fees.joinProgram ? (net * num(fees.serviceProgram)) / 100 : 0;
    const payment = ((net + num(calc.shippingBuyerPays)) * num(fees.payment)) / 100;
    const infra = num(fees.infra);
    const shopeeTotal = commission + service + payment + infra;
    const received = net - shopeeTotal;
    const profitBeforeHidden = received - num(calc.cogs);
    const profit = profitBeforeHidden - hidden.perOrder;
    const margin = net > 0 ? (profit / net) * 100 : 0;
    const feePct = net > 0 ? (shopeeTotal / net) * 100 : 0;
    return { net, commission, service, payment, infra, shopeeTotal, received, profit, margin, feePct };
  }, [calc, fees, hidden]);

  /* --- report upload --- */
  const [report, setReport] = useState(null); // {rows, cols, mapping, fileName}
  const [reportCogs, setReportCogs] = useState({ mode: "percent", value: 50 });
  const [reportErr, setReportErr] = useState("");
  const fileRef = useRef(null);

  const KEYS = {
    orderId: ["หมายเลขคำสั่งซื้อ", "order id", "order sn", "เลขที่คำสั่งซื้อ"],
    sales: ["ราคาขายสุทธิ", "ราคาสินค้าที่ชำระ", "original price", "product price", "ราคาตั้งต้น", "ราคาขาย"],
    commission: ["คอมมิชชั่น", "commission"],
    txFee: ["ธุรกรรม", "transaction fee"],
    serviceFee: ["ค่าบริการ", "service fee"],
    released: ["ปล่อย", "released", "รายรับ", "ยอดเงินที่", "total release"],
  };

  const matchCol = (cols, keys) => {
    const low = cols.map((c) => String(c).toLowerCase());
    for (const k of keys) {
      const i = low.findIndex((c) => c.includes(k.toLowerCase()));
      if (i !== -1) return cols[i];
    }
    return "";
  };

  const handleFile = async (file) => {
    setReportErr("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      // find header row within first 15 rows
      let hi = 0;
      for (let i = 0; i < Math.min(raw.length, 15); i++) {
        const rowStr = raw[i].join("|").toLowerCase();
        if (rowStr.includes("คำสั่งซื้อ") || rowStr.includes("order")) {
          hi = i;
          break;
        }
      }
      const cols = raw[hi].map((c) => String(c).trim()).filter(Boolean);
      const rows = raw
        .slice(hi + 1)
        .filter((r) => r.some((v) => v !== ""))
        .map((r) => {
          const o = {};
          raw[hi].forEach((c, i) => {
            if (String(c).trim()) o[String(c).trim()] = r[i];
          });
          return o;
        });
      if (!rows.length) throw new Error("ไม่พบข้อมูลออเดอร์ในไฟล์");
      const mapping = {
        sales: matchCol(cols, KEYS.sales),
        commission: matchCol(cols, KEYS.commission),
        txFee: matchCol(cols, KEYS.txFee),
        serviceFee: matchCol(cols, KEYS.serviceFee),
        released: matchCol(cols, KEYS.released),
      };
      setReport({ rows, cols, mapping, fileName: file.name });
    } catch (e) {
      setReportErr("อ่านไฟล์ไม่สำเร็จ: " + e.message + " — ลองใช้ไฟล์ .xlsx จากเมนู รายรับของฉัน ใน Seller Centre");
    }
  };

  const reportResult = useMemo(() => {
    if (!report) return null;
    const m = report.mapping;
    const sum = (col) => (col ? report.rows.reduce((s, r) => s + num(r[col]), 0) : 0);
    const orders = report.rows.length;
    const sales = sum(m.sales);
    const commission = Math.abs(sum(m.commission));
    const txFee = Math.abs(sum(m.txFee));
    const serviceFee = Math.abs(sum(m.serviceFee));
    const released = sum(m.released);
    const shopeeFees = commission + txFee + serviceFee;
    const income = m.released ? released : sales - shopeeFees;
    const cogsTotal =
      reportCogs.mode === "percent" ? (sales * num(reportCogs.value)) / 100 : num(reportCogs.value) * orders;
    const hiddenTotal = hidden.packPerOrder * orders + (hidden.fixedMonthly * orders) / Math.max(num(hc.ordersPerMonth), 1);
    const profit = income - cogsTotal - hiddenTotal;
    const margin = sales > 0 ? (profit / sales) * 100 : 0;
    return { orders, sales, commission, txFee, serviceFee, shopeeFees, income, cogsTotal, hiddenTotal, profit, margin };
  }, [report, reportCogs, hidden, hc.ordersPerMonth]);

  const setMap = (k, v) => setReport((r) => ({ ...r, mapping: { ...r.mapping, [k]: v } }));

  /* ================= RENDER ================= */

  const tabs = [
    ["calc", "คำนวณต่อออเดอร์"],
    ["hidden", "ต้นทุนแฝง"],
    ["report", "อัปโหลด Report"],
    ["fees", "ตั้งค่าเรท"],
  ];

  return (
    <div style={{ minHeight: "100vh", background: T.paper, fontFamily: "'Noto Sans Thai','IBM Plex Sans Thai',system-ui,sans-serif", color: T.ink }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;600;700;800&display=swap');
        input:focus{border-color:${T.ink}!important}
        button{cursor:pointer}
      `}</style>

      {/* header */}
      <div style={{ background: T.ink, color: "#fff", padding: "22px 20px 18px" }}>
        <div style={{ maxWidth: 960, margin: "0 auto" }}>
          <div style={{ fontSize: 12, letterSpacing: 2, opacity: 0.7, textTransform: "uppercase" }}>สมุดกำไรแม่ค้า</div>
          <div style={{ fontSize: 26, fontWeight: 800, marginTop: 2 }}>คำนวณกำไรจริง Shopee</div>
          <div style={{ fontSize: 13, opacity: 0.75, marginTop: 4 }}>
            หักครบทุกค่าธรรมเนียม + ต้นทุนแฝงที่มองไม่เห็น · เรทอ้างอิง มิ.ย. 2026 (แก้ไขได้)
          </div>
        </div>
      </div>

      {/* tabs */}
      <div style={{ background: "#fff", borderBottom: `1px solid ${T.line}`, position: "sticky", top: 0, zIndex: 5 }}>
        <div style={{ maxWidth: 960, margin: "0 auto", display: "flex", overflowX: "auto" }}>
          {tabs.map(([k, label]) => (
            <button
              key={k}
              onClick={() => setTab(k)}
              style={{
                padding: "13px 16px",
                border: "none",
                background: "none",
                fontSize: 14.5,
                fontWeight: tab === k ? 700 : 400,
                color: tab === k ? T.ink : T.dim,
                borderBottom: tab === k ? `3px solid ${T.ink}` : "3px solid transparent",
                whiteSpace: "nowrap",
                fontFamily: "inherit",
              }}
            >
              {label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 16px 60px" }}>
        {/* ============ TAB: MANUAL CALC ============ */}
        {tab === "calc" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16 }}>
            <div>
              <Card title="ข้อมูลออเดอร์">
                <Field label="ราคาขาย" value={calc.price} onChange={(v) => setCalc({ ...calc, price: v })} suffix="บาท" />
                <Field
                  label="ส่วนลดที่ร้านออกเอง (โค้ด/โปรร้าน)"
                  value={calc.sellerDiscount}
                  onChange={(v) => setCalc({ ...calc, sellerDiscount: v })}
                  suffix="บาท"
                />
                <Field
                  label="ค่าส่งที่ลูกค้าจ่าย"
                  value={calc.shippingBuyerPays}
                  onChange={(v) => setCalc({ ...calc, shippingBuyerPays: v })}
                  suffix="บาท"
                  hint="ใช้คิดค่าธรรมเนียมธุรกรรม (คิดจากราคา + ค่าส่ง)"
                />
                <Field label="ต้นทุนสินค้า (COGS)" value={calc.cogs} onChange={(v) => setCalc({ ...calc, cogs: v })} suffix="บาท" />
              </Card>
              <Card title="เรทที่ใช้คำนวณ">
                <div style={{ fontSize: 13.5, color: T.dim, lineHeight: 1.9 }}>
                  {fees.shopType === "mall" ? "Shopee Mall" : "ร้านทั่วไป"} · {(CATS.find((c) => c.key === fees.category) || {}).label || "กำหนดเอง"}
                  <br />
                  คอมมิชชั่น {fees.commission}% · โปรแกรมส่งฟรี/โค้ดคุ้ม {fees.joinProgram ? fees.serviceProgram + "%" : "ไม่เข้าร่วม"} ·
                  ธุรกรรม {fees.payment}% · Infra ฿{fees.infra}/ออเดอร์ · ต้นทุนแฝง ฿{fmt(hidden.perOrder)}/ออเดอร์
                  <br />
                  <button
                    onClick={() => setTab("fees")}
                    style={{ border: "none", background: "none", color: T.profit, fontWeight: 600, padding: 0, fontSize: 13.5, fontFamily: "inherit", textDecoration: "underline" }}
                  >
                    แก้ไขเรท →
                  </button>
                </div>
              </Card>
            </div>

            {/* signature receipt */}
            <Card title="ใบเสร็จกำไรจริง" accent={result.profit >= 0 ? T.profit : T.loss}>
              <RLine label="ราคาหลังหักส่วนลดร้าน" value={result.net} />
              <RLine label="คอมมิชชั่น Shopee" value={result.commission} color={T.loss} sign="−" indent />
              {fees.joinProgram && <RLine label="ค่าโปรแกรมส่งฟรี/โค้ดคุ้ม" value={result.service} color={T.loss} sign="−" indent />}
              <RLine label="ค่าธรรมเนียมธุรกรรม" value={result.payment} color={T.loss} sign="−" indent />
              <RLine label="ค่า Infrastructure" value={result.infra} color={T.loss} sign="−" indent />
              <RLine label="เงินเข้าร้านจาก Shopee" value={result.received} bold />
              <div style={{ borderTop: `1px solid ${T.line}`, margin: "6px 0" }} />
              <RLine label="ต้นทุนสินค้า" value={num(calc.cogs)} color={T.loss} sign="−" indent />
              <RLine label="ต้นทุนแฝง (แพ็ค+แรงงาน+Fix cost)" value={hidden.perOrder} color={T.amber} sign="−" indent />
              <div
                style={{
                  marginTop: 12,
                  background: result.profit >= 0 ? "#EAF5EE" : "#FBEBEB",
                  border: `1px solid ${result.profit >= 0 ? T.profit : T.loss}`,
                  borderRadius: 10,
                  padding: "14px 16px",
                }}
              >
                <div style={{ fontSize: 13, color: T.dim }}>กำไรจริงต่อออเดอร์</div>
                <div style={{ fontSize: 34, fontWeight: 800, color: result.profit >= 0 ? T.profit : T.loss, fontVariantNumeric: "tabular-nums" }}>
                  ฿{fmt(result.profit)}
                </div>
                <div style={{ fontSize: 13.5, color: T.dim, marginTop: 2 }}>
                  มาร์จิ้น {fmt(result.margin, 1)}% · Shopee หักไป {fmt(result.feePct, 1)}% ของยอดขาย
                </div>
                {result.profit < 0 && (
                  <div style={{ fontSize: 13.5, color: T.loss, marginTop: 6, fontWeight: 600 }}>
                    ⚠ ขายราคานี้ขาดทุน — ต้องตั้งราคาอย่างน้อย ฿
                    {fmt((num(calc.cogs) + hidden.perOrder + num(fees.infra)) / (1 - (num(fees.commission) + (fees.joinProgram ? num(fees.serviceProgram) : 0) + num(fees.payment)) / 100), 0)}{" "}
                    ถึงจะเท่าทุน
                  </div>
                )}
              </div>
            </Card>
          </div>
        )}

        {/* ============ TAB: HIDDEN COSTS ============ */}
        {tab === "hidden" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16 }}>
            <div>
              <Card title="พนักงาน (มีได้ทั้งรายวันและรายเดือน)">
                {staff.map((g, i) => {
                  const monthlyOfGroup =
                    g.payType === "monthly"
                      ? num(g.count) * num(g.rate)
                      : num(g.count) * num(g.rate) * num(g.workDays);
                  return (
                    <div
                      key={g.id}
                      style={{
                        border: `1px solid ${T.line}`,
                        borderRadius: 10,
                        padding: 14,
                        marginBottom: 12,
                        background: "#FAFCFA",
                      }}
                    >
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: T.dim }}>
                          กลุ่มที่ {i + 1} · {g.payType === "monthly" ? "รายเดือน" : "รายวัน"}
                        </div>
                        {staff.length > 1 && (
                          <button
                            onClick={() => delStaff(g.id)}
                            style={{ border: "none", background: "none", color: T.loss, fontSize: 13, fontWeight: 600, fontFamily: "inherit", padding: 0 }}
                          >
                            ลบกลุ่มนี้
                          </button>
                        )}
                      </div>
                      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
                        {[
                          ["monthly", "รายเดือน"],
                          ["daily", "รายวัน"],
                        ].map(([k, label]) => (
                          <button
                            key={k}
                            onClick={() => updStaff(g.id, { payType: k })}
                            style={{
                              flex: 1,
                              padding: "8px 0",
                              borderRadius: 8,
                              border: `1px solid ${g.payType === k ? T.ink : T.line}`,
                              background: g.payType === k ? T.ink : "#fff",
                              color: g.payType === k ? "#fff" : T.dim,
                              fontWeight: 600,
                              fontSize: 13.5,
                              fontFamily: "inherit",
                            }}
                          >
                            {label}
                          </button>
                        ))}
                      </div>
                      <Field label="จำนวนพนักงานกลุ่มนี้" value={g.count} onChange={(v) => updStaff(g.id, { count: v })} suffix="คน" />
                      <Field
                        label={g.payType === "monthly" ? "เงินเดือนต่อคน" : "ค่าแรงรายวันต่อคน"}
                        value={g.rate}
                        onChange={(v) => updStaff(g.id, { rate: v })}
                        suffix={g.payType === "monthly" ? "บาท/เดือน" : "บาท/วัน"}
                      />
                      {g.payType === "daily" && (
                        <Field label="วันทำงานต่อเดือน" value={g.workDays} onChange={(v) => updStaff(g.id, { workDays: v })} suffix="วัน" />
                      )}
                      <div style={{ fontSize: 13, color: T.dim, fontVariantNumeric: "tabular-nums" }}>
                        = ฿{fmt(monthlyOfGroup, 0)}/เดือน
                      </div>
                    </div>
                  );
                })}
                <button
                  onClick={addStaff}
                  style={{
                    width: "100%",
                    padding: "11px 0",
                    borderRadius: 10,
                    border: `1.5px dashed ${T.ink}`,
                    background: "none",
                    color: T.ink,
                    fontWeight: 700,
                    fontSize: 14,
                    fontFamily: "inherit",
                  }}
                >
                  + เพิ่มกลุ่มพนักงาน
                </button>
              </Card>
              <Card title="ค่าแพ็คเกจ ต่อ 1 ออเดอร์">
                <Field label="กล่อง/ซองไปรษณีย์" value={hc.boxCost} onChange={(v) => setHc({ ...hc, boxCost: v })} suffix="บาท" />
                <Field label="บับเบิ้ล/กันกระแทก" value={hc.bubbleCost} onChange={(v) => setHc({ ...hc, bubbleCost: v })} suffix="บาท" />
                <Field label="เทป/สติกเกอร์/ใบปะหน้า" value={hc.tapeCost} onChange={(v) => setHc({ ...hc, tapeCost: v })} suffix="บาท" />
                <Field label="อื่นๆ (การ์ดขอบคุณ ของแถม ฯลฯ)" value={hc.otherPack} onChange={(v) => setHc({ ...hc, otherPack: v })} suffix="บาท" />
              </Card>
              <Card title="ค่าใช้จ่ายคงที่รายเดือน">
                <Field label="ค่าเช่าที่/โกดัง" value={hc.rentMonthly} onChange={(v) => setHc({ ...hc, rentMonthly: v })} suffix="บาท" />
                <Field label="ค่าน้ำ-ไฟ-เน็ต" value={hc.utilMonthly} onChange={(v) => setHc({ ...hc, utilMonthly: v })} suffix="บาท" />
                <Field label="อื่นๆ (ค่าโฆษณา sub app ฯลฯ)" value={hc.otherMonthly} onChange={(v) => setHc({ ...hc, otherMonthly: v })} suffix="บาท" />
                <Field label="ออเดอร์เฉลี่ยต่อเดือน" value={hc.ordersPerMonth} onChange={(v) => setHc({ ...hc, ordersPerMonth: v })} suffix="ออเดอร์" hint="ใช้เฉลี่ยต้นทุนคงที่ลงต่อออเดอร์" />
              </Card>
            </div>
            <Card title="สรุปต้นทุนแฝง" accent={T.amber}>
              <RLine label="ค่าพนักงานรวม/เดือน" value={hidden.staffMonthly} />
              <RLine label="ค่าใช้จ่ายคงที่รวม/เดือน" value={hidden.fixedMonthly} bold />
              <div style={{ borderTop: `1px solid ${T.line}`, margin: "8px 0" }} />
              <RLine label="แพ็คเกจต่อออเดอร์" value={hidden.packPerOrder} />
              <RLine label="Fix cost เฉลี่ยต่อออเดอร์" value={hidden.fixedPerOrder} />
              <div style={{ marginTop: 12, background: "#FBF4E4", border: `1px solid ${T.amber}`, borderRadius: 10, padding: "14px 16px" }}>
                <div style={{ fontSize: 13, color: T.dim }}>ต้นทุนแฝงรวมต่อ 1 ออเดอร์</div>
                <div style={{ fontSize: 32, fontWeight: 800, color: T.amber, fontVariantNumeric: "tabular-nums" }}>฿{fmt(hidden.perOrder)}</div>
                <div style={{ fontSize: 13, color: T.dim, marginTop: 4 }}>
                  ตัวเลขนี้ถูกหักอัตโนมัติในหน้า คำนวณต่อออเดอร์ และ อัปโหลด Report — นี่คือเงินที่แม่ค้าส่วนใหญ่ลืมคิด แล้วนึกว่าตัวเองกำไร
                </div>
              </div>
            </Card>
          </div>
        )}

        {/* ============ TAB: REPORT ============ */}
        {tab === "report" && (
          <div>
            <Card title="อัปโหลดรายงานรายรับจาก Shopee">
              <div style={{ fontSize: 13.5, color: T.dim, marginBottom: 12, lineHeight: 1.8 }}>
                ดาวน์โหลดจาก Seller Centre → การเงิน → <b>รายรับของฉัน</b> → Export เป็น Excel แล้วลากมาวางที่นี่
              </div>
              <div
                onClick={() => fileRef.current?.click()}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
                }}
                style={{
                  border: `2px dashed ${T.line}`,
                  borderRadius: 12,
                  padding: "36px 20px",
                  textAlign: "center",
                  color: T.dim,
                  cursor: "pointer",
                  background: "#FAFCFA",
                }}
              >
                {report ? (
                  <span style={{ color: T.profit, fontWeight: 700 }}>✓ {report.fileName} · {report.rows.length} ออเดอร์</span>
                ) : (
                  <span>คลิกเลือกไฟล์ หรือลากไฟล์ .xlsx / .csv มาวาง</span>
                )}
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
              {reportErr && <div style={{ color: T.loss, fontSize: 13.5, marginTop: 10 }}>{reportErr}</div>}
            </Card>

            {report && (
              <>
                <Card title="ตรวจสอบการจับคู่คอลัมน์ (แก้ได้ถ้าไม่ตรง)">
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 12 }}>
                    {[
                      ["sales", "ยอดขาย/ราคาสินค้า"],
                      ["commission", "ค่าคอมมิชชั่น"],
                      ["txFee", "ค่าธรรมเนียมธุรกรรม"],
                      ["serviceFee", "ค่าบริการ"],
                      ["released", "เงินที่ปล่อยให้ผู้ขาย"],
                    ].map(([k, label]) => (
                      <label key={k} style={{ fontSize: 13 }}>
                        <div style={{ color: T.dim, marginBottom: 4 }}>{label}</div>
                        <select
                          value={report.mapping[k]}
                          onChange={(e) => setMap(k, e.target.value)}
                          style={{ width: "100%", padding: "9px 10px", borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 13.5, fontFamily: "inherit", background: "#fff", color: T.ink }}
                        >
                          <option value="">— ไม่ใช้ —</option>
                          {report.cols.map((c) => (
                            <option key={c} value={c}>{c}</option>
                          ))}
                        </select>
                      </label>
                    ))}
                  </div>
                </Card>

                <Card title="ต้นทุนสินค้าของร้าน">
                  <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                    {[
                      ["percent", "% ของยอดขาย"],
                      ["perOrder", "บาท/ออเดอร์ (เฉลี่ย)"],
                    ].map(([k, label]) => (
                      <button
                        key={k}
                        onClick={() => setReportCogs({ ...reportCogs, mode: k })}
                        style={{
                          flex: 1,
                          padding: "9px 0",
                          borderRadius: 8,
                          border: `1px solid ${reportCogs.mode === k ? T.ink : T.line}`,
                          background: reportCogs.mode === k ? T.ink : "#fff",
                          color: reportCogs.mode === k ? "#fff" : T.dim,
                          fontWeight: 600,
                          fontSize: 13.5,
                          fontFamily: "inherit",
                        }}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                  <Field
                    label={reportCogs.mode === "percent" ? "ต้นทุนสินค้าคิดเป็น % ของยอดขาย" : "ต้นทุนสินค้าเฉลี่ยต่อออเดอร์"}
                    value={reportCogs.value}
                    onChange={(v) => setReportCogs({ ...reportCogs, value: v })}
                    suffix={reportCogs.mode === "percent" ? "%" : "บาท"}
                  />
                </Card>

                {reportResult && (
                  <Card title={`สรุปกำไรจริงจาก Report · ${reportResult.orders} ออเดอร์`} accent={reportResult.profit >= 0 ? T.profit : T.loss}>
                    <RLine label="ยอดขายรวม" value={reportResult.sales} />
                    <RLine label="ค่าคอมมิชชั่น" value={reportResult.commission} color={T.loss} sign="−" indent />
                    <RLine label="ค่าธรรมเนียมธุรกรรม" value={reportResult.txFee} color={T.loss} sign="−" indent />
                    <RLine label="ค่าบริการ (โปรแกรมต่างๆ)" value={reportResult.serviceFee} color={T.loss} sign="−" indent />
                    <RLine label="เงินเข้าร้านจริง" value={reportResult.income} bold />
                    <div style={{ borderTop: `1px solid ${T.line}`, margin: "6px 0" }} />
                    <RLine label="ต้นทุนสินค้ารวม" value={reportResult.cogsTotal} color={T.loss} sign="−" indent />
                    <RLine label="ต้นทุนแฝงรวม (จากแท็บต้นทุนแฝง)" value={reportResult.hiddenTotal} color={T.amber} sign="−" indent />
                    <div
                      style={{
                        marginTop: 12,
                        background: reportResult.profit >= 0 ? "#EAF5EE" : "#FBEBEB",
                        border: `1px solid ${reportResult.profit >= 0 ? T.profit : T.loss}`,
                        borderRadius: 10,
                        padding: "14px 16px",
                      }}
                    >
                      <div style={{ fontSize: 13, color: T.dim }}>กำไรจริงของช่วงนี้</div>
                      <div style={{ fontSize: 34, fontWeight: 800, color: reportResult.profit >= 0 ? T.profit : T.loss, fontVariantNumeric: "tabular-nums" }}>
                        ฿{fmt(reportResult.profit)}
                      </div>
                      <div style={{ fontSize: 13.5, color: T.dim, marginTop: 2 }}>
                        มาร์จิ้นสุทธิ {fmt(reportResult.margin, 1)}% · เฉลี่ยกำไร ฿{fmt(reportResult.profit / Math.max(reportResult.orders, 1))}/ออเดอร์
                      </div>
                    </div>
                  </Card>
                )}
              </>
            )}
          </div>
        )}

        {/* ============ TAB: FEE SETTINGS ============ */}
        {tab === "fees" && (
          <div style={{ maxWidth: 520 }}>
            <Card title="เรทค่าธรรมเนียม Shopee (รวม VAT แล้ว)" accent={T.ink}>
              <div style={{ fontSize: 13, color: T.dim, marginBottom: 6 }}>ประเภทร้าน</div>
              <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                {[
                  ["nonmall", "ร้านทั่วไป (Non-Mall)"],
                  ["mall", "Shopee Mall"],
                ].map(([k, label]) => (
                  <button
                    key={k}
                    onClick={() => pickCategory(fees.category, k)}
                    style={{
                      flex: 1,
                      padding: "10px 0",
                      borderRadius: 8,
                      border: `1px solid ${fees.shopType === k ? T.ink : T.line}`,
                      background: fees.shopType === k ? T.ink : "#fff",
                      color: fees.shopType === k ? "#fff" : T.dim,
                      fontWeight: 600,
                      fontSize: 13.5,
                      fontFamily: "inherit",
                    }}
                  >
                    {label}
                  </button>
                ))}
              </div>
              <label style={{ display: "block", marginBottom: 14 }}>
                <div style={{ fontSize: 13, color: T.dim, marginBottom: 4 }}>หมวดหมู่สินค้าหลักของร้าน</div>
                <select
                  value={fees.category}
                  onChange={(e) => pickCategory(e.target.value, fees.shopType)}
                  style={{
                    width: "100%",
                    padding: "11px 12px",
                    borderRadius: 8,
                    border: `1px solid ${T.line}`,
                    fontSize: 14.5,
                    fontFamily: "inherit",
                    background: "#fff",
                    color: T.ink,
                  }}
                >
                  {CATS.map((c) => (
                    <option key={c.key} value={c.key}>
                      {c.label}
                      {c.nonmall !== null ? ` — ${fees.shopType === "mall" ? c.mall : c.nonmall}%` : ""}
                    </option>
                  ))}
                </select>
              </label>
              <Field
                label="ค่าคอมมิชชั่นที่ใช้คำนวณ (ปรับเองได้)"
                value={fees.commission}
                onChange={(v) => setFees({ ...fees, commission: v, category: "custom" })}
                suffix="%"
                hint="เรทในลิสต์เป็นค่าประมาณตามประกาศ มิ.ย. 2026 — เรทจริงต่างตาม leaf category เช็คของร้านคุณใน Seller Centre แล้วปรับเลขตรงนี้ให้ตรง"
              />
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
                <input
                  type="checkbox"
                  id="joinP"
                  checked={fees.joinProgram}
                  onChange={(e) => setFees({ ...fees, joinProgram: e.target.checked })}
                  style={{ width: 18, height: 18, accentColor: T.ink }}
                />
                <label htmlFor="joinP" style={{ fontSize: 14 }}>เข้าร่วมโปรแกรมส่งฟรีพิเศษ / โค้ดคุ้ม</label>
              </div>
              {fees.joinProgram && (
                <Field label="ค่าบริการโปรแกรม" value={fees.serviceProgram} onChange={(v) => setFees({ ...fees, serviceProgram: v })} suffix="%" />
              )}
              <Field
                label="ค่าธรรมเนียมธุรกรรม (Payment)"
                value={fees.payment}
                onChange={(v) => setFees({ ...fees, payment: v })}
                suffix="%"
                hint="คิดจากราคาสุทธิ + ค่าส่งที่ลูกค้าจ่าย (3% + VAT = 3.21%)"
              />
              <Field label="ค่า Infrastructure ต่อออเดอร์" value={fees.infra} onChange={(v) => setFees({ ...fees, infra: v })} suffix="บาท" />
              <div style={{ fontSize: 12.5, color: T.dim, lineHeight: 1.7, marginTop: 4 }}>
                ⓘ ค่าเริ่มต้นอ้างอิงประกาศ Shopee TH ล่าสุด (มิ.ย. 2026) — Shopee ปรับเรทบ่อย ควรเทียบกับ Seller Centre ของร้านคุณเสมอ
                กรณีลูกค้าผ่อนบัตร/SPayLater ค่าธรรมเนียมธุรกรรมจะสูงกว่านี้ (4–7% ตามงวดผ่อน)
              </div>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}
